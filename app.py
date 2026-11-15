from __future__ import annotations

import re
import shutil
import sqlite3
from contextlib import contextmanager
from pathlib import Path

from fastapi import FastAPI, Query, HTTPException, Body
from fastapi.responses import HTMLResponse
from export_service import (
    load_settings, save_settings, export_excel, start_scheduler, stop_scheduler,
    APP_VERSION, APP_AUTHOR, DB_PATH,
)
from extractor import collect_excel_files, import_files, parse_report, initialize_db

app = FastAPI(title="CTEP Shift Report Archive", version=APP_VERSION)


@app.on_event("startup")
def startup_event():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    settings = load_settings()
    Path(settings["import_path"]).mkdir(parents=True, exist_ok=True)
    Path(settings["archive_path"]).mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    initialize_db(conn)
    conn.close()
    start_scheduler()


@app.on_event("shutdown")
def shutdown_event():
    stop_scheduler()


@contextmanager
def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


@app.get("/api/settings")
def get_settings():
    return {**load_settings(), "version": APP_VERSION, "author": APP_AUTHOR}


@app.post("/api/settings")
def update_settings(payload: dict = Body(...)):
    try:
        return save_settings(
            payload.get("export_path", ""),
            payload.get("import_path", ""),
            payload.get("archive_path", ""),
            payload.get("export_interval_hours", 3),
            payload.get("auto_export_enabled", True),
        )
    except (ValueError, OSError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/export")
def run_export():
    try:
        return export_excel()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/api/import")
def run_import():
    settings = load_settings()
    source = Path(settings["import_path"]).expanduser().resolve()
    archive = Path(settings["archive_path"]).expanduser().resolve()
    if not source.is_dir():
        raise HTTPException(status_code=400, detail=f"Import folder does not exist: {source}")
    if source == archive or source in archive.parents:
        raise HTTPException(status_code=400, detail="Archive folder must be outside the Outlook Import folder")
    files = collect_excel_files([str(source)])
    if not files:
        with db() as conn:
            total = conn.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
        return {"found": 0, "imported": 0, "moved": 0, "total": total,
                "skipped_or_failed": [], "move_errors": [], "message": "No Excel reports found."}

    try:
        result = import_files(files, DB_PATH)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Database import failed; source files were not moved. {exc}")

    moved, move_errors = 0, []
    for raw_path in result["processed_paths"]:
        original = Path(raw_path)
        try:
            relative = original.relative_to(source)
            destination = archive / relative
            destination.parent.mkdir(parents=True, exist_ok=True)
            original_destination = destination
            suffix = 1
            while destination.exists():
                destination = original_destination.with_name(
                    f"{original_destination.stem}_{suffix}{original_destination.suffix}"
                )
                suffix += 1
            shutil.move(str(original), str(destination))
            moved += 1
            try:
                report = parse_report(destination)
                with db() as conn:
                    existing = conn.execute(
                        "SELECT id, source_path FROM reports WHERE doc_no=? AND content_hash=?",
                        (report["doc_no"], report["content_hash"]),
                    ).fetchone()
                    if existing and (not existing["source_path"] or not Path(existing["source_path"]).is_file()):
                        conn.execute("UPDATE reports SET source_path=? WHERE id=?",
                                     (str(destination), existing["id"]))
            except Exception as update_exc:
                move_errors.append(f"{original.name}: moved, but source link update failed: {update_exc}")
        except Exception as exc:
            move_errors.append(f"{original.name}: {exc}")

    return {
        "found": len(files),
        "imported": result["imported"],
        "moved": moved,
        "total": result["total"],
        "skipped_or_failed": result["skipped"],
        "move_errors": move_errors,
        "message": f"Imported {result['imported']} new report(s); moved {moved} processed file(s).",
    }


@app.get("/api/summary")
def summary():
    with db() as conn:
        total = conn.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
        shifts = conn.execute(
            "SELECT COALESCE(shift_type,'Unknown') AS shift_type, COUNT(*) AS n FROM reports GROUP BY shift_type ORDER BY shift_type"
        ).fetchall()
        dates = conn.execute("SELECT MIN(report_date), MAX(report_date) FROM reports").fetchone()
    return {"total": total, "shift_types": {row["shift_type"]: row["n"] for row in shifts},
            "from_date": dates[0], "to_date": dates[1]}


@app.get("/api/reports")
def reports(q: str = "", from_date: str = "", to_date: str = "",
            shift_type: str = "", limit: int = Query(500, ge=1, le=1000)):
    clauses, params = [], []
    terms = re.findall(r"[A-Za-z0-9#_-]+", q)
    if q.strip() and terms:
        clauses.append("report_search MATCH ?")
        params.append(" AND ".join('"' + t.replace('"', '""') + '"' for t in terms))
    if from_date:
        clauses.append("r.report_date >= ?")
        params.append(from_date)
    if to_date:
        clauses.append("r.report_date <= ?")
        params.append(to_date)
    if shift_type:
        clauses.append("r.shift_type = ?")
        params.append(shift_type)
    where = (" WHERE " + " AND ".join(clauses)) if clauses else ""
    if q.strip() and terms:
        sql = """SELECT r.id,r.doc_no,r.report_date,r.shift_type,r.shift,r.supervisor,
                 r.source_filename,snippet(report_search,5,'[',']',' ... ',18) AS excerpt
                 FROM report_search JOIN reports r ON r.id=report_search.rowid""" + where + " ORDER BY r.report_date DESC,r.shift LIMIT ?"
    else:
        sql = """SELECT r.id,r.doc_no,r.report_date,r.shift_type,r.shift,r.supervisor,
                 r.source_filename,substr(r.activity_report,1,240) AS excerpt
                 FROM reports r""" + where + " ORDER BY r.report_date DESC,r.shift LIMIT ?"
    params.append(limit)
    with db() as conn:
        rows = conn.execute(sql, params).fetchall()
    return {"count": len(rows), "items": [dict(row) for row in rows]}


@app.get("/api/reports/{report_id}")
def report_detail(report_id: int):
    with db() as conn:
        row = conn.execute(
            """SELECT id,doc_no,report_date,shift_type,shift,supervisor,source_filename,
                      source_path,activity_report,full_text
               FROM reports WHERE id=?""", (report_id,)
        ).fetchone()
    if row is None:
        raise HTTPException(status_code=404, detail="Report not found")
    return dict(row)


@app.get("/", response_class=HTMLResponse)
def home():
    return HTML


HTML = r"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>CTEP Shift Report Archive</title>
<style>
:root{font-family:Segoe UI,Arial,sans-serif;color:#182433;background:#f3f6fa}*{box-sizing:border-box}
body{margin:0}.top{background:#173b5e;color:white;padding:22px 5vw}.top h1{margin:0 0 6px;font-size:25px}
.top p{margin:0;color:#d7e5f2}.top small{display:block;margin-top:8px;color:#d7e5f2}
main{max-width:1400px;margin:22px auto;padding:0 20px}.stats{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px}
.stat,.panel{background:white;border:1px solid #dfe6ee;border-radius:12px;padding:16px}.stat{min-width:145px;flex:1}
.stat b{display:block;font-size:24px;color:#173b5e}.stat span,.hint{font-size:12px;color:#65758a}
.actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap}button{padding:10px 16px;border:0;border-radius:8px;background:#1e5b8f;color:white;cursor:pointer}
button.secondary{background:#e8eef5;color:#24445f}button:disabled{opacity:.55;cursor:wait}
.settings{display:none;margin-top:15px;grid-template-columns:2fr 2fr 2fr 1fr;gap:12px;align-items:end}
label{display:block;font-size:12px;color:#536477;margin-bottom:5px}input,select{width:100%;padding:10px;border:1px solid #cbd5e1;border-radius:8px;background:white;color:#182433}
.filters{display:grid;grid-template-columns:2fr 1fr 1fr 1fr auto;gap:10px;align-items:end}
.meta{display:flex;justify-content:space-between;align-items:center;margin:14px 0 8px;color:#617286;font-size:13px}
table{width:100%;border-collapse:collapse}th,td{text-align:left;padding:11px 10px;border-bottom:1px solid #e8edf2;vertical-align:top}
th{font-size:12px;color:#66788a;background:#f8fafc}td{font-size:14px}.doc{font-weight:650;color:#14558a;cursor:pointer}
.excerpt{font-size:12px;color:#65758a;max-width:700px}.pill{display:inline-block;padding:3px 8px;border-radius:20px;background:#eaf2f9;color:#244f72;font-size:12px}
.detail{margin-top:14px;display:none}.detail pre{white-space:pre-wrap;overflow-wrap:anywhere;background:#f7f9fc;padding:14px;border-radius:8px;max-height:500px;overflow:auto;font:13px/1.5 Consolas,monospace}
.empty{text-align:center;padding:28px;color:#738297}@media(max-width:850px){.filters,.settings{grid-template-columns:1fr 1fr}.wide{grid-column:1/-1}table{display:block;overflow-x:auto}}
</style></head><body>
<header class="top"><h1>CTEP Shift Report Archive</h1><p>Local searchable archive · Day and Night shift reports</p>
<small>Created by Farzad Meshkipour · Version 1.2.0</small></header>
<main>
<section class="stats">
<div class="stat"><b id="total">—</b><span>Report versions indexed</span></div>
<div class="stat"><b id="day">—</b><span>Day shift reports</span></div>
<div class="stat"><b id="night">—</b><span>Night shift reports</span></div>
<div class="stat"><b id="range">—</b><span>Report date range</span></div></section>
<section class="panel" style="margin-bottom:14px">
<div class="actions"><button id="import">Import Outlook Folder</button><button id="export">Export to Excel</button>
<button class="secondary" id="settingsToggle">Settings</button><span id="actionStatus" class="hint"></span></div>
<div id="settingsPanel" class="settings">
<div><label>Outlook Import folder (recursive)</label><input id="importPath"></div>
<div><label>Processed reports archive folder</label><input id="archivePath"></div>
<div><label>Excel export file (.xlsx)</label><input id="exportPath"></div>
<div><label>Export interval (hours)</label><input id="exportInterval" type="number" min="1" max="168"></div>
<div><label><input id="autoExport" type="checkbox" style="width:auto"> Automatic export</label></div>
<div><button id="saveSettings">Save Settings</button></div>
<p id="settingsStatus" class="hint" style="grid-column:1/-1"></p></div></section>
<section class="panel"><div class="filters">
<div class="wide"><label>Search report content, equipment tags, wells, activities</label><input id="q" placeholder="e.g. F07, GTG, PSV-02-954-27, shutdown"></div>
<div><label>From date</label><input id="from" type="date"></div><div><label>To date</label><input id="to" type="date"></div>
<div><label>Shift</label><select id="shift"><option value="">All shifts</option><option>Day</option><option>Night</option></select></div>
<div><button id="search">Search</button></div></div></section>
<div class="meta"><span id="count">Loading reports…</span><button class="secondary" id="clear">Clear filters</button></div>
<section class="panel"><table><thead><tr><th>Document No.</th><th>Date</th><th>Shift</th><th>Supervisor</th><th>Matched content</th></tr></thead>
<tbody id="results"><tr><td colspan="5" class="empty">Loading…</td></tr></tbody></table></section>
<section class="panel detail" id="detail"><h3 id="detailTitle"></h3><div id="detailMeta" class="hint"></div><pre id="detailText"></pre></section>
</main>
<script>
const $=id=>document.getElementById(id);
async function api(url,options){const r=await fetch(url,options);const d=await r.json();if(!r.ok)throw Error(d.detail||'Request failed');return d;}
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
async function loadSummary(){const s=await api('/api/summary');$('total').textContent=s.total;$('day').textContent=s.shift_types.Day||0;$('night').textContent=s.shift_types.Night||0;$('range').textContent=(s.from_date||'—')+' → '+(s.to_date||'—');}
async function search(){const p=new URLSearchParams({q:$('q').value,from_date:$('from').value,to_date:$('to').value,shift_type:$('shift').value});const d=await api('/api/reports?'+p);$('count').textContent=d.count+' report(s) found';
$('results').innerHTML=d.items.length?d.items.map(x=>'<tr><td><span class="doc" data-id="'+x.id+'">'+esc(x.doc_no)+'</span></td><td>'+esc(x.report_date)+'</td><td><span class="pill">'+esc(x.shift_type)+' (Shift '+esc(x.shift)+')</span></td><td>'+esc(x.supervisor)+'</td><td class="excerpt">'+esc(x.excerpt||'')+'</td></tr>').join(''):'<tr><td colspan="5" class="empty">No matching reports found.</td></tr>';
document.querySelectorAll('.doc').forEach(el=>el.addEventListener('click',()=>details(el.dataset.id)));}
async function details(id){const r=await api('/api/reports/'+id);$('detailTitle').textContent=r.doc_no+' — '+r.report_date;$('detailMeta').textContent=r.shift_type+' (Shift '+r.shift+') · Supervisor: '+r.supervisor+' · Source: '+r.source_filename;$('detailText').textContent=r.full_text;$('detail').style.display='block';$('detail').scrollIntoView({behavior:'smooth',block:'start'});}
async function initSettings(){const s=await api('/api/settings');$('importPath').value=s.import_path;$('archivePath').value=s.archive_path;$('exportPath').value=s.export_path;$('exportInterval').value=s.export_interval_hours;$('autoExport').checked=s.auto_export_enabled;}
$('search').onclick=()=>search().catch(e=>$('actionStatus').textContent=e.message);
$('clear').onclick=()=>{$('q').value='';$('from').value='';$('to').value='';$('shift').value='';search();};
$('q').onkeydown=e=>{if(e.key==='Enter')search();};
$('settingsToggle').onclick=()=>{$('settingsPanel').style.display=$('settingsPanel').style.display==='grid'?'none':'grid';};
$('saveSettings').onclick=async()=>{try{await api('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({import_path:$('importPath').value,archive_path:$('archivePath').value,export_path:$('exportPath').value,export_interval_hours:Number($('exportInterval').value),auto_export_enabled:$('autoExport').checked})});$('settingsStatus').textContent='Settings saved.';}catch(e){$('settingsStatus').textContent=e.message;}};
$('export').onclick=async()=>{const b=$('export');b.disabled=true;$('actionStatus').textContent='Exporting…';try{const d=await api('/api/export',{method:'POST'});$('actionStatus').textContent='Exported '+d.reports+' reports: '+d.path;}catch(e){$('actionStatus').textContent='Export failed: '+e.message;}finally{b.disabled=false;}};
$('import').onclick=async()=>{if(!confirm('Import all Excel reports from the configured folder and subfolders? Successfully imported or duplicate files will be moved to the processed folder.'))return;const b=$('import');b.disabled=true;$('actionStatus').textContent='Scanning and importing…';try{const d=await api('/api/import',{method:'POST'});$('actionStatus').textContent=d.message+' Total in database: '+d.total+'. '+(d.skipped_or_failed?.length||0)+' skipped/error(s).';await loadSummary();await search();if(d.skipped_or_failed?.length||d.move_errors?.length)alert([...d.skipped_or_failed,...d.move_errors].join('\n'));}catch(e){$('actionStatus').textContent='Import failed: '+e.message;}finally{b.disabled=false;}};
Promise.all([loadSummary(),search(),initSettings()]).catch(e=>$('actionStatus').textContent=e.message);
</script></body></html>"""
