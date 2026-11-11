from __future__ import annotations

import re
import sqlite3
from pathlib import Path
from fastapi import FastAPI, Query, HTTPException, Body
from export_service import load_settings, save_settings, export_excel, start_scheduler, stop_scheduler, APP_VERSION, APP_AUTHOR
from fastapi.responses import HTMLResponse

BASE = Path(__import__('sys').executable).parent if getattr(__import__('sys'), 'frozen', False) else Path(__file__).parent
DB_PATH = BASE / "data" / "ctep_archive.db"
app = FastAPI(title="CTEP Shift Report Archive", version=APP_VERSION)


@app.on_event("startup")
def startup_event():
    start_scheduler()


@app.on_event("shutdown")
def shutdown_event():
    stop_scheduler()


@app.get("/api/settings")
def get_settings():
    return {**load_settings(), "version": APP_VERSION, "author": APP_AUTHOR}


@app.post("/api/settings")
def update_settings(payload: dict = Body(...)):
    try:
        return save_settings(payload.get("export_path", ""), payload.get("export_interval_hours", 3), payload.get("auto_export_enabled", True))
    except (ValueError, OSError) as exc:
        raise HTTPException(status_code=400, detail=str(exc))


@app.post("/api/export")
def run_export():
    try:
        return export_excel()
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


def db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


@app.get("/", response_class=HTMLResponse)
def home():
    return HTML


@app.get("/api/summary")
def summary():
    with db() as conn:
        total = conn.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
        shifts = conn.execute(
            "SELECT COALESCE(shift_type,'Unknown') AS shift_type, COUNT(*) AS n FROM reports GROUP BY shift_type ORDER BY shift_type"
        ).fetchall()
        dates = conn.execute("SELECT MIN(report_date), MAX(report_date) FROM reports").fetchone()
    return {
        "total": total,
        "shift_types": {row["shift_type"]: row["n"] for row in shifts},
        "from_date": dates[0],
        "to_date": dates[1],
    }


@app.get("/api/reports")
def reports(
    q: str = "",
    from_date: str = "",
    to_date: str = "",
    shift_type: str = "",
    limit: int = Query(200, ge=1, le=500),
):
    clauses = []
    params = []
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
        sql = """
            SELECT r.id, r.doc_no, r.report_date, r.shift_type, r.shift, r.supervisor,
                   r.source_filename,
                   snippet(report_search, 5, '[', ']', ' ... ', 18) AS excerpt
            FROM report_search
            JOIN reports r ON r.id = report_search.rowid
        """ + where + " ORDER BY r.report_date DESC, r.shift LIMIT ?"
    else:
        sql = """
            SELECT r.id, r.doc_no, r.report_date, r.shift_type, r.shift, r.supervisor,
                   r.source_filename, substr(r.activity_report,1,220) AS excerpt
            FROM reports r
        """ + where + " ORDER BY r.report_date DESC, r.shift LIMIT ?"
    params.append(limit)
    with db() as conn:
        rows = conn.execute(sql, params).fetchall()
    return {"count": len(rows), "items": [dict(row) for row in rows]}


@app.get("/api/reports/{report_id}")
def report_detail(report_id: int):
    with db() as conn:
        row = conn.execute(
            """SELECT id, doc_no, report_date, shift_type, shift, supervisor, source_filename,
                      activity_report, shutdown_info, team_info, remarks,
                      operational_json, full_text
               FROM reports WHERE id = ?""",
            (report_id,),
        ).fetchone()
    if row is None:
        raise HTTPException(status_code=404, detail="Report not found")
    return dict(row)


HTML = r"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>CTEP Shift Report Archive</title>
<style>
:root{font-family:Segoe UI,Arial,sans-serif;color:#182433;background:#f3f6fa}
*{box-sizing:border-box}body{margin:0}.top{background:#173b5e;color:white;padding:24px 5vw}
.top h1{margin:0 0 6px;font-size:25px}.top p{margin:0;color:#d7e5f2}.top small{display:block;margin-top:8px;color:#d7e5f2}
main{max-width:1200px;margin:24px auto;padding:0 20px}
.stats{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px}
.stat,.panel{background:white;border:1px solid #dfe6ee;border-radius:12px;padding:16px}
.stat{min-width:150px;flex:1}.stat b{display:block;font-size:25px;color:#173b5e}
.stat span{font-size:12px;color:#65758a}
.filters{display:grid;grid-template-columns:2fr 1fr 1fr 1fr auto;gap:10px;align-items:end}
label{display:block;font-size:12px;color:#536477;margin-bottom:5px}
input,select{width:100%;padding:10px;border:1px solid #cbd5e1;border-radius:8px;background:white;color:#182433}
button{padding:10px 16px;border:0;border-radius:8px;background:#1e5b8f;color:white;cursor:pointer}
button.secondary{background:#e8eef5;color:#24445f}
.meta{display:flex;justify-content:space-between;align-items:center;margin:16px 0 8px;color:#617286;font-size:13px}
table{width:100%;border-collapse:collapse}th,td{text-align:left;padding:12px 10px;border-bottom:1px solid #e8edf2;vertical-align:top}
th{font-size:12px;color:#66788a;background:#f8fafc}td{font-size:14px}
.doc{font-weight:650;color:#14558a;cursor:pointer}.excerpt{font-size:12px;color:#65758a;max-width:560px}
.pill{display:inline-block;padding:3px 8px;border-radius:20px;background:#eaf2f9;color:#244f72;font-size:12px}
.detail{margin-top:16px;display:none}.detail h3{margin:0 0 10px}.detail pre{white-space:pre-wrap;overflow-wrap:anywhere;background:#f7f9fc;padding:14px;border-radius:8px;max-height:520px;overflow:auto;font:13px/1.5 Consolas,monospace}
.empty{text-align:center;padding:30px;color:#738297}
@media(max-width:800px){.filters{grid-template-columns:1fr 1fr}.filters .wide{grid-column:1/-1}table{display:block;overflow-x:auto}}
</style>
</head>
<body>
<header class="top"><h1>CTEP Shift Report Archive</h1><p>Local searchable archive · Day and Night shift reports</p><small>Created by Farzad Meshkipour · Version 1.1.0</small></header>
<main>
<section class="stats">
 <div class="stat"><b id="total">—</b><span>Report versions indexed</span></div>
 <div class="stat"><b id="day">—</b><span>Day shift reports</span></div>
 <div class="stat"><b id="night">—</b><span>Night shift reports</span></div>
 <div class="stat"><b id="range">—</b><span>Report date range</span></div>
</section>
<section class="panel" style="margin-bottom:14px">
 <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap"><button id="export">Export to Excel</button><span id="exportStatus" class="excerpt"></span><button class="secondary" id="settingsToggle">Export Settings</button></div>
 <div id="settingsPanel" style="display:none;margin-top:16px;max-width:760px">
  <div style="display:grid;grid-template-columns:2fr 1fr auto;gap:10px;align-items:end">
   <div><label for="exportPath">Excel file location (.xlsx)</label><input id="exportPath" placeholder="C:\\Reports\\CTEP_Archive.xlsx"></div>
   <div><label for="exportInterval">Update interval (hours)</label><input id="exportInterval" type="number" min="1" max="168" value="3"></div>
   <div><label><input id="autoExport" type="checkbox" checked style="width:auto"> Enable automatic export</label><button id="saveSettings">Save Settings</button></div>
  </div><p id="settingsStatus" class="excerpt"></p>
 </div>
</section>
<section class="panel">
 <div class="filters">
  <div class="wide"><label for="q">Search report content, equipment tags, wells, activities</label><input id="q" placeholder="e.g. F07, GTG, PSV-02-954-27, shutdown"></div>
  <div><label for="from">From date</label><input id="from" type="date"></div>
  <div><label for="to">To date</label><input id="to" type="date"></div>
  <div><label for="shift">Shift</label><select id="shift"><option value="">All shifts</option><option value="Day">Day</option><option value="Night">Night</option></select></div>
  <div><button id="search">Search</button></div>
 </div>
</section>
<div class="meta"><span id="count">Loading reports…</span><button class="secondary" id="clear">Clear filters</button></div>
<section class="panel">
 <table><thead><tr><th>Document No.</th><th>Date</th><th>Shift</th><th>Supervisor</th><th>Matched content</th></tr></thead>
 <tbody id="results"><tr><td colspan="5" class="empty">Loading…</td></tr></tbody></table>
</section>
<section class="panel detail" id="detail"><h3 id="detailTitle"></h3><div id="detailMeta" class="excerpt"></div><pre id="detailText"></pre></section>
</main>
<script>
const $=id=>document.getElementById(id);
async function loadSummary(){
 const s=await fetch('/api/summary').then(r=>r.json());
 $('total').textContent=s.total;
 $('day').textContent=s.shift_types.Day||0;
 $('night').textContent=s.shift_types.Night||0;
 $('range').textContent=(s.from_date||'—')+' → '+(s.to_date||'—');
}
function esc(v){return String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
async function search(){
 const p=new URLSearchParams({q:$('q').value,from_date:$('from').value,to_date:$('to').value,shift_type:$('shift').value});
 const data=await fetch('/api/reports?'+p).then(r=>r.json());
 $('count').textContent=data.count+' report(s) found';
 $('results').innerHTML=data.items.length?data.items.map(x=>'<tr>'+
 '<td><span class="doc" data-id="'+x.id+'">'+esc(x.doc_no)+'</span></td>'+
 '<td>'+esc(x.report_date)+'</td><td><span class="pill">'+esc(x.shift_type)+' (Shift '+esc(x.shift)+')</span></td>'+
 '<td>'+esc(x.supervisor)+'</td><td class="excerpt">'+esc(x.excerpt||'')+'</td></tr>').join(''):'<tr><td colspan="5" class="empty">No matching reports found.</td></tr>';
 document.querySelectorAll('.doc').forEach(el=>el.addEventListener('click',()=>details(el.dataset.id)));
}
async function details(id){
 const r=await fetch('/api/reports/'+id).then(r=>r.json());
 $('detailTitle').textContent=r.doc_no+' — '+r.report_date;
 $('detailMeta').textContent=r.shift_type+' (Shift '+r.shift+') · Supervisor: '+r.supervisor+' · Source: '+r.source_filename;
 $('detailText').textContent=r.full_text;
 $('detail').style.display='block';
 $('detail').scrollIntoView({behavior:'smooth',block:'start'});
}
$('search').addEventListener('click',search);
$('clear').addEventListener('click',()=>{$('q').value='';$('from').value='';$('to').value='';$('shift').value='';search();});
$('q').addEventListener('keydown',e=>{if(e.key==='Enter')search();});
async function initExportSettings(){\n const s=await fetch('/api/settings').then(r=>r.json());\n $('exportPath').value=s.export_path; $('exportInterval').value=s.export_interval_hours; $('autoExport').checked=s.auto_export_enabled;\n}\n$('settingsToggle').addEventListener('click',()=>{$('settingsPanel').style.display=$('settingsPanel').style.display==='none'?'block':'none';});\n$('saveSettings').addEventListener('click',async()=>{\n const r=await fetch('/api/settings',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({export_path:$('exportPath').value,export_interval_hours:Number($('exportInterval').value),auto_export_enabled:$('autoExport').checked})});\n const d=await r.json(); $('settingsStatus').textContent=r.ok?'Settings saved. Automatic export schedule updated.':(d.detail||'Could not save settings');\n});\n$('export').addEventListener('click',async()=>{\n $('exportStatus').textContent='Exporting…'; const r=await fetch('/api/export',{method:'POST'}); const d=await r.json();\n $('exportStatus').textContent=r.ok?('Exported '+d.reports+' reports · '+d.path):('Export failed: '+(d.detail||'Unknown error'));\n});\nloadSummary().then(search); initExportSettings();
</script>
</body></html>"""
