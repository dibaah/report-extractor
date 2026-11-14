from __future__ import annotations

import argparse
import hashlib
import json
import re
import sqlite3
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

DEFAULT_DB = Path(__file__).parent / "data" / "ctep_archive.db"
DEFAULT_INPUT = Path(__file__).parent / "outlook_import"
EXCEL_EXTS = {".xlsx", ".xlsm"}


def clean(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def unique_row_values(row):
    result = []
    previous = None
    for cell in row:
        value = clean(cell.value)
        if value and value != previous:
            result.append(value)
        previous = value
    return result


def parse_report(path: Path):
    wb = load_workbook(path, data_only=True, read_only=True)
    if "CTEP Shift Report" not in wb.sheetnames:
        raise ValueError(f"Missing CTEP Shift Report sheet: {path.name}")
    ws = wb["CTEP Shift Report"]

    doc_no = clean(ws["G3"].value)
    if not doc_no:
        for row in ws.iter_rows(min_row=1, max_row=5):
            for cell in row:
                match = re.search(r"CWRS-\d{8}[DN]", clean(cell.value))
                if match:
                    doc_no = match.group(0)
                    break
            if doc_no:
                break

    shift_text = clean(ws["C3"].value)
    shift_match = re.search(r"Shift:\s*(.+)", shift_text, re.I)
    shift = shift_match.group(1).strip() if shift_match else ""
    shift_type = "Day" if doc_no.endswith("D") else "Night" if doc_no.endswith("N") else ""

    supervisor_text = clean(ws["A3"].value)
    supervisor = re.sub(r"^Supervisor:\s*", "", supervisor_text, flags=re.I).strip()

    date_text = clean(ws["I3"].value)
    date_match = re.search(r"Date:\s*(.+)", date_text, re.I)
    date_value = date_match.group(1).strip() if date_match else ""
    report_date = ""
    if date_value:
        for fmt in ("%Y-%b-%d", "%Y-%B-%d", "%Y-%m-%d"):
            try:
                report_date = datetime.strptime(date_value, fmt).date().isoformat()
                break
            except ValueError:
                pass

    activity_report = clean(ws["F6"].value)
    if not activity_report:
        for row in ws.iter_rows(min_row=6, max_row=min(ws.max_row, 15)):
            for cell in row:
                value = clean(cell.value)
                if len(value) > 100:
                    activity_report = value
                    break
            if activity_report:
                break

    rows = []
    all_text = []
    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        values = unique_row_values(row)
        if values:
            row_text = " | ".join(values)
            rows.append({"row": row_idx, "values": values, "text": row_text})
            all_text.extend(values)

    # Retain a structured snapshot of the operational table without merged-cell duplicates.
    operational_rows = [r for r in rows if 5 <= r["row"] <= 34]
    shutdown_rows = [r for r in rows if 33 <= r["row"] <= 34]
    member_rows = [r for r in rows if 35 <= r["row"] <= 37]
    remark_rows = [r for r in rows if r["row"] >= 38 and r["row"] <= 40]

    full_text = "\n".join(r["text"] for r in rows)
    content_hash = hashlib.sha256(full_text.encode("utf-8")).hexdigest()

    report = {
        "doc_no": doc_no,
        "report_date": report_date,
        "shift": shift,
        "shift_type": shift_type,
        "supervisor": supervisor,
        "source_filename": path.name,
        "source_path": str(path),
        "activity_report": activity_report,
        "shutdown_info": "\n".join(r["text"] for r in shutdown_rows),
        "team_info": "\n".join(r["text"] for r in member_rows),
        "remarks": "\n".join(r["text"] for r in remark_rows),
        "operational_rows": operational_rows,
        "all_rows": rows,
        "full_text": full_text,
        "content_hash": content_hash,
    }
    wb.close()
    return report



def collect_excel_files(paths):
    """Accept single Excel files AND folders; scan folders recursively (e.g. year subfolders).

    Examples:
        py -3 extractor.py outlook_import                 -> folder + all subfolders
        py -3 extractor.py outlook_import/CWRS-xxx.xlsx   -> a single loose file
        py -3 extractor.py folder1 file1.xlsx file2.xlsx  -> any mix of both
    """
    files, seen = [], set()
    for raw in paths:
        p = Path(raw).expanduser().resolve()
        print(f"[i] Scanning: {p}")
        if p.is_dir():
            candidates = sorted(p.rglob("*"))   # شامل همه‌ی زیرپوشه‌ها (مثلاً 2018..2026)
        elif p.is_file():
            candidates = [p]
        else:
            print(f"[!] Path not found: {p}")
            continue
        stats, xls_count = {}, 0
        for f in candidates:
            if not f.is_file():
                continue
            ext = f.suffix.lower()
            folder = str(f.parent.relative_to(p)) if p.is_dir() else "."
            stats.setdefault(folder, {})
            stats[folder][ext] = stats[folder].get(ext, 0) + 1
            if ext == ".xls":
                xls_count += 1
            if ext not in EXCEL_EXTS or f.name.startswith("~$"):
                continue
            key = str(f)
            if key not in seen:
                seen.add(key)
                files.append(f)
        for folder, exts in sorted(stats.items()):
            print(f"    [{folder}] " + ", ".join(f"{n} x {e}" for e, n in sorted(exts.items())))
        if xls_count:
            print(f"[!] {xls_count} old .xls file(s) found - convert to .xlsx first (Excel > Save As).")
    return sorted(set_by_path(files), key=lambda x: str(x).lower())


def set_by_path(files):
    """De-duplicate the final list by resolved path (case-insensitive on Windows)."""
    unique, seen = [], set()
    for f in files:
        key = str(Path(f).resolve()).lower()
        if key not in seen:
            seen.add(key)
            unique.append(f)
    return unique

def initialize_db(conn):
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS reports (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        doc_no TEXT NOT NULL,
        report_date TEXT,
        shift TEXT,
        shift_type TEXT,
        supervisor TEXT,
        source_filename TEXT NOT NULL,
        source_path TEXT NOT NULL,
        content_hash TEXT NOT NULL,
        activity_report TEXT,
        shutdown_info TEXT,
        team_info TEXT,
        remarks TEXT,
        operational_json TEXT NOT NULL,
        full_text TEXT NOT NULL,
        imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(doc_no, content_hash)
    );
    CREATE INDEX IF NOT EXISTS idx_reports_date ON reports(report_date);
    CREATE INDEX IF NOT EXISTS idx_reports_shift ON reports(shift);
    CREATE INDEX IF NOT EXISTS idx_reports_doc ON reports(doc_no);
    CREATE VIRTUAL TABLE IF NOT EXISTS report_search USING fts5(
        doc_no, report_date, shift, supervisor, activity_report, full_text,
        content='reports', content_rowid='id'
    );
    CREATE TRIGGER IF NOT EXISTS reports_ai AFTER INSERT ON reports BEGIN
        INSERT INTO report_search(rowid, doc_no, report_date, shift, supervisor, activity_report, full_text)
        VALUES (new.id, new.doc_no, new.report_date, new.shift, new.supervisor, new.activity_report, new.full_text);
    END;
    CREATE TRIGGER IF NOT EXISTS reports_ad AFTER DELETE ON reports BEGIN
        INSERT INTO report_search(report_search, rowid, doc_no, report_date, shift, supervisor, activity_report, full_text)
        VALUES ('delete', old.id, old.doc_no, old.report_date, old.shift, old.supervisor, old.activity_report, old.full_text);
    END;
    """)
    columns = {row[1] for row in conn.execute("PRAGMA table_info(reports)")}
    if "shift_type" not in columns:
        conn.execute("ALTER TABLE reports ADD COLUMN shift_type TEXT")
    conn.execute("""
        UPDATE reports
        SET shift_type = CASE
            WHEN doc_no LIKE '%D' THEN 'Day'
            WHEN doc_no LIKE '%N' THEN 'Night'
            ELSE ''
        END
        WHERE shift_type IS NULL
    """)


def import_files(files, db_path):
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(db_path)
    initialize_db(conn)
    imported = 0
    skipped = []
    for file_path in files:
        try:
            report = parse_report(Path(file_path))
            if not report["doc_no"]:
                raise ValueError("Could not identify document number")
            cur = conn.execute("""
                INSERT OR IGNORE INTO reports
                (doc_no, report_date, shift, shift_type, supervisor, source_filename, source_path,
                 content_hash, activity_report, shutdown_info, team_info, remarks,
                 operational_json, full_text)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report["doc_no"], report["report_date"], report["shift"], report["shift_type"], report["supervisor"],
                report["source_filename"], report["source_path"], report["content_hash"],
                report["activity_report"], report["shutdown_info"], report["team_info"],
                report["remarks"], json.dumps(report["operational_rows"], ensure_ascii=False),
                report["full_text"]
            ))
            if cur.rowcount:
                imported += 1
            else:
                skipped.append(f"{Path(file_path).name}: already imported")
        except Exception as exc:
            skipped.append(f"{Path(file_path).name}: {exc}")
    conn.commit()
    total = conn.execute("SELECT COUNT(*) FROM reports").fetchone()[0]
    print(f"Database: {db_path}")
    print(f"New reports imported: {imported}")
    print(f"Total report versions: {total}")
    for row in conn.execute("SELECT doc_no, report_date, shift_type, shift, supervisor FROM reports ORDER BY report_date, shift_type"):
        print(" | ".join(str(x or "") for x in row))
    if skipped:
        print("Skipped / errors:")
        for item in skipped:
            print(" - " + item)
    conn.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Import CTEP shift-report Excel files into a local SQLite archive.")
    parser.add_argument("paths", nargs="*", default=[str(DEFAULT_INPUT)],
                        help="Excel files and/or folders (folders are scanned recursively)")
    parser.add_argument("--db", default=str(DEFAULT_DB), help="SQLite database path")
    args = parser.parse_args()
    files = collect_excel_files(args.paths)
    if not files:
        print("No .xlsx/.xlsm files found - nothing to import.")
        raise SystemExit(1)
    print(f"Found {len(files)} Excel file(s) incl. subfolders.")
    import_files(files, Path(args.db))
