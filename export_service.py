from __future__ import annotations

import json
import os
import re
import sqlite3
import threading
from datetime import datetime
from pathlib import Path
from contextlib import closing

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

FROZEN = getattr(__import__("sys"), "frozen", False)
APP_DIR = Path(__import__("sys").executable).parent if FROZEN else Path(__file__).parent
USER_DATA_DIR = (
    Path(os.environ.get("LOCALAPPDATA", str(APP_DIR))) / "CTEP Shift Report Archive"
    if FROZEN else APP_DIR
)
DB_PATH = USER_DATA_DIR / "data" / "ctep_archive.db"
SETTINGS_PATH = USER_DATA_DIR / "archive_settings.json"
APP_VERSION = "1.2.0"
APP_AUTHOR = "Farzad Meshkipour"
DEFAULT_SETTINGS = {
    "export_path": str(USER_DATA_DIR / "exports" / "CTEP_Shift_Report_Archive.xlsx"),
    "import_path": str(USER_DATA_DIR / "outlook_import"),
    "archive_path": str(USER_DATA_DIR / "processed_reports"),
    "export_interval_hours": 3,
    "auto_export_enabled": True,
}
_stop = threading.Event()
_wake = threading.Event()
_thread = None


def load_settings():
    settings = DEFAULT_SETTINGS.copy()
    try:
        settings.update(json.loads(SETTINGS_PATH.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    return settings


def save_settings(export_path: str, import_path: str, archive_path: str, interval_hours: int, enabled: bool):
    export = Path(export_path).expanduser()
    source = Path(import_path).expanduser()
    archive = Path(archive_path).expanduser()
    if export.suffix.lower() != ".xlsx":
        raise ValueError("Export file must have .xlsx extension")
    if not str(source).strip():
        raise ValueError("Please set the Outlook Import folder")
    if not str(archive).strip():
        raise ValueError("Please set the processed/archive folder")
    if not 1 <= int(interval_hours) <= 168:
        raise ValueError("Interval must be between 1 and 168 hours")
    settings = {
        "export_path": str(export.resolve()),
        "import_path": str(source.resolve()),
        "archive_path": str(archive.resolve()),
        "export_interval_hours": int(interval_hours),
        "auto_export_enabled": bool(enabled),
    }
    SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() == archive.resolve() or source.resolve() in archive.resolve().parents:
        raise ValueError("Processed/archive folder must be outside the Outlook Import folder")
    source.mkdir(parents=True, exist_ok=True)
    archive.mkdir(parents=True, exist_ok=True)
    SETTINGS_PATH.write_text(json.dumps(settings, indent=2), encoding="utf-8")
    _wake.set()
    return settings


def export_excel():
    settings = load_settings()
    target = Path(settings["export_path"])
    target.parent.mkdir(parents=True, exist_ok=True)
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    with closing(sqlite3.connect(DB_PATH)) as conn, conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute("""
            SELECT doc_no, report_date, shift_type, shift, supervisor,
                   activity_report, source_path
            FROM reports ORDER BY report_date DESC, shift_type, doc_no
        """).fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "CTEP Reports"
    headers = ["Document No.", "Report Date", "Shift Type", "Excel Shift",
               "Supervisor", "Activity Report", "Original Excel"]
    ws.append(headers)
    for row in rows:
        activity = re.sub(r"\s+", " ", row["activity_report"] or "").strip()
        ws.append([row["doc_no"], row["report_date"], row["shift_type"], row["shift"],
                   row["supervisor"], activity, "Open source Excel"])
        link = ws.cell(ws.max_row, 7)
        source_path = row["source_path"] or ""
        if source_path and Path(source_path).is_file():
            link.hyperlink = Path(source_path).resolve().as_uri()
            link.style = "Hyperlink"
        else:
            link.value = "Source unavailable"

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="173B5E")
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    widths = [25, 15, 14, 13, 23, 105, 20]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=(cell.column == 6))
        ws.row_dimensions[row[0].row].height = 30
    ws.sheet_view.showGridLines = False

    info = wb.create_sheet("Archive Info")
    info.append(["CTEP Shift Report Archive"])
    info.append(["Created by", APP_AUTHOR])
    info.append(["Version", APP_VERSION])
    info.append(["Exported at", datetime.now().astimezone().isoformat(timespec="seconds")])
    info.append(["Report versions", len(rows)])
    info.column_dimensions["A"].width = 22
    info.column_dimensions["B"].width = 55

    temp = target.with_name(target.stem + ".tmp.xlsx")
    wb.save(temp)
    os.replace(temp, target)
    return {"path": str(target), "reports": len(rows),
            "exported_at": datetime.now().astimezone().isoformat(timespec="seconds")}


def _scheduler():
    while not _stop.is_set():
        settings = load_settings()
        if settings.get("auto_export_enabled", True):
            try:
                export_excel()
            except Exception as exc:
                print(f"[scheduled Excel export] {exc}", flush=True)
            wait_seconds = int(settings.get("export_interval_hours", 3)) * 3600
        else:
            wait_seconds = 60
        _wake.wait(max(30, wait_seconds))
        _wake.clear()


def start_scheduler():
    global _thread
    if _thread and _thread.is_alive():
        return
    _stop.clear()
    _wake.clear()
    _thread = threading.Thread(target=_scheduler, name="ctep-excel-export", daemon=True)
    _thread.start()


def stop_scheduler():
    _stop.set()
    _wake.set()
