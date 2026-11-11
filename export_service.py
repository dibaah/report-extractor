from __future__ import annotations

import json
import os
import sqlite3
import threading
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = Path(__import__('sys').executable).parent if getattr(__import__('sys'), 'frozen', False) else Path(__file__).parent
DB_PATH = BASE / "data" / "ctep_archive.db"
SETTINGS_PATH = BASE / "archive_settings.json"
APP_VERSION = "1.1.0"
APP_AUTHOR = "Farzad Meshkipour"
DEFAULT_SETTINGS = {"export_path": str(BASE / "exports" / "CTEP_Shift_Report_Archive.xlsx"), "export_interval_hours": 3, "auto_export_enabled": True}
_stop = threading.Event()
_thread = None


def load_settings():
    settings = DEFAULT_SETTINGS.copy()
    try:
        settings.update(json.loads(SETTINGS_PATH.read_text(encoding="utf-8")))
    except (OSError, ValueError):
        pass
    return settings


def save_settings(export_path: str, interval_hours: int, enabled: bool):
    path = str(Path(export_path).expanduser())
    if not path.lower().endswith(".xlsx"):
        raise ValueError("Export file must have .xlsx extension")
    if not 1 <= int(interval_hours) <= 168:
        raise ValueError("Interval must be between 1 and 168 hours")
    settings = {"export_path": path, "export_interval_hours": int(interval_hours), "auto_export_enabled": bool(enabled)}
    SETTINGS_PATH.write_text(json.dumps(settings, indent=2), encoding="utf-8")
    return settings


def export_excel():
    settings = load_settings()
    target = Path(settings["export_path"])
    target.parent.mkdir(parents=True, exist_ok=True)
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute("""SELECT doc_no, report_date, shift_type, shift, supervisor, source_filename,
            activity_report, shutdown_info, team_info, remarks, full_text, imported_at
            FROM reports ORDER BY report_date DESC, shift_type, doc_no""").fetchall()
    wb = Workbook()
    ws = wb.active
    ws.title = "CTEP Reports"
    headers = ["Document No.", "Report Date", "Shift Type", "Excel Shift", "Supervisor", "Source File",
               "Activity Report", "Shutdown Info", "Team Info", "Remarks", "Full Text", "Imported At"]
    ws.append(headers)
    for row in rows:
        ws.append([row[h.lower().replace(" ", "_").replace(".", "")] if False else None for h in []]) if False else None
        ws.append([row[k] for k in ("doc_no", "report_date", "shift_type", "shift", "supervisor", "source_filename",
                                    "activity_report", "shutdown_info", "team_info", "remarks", "full_text", "imported_at")])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="173B5E")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    widths = [24, 14, 14, 12, 22, 28, 45, 35, 35, 35, 80, 22]
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = cell.alignment.copy(wrap_text=True, vertical="top")
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
    return {"path": str(target), "reports": len(rows), "exported_at": datetime.now().astimezone().isoformat(timespec="seconds")}


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
        _stop.wait(max(30, wait_seconds))


def start_scheduler():
    global _thread
    if _thread and _thread.is_alive():
        return
    _stop.clear()
    _thread = threading.Thread(target=_scheduler, name="ctep-excel-export", daemon=True)
    _thread.start()


def stop_scheduler():
    _stop.set()
