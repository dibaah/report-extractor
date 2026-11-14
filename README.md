# CTEP Shift Report Archive

Local searchable archive for CTEP Day/Night shift reports.

**Created by:** Farzad Meshkipour  
**Version:** 1.1.0

## Scope and safeguards
- CTEP Shift Reports only (Day/Night); Biweekly reports are excluded.
- Original Outlook emails and Excel attachments are not modified.
- Day/Night is identified from the document number suffix (D/N); Excel Shift A/B is retained separately.
- SQLite stores searchable extracted report content and structured data.
- Identical report versions are skipped during import.

## Run locally
Run `start_archive.bat`, then open http://127.0.0.1:8765.

The dashboard binds to localhost only.

## Excel export
- Use **Export to Excel** for an immediate export.
- Open **Export Settings** to set the full .xlsx destination, interval in hours (1–168), and enable/disable automatic export.
- Default interval is 3 hours. The same destination file is replaced on each successful export.
- The scheduled exporter runs while the dashboard application is running.

## Import reports
Both a **folder** and a **single Excel file** are supported (folders are scanned recursively, including subfolders):
```powershell
py -3 .\extractor.py                                  # default: whole .\outlook_import folder + subfolders
py -3 .\extractor.py "C:\path\to\folder"              # a folder (all .xlsx/.xlsm inside, recursively)
py -3 .\extractor.py "C:\path\to\CWRS-report.xlsx"    # a single Excel file outside any folder
py -3 .\extractor.py "C:\folder" "C:\one\report.xlsx" # any mix of folders and files at once
```
Duplicate paths (e.g. passing a file that is also inside a passed folder) are ignored automatically.

## Automatic Outlook import (every 3 hours)
Two files handle the fully automatic flow:

- `auto_outlook_import.ps1` – one run = connects to Outlook, saves new `CWRS-*.xlsx/.xlsm` attachments from *Inbox ▸ 1. Report ▸ C. Reports Archive* into `outlook_import\`, then runs `extractor.py` to import them into the archive DB. Already-imported reports are skipped automatically, so repeated runs are safe. Log: `logs\auto_outlook_import.log`.
- `install_autorun_task.bat` – creates a Windows Scheduled Task (**CTEP Outlook Auto Import**) that runs the script **every 3 hours** under your user account. Double-click it once to install.

Useful commands:
```powershell
schtasks /Run   /TN "CTEP Outlook Auto Import"   # test run right now
schtasks /Query /TN "CTEP Outlook Auto Import" /V # next/last run times
schtasks /Delete /TN "CTEP Outlook Auto Import" /F # uninstall
```

Requirements: Windows + desktop Outlook installed and logged in (the task only works while you're signed in; Outlook may show an antivirus/permission prompt on first attachment access). To change mailbox or folder, edit `$mailboxName` / `$folderPath` at the top of `auto_outlook_import.ps1`.

## Build
The GitHub Actions workflow at `.github/workflows/build-windows.yml` builds a Windows portable application folder and packages it as a ZIP when manually triggered or when a `v*` tag is pushed. Add this project to a GitHub repository to use the workflow.

## Current sample data
- CWRS-20261106N
- CWRS-20261107D
- CWRS-20261107N
