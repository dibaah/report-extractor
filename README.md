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
```powershell
py -3 .\extractor.py "C:\path\to\CWRS-report.xlsx"
```

## Build
The GitHub Actions workflow at `.github/workflows/build-windows.yml` builds a Windows portable application folder and packages it as a ZIP when manually triggered or when a `v*` tag is pushed. Add this project to a GitHub repository to use the workflow.

## Current sample data
- CWRS-20261106N
- CWRS-20261107D
- CWRS-20261107N
