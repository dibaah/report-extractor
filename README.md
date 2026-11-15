# CTEP Shift Report Archive

Local searchable archive for CTEP Day/Night shift reports.

**Created by:** Farzad Meshkipour  
**Version:** 1.2.0

## Dashboard

Run `start_archive.bat`. The dashboard opens in a dedicated Microsoft Edge app window at localhost. Closing that window stops the local API server. The app does not connect to Outlook directly.

## Manual import workflow

1. Save Outlook Excel attachments into the configured Outlook Import folder.
2. Click **Import Outlook Folder** in the dashboard.
3. The importer scans all subfolders for `.xlsx` and `.xlsm` reports.
4. SQLite duplicate protection skips identical report versions.
5. Successfully imported reports and confirmed duplicates are moved to the configured processed folder, retaining their subfolder structure. Files that fail parsing/import remain in the import folder.
6. The database records the processed file path so the Excel export can link to the original workbook.

The import action does not delete source files. Do not configure the processed folder inside the import folder.

## Settings

The dashboard Settings panel controls:
- Outlook Import folder (recursive scan)
- Processed reports archive folder
- Excel export destination
- Automatic Excel export interval (1–168 hours), with enable/disable

In development mode, settings and the SQLite database stay in the repository's local working folder. In the packaged Windows application, runtime data (SQLite database, settings, Outlook Import folder, processed reports and default Excel export) is stored under `%LOCALAPPDATA%\CTEP Shift Report Archive`, so updating the executable does not replace the user's archive. The SQLite database is excluded from Git to prevent operational report data from being synchronized to GitHub.

## Excel export

The exported workbook contains:
- Document No.
- Report Date
- Shift Type
- Excel Shift
- Supervisor
- Activity Report
- Original Excel (clickable hyperlink when the source file is available)

Shutdown Info, Team Info, Remarks, Full Text and Imported At are intentionally excluded. Activity Report is flattened to reduce unnecessary line breaks and given a wider column.

## Command-line importer

```powershell
py -3 .\extractor.py
py -3 .\extractor.py "C:\path\to\folder"
py -3 .\extractor.py "C:\path\to\report.xlsx"
```

Folders are scanned recursively. Only `.xlsx` and `.xlsm` are supported. Legacy `.xls` files must be converted first.

## Build

GitHub Actions workflow: `.github/workflows/main.yml`. Run it manually from Actions or push a tag beginning with `v` (for example `v1.2.0`). It builds a Windows portable application folder and publishes a ZIP artifact.

## Scope and safeguards

- CTEP Shift Reports only (Day/Night); Biweekly reports are excluded.
- Original Outlook emails are not modified.
- Day/Night is identified from document number suffix (D/N); Excel Shift A/B is retained separately.
- Identical report versions are skipped.
- Files are moved only after a successful database commit; failed files remain in the import folder.
