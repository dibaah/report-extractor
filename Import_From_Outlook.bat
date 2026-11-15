@echo off
setlocal
cd /d "%~dp0"

echo ============================================================
echo CTEP Outlook Attachment Export
echo Saves CWRS/CSR Excel attachments into outlook_import.
echo This script does NOT import reports into the SQLite database.
echo ============================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0export_outlook_ctep.ps1"
set "EXIT_CODE=%ERRORLEVEL%"

echo.
if "%EXIT_CODE%"=="0" (
    echo Outlook attachment export completed.
    echo Check the outlook_import folder and logs folder.
) else (
    echo Outlook attachment export failed. Check logs\outlook_attachment_export.log.
)
echo.
pause
exit /b %EXIT_CODE%
