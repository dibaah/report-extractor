@echo off
REM Installs a Windows Scheduled Task that runs auto_outlook_import.ps1
REM every 3 hours.
REM Run this file normally (no admin needed - task is created for current user).

setlocal
cd /d "%~dp0"
set "TASKNAME=CTEP Outlook Auto Import"

schtasks /Create ^
  /TN "%TASKNAME%" ^
  /SC HOURLY /MO 3 ^
  /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"%~dp0auto_outlook_import.ps1\"" ^
  /ST 00:00 /F

if errorlevel 1 (
  echo [X] Failed to create task. Try "Run as administrator".
  pause
  exit /b 1
)

echo.
echo [+] Task "%TASKNAME%" created - will check Outlook every 3 hours.
echo    To run it right now for testing:
echo      schtasks /Run /TN "%TASKNAME%"
echo    To remove it:
echo      schtasks /Delete /TN "%TASKNAME%" /F
pause
