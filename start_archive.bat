@echo off
setlocal
cd /d "%~dp0"
if exist "%~dp0CTEP-Shift-Report-Archive.exe" (
  start "" "%~dp0CTEP-Shift-Report-Archive.exe"
  exit /b 0
)
py -3 run_archive.py
if errorlevel 1 (
  echo.
  echo Dashboard failed to start. Check Python dependencies and Edge installation.
  pause
)
