@echo off
setlocal
cd /d "%~dp0"
set "PORT=8765"

echo [+] Checking port %PORT% ...
set "BUSY="
for /f "tokens=5" %%a in ('netstat -ano ^| findstr /C:":%PORT% " ^| findstr /C:"LISTENING"') do (
    set "BUSY=1"
    echo [!] Port %PORT% is busy - killing PID %%a
    taskkill /F /T /PID %%a >nul 2>&1
)

if defined BUSY (
    timeout /t 1 /nobreak >nul
    netstat -ano | findstr /C:":%PORT% " | findstr /C:"LISTENING" >nul 2>&1
    if errorlevel 1 (
        echo [+] Port %PORT% is now free.
    ) else (
        echo [X] Port %PORT% is still busy. Run this file as Administrator.
        pause
        exit /b 1
    )
) else (
    echo [+] Port %PORT% is already free. Nothing to kill.
)

echo [+] Starting server on http://127.0.0.1:%PORT% ...
py -3 -m uvicorn app:app --host 127.0.0.1 --port %PORT%
pause
