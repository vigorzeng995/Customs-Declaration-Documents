@echo off
setlocal enabledelayedexpansion
title Shared mode - LAN only - close this window to stop
cd /d "%~dp0"
set PORT=8765
set ODIN_HOST=0.0.0.0

rem NOTE: keep this file pure ASCII. cmd reads .bat in the OEM codepage, and
rem non-ASCII bytes get garbled into broken commands.

rem ---- 1. find a python that can really import our dependencies ----
rem The bare name "python" often points at the Microsoft Store stub, which
rem starts nothing and just exits with an error. So we test by importing.
set "PY="
for /d %%D in ("%USERPROFILE%\.workbuddy\binaries\python\versions\*" "%LOCALAPPDATA%\Programs\WorkBuddy\resources\vendor\python") do (
    if not defined PY if exist "%%D\python.exe" (
        "%%D\python.exe" -c "import flask,openpyxl,PIL" >nul 2>nul && set "PY=%%D\python.exe"
    )
)
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*" "C:\Program Files\Python3*") do (
    if not defined PY if exist "%%D\python.exe" (
        "%%D\python.exe" -c "import flask,openpyxl,PIL" >nul 2>nul && set "PY=%%D\python.exe"
    )
)
if not defined PY (
    python -c "import flask,openpyxl,PIL" >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo.
    echo [ERROR] No usable Python found ^(needs flask, openpyxl and Pillow^).
    echo.
    echo   1. Install Python 3 from https://www.python.org/downloads/
    echo      and tick "Add python.exe to PATH" during setup.
    echo   2. Open this folder in a terminal and run:
    echo        pip install -r requirements.txt
    echo   3. Run this file again.
    echo.
    pause
    exit /b 1
)
echo Using Python: %PY%
echo.

rem ---- 2. already listening? do not start a second instance ----
netstat -ano | findstr /c:":%PORT%" | findstr /c:"LISTENING" >nul
if not errorlevel 1 (
    echo Port %PORT% is already in use - probably this app is already running.
    echo Open the URL printed in that window.
    echo.
    netstat -ano | findstr /c:":%PORT%" | findstr /c:"LISTENING"
    pause
    exit /b 0
)

echo ============================================================
echo   Customs Doc Generator - SHARED MODE
echo.
echo   On this PC : http://127.0.0.1:%PORT%
echo   For others : see the addresses printed below
echo.
echo   Only devices on the same network can reach it.
echo   Windows will ask about the firewall once - allow "Private network".
echo   Keep this window open. Minimising is fine, closing stops sharing.
echo ============================================================
echo.
echo Your LAN addresses:
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /c:"IPv4"') do for /f "tokens=* delims= " %%B in ("%%A") do echo    http://%%B:%PORT%
echo.

set OPEN_BROWSER=1
set FAILS=0

:loop
"%PY%" app.py
set OPEN_BROWSER=
if errorlevel 1 (set /a FAILS+=1) else (set FAILS=0)
if !FAILS! GEQ 3 (
    echo.
    echo [ERROR] The service keeps crashing ^(3 times in a row^).
    echo         Read the error above, fix it, then run this file again.
    echo.
    pause
    exit /b 1
)
echo.
echo service stopped, restarting in 2 seconds ...  ^(close this window to stop^)
ping -n 3 127.0.0.1 >nul 2>nul
goto loop
