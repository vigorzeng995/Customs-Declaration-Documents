@echo off
setlocal enabledelayedexpansion
title Customs Doc Generator - http://127.0.0.1:8765 - close this window to stop
cd /d "%~dp0"
set PORT=8765

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

rem ---- 2. already running? just open the page, do not start a 2nd instance ----
rem match any bind address (127.0.0.1 or 0.0.0.0), so shared mode is detected too
netstat -ano | findstr /c:":%PORT%" | findstr /c:"LISTENING" >nul
if not errorlevel 1 (
    echo Service already running on port %PORT%. Opening the page ...
    echo If the page ever stops working, close the window that runs the service,
    echo then run this file again to get a fresh one.
    start "" http://127.0.0.1:%PORT%
    ping -n 3 127.0.0.1 >nul 2>nul
    exit /b 0
)

echo ============================================
echo   Customs Doc Generator
echo   URL : http://127.0.0.1:%PORT%
echo   Keep this window open. Close it to stop.
echo   Crashed? It restarts itself after 2 seconds.
echo ============================================
echo.

rem ---- 3. open the browser only on the first start ----
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
echo service stopped, restarting in 2 seconds ...  ^(Ctrl+C or close this window to stop^)
rem use ping, not timeout: on machines where GNU coreutils is on PATH, `timeout /t` fails
ping -n 3 127.0.0.1 >nul 2>nul
goto loop
