@echo off
setlocal enabledelayedexpansion
title Customs Doc Generator - http://127.0.0.1:8765 - close this window to stop
cd /d "%~dp0"
set PORT=8765

rem ---- 0. python available? ----
where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] Python not found. Install Python 3 and tick "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

rem ---- 1. already running? just open the page, do not start a 2nd instance ----
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

rem ---- 2. open the browser only on the first start ----
set OPEN_BROWSER=1
set FAILS=0

:loop
python app.py
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
