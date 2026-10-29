@echo off
setlocal enabledelayedexpansion
title Shared mode - LAN only - close this window to stop
cd /d "%~dp0"
set PORT=8765
set ODIN_HOST=0.0.0.0

rem NOTE: keep this file pure ASCII. cmd reads .bat in the OEM codepage, and
rem non-ASCII bytes get garbled into broken commands.

where python >nul 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] Python not found. Install Python 3 and tick "Add python.exe to PATH".
    echo.
    pause
    exit /b 1
)

rem ---- already listening? do not start a second instance ----
netstat -ano | findstr /c:":%PORT%" | findstr /c:"LISTENING" >nul
if not errorlevel 1 (
    echo Port %PORT% is already in use - probably this app is already running.
    echo Open the URL printed in that window.
    echo.
    netstat -ano | findstr /c:":%PORT%" | findstr /c:"LISTENING"
    ping -n 4 127.0.0.1 >nul 2>nul
    exit /b 0
)

echo ============================================================
echo   Customs Doc Generator - SHARED MODE
echo.
echo   On this PC : http://127.0.0.1:%PORT%
echo   For others : http://YOUR-LAN-IP:%PORT%   ^(see below^)
echo.
echo   Only devices on the same network can reach it.
echo   Windows will ask about the firewall once - allow "Private network".
echo   Keep this window open. Close it to stop sharing.
echo ============================================================
echo.
echo Your LAN addresses:
for /f "tokens=2 delims=:" %%A in ('ipconfig ^| findstr /c:"IPv4"') do echo    http://%%A:%PORT%
echo.

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
echo service stopped, restarting in 2 seconds ...  ^(close this window to stop^)
ping -n 3 127.0.0.1 >nul 2>nul
goto loop
