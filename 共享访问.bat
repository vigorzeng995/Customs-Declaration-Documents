@echo off
setlocal enabledelayedexpansion
title Customs Doc Generator - SHARED MODE (LAN) - close this window to stop
cd /d "%~dp0"
set PORT=8765
set ODIN_HOST=0.0.0.0
set RULENAME=CustomsDoc8765

rem NOTE: keep this file pure ASCII. cmd reads .bat in the OEM codepage,
rem and non-ASCII bytes get garbled into broken commands.

echo ============================================================
echo   Customs Doc Generator - SHARED MODE
echo   Other people on the same network can open this tool.
echo ============================================================
echo.

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
echo [1/3] Python OK: %PY%
echo.

rem ---- 2. make sure the port is reachable from other machines ----
rem The current network is often "Public", where Windows blocks inbound by
rem default. Adding the rule needs admin rights, so we ask for it once.
netsh advfirewall firewall show rule name="%RULENAME%" >nul 2>nul
if errorlevel 1 (
    echo [2/3] First run: allowing port %PORT% in Windows Firewall.
    echo       A permission prompt will pop up - please click YES.
    powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Start-Process -FilePath netsh -Verb RunAs -Wait -ArgumentList 'advfirewall firewall add rule name=%RULENAME% dir=in action=allow protocol=TCP localport=%PORT% profile=any' } catch { exit 1 }" >nul 2>nul
    netsh advfirewall firewall show rule name="%RULENAME%" >nul 2>nul
    if errorlevel 1 (
        echo       [WARN] Rule NOT added. This PC still works, but others
        echo              may not connect. To fix it, run this once in an
        echo              ADMIN terminal:
        echo              netsh advfirewall firewall add rule name=%RULENAME% dir=in action=allow protocol=TCP localport=%PORT% profile=any
    ) else (
        echo       Firewall rule added.
    )
) else (
    echo [2/3] Firewall rule already in place.
)
echo.

rem ---- 3. already running? do not start a second instance ----
netstat -ano | findstr /c:":%PORT%" | findstr /c:"LISTENING" >nul
if not errorlevel 1 (
    echo [3/3] The service is ALREADY RUNNING on port %PORT%.
    echo       No need to start it twice - just share the address below.
    echo.
    call :showurls
    start "" http://127.0.0.1:%PORT%
    echo This window can be closed. The running service is not affected.
    echo.
    pause
    exit /b 0
)

echo [3/3] Starting the service ...
echo.
call :showurls
echo   Keep this window open. Closing it stops sharing.
echo   Minimising is fine. If it crashes it restarts by itself.
echo ============================================================
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
rem use ping, not timeout: on machines where GNU coreutils is on PATH, `timeout /t` fails
ping -n 3 127.0.0.1 >nul 2>nul
goto loop

:showurls
echo   Send this address to your colleagues:
set "LANIP="
for /f "usebackq delims=" %%A in (`"%PY%" app.py --lan`) do if not defined LANIP set "LANIP=%%A"
if defined LANIP (
    echo       http://!LANIP!:%PORT%
) else (
    echo       [WARN] Could not detect the network address.
    echo              Run "ipconfig" in a terminal to look it up.
)
echo.
echo   On this computer: http://127.0.0.1:%PORT%
echo.
exit /b 0
