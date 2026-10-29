@echo off
cd /d "%~dp0"
set OPEN_BROWSER=1
echo ============================================
echo  Customs Doc Generator  /  Baoguan Tool
echo  URL: http://127.0.0.1:8765
echo  Keep this window open. Close it to stop.
echo ============================================
python app.py
pause
