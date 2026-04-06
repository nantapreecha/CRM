@echo off
echo ================================================
echo   Sourcing BU CRM - Starting...
echo ================================================
cd /d "%~dp0"
pip install -r requirements.txt --quiet
echo.
echo   Opening browser at: http://localhost:5000
echo ================================================
start "" "http://localhost:5000"
python app.py
pause
