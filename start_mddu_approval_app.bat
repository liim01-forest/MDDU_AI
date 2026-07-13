@echo off
cd /d "%~dp0"

netstat -ano | findstr ":8510" | findstr "LISTENING" >nul 2>nul
if %errorlevel%==0 (
    echo App is already running. Opening browser...
    start "" "http://localhost:8510"
    pause
    exit /b
)

where python >nul 2>nul
if %errorlevel%==0 (
    start "" "http://localhost:8510"
    python -m streamlit run app.py --server.port 8510 --server.address localhost
    pause
    exit /b
)

where py >nul 2>nul
if %errorlevel%==0 (
    start "" "http://localhost:8510"
    py -m streamlit run app.py --server.port 8510 --server.address localhost
    pause
    exit /b
)

echo Python was not found.
echo Please check that Python is installed and available in PATH.
pause
