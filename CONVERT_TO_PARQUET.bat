@echo off
title RFAI Backend Server
color 0A
cd /d "%~dp0"
echo.
echo  Stopping any existing RFAI server...
for /f "tokens=5" %%a in ('netstat -aon ^| find ":8765" ^| find "LISTENING"') do (
    taskkill /f /pid %%a >nul 2>&1
)
timeout /t 1 /nobreak >nul
echo  Installing libraries (first time only)...
pip install pandas pyarrow openpyxl fastapi uvicorn "python-multipart" -q --disable-pip-version-check 2>nul
if errorlevel 1 pip install pandas pyarrow openpyxl fastapi uvicorn "python-multipart" -q
echo  Ready.
echo.
echo  Starting server at http://localhost:8765
echo  Opening browser...
start "" "http://localhost:8765"
python "%~dp0converter.py"
if %errorlevel% neq 0 (
    echo.
    echo  ERROR - see above.
    pause
)
