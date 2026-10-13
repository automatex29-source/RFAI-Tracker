@echo off
title RFAI Tracker Server
color 0A
cd /d "%~dp0"

REM ======================= SETTINGS (edit if needed) ==========================
REM Save data in SQL Server (1 = yes, 0 = no). Uses your Windows login.
set USE_SQL_SERVER=1
REM Database name (created automatically if it does not exist)
set SQL_DATABASE=RFAI
REM Server name: leave empty to auto-detect. If not found, copy the
REM "Server name" from the SSMS login window, e.g.  SNEHA-SHARMA\SQLEXPRESS
set SQL_SERVER=
REM Password needed to UPLOAD files (viewing stays open on the office network).
REM Leave empty only if you are the only person using this PC.
set ADMIN_PASSWORD=
REM ============================================================================

echo.
echo  Stopping any existing RFAI server...
for /f "tokens=5" %%a in ('netstat -aon ^| find ":8765" ^| find "LISTENING"') do (
    taskkill /f /pid %%a >nul 2>&1
)
timeout /t 1 /nobreak >nul
echo  Installing libraries (first time only)...
pip install pandas pyarrow openpyxl python-calamine lxml pyodbc fastapi uvicorn "python-multipart" -q --disable-pip-version-check 2>nul
if errorlevel 1 pip install pandas pyarrow openpyxl python-calamine lxml pyodbc fastapi uvicorn "python-multipart" -q
echo  Ready.
echo.
echo  ============================================================
echo   On this PC:          http://localhost:8765
echo   For colleagues:      http://%COMPUTERNAME%:8765
echo   (same office network / VPN; keep this window open)
echo  ============================================================
echo.
python "%~dp0converter.py"
if %errorlevel% neq 0 (
    echo.
    echo  ERROR - see above.
    pause
)
