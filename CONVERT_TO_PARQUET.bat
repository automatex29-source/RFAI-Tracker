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
REM Password of the built-in "admin" sign-in user. When set, everyone must sign in
REM (admin adds colleagues under the account button - Manage users).
REM Leave empty only if you are the only person using this PC (no sign-in then).
set ADMIN_PASSWORD=
REM ============================================================================
REM Personal settings saved by START_HERE.bat (Python path, upload password)
set "PY=python"
if exist "%~dp0rfai_local_settings.bat" call "%~dp0rfai_local_settings.bat"

echo.
echo  Stopping any existing RFAI server...
for /f "tokens=5" %%a in ('netstat -aon ^| find ":8765" ^| find "LISTENING"') do (
    taskkill /f /pid %%a >nul 2>&1
)
timeout /t 1 /nobreak >nul
"%PY%" -c "import fastapi,uvicorn,pandas,pyarrow,pyodbc,lxml,python_calamine,multipart" >nul 2>&1
if errorlevel 1 (
    echo  Installing libraries ^(first time only^)...
    "%PY%" -m pip install pandas pyarrow openpyxl python-calamine lxml pyodbc fastapi uvicorn python-multipart -q --disable-pip-version-check
)
echo  Ready.
echo.
echo  ============================================================
echo   On this PC:          http://localhost:8765
echo   For colleagues:      http://%COMPUTERNAME%:8765
echo   (same office network / VPN; keep this window open)
echo  ============================================================
echo.
"%PY%" "%~dp0converter.py"
if %errorlevel% neq 0 (
    echo.
    echo  ERROR - see above.
    pause
)
