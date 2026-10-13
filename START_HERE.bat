@echo off
title RFAI Tracker - one-time setup
color 0B
cd /d "%~dp0"
echo.
echo  ==========================================================
echo     RFAI Tracker - one-time setup
echo  ==========================================================
echo.

REM ---------- 1. Python ----------
REM PY = full path to python.exe (always used in quotes: user names can contain spaces)
set "PY="
for /f "delims=" %%i in ('python -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%i"
if not defined PY for /f "delims=" %%i in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do set "PY=%%i"
if not defined PY if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not defined PY (
    echo  [1/4] Python not found - installing Python 3.12 ^(takes 1-3 minutes^)...
    winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
)
if not defined PY (
    echo.
    echo  Could not install Python automatically.
    echo  Please install it from https://www.python.org/downloads/
    echo  ^(tick "Add python.exe to PATH"^), then double-click START_HERE.bat again.
    pause
    exit /b 1
)
echo  [1/4] Python OK: %PY%

REM ---------- 2. Libraries ----------
echo  [2/4] Installing libraries ^(first time: a few minutes^)...
"%PY%" -m pip install --upgrade pip -q --disable-pip-version-check >nul 2>&1
"%PY%" -m pip install pandas pyarrow openpyxl python-calamine lxml pyodbc fastapi uvicorn python-multipart -q --disable-pip-version-check
if errorlevel 1 (
    echo  Library install failed - check your internet connection and run START_HERE.bat again.
    pause
    exit /b 1
)
echo  [2/4] Libraries OK

REM ---------- 3. Settings ----------
if exist "%~dp0rfai_local_settings.bat" (
    echo  [3/4] Settings already saved ^(delete rfai_local_settings.bat to change them^)
) else (
    echo.
    echo  [3/4] Choose a password for UPLOADING files.
    echo        Colleagues can view without it. Avoid the characters ! %% ^& ^< ^> ^|
    set /p "ADMINPW=        Upload password: "
    call :save_settings
)

REM ---------- 4. Desktop shortcut ----------
powershell -NoProfile -Command "$s=(New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop')+'\RFAI Tracker.lnk');$s.TargetPath='%~dp0CONVERT_TO_PARQUET.bat';$s.WorkingDirectory='%~dp0';$s.IconLocation='%SystemRoot%\System32\SHELL32.dll,13';$s.Save()" >nul 2>&1
echo  [4/4] Desktop icon "RFAI Tracker" created

echo.
echo  Setup complete. Next time just double-click "RFAI Tracker" on your desktop.
echo  Starting the tool now...
timeout /t 3 /nobreak >nul
call "%~dp0CONVERT_TO_PARQUET.bat"
exit /b 0

:save_settings
> "%~dp0rfai_local_settings.bat" (
    echo @echo off
    echo REM Saved by START_HERE.bat - personal settings, not uploaded anywhere
    echo set "PY=%PY%"
    echo set "ADMIN_PASSWORD=%ADMINPW%"
)
exit /b 0
