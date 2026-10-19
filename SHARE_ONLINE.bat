@echo off
setlocal
title RFAI Tracker - share online
color 0B
cd /d "%~dp0"

REM  Runs the tracker on THIS PC (full data, full speed) and gives it a public
REM  https link through a free Cloudflare quick tunnel (no account, no card).
REM  The link works while this window is open. It changes every time you start.

set "PORT=8765"
set "PY=python"
set "ADMIN_PASSWORD="
if exist "%~dp0rfai_local_settings.bat" call "%~dp0rfai_local_settings.bat"

echo.
echo  ==========================================================
echo     RFAI Tracker - share online
echo  ==========================================================
echo.

REM ---------- 1. A sign-in password is required for a public link ----------
if not defined ADMIN_PASSWORD (
    echo  The link will be on the internet, so everyone must sign in.
    echo  Choose the password of the "admin" user ^(avoid ! %% ^& ^< ^> ^|^):
    set /p "ADMIN_PASSWORD=  Admin password: "
)
if not defined ADMIN_PASSWORD (
    echo  No password given - not sharing. Run SHARE_ONLINE.bat again.
    pause & exit /b 1
)
> "%~dp0rfai_local_settings.bat" (
    echo @echo off
    echo REM Saved by START_HERE.bat / SHARE_ONLINE.bat - personal settings, not uploaded anywhere
    echo set "PY=%PY%"
    echo set "ADMIN_PASSWORD=%ADMIN_PASSWORD%"
)

REM ---------- 2. Cloudflare tunnel program (downloaded once, ~60 MB) ----------
set "CF=%~dp0tools\cloudflared.exe"
if not exist "%CF%" (
    echo  [1/3] Downloading the free Cloudflare tunnel program ^(first time only^)...
    if not exist "%~dp0tools" mkdir "%~dp0tools"
    curl -L --fail -s -o "%CF%" https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe
    if not exist "%CF%" powershell -NoProfile -Command "Invoke-WebRequest -UseBasicParsing -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile '%CF%'"
)
if not exist "%CF%" (
    echo.
    echo  Could not download cloudflared.exe ^(the office network may block it^).
    echo  Download it yourself from
    echo    https://github.com/cloudflare/cloudflared/releases/latest
    echo  ^(file: cloudflared-windows-amd64.exe^), rename it to cloudflared.exe
    echo  and put it in:  %~dp0tools\
    pause & exit /b 1
)
echo  [1/3] Tunnel program OK

REM ---------- 3. Start the tracker (its own window) ----------
echo  [2/3] Starting the tracker...
taskkill /f /im cloudflared.exe >nul 2>&1
set "RFAI_PUBLIC=1"
start "RFAI Tracker Server" cmd /c ""%~dp0CONVERT_TO_PARQUET.bat""
set /a n=0
:wait_server
timeout /t 2 /nobreak >nul
curl -s -f -o nul "http://localhost:%PORT%/api/ping" && goto server_up
set /a n+=1
if %n% lss 90 goto wait_server
echo.
echo  The tracker did not start - look at the "RFAI Tracker Server" window for the error.
pause & exit /b 1
:server_up
echo  [2/3] Tracker running on this PC: http://localhost:%PORT%

REM ---------- 4. Open the tunnel and read its link ----------
echo  [3/3] Opening the public link...
set "LOG=%~dp0_tunnel.log"
if exist "%LOG%" del "%LOG%" >nul 2>&1
start "RFAI tunnel" /min cmd /c ""%CF%" tunnel --no-autoupdate --protocol http2 --url http://localhost:%PORT% 2> "%LOG%""
set "URL="
set /a n=0
:wait_url
timeout /t 2 /nobreak >nul
for /f "delims=" %%u in ('powershell -NoProfile -Command "$m = Select-String -Path '%LOG%' -Pattern 'https://(?!api\.)[a-z0-9]+(-[a-z0-9]+)+\.trycloudflare\.com' -ErrorAction SilentlyContinue; if ($m) { $m[0].Matches[0].Value }"') do set "URL=%%u"
if defined URL goto url_up
set /a n+=1
if %n% lss 45 goto wait_url
echo.
echo  Could not open the public link. The office network may block Cloudflare.
echo  The tracker still works on this PC and the office network:
echo    http://localhost:%PORT%      http://%COMPUTERNAME%:%PORT%
echo  Details: %LOG%
pause & exit /b 1
:url_up
> "%~dp0RFAI_link.txt" echo %URL%
echo %URL%| clip
color 0A
echo.
echo  ==========================================================
echo    PUBLIC LINK  ^(copied - paste it in WhatsApp / email^):
echo.
echo      %URL%
echo.
echo    Sign in: admin + your password ^(add colleagues under the
echo    account button - Manage users^).
echo    Upload big files on THIS PC at http://localhost:%PORT%
echo  ==========================================================
echo.
echo  Keep this window open. Close it, or press any key, to stop sharing.
start "" "%URL%"
pause >nul

REM ---------- stop ----------
taskkill /f /im cloudflared.exe >nul 2>&1
for /f "tokens=5" %%a in ('netstat -aon ^| find ":%PORT%" ^| find "LISTENING"') do taskkill /f /pid %%a >nul 2>&1
echo  Sharing stopped.
timeout /t 2 /nobreak >nul
exit /b 0
