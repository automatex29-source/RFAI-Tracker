@echo off
title RFAI Bulk Converter
cd /d "%~dp0"
if not exist "%~dp0r2_settings.bat" (
  echo  r2_settings.bat not found - copy r2_settings.example.bat to r2_settings.bat and fill in your R2 keys.
  pause & exit /b 1
)
call "%~dp0r2_settings.bat"
if "%~1"=="" (
  echo Usage: BULK_CONVERT.bat  FOLDER_WITH_EXCEL_OR_CSV_FILES
  pause & exit /b 1
)
pip install pandas pyarrow openpyxl fastapi uvicorn python-multipart duckdb boto3 -q --disable-pip-version-check
python "%~dp0bulk_convert.py" "%~1" "%~dp0rfai_bulk_out" --upload
pause
