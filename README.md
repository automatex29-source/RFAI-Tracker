# RFAI Tracker

RFAI Analysis — Monthly Exit Tracker for Airtel site rollouts.

## Files
- `RFAI_Analysis_Airtel_Interactive.html` — interactive dashboard (frontend)
- `converter.py` — FastAPI backend that converts Excel/CSV uploads to Parquet and serves the dashboard on http://localhost:8765
- `CONVERT_TO_PARQUET.bat` — Windows launcher: installs dependencies, starts the backend and opens the browser

## Usage (Windows)
Keep all three files in the same folder and double-click `CONVERT_TO_PARQUET.bat`.

Requirements: Python 3 with `pandas pyarrow openpyxl fastapi uvicorn python-multipart` (installed automatically by the .bat).
