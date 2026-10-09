# RFAI Tracker

RFAI Analysis — Monthly Exit Tracker for Airtel site rollouts.

## Files
- `RFAI_Analysis_Airtel_Interactive.html` — interactive dashboard (frontend)
- `converter.py` — FastAPI backend that converts Excel/CSV uploads to Parquet and serves the dashboard on http://localhost:8765
- `CONVERT_TO_PARQUET.bat` — Windows launcher: installs dependencies, starts the backend and opens the browser

## Usage (Windows)
Keep all three files in the same folder and double-click `CONVERT_TO_PARQUET.bat`.

Requirements: Python 3 with `pandas pyarrow openpyxl fastapi uvicorn python-multipart` (installed automatically by the .bat).

## Deploy on Render
1. Go to https://dashboard.render.com → **New** → **Blueprint**.
2. Connect this GitHub repo (`automatex29-source/RFAI-Tracker`) and pick the branch.
3. Render reads `render.yaml` and creates the `rfai-tracker` web service. Click **Apply**.
4. Open the `https://rfai-tracker-xxxx.onrender.com` URL and upload your Excel/CSV.

Notes: the free plan has 512 MB RAM and sleeps after 15 min idle; uploaded data
is kept in memory/local disk and is lost when the service restarts — re-upload after
a restart, or use a paid plan for large files.
