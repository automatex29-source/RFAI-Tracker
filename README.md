# RFAI Tracker

RFAI Analysis — Monthly Exit Tracker for Airtel site rollouts.

## Files
- `RFAI_Analysis_Airtel_Interactive.html` — interactive dashboard (frontend)
- `converter.py` — FastAPI backend that converts Excel/CSV uploads to Parquet and serves the dashboard on http://localhost:8765
- `CONVERT_TO_PARQUET.bat` — Windows launcher: installs dependencies, starts the backend and opens the browser

## Usage (Windows)
Keep all three files in the same folder and double-click `CONVERT_TO_PARQUET.bat`.

Requirements: Python 3 with `pandas pyarrow openpyxl fastapi uvicorn python-multipart` (installed automatically by the .bat).

## Run on your PC with SQL Server (office network)

Data stays inside your own network, in your own SQL Server.

1. Install **Python** from python.org and tick **"Add python.exe to PATH"** during setup.
2. Download this repository (**Code → Download ZIP**) and unzip it, e.g. to `D:\RFAI`.
3. Double-click **`CONVERT_TO_PARQUET.bat`**. The first run installs the libraries (a few minutes).
   - If Windows asks whether to allow Python on the network, click **Allow** (private networks).
4. The browser opens **http://localhost:8765**. Drop your Excel file in. The data is saved to
   SQL Server in database **`RFAI`**, table **`dbo.rfai_data`** (each upload is logged in
   `dbo.rfai_upload_log`), and loaded from there every time the tool starts.
5. Colleagues on the same office network or VPN open **http://YOUR-PC-NAME:8765**. The black window
   shows the exact link. Keep that window open while others use the tool.

Settings are at the top of `CONVERT_TO_PARQUET.bat`:
- `SQL_SERVER`: leave empty to auto-detect, or paste the **Server name** from the SSMS login window.
- `ADMIN_PASSWORD`: the password needed to upload files. Viewing stays open on the network.
- The tool uses Windows Authentication, so no database password is needed. For a SQL login, add
  `set SQL_USER=...` and `set SQL_PASSWORD=...` instead.
- If it says *"No SQL Server ODBC driver found"*, install **ODBC Driver 18 for SQL Server** (free,
  from Microsoft) and run the .bat again.

## Large data (hundreds of GB / TB)

Raw files are **never uploaded to Render**. You convert them on your own PC, which
shrinks TBs of Excel/CSV into one small, deduplicated Parquet file (one row per SR).
That file goes to free cloud storage, and the Render app reloads it from there every time
it starts. Nothing is lost on restart or redeploy.

**Cost: $0.** The raw TBs stay on your own disk. Only the deduplicated file goes to the cloud
(about 2 MB per 1 million SRs), which fits easily in the free tiers:

| Piece | Free option | Free limit |
|---|---|---|
| Storage | Cloudflare R2 *or* Backblaze B2 | 10 GB each |
| App hosting | Render free web service | 512 MB RAM, sleeps after 15 min idle |
| Converting the raw files | Your own PC (`BULK_CONVERT.bat`) | free |

The free Render plan handles about **300,000 unique SRs**. Beyond that, switch `plan:` in
`render.yaml` to `starter`/`standard`.

```
 Your PC: TB of Excel/CSV ──bulk_convert.py──▶ rfai_dedup.parquet ──upload──▶ Cloudflare R2
                                                                                │
                                         Render app (converter.py) ◀── loads on start / Reload
```

### 1. Create the free storage (once). Pick one:
**Supabase (easiest: sign in with GitHub, no card).**
1. https://supabase.com → **Start your project** → **Continue with GitHub** → **New project**
   (any name, any database password, pick the closest region) → wait ~2 min.
2. **Storage** (left menu) → **New bucket** → name `rfai-data` (leave *Public* OFF) → Save.
3. **Project Settings → Storage → S3 Connection**: copy **Endpoint** and **Region**.
   Under **S3 Access Keys** → **New access key** → copy **Access key ID** and **Secret access key**.
4. Set `S3_ENDPOINT`=Endpoint, `S3_REGION`=Region, `S3_BUCKET`=`rfai-data`,
   `S3_ACCESS_KEY`/`S3_SECRET_KEY`=the keys. Free tier: 1 GB. Free projects pause after
   7 days with no use; click **Restore** in Supabase if that happens.

**Cloudflare R2.** Cloudflare may ask for a card to enable R2, but there is no charge
under 10 GB.
1. https://dash.cloudflare.com → **R2** → **Create bucket**, e.g. `rfai-data`.
2. **R2 → Manage API tokens → Create API token** with *Object Read & Write* on that bucket.
   Note the **Access Key ID**, **Secret Access Key** and the **S3 endpoint**
   (`https://<account-id>.r2.cloudflarestorage.com`).

**Backblaze B2.** No card needed.
1. https://www.backblaze.com/sign-up/cloud-storage → **Buckets → Create a Bucket** (Private).
2. **Application Keys → Add a New Application Key** for that bucket. Use the *keyID* as
   `S3_ACCESS_KEY`, the *applicationKey* as `S3_SECRET_KEY`, and the bucket's **Endpoint**
   with `https://` in front (e.g. `https://s3.us-west-004.backblazeb2.com`) as `S3_ENDPOINT`.

### 2. Convert + upload (on your PC, every time you have new data)
Copy `r2_settings.example.bat` to `r2_settings.bat` and fill in the R2 keys (this file is
git-ignored, so the keys stay on your PC). Put all your Excel/CSV files in one
folder (sub-folders are fine), and run:

```
BULK_CONVERT.bat  D:\RFAI_raw_files
```

- Uses all CPU cores; you can stop it and run again. Files already converted are skipped,
  so the next run only processes new or changed files.
- Duplicates across files are merged with DuckDB on disk. Each SR keeps its
  highest-priority status, and the newest file wins on ties.
- Needs free disk space of roughly 5–10% of the raw size for the `rfai_bulk_out` folder.

### 3. Point Render at the storage
Set `S3_ENDPOINT`, `S3_BUCKET`, `S3_ACCESS_KEY`, `S3_SECRET_KEY` in the Render service's
**Environment** tab.

**Updating data without the website:** upload your Excel/CSV straight into the bucket
(e.g. Supabase → Storage → rfai-data → **Upload files**), then click **↻ Reload stored data**
on the site. Files newer than the last conversion are converted automatically. Files over
the storage's own upload limit (Supabase free: 50 MB) need the website upload or `BULK_CONVERT.bat`.

**Who can see it:**
- By default **anyone with the link can view** the dashboard.
- To require a login for viewers, add a `VIEW_PASSWORD` variable (any username, that password).
- Uploading through the browser is **disabled** on Render, so visitors can't replace your
  data. Set `ADMIN_PASSWORD` to upload from the browser yourself. Uploaded data is then
  saved to the storage bucket automatically, so it survives restarts. No PC setup is needed
  for files the browser can handle (up to a few hundred MB). After each new upload, click **↻ Reload stored data**
on the start screen (or restart the service).

Pending SRs' aging is recalculated against today's date every time the data loads,
so stored data never shows stale aging.

## Deploy on Render
1. Go to https://dashboard.render.com → **New** → **Blueprint**.
2. Connect this GitHub repo (`automatex29-source/RFAI-Tracker`) and pick the branch.
3. Render reads `render.yaml` and creates the `rfai-tracker` web service. Click **Apply**.
4. Open the `https://rfai-tracker-xxxx.onrender.com` URL and upload your Excel/CSV.

Notes: the free plan sleeps after 15 min idle. The first visit after that takes about a
minute while it wakes up and reloads the data from storage. Data uploaded through the
browser is not kept after a restart. Data loaded from R2/B2 (above) is.
