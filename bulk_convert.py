"""
RFAI bulk converter — for very large data (hundreds of GB / TB of Excel/CSV).
=============================================================================
Run this on your own PC/server (NOT on Render). It:

  1. Converts every .csv / .xlsx / .xlsm file under INPUT_DIR to Parquet,
     using several CPU cores. Resumable: files already converted (same size +
     modified time) are skipped, so you can stop and re-run any time.
  2. Deduplicates all rows with DuckDB (works on disk, not in RAM) into ONE
     small file:  <OUT_DIR>/dedup/rfai_dedup.parquet
     (one row per SR — highest-priority status, newest file wins on ties).
  3. Optionally uploads that file to Cloudflare R2 / AWS S3, from where the
     Render app reloads it on every start.

Usage:
    python bulk_convert.py  INPUT_DIR  [OUT_DIR]  [--workers N] [--upload]

Upload settings come from environment variables (see README):
    S3_ENDPOINT, S3_BUCKET, S3_ACCESS_KEY, S3_SECRET_KEY, S3_PREFIX (optional)
"""
import os, re, sys, json, time, hashlib, argparse, io
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

import pandas as pd
import pyarrow as pa
import pyarrow.parquet as pq

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from converter import map_headers, process_row, iter_excel_rows, OUT, COLS, CHUNK  # noqa: E402

EXTS = {'.csv', '.xlsx', '.xlsm', '.xls', '.xlsb'}
DEDUP_NAME = 'rfai_dedup.parquet'


def _write(rows, path, ord_):
    arrays = {}
    for col in OUT:
        vals = [r.get(col, '') for r in rows]
        if col in ('delay', 'priDays'):
            arrays[col] = pa.array([v if isinstance(v, int) else -1 for v in vals], type=pa.int16())
        else:
            arrays[col] = pa.array([str(v) if v else '' for v in vals], type=pa.string())
    arrays['_ord'] = pa.array([ord_] * len(rows), type=pa.int64())
    tmp = str(path) + '.tmp'
    pq.write_table(pa.table(arrays), tmp, compression='zstd', use_dictionary=True)
    os.replace(tmp, path)


def _rows_csv(path):
    with open(path, 'rb') as fh:
        head = fh.read(2000)
    enc = 'utf-8'
    for e in ('utf-8', 'utf-8-sig', 'latin-1'):
        try: head.decode(e); enc = e; break
        except Exception: pass
    first = head.decode(enc, errors='replace')
    delim = next((d for d in (',', '\t', '|', ';') if d in first), ',')
    cm = None
    for chunk in pd.read_csv(path, sep=delim, dtype=str, keep_default_na=False,
                             chunksize=CHUNK, encoding=enc, on_bad_lines='skip',
                             encoding_errors='replace'):
        if cm is None: cm = map_headers(list(chunk.columns))
        for row in chunk.itertuples(index=False, name=None):
            yield row, cm


def _rows_xlsx(path):
    it = iter_excel_rows(path)
    hdr = next(it, None)
    if not hdr: return
    cm = map_headers([str(c).strip() if c is not None else '' for c in hdr])
    for row in it:
        yield row, cm


def convert_one(src, raw_dir, ord_):
    """Worker: one source file -> one or more raw parquet parts."""
    src = Path(src)
    tag = hashlib.md5(str(src).encode()).hexdigest()[:10]
    for old in Path(raw_dir).glob(f'{tag}_*.parquet'): old.unlink()
    rows_iter = _rows_csv(src) if src.suffix.lower() == '.csv' else _rows_xlsx(src)
    scanned = kept = part = 0; batch = []
    for row, cm in rows_iter:
        scanned += 1
        r = process_row(row, cm)
        if r: batch.append(r); kept += 1
        if len(batch) >= CHUNK:
            _write(batch, Path(raw_dir) / f'{tag}_{part:05d}.parquet', ord_); batch = []; part += 1
    if batch:
        _write(batch, Path(raw_dir) / f'{tag}_{part:05d}.parquet', ord_); part += 1
    return str(src), scanned, kept, part


def dedup(raw_dir, out_file, mem_limit):
    import duckdb
    tmp_dir = Path(out_file).parent / '_duckdb_tmp'; tmp_dir.mkdir(parents=True, exist_ok=True)
    con = duckdb.connect()
    con.execute(f"SET temp_directory='{tmp_dir.as_posix()}'")
    con.execute(f"SET memory_limit='{mem_limit}'")
    con.execute("SET preserve_insertion_order=false")
    src = f"read_parquet('{Path(raw_dir).as_posix()}/*.parquet')"
    cols = ', '.join(f'"{c}"' for c in OUT)
    pri = ("CASE srStatus WHEN 'OnAir' THEN 5 WHEN 'OnAir Pending' THEN 4 WHEN 'Open PRI' THEN 3 "
           "WHEN 'RFAI Pending' THEN 2 WHEN 'Rejected RFAI' THEN 1 ELSE 0 END")
    tmp_out = str(out_file) + '.tmp'
    con.execute(f"""
      COPY (
        SELECT {cols} FROM (
          SELECT *, row_number() OVER (PARTITION BY srNo ORDER BY {pri} DESC, _ord DESC) AS _rn
          FROM {src} WHERE srNo <> ''
        ) WHERE _rn = 1
        UNION ALL
        SELECT DISTINCT {cols} FROM {src} WHERE srNo = ''
      ) TO '{Path(tmp_out).as_posix()}' (FORMAT parquet, COMPRESSION zstd, ROW_GROUP_SIZE 100000)
    """)
    os.replace(tmp_out, out_file)
    n = con.execute(f"SELECT count(*) FROM read_parquet('{Path(out_file).as_posix()}')").fetchone()[0]
    con.close()
    return n


def _s3_region():
    """R2 uses 'auto'; Backblaze B2 needs the region from its endpoint."""
    if os.environ.get('S3_REGION'): return os.environ['S3_REGION']
    m = re.search(r's3\.([a-z0-9-]+)\.backblazeb2\.com', os.environ.get('S3_ENDPOINT', ''))
    return m.group(1) if m else 'auto'

def upload(path):
    import boto3
    bucket = os.environ.get('S3_BUCKET')
    if not bucket:
        print('  ! S3_BUCKET not set — skipping upload'); return
    prefix = os.environ.get('S3_PREFIX', 'rfai/').rstrip('/') + '/'
    s3 = boto3.client('s3', endpoint_url=os.environ.get('S3_ENDPOINT') or None,
                      aws_access_key_id=os.environ.get('S3_ACCESS_KEY'),
                      aws_secret_access_key=os.environ.get('S3_SECRET_KEY'),
                      region_name=_s3_region(),
                        config=__import__('botocore.config',fromlist=['Config']).Config(
                            request_checksum_calculation='when_required',
                            s3={'addressing_style': 'path'},
                            response_checksum_validation='when_required'))
    key = prefix + DEDUP_NAME
    print(f'  Uploading {Path(path).stat().st_size/1e6:,.1f} MB -> s3://{bucket}/{key}')
    s3.upload_file(str(path), bucket, key)
    # remove older files (e.g. from a browser upload) so only this data is loaded
    for page in s3.get_paginator('list_objects_v2').paginate(Bucket=bucket, Prefix=prefix):
        for o in page.get('Contents', []):
            if o['Key'] != key and o['Key'].endswith('.parquet'):
                s3.delete_object(Bucket=bucket, Key=o['Key'])
    print('  ✓ Uploaded. Open the Render app and click Reload (or restart it).')


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('input_dir'); ap.add_argument('out_dir', nargs='?', default='rfai_bulk_out')
    ap.add_argument('--workers', type=int, default=max(1, (os.cpu_count() or 2) - 1))
    ap.add_argument('--memory', default='4GB', help='DuckDB memory limit for dedup (default 4GB)')
    ap.add_argument('--upload', action='store_true', help='upload result to S3/R2 after dedup')
    a = ap.parse_args()

    inp = Path(a.input_dir); out = Path(a.out_dir)
    raw = out / 'raw'; ded = out / 'dedup'
    raw.mkdir(parents=True, exist_ok=True); ded.mkdir(parents=True, exist_ok=True)
    man_p = out / 'manifest.json'
    manifest = json.loads(man_p.read_text()) if man_p.exists() else {}

    files = sorted(p for p in inp.rglob('*') if p.is_file() and p.suffix.lower() in EXTS
                   and not p.name.startswith('~$'))
    todo = []
    for p in files:
        st = p.stat(); sig = [st.st_size, int(st.st_mtime)]
        if manifest.get(str(p), {}).get('sig') != sig: todo.append((p, sig))
    total_gb = sum(p.stat().st_size for p in files) / 1e9
    print(f'\n  {len(files):,} files ({total_gb:,.1f} GB) found · {len(todo):,} to convert · '
          f'{a.workers} workers\n')

    t0 = time.time(); done = 0
    with ProcessPoolExecutor(max_workers=a.workers) as ex:
        futs = {ex.submit(convert_one, str(p), str(raw), sig[1]): (p, sig) for p, sig in todo}
        for f in as_completed(futs):
            p, sig = futs[f]; done += 1
            try:
                _, scanned, kept, parts = f.result()
                manifest[str(p)] = {'sig': sig, 'scanned': scanned, 'kept': kept, 'parts': parts}
                man_p.write_text(json.dumps(manifest))
                print(f'  [{done}/{len(todo)}] {p.name}: {scanned:,} rows, {kept:,} with RFAI '
                      f'({time.time()-t0:,.0f}s)')
            except Exception as e:
                print(f'  [{done}/{len(todo)}] ERROR {p.name}: {e}')

    # drop parts of source files that were deleted since the last run
    live = {hashlib.md5(k.encode()).hexdigest()[:10] for k in manifest if Path(k).exists()}
    for k in [k for k in manifest if not Path(k).exists()]: manifest.pop(k)
    man_p.write_text(json.dumps(manifest))
    for f in raw.glob('*.parquet'):
        if f.name.split('_')[0] not in live: f.unlink()

    if not any(raw.glob('*.parquet')):
        print('  No data converted.'); return
    print('\n  Deduplicating (DuckDB, on disk)…')
    n = dedup(raw, ded / DEDUP_NAME, a.memory)
    sz = (ded / DEDUP_NAME).stat().st_size / 1e6
    print(f'  ✓ {n:,} unique SR rows -> {ded / DEDUP_NAME} ({sz:,.1f} MB)')
    if a.upload: upload(ded / DEDUP_NAME)


if __name__ == '__main__':
    main()
