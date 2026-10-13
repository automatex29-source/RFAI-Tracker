"""
RFAI Backend — v4 (Python 3.13 / PyArrow-safe)
================================================
Root cause of previous crash: pandas 3.x on Python 3.13 uses PyArrow-backed
string arrays by default. PyArrow's .take() fails with malloc errors when the
filtered result is large because it tries to allocate contiguous memory.

Fixes in v4:
  • Force numpy-backed dtypes on every parquet read (dtype_backend='numpy_nullable')
  • Dedup with groupby+idxmax — no boolean filter on large frame, no .take() call
  • CSV reads via pandas.read_csv (no csv.Sniffer / import collision)
  • int16 for delay, int8 for _sp, Categorical for string cols
  • gc.collect() after every heavy step
  • Full traceback logged on errors
"""
import os, sys, re, json, time, logging, threading, io, gc
# system allocator hands freed memory back to the OS (matters on 512 MB hosts)
os.environ.setdefault('ARROW_DEFAULT_MEMORY_POOL', 'system')
from pathlib import Path
from datetime import datetime, date

try:
    import pandas as pd
    import numpy as np
    import pyarrow as pa
    import pyarrow.parquet as pq
    import openpyxl
except ImportError:
    print("\n  Run: pip install pandas pyarrow openpyxl fastapi uvicorn python-multipart\n")
    sys.exit(1)

try:
    import uvicorn
    from fastapi import FastAPI, UploadFile, File, Query, BackgroundTasks
    from fastapi.responses import JSONResponse, FileResponse
    from fastapi.middleware.cors import CORSMiddleware
    from typing import List
except ImportError:
    print("\n  Run: pip install fastapi uvicorn python-multipart\n")
    sys.exit(1)

logging.basicConfig(level=logging.INFO,
                    format="%(asctime)s  %(message)s", datefmt="%H:%M:%S")
log = logging.getLogger("rfai")

SCRIPT_DIR = Path(os.path.dirname(os.path.abspath(__file__)))
HTML_FILE  = next(SCRIPT_DIR.glob("RFAI_Analysis_Airtel*.html"),
                  SCRIPT_DIR / "RFAI_Analysis_Airtel_v10.html")
PORT  = int(os.environ.get("PORT", 8765))
CHUNK = 50_000
TODAY = date.today()

# ── Column aliases ─────────────────────────────────────────────────────────
COLS = {
    'CIRCLE':        ['CIRCLE','Circle','circle','STATE'],
    'SR_NO':         ['SR_NO','SR No','SR No.','SRNO','SR_NUMBER'],
    'SITE_ID':       ['2G_SITE_ID','SITE_ID','Site ID','SiteId'],
    'TOCO':          ['TOCO_NAME','TOCO','Toco'],
    'PARTNER':       ['Integration Partner','INTEGRATION_PARTNER','Infra Partner'],
    'BROAD_PROJECT': ['BROAD_PROJECT','Broad Project','Project','PROJECT'],
    'DOMAIN':        ['Domain','DOMAIN','domain','Sub Project','SUB_PROJECT'],
    'OEM':           ['OEM_NAME','OEM','OEM Name'],
    'TSP':           ['TSP_PROJECT','TSP','TSP Project'],
    'RFAI_NOTICE':   ['RFAI_NOTICE_FORM_DATE','RFAI Notice','RFAI Date','rfai_date'],
    'RFAI_REJ':      ['RFAI_LAST_REJECTION_DATE','RFAI Rejection','Rejection Date'],
    'RFAI_REGEN':    ['RFAI_REGENERATION_DATE','RFAI Regen'],
    'RFAI_OFFERED':  ['RFAI_LAST_OFFERED_DATE','RFAI Offered'],
    'FINAL_RFAI':    ['Final RFAI','FINAL_RFAI','Final RFAI Date'],
    'OA_DATE':       ['OA_(COMMERCIAL_TRAFFIC_PUT_ON_AIR)_(MS1)_DATE','OA_DATE',
                      'OnAir Date','On Air Date','Go Live Date'],
    'PRI_HIST':      ['PriHistory','PRI_HISTORY','PRI History','PRI'],
    'MO_DATE':       ['MO_PUNCHING_DATE','MO Date'],
    'MD_DATE':       ['MATERIAL_DISPATCH_(MD)_DATE','MD Date'],
    'MOS_DATE':      ['MATERIAL_DELIVERY_(MOS)_DATE','MOS Date'],
    'INST_START':    ['INSTALLATION_START_DATE','Install Start'],
    'INST_END':      ['INSTALLATION_END_DATE','Install End'],
    'PMI_OPEN':      ['POST_MEDIA_ISSUE_OPEN_DATE','PMI Open'],
    'PMI_CLOSE':     ['POST_MEDIA_ISSUE_CLOSE_DATE','PMI Close'],
}

MON     = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
MON_IDX = {m.lower(): i for i, m in enumerate(MON)}
_dc: dict = {}

# ── Date parser ────────────────────────────────────────────────────────────
def pd_(v):
    if v is None: return None
    if isinstance(v, date): return v
    if isinstance(v, datetime): return v.date()
    s = str(v).strip()
    if not s or s in ('-','—','nan','None','NaT','0',''): return None
    cached = _dc.get(s)
    if cached is not None: return cached if cached != '' else None
    result = None
    try:
        m = re.match(
            r'(\d{1,2})[\-\/](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[\-\/](\d{2,4})$',
            s, re.IGNORECASE)
        if m:
            day = int(m.group(1)); mo = MON_IDX.get(m.group(2).lower())
            yr  = int(m.group(3))
            if yr < 100: yr += 2000
            if mo is not None and 1 <= day <= 31:
                result = date(yr, mo+1, day)
        elif len(s) >= 10 and s[4] == '-':
            result = date(int(s[:4]), int(s[5:7]), int(s[8:10]))
        elif re.search(r'[/\-]', s):
            sep = '/' if '/' in s else '-'
            p = s.split(sep)
            if len(p) == 3:
                a, b, y = int(p[0]), int(p[1]), int(p[2])
                if y < 100: y += 2000
                if y > 1900:
                    if 1 <= b <= 12 and 1 <= a <= 31: result = date(y, b, a)
                elif 1 <= a <= 12 and 1 <= b <= 31:   result = date(y, a, b)
        elif s.replace('.','').lstrip('-').isdigit():
            n = float(s)
            if 1000 < n < 70000:
                from datetime import timedelta
                result = date(1899, 12, 30) + timedelta(days=int(n))
    except Exception: pass
    if result and not (2000 <= result.year <= 2035): result = None
    _dc[s] = result if result else ''
    return result

def fd(d):  return f"{d.day:02d}-{MON[d.month-1]}-{d.year}" if d else ''
def mk(d):
    if not d: return None
    m, y = d.month-1, d.year
    if d.day > 25:
        m += 1
        if m > 11: m = 0; y += 1
    return f"{MON[m]}'{str(y)[2:]}"
def jmk(d): return f"{MON[d.month-1]}'{str(d.year)[2:]}" if d else None

# ── Category ───────────────────────────────────────────────────────────────
CAT = {
    'SWAP':'SWAP','COLD SWAP':'SWAP',
    '5G':'5G Upgrade','5G UPGRADE':'5G Upgrade','5G_UPGRADE':'5G Upgrade',
    '5G SEC ADDITION':'5G Upgrade','5G_SEC_ADDITION':'5G Upgrade',
    'FDD_UPGRADE':'4G Upgrade','FDD UPGRADE':'4G Upgrade',
    'TDD_UPGRADE':'4G Upgrade','TDD UPGRADE':'4G Upgrade',
    '4G UPGRADE':'4G Upgrade','4G':'4G Upgrade',
    '2G_UPGRADE':'4G Upgrade','NEW_SECTOR_ADDITION':'4G Upgrade','TWIN_BEAM':'4G Upgrade',
    'TPIP':'Fibre','OLT/FTTH':'Fibre','FIBER':'Fibre','FIBRE':'Fibre',
    'OLT':'Fibre','FTTH':'Fibre',
    'NT':'NT','NEW_TOWER':'NT','NEW TOWER':'NT','NEW SITE':'NT',
    'ULS':'ULS','NEW_TOWER_ULS':'ULS',
    'MW':'MW','MICROWAVE':'MW','E-BAND':'MW',
    'RELOCATION':'Relocation','RELOCENT':'Relocation',
}
def get_cat(proj):
    if not proj: return None
    u = str(proj).strip().upper()
    v = CAT.get(u) or CAT.get(u.replace('_',' '))
    if v: return v
    if u.startswith('5G'):                               return '5G Upgrade'
    if u.startswith('4G') or 'FDD' in u or 'TDD' in u: return '4G Upgrade'
    if 'FIBER' in u or 'FIBRE' in u:                    return 'Fibre'
    if 'RELOC' in u:                                     return 'Relocation'
    if 'NEW TOWER' in u or 'NEW_TOWER' in u:            return 'NT'
    if u.startswith('MW') or 'MICROWAVE' in u:          return 'MW'
    if 'SWAP' in u:                                      return 'SWAP'
    return None

# ── PRI parser ─────────────────────────────────────────────────────────────
def parse_pri(v):
    if not v or str(v).strip() in ('','nan','None','-','—'):
        return None, None, 0, False
    s = str(v).strip(); ranges = []
    for part in re.split(r'[|;\n]', s):
        t = part.strip()
        if not t: continue
        m = re.search(r'[-–—]>', t)
        if m: a, b = pd_(t[:m.start()].strip()), pd_(t[m.end():].strip())
        else: a, b = pd_(t), None
        if a: ranges.append((a, b))
    if not ranges: return None, None, 0, False
    ranges.sort(key=lambda x: x[0])
    first = ranges[0][0]; is_open = any(b is None for _, b in ranges)
    tot = 0; last = None
    for a, b in ranges:
        if b:
            d = (b-a).days
            if d > 0: tot += d
            if not last or b > last: last = b
    return first, last, tot, is_open

# ── Header resolver ────────────────────────────────────────────────────────
def _norm(s): return re.sub(r'[\s_\-\(\)/\.]+','',str(s)).lower()
def map_headers(headers):
    hmap = {_norm(h): i for i, h in enumerate(headers)}
    result = {}
    for col, aliases in COLS.items():
        idx = None
        for a in aliases:
            if a in headers: idx = headers.index(a); break
        if idx is None:
            for a in aliases:
                n = _norm(a)
                if n in hmap: idx = hmap[n]; break
        if idx is None:
            for a in aliases:
                an = _norm(a)
                for hn, hi in hmap.items():
                    if len(an) > 3 and (hn == an or hn in an or an in hn):
                        idx = hi; break
                if idx is not None: break
        result[col] = idx
    return result

def _g(row, cm, col):
    idx = cm.get(col)
    if idx is None or idx >= len(row): return None
    v = row[idx]
    return None if v in (None,'','nan','None','NaT') else str(v).strip() or None

# ── Row processor ──────────────────────────────────────────────────────────
def process_row(row, cm):
    rfai = pd_(_g(row, cm, 'RFAI_NOTICE'))
    if not rfai: return None
    if not (2020 <= rfai.year <= 2030): return None
    rej = pd_(_g(row,cm,'RFAI_REJ')); regen = pd_(_g(row,cm,'RFAI_REGEN'))
    final_c = pd_(_g(row,cm,'FINAL_RFAI')); oa_raw = pd_(_g(row,cm,'OA_DATE'))
    mo_d=pd_(_g(row,cm,'MO_DATE')); md_d=pd_(_g(row,cm,'MD_DATE'))
    mos_d=pd_(_g(row,cm,'MOS_DATE')); is_d=pd_(_g(row,cm,'INST_START'))
    ie_d=pd_(_g(row,cm,'INST_END')); po_d=pd_(_g(row,cm,'PMI_OPEN'))
    pc_d=pd_(_g(row,cm,'PMI_CLOSE'))
    pf, pl, pd_days, p_open = parse_pri(_g(row,cm,'PRI_HIST'))
    oa    = oa_raw if oa_raw and 2000<=oa_raw.year<=2035 and oa_raw>=rfai else None
    final = final_c or (rej and (regen or oa)) or rfai
    reoffer = regen if rej else None
    bd = None; bl = ''
    for d_,l in [(mo_d,'MO'),(md_d,'MD'),(mos_d,'MOS'),(is_d,'Inst Start'),
                 (ie_d,'Inst End'),(po_d,'PMI Open'),(pc_d,'PMI Close'),(rfai,'RFAI')]:
        if d_ and (bd is None or d_ < bd): bd = d_; bl = l
    if not bd: bd = rfai; bl = 'RFAI'
    if final and (not bd or final > bd): bd = final; bl = 'Final RFAI'
    if oa:              st = 'OnAir'
    elif p_open:        st = 'Open PRI'
    elif rej and not regen: st = 'Rejected RFAI'
    elif final or pf:   st = 'OnAir Pending'
    else:               st = 'RFAI Pending'
    anchor = final or rfai or bd; delay = None
    if anchor and oa:
        d_ = (oa-anchor).days
        if pd_days > 0: d_ = max(0, d_-pd_days)
        delay = max(0, d_)
    elif anchor and st in ('OnAir Pending','Open PRI','RFAI Pending'):
        d_ = (TODAY-anchor).days
        if pd_days > 0: d_ = max(0, d_-pd_days)
        delay = max(0, d_)
    grp = ('N/A' if delay is None else
           '0-30' if delay<=30 else '30-60' if delay<=60 else
           '60-90' if delay<=90 else '>90')
    exit_mk = (jmk(oa) if st=='OnAir' and oa
               else mk(rfai if st!='Rejected RFAI' else (final or rfai)))
    bp = _g(row,cm,'BROAD_PROJECT') or ''; dm = _g(row,cm,'DOMAIN') or ''
    return {
        'srNo':_g(row,cm,'SR_NO') or '','siteId':_g(row,cm,'SITE_ID') or '',
        'circle':_g(row,cm,'CIRCLE') or '','toco':_g(row,cm,'TOCO') or '',
        'partner':_g(row,cm,'PARTNER') or '','bPrj':bp,'domain':dm,
        'oem':_g(row,cm,'OEM') or '','tsp':_g(row,cm,'TSP') or '',
        'category':get_cat(bp) or get_cat(dm) or '',
        'srStatus':st,'delay':delay if delay is not None else -1,'group':grp,
        'exitMK':exit_mk or '','journeyMK':jmk(rfai) or '',
        'rfaiNotice':fd(rfai),'rfaiRej':fd(rej),'rfaiRegen':fd(regen),
        'rfaiReoffer':fd(reoffer),'finalRfai':fd(final),
        'priFirst':fd(pf),'priLast':fd(pl),'priDays':pd_days,
        'onAir':fd(oa),'baseDate':fd(bd),'baseLabel':bl,
    }

OUT = ['srNo','siteId','circle','toco','partner','bPrj','domain','oem','tsp','category',
       'srStatus','delay','group','exitMK','journeyMK',
       'rfaiNotice','rfaiRej','rfaiRegen','rfaiReoffer','finalRfai',
       'priFirst','priLast','priDays','onAir','baseDate','baseLabel']

_SHEET_KEYS = ('RFAI','CIRCLE','SR_NO','SITE','BROAD','OA')

def _norm_cell(v):
    # calamine returns every number as float: 123456.0 -> 123456 (as openpyxl did)
    if isinstance(v, float) and v.is_integer(): return int(v)
    return None if v == '' else v

def iter_excel_rows(path):
    """Yield rows (tuples) of the best-matching sheet, header row first.
    Uses calamine (Rust, ~10x faster than openpyxl, also reads .xls/.xlsb/.ods);
    falls back to openpyxl if calamine is unavailable."""
    try:
        from python_calamine import CalamineWorkbook
    except ImportError:
        CalamineWorkbook = None
    if CalamineWorkbook is not None:
        wb = CalamineWorkbook.from_path(str(path))
        best, best_sc = 0, -1
        for i, sn in enumerate(wb.sheet_names):
            hdr = next(iter(wb.get_sheet_by_index(i).iter_rows()), None) if len(wb.sheet_names) > 1 else []
            sc = sum(1 for c in (hdr or []) if c and any(k in str(c).upper() for k in _SHEET_KEYS))
            if sc > best_sc: best, best_sc = i, sc
        for row in wb.get_sheet_by_index(best).iter_rows():
            yield tuple(_norm_cell(v) for v in row)
        return
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    try:
        best, best_sc = wb.sheetnames[0], -1
        for sn in wb.sheetnames:
            r1 = next(wb[sn].iter_rows(max_row=1, values_only=True), None) or ()
            sc = sum(1 for c in r1 if c and any(k in str(c).upper() for k in _SHEET_KEYS))
            if sc > best_sc: best, best_sc = sn, sc
        yield from wb[best].iter_rows(values_only=True)
    finally:
        wb.close()

def to_pq(rows, path):
    if not rows: return
    arrays = {}
    for col in OUT:
        vals = [r.get(col,'') for r in rows]
        if col in ('delay','priDays'):
            arrays[col] = pa.array([v if isinstance(v,int) else -1 for v in vals], type=pa.int16())
        else:
            arrays[col] = pa.array([str(v) if v else '' for v in vals], type=pa.string())
    pq.write_table(pa.table(arrays), str(path), compression='snappy', use_dictionary=True)

# ── Parquet reader — ALWAYS numpy-backed, never PyArrow-backed ─────────────
def _read_pq(path, columns=None):
    """Read parquet and FORCE numpy dtypes — avoids PyArrow-backed string
    arrays that cause malloc failures on pandas 3.x / Python 3.13."""
    tbl = pq.read_table(str(path), columns=columns)
    # Convert to pandas with explicit numpy dtype backend
    df = tbl.to_pandas(types_mapper=lambda t: (
        pd.StringDtype()  # numpy object-backed strings
        if pa.types.is_string(t) or pa.types.is_large_string(t) else None
    ))
    # Belt-and-braces: force any remaining ArrowDtype → numpy object
    for col in df.columns:
        if hasattr(df[col].dtype, 'pyarrow_dtype'):
            df[col] = df[col].astype(object).fillna('')
        # Ensure string columns are plain Python str (object dtype)
        if df[col].dtype.name in ('string','large_string') or str(df[col].dtype).startswith('string'):
            df[col] = df[col].astype(str).replace('<NA>','').replace('nan','')
    return df

# ── Job state ──────────────────────────────────────────────────────────────
class Job:
    def __init__(self): self.reset()
    def reset(self):
        self.running=False; self.done=False; self.error=''
        self.progress=0; self.message=''; self.current_file=''
        self.rows_processed=0; self.rows_kept=0; self.total_files=0
        self.cache_dir=None; self.file_stats=[]; self.rows_no_sr=0

_job       = Job()
_cache_dir = Path(os.environ.get('DATA_DIR') or (SCRIPT_DIR / '_rfai_cache'))
_dash_cache    = None
_df_cache      = None
_df_full_cache = None

_df_cache_cols = ['circle','category','bPrj','toco','domain','oem',
                  'exitMK','journeyMK','srStatus','delay','srNo','finalRfai','priDays']
_CAT_COLS      = ['circle','category','bPrj','toco','domain','oem',
                  'srStatus','group','exitMK','journeyMK']

def _refresh_pending_delay(df):
    """Stored data can be days/months old: recompute delay of still-pending
    SRs as of TODAY (same formula as process_row: TODAY - Final RFAI - PRI days)."""
    if 'finalRfai' not in df.columns or 'srStatus' not in df.columns: return df
    pend = df['srStatus'].astype(str).isin(['OnAir Pending','Open PRI','RFAI Pending']).to_numpy()
    if not pend.any(): return df
    anchor = pd.to_datetime(df['finalRfai'].astype(str), format='%d-%b-%Y', errors='coerce')
    days = (pd.Timestamp(date.today()) - anchor).dt.days
    pri = (pd.to_numeric(df['priDays'], errors='coerce').fillna(0).clip(lower=0)
           if 'priDays' in df.columns else 0)
    new = (days - pri).clip(lower=0)
    upd = pend & new.notna().to_numpy()
    delay = pd.to_numeric(df['delay'], errors='coerce').fillna(-1).to_numpy().astype('int32')
    delay[upd] = new.to_numpy()[upd].astype('int32')
    df['delay'] = np.minimum(delay, 32767).astype('int16')
    if 'group' in df.columns:
        d = df['delay'].to_numpy()
        g = np.where(d < 0, 'N/A', np.where(d <= 30, '0-30', np.where(d <= 60, '30-60',
                     np.where(d <= 90, '60-90', '>90'))))
        df['group'] = g
    return df

def _optimise_df(df):
    """Categorical strings + int16 numerics."""
    df = _refresh_pending_delay(df)
    for col in _CAT_COLS:
        if col in df.columns:
            try: df[col] = df[col].astype('category')
            except Exception: pass
    for col in ('delay','priDays'):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(-1).astype('int16')
    return df

# ── Memory-safe dedup — groupby+idxmax, NO boolean filter, NO sort_values ──
_SP = {'OnAir':5,'OnAir Pending':4,'Open PRI':3,'RFAI Pending':2,'Rejected RFAI':1}

def _dedup_df(df):
    """Keep highest-priority status per srNo.
    Uses integer index selection only — no .take(), no boolean filter copy."""
    if df.empty: return df
    # Work on plain numpy object arrays to avoid PyArrow path entirely
    sr_arr  = df['srNo'].to_numpy(dtype=object, na_value='')
    st_arr  = df['srStatus'].to_numpy(dtype=object, na_value='')

    # Vectorised priority lookup
    sp_arr  = np.array([_SP.get(str(s), 0) for s in st_arr], dtype=np.int8)

    named_mask = sr_arr != ''
    named_idx  = np.where(named_mask)[0]
    blank_idx  = np.where(~named_mask)[0]

    if len(named_idx):
        # groupby via dictionary — pure Python, zero PyArrow allocation
        best: dict = {}          # srNo -> (priority, positional_index)
        for pos in named_idx:
            sr  = sr_arr[pos]
            pri = int(sp_arr[pos])
            if sr not in best or pri > best[sr][0]:
                best[sr] = (pri, int(pos))
        keep_named = np.array([v[1] for v in best.values()], dtype=np.intp)
        keep_named.sort()
    else:
        keep_named = np.array([], dtype=np.intp)

    keep_all = np.concatenate([keep_named, blank_idx]) if len(blank_idx) else keep_named
    return df.iloc[keep_all].reset_index(drop=True)

# ── CSV reader — pandas.read_csv, no Sniffer ──────────────────────────────
def _process_csv(path, fname: str, grand: int, kept: int, chunk_idx: int):
    with open(path, 'rb') as fh: head = fh.read(65536)
    enc = 'cp1252'
    for e in ('utf-8-sig', 'utf-8'):
        try: head.decode(e); enc = e; break
        except UnicodeDecodeError as ex:
            if ex.start > len(head) - 4: enc = e; break   # cut mid-character at buffer end
    first_line = head.decode(enc, errors='replace')[:2000]
    delim = ','
    for d in (',','\t','|',';'):
        if d in first_line: delim = d; break
    try:
        reader = pd.read_csv(
            path, sep=delim, dtype=str, keep_default_na=False, encoding=enc,
            encoding_errors='replace', chunksize=CHUNK, low_memory=False, on_bad_lines='skip')
    except Exception as e:
        log.error(f"  CSV open failed {fname}: {e}")
        return grand, kept, chunk_idx

    headers = None; cm = None; batch = []
    for chunk_df in reader:
        if headers is None:
            headers = list(chunk_df.columns)
            cm = map_headers(headers)
            log.info(f"  Matched {sum(1 for v in cm.values() if v is not None)}/{len(COLS)} cols")
        for row in chunk_df.itertuples(index=False, name=None):
            grand += 1
            r = process_row(row, cm)
            if r: batch.append(r); kept += 1
        _job.rows_processed = grand; _job.rows_kept = kept
        _job.message = f'{fname} — {grand:,} rows read'
        if len(batch) >= CHUNK:
            to_pq(batch, _stage_dir / f"{Path(fname).stem}_c{chunk_idx:04d}.parquet")
            batch = []; chunk_idx += 1
            _job.rows_processed = grand; _job.rows_kept = kept
    if batch:
        to_pq(batch, _stage_dir / f"{Path(fname).stem}_c{chunk_idx:04d}.parquet")
        chunk_idx += 1
    return grand, kept, chunk_idx

# ── Background file processor ──────────────────────────────────────────────
_stage_dir = _cache_dir / '_incoming'   # new upload is built here, swapped in only on success

def process_files_bg(file_infos):
    global _job, _cache_dir, _df_cache, _df_full_cache, _dash_cache
    _job.reset(); _job.running=True
    _stage_dir.mkdir(parents=True, exist_ok=True)
    for f in _stage_dir.glob("*.parquet"): f.unlink()
    _job.cache_dir=_cache_dir; _job.total_files=len(file_infos)
    grand=0; kept=0; errors=[]

    for fi,(fname,content) in enumerate(file_infos):
        _job.current_file=fname
        _job.progress=int(fi/len(file_infos)*90)
        _job.message=f'[{fi+1}/{len(file_infos)}] {fname}'
        log.info(f"  {fname}")
        ext=Path(fname).suffix.lower(); chunk_idx=0
        try:
            if ext=='.csv':
                grand, kept, chunk_idx = _process_csv(content, fname, grand, kept, chunk_idx)
            else:
                it=iter_excel_rows(content)
                hdr=next(it,None) or ()
                headers=[str(c_).strip() if c_ is not None else '' for c_ in hdr]
                cm=map_headers(headers)
                log.info(f"  Matched {sum(1 for v in cm.values() if v is not None)}/{len(COLS)} cols")
                batch=[]
                for row in it:
                    grand+=1; r=process_row(row,cm)
                    if r: batch.append(r); kept+=1
                    if grand % 5000 == 0:
                        _job.rows_processed=grand; _job.rows_kept=kept
                        _job.message=f'[{fi+1}/{len(file_infos)}] {fname} — {grand:,} rows read'
                    if len(batch)>=CHUNK:
                        to_pq(batch,_stage_dir/f"{Path(fname).stem}_c{chunk_idx:04d}.parquet")
                        batch=[]; chunk_idx+=1
                if batch:
                    to_pq(batch,_stage_dir/f"{Path(fname).stem}_c{chunk_idx:04d}.parquet")
        except Exception as e:
            log.error(f"  ERROR {fname}: {e}")
            import traceback; traceback.print_exc()
            errors.append(f'{fname}: {e}')
        finally:
            try: Path(content).unlink()
            except Exception: pass

        ps=sum(s['scanned'] for s in _job.file_stats)
        pk=sum(s['kept']    for s in _job.file_stats)
        _job.file_stats.append({'name':fname,'scanned':grand-ps,'kept':kept-pk})
        _job.rows_processed=grand; _job.rows_kept=kept

    _job.rows_no_sr=0
    for pf in sorted(_stage_dir.glob("*.parquet")):
        try:
            t=pq.read_table(str(pf),columns=['srNo'])
            _job.rows_no_sr+=int(t.column('srNo').to_pylist().count(''))
        except Exception: pass

    log.info(f"  Total scanned: {grand:,} | With RFAI date: {kept:,} | "
             f"Dropped (no RFAI): {grand-kept:,} | Blank SR#: {_job.rows_no_sr:,}")
    if kept == 0:
        _job.error = ('No rows with an RFAI date were found. ' + ' | '.join(errors)
                      if errors else 'No rows with an RFAI date were found in the file(s).')
        for f in _stage_dir.glob("*.parquet"): f.unlink()
        _job.progress=100; _job.running=False; return   # existing data left untouched
    _job.progress=92; _job.message='Preparing dashboard…'
    # success: replace the old data with the new upload
    _df_cache=None; _df_full_cache=None; _dash_cache=None; gc.collect()
    for f in _cache_dir.glob("*.parquet"): f.unlink()
    for f in _stage_dir.glob("*.parquet"): os.replace(f, _cache_dir / f.name)
    _build_caches()                      # finish BEFORE reporting done, so the page
    _job.progress=100; _job.done=True    # reads the ready cache instead of redoing it
    _job.message=f'✓ Done — {kept:,} rows with RFAI date from {len(file_infos)} file(s)'
    if errors: _job.message += ' (skipped: ' + ' | '.join(errors) + ')'
    log.info(_job.message)
    if S3_BUCKET: save_to_storage()
    _job.running=False

# ── In-memory cache builder — v4 PyArrow-safe ─────────────────────────────
def _build_caches():
    """Read all parquet ONCE as one Arrow table, convert every string column to
    pandas Categorical (low memory: ~5x smaller than Python str objects),
    dedup if needed, then share it between the dashboard and journey caches."""
    global _df_cache, _dash_cache, _df_full_cache
    try:
        fs = sorted(_cache_dir.glob("*.parquet")) if _cache_dir.exists() else []
        if not fs: return
        log.info(f"  Building in-memory cache from {len(fs)} parquet file(s)…")
        tables = []
        for f in fs:
            try:
                # read strings straight into dictionary form -> pandas Categorical,
                # never materialising millions of individual Python strings
                str_cols = [fld.name for fld in pq.read_schema(str(f))
                            if pa.types.is_string(fld.type) or pa.types.is_large_string(fld.type)]
                tables.append(pq.read_table(str(f), read_dictionary=str_cols))
            except Exception as e: log.warning(f"  Skipping {f.name}: {e}")
        if not tables:
            log.error("  No parquet files readable"); return
        tbl = pa.concat_tables(tables, promote_options='default'); del tables
        if '_ord' in tbl.column_names: tbl = tbl.drop_columns(['_ord'])
        tbl = tbl.unify_dictionaries()
        df = tbl.to_pandas(strings_to_categorical=True, self_destruct=True); del tbl
        gc.collect(); pa.default_memory_pool().release_unused()

        # dedup (skipped for bulk_convert.py output — already one row per SR)
        if not all(f.name.startswith('rfai_dedup') for f in fs):
            df = _dedup_df(df); gc.collect()

        df = _optimise_df(df)
        _df_full_cache = df
        _df_cache = df[[c for c in _df_cache_cols if c in df.columns]]
        _dash_cache = _compute_dash_aggs(_df_cache)
        gc.collect(); pa.default_memory_pool().release_unused()
        log.info(f"  Cache ready — {len(df):,} unique SRs | {len(df.columns)} cols")

    except Exception as e:
        log.error(f"  Cache build error: {e}")
        import traceback; traceback.print_exc()
        _df_cache=None; _dash_cache=None; _df_full_cache=None

# ── Persistent storage: Cloudflare R2 / AWS S3 ────────────────────────────
# bulk_convert.py --upload puts rfai_dedup.parquet in the bucket; the app
# downloads it on start (and on POST /api/reload) so data survives restarts.
S3_BUCKET = os.environ.get('S3_BUCKET', '')
S3_PREFIX = os.environ.get('S3_PREFIX', 'rfai/').rstrip('/') + '/'

def _s3_region():
    """R2 uses 'auto'; Backblaze B2 needs the region from its endpoint."""
    if os.environ.get('S3_REGION'): return os.environ['S3_REGION']
    m = re.search(r's3\.([a-z0-9-]+)\.backblazeb2\.com', os.environ.get('S3_ENDPOINT', ''))
    return m.group(1) if m else 'auto'

def _s3():
    import boto3
    return boto3.client('s3', endpoint_url=os.environ.get('S3_ENDPOINT') or None,
                        aws_access_key_id=os.environ.get('S3_ACCESS_KEY'),
                        aws_secret_access_key=os.environ.get('S3_SECRET_KEY'),
                        region_name=_s3_region(),
                        config=__import__('botocore.config',fromlist=['Config']).Config(
                            request_checksum_calculation='when_required',
                            s3={'addressing_style': 'path'},
                            response_checksum_validation='when_required'))

def save_to_storage():
    """After a browser upload: store the deduplicated data in the bucket so it
    survives restarts (Render free plan restarts after 15 min idle)."""
    if _df_full_cache is None: return
    try:
        _job.message = 'Saving to cloud storage…'
        out = _cache_dir / 'rfai_dedup.parquet.save'
        df = _df_full_cache.copy()
        for c in df.columns:
            if c in ('delay', 'priDays'): df[c] = df[c].astype('int16')
            else: df[c] = df[c].astype(str).replace({'nan': ''})
        pq.write_table(pa.Table.from_pandas(df[[c for c in OUT if c in df.columns]],
                                            preserve_index=False),
                       str(out), compression='zstd')
        del df
        _s3().upload_file(str(out), S3_BUCKET, S3_PREFIX + 'rfai_dedup.parquet')
        out.unlink()
        _job.message += ' ✓ saved'
        log.info(f"  Saved to s3://{S3_BUCKET}/{S3_PREFIX}rfai_dedup.parquet")
    except Exception as e:
        log.error(f"  Save to storage failed: {e}")
        _job.message = f'Loaded, but saving to storage failed: {e}'

def load_stored_data():
    """Pull parquet from the bucket (if configured), then build caches from disk."""
    global _df_cache, _df_full_cache, _dash_cache
    if _job.running: return
    _job.reset(); _job.running = True
    try:
        _cache_dir.mkdir(parents=True, exist_ok=True)
        if S3_BUCKET:
            _job.message = 'Downloading data from storage…'
            s3 = _s3(); keys = []
            for page in s3.get_paginator('list_objects_v2').paginate(Bucket=S3_BUCKET, Prefix=S3_PREFIX):
                keys += [o['Key'] for o in page.get('Contents', []) if o['Key'].endswith('.parquet')]
            if keys:
                for f in _cache_dir.glob('*.parquet'): f.unlink()
                for k in keys:
                    dst = _cache_dir / Path(k).name
                    log.info(f"  Downloading s3://{S3_BUCKET}/{k}")
                    s3.download_file(S3_BUCKET, k, str(dst) + '.part')
                    os.replace(str(dst) + '.part', dst)
            else:
                log.info(f"  No parquet under s3://{S3_BUCKET}/{S3_PREFIX} yet")
        if any(_cache_dir.glob('*.parquet')):
            _job.message = 'Loading data…'
            _df_cache = _df_full_cache = _dash_cache = None; gc.collect()
            _build_caches()
            n = len(_df_cache) if _df_cache is not None else 0
            _job.rows_processed = _job.rows_kept = n
            _job.message = f'✓ Loaded {n:,} SRs from storage'
        else:
            _job.message = 'No stored data yet'
        _job.done = True
    except Exception as e:
        log.error(f"  Storage load failed: {e}")
        import traceback; traceback.print_exc()
        _job.error = str(e); _job.message = 'Storage load failed: ' + str(e)
    finally:
        _job.progress = 100; _job.running = False
        log.info('  ' + _job.message)

# ── Dashboard aggregation ──────────────────────────────────────────────────
def _compute_dash_aggs(df):
    try: return _compute_dash_aggs_inner(df)
    except Exception as e:
        log.error(f"  _compute_dash_aggs error: {e}")
        return {"total":len(df),"statuses":{},"byMonth":[],"byCircle":[],
                "byOem":[],"byCat":[],"agingBuckets":{},
                "avgDelay":None,"avgDelayPending":None}

_PEND_ST = ('OnAir Pending','RFAI Pending','Open PRI')
_BUCKETS = ('0-30','30-60','60-90','>90')

def _bucket_counts(d):
    """d: numpy int array of delays (>=0). Returns bucket dict."""
    return {'0-30': int((d<=30).sum()), '30-60': int(((d>30)&(d<=60)).sum()),
            '60-90': int(((d>60)&(d<=90)).sum()), '>90': int((d>90).sum())}

def _group_aggs(keys, st, delay, is_pend, is_oa):
    """Vectorised per-group stats: status counts, pending aging buckets,
    OnAir delay buckets and averages. One pass, no per-row Python."""
    out = {}
    codes, uniq = pd.factorize(keys, sort=False)
    order = np.argsort(codes, kind='stable')
    bounds = np.searchsorted(codes[order], np.arange(len(uniq)+1))
    for gi, k in enumerate(uniq):
        if not k: continue
        idx = order[bounds[gi]:bounds[gi+1]]
        st_g = st[idx]; d_g = delay[idx]
        vals, cnts = np.unique(st_g, return_counts=True)
        pd_mask = is_pend[idx] & (d_g >= 0); oa_mask = is_oa[idx] & (d_g >= 0)
        pdv = d_g[pd_mask]; oav = d_g[oa_mask]; anyv = d_g[d_g >= 0]
        out[k] = {"total": int(len(idx)),
                  "statuses": {str(v): int(c) for v, c in zip(vals, cnts)},
                  "agingBreakdown": _bucket_counts(pdv),
                  "avgPendAging": int(round(pdv.mean())) if len(pdv) else None,
                  "oaAging": _bucket_counts(oav),
                  "avgOaDelay": int(round(oav.mean())) if len(oav) else None,
                  "avgDelay": int(round(anyv.mean())) if len(anyv) else None}
    return out

def _compute_dash_aggs_inner(df):
    MON_O=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec']
    def mk_sort(m):
        if not m or "'" not in m: return 0
        p=m.split("'")
        try: return int('20'+p[1])*12+MON_O.index(p[0][:3].lower())
        except: return 0

    total=len(df)
    if not total:
        return {"total":0,"statuses":{},"byMonth":[],"byCircle":[],
                "byOem":[],"byCat":[],"agingBuckets":{},
                "avgDelay":None,"avgDelayPending":None}

    st    = df['srStatus'].astype(str).to_numpy()
    delay = pd.to_numeric(df['delay'], errors='coerce').fillna(-1).astype(np.int32).to_numpy()
    is_pend = np.isin(st, _PEND_ST); is_oa = st == 'OnAir'
    def col(c): return df[c].astype(str).replace({'nan':'','None':''}).to_numpy()

    vals, cnts = np.unique(st, return_counts=True)
    statuses = {str(v): int(c) for v, c in zip(vals, cnts)}

    bm = _group_aggs(col('exitMK'), st, delay, is_pend, is_oa)
    byMonth = [dict(mk=k, **v) for k, v in bm.items()]
    byMonth.sort(key=lambda x: mk_sort(x['mk']))

    bc = _group_aggs(col('circle'), st, delay, is_pend, is_oa)
    byCircle = [dict(circle=k, oaRate=round(v['statuses'].get('OnAir',0)/v['total']*100), **v)
                for k, v in bc.items()]
    byCircle.sort(key=lambda x: -x['total']); byCircle = byCircle[:20]

    bo = _group_aggs(col('oem'), st, delay, is_pend, is_oa)
    byOem = sorted([dict(oem=k, **v) for k, v in bo.items()], key=lambda x: -x['total'])[:15]

    bk = _group_aggs(col('category'), st, delay, is_pend, is_oa)
    byCat = sorted([dict(category=k, **v) for k, v in bk.items()], key=lambda x: -x['total'])

    valid = delay >= 0
    oa_d = delay[is_oa & valid]; pend_d = delay[is_pend & valid]; all_d = delay[valid]
    return {"total":total,"statuses":statuses,"byMonth":byMonth,
            "byCircle":byCircle,"byOem":byOem,"byCat":byCat,
            "agingBuckets":_bucket_counts(oa_d),
            "pendAgingBuckets":_bucket_counts(pend_d),
            "avgDelay":    int(round(all_d.mean())) if len(all_d) else None,
            "avgOaDelay":  int(round(oa_d.mean())) if len(oa_d) else None,
            "avgDelayPending":int(round(pend_d.mean())) if len(pend_d) else None}

# ── HTML row builder ───────────────────────────────────────────────────────
SCS={'OnAir':'bo','OnAir Pending':'bp','Rejected RFAI':'br-rej',
     'Open PRI':'bb','RFAI Pending':'bg2'}
GCS={'0-30':'g0 gb','30-60':'g2 gb','60-90':'g4 gb','>90':'g5 gb','N/A':'gn'}
GLS={'0-30':'0–30','30-60':'30–60','60-90':'60–90','>90':'>90','N/A':'N/A'}

def mk_sort_val(m):
    if not m or "'" not in m: return 0
    p=m.split("'")
    try: return int('20'+p[1])*12+['jan','feb','mar','apr','may','jun',
         'jul','aug','sep','oct','nov','dec'].index(p[0][:3].lower())
    except: return 0

def row_html(r, months=None):
    pf=r.get('priFirst',''); pd2=r.get('priDays',0)
    base=(f'<tr>'
          f'<td class="mn">{r.get("srNo","")}</td><td class="mn">{r.get("siteId","")}</td>'
          f'<td>{r.get("circle","")}</td><td>{r.get("toco","")}</td>'
          f'<td>{r.get("domain","")}</td><td>{r.get("oem","")}</td>'
          f'<td>{r.get("tsp","")}</td><td>{r.get("partner","")}</td>'
          f'<td>{r.get("bPrj","")}</td><td>{r.get("category","")}</td>'
          f'<td class="mn">{r.get("rfaiNotice","")}</td>'
          f'<td class="mn">{r.get("rfaiRej","")}</td>'
          f'<td class="mn">{r.get("rfaiReoffer","")}</td>'
          f'<td class="mn">{r.get("finalRfai","")}</td>'
          f'<td class="mn">{r.get("priFirst","")}</td>'
          f'<td class="mn">{r.get("priLast","")}</td>'
          f'<td class="mn tc">{pd2 if pf and pd2 else ""}</td>'
          f'<td class="mn">{r.get("onAir","")}</td>'
          f'<td class="mn">{r.get("baseDate","")}</td>'
          f'<td style="font-size:11px;color:var(--ac);font-weight:600">'
          f'{r.get("baseLabel","")}</td>')
    if not months:
        st=r.get('srStatus',''); sc=SCS.get(st,'bg2')
        stl=('✕ '+st) if st=='Rejected RFAI' else st
        dl=r.get('delay',-1); dv='' if dl==-1 else str(dl)
        grp=r.get('group','N/A'); gc2=GCS.get(grp,'gn'); gl=GLS.get(grp,'N/A')
        base+=(f'<td class="tm"><span class="badge {sc}">{stl}</span></td>'
               f'<td class="tm tc mn" style="font-weight:600;color:var(--tx2)">'
               f'{dv if dv else "—"}</td>'
               f'<td class="tm tc"><span class="{gc2}">{gl}</span></td>')
        return base+'</tr>'
    jmk_r=r.get('journeyMK',''); emk_r=r.get('exitMK','')
    js=mk_sort_val(jmk_r) if jmk_r else 0
    es=mk_sort_val(emk_r) if emk_r else 999999
    st_final=r.get('srStatus',''); dl_cur=r.get('delay',-1); grp_cur=r.get('group','N/A')
    oa_str=r.get('onAir',''); oa_mk_str=''
    if oa_str and st_final=='OnAir':
        try:
            fp=oa_str.split('-')
            if len(fp)==3:
                fmo=MON_IDX.get(fp[1].lower())
                if fmo is not None: oa_mk_str=f"{MON[fmo]}'{str(int(fp[2]))[2:]}"
        except Exception: pass
    oa_ms=mk_sort_val(oa_mk_str) if oa_mk_str else 999999
    anchor_d=None
    for af in [r.get('finalRfai',''),r.get('rfaiNotice','')]:
        if af:
            try:
                fp=af.split('-')
                if len(fp)==3:
                    fmo=MON_IDX.get(fp[1].lower())
                    if fmo is not None:
                        anchor_d=date(int(fp[2]),fmo+1,int(fp[0])); break
            except Exception: pass
    pri_days_val=r.get('priDays',0) or 0
    for mk_col in months:
        ms=mk_sort_val(mk_col)
        if js==0 or ms<js or ms>es:
            base+='<td class="tm"></td><td class="tm"></td><td class="tm"></td>'
        elif st_final=='OnAir' and oa_mk_str:
            if ms<oa_ms:
                try:
                    import calendar as _cal
                    col_p=mk_col.split("'")
                    mo_i=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].index(col_p[0][:3].lower())
                    yr_c=int('20'+col_p[1]); last_d=_cal.monthrange(yr_c,mo_i+1)[1]
                    mend=date(yr_c,mo_i+1,last_d)
                    if anchor_d and mend>=anchor_d:
                        live_d=max(0,(mend-anchor_d).days-(pri_days_val if isinstance(pri_days_val,int) else 0))
                        lg='0-30' if live_d<=30 else '30-60' if live_d<=60 else '60-90' if live_d<=90 else '>90'
                        gc2=GCS.get(lg,'gn'); gl2=GLS.get(lg,'N/A'); dv2=str(live_d)
                    else: gc2='gn'; gl2='N/A'; dv2='—'
                except Exception: gc2='gn'; gl2='N/A'; dv2='—'
                sc2=SCS.get('OnAir Pending','bg2')
                base+=(f'<td class="tm"><span class="badge {sc2}">OnAir Pending</span></td>'
                       f'<td class="tm tc mn" style="font-weight:600;color:var(--tx2)">{dv2}</td>'
                       f'<td class="tm tc"><span class="{gc2}">{gl2}</span></td>')
            elif ms==oa_ms:
                sc2=SCS.get('OnAir','bg2'); dv2='' if dl_cur==-1 else str(dl_cur)
                gc2=GCS.get(grp_cur,'gn'); gl2=GLS.get(grp_cur,'N/A')
                base+=(f'<td class="tm"><span class="badge {sc2}">OnAir</span></td>'
                       f'<td class="tm tc mn" style="font-weight:600;color:var(--tx2)">'
                       f'{dv2 if dv2 else "—"}</td>'
                       f'<td class="tm tc"><span class="{gc2}">{gl2}</span></td>')
            else:
                base+='<td class="tm"></td><td class="tm"></td><td class="tm"></td>'
        else:
            sc=SCS.get(st_final,'bg2')
            stl=('✕ '+st_final) if st_final=='Rejected RFAI' else st_final
            dv='' if dl_cur==-1 else str(dl_cur)
            gc2=GCS.get(grp_cur,'gn'); gl=GLS.get(grp_cur,'N/A')
            base+=(f'<td class="tm"><span class="badge {sc}">{stl}</span></td>'
                   f'<td class="tm tc mn" style="font-weight:600;color:var(--tx2)">'
                   f'{dv if dv else "—"}</td>'
                   f'<td class="tm tc"><span class="{gc2}">{gl}</span></td>')
    return base+'</tr>'

# ── FastAPI ────────────────────────────────────────────────────────────────
app=FastAPI()
app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_methods=["*"],allow_headers=["*"])
from fastapi import Request
from fastapi.responses import JSONResponse as _JR

VIEW_PASSWORD  = os.environ.get('VIEW_PASSWORD', '')   # set = whole site needs login (default: public)
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', '')  # set = needed to upload data
# On a public server (Render) browser uploads are blocked unless ADMIN_PASSWORD
# is set, so visitors can view the dashboard but never replace the data.
_ON_SERVER     = bool(os.environ.get('RENDER'))
_ADMIN_PATHS   = ('/api/upload',)

def _basic_pw(request):
    import base64
    h = request.headers.get('authorization', '')
    if not h.startswith('Basic '): return None
    try: return base64.b64decode(h[6:]).decode('utf-8', 'replace').split(':', 1)[1]
    except Exception: return None

def _pw_ok(given, expected):
    import secrets
    return given is not None and secrets.compare_digest(given.encode(), expected.encode())

@app.middleware("http")
async def _auth(request:Request,call_next):
    from fastapi.responses import Response
    path = request.url.path
    if path in _ADMIN_PATHS and (ADMIN_PASSWORD or _ON_SERVER):
        if not ADMIN_PASSWORD:
            return JSONResponse({"error":"Upload is locked: ADMIN_PASSWORD is not set on the server. "
                                 "In Render open rfai-tracker → Environment, add ADMIN_PASSWORD, "
                                 "then click 'Save, rebuild, and deploy'."}, status_code=403)
        if not _pw_ok(_basic_pw(request), ADMIN_PASSWORD):
            return Response('Admin login required', status_code=401,
                            headers={'WWW-Authenticate': 'Basic realm="RFAI admin"'})
    elif VIEW_PASSWORD and path != '/api/ping':
        if not _pw_ok(_basic_pw(request), VIEW_PASSWORD):
            return Response('Login required', status_code=401,
                            headers={'WWW-Authenticate': 'Basic realm="RFAI"'})
    return await call_next(request)

@app.on_event("startup")
def _startup_load():
    threading.Thread(target=load_stored_data, daemon=True).start()

@app.post("/api/reload")
def reload_data():
    if _job.running: return JSONResponse({"error":"Busy"},status_code=409)
    threading.Thread(target=load_stored_data, daemon=True).start()
    return {"ok":True}

@app.middleware("http")
async def _catch_all(request:Request,call_next):
    try: return await call_next(request)
    except Exception as e:
        log.error(f"  Unhandled {request.url.path}: {e}")
        return _JR({"error":str(e),"path":str(request.url.path)},status_code=500)

@app.post("/api/upload")
async def upload(background_tasks:BackgroundTasks,files:List[UploadFile]=File(...)):
    if _job.running: return JSONResponse({"error":"Already processing"},status_code=409)
    # stream to disk (not RAM) — the free server only has 512 MB
    up_dir=_cache_dir/'_uploads'; up_dir.mkdir(parents=True,exist_ok=True)
    import shutil, uuid
    file_infos=[]
    for f in files:
        dst=up_dir/f"{uuid.uuid4().hex}{Path(f.filename or 'x').suffix.lower()}"
        with open(dst,'wb') as out: shutil.copyfileobj(f.file,out,1024*1024)
        file_infos.append((f.filename,str(dst)))
    background_tasks.add_task(process_files_bg,file_infos)
    return {"ok":True,"files":[f[0] for f in file_infos],"count":len(file_infos)}

@app.get("/api/progress")
def progress():
    return {"running":_job.running,"done":_job.done,"error":_job.error,
            "progress":_job.progress,"message":_job.message,
            "current_file":_job.current_file,"total_files":_job.total_files,
            "rows_processed":_job.rows_processed,"rows_kept":_job.rows_kept}

@app.get("/api/ping")
def ping(): return {"ok":True}

@app.get("/api/config")
def config():
    """Non-secret setup status, shown on the start screen."""
    env_names = sorted(k for k in os.environ if 'PASS' in k.upper())
    return {"upload_enabled": bool(ADMIN_PASSWORD) or not _ON_SERVER,
            "storage_configured": bool(S3_BUCKET),
            "password_like_vars": env_names}

@app.get("/api/summary")
def summary():
    fs=sorted(_cache_dir.glob("*.parquet")) if _cache_dir.exists() else []
    if not fs: return JSONResponse({"error":"No data"},status_code=404)
    cached=_df_cache is not None
    if cached:
        df=_df_cache
    else:
        parts=[]
        for f in fs:
            try: parts.append(_read_pq(f,columns=['circle','category','bPrj','toco',
                'domain','oem','tsp','srStatus','exitMK','journeyMK','srNo']))
            except Exception: pass
        df=pd.concat(parts,ignore_index=True) if parts else pd.DataFrame()

    total_rows=len(df); rows_no_sr=int((df['srNo']=='').sum())
    unique_srs=df['srNo'].replace('',pd.NA).dropna().nunique()
    dup_sr_rows=(total_rows-rows_no_sr)-unique_srs
    no_rfai_drop=max(0,_job.rows_processed-total_rows)
    dedup=df if cached else _dedup_df(df.copy()); status_counts=dedup['srStatus'].value_counts().to_dict()
    total_unique=unique_srs+rows_no_sr

    def _sk(m):
        if not m or "'" not in m: return 0
        p=m.split("'")
        try: return int('20'+p[1])*12+['jan','feb','mar','apr','may','jun',
             'jul','aug','sep','oct','nov','dec'].index(p[0][:3].lower())
        except: return 0

    mons=sorted(df['exitMK'].dropna().replace('',pd.NA).dropna().unique().tolist(),key=_sk)
    ST_ORD=['OnAir','OnAir Pending','Open PRI','Rejected RFAI','RFAI Pending']
    rfai_by_month=[]
    for jm in sorted(dedup['journeyMK'].replace('',pd.NA).dropna().unique().tolist(),key=_sk):
        grp=dedup[dedup['journeyMK']==jm]
        row={'rfai_month':jm,'total':len(grp)}
        for st in ST_ORD: row[st]=int((grp['srStatus']==st).sum())
        rfai_by_month.append(row)

    def _ul(col): return sorted(df[col].dropna().replace('',pd.NA).dropna().unique().tolist())
    return {"total":total_unique,"total_rows":total_rows,"unique_srs":unique_srs,
            "rows_no_sr":rows_no_sr,"duplicate_rows":dup_sr_rows,
            "total_scanned":_job.rows_processed,"no_rfai_dropped":no_rfai_drop,
            "statuses":status_counts,"rfai_by_month":rfai_by_month,
            "file_stats":getattr(_job,'file_stats',[]),
            "circles":_ul('circle'),"categories":_ul('category'),
            "projects":_ul('bPrj'),"tocos":_ul('toco'),
            "domains":_ul('domain'),"oems":_ul('oem'),"months":mons}

@app.get("/api/filter")
def filter_data(
    circle:str=Query(default=''),category:str=Query(default=''),
    project:str=Query(default=''),toco:str=Query(default=''),
    domain:str=Query(default=''),oem:str=Query(default=''),
    status:str=Query(default=''),month:str=Query(default=''),
    search:str=Query(default=''),page:int=Query(default=1),
    page_size:int=Query(default=75)):
    cached=_df_full_cache is not None
    if cached:
        df=_df_full_cache
    else:
        fs=sorted(_cache_dir.glob("*.parquet")) if _cache_dir.exists() else []
        if not fs: return JSONResponse({"error":"No data"},status_code=404)
        parts=[]
        for f in fs:
            try:
                flt=[]
                if circle:   flt.append(('circle','in',circle.split(',')))
                if status:   flt.append(('srStatus','in',status.split(',')))
                if category: flt.append(('category','in',category.split(',')))
                parts.append(_read_pq(f))
            except Exception: pass
        if not parts: return JSONResponse({"error":"No data"},status_code=404)
        df=pd.concat(parts,ignore_index=True)
    if circle:   df=df[df['circle'].isin(circle.split(','))]
    if category: df=df[df['category'].isin(category.split(','))]
    if project:  df=df[df['bPrj'].isin(project.split(','))]
    if toco:     df=df[df['toco'].isin(toco.split(','))]
    if domain:   df=df[df['domain'].isin(domain.split(','))]
    if oem:      df=df[df['oem'].isin(oem.split(','))]
    if status:   df=df[df['srStatus'].isin(status.split(','))]
    if month:
        ms=set(month.split(','))
        df=df[df['exitMK'].isin(ms)|df['journeyMK'].isin(ms)]
    if search:
        s=search.upper()
        df=df[df['srNo'].astype(str).str.upper().str.contains(s,na=False,regex=False)|
              df['siteId'].astype(str).str.upper().str.contains(s,na=False,regex=False)]
    if len(df) and not cached: df=_dedup_df(df)
    df=df.reset_index(drop=True)
    total=len(df); pages=max(1,(total+page_size-1)//page_size)
    page=max(1,min(page,pages)); s=(page-1)*page_size
    if month:
        show_months=sorted(month.split(','),key=mk_sort_val)
    else:
        all_mks=(set(df['exitMK'].dropna().replace('',pd.NA).dropna().unique())|
                 set(df['journeyMK'].dropna().replace('',pd.NA).dropna().unique()))
        show_months=sorted([m for m in all_mks if m],key=mk_sort_val)
    page_df=df.iloc[s:s+page_size]
    html=''.join(row_html(r.to_dict(),show_months) for _,r in page_df.iterrows())
    return {"html":html,"total":total,"page":page,"pages":pages,
            "start":s+1,"end":min(s+page_size,total),"months":show_months}

@app.get("/api/dashboard")
def dashboard_data(
    circle:str=Query(default=''),category:str=Query(default=''),
    oem:str=Query(default=''),toco:str=Query(default=''),
    domain:str=Query(default=''),status:str=Query(default=''),
    month:str=Query(default='')):
    try: return _dash_inner(circle,category,oem,toco,domain,status,month)
    except Exception as e:
        log.error(f"  /api/dashboard: {e}")
        return {"total":0,"statuses":{},"byMonth":[],"byCircle":[],"byOem":[],
                "byCat":[],"agingBuckets":{},"avgDelay":None,"avgDelayPending":None}

def _dash_inner(circle,category,oem,toco,domain,status,month):
    has_filter=any([circle,category,oem,toco,domain,status,month])
    if not has_filter and _dash_cache is not None: return _dash_cache
    if _df_cache is None:
        fs=sorted(_cache_dir.glob("*.parquet")) if _cache_dir.exists() else []
        if not fs: return JSONResponse({"error":"No data"},status_code=404)
        _build_caches()
    if _df_cache is None: return JSONResponse({"error":"Cache unavailable"},status_code=503)
    df=_df_cache
    if circle:   df=df[df['circle'].isin(circle.split(','))]
    if category: df=df[df['category'].isin(category.split(','))]
    if oem:      df=df[df['oem'].isin(oem.split(','))]
    if toco:     df=df[df['toco'].isin(toco.split(','))]
    if domain:   df=df[df['domain'].isin(domain.split(','))]
    if status:   df=df[df['srStatus'].isin(status.split(','))]
    if month:
        ms=set(month.split(','))
        df=df[df['exitMK'].isin(ms)|df['journeyMK'].isin(ms)]
    if not len(df):
        return {"total":0,"statuses":{},"byMonth":[],"byCircle":[],"byOem":[],
                "byCat":[],"agingBuckets":{},"avgDelay":None,"avgDelayPending":None}
    return _compute_dash_aggs(df)

@app.get("/api/rows")
def get_rows():
    return {"rows":[],"total":0,"capped":False,"deprecated":True}

@app.get("/api/drill")
def drill_data(
    circle:str=Query(default=''),status:str=Query(default=''),
    aging:str=Query(default=''),category:str=Query(default=''),
    limit:int=Query(default=200)):
    DCOLS=['srNo','siteId','circle','toco','oem','tsp','bPrj','domain','category',
           'srStatus','delay','group','exitMK','rfaiNotice','finalRfai','onAir']
    if _df_full_cache is not None:
        avail=[c for c in DCOLS if c in _df_full_cache.columns]
        df=_df_full_cache[avail].copy()
        if circle:   df=df[df['circle'].isin(circle.split(','))]
        if status:   df=df[df['srStatus'].isin(status.split(','))]
        if category: df=df[df['category'].isin(category.split(','))]
    else:
        fs=sorted(_cache_dir.glob("*.parquet")) if _cache_dir.exists() else []
        if not fs: return JSONResponse({"error":"No data"},status_code=404)
        parts=[]
        for f in fs:
            try: parts.append(_read_pq(f,columns=DCOLS))
            except Exception: pass
        if not parts: return {"rows":[],"total":0}
        df=pd.concat(parts,ignore_index=True)
        if circle:   df=df[df['circle'].isin(circle.split(','))]
        if status:   df=df[df['srStatus'].isin(status.split(','))]
        if category: df=df[df['category'].isin(category.split(','))]
    if aging:
        bounds={'0-30':(0,30),'30-60':(31,60),'60-90':(61,90),'>90':(91,9999)}
        lo,hi=bounds.get(aging,(0,9999))
        dv=pd.to_numeric(df['delay'],errors='coerce').where(df['delay']!=-1)
        df=df[(dv>=lo)&(dv<=hi)]
    df=df.sort_values('delay',ascending=False).head(limit)
    OUTC=['srNo','siteId','circle','toco','oem','category','srStatus','group',
          'exitMK','rfaiNotice','finalRfai','onAir']
    out=pd.DataFrame({c:(df[c].astype(str).replace({'nan':'','None':''}).to_numpy()
                         if c in df.columns else np.full(len(df),''))
                      for c in OUTC})
    out['group']=out['group'].replace('', 'N/A')
    dv=pd.to_numeric(df['delay'],errors='coerce').to_numpy() if 'delay' in df.columns else np.full(len(df),np.nan)
    out['delay']=[None if (v!=v or v==-1) else int(v) for v in dv]
    rows=out.to_dict('records')
    return {"rows":rows,"total":len(rows)}

@app.get("/")
def root():
    if HTML_FILE.exists(): return FileResponse(str(HTML_FILE))
    return JSONResponse({"error":"HTML not found"},status_code=404)

@app.get("/{path:path}")
def static(path:str):
    f=(SCRIPT_DIR/path).resolve()
    if (f.is_file() and f.is_relative_to(SCRIPT_DIR.resolve())
            and f.suffix.lower() in ('.html','.css','.js','.png','.jpg','.svg','.ico','.json')):
        return FileResponse(str(f))
    return JSONResponse({"error":"Not found"},status_code=404)

if __name__=="__main__":
    print("\n  ╔══════════════════════════════════════════════════╗")
    print("  ║  RFAI Backend — v4 PyArrow-Safe Build           ║")
    print("  ║  Open: http://localhost:8765                    ║")
    print("  ╚══════════════════════════════════════════════════╝\n")
    if not os.environ.get("RENDER"):   # local run only — no browser on a server
        import webbrowser
        threading.Thread(target=lambda:(time.sleep(1.2),webbrowser.open(f"http://localhost:{PORT}")),
                         daemon=True).start()
    uvicorn.run(app,host="0.0.0.0",port=PORT,log_level="warning",timeout_keep_alive=600)
