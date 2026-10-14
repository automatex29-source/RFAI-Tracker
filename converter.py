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
FLUSH = 10_000   # converted rows held in RAM before writing to disk
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

def _mem_limit_bytes():
    """Memory this process may use: container limit (Render) or free RAM."""
    if os.environ.get('MEMORY_LIMIT_MB'):
        return int(os.environ['MEMORY_LIMIT_MB']) << 20
    if os.environ.get('RENDER'):
        # Render shows the HOST's memory in /proc and cgroups, not the plan's limit;
        # assume the free plan (512 MB) so big files use the low-memory reader.
        return 512 << 20
    for f in ('/sys/fs/cgroup/memory.max', '/sys/fs/cgroup/memory/memory.limit_in_bytes'):
        try:
            v = open(f).read().strip()
            if v.isdigit() and int(v) < 1 << 50: return int(v)
        except Exception: pass
    try:
        for line in open('/proc/meminfo'):
            if line.startswith('MemAvailable:'): return int(line.split()[1]) * 1024
    except Exception: pass
    return 8 << 30   # Windows PC etc.: assume plenty

def _xlsx_stream(path):
    """Low-memory .xlsx reader (lxml iterparse): ~20 MB RAM for any file size.
    Dates come out as Excel serial numbers, which the date parser understands."""
    import zipfile
    from lxml import etree
    NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    z = zipfile.ZipFile(str(path))
    names = z.namelist()
    ss = []
    if 'xl/sharedStrings.xml' in names:
        for _, el in etree.iterparse(z.open('xl/sharedStrings.xml'), tag=NS + 'si'):
            ss.append(''.join(t.text or '' for t in el.iter(NS + 't'))); el.clear()

    def col_idx(ref):
        n = 0
        for ch in ref:
            if 'A' <= ch <= 'Z': n = n * 26 + ord(ch) - 64
            else: break
        return n - 1

    def rows(member):
        for _, row in etree.iterparse(z.open(member), tag=NS + 'row'):
            out = []
            for c in row.iterchildren(NS + 'c'):
                ref = c.get('r'); i = col_idx(ref) if ref else len(out)
                while len(out) < i: out.append(None)
                t = c.get('t'); v = c.find(NS + 'v')
                if t == 's': val = ss[int(v.text)] if v is not None and v.text else None
                elif t == 'inlineStr': val = ''.join(x.text or '' for x in c.iter(NS + 't'))
                elif v is None or v.text is None: val = None
                elif t in ('str', 'e'): val = v.text
                elif t == 'b': val = v.text == '1'
                else:
                    x = v.text
                    try: f = float(x); val = int(f) if f.is_integer() else f
                    except ValueError: val = x
                out.append(val)
            row.clear()
            while row.getprevious() is not None: del row.getparent()[0]
            yield tuple(out)

    sheets = sorted((n for n in names if n.startswith('xl/worksheets/sheet') and n.endswith('.xml')),
                    key=lambda n: int(re.sub(r'\D', '', n) or 0))
    best, best_sc = sheets[0], -1
    if len(sheets) > 1:
        for sh in sheets:
            hdr = next(rows(sh), ())
            sc = sum(1 for c in hdr if c and any(k in str(c).upper() for k in _SHEET_KEYS))
            if sc > best_sc: best, best_sc = sh, sc
    yield from rows(best)

def iter_excel_rows(path):
    """Yield rows (tuples) of the best-matching sheet, header row first.
    • calamine (Rust, fastest) when the file fits comfortably in memory
      (it needs ~16x the .xlsx size in RAM);
    • streaming lxml reader for big .xlsx on small servers (Render free 512 MB);
    • openpyxl as last resort."""
    size = os.path.getsize(path)
    ext = Path(str(path)).suffix.lower()
    fits = size * 16 < _mem_limit_bytes() * 0.5
    try:
        from python_calamine import CalamineWorkbook
    except ImportError:
        CalamineWorkbook = None
    if ext in ('.xlsx', '.xlsm') and not fits:
        try:
            import lxml  # noqa: F401
            log.info(f"  Large file ({size/2**20:.0f} MB) — using low-memory reader")
            yield from _xlsx_stream(path); return
        except ImportError:
            CalamineWorkbook = None   # fall through to openpyxl streaming
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
        # read only the columns the tool uses (often ~20 of 60+): far less memory
        headers = list(pd.read_csv(path, sep=delim, nrows=0, encoding=enc,
                                   encoding_errors='replace').columns)
        full_cm = map_headers(headers)
        log.info(f"  Matched {sum(1 for v in full_cm.values() if v is not None)}/{len(COLS)} cols")
        used = sorted({i for i in full_cm.values() if i is not None})
        pos = {orig: k for k, orig in enumerate(used)}
        cm = {col: (pos[i] if i is not None else None) for col, i in full_cm.items()}
        reader = pd.read_csv(
            path, sep=delim, dtype=str, keep_default_na=False, encoding=enc, usecols=used,
            encoding_errors='replace', chunksize=20000, low_memory=False, on_bad_lines='skip')
    except Exception as e:
        log.error(f"  CSV open failed {fname}: {e}")
        return grand, kept, chunk_idx

    batch = []
    for chunk_df in reader:
        chunk_df = chunk_df[[headers[i] for i in used]] if len(used) else chunk_df
        for row in chunk_df.itertuples(index=False, name=None):
            grand += 1
            r = process_row(row, cm)
            if r: batch.append(r); kept += 1
        _job.rows_processed = grand; _job.rows_kept = kept
        _job.message = f'{fname} — {grand:,} rows read'
        if len(batch) >= FLUSH:
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
    _current_raw_keys.clear()            # website upload replaces all bucket Excel files
    try: _process_files(file_infos)
    except Exception as e:
        log.error(f"  Upload processing crashed: {e}")
        import traceback; traceback.print_exc()
        _job.error=f'Processing failed: {e}'
    finally:
        _job.running=False

def _process_files(file_infos, source='website upload'):
    global _df_cache, _df_full_cache, _dash_cache
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
                    if len(batch)>=FLUSH:
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
    _set_meta([n for n, _ in file_infos], source)
    _job.progress=100; _job.done=True    # reads the ready cache instead of redoing it
    _job.message=f'✓ Done — {kept:,} rows with RFAI date from {len(file_infos)} file(s)'
    if errors: _job.message += ' (skipped: ' + ' | '.join(errors) + ')'
    log.info(_job.message)
    if S3_BUCKET: save_to_storage()
    if SQL_ENABLED: save_to_sql()
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
def _clean_storage_env():
    """Forgive copy-paste slips in the Render settings: spaces, quotes,
    a missing https://, 's3://' or slashes around the bucket name."""
    for k in ('S3_ENDPOINT', 'S3_BUCKET', 'S3_ACCESS_KEY', 'S3_SECRET_KEY', 'S3_REGION', 'S3_PREFIX'):
        v = os.environ.get(k)
        if v is None: continue
        v = v.strip().strip('"').strip("'").strip()
        if k == 'S3_ENDPOINT' and v:
            if not v.startswith(('http://', 'https://')): v = 'https://' + v
            v = v.rstrip('/')
        if k == 'S3_BUCKET': v = v.replace('s3://', '').strip('/')
        os.environ[k] = v
_clean_storage_env()

S3_BUCKET = os.environ.get('S3_BUCKET', '')
S3_PREFIX = os.environ.get('S3_PREFIX', 'rfai/').rstrip('/') + '/'

def _s3_region():
    """R2 uses 'auto'; Backblaze B2 needs the region from its endpoint."""
    if os.environ.get('S3_REGION'): return os.environ['S3_REGION']
    m = re.search(r's3\.([a-z0-9-]+)\.backblazeb2\.com', os.environ.get('S3_ENDPOINT', ''))
    return m.group(1) if m else 'auto'

def _s3():
    import boto3
    c = boto3.client('s3', endpoint_url=os.environ.get('S3_ENDPOINT') or None,
                        aws_access_key_id=os.environ.get('S3_ACCESS_KEY'),
                        aws_secret_access_key=os.environ.get('S3_SECRET_KEY'),
                        region_name=_s3_region(),
                        config=__import__('botocore.config',fromlist=['Config']).Config(
                            request_checksum_calculation='when_required',
                            s3={'addressing_style': 'path'},
                            response_checksum_validation='when_required'))
    c.meta.events.register('after-call.s3', _capture_body)
    return c

_storage_status = {'last_save': None, 'last_load': None}
# What is loaded right now — shown on the page header ("Book2.xlsx · 13-Oct 14:05")
_data_meta = {}
# Excel/CSV keys in the bucket that make up the current data (kept); any other
# Excel/CSV in the bucket is older history and is deleted after a successful save.
_current_raw_keys = set()

def _set_meta(files, source):
    _data_meta.clear()
    _data_meta.update({'files': list(files), 'source': source,
                       'uploaded_at': datetime.now().strftime('%d-%b-%Y %H:%M'),
                       'srs': len(_df_cache) if _df_cache is not None else 0})

def _s3_err(e):
    """Readable storage error: HTTP status + code + message + raw reply."""
    r = getattr(e, 'response', None) or {}
    meta = r.get('ResponseMetadata', {}); err = r.get('Error', {})
    msg = f"HTTP {meta.get('HTTPStatusCode','?')} {err.get('Code','')} {err.get('Message','')}".strip()
    if _last_s3_body.get('body'): msg += ' | reply: ' + _last_s3_body['body']
    return msg if r else str(e)

_last_s3_body = {}
def _capture_body(http_response=None, **kw):
    try:
        if http_response is not None and http_response.status_code >= 400:
            _last_s3_body['body'] = http_response.content[:300].decode('utf-8', 'replace')
    except Exception: pass

def save_to_storage():
    """After a browser upload: copy the converted parquet files to the bucket
    (no extra RAM — files are streamed from disk) and remove older files there,
    so the data survives restarts (Render free plan restarts after 15 min idle)."""
    files = sorted(_cache_dir.glob('*.parquet'))
    if not files: return
    try:
        _job.message = 'Saving to cloud storage…'
        s3 = _s3()
        new_keys = set()
        for f in files:
            key = S3_PREFIX + f.name
            s3.upload_file(str(f), S3_BUCKET, key); new_keys.add(key)
        old = []
        for page in s3.get_paginator('list_objects_v2').paginate(Bucket=S3_BUCKET, Prefix=S3_PREFIX):
            old += [o['Key'] for o in page.get('Contents', [])
                    if o['Key'].endswith('.parquet') and o['Key'] not in new_keys]
        for k in old: s3.delete_object(Bucket=S3_BUCKET, Key=k)
        # remove history: older Excel/CSV files sitting in the bucket
        removed = 0
        for page in s3.get_paginator('list_objects_v2').paginate(Bucket=S3_BUCKET):
            for o in page.get('Contents', []):
                k = o['Key']
                if (not k.startswith(S3_PREFIX) and k.lower().endswith(_RAW_EXTS)
                        and k not in _current_raw_keys):
                    s3.delete_object(Bucket=S3_BUCKET, Key=k); removed += 1
        if removed: log.info(f"  Removed {removed} older Excel/CSV file(s) from storage")
        s3.put_object(Bucket=S3_BUCKET, Key=S3_PREFIX + '_meta.json',
                      Body=json.dumps(_data_meta).encode(), ContentType='application/json')
        _job.message = f'✓ Saved to cloud storage ({len(files)} file(s))'
        _storage_status['last_save'] = f'ok: {len(files)} file(s) at {datetime.now():%d-%b %H:%M}'
        log.info(f"  Saved {len(files)} file(s) to s3://{S3_BUCKET}/{S3_PREFIX}")
    except Exception as e:
        log.error(f"  Save to storage failed: {e}")
        _job.message = f'Loaded, but saving to storage FAILED: {_s3_err(e)}'
        _storage_status['last_save'] = f'FAILED: {_s3_err(e)}'

# ── SQL Server storage (local / office-network mode) ───────────────────────
# Enable with USE_SQL_SERVER=1 (server auto-detected) or SQL_SERVER=<name>.
# Windows Authentication by default; SQL_USER/SQL_PASSWORD for SQL logins.
SQL_DATABASE = (os.environ.get('SQL_DATABASE') or 'RFAI').strip()
SQL_ENABLED  = bool(os.environ.get('SQL_SERVER', '').strip()
                    or os.environ.get('USE_SQL_SERVER', '').strip() in ('1', 'true', 'yes'))
SQL_TABLE    = 'rfai_data'
_sql_server  = {'name': os.environ.get('SQL_SERVER', '').strip() or None}

def _sql_driver():
    import pyodbc
    have = pyodbc.drivers()
    for d in ('ODBC Driver 18 for SQL Server', 'ODBC Driver 17 for SQL Server',
              'ODBC Driver 13 for SQL Server', 'SQL Server Native Client 11.0', 'SQL Server'):
        if d in have: return d
    raise RuntimeError('No SQL Server ODBC driver found. Install "ODBC Driver 18 for SQL Server" '
                       'from Microsoft (free), then restart.')

def _sql_conn_str(server, database):
    drv = _sql_driver()
    cs = f'DRIVER={{{drv}}};SERVER={server};DATABASE={database};'
    if os.environ.get('SQL_USER'):
        cs += f"UID={os.environ['SQL_USER']};PWD={os.environ.get('SQL_PASSWORD', '')};"
    else:
        cs += 'Trusted_Connection=yes;'
    if 'ODBC Driver' in drv: cs += 'TrustServerCertificate=yes;'
    return cs

def _sql_find_server():
    """Use SQL_SERVER if given, else try the usual local instance names."""
    import pyodbc
    if _sql_server['name']: return _sql_server['name']
    pc = os.environ.get('COMPUTERNAME', '')
    tried = []
    for cand in ('localhost', r'localhost\SQLEXPRESS', r'.\SQLEXPRESS', r'(localdb)\MSSQLLocalDB',
                 pc, pc + r'\SQLEXPRESS'):
        if not cand or cand in tried or cand.startswith('\\'): continue
        tried.append(cand)
        try:
            pyodbc.connect(_sql_conn_str(cand, 'master'), timeout=3).close()
            _sql_server['name'] = cand
            log.info(f"  SQL Server found: {cand}")
            return cand
        except Exception:
            continue
    raise RuntimeError('Could not find SQL Server (tried ' + ', '.join(tried) +
                       '). Set SQL_SERVER to the "Server name" shown in SSMS.')

def _sql_connect(database=None, autocommit=False):
    import pyodbc
    return pyodbc.connect(_sql_conn_str(_sql_find_server(), database or SQL_DATABASE),
                          timeout=10, autocommit=autocommit)

def _sql_ensure_db():
    c = _sql_connect('master', autocommit=True)
    try:
        c.execute(f"IF DB_ID(N'{SQL_DATABASE}') IS NULL CREATE DATABASE [{SQL_DATABASE}]")
    finally:
        c.close()

def save_to_sql():
    """Write the current (deduplicated) data to <SQL_DATABASE>.dbo.rfai_data.
    Built in a side table and swapped in at the end, so a failure never
    leaves the old data half-replaced. Each upload is also logged."""
    df = _df_full_cache
    if df is None: return
    try:
        _job.message = 'Saving to SQL Server…'
        _sql_ensure_db()
        cols = [c for c in OUT if c in df.columns]
        ddl = ', '.join(f'[{c}] SMALLINT' if c in ('delay', 'priDays') else f'[{c}] NVARCHAR(400)'
                        for c in cols)
        data = []
        for c in cols:
            if c in ('delay', 'priDays'):
                data.append([int(v) for v in pd.to_numeric(df[c], errors='coerce').fillna(-1)])
            else:
                data.append(df[c].astype(str).replace({'nan': '', 'None': ''}).tolist())
        rows = list(zip(*data)); del data
        con = _sql_connect()
        try:
            cur = con.cursor(); cur.fast_executemany = True
            cur.execute(f"IF OBJECT_ID('dbo.{SQL_TABLE}_new') IS NOT NULL DROP TABLE dbo.{SQL_TABLE}_new")
            cur.execute(f"CREATE TABLE dbo.{SQL_TABLE}_new ({ddl})")
            ins = (f"INSERT INTO dbo.{SQL_TABLE}_new ({', '.join(f'[{c}]' for c in cols)}) "
                   f"VALUES ({', '.join('?' * len(cols))})")
            for i in range(0, len(rows), 5000):
                cur.executemany(ins, rows[i:i + 5000])
            cur.execute(f"IF OBJECT_ID('dbo.{SQL_TABLE}') IS NOT NULL DROP TABLE dbo.{SQL_TABLE}")
            cur.execute(f"EXEC sp_rename 'dbo.{SQL_TABLE}_new', '{SQL_TABLE}'")
            cur.execute("IF OBJECT_ID('dbo.rfai_upload_log') IS NULL CREATE TABLE dbo.rfai_upload_log "
                        "(uploaded_at DATETIME2 DEFAULT SYSDATETIME(), srs INT, files NVARCHAR(1000))")
            cur.execute("INSERT INTO dbo.rfai_upload_log (srs, files) VALUES (?, ?)",
                        len(rows), ', '.join(s_['name'] for s_ in _job.file_stats)[:1000])
            con.commit()
        finally:
            con.close()
        _storage_status['last_save'] = (f'ok: {len(rows):,} SRs → SQL Server '
                                        f'{_sql_server["name"]}/{SQL_DATABASE} at {datetime.now():%d-%b %H:%M}')
        _job.message = f'✓ Saved {len(rows):,} SRs to SQL Server'
        log.info('  ' + _storage_status['last_save'])
    except Exception as e:
        log.error(f"  Save to SQL Server failed: {e}")
        _storage_status['last_save'] = f'FAILED (SQL Server): {e}'
        _job.message = f'Loaded, but saving to SQL Server FAILED: {e}'

def _load_from_sql(force=False):
    """Read dbo.rfai_data into the cache (skipped if unchanged since last load)."""
    global _df_cache, _df_full_cache, _dash_cache
    _job.message = 'Connecting to SQL Server…'
    _sql_ensure_db()
    con = _sql_connect()
    try:
        cur = con.cursor()
        if not cur.execute(f"SELECT OBJECT_ID('dbo.{SQL_TABLE}')").fetchone()[0]:
            _storage_status['last_load'] = f'SQL Server {_sql_server["name"]}/{SQL_DATABASE}: no data yet'
            _job.message = 'No stored data yet'; _job.done = True; return
        try:
            sig = tuple(cur.execute("SELECT COUNT(*), MAX(uploaded_at) FROM dbo.rfai_upload_log").fetchone())
            last = cur.execute("SELECT TOP 1 files, uploaded_at FROM dbo.rfai_upload_log "
                               "ORDER BY uploaded_at DESC").fetchone()
            if last:
                _data_meta.clear()
                _data_meta.update({'files': [f for f in (last[0] or '').split(', ') if f],
                                   'source': 'SQL Server',
                                   'uploaded_at': last[1].strftime('%d-%b-%Y %H:%M')})
        except Exception:
            sig = None
        if sig and _df_cache is not None and _loaded_sig.get('sql') == sig:
            _job.message = f'✓ Already up to date — {len(_df_cache):,} SRs'; _job.done = True; return
        cols = [r[0] for r in cur.execute(
            "SELECT name FROM sys.columns WHERE object_id = OBJECT_ID(?) ORDER BY column_id",
            f'dbo.{SQL_TABLE}').fetchall() if r[0] in OUT]
        _job.message = 'Reading data from SQL Server…'
        cur.execute(f"SELECT {', '.join(f'[{c}]' for c in cols)} FROM dbo.{SQL_TABLE}")
        batches = []
        while True:
            chunk = cur.fetchmany(50000)
            if not chunk: break
            arrays = {}
            for i, c in enumerate(cols):
                vals = [r[i] for r in chunk]
                arrays[c] = (pa.array([v if v is not None else -1 for v in vals], type=pa.int16())
                             if c in ('delay', 'priDays')
                             else pa.array([v or '' for v in vals], type=pa.string()))
            batches.append(pa.table(arrays))
    finally:
        con.close()
    for f in _cache_dir.glob('*.parquet'): f.unlink()
    if batches:
        # 'rfai_dedup' prefix: data in SQL is already one row per SR
        pq.write_table(pa.concat_tables(batches), str(_cache_dir / 'rfai_dedup_sql.parquet'),
                       compression='zstd')
    n_rows = sum(b.num_rows for b in batches); del batches
    _storage_status['last_load'] = (f'SQL Server {_sql_server["name"]}/{SQL_DATABASE}: '
                                    f'{n_rows:,} SRs')
    _df_cache = _df_full_cache = _dash_cache = None; gc.collect()
    if n_rows: _build_caches()
    _loaded_sig['sql'] = sig
    n = len(_df_cache) if _df_cache is not None else 0
    _job.rows_processed = _job.rows_kept = n
    _job.message = f'✓ Loaded {n:,} SRs from SQL Server' if n else 'No stored data yet'
    _job.done = True

_loaded_sig = {}
_RAW_EXTS = ('.xlsx', '.xlsm', '.xls', '.xlsb', '.ods', '.csv')

def load_stored_data(retries=1, force=False):
    """Load data from the bucket:
    • Excel/CSV files you put in the bucket yourself (e.g. via the Supabase
      dashboard) are converted when they are newer than the last conversion;
    • otherwise the converted parquet under S3_PREFIX is loaded (fast)."""
    global _df_cache, _df_full_cache, _dash_cache
    if _job.running: return
    for attempt in range(retries):
        _job.reset(); _job.running = True
        try:
            _load_once(force)
            return
        except Exception as e:
            log.error(f"  Storage load failed (attempt {attempt+1}/{retries}): {_s3_err(e)}")
            import traceback; traceback.print_exc()
            _job.error = _s3_err(e); _job.message = 'Storage load failed: ' + _s3_err(e)
            _storage_status['last_load'] = f'FAILED: {_s3_err(e)}'
        finally:
            _job.progress = 100; _job.running = False
            log.info('  ' + _job.message)
        if attempt + 1 < retries: time.sleep(15)   # storage may still be waking up

def _load_once(force=False):
    global _df_cache, _df_full_cache, _dash_cache
    _cache_dir.mkdir(parents=True, exist_ok=True)
    if SQL_ENABLED and not S3_BUCKET:
        _load_from_sql(force); return
    if S3_BUCKET:
        _job.message = 'Checking storage…'
        s3 = _s3(); objs = []
        for page in s3.get_paginator('list_objects_v2').paginate(Bucket=S3_BUCKET):
            objs += page.get('Contents', [])
        pq_objs  = [o for o in objs if o['Key'].startswith(S3_PREFIX) and o['Key'].endswith('.parquet')]
        raw_objs = [o for o in objs if not o['Key'].startswith(S3_PREFIX)
                    and o['Key'].lower().endswith(_RAW_EXTS)
                    and not o['Key'].rsplit('/', 1)[-1].startswith(('~$', '.'))]
        newest_pq  = max((o['LastModified'] for o in pq_objs), default=None)
        newest_raw = max((o['LastModified'] for o in raw_objs), default=None)
        _storage_status['last_load'] = (f'{len(pq_objs)} converted file(s), '
                                        f'{len(raw_objs)} Excel/CSV file(s) in storage')
        raw_sig = sorted(f"{o['Key']}|{o['LastModified']}|{o['Size']}" for o in raw_objs)
        marker = S3_PREFIX + '_converting.json'
        if raw_objs and (newest_pq is None or newest_raw > newest_pq):
            # Crash-loop guard: if the last start died while converting these exact
            # files (e.g. out of memory), don't try again on every restart.
            try:
                prev = json.loads(s3.get_object(Bucket=S3_BUCKET, Key=marker)['Body'].read())
            except Exception:
                prev = None
            if prev and prev.get('files') == raw_sig and prev.get('attempts', 0) >= 2 and not force:
                _storage_status['last_load'] += ' — conversion skipped (crashed twice)'
                raise RuntimeError('Converting ' + ', '.join(o['Key'] for o in raw_objs) +
                                   ' crashed the server twice (too big to convert on the free server). '
                                   'Upload it on the website instead — Excel is converted in your browser there.')
            attempts = (prev.get('attempts', 0) if prev and prev.get('files') == raw_sig else 0) + 1
            s3.put_object(Bucket=S3_BUCKET, Key=marker,
                          Body=json.dumps({'files': raw_sig, 'attempts': attempts}).encode())
            # Excel/CSV uploaded straight into the bucket → convert them now
            up_dir = _cache_dir / '_uploads'; up_dir.mkdir(parents=True, exist_ok=True)
            # new data = Excel/CSV added since the last conversion; anything older
            # is history and gets deleted after a successful save. First time
            # (nothing converted yet): files added within 15 min of the newest one.
            from datetime import timedelta
            if newest_pq is not None:
                batch = [o for o in raw_objs if o['LastModified'] > newest_pq]
            else:
                batch = [o for o in raw_objs if newest_raw - o['LastModified'] <= timedelta(minutes=15)]
            _current_raw_keys.clear(); _current_raw_keys.update(o['Key'] for o in batch)
            infos = []
            for o in batch:
                name = o['Key'].rsplit('/', 1)[-1]
                dst = up_dir / f"{len(infos):03d}_{name}"
                _job.message = f'Downloading {name} from storage…'
                log.info(f"  Downloading s3://{S3_BUCKET}/{o['Key']}")
                s3.download_file(S3_BUCKET, o['Key'], str(dst))
                infos.append((name, str(dst)))
            _job.total_files = len(infos)
            _process_files(infos, 'Supabase upload')   # converts, builds caches, saves back
            try: s3.delete_object(Bucket=S3_BUCKET, Key=marker)
            except Exception: pass
            if _job.error: raise RuntimeError(_job.error)
            _loaded_sig['sig'] = None      # next reload re-checks storage
            return
        # Fast path: nothing changed in storage since the last load → done in ~1 s
        pq_sig = sorted(f"{o['Key']}|{o.get('ETag')}" for o in pq_objs)
        if pq_objs and _df_cache is not None and _loaded_sig.get('sig') == pq_sig:
            _job.message = f'✓ Already up to date — {len(_df_cache):,} SRs'
            _job.done = True
            return
        if pq_objs:
            _loaded_sig['sig'] = pq_sig
            try:
                _data_meta.clear()
                _data_meta.update(json.loads(s3.get_object(
                    Bucket=S3_BUCKET, Key=S3_PREFIX + '_meta.json')['Body'].read()))
            except Exception:
                pass
            for f in _cache_dir.glob('*.parquet'): f.unlink()
            for o in pq_objs:
                k = o['Key']; dst = _cache_dir / Path(k).name
                _job.message = 'Downloading data from storage…'
                log.info(f"  Downloading s3://{S3_BUCKET}/{k}")
                s3.download_file(S3_BUCKET, k, str(dst) + '.part')
                os.replace(str(dst) + '.part', dst)
        else:
            log.info(f"  No data in s3://{S3_BUCKET} yet")
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

def month_cells(r, months):
    """Status / days / ageing group of one SR in each journey month (None = blank).
    Shared by the Journey Table and the Excel export so both always agree."""
    out=[]
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
            out.append(None)
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
                        out.append(('OnAir Pending',str(live_d),lg))
                    else: out.append(('OnAir Pending','—','N/A'))
                except Exception: out.append(('OnAir Pending','—','N/A'))
            elif ms==oa_ms:
                out.append(('OnAir','' if dl_cur==-1 else str(dl_cur),grp_cur))
            else:
                out.append(None)
        else:
            out.append((st_final,'' if dl_cur==-1 else str(dl_cur),grp_cur))
    return out

def row_html(r, months=None):
    from html import escape as _he
    raw=r                                   # month logic needs the original values (e.g. "Feb'26")
    r={k:(_he(str(v)) if isinstance(v,str) else v) for k,v in r.items()}   # uploaded text is untrusted
    pf=r.get('priFirst',''); pd2=r.get('priDays',0)
    base=(f'<tr data-sr="{r.get("srNo","")}">'
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
    for c in month_cells(raw, months):
        if c is None:
            base+='<td class="tm"></td><td class="tm"></td><td class="tm"></td>'
            continue
        st,dv,grp=c
        sc=SCS.get(st,'bg2'); stl=('✕ '+st) if st=='Rejected RFAI' else st
        base+=(f'<td class="tm"><span class="badge {sc}">{stl}</span></td>'
               f'<td class="tm tc mn" style="font-weight:600;color:var(--tx2)">'
               f'{dv if dv else "—"}</td>'
               f'<td class="tm tc"><span class="{GCS.get(grp,"gn")}">{GLS.get(grp,"N/A")}</span></td>')
    return base+'</tr>'

# ── FastAPI ────────────────────────────────────────────────────────────────
app=FastAPI()
app.add_middleware(CORSMiddleware,allow_origins=["*"],allow_methods=["*"],allow_headers=["*"])
from fastapi import Request
from fastapi.responses import JSONResponse as _JR

ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', '')  # password of the built-in "admin" user
_ON_SERVER     = bool(os.environ.get('RENDER'))
# Login is required whenever an admin password exists, and always on a server.
_AUTH_ON       = bool(ADMIN_PASSWORD) or _ON_SERVER
_PUBLIC_PATHS  = ('/login', '/api/login', '/api/logout', '/api/ping', '/favicon.ico',
                  # phone app ("Add to Home screen"): manifest, icons, service worker
                  '/manifest.webmanifest', '/sw.js', '/icon-192.png', '/icon-512.png',
                  '/icon-maskable-512.png', '/apple-touch-icon.png')
_ADMIN_PATHS   = ('/api/upload', '/api/reload', '/api/storage-test', '/api/users')
LOGIN_FILE     = SCRIPT_DIR / 'login.html'
_SESSION_COOKIE = 'rfai_session'
_SESSION_HOURS  = 12          # normal sign-in
_REMEMBER_DAYS  = 30          # "keep me signed in"

def _basic_pw(request):
    import base64
    h = request.headers.get('authorization', '')
    if not h.startswith('Basic '): return None
    try: return base64.b64decode(h[6:]).decode('utf-8', 'replace').split(':', 1)[1]
    except Exception: return None

def _pw_ok(given, expected):
    import secrets
    return given is not None and secrets.compare_digest(given.encode(), expected.encode())

# ── Users ─────────────────────────────────────────────────────────────────────
# Stored as rfai/_users.json in the bucket (or _rfai_cache/_users.json on a PC)
# with salted PBKDF2 password hashes. "admin" is built in: its password is
# ADMIN_PASSWORD, so it always works even if the users file is lost.
_users = {}
_users_state = {'loaded': False}
_users_lock = threading.Lock()
_USERS_LOCAL = _cache_dir / '_users.json'
_USER_RE = re.compile(r'^[a-z0-9._@-]{2,40}$')

def _hash_pw(pw, salt=None):
    import hashlib, secrets
    salt = salt or secrets.token_hex(16)
    return salt, hashlib.pbkdf2_hmac('sha256', pw.encode(), bytes.fromhex(salt), 200_000).hex()

def _load_users():
    data = None
    try:
        if S3_BUCKET:
            try:
                data = json.loads(_s3().get_object(Bucket=S3_BUCKET, Key=S3_PREFIX + '_users.json')['Body'].read())
            except Exception as e:
                if 'NoSuchKey' not in str(e) and '404' not in str(e): raise
        elif _USERS_LOCAL.exists():
            data = json.loads(_USERS_LOCAL.read_text())
    except Exception as e:
        log.error(f"  Could not read users: {e}")
        return False
    with _users_lock:
        _users.clear(); _users.update(data or {})
    _users_state['loaded'] = True
    return True

def _save_users():
    body = json.dumps(_users, indent=1)
    if S3_BUCKET:
        _s3().put_object(Bucket=S3_BUCKET, Key=S3_PREFIX + '_users.json',
                         Body=body.encode(), ContentType='application/json')
    else:
        _USERS_LOCAL.parent.mkdir(parents=True, exist_ok=True)
        _USERS_LOCAL.write_text(body)

def _check_user(name, pw):
    """-> role or None"""
    import secrets
    if name == 'admin':
        return 'admin' if ADMIN_PASSWORD and _pw_ok(pw, ADMIN_PASSWORD) else None
    u = _users.get(name)
    if not u: _hash_pw(pw or 'x'); return None          # same timing as a real check
    _, h = _hash_pw(pw, u['s'])
    return u.get('r', 'viewer') if secrets.compare_digest(h, u['h']) else None

def _user_fp(name):
    """Changes when the user's password changes or the user is deleted -> old sessions end."""
    import hashlib
    if name == 'admin': src = ADMIN_PASSWORD
    else:
        u = _users.get(name)
        if not u: return None
        src = u['h']
    return hashlib.sha256(('fp' + src).encode()).hexdigest()[:12]

# ── Sessions (signed cookie) ─────────────────────────────────────────────────
def _session_secret():
    import hashlib
    s = os.environ.get('SESSION_SECRET') or (ADMIN_PASSWORD + os.environ.get('S3_SECRET_KEY', '') + 'rfai-session')
    return hashlib.sha256(s.encode()).digest()

def _make_token(name, hours):
    import base64, hmac, hashlib
    body = base64.urlsafe_b64encode(json.dumps({'u': name, 'e': int(time.time() + hours * 3600),
                                                'f': _user_fp(name)}).encode()).decode().rstrip('=')
    sig = hmac.new(_session_secret(), body.encode(), hashlib.sha256).hexdigest()
    return body + '.' + sig

def _session_user(request):
    """-> (name, role) from a valid session cookie, else None"""
    import base64, hmac, hashlib
    tok = request.cookies.get(_SESSION_COOKIE, '')
    if '.' not in tok: return None
    body, sig = tok.rsplit('.', 1)
    if not hmac.compare_digest(sig, hmac.new(_session_secret(), body.encode(), hashlib.sha256).hexdigest()):
        return None
    try: d = json.loads(base64.urlsafe_b64decode(body + '=' * (-len(body) % 4)))
    except Exception: return None
    name = d.get('u')
    if d.get('e', 0) < time.time() or not name or d.get('f') != _user_fp(name): return None
    if name == 'admin': return (name, 'admin')
    return (name, _users[name].get('r', 'viewer'))

# brute-force guard: max 8 failed sign-ins per username + address in 10 minutes
# (per username, so one person's typos never lock out a whole office behind one IP)
_login_fails = {}
def _client_ip(request):
    return (request.headers.get('x-forwarded-for', '').split(',')[0].strip()
            or (request.client.host if request.client else '?'))

# Kill-switch for the public cloud copy (Render): set CLOUD_DISABLED=1 there to show
# only a "moved" notice and serve no data.
_CLOUD_OFF = _ON_SERVER and os.environ.get('CLOUD_DISABLED', '').strip() in ('1', 'true', 'yes')
_MOVED_PAGE = """<!doctype html><html><head><meta charset="utf-8"><meta name="viewport"
content="width=device-width,initial-scale=1"><title>RFAI Tracker</title><style>
body{font-family:system-ui,sans-serif;background:#f5f7fb;color:#1a1d23;display:flex;
align-items:center;justify-content:center;min-height:100vh;margin:0;padding:16px}
.c{background:#fff;border-radius:16px;padding:32px;max-width:460px;box-shadow:0 4px 24px rgba(0,0,0,.08)}
h1{font-size:20px;margin:0 0 12px}p{color:#4b5563;line-height:1.5;margin:0}</style></head>
<body><div class="c"><h1>RFAI Tracker has moved</h1><p>This tool now runs on the internal
office network. Please ask the tool owner for the internal link.</p></div></body></html>"""

@app.middleware("http")
async def _auth(request:Request,call_next):
    from fastapi.responses import HTMLResponse, RedirectResponse
    path = request.url.path
    if _CLOUD_OFF and path != '/api/ping':
        return HTMLResponse(_MOVED_PAGE, status_code=410 if path.startswith('/api/') else 200)
    request.state.user = ('admin', 'admin') if not _AUTH_ON else None
    if not _AUTH_ON or path in _PUBLIC_PATHS:
        return await call_next(request)
    user = _session_user(request)
    if user is None and ADMIN_PASSWORD and _pw_ok(_basic_pw(request), ADMIN_PASSWORD):
        user = ('admin', 'admin')                    # scripts: HTTP basic auth with ADMIN_PASSWORD
    if user is None:
        if path.startswith('/api/'):
            return JSONResponse({"error": "Please sign in", "login": True}, status_code=401)
        return RedirectResponse('/login', status_code=303)
    if any(path == a or path.startswith(a + '/') for a in _ADMIN_PATHS) and user[1] != 'admin':
        return JSONResponse({"error": "Only an admin can do this. Ask the tool owner."}, status_code=403)
    request.state.user = user
    return await call_next(request)

@app.get("/login")
def login_page():
    if LOGIN_FILE.exists(): return FileResponse(str(LOGIN_FILE), headers={'Cache-Control': 'no-store'})
    return JSONResponse({"error": "login.html not found"}, status_code=404)

@app.post("/api/login")
async def api_login(request:Request):
    from fastapi.responses import JSONResponse as J
    if _CLOUD_OFF: return J({"error": "off"}, status_code=410)
    if not ADMIN_PASSWORD and _ON_SERVER:
        return J({"error": "Sign-in is not set up yet: add ADMIN_PASSWORD in Render → Environment."}, status_code=503)
    try: d = await request.json()
    except Exception: d = {}
    name = str(d.get('username', '')).strip().lower(); pw = str(d.get('password', ''))
    key = _client_ip(request) + '|' + name; now = time.time()
    fails = [t for t in _login_fails.get(key, []) if now - t < 600]
    if len(fails) >= 8:
        return J({"error": "Too many wrong attempts. Try again in 10 minutes."}, status_code=429)
    if name != 'admin' and not _users_state['loaded']: _load_users()
    role = _check_user(name, pw) if name and pw else None
    if not role:
        fails.append(now); _login_fails[key] = fails
        if len(_login_fails) > 5000: _login_fails.clear()
        return J({"error": "Wrong username or password."}, status_code=401)
    _login_fails.pop(key, None)
    remember = bool(d.get('remember'))
    hours = _REMEMBER_DAYS * 24 if remember else _SESSION_HOURS
    r = J({"ok": True, "user": name, "role": role})
    secure = request.url.scheme == 'https' or request.headers.get('x-forwarded-proto') == 'https'
    r.set_cookie(_SESSION_COOKIE, _make_token(name, hours), httponly=True, samesite='lax', secure=secure,
                 max_age=hours * 3600 if remember else None, path='/')
    log.info(f"  Sign-in: {name} ({role})")
    return r

@app.post("/api/logout")
def api_logout():
    r = JSONResponse({"ok": True}); r.delete_cookie(_SESSION_COOKIE, path='/'); return r

@app.get("/api/me")
def api_me(request:Request):
    u = request.state.user or ('', '')
    return {"user": u[0], "role": u[1], "auth": _AUTH_ON}

@app.post("/api/me/password")
async def api_my_password(request:Request):
    name = request.state.user[0]
    if name == 'admin':
        return JSONResponse({"error": "The admin password is ADMIN_PASSWORD in the server settings."}, status_code=400)
    d = await request.json()
    if not _check_user(name, str(d.get('old', ''))):
        return JSONResponse({"error": "Current password is wrong."}, status_code=400)
    new = str(d.get('new', ''))
    if len(new) < 8: return JSONResponse({"error": "New password must be at least 8 characters."}, status_code=400)
    with _users_lock:
        u = _users[name]; u['s'], u['h'] = _hash_pw(new)
        _save_users()
    return {"ok": True}

@app.get("/api/users")
def api_users():
    with _users_lock:
        rows = [{"username": k, "role": v.get('r', 'viewer'), "created": v.get('c', ''), "by": v.get('by', '')}
                for k, v in sorted(_users.items())]
    return {"users": [{"username": "admin", "role": "admin", "created": "built in", "by": ""}] + rows}

@app.post("/api/users")
async def api_user_save(request:Request):
    d = await request.json()
    name = str(d.get('username', '')).strip().lower(); pw = str(d.get('password', ''))
    role = 'admin' if d.get('role') == 'admin' else 'viewer'
    if name == 'admin': return JSONResponse({"error": "\"admin\" is the built-in user."}, status_code=400)
    if not _USER_RE.match(name):
        return JSONResponse({"error": "Username: 2-40 characters, letters, numbers, . _ @ - only."}, status_code=400)
    if not _users_state['loaded'] and not _load_users():
        return JSONResponse({"error": "Could not read the user list from storage. Try again."}, status_code=503)
    if name not in _users and len(pw) < 8:
        return JSONResponse({"error": "Password must be at least 8 characters."}, status_code=400)
    if pw and len(pw) < 8:
        return JSONResponse({"error": "Password must be at least 8 characters."}, status_code=400)
    with _users_lock:
        u = _users.setdefault(name, {'c': datetime.now().strftime('%d-%b-%Y'), 'by': request.state.user[0]})
        u['r'] = role
        if pw: u['s'], u['h'] = _hash_pw(pw)
        try: _save_users()
        except Exception as e:
            _load_users(); return JSONResponse({"error": f"Could not save: {e}"}, status_code=500)
    return {"ok": True}

@app.delete("/api/users/{name}")
def api_user_delete(name:str):
    name = name.strip().lower()
    if not _users_state['loaded'] and not _load_users():
        return JSONResponse({"error": "Could not read the user list from storage. Try again."}, status_code=503)
    with _users_lock:
        if name not in _users: return JSONResponse({"error": "No such user."}, status_code=404)
        _users.pop(name)
        try: _save_users()
        except Exception as e:
            _load_users(); return JSONResponse({"error": f"Could not save: {e}"}, status_code=500)
    return {"ok": True}

# ── Keep the free Supabase project awake ──────────────────────────────────────
# Supabase pauses free projects after ~a week without *database* activity; this app
# only uses its file storage. With SUPABASE_DB_URL set (Supabase → Connect →
# Session pooler URI), the server writes one tiny row twice a day so it never pauses.
SUPABASE_DB_URL=os.environ.get('SUPABASE_DB_URL','').strip().strip('"').strip("'")
_keepalive={'last':None,'error':None}
def _supabase_touch():
    import psycopg2
    with psycopg2.connect(SUPABASE_DB_URL, connect_timeout=15) as con, con.cursor() as cur:
        cur.execute("CREATE TABLE IF NOT EXISTS rfai_keepalive (id int PRIMARY KEY, at timestamptz, note text)")
        cur.execute("INSERT INTO rfai_keepalive VALUES (1, now(), %s) ON CONFLICT (id) DO UPDATE SET at=now(), note=EXCLUDED.note",
                    (f"rfai-tracker {_data_meta.get('srs','')} SRs",))
    con.close()
def _keepalive_loop():
    while True:
        try:
            _supabase_touch(); _keepalive.update(last=datetime.now().strftime('%d-%b-%Y %H:%M'),error=None)
            log.info("  Supabase keep-alive OK")
        except Exception as e:
            _keepalive['error']=str(e)[:300]; log.error(f"  Supabase keep-alive failed: {e}")
        time.sleep(12*3600)

@app.on_event("startup")
def _startup_load():
    if _CLOUD_OFF:
        log.info("  Cloud copy switched off (CLOUD_DISABLED=1)"); return
    if _AUTH_ON: threading.Thread(target=_load_users, daemon=True).start()
    threading.Thread(target=load_stored_data, kwargs={'retries': 3}, daemon=True).start()
    if SUPABASE_DB_URL: threading.Thread(target=_keepalive_loop, daemon=True).start()

@app.post("/api/reload")
def reload_data():
    if _job.running: return JSONResponse({"error":"Busy"},status_code=409)
    # a manual Reload gives a previously crashed conversion one more try
    threading.Thread(target=load_stored_data, kwargs={'force': True}, daemon=True).start()
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
    # mark busy NOW: the page polls /api/progress right away, and on a slow server
    # the background task may not have started yet (it would see the old 'done')
    _job.reset(); _job.running=True; _job.message='Receiving file(s)…'
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

@app.get("/api/storage-test")
def storage_test():
    """Try the storage connection and explain what is wrong (no secrets shown)."""
    ak = os.environ.get('S3_ACCESS_KEY', ''); sk = os.environ.get('S3_SECRET_KEY', '')
    ep = os.environ.get('S3_ENDPOINT', '')
    info = {"endpoint": ep, "bucket": S3_BUCKET, "region": _s3_region(), "prefix": S3_PREFIX,
            "access_key_length": len(ak), "secret_key_length": len(sk),
            "access_key_starts": ak[:4] + '…' if ak else ''}
    hints = []
    if not ep: hints.append('S3_ENDPOINT is empty')
    if not S3_BUCKET: hints.append('S3_BUCKET is empty')
    if not ak or not sk: hints.append('S3_ACCESS_KEY / S3_SECRET_KEY missing')
    if ak and sk and len(ak) > len(sk): hints.append('Access key is longer than secret key — they may be swapped')
    if 'supabase' in ep and not ep.endswith('/storage/v1/s3'):
        hints.append("Supabase endpoint should end with /storage/v1/s3")
    if 'supabase' in ep and not os.environ.get('S3_REGION'):
        hints.append('S3_REGION is empty — Supabase needs it (e.g. ap-south-1)')
    try:
        _last_s3_body.clear()
        s3 = _s3()
        names = [b['Name'] for b in s3.list_buckets().get('Buckets', [])]
        info['buckets_visible_with_these_keys'] = names
        if S3_BUCKET and names and S3_BUCKET not in names:
            hints.append(f"Bucket '{S3_BUCKET}' not found — these keys see: {names}")
    except Exception as e:
        info['list_buckets_error'] = _s3_err(e)
    try:
        _last_s3_body.clear()
        r = _s3().list_objects_v2(Bucket=S3_BUCKET, Prefix=S3_PREFIX)
        info['files_in_storage'] = [o['Key'] for o in r.get('Contents', [])]
        info['ok'] = True
    except Exception as e:
        info['ok'] = False; info['error'] = _s3_err(e)
        err = info['error']
        if 'SignatureDoesNotMatch' in err or '403' in err:
            hints.append('Keys or region are wrong: re-copy Access key ID / Secret access key, check S3_REGION')
        if 'NoSuchBucket' in err or '404' in err:
            hints.append('Bucket name or endpoint is wrong')
    info['hints'] = hints
    return info

@app.get("/api/config")
def config():
    """Non-secret setup status, shown on the start screen."""
    env_names = sorted(k for k in os.environ if 'PASS' in k.upper())
    return {"upload_enabled": bool(ADMIN_PASSWORD) or not _ON_SERVER,
            "storage_configured": bool(S3_BUCKET) or SQL_ENABLED,
            "sql_server": ({"server": _sql_server['name'], "database": SQL_DATABASE}
                           if SQL_ENABLED else None),
            "password_like_vars": env_names,
            "data_loaded": _df_cache is not None,
            "srs_loaded": len(_df_cache) if _df_cache is not None else 0,
            "storage": _storage_status,
            "supabase_keepalive": ({"on":True,**_keepalive} if SUPABASE_DB_URL else {"on":False})}

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
            "domains":_ul('domain'),"oems":_ul('oem'),"months":mons,
            "data_info":dict(_data_meta, srs=total_unique)}

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

# ── Circle velocity: milestone events per circle per period ──────────────────
_VEL_COL={'rfai':'rfaiNotice','final':'finalRfai','pri_open':'priFirst','pri_close':'priLast',
          'onair':'onAir','rej':'rfaiRej'}
def _dcol(df,col):
    if col not in df.columns: return pd.Series(pd.NaT,index=df.index)
    return pd.to_datetime(df[col].astype(str),format='%d-%b-%Y',errors='coerce')
def _periods(start,end,gran,default_weeks=8):
    today=pd.Timestamp(date.today())
    try: e=pd.Timestamp(end).normalize() if end else today
    except Exception: e=today
    try: b=pd.Timestamp(start).normalize() if start else None
    except Exception: b=None
    gran=gran if gran in ('day','week','month') else 'week'
    if gran=='week':
        # 7-day blocks ending on the end date (newest block = last 7 days)
        if b is None: b=e-pd.Timedelta(days=7*default_weeks-1)
        n=max(1,min(156,-(-((e-b).days+1)//7)))
        starts=[e-pd.Timedelta(days=7*(n-i)-1) for i in range(n)]
        ends=[x+pd.Timedelta(days=6) for x in starts]; lab=[x.strftime('%d %b') for x in starts]
    elif gran=='day':
        if b is None: b=e-pd.Timedelta(days=29)
        if (e-b).days>366: b=e-pd.Timedelta(days=366)
        starts=list(pd.date_range(b,e,freq='D')); ends=starts; lab=[x.strftime('%d %b') for x in starts]
    else:
        if b is None: b=(e-pd.DateOffset(months=5)).replace(day=1)
        starts=list(pd.date_range(b.replace(day=1),e,freq='MS'))[-60:]
        ends=[min(x+pd.offsets.MonthEnd(0),e) for x in starts]; lab=[x.strftime("%b'%y") for x in starts]
    return starts,ends,lab,gran

@app.get("/api/velocity")
def velocity(metric:str=Query(default='pri_open'), start:str=Query(default=''), end:str=Query(default=''),
             gran:str=Query(default='week')):
    """Per circle and period: how many SRs reached a milestone (or, for pri_stock,
    how many had a PRI open at the end of the period)."""
    df=_df_full_cache
    if df is None or not len(df): return {"labels":[],"rows":[]}
    starts,ends,lab,gran=_periods(start,end,gran)
    circ=df['circle'].astype(str).to_numpy()
    today=pd.Timestamp(date.today())
    if metric=='pri_stock':
        pf=_dcol(df,'priFirst'); pl=_dcol(df,'priLast'); oa=_dcol(df,'onAir')
        def at(t): return ((pf<=t)&~(pl<=t)&~(oa<=t)).to_numpy()
        cols=[at(min(x,today)) for x in ends]
        last7=at(today); prev7=at(today-pd.Timedelta(days=7))
    else:
        d=_dcol(df,_VEL_COL.get(metric,'priFirst'))
        cols=[((d>=a)&(d<=b)).to_numpy() for a,b in zip(starts,ends)]
        last7=((d>today-pd.Timedelta(days=7))&(d<=today)).to_numpy()
        prev7=((d>today-pd.Timedelta(days=14))&(d<=today-pd.Timedelta(days=7))).to_numpy()
    names=sorted(c for c in set(circ) if c and c!='nan')
    rows=[]
    for c in names:
        m=circ==c
        rows.append({"circle":c,"counts":[int((m&x).sum()) for x in cols],"last7":int((m&last7).sum()),"prev7":int((m&prev7).sum())})
    rows.sort(key=lambda r:-(sum(r['counts']) if metric!='pri_stock' else r['counts'][-1]))
    tot={"counts":[int(x.sum()) for x in cols],"last7":int(last7.sum()),"prev7":int(prev7.sum())}
    return {"labels":lab,"starts":[x.strftime('%Y-%m-%d') for x in starts],"ends":[min(x,today).strftime('%Y-%m-%d') for x in ends],
            "gran":gran,"metric":metric,"rows":rows,"total":tot}

@app.get("/api/velocity/detail")
def velocity_detail(metric:str=Query(default='pri_open'), circle:str=Query(default=''),
                    start:str=Query(default=''), end:str=Query(default=''),
                    sel_from:str=Query(default=''), sel_to:str=Query(default='')):
    """Daily series + running total over start..end (chart), and key facts for the
    selected period sel_from..sel_to (defaults to the whole range)."""
    df=_df_full_cache
    if df is None or not len(df): return {}
    if circle: df=df[df['circle'].astype(str)==circle]
    today=pd.Timestamp(date.today())
    try: e=pd.Timestamp(end).normalize() if end else today
    except Exception: e=today
    try: b=pd.Timestamp(start).normalize() if start else e-pd.Timedelta(days=55)
    except Exception: b=e-pd.Timedelta(days=55)
    if (e-b).days>730: b=e-pd.Timedelta(days=730)
    days=pd.date_range(b,e,freq='D')
    st=df['srStatus'].astype(str)
    out={"circle":circle or 'All circles',"metric":metric,"start":b.strftime('%Y-%m-%d'),"end":e.strftime('%Y-%m-%d'),
         "labels":[x.strftime('%d %b') for x in days]}
    pf=_dcol(df,'priFirst'); pl=_dcol(df,'priLast'); oa=_dcol(df,'onAir')
    try: sb=pd.Timestamp(sel_from).normalize() if sel_from else b
    except Exception: sb=b
    try: se=pd.Timestamp(sel_to).normalize() if sel_to else e
    except Exception: se=e
    insel=[(x>=sb)&(x<=se) for x in days]
    if metric=='pri_stock':
        vals=[int(((pf<=t)&~(pl<=t)&~(oa<=t)).sum()) for t in days]
        out.update(daily=vals,cumulative=None)
        sel=((pf<=se)&~(pl<=se)&~(oa<=se)); out["total"]=int(sel.sum())
    else:
        d=_dcol(df,_VEL_COL.get(metric,'priFirst'))
        cnt=d[(d>=b)&(d<=e)].dt.normalize().value_counts()
        vals=[int(cnt.get(x,0)) for x in days]
        out.update(daily=vals,cumulative=[int(v) for v in np.cumsum(vals)])
        sel=(d>=sb)&(d<=se); out["total"]=int(sel.sum())
    sub=df[sel.to_numpy()]
    sv=[v for v,i in zip(vals,insel) if i]; sd=[x for x,i in zip(days,insel) if i]
    out["avg_per_day"]=round(out['total']/max(1,len(sv)),1) if metric!='pri_stock' else None
    if sv:
        k=int(np.argmax(sv)); out["peak"]={"day":sd[k].strftime('%d %b %Y'),"value":sv[k]}
    out["status_mix"]={k:int(v) for k,v in sub['srStatus'].astype(str).value_counts().items() if k and k!='nan'}
    out["top_projects"]=[{"name":str(k),"n":int(v)} for k,v in sub['bPrj'].astype(str).value_counts().head(6).items() if k and k!='nan'] if 'bPrj' in sub else []
    if metric in ('pri_open','pri_close','pri_stock'):
        pdays=pd.to_numeric(sub.get('priDays'),errors='coerce') if 'priDays' in sub else pd.Series(dtype=float)
        closed=pdays[pdays>0]
        out["pri_avg_days"]=round(float(closed.mean()),1) if len(closed) else None
        openmask=(~(pl<=today)&(pf<=today)&~(oa<=today))[sel]
        age=(today-pf[sel][openmask]).dt.days
        out["pri_open_now"]=int(openmask.sum()); out["pri_oldest_open_days"]=int(age.max()) if len(age) else None
    if circle=='' :
        g=sub['circle'].astype(str).value_counts()
        out["by_circle"]=[{"name":str(k),"n":int(v)} for k,v in g.head(12).items() if k and k!='nan']
    return out

@app.get("/api/momentum")
def momentum(circle:str=Query(default=''), start:str=Query(default=''), end:str=Query(default=''),
             gran:str=Query(default='week')):
    """OnAir achieved, RFAI received and SRs pending, per day / week / month, for any date range."""
    df=_df_full_cache
    if df is None or not len(df): return {"labels":[]}
    if circle: df=df[df['circle'].astype(str)==circle]
    today=pd.Timestamp(date.today())
    try: e=pd.Timestamp(end) if end else today
    except Exception: e=today
    try: b=pd.Timestamp(start) if start else e-pd.Timedelta(weeks=12)
    except Exception: b=e-pd.Timedelta(weeks=12)
    if b>e: b,e=e,b
    gran=gran if gran in ('day','week','month') else 'week'
    if gran=='day':
        b=b.normalize(); e=e.normalize()
        if (e-b).days>366: b=e-pd.Timedelta(days=366)
        starts=list(pd.date_range(b,e,freq='D')); lab=lambda x:x.strftime('%d %b')
        nxt=lambda x:x+pd.Timedelta(days=1)
    elif gran=='week':
        b=(b-pd.Timedelta(days=b.weekday())).normalize()
        if (e-b).days>7*156: b=(e-pd.Timedelta(weeks=156)); b=(b-pd.Timedelta(days=b.weekday())).normalize()
        starts=list(pd.date_range(b,e,freq='7D')); lab=lambda x:x.strftime('%d %b')+("'"+x.strftime('%y') if (e-b).days>330 else '')
        nxt=lambda x:x+pd.Timedelta(days=7)
    else:
        b=b.replace(day=1).normalize()
        starts=list(pd.date_range(b,e,freq='MS')); lab=lambda x:x.strftime("%b'%y")
        nxt=lambda x:(x+pd.offsets.MonthBegin(1))
    starts=starts[-200:]
    if not starts: return {"labels":[]}
    def dt(col):
        if col not in df.columns: return pd.Series(pd.NaT,index=df.index)
        return pd.to_datetime(df[col].astype(str),format='%d-%b-%Y',errors='coerce')
    d_rfai=dt('rfaiNotice'); d_oa=dt('onAir'); d_rej=dt('rfaiRej')
    st=df['srStatus'].astype(str).to_numpy(); isrej=(st=='Rejected RFAI')
    edges=[x for x in starts]+[nxt(starts[-1])]
    def counts(d):
        v=d.dropna(); v=v[(v>=edges[0])&(v<edges[-1])]
        idx=np.searchsorted(np.array(edges[1:],dtype='datetime64[ns]'),v.to_numpy(dtype='datetime64[ns]'),side='right')
        return [int(x) for x in np.bincount(idx,minlength=len(starts))[:len(starts)]]
    pend=[]
    for x in starts:
        en=min(nxt(x)-pd.Timedelta(days=1),today)
        sc=(d_rfai<=en).to_numpy(); oa=(d_oa<=en).to_numpy(); rj=isrej&(d_rej<=en).to_numpy()
        pend.append(int((sc&~oa&~rj).sum()))
    return {"labels":[lab(x) for x in starts],"onair":counts(d_oa),"rfai":counts(d_rfai),"pending":pend,
            "gran":gran,"start":edges[0].strftime('%Y-%m-%d'),"end":min(edges[-1]-pd.Timedelta(days=1),e).strftime('%Y-%m-%d')}

@app.get("/api/home")
def home_stats(circle:str=Query(default='')):
    """Milestone journey, recent movement, weekly momentum and circle scorecard
    for the Home page. ?circle=X scopes everything to one circle."""
    df=_df_full_cache
    if df is None or not len(df): return {"total":0}
    all_circles=sorted(c for c in df['circle'].astype(str).unique() if c and c!='nan')
    if circle: df=df[df['circle'].astype(str)==circle]
    if not len(df): return {"total":0,"circles_all":all_circles,"circle":circle}
    st=df['srStatus'].astype(str).to_numpy()
    tot=len(df)
    def cnt(x): return int((st==x).sum())
    def dates(col):
        if col not in df.columns: return pd.Series(pd.NaT,index=df.index)
        return pd.to_datetime(df[col].astype(str),format='%d-%b-%Y',errors='coerce')
    today=pd.Timestamp(date.today()); wk=today-pd.Timedelta(days=7)
    d_rfai=dates('rfaiNotice'); d_fin=dates('finalRfai'); d_pri=dates('priLast'); d_oa=dates('onAir')
    def last7(d): return int(((d>wk)&(d<=today)).sum())
    def nonblank(col): return (df[col].astype(str).str.len()>0).to_numpy() if col in df.columns else np.zeros(tot,bool)
    delay=pd.to_numeric(df['delay'],errors='coerce').fillna(-1).to_numpy()
    pend=np.isin(st,['OnAir Pending','Open PRI','RFAI Pending'])
    # weekly momentum: last 12 weeks (Mon-start), RFAI received vs OnAir achieved
    start=(today-pd.Timedelta(days=today.weekday()))-pd.Timedelta(weeks=11)
    weeks=[start+pd.Timedelta(weeks=i) for i in range(12)]
    def per_week(d):
        d=d[(d>=start)&(d<=today)]
        idx=((d-start).dt.days//7).astype(int)
        c=np.bincount(idx.to_numpy(),minlength=12)[:12] if len(idx) else np.zeros(12,int)
        return [int(x) for x in c]
    weekly={"labels":[w.strftime('%d %b') for w in weeks],"rfai":per_week(d_rfai),"onair":per_week(d_oa)}
    # 12-week trends for the KPI sparklines: state as at the end of each week
    d_rej=dates('rfaiRej'); d_pf=dates('priFirst')
    isrej=(st=='Rejected RFAI')
    dl_oa=pd.to_numeric(df['delay'],errors='coerce').to_numpy()
    tr={k:[] for k in ('scope','accepted','onair_pct','pending','pending_30','pri_open','onair_week','avg_delay')}
    for i,w in enumerate(weeks):
        e=min(w+pd.Timedelta(days=6),today)
        sc=(d_rfai<=e).to_numpy(); oa=(d_oa<=e).to_numpy()
        rj=(isrej&(d_rej<=e).to_numpy())
        pe=sc&~oa&~rj
        p30=pe&(d_fin<=e-pd.Timedelta(days=30)).to_numpy()
        po=(d_pf<=e).to_numpy()&~((d_pri<=e).to_numpy())&~oa
        inwk=((d_oa>=w)&(d_oa<=e)).to_numpy()
        n_sc=int(sc.sum())
        tr['scope'].append(n_sc); tr['accepted'].append(int((sc&~rj).sum()))
        tr['onair_pct'].append(round(float(oa.sum())/n_sc*100,1) if n_sc else 0)
        tr['pending'].append(int(pe.sum())); tr['pending_30'].append(int(p30.sum()))
        tr['pri_open'].append(int(po.sum())); tr['onair_week'].append(int(inwk.sum()))
        v=dl_oa[inwk&(dl_oa>=0)]
        tr['avg_delay'].append(int(round(v.mean())) if len(v) else None)
    weekly['pending']=tr['pending']
    rd=d_rfai.dropna()
    date_range=[rd.min().strftime('%d %b %Y'),rd.max().strftime('%d %b %Y')] if len(rd) else None
    # circle scorecard
    g=pd.DataFrame({'c':df['circle'].astype(str).to_numpy(),'on':st=='OnAir','pe':pend,
                    'p30':pend&(delay>30),'pri':st=='Open PRI','rj':st=='Rejected RFAI',
                    'age':np.where(pend&(delay>=0),delay,np.nan)})
    g=g[g['c'].str.len()>0]
    sc=g.groupby('c').agg(total=('on','size'),onair=('on','sum'),pending=('pe','sum'),p30=('p30','sum'),
                          pri=('pri','sum'),rej=('rj','sum'),age=('age','mean')).reset_index()
    circles=[{"circle":r.c,"total":int(r.total),"onair":int(r.onair),"pending":int(r.pending),
              "pending_30":int(r.p30),"pri_open":int(r.pri),"rejected":int(r.rej),
              "avg_age":None if pd.isna(r.age) else int(round(r.age))} for r in sc.itertuples()]
    circles.sort(key=lambda x:-x['total'])
    return {
        "total":tot,"circle":circle,"circles_all":all_circles,
        "statuses":{k:cnt(k) for k in ('OnAir','OnAir Pending','Open PRI','Rejected RFAI','RFAI Pending')},
        "rfai_rejected_ever":int(nonblank('rfaiRej').sum()), "rfai_reoffered":int(nonblank('rfaiReoffer').sum()),
        "pri_ever":int(nonblank('priFirst').sum()), "pri_open":cnt('Open PRI'),
        "pending":int(pend.sum()), "pending_30":int((pend&(delay>30)).sum()),
        "pending_90":int((pend&(delay>90)).sum()),
        "last7":{"rfai":last7(d_rfai),"final":last7(d_fin),"pri_closed":last7(d_pri),"onair":last7(d_oa)},
        "weekly":weekly,"circles":circles,"trends":tr,"date_range":date_range,
        "as_of":date.today().strftime('%d-%b-%Y'),
    }

@app.get("/api/sr")
def sr_detail(sr:str=Query(default='')):
    """Everything known about one SR (for the detail panel)."""
    df=_df_full_cache
    if df is None or not sr: return JSONResponse({"error":"Not found"},status_code=404)
    hit=df[df['srNo'].astype(str)==sr]
    if not len(hit): return JSONResponse({"error":"Not found"},status_code=404)
    r=hit.iloc[0]
    out={}
    for c in OUT:
        if c not in hit.columns: continue
        v=r[c]
        if c in ('delay','priDays'):
            try: v=int(v)
            except Exception: v=-1
        else:
            v='' if v is None or str(v) in ('nan','None') else str(v)
        out[c]=v
    return out

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

@app.get("/api/explore")
def explore(status:str=Query(default=''), circle:str=Query(default=''), category:str=Query(default=''),
            oem:str=Query(default=''), toco:str=Query(default=''), month:str=Query(default=''),
            aging:str=Query(default=''), recent:str=Query(default=''), q:str=Query(default=''),
            ev:str=Query(default=''), dfrom:str=Query(default=''), dto:str=Query(default=''),
            sort:str=Query(default='delay'), dir:str=Query(default='desc'),
            page:int=Query(default=1), size:int=Query(default=50)):
    """SR Explorer: the SRs behind any number clicked on Home or the Dashboard,
    with summaries (status mix, circles, ageing) and a sortable, paged list."""
    df=_df_full_cache
    if df is None or not len(df): return {"total":0,"rows":[],"summary":{}}
    m=np.ones(len(df),bool)
    def col(c): return df[c].astype(str).to_numpy() if c in df.columns else np.full(len(df),'')
    st=col('srStatus')
    PEND=['OnAir Pending','Open PRI','RFAI Pending']
    if status:
        want=[]
        for x in status.split(','):
            want+=PEND if x=='pending' else [x]
        m&=np.isin(st,want)
    for key,val in (('circle',circle),('category',category),('oem',oem),('toco',toco),('exitMK',month)):
        if val: m&=np.isin(col(key),[v for v in val.split(',')])
    delay=pd.to_numeric(df['delay'],errors='coerce').fillna(-1).to_numpy()
    if aging:
        lo,hi={'0-30':(0,30),'30-60':(31,60),'60-90':(61,90),'>90':(91,10**6),'>30':(31,10**6)}.get(aging,(0,10**6))
        m&=(delay>=lo)&(delay<=hi)
    if recent:
        cname={'onair7':'onAir','final7':'finalRfai','pri7':'priLast','rfai7':'rfaiNotice'}.get(recent)
        if cname and cname in df.columns:
            d=pd.to_datetime(df[cname].astype(str),format='%d-%b-%Y',errors='coerce')
            today=pd.Timestamp(date.today())
            m&=((d>today-pd.Timedelta(days=7))&(d<=today)).to_numpy()
    if ev and (ev in _VEL_COL or ev=='pri_stock'):
        try:
            b_=pd.Timestamp(dfrom) if dfrom else pd.Timestamp('1900-01-01'); e_=pd.Timestamp(dto) if dto else pd.Timestamp(date.today())
            if ev=='pri_stock':
                pf=_dcol(df,'priFirst'); pl=_dcol(df,'priLast'); oa=_dcol(df,'onAir')
                m&=((pf<=e_)&~(pl<=e_)&~(oa<=e_)).to_numpy()
            else:
                d=_dcol(df,_VEL_COL[ev]); m&=((d>=b_)&(d<=e_)).to_numpy()
        except Exception: pass
    if q:
        ql=q.strip().lower()
        m&=(pd.Series(col('srNo')).str.lower().str.contains(ql,regex=False).to_numpy()
            |pd.Series(col('siteId')).str.lower().str.contains(ql,regex=False).to_numpy())
    sub=df[m]; total=int(m.sum())
    sst=st[m]; sd=delay[m]
    # summaries
    stc={k:int(v) for k,v in pd.Series(sst).value_counts().items() if k and k!='nan'}
    circ=[]
    if total:
        g=pd.DataFrame({'c':col('circle')[m],'on':sst=='OnAir'}).groupby('c')
        cs=g.size().sort_values(ascending=False)
        on=g['on'].sum()
        circ=[{'circle':str(k),'n':int(v),'onair':int(on.get(k,0))} for k,v in cs.head(12).items() if k and k!='nan']
    pend_mask=np.isin(sst,PEND)
    ages={}
    for lab,(lo,hi) in (('0-30',(0,30)),('30-60',(31,60)),('60-90',(61,90)),('>90',(91,10**6))):
        ages[lab]=int(((sd>=lo)&(sd<=hi)).sum())
    valid=sd[sd>=0]
    summ={'statuses':stc,'circles':circ,'aging':ages,
          'avg_days':int(round(valid.mean())) if len(valid) else None,
          'max_days':int(valid.max()) if len(valid) else None,
          'pending':int(pend_mask.sum()),'onair':int((sst=='OnAir').sum())}
    # sort + page
    COLS=['srNo','siteId','circle','srStatus','delay','bPrj','category','toco','oem','tsp',
          'rfaiNotice','finalRfai','priLast','onAir','exitMK']
    sort=sort if sort in COLS else 'delay'
    asc=dir=='asc'
    if total:
        if sort in ('rfaiNotice','finalRfai','priLast','onAir'):
            key=pd.to_datetime(sub[sort].astype(str),format='%d-%b-%Y',errors='coerce')
        elif sort=='delay':
            key=pd.Series(delay[m],index=sub.index).where(lambda x:x>=0)
        else:
            key=sub[sort].astype(str).str.lower() if sort in sub.columns else pd.Series('',index=sub.index)
        order=key.sort_values(ascending=asc,na_position='last').index
        size=max(10,min(int(size),500)); pages=max(1,-(-total//size)); page=max(1,min(int(page),pages))
        pg=sub.loc[order[(page-1)*size:page*size]]
    else:
        size=max(10,min(int(size),500)); pages=1; page=1; pg=sub
    rows=[]
    for _,r in pg.iterrows():
        d={c:('' if c not in pg.columns or str(r[c]) in ('nan','None') else str(r[c])) for c in COLS if c!='delay'}
        dv=pd.to_numeric(r.get('delay'),errors='coerce')
        d['delay']=None if dv!=dv or dv==-1 else int(dv)
        rows.append(d)
    return {"total":total,"page":page,"pages":pages,"size":size,"rows":rows,"summary":summ}

# ── Excel export ──────────────────────────────────────────────────────────────
_XL_ST={'OnAir':('#E6F7EE','#0B7A43'),'OnAir Pending':('#FFF3D6','#8A5300'),'Open PRI':('#E6EEFF','#1D4ED8'),
        'Rejected RFAI':('#FDE8E8','#B91C1C'),'RFAI Pending':('#EEF0F4','#475569')}
_XL_GR={'0-30':('#DCF5E8','#0B7A43'),'30-60':('#FFF4CC','#8A6100'),'60-90':('#FFE4CC','#B34700'),
        '>90':('#FDDCDC','#B91C1C'),'N/A':('#F1F2F5','#6B7280')}
_XL_STATUSES=['OnAir','OnAir Pending','Open PRI','Rejected RFAI','RFAI Pending']
_XL_MAXROWS=1_000_000          # Excel holds 1,048,576 rows per sheet; bigger lists continue on the next sheet

def _xl_filter(df, p):
    """Same filters as the Journey Table (/api/filter) plus the SR Explorer ones."""
    PEND=['OnAir Pending','Open PRI','RFAI Pending']
    for key,colname in (('circle','circle'),('category','category'),('project','bPrj'),('toco','toco'),
                        ('domain','domain'),('oem','oem')):
        if p.get(key): df=df[df[colname].isin(p[key].split(','))]
    if p.get('status'):
        want=[]
        for x in p['status'].split(','): want+=PEND if x=='pending' else [x]
        df=df[df['srStatus'].isin(want)]
    if p.get('month'):
        ms=set(p['month'].split(','))
        df=df[df['exitMK'].isin(ms)|df['journeyMK'].isin(ms)] if p.get('scope')!='explore' else df[df['exitMK'].isin(ms)]
    q=(p.get('search') or p.get('q') or '').strip().upper()
    if q:
        df=df[df['srNo'].astype(str).str.upper().str.contains(q,na=False,regex=False)|
              df['siteId'].astype(str).str.upper().str.contains(q,na=False,regex=False)]
    if p.get('aging'):
        lo,hi={'0-30':(0,30),'30-60':(31,60),'60-90':(61,90),'>90':(91,10**6),'>30':(31,10**6)}.get(p['aging'],(0,10**6))
        dv=pd.to_numeric(df['delay'],errors='coerce').fillna(-1)
        df=df[(dv>=lo)&(dv<=hi)]
    ev=p.get('ev')
    if ev and (ev in _VEL_COL or ev=='pri_stock'):
        b_=pd.Timestamp(p.get('dfrom') or '1900-01-01'); e_=pd.Timestamp(p.get('dto') or date.today())
        if ev=='pri_stock':
            pf=_dcol(df,'priFirst'); pl=_dcol(df,'priLast'); oa=_dcol(df,'onAir')
            df=df[(pf<=e_)&~(pl<=e_)&~(oa<=e_)]
        else:
            d=_dcol(df,_VEL_COL[ev]); df=df[(d>=b_)&(d<=e_)]
    if p.get('recent'):
        cname={'onair7':'onAir','final7':'finalRfai','pri7':'priLast','rfai7':'rfaiNotice'}.get(p['recent'])
        if cname and cname in df.columns:
            d=pd.to_datetime(df[cname].astype(str),format='%d-%b-%Y',errors='coerce')
            today=pd.Timestamp(date.today())
            df=df[(d>today-pd.Timedelta(days=7))&(d<=today)]
    return df

def build_export_xlsx(df, path, filters_text, title, include_journey=True):
    import xlsxwriter
    wb=xlsxwriter.Workbook(path,{'constant_memory':True,'strings_to_numbers':False,'default_date_format':'dd-mmm-yyyy'})
    wb.set_properties({'title':title,'subject':'RFAI Tracker export','company':'Airtel'})
    RED='#E40000'; INK='#1A1D24'; LINE='#E6D9DB'
    F=lambda **k: wb.add_format({'font_name':'Calibri','font_size':10,'valign':'vcenter',**k})
    f_banner=F(bold=True,font_size=20,font_color='#FFFFFF',bg_color=RED,indent=1)
    f_bsub=F(font_color='#FFE3E3',bg_color=RED,indent=1,italic=True)
    f_h2=F(bold=True,font_size=13,font_color=INK,bottom=2,bottom_color=RED)
    f_note=F(font_color='#6B7280',italic=True)
    f_hdr=F(bold=True,font_color='#FFFFFF',bg_color='#2B2F3A',border=1,border_color='#2B2F3A',align='center',text_wrap=True)
    f_band=F(bold=True,font_color='#FFFFFF',bg_color=RED,border=1,border_color='#FFFFFF',align='center')
    f_band2=F(bold=True,font_color=INK,bg_color='#F6E7E9',border=1,border_color=LINE,align='left',indent=1)
    f_txt=F(border=1,border_color=LINE)
    f_txtb=F(border=1,border_color=LINE,bold=True,font_color=INK)
    f_num=F(border=1,border_color=LINE,num_format='#,##0',align='center')
    f_pct=F(border=1,border_color=LINE,num_format='0%',align='center')
    f_date=F(border=1,border_color=LINE,num_format='dd-mmm-yyyy',align='center')
    f_tot=F(border=1,border_color=LINE,bold=True,bg_color='#FBEFF0',num_format='#,##0',align='center')
    f_totl=F(border=1,border_color=LINE,bold=True,bg_color='#FBEFF0')
    f_link=F(font_color='#1D4ED8',underline=1)
    f_st={k:F(border=1,border_color=LINE,bg_color=b,font_color=c,bold=True,align='center') for k,(b,c) in _XL_ST.items()}
    f_gr={k:F(border=1,border_color=LINE,bg_color=b,font_color=c,bold=True,align='center') for k,(b,c) in _XL_GR.items()}
    f_kl={k:F(font_size=9,bold=True,font_color='#6B7280',bg_color='#FFFFFF',align='center',top=5,top_color=c,left=1,right=1,left_color=LINE,right_color=LINE) for k,c in
          {'all':'#6366F1','OnAir':'#10B981','OnAir Pending':'#F59E0B','Open PRI':'#3B82F6','Rejected RFAI':'#EF4444','RFAI Pending':'#94A3B8','p30':'#F97316','avg':'#A855F7'}.items()}
    f_kv=F(font_size=20,bold=True,font_color=INK,align='center',num_format='#,##0',bg_color='#FFFFFF',left=1,right=1,bottom=1,left_color=LINE,right_color=LINE,bottom_color=LINE)
    f_kvd=F(font_size=20,bold=True,font_color=INK,align='center',num_format='0" d"',bg_color='#FFFFFF',left=1,right=1,bottom=1,left_color=LINE,right_color=LINE,bottom_color=LINE)

    def dts(series):
        return pd.to_datetime(series.astype(str),format='%d-%b-%Y',errors='coerce')
    def dv_int(v):
        try: v=int(v); return None if v<0 else v
        except Exception: return None
    n=len(df)
    st=df['srStatus'].astype(str)
    delay=pd.to_numeric(df['delay'],errors='coerce').fillna(-1).astype(int)
    pend=st.isin(['OnAir Pending','Open PRI','RFAI Pending'])
    cnt={k:int((st==k).sum()) for k in _XL_STATUSES}
    p30=int((pend&(delay>30)).sum())
    vd=delay[delay>=0]; avg=int(round(vd.mean())) if len(vd) else 0

    # sheet names first (for the contents links)
    sheets=[]
    if include_journey:
        jparts=max(1,-(-n//_XL_MAXROWS)); sheets+=[('Journey Table' if jparts==1 else f'Journey Table {i+1}','Every SR with dates and its status month by month') for i in range(jparts)]
    DETAIL=[('Pending SRs',pend,'Not yet OnAir, oldest first (days since Final RFAI, PRI excl.)'),
            ('Open PRI',st=='Open PRI','Post-RFAI issue still open, oldest first'),
            ('Rejected RFAI',st=='Rejected RFAI','RFAI rejected, waiting re-offer'),
            ('OnAir SRs',st=='OnAir','On air, with days from Final RFAI to OnAir')]
    sheets+=[(nm,ds) for nm,_,ds in DETAIL]

    # ═══ Summary ═══
    ws=wb.add_worksheet('Summary'); ws.hide_gridlines(2); ws.set_zoom(90)
    ws.set_landscape(); ws.fit_to_pages(1,0); ws.set_paper(9)
    ws.set_column('A:A',2); ws.set_column('B:Q',13)
    ws.set_row(0,40); ws.merge_range('B1:Q1','  airtel  ·  RFAI Tracker  —  '+title,f_banner)
    ws.merge_range('B2:Q2','  Generated '+datetime.now().strftime('%d-%b-%Y %H:%M')+'   ·   '+f'{n:,}'+' SRs   ·   Filters: '+(filters_text or 'none (all SRs)'),f_bsub)
    ws.write('B4','Key numbers',f_h2)
    K=[('all','Total SRs',n,None),('OnAir','OnAir',cnt['OnAir'],None),('OnAir Pending','OnAir Pending',cnt['OnAir Pending'],None),
       ('Open PRI','Open PRI',cnt['Open PRI'],None),('Rejected RFAI','Rejected RFAI',cnt['Rejected RFAI'],None),
       ('RFAI Pending','RFAI Pending',cnt['RFAI Pending'],None),('p30','Pending > 30 days',p30,None),('avg','Average days',avg,'d')]
    for i,(k,l,v,u) in enumerate(K): ws.merge_range(5,1+i*2,5,2+i*2,l.upper(),f_kl[k])   # rows strictly top-down (streamed file)
    for i,(k,l,v,u) in enumerate(K): ws.merge_range(6,1+i*2,6,2+i*2,v,f_kvd if u else f_kv)
    ws.set_row(6,34)
    pct=lambda v:(v/n if n else 0)
    ws.write(7,1,f'OnAir {pct(cnt["OnAir"]):.0%} of SRs  ·  {pct(int(pend.sum())):.0%} still pending  ·  click a sheet name below to open it',f_note)
    # contents
    ws.write(9,1,'Sheets in this file',f_h2)
    r0=10
    for i,(nm,ds) in enumerate(sheets):
        ws.write_url(r0+i,1,f"internal:'{nm}'!A1",f_link,nm); ws.write(r0+i,3,ds,f_note)
    r=r0+len(sheets)+1
    # status mix
    ws.write(r,1,'Status mix',f_h2); r+=1
    hdr_r=r; ws.write_row(r,1,['Status','SRs','Share'],f_hdr); r+=1
    st_first=r
    for k in _XL_STATUSES:
        ws.write(r,1,k,f_st[k]); ws.write(r,2,cnt[k],f_num); ws.write(r,3,pct(cnt[k]),f_pct); r+=1
    ws.write(r,1,'Total',f_totl); ws.write(r,2,n,f_tot); ws.write(r,3,1 if n else 0,F(border=1,border_color=LINE,bold=True,bg_color='#FBEFF0',num_format='0%',align='center')); r+=1
    ch=wb.add_chart({'type':'doughnut'})
    ch.add_series({'name':'Status mix','categories':['Summary',st_first,1,st_first+4,1],'values':['Summary',st_first,2,st_first+4,2],
                   'points':[{'fill':{'color':{'OnAir':'#10B981','OnAir Pending':'#F59E0B','Open PRI':'#3B82F6','Rejected RFAI':'#EF4444','RFAI Pending':'#94A3B8'}[k]},'border':{'color':'#FFFFFF'}} for k in _XL_STATUSES],
                   'data_labels':{'percentage':True,'font':{'bold':True,'color':'#FFFFFF'}}})
    ch.set_hole_size(55); ch.set_title({'name':'Status mix','name_font':{'size':12}}); ch.set_legend({'position':'right'})
    ch.set_size({'width':420,'height':250}); ws.insert_chart(hdr_r,5,ch,{'x_offset':10})
    r=max(r,hdr_r+13)+1
    # by circle
    ws.write(r,1,'By circle',f_h2); r+=1
    g=pd.DataFrame({'c':df['circle'].astype(str),'s':st,'p30':(pend&(delay>30))})
    ct=pd.crosstab(g['c'],g['s']).reindex(columns=_XL_STATUSES,fill_value=0)
    ct['Total']=ct.sum(axis=1); ct=ct.sort_values('Total',ascending=False)
    p30c=g.groupby('c')['p30'].sum()
    cols=['Circle']+_XL_STATUSES+['Total','OnAir %','Pending >30 d']
    c_hdr=r; ws.write_row(r,1,cols,f_hdr); ws.set_row(r,30); r+=1; c_first=r
    for circ,row in ct.iterrows():
        if circ in ('','nan'): continue
        ws.write(r,1,circ,f_txtb)
        for j,k in enumerate(_XL_STATUSES): ws.write(r,2+j,int(row[k]),f_num)
        ws.write(r,7,int(row['Total']),f_num); ws.write(r,8,(row['OnAir']/row['Total']) if row['Total'] else 0,f_pct)
        ws.write(r,9,int(p30c.get(circ,0)),f_num); r+=1
    c_last=r-1
    if c_last>=c_first:
        ws.write(r,1,'Total',f_totl)
        for j,k in enumerate(_XL_STATUSES): ws.write(r,2+j,cnt[k],f_tot)
        ws.write(r,7,n,f_tot)
        ws.write(r,8,(cnt['OnAir']/n) if n else 0,F(border=1,border_color=LINE,bold=True,bg_color='#FBEFF0',num_format='0%',align='center'))
        ws.write(r,9,p30,f_tot); r+=1
        ws.conditional_format(c_first,8,c_last,8,{'type':'data_bar','bar_color':'#34D399','bar_solid':True,'min_type':'num','min_value':0,'max_type':'num','max_value':1})
        ws.conditional_format(c_first,9,c_last,9,{'type':'data_bar','bar_color':'#FB923C','bar_solid':True})
        ch=wb.add_chart({'type':'bar','subtype':'stacked'})
        for j,k in enumerate(_XL_STATUSES):
            ch.add_series({'name':['Summary',c_hdr,2+j],'categories':['Summary',c_first,1,c_last,1],'values':['Summary',c_first,2+j,c_last,2+j],
                           'fill':{'color':_XL_ST[k][1] if k!='OnAir' else '#10B981'},'gap':60})
        ch.set_title({'name':'SRs by circle and status','name_font':{'size':12}}); ch.set_legend({'position':'bottom'})
        chh=max(260,22*(c_last-c_first+1)+120)
        ch.set_y_axis({'reverse':True}); ch.set_size({'width':760,'height':chh})
        ws.insert_chart(c_hdr,11,ch,{'x_offset':10})
        r=max(r,c_hdr+int(chh/20)+2)
    r+=1
    # by month (exit month)
    ws.write(r,1,'By month (exit month: OnAir month for live sites, RFAI month otherwise)',f_h2); r+=1
    g=pd.DataFrame({'m':df['exitMK'].astype(str),'s':st})
    mt=pd.crosstab(g['m'],g['s']).reindex(columns=_XL_STATUSES,fill_value=0)
    mt=mt.loc[sorted([m for m in mt.index if m and m!='nan'],key=mk_sort_val)]
    m_hdr=r; ws.write_row(r,1,['Month']+_XL_STATUSES+['Total','OnAir %'],f_hdr); ws.set_row(r,30); r+=1; m_first=r
    for m,row in mt.iterrows():
        t=int(row.sum()); ws.write(r,1,m,f_txtb)
        for j,k in enumerate(_XL_STATUSES): ws.write(r,2+j,int(row[k]),f_num)
        ws.write(r,7,t,f_num); ws.write(r,8,(row['OnAir']/t) if t else 0,f_pct); r+=1
    m_last=r-1
    if m_last>=m_first:
        ws.conditional_format(m_first,7,m_last,7,{'type':'data_bar','bar_color':'#F87171','bar_solid':True})
        ch=wb.add_chart({'type':'column','subtype':'stacked'})
        for j,k in enumerate(_XL_STATUSES):
            ch.add_series({'name':['Summary',m_hdr,2+j],'categories':['Summary',m_first,1,m_last,1],'values':['Summary',m_first,2+j,m_last,2+j],
                           'fill':{'color':_XL_ST[k][1] if k!='OnAir' else '#10B981'},'gap':50})
        ch.set_title({'name':'SRs by month and status','name_font':{'size':12}}); ch.set_legend({'position':'bottom'})
        ch.set_size({'width':760,'height':300}); ws.insert_chart(m_hdr,11,ch,{'x_offset':10})
    r=max(r,m_hdr+16)+1
    # pending ageing
    ws.write(r,1,'Pending ageing (days since Final RFAI, PRI days excluded)',f_h2); r+=1
    ws.write_row(r,1,['Days','Pending SRs','Share'],f_hdr); r+=1
    pd_=delay[pend]; tp=int(pend.sum())
    for lab,lo,hi in (('0-30',0,30),('30-60',31,60),('60-90',61,90),('>90',91,10**6)):
        v=int(((pd_>=lo)&(pd_<=hi)).sum())
        ws.write(r,1,lab+' days',f_gr[lab]); ws.write(r,2,v,f_num); ws.write(r,3,(v/tp) if tp else 0,f_pct); r+=1
    ws.write(r,1,'Total pending',f_totl); ws.write(r,2,tp,f_tot); r+=1
    ws.activate()

    def table_sheets(base_name, cols, row_iter, total, freeze_col=2, bands=None, widths=None):
        """Write a list over as many sheets as needed (1,000,000 rows each).
        cols: [(header, kind)] kind in txt|txtb|num|date|status|group|days."""
        parts=max(1,-(-total//_XL_MAXROWS)); it=iter(row_iter)
        for pi in range(parts):
            name=base_name if parts==1 else f'{base_name} {pi+1}'
            w=wb.add_worksheet(name); w.set_zoom(90); w.hide_gridlines(2)
            w.set_landscape(); w.set_paper(9); w.repeat_rows(0,1 if bands else 0)
            if not bands: w.fit_to_pages(1,0)
            else: w.repeat_columns(0,1)
            hr=0
            if bands:
                for (c0,c1,label,fmt) in bands:
                    if c1>c0: w.merge_range(0,c0,0,c1,label,fmt)
                    else: w.write(0,c0,label,fmt)
                hr=1
            w.write_row(hr,0,[c[0] for c in cols],f_hdr); w.set_row(hr,30)
            for j,(h,k) in enumerate(cols):
                wd=(widths or {}).get(j) or {'date':12.5,'num':9,'days':8,'status':15,'group':9,'txtb':18}.get(k,14)
                w.set_column(j,j,wd)
            w.freeze_panes(hr+1,freeze_col)
            rr=hr+1; cnt_rows=0
            for vals in it:
                for j,((h,k),v) in enumerate(zip(cols,vals)):
                    if v is None or v=='' : w.write_blank(rr,j,None,f_txt); continue
                    if k=='date':
                        if isinstance(v,pd.Timestamp) and not pd.isna(v): w.write_datetime(rr,j,v.to_pydatetime(),f_date)
                        else: w.write_string(rr,j,str(v),f_txt)
                    elif k in ('num','days'): w.write_number(rr,j,v,f_num)
                    elif k=='status': w.write_string(rr,j,v,f_st.get(v,f_txt))
                    elif k=='group': w.write_string(rr,j,GLS.get(v,v),f_gr.get(v,f_txt))
                    elif k=='txtb': w.write_string(rr,j,str(v),f_txtb)
                    else: w.write_string(rr,j,str(v),f_txt)
                rr+=1; cnt_rows+=1
                if cnt_rows>=_XL_MAXROWS: break
            last=max(rr-1,hr+1)
            w.autofilter(hr,0,last,len(cols)-1)
            for j,(h,k) in enumerate(cols):
                if k=='days' and cnt_rows: w.conditional_format(hr+1,j,last,j,{'type':'data_bar','bar_color':'#FCA5A5','bar_solid':True,'min_type':'num','min_value':0})
            w.write_url(last+2,0,"internal:'Summary'!A1",f_link,'◂ Back to Summary')

    DATE_COLS=['rfaiNotice','rfaiRej','rfaiReoffer','finalRfai','priFirst','priLast','onAir','baseDate']
    D={c:(dts(df[c]) if c in df.columns else pd.Series(pd.NaT,index=df.index)) for c in DATE_COLS}
    def g(c): return df[c].astype(str).where(lambda x:~x.isin(['nan','None']),'') if c in df.columns else pd.Series('',index=df.index)

    if include_journey:
        all_mks=(set(df['exitMK'].dropna().astype(str))|set(df['journeyMK'].dropna().astype(str)))
        months=sorted([m for m in all_mks if m and m!='nan'],key=mk_sort_val)
        base=[('SR #','txtb'),('2G Site ID','txt'),('Circle','txt'),('ToCo Name','txt'),('Project Type','txt'),('OEM','txt'),('TSP','txt'),
              ('Partner','txt'),('TOCO Project','txt'),('Broad Project','txt'),('Initial RFAI','date'),('RFAI Rejection','date'),
              ('RFAI Re-offer','date'),('Final RFAI','date'),('1st PRI Start','date'),('Last PRI Close','date'),('PRI Days','num'),
              ('OnAir Date','date'),('Base Date','date'),('Base','txt')]
        cols=list(base)
        for m in months: cols+=[('Status','status'),('Days','days'),('Group','group')]
        nb=len(base)
        bands=[(0,nb-1,'BASE INFORMATION',f_band2)]+[(nb+3*i,nb+3*i+2,m,f_band) for i,m in enumerate(months)]
        GV={c:g(c).tolist() for c in ['srNo','siteId','circle','toco','domain','oem','tsp','partner','bPrj','category','baseLabel']}
        DV={c:D[c].tolist() for c in DATE_COLS}
        def jrows():
            def recs():            # in chunks: a dict per row for the whole table would not fit in 512 MB
                for a in range(0,n,20000):
                    yield from df.iloc[a:a+20000].to_dict('records')
            for i,rec in enumerate(recs()):
                pf=rec.get('priFirst',''); pdays=dv_int(rec.get('priDays',0))
                row=[GV['srNo'][i],GV['siteId'][i],GV['circle'][i],GV['toco'][i],GV['domain'][i],GV['oem'][i],GV['tsp'][i],GV['partner'][i],
                     GV['bPrj'][i],GV['category'][i],DV['rfaiNotice'][i],DV['rfaiRej'][i],DV['rfaiReoffer'][i],DV['finalRfai'][i],
                     DV['priFirst'][i],DV['priLast'][i],(pdays if pf and pdays else None),DV['onAir'][i],DV['baseDate'][i],GV['baseLabel'][i]]
                row=[None if (isinstance(v,float) and v!=v) or (v is pd.NaT) else v for v in row]
                for c in month_cells(rec,months):
                    if c is None: row+=[None,None,None]
                    else:
                        stt,dvs,grp=c
                        try: dn=int(dvs)
                        except Exception: dn=None
                        row+=[stt,dn,grp]
                yield row
        table_sheets('Journey Table',cols,jrows(),n,freeze_col=2,bands=bands)

    # detail lists
    det_cols=[('SR #','txtb'),('2G Site ID','txt'),('Circle','txt'),('Status','status'),('Days','days'),('Ageing','group'),
              ('Project Type','txt'),('TOCO Project','txt'),('Broad Project','txt'),('ToCo Name','txt'),('OEM','txt'),('Initial RFAI','date'),
              ('Final RFAI','date'),('1st PRI Start','date'),('Last PRI Close','date'),('PRI Days','num'),('OnAir Date','date')]
    for nm,mask,_ in DETAIL:
        idx=np.where(mask.to_numpy())[0]
        order=idx[np.argsort(-delay.to_numpy()[idx],kind='stable')]
        S={c:g(c).to_numpy() for c in ['srNo','siteId','circle','srStatus','group','domain','bPrj','category','toco','oem']}
        Dd={c:D[c].to_numpy() for c in ['rfaiNotice','finalRfai','priFirst','priLast','onAir']}
        pdv=pd.to_numeric(df['priDays'],errors='coerce').fillna(0).astype(int).to_numpy() if 'priDays' in df.columns else np.zeros(n,int)
        dl=delay.to_numpy()
        def drows(order=order):
            for i in order:
                def dd(c):
                    v=Dd[c][i]; return None if pd.isna(v) else pd.Timestamp(v)
                yield [S['srNo'][i],S['siteId'][i],S['circle'][i],S['srStatus'][i],(int(dl[i]) if dl[i]>=0 else None),
                       (S['group'][i] or 'N/A'),S['domain'][i],S['bPrj'][i],S['category'][i],S['toco'][i],S['oem'][i],dd('rfaiNotice'),dd('finalRfai'),
                       dd('priFirst'),dd('priLast'),(int(pdv[i]) if pdv[i]>0 else None),dd('onAir')]
        table_sheets(nm,det_cols,drows(),len(order),freeze_col=1)
    wb.close()

@app.get("/api/export-xlsx")
def export_xlsx(request:Request, background_tasks:BackgroundTasks, journey:int=Query(default=1), title:str=Query(default='')):
    """Formatted Excel: Summary (key numbers, charts, contents), the Journey Table
    (split over several sheets when needed) and Pending / Open PRI / Rejected / OnAir lists."""
    import tempfile
    df=_df_full_cache
    if df is None or not len(df): return JSONResponse({"error":"No data loaded"},status_code=404)
    p={k:v for k,v in request.query_params.items() if v}
    df=_xl_filter(df,p).reset_index(drop=True)
    if not len(df): return JSONResponse({"error":"No SRs match these filters"},status_code=404)
    LBL={'circle':'Circle','category':'Category','project':'Project','toco':'ToCo','domain':'Domain','oem':'OEM','status':'Status',
         'month':'Month','search':'Search','q':'Search','aging':'Days','recent':'Recent'}
    ftxt='  ·  '.join(f'{LBL[k]}: {v}' for k,v in p.items() if k in LBL)
    fd,path=tempfile.mkstemp(suffix='.xlsx',dir=str(_cache_dir) if _cache_dir.exists() else None); os.close(fd)
    try:
        build_export_xlsx(df,path,ftxt,title or 'Journey export',include_journey=bool(journey))
    except Exception as e:
        try: os.remove(path)
        except Exception: pass
        log.error(f"  Export failed: {e}"); return JSONResponse({"error":f"Export failed: {e}"},status_code=500)
    background_tasks.add_task(lambda: os.path.exists(path) and os.remove(path))
    name='RFAI_'+re.sub(r'[^A-Za-z0-9]+','_',title or 'Journey')[:40].strip('_')+'_'+datetime.now().strftime('%d%b%Y_%H%M')+'.xlsx'
    return FileResponse(path,filename=name,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.get("/")
def root():
    if HTML_FILE.exists(): return FileResponse(str(HTML_FILE))
    return JSONResponse({"error":"HTML not found"},status_code=404)

@app.get("/{path:path}")
def static(path:str):
    f=(SCRIPT_DIR/path).resolve()
    if (f.is_file() and f.is_relative_to(SCRIPT_DIR.resolve())
            and f.suffix.lower() in ('.html','.css','.js','.png','.jpg','.svg','.ico','.json','.webmanifest')):
        if f.suffix.lower()=='.webmanifest': return FileResponse(str(f),media_type='application/manifest+json')
        if f.name=='sw.js': return FileResponse(str(f),media_type='text/javascript',headers={'Cache-Control':'no-cache'})
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
