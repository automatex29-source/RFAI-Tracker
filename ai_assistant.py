"""
RFAI AI assistant — answers questions about the loaded RFAI data and builds files
(Excel, PowerPoint, chart images) on request, using Claude with tools.

The model never sees the raw table: it calls the tools below, which run on this
server against the in-memory data, and gets back only the rows/aggregates it asked for.
Engines, picked from the environment (first one set wins):
  ANTHROPIC_API_KEY  -> Claude (paid)
  GEMINI_API_KEY     -> Google Gemini free tier (no card needed)
  GROQ_API_KEY       -> Groq free tier (no card needed)
  none               -> built-in rules engine (ai_local.py): free, offline, data never leaves the server
If a free engine is over its limit or fails, the answer falls back to the built-in engine.
"""
import os, re, json, uuid, time, threading
from datetime import date, datetime
from pathlib import Path

import numpy as np
import pandas as pd

MODEL = os.environ.get('AI_MODEL', 'claude-opus-5-5')
FREE_ENGINES = {   # OpenAI-compatible chat endpoints with free tiers
    'gemini': ('https://generativelanguage.googleapis.com/v1beta/openai/chat/completions', 'gemini-flash-latest', 'Gemini'),
    'groq': ('https://api.groq.com/openai/v1/chat/completions', 'llama-3.3-70b-versatile', 'Groq'),
}
MAX_TOOL_ROUNDS = 12
_conversations = {}            # conv_id -> {"user":..., "messages":[...], "t":...}
_conv_lock = threading.Lock()

STATUSES = ['OnAir', 'OnAir Pending', 'Open PRI', 'Rejected RFAI', 'RFAI Pending']
GROUP_COLS = {'circle': 'circle', 'status': 'srStatus', 'category': 'category', 'project': 'bPrj',
              'toco': 'toco', 'oem': 'oem', 'domain': 'domain', 'tsp': 'tsp', 'partner': 'partner',
              'exit_month': 'exitMK', 'rfai_month': 'journeyMK', 'ageing_group': 'group'}
DATE_COLS = {'initial_rfai': 'rfaiNotice', 'rfai_rejection': 'rfaiRej', 'rfai_reoffer': 'rfaiReoffer',
             'final_rfai': 'finalRfai', 'first_pri_start': 'priFirst', 'last_pri_close': 'priLast', 'onair': 'onAir'}

SYSTEM = """You are the RFAI Tracker assistant for Airtel's 5G network deployment team.
You answer questions about the RFAI (Ready For Installation) data loaded in this tool and you
create files people ask for: Excel workbooks, PowerPoint decks and charts.

How the data works:
- One row per SR (site request). Key fields: SR number, 2G Site ID, circle, ToCo, project type,
  OEM, TSP, partner, TOCO project (category), broad project, and dates: initial RFAI, RFAI
  rejection, RFAI re-offer, final RFAI, first PRI start, last PRI close, OnAir.
- Status: OnAir, OnAir Pending, Open PRI (a Post-RFAI Issue is still open), Rejected RFAI,
  RFAI Pending. "Pending" means OnAir Pending + Open PRI + RFAI Pending.
- Days (delay): for pending SRs, days since Final RFAI with PRI days excluded; for OnAir SRs,
  days from Final RFAI to OnAir. Ageing groups: 0-30, 30-60, 60-90, >90 days.
- Exit month = OnAir month for live sites, otherwise the RFAI month.

Rules:
- Always get numbers from the tools. Never guess or invent figures. If the data can't answer
  a question, say so plainly.
- Start with get_overview when you need the available circles, months or field values.
- Prefer aggregate for counts and breakdowns; use list_srs only when the user wants specific SRs.
- When the user asks for a file (Excel, PPT, chart, report, presentation), build it with the
  file tools and mention it; the download button appears in the chat automatically.
- Keep answers short and scannable: a one-line answer first, then a compact markdown table or
  bullets if useful. Use Indian number formatting only if the user does. Today is {today}.
"""

FILTER_PROPS = {
    "circle": {"type": "string", "description": "Comma-separated circle names, e.g. 'MH,DL'"},
    "status": {"type": "string", "description": "Comma-separated statuses (OnAir, OnAir Pending, Open PRI, Rejected RFAI, RFAI Pending) or 'pending' for all not-yet-OnAir"},
    "category": {"type": "string", "description": "Comma-separated TOCO project / category values"},
    "project": {"type": "string", "description": "Comma-separated broad project values"},
    "toco": {"type": "string"}, "oem": {"type": "string"}, "domain": {"type": "string", "description": "Project type"},
    "month": {"type": "string", "description": "Comma-separated exit months like \"Jan'26,Feb'26\""},
    "aging": {"type": "string", "enum": ["0-30", "30-60", "60-90", ">90", ">30"]},
    "search": {"type": "string", "description": "Part of an SR number or site ID"},
    "date_field": {"type": "string", "enum": list(DATE_COLS), "description": "Restrict to SRs whose date in this field falls between date_from and date_to"},
    "date_from": {"type": "string", "description": "YYYY-MM-DD"},
    "date_to": {"type": "string", "description": "YYYY-MM-DD"},
}

def _tools():
    flt = {"type": "object", "properties": FILTER_PROPS}
    return [
        {"name": "get_overview",
         "description": "Totals, status mix, the list of circles, exit months, categories, projects, OEMs and ToCos in the loaded data, and the data's date range. Call first when unsure what values exist.",
         "input_schema": {"type": "object", "properties": {}}},
        {"name": "aggregate",
         "description": "Count SRs (and average/max days) grouped by up to two fields, with optional filters. Use for any 'how many', 'by circle', 'per month', 'breakdown', 'top N' question.",
         "input_schema": {"type": "object", "properties": {
             "group_by": {"type": "array", "items": {"type": "string", "enum": list(GROUP_COLS)}, "maxItems": 2,
                          "description": "Fields to group by; empty for a single total"},
             "date_group": {"type": "object", "description": "Group instead by a date field per day/week/month",
                            "properties": {"field": {"type": "string", "enum": list(DATE_COLS)},
                                           "period": {"type": "string", "enum": ["day", "week", "month"]}}},
             "filters": flt,
             "top": {"type": "integer", "description": "Keep only the largest N groups"}}}},
        {"name": "list_srs",
         "description": "List individual SRs with their key fields, filtered and sorted. Returns at most 100 rows plus the total match count.",
         "input_schema": {"type": "object", "properties": {
             "filters": flt,
             "sort_by": {"type": "string", "enum": ["days", "sr", "circle", "initial_rfai", "final_rfai", "onair"]},
             "descending": {"type": "boolean"}, "limit": {"type": "integer", "maximum": 100}}}},
        {"name": "get_sr",
         "description": "Everything known about one SR (all fields and its journey dates).",
         "input_schema": {"type": "object", "properties": {"sr": {"type": "string"}}, "required": ["sr"]}},
        {"name": "make_data_excel",
         "description": "Build the standard formatted RFAI Excel workbook (summary sheet with charts, Journey Table, pending / PRI / rejected / OnAir lists) for SRs matching the filters.",
         "input_schema": {"type": "object", "properties": {"filters": flt, "title": {"type": "string"}}, "required": ["title"]}},
        {"name": "make_custom_excel",
         "description": "Build a formatted Excel workbook from tables you supply (e.g. aggregate results). Each sheet: name, columns, rows; optional chart.",
         "input_schema": {"type": "object", "properties": {
             "title": {"type": "string"},
             "sheets": {"type": "array", "items": {"type": "object", "properties": {
                 "name": {"type": "string"}, "columns": {"type": "array", "items": {"type": "string"}},
                 "rows": {"type": "array", "items": {"type": "array"}},
                 "chart": {"type": "object", "properties": {"type": {"type": "string", "enum": ["column", "bar", "line", "pie"]},
                                                            "title": {"type": "string"}}}},
                 "required": ["name", "columns", "rows"]}}},
             "required": ["title", "sheets"]}},
        {"name": "make_powerpoint",
         "description": "Build an Airtel-styled PowerPoint deck. Slide kinds: 'title' (title, subtitle), 'bullets' (title, bullets), 'kpis' (title, kpis:[{label,value}]), 'table' (title, columns, rows), 'chart' (title, chart_type, categories, series:[{name,values}]).",
         "input_schema": {"type": "object", "properties": {
             "title": {"type": "string"},
             "slides": {"type": "array", "items": {"type": "object", "properties": {
                 "kind": {"type": "string", "enum": ["title", "bullets", "kpis", "table", "chart"]},
                 "title": {"type": "string"}, "subtitle": {"type": "string"},
                 "bullets": {"type": "array", "items": {"type": "string"}},
                 "kpis": {"type": "array", "items": {"type": "object", "properties": {"label": {"type": "string"}, "value": {"type": "string"}}}},
                 "columns": {"type": "array", "items": {"type": "string"}}, "rows": {"type": "array", "items": {"type": "array"}},
                 "chart_type": {"type": "string", "enum": ["column", "bar", "line", "pie", "stacked_column"]},
                 "categories": {"type": "array", "items": {"type": "string"}},
                 "series": {"type": "array", "items": {"type": "object", "properties": {"name": {"type": "string"}, "values": {"type": "array", "items": {"type": "number"}}}}},
                 "notes": {"type": "string"}},
                 "required": ["kind", "title"]}}},
             "required": ["title", "slides"]}},
        {"name": "make_chart_image",
         "description": "Draw a chart as a PNG image the user can download or paste anywhere.",
         "input_schema": {"type": "object", "properties": {
             "title": {"type": "string"}, "chart_type": {"type": "string", "enum": ["column", "bar", "line", "pie", "stacked_column"]},
             "categories": {"type": "array", "items": {"type": "string"}},
             "series": {"type": "array", "items": {"type": "object", "properties": {"name": {"type": "string"}, "values": {"type": "array", "items": {"type": "number"}}}}}},
             "required": ["title", "chart_type", "categories", "series"]}},
    ]


def _loose(schema):
    """Schema for OpenAI-compatible engines: drop limits some providers reject, give every array its items."""
    if isinstance(schema, dict):
        out = {k: _loose(v) for k, v in schema.items() if k not in ('maxItems', 'maximum')}
        if out.get('type') == 'array' and not out.get('items'): out['items'] = {"type": "string"}
        if out.get('type') == 'object' and 'properties' in out and not out['properties']: out.pop('properties')
        return out
    if isinstance(schema, list): return [_loose(x) for x in schema]
    return schema


def _num(v):
    if isinstance(v, str) and re.fullmatch(r'-?\d[\d,]*(\.\d+)?', v.strip()):
        try: return float(v.replace(',', '')) if '.' in v else int(v.replace(',', ''))
        except ValueError: return v
    return v


ST_COL = {'onair': '#10B981', 'onair pending': '#F59E0B', 'pending': '#F59E0B', 'open pri': '#3B82F6',
          'rejected rfai': '#EF4444', 'rejected': '#EF4444', 'rfai pending': '#94A3B8'}
def _col(name, i, pal):
    return ST_COL.get(str(name).strip().lower(), pal[i % len(pal)])


class Assistant:
    """Bound to the host app: get_df() returns the current data, build_xlsx is the export builder."""

    def __init__(self, get_df, xl_filter, build_xlsx, mk_sort_val, files_dir):
        self.get_df, self.xl_filter, self.build_xlsx, self.mk_sort_val = get_df, xl_filter, build_xlsx, mk_sort_val
        self.files_dir = Path(files_dir); self.files_dir.mkdir(parents=True, exist_ok=True)
        self.files = {}           # file_id -> {"path", "name", "user", "t"}
        self._client = None

    # ── public ──────────────────────────────────────────────────────────────
    def enabled(self):
        return True

    def engine(self):
        if os.environ.get('ANTHROPIC_API_KEY') or os.environ.get('ANTHROPIC_AUTH_TOKEN'): return 'claude'
        if os.environ.get('GEMINI_API_KEY') or os.environ.get('GOOGLE_API_KEY'): return 'gemini'
        if os.environ.get('GROQ_API_KEY'): return 'groq'
        return 'builtin'

    def engine_label(self):
        return {'claude': 'Claude AI', 'gemini': 'Gemini AI · free', 'groq': 'Groq AI · free'}.get(self.engine(), 'Built-in · free & offline')

    def chat(self, user, conv_id, text):
        """Returns (conv_id, answer_markdown, files, suggestions)."""
        with _conv_lock:
            self._gc()
            conv = _conversations.get(conv_id) if conv_id else None
            if not conv or conv['user'] != user:
                conv_id = uuid.uuid4().hex; conv = {'user': user, 'messages': [], 'omsgs': [], 'last': None, 't': time.time()}
                _conversations[conv_id] = conv
        conv['t'] = time.time()
        eng = self.engine()
        if eng == 'builtin':
            return (conv_id,) + self._local(conv, text, user)
        try:
            if eng == 'claude':
                ans, files = self._chat_claude(conv, text, user)
            else:
                ans, files = self._chat_openai(conv, text, user, eng)
            return conv_id, ans, files, []
        except Exception as e:
            if eng == 'claude': raise
            ans, files, sug = self._local(conv, text, user)
            why = 'is over its free limit right now' if '429' in str(e) else 'did not answer'
            return conv_id, f"_{FREE_ENGINES[eng][2]} {why}, so the built-in assistant answered._\n\n" + ans, files, sug

    def _local(self, conv, text, user):
        from ai_local import LocalBrain
        ans, files, state, sug = LocalBrain(self).answer(text, user, conv.get('last'))
        if state is not None: conv['last'] = state
        return ans, files, sug

    def _chat_openai(self, conv, text, user, eng):
        """Tool loop against an OpenAI-compatible endpoint (Gemini / Groq free tiers)."""
        import requests
        url, model, _ = FREE_ENGINES[eng]
        key = (os.environ.get('GEMINI_API_KEY') or os.environ.get('GOOGLE_API_KEY')) if eng == 'gemini' else os.environ.get('GROQ_API_KEY')
        model = os.environ.get('AI_FREE_MODEL') or model
        tools = [{"type": "function", "function": {"name": t['name'], "description": t['description'],
                  **({"parameters": _loose(t['input_schema'])} if t['input_schema'].get('properties') else {})}} for t in _tools()]
        msgs = conv['omsgs']; start = len(msgs)
        msgs.append({"role": "user", "content": text}); made = []
        sysmsg = {"role": "system", "content": SYSTEM.replace('{today}', date.today().strftime('%d %b %Y'))}
        try:
            for _ in range(MAX_TOOL_ROUNDS):
                r = requests.post(url, headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                                  json={"model": model, "messages": [sysmsg] + msgs[-40:], "tools": tools, "temperature": 0.2}, timeout=120)
                if r.status_code >= 400:
                    raise RuntimeError(f"{r.status_code}: {r.text[:300]}")
                m = r.json()['choices'][0]['message']
                calls = m.get('tool_calls') or []
                msgs.append({"role": "assistant", "content": m.get('content') or '', **({"tool_calls": calls} if calls else {})})
                if not calls:
                    return (m.get('content') or 'Done.').strip(), made
                for c in calls:
                    try:
                        args = json.loads(c['function'].get('arguments') or '{}') or {}
                        out = self._run_tool(c['function']['name'], args, user)
                        if isinstance(out, dict) and out.get('_file'):
                            made.append(out['_file']); out = {k: v for k, v in out.items() if k != '_file'}
                        content = json.dumps(out, default=str)[:30000]
                    except Exception as e:
                        content = f"Error: {e}"
                    msgs.append({"role": "tool", "tool_call_id": c.get('id', ''), "name": c['function']['name'], "content": content})
            return "That took too many steps. Try a more specific question.", made
        except Exception:
            del msgs[start:]   # keep the history valid
            raise

    def _chat_claude(self, conv, text, user):
        import anthropic
        if self._client is None:
            self._client = anthropic.Anthropic(max_retries=3)
        msgs = conv['messages']; msgs.append({"role": "user", "content": text})
        made = []; tools = _tools()
        system = [{"type": "text", "text": SYSTEM.replace('{today}', date.today().strftime('%d %b %Y')),
                   "cache_control": {"type": "ephemeral"}}]
        for _ in range(MAX_TOOL_ROUNDS):
            resp = self._client.beta.messages.create(
                model=MODEL, max_tokens=16000, system=system, tools=tools, messages=msgs,
                output_config={"effort": "medium"},
                betas=["server-side-fallback-2026-07-01"], fallbacks="default",
            )
            if resp.stop_reason == "refusal":
                msgs.pop()   # keep history valid for the next question
                return "Sorry, I can't help with that request.", made
            msgs.append({"role": "assistant", "content": resp.content})
            if resp.stop_reason != "tool_use":
                answer = "\n".join(b.text for b in resp.content if b.type == "text").strip()
                return answer or "Done.", made
            results = []
            for b in resp.content:
                if b.type != "tool_use": continue
                try:
                    out = self._run_tool(b.name, b.input or {}, user)
                    if isinstance(out, dict) and out.get('_file'):
                        made.append(out['_file'])
                        out = {k: v for k, v in out.items() if k != '_file'}
                    results.append({"type": "tool_result", "tool_use_id": b.id, "content": json.dumps(out, default=str)[:60000]})
                except Exception as e:
                    results.append({"type": "tool_result", "tool_use_id": b.id, "content": f"Error: {e}", "is_error": True})
            msgs.append({"role": "user", "content": results})
        return "That took too many steps. Try a more specific question.", made

    def file(self, file_id, user):
        f = self.files.get(file_id)
        if not f or f['user'] != user or not os.path.exists(f['path']): return None
        return f

    # ── tools ───────────────────────────────────────────────────────────────
    def _df(self, filters=None):
        df = self.get_df()
        if df is None or not len(df): raise RuntimeError("No data is loaded yet.")
        p = {k: str(v) for k, v in (filters or {}).items() if v not in (None, '') and k in FILTER_PROPS}
        if p.get('date_field'):
            p['ev'] = {'initial_rfai': 'rfai', 'final_rfai': 'final', 'first_pri_start': 'pri_open',
                       'last_pri_close': 'pri_close', 'onair': 'onair', 'rfai_rejection': 'rej'}.get(p.pop('date_field'), '')
            if not p['ev']: p.pop('ev')
            if 'date_from' in p: p['dfrom'] = p.pop('date_from')
            if 'date_to' in p: p['dto'] = p.pop('date_to')
        else:
            p.pop('date_from', None); p.pop('date_to', None)
        return self.xl_filter(df, p)

    @staticmethod
    def _days(df):
        return pd.to_numeric(df['delay'], errors='coerce').where(lambda x: x >= 0)

    def _run_tool(self, name, a, user):
        if name == 'get_overview':
            df = self.get_df()
            if df is None or not len(df): return {"error": "No data is loaded yet."}
            st = df['srStatus'].astype(str).value_counts().to_dict()
            vals = lambda c: sorted(x for x in df[c].astype(str).unique() if x and x != 'nan')[:80] if c in df.columns else []
            d = pd.to_datetime(df['rfaiNotice'].astype(str), format='%d-%b-%Y', errors='coerce')
            return {"total_srs": len(df), "statuses": st, "circles": vals('circle'),
                    "exit_months": sorted(vals('exitMK'), key=self.mk_sort_val), "categories": vals('category'),
                    "projects": vals('bPrj'), "oems": vals('oem'), "tocos": vals('toco'), "project_types": vals('domain'),
                    "initial_rfai_from": str(d.min().date()) if d.notna().any() else None,
                    "initial_rfai_to": str(d.max().date()) if d.notna().any() else None}
        if name == 'aggregate':
            df = self._df(a.get('filters'))
            days = self._days(df)
            keys = []
            dg = a.get('date_group') or {}
            if dg.get('field'):
                d = pd.to_datetime(df[DATE_COLS[dg['field']]].astype(str), format='%d-%b-%Y', errors='coerce')
                per = dg.get('period', 'month')
                lab = d.dt.strftime('%Y-%m-%d') if per == 'day' else (d - pd.to_timedelta(d.dt.weekday, unit='D')).dt.strftime('week of %Y-%m-%d') if per == 'week' else d.dt.strftime('%Y-%m')
                keys.append(lab.rename(dg['field'] + '_' + per))
            for g in (a.get('group_by') or [])[:2]:
                keys.append(df[GROUP_COLS[g]].astype(str).rename(g))
            if not keys:
                return {"total": len(df), "avg_days": None if days.isna().all() else round(float(days.mean()), 1),
                        "statuses": df['srStatus'].astype(str).value_counts().to_dict()}
            g = pd.DataFrame({k.name: k for k in keys}).assign(_d=days)
            out = g.groupby([k.name for k in keys], dropna=True).agg(srs=('_d', 'size'), avg_days=('_d', 'mean'), max_days=('_d', 'max')).reset_index()
            out = out.sort_values('srs', ascending=False)
            if dg.get('field'): out = out.sort_values(keys[0].name)
            if a.get('top'): out = out.nlargest(int(a['top']), 'srs')
            out['avg_days'] = out['avg_days'].round(1)
            rows = out.head(300).replace({np.nan: None}).to_dict('records')
            return {"total_srs": len(df), "groups": len(out), "rows": rows, "truncated": len(out) > 300}
        if name == 'list_srs':
            df = self._df(a.get('filters'))
            sk = {'days': None, 'sr': 'srNo', 'circle': 'circle', 'initial_rfai': 'rfaiNotice', 'final_rfai': 'finalRfai', 'onair': 'onAir'}.get(a.get('sort_by') or 'days')
            desc = a.get('descending', True)
            if sk is None: key = self._days(df)
            elif sk in ('rfaiNotice', 'finalRfai', 'onAir'): key = pd.to_datetime(df[sk].astype(str), format='%d-%b-%Y', errors='coerce')
            else: key = df[sk].astype(str)
            df = df.loc[key.sort_values(ascending=not desc, na_position='last').index]
            n = min(int(a.get('limit') or 25), 100)
            cols = ['srNo', 'siteId', 'circle', 'srStatus', 'delay', 'bPrj', 'category', 'oem', 'rfaiNotice', 'finalRfai', 'priFirst', 'priLast', 'onAir']
            sub = df[[c for c in cols if c in df.columns]].head(n).astype(str).replace({'nan': '', 'None': '', '-1': ''})
            return {"total_matches": len(df), "shown": len(sub), "rows": sub.to_dict('records')}
        if name == 'get_sr':
            df = self.get_df()
            hit = df[df['srNo'].astype(str).str.upper() == str(a['sr']).strip().upper()]
            if not len(hit): hit = df[df['srNo'].astype(str).str.upper().str.contains(str(a['sr']).strip().upper(), regex=False)].head(5)
            if not len(hit): return {"error": "SR not found"}
            return {"matches": hit.astype(str).replace({'nan': ''}).head(5).to_dict('records')}
        if name == 'make_data_excel':
            df = self._df(a.get('filters')).reset_index(drop=True)
            if not len(df): return {"error": "No SRs match these filters."}
            ftxt = ' · '.join(f"{k}: {v}" for k, v in (a.get('filters') or {}).items() if v)
            path, fid, fname = self._new_file(a.get('title') or 'RFAI export', '.xlsx', user)
            self.build_xlsx(df, path, ftxt, a.get('title') or 'RFAI export', include_journey=True)
            return {"ok": True, "srs": len(df), "file": fname, "_file": self._meta(fid)}
        if name == 'make_custom_excel':
            path, fid, fname = self._new_file(a['title'], '.xlsx', user)
            self._custom_xlsx(path, a['title'], a['sheets'])
            return {"ok": True, "file": fname, "_file": self._meta(fid)}
        if name == 'make_powerpoint':
            path, fid, fname = self._new_file(a['title'], '.pptx', user)
            self._pptx(path, a['title'], a['slides'])
            return {"ok": True, "file": fname, "slides": len(a['slides']), "_file": self._meta(fid)}
        if name == 'make_chart_image':
            path, fid, fname = self._new_file(a['title'], '.png', user)
            self._png(path, a)
            return {"ok": True, "file": fname, "_file": self._meta(fid)}
        return {"error": f"Unknown tool {name}"}

    # ── files ───────────────────────────────────────────────────────────────
    def _new_file(self, title, ext, user):
        fid = uuid.uuid4().hex
        fname = 'RFAI_' + re.sub(r'[^A-Za-z0-9]+', '_', title)[:50].strip('_') + '_' + datetime.now().strftime('%d%b%Y_%H%M') + ext
        path = str(self.files_dir / (fid + ext))
        self.files[fid] = {"path": path, "name": fname, "user": user, "t": time.time()}
        return path, fid, fname

    def _meta(self, fid):
        f = self.files[fid]; return {"id": fid, "name": f['name'], "kind": f['name'].rsplit('.', 1)[-1]}

    def _gc(self):
        now = time.time()
        for k in [k for k, v in _conversations.items() if now - v['t'] > 6 * 3600]: _conversations.pop(k, None)
        for k in [k for k, v in self.files.items() if now - v['t'] > 6 * 3600]:
            try: os.remove(self.files[k]['path'])
            except Exception: pass
            self.files.pop(k, None)

    def _custom_xlsx(self, path, title, sheets):
        import xlsxwriter
        wb = xlsxwriter.Workbook(path)
        F = lambda **k: wb.add_format({'font_name': 'Calibri', 'font_size': 10, 'valign': 'vcenter', **k})
        ban = F(bold=True, font_size=16, font_color='#FFFFFF', bg_color='#E40000', indent=1)
        hdr = F(bold=True, font_color='#FFFFFF', bg_color='#2B2F3A', border=1, border_color='#2B2F3A', align='center', text_wrap=True)
        txt = F(border=1, border_color='#E6D9DB'); num = F(border=1, border_color='#E6D9DB', num_format='#,##0.##', align='center')
        used = set()
        for sh in sheets[:20]:
            nm = re.sub(r'[\[\]:*?/\\]', ' ', sh.get('name') or 'Sheet')[:31] or 'Sheet'
            while nm in used: nm = nm[:28] + str(len(used))
            used.add(nm)
            ws = wb.add_worksheet(nm); ws.hide_gridlines(2)
            cols = sh['columns']; rows = [[_num(v) for v in r] for r in sh['rows']]
            ws.set_row(0, 30); ws.merge_range(0, 0, 0, max(1, len(cols) - 1), '  ' + title + (' — ' + sh['name'] if sh.get('name') else ''), ban)
            ws.write_row(2, 0, cols, hdr); ws.set_row(2, 28)
            for i, r in enumerate(rows):
                for j, v in enumerate(r[:len(cols)]):
                    if isinstance(v, (int, float)) and not isinstance(v, bool): ws.write_number(3 + i, j, v, num)
                    else: ws.write(3 + i, j, '' if v is None else str(v), txt)
            for j, c in enumerate(cols):
                w = max([len(str(c))] + [len(str(r[j])) for r in rows[:200] if j < len(r)]) + 3
                ws.set_column(j, j, min(max(w, 9), 45))
            ws.freeze_panes(3, 1)
            if rows: ws.autofilter(2, 0, 2 + len(rows), len(cols) - 1)
            chp = sh.get('chart')
            if chp and rows and len(cols) > 1:
                ch = wb.add_chart({'type': chp.get('type', 'column')})
                pal = ['#E40000', '#10B981', '#F59E0B', '#3B82F6', '#94A3B8', '#A855F7']
                for j in range(1, len(cols)):
                    if not all(isinstance(r[j], (int, float)) for r in rows if j < len(r) and r[j] is not None): continue
                    ser = {'name': [nm, 2, j], 'categories': [nm, 3, 0, 2 + len(rows), 0], 'values': [nm, 3, j, 2 + len(rows), j]}
                    if chp.get('type') != 'pie': ser['fill'] = {'color': pal[(j - 1) % len(pal)]}
                    ch.add_series(ser)
                ch.set_title({'name': chp.get('title') or sh['name']}); ch.set_size({'width': 720, 'height': 360})
                ws.insert_chart(2, len(cols) + 1, ch)
        wb.close()

    def _pptx(self, path, title, slides):
        from pptx import Presentation
        from pptx.util import Inches, Pt, Emu
        from pptx.dml.color import RGBColor
        from pptx.chart.data import CategoryChartData
        from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
        from pptx.enum.shapes import MSO_SHAPE
        RED = RGBColor(0xE4, 0x00, 0x00); INK = RGBColor(0x1A, 0x1D, 0x24); GREY = RGBColor(0x6B, 0x71, 0x80); WHITE = RGBColor(255, 255, 255)
        prs = Presentation(); prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
        blank = prs.slide_layouts[6]
        W, H = prs.slide_width, prs.slide_height

        def text(sl, x, y, w, h, s, size=18, bold=False, color=INK, align=None):
            tb = sl.shapes.add_textbox(x, y, w, h); tf = tb.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]; p.text = s; p.font.size = Pt(size); p.font.bold = bold; p.font.color.rgb = color; p.font.name = 'Calibri'
            if align: p.alignment = align
            return tf

        def frame(sl, t, n):
            bar = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, Inches(0.12)); bar.fill.solid(); bar.fill.fore_color.rgb = RED; bar.line.fill.background()
            text(sl, Inches(0.6), Inches(0.35), Inches(11.5), Inches(0.9), t, 30, True)
            text(sl, Inches(0.6), H - Inches(0.5), Inches(8), Inches(0.35), 'airtel · RFAI Tracker · ' + title, 10, color=GREY)
            text(sl, W - Inches(1.2), H - Inches(0.5), Inches(0.8), Inches(0.35), str(n), 10, color=GREY)

        CT = {'column': XL_CHART_TYPE.COLUMN_CLUSTERED, 'bar': XL_CHART_TYPE.BAR_CLUSTERED, 'line': XL_CHART_TYPE.LINE_MARKERS,
              'pie': XL_CHART_TYPE.PIE, 'stacked_column': XL_CHART_TYPE.COLUMN_STACKED}
        PALH = ['#E40000', '#10B981', '#F59E0B', '#3B82F6', '#94A3B8', '#A855F7']
        PAL = [RED, RGBColor(0x10, 0xB9, 0x81), RGBColor(0xF5, 0x9E, 0x0B), RGBColor(0x3B, 0x82, 0xF6), RGBColor(0x94, 0xA3, 0xB8), RGBColor(0xA8, 0x55, 0xF7)]
        for n, s in enumerate(slides[:40], 1):
            sl = prs.slides.add_slide(blank); k = s.get('kind', 'bullets')
            if k == 'title':
                bg = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H); bg.fill.solid(); bg.fill.fore_color.rgb = RED; bg.line.fill.background()
                text(sl, Inches(0.9), Inches(0.6), Inches(4), Inches(0.8), 'airtel', 34, True, WHITE)
                text(sl, Inches(0.9), Inches(2.6), Inches(11.5), Inches(1.6), s.get('title', title), 44, True, WHITE)
                text(sl, Inches(0.9), Inches(4.2), Inches(11.5), Inches(1), s.get('subtitle') or datetime.now().strftime('%d %B %Y'), 20, color=RGBColor(0xFF, 0xE3, 0xE3))
            else:
                frame(sl, s.get('title', ''), n)
                top = Inches(1.45); left = Inches(0.6); w = W - Inches(1.2); h = H - Inches(2.2)
                if k == 'bullets':
                    tf = text(sl, left, top, w, h, '', 20)
                    for i, bt in enumerate(s.get('bullets') or []):
                        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                        p.text = '•  ' + str(bt); p.font.size = Pt(20); p.font.color.rgb = INK; p.space_after = Pt(10)
                elif k == 'kpis':
                    ks = (s.get('kpis') or [])[:8]; cols = min(4, max(1, len(ks))); cw = (w - Inches(0.3) * (cols - 1)) / cols
                    for i, kp in enumerate(ks):
                        x = left + (cw + Inches(0.3)) * (i % cols); y = top + Inches(2.1) * (i // cols)
                        box = sl.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, int(x), int(y), int(cw), Inches(1.8))
                        box.fill.solid(); box.fill.fore_color.rgb = RGBColor(0xFF, 0xF5, 0xF5); box.line.color.rgb = RGBColor(0xF0, 0xD6, 0xD9)
                        text(sl, int(x) + Inches(0.25), int(y) + Inches(0.2), int(cw) - Inches(0.5), Inches(0.5), str(kp.get('label', '')).upper(), 12, True, GREY)
                        text(sl, int(x) + Inches(0.25), int(y) + Inches(0.7), int(cw) - Inches(0.5), Inches(0.9), str(kp.get('value', '')), 32, True, RED)
                elif k == 'table':
                    cols = s.get('columns') or []; rows = (s.get('rows') or [])[:18]
                    if cols:
                        tb = sl.shapes.add_table(len(rows) + 1, len(cols), left, top, w, Inches(0.4) * (len(rows) + 1)).table
                        for j, c in enumerate(cols):
                            cell = tb.cell(0, j); cell.text = str(c); cell.fill.solid(); cell.fill.fore_color.rgb = INK
                            pp = cell.text_frame.paragraphs[0]; pp.font.bold = True; pp.font.size = Pt(12); pp.font.color.rgb = WHITE
                        for i, r in enumerate(rows, 1):
                            for j in range(len(cols)):
                                cell = tb.cell(i, j); v = r[j] if j < len(r) else ''
                                cell.text = f'{v:,}' if isinstance(v, (int, float)) and not isinstance(v, bool) else str(v if v is not None else '')
                                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFA, 0xFA) if i % 2 else WHITE
                                pp = cell.text_frame.paragraphs[0]; pp.font.size = Pt(11); pp.font.color.rgb = INK
                elif k == 'chart':
                    cd = CategoryChartData(); cd.categories = [str(c) for c in (s.get('categories') or [])]
                    for se in s.get('series') or []: cd.add_series(str(se.get('name', '')), [float(v or 0) for v in se.get('values') or []])
                    ct = s.get('chart_type', 'column')
                    ch = sl.shapes.add_chart(CT.get(ct, XL_CHART_TYPE.COLUMN_CLUSTERED), left, top, w, h, cd).chart
                    ch.has_legend = len(s.get('series') or []) > 1 or ct == 'pie'
                    if ch.has_legend: ch.legend.position = XL_LEGEND_POSITION.BOTTOM; ch.legend.include_in_layout = False
                    plot = ch.plots[0]
                    if ct != 'line':
                        plot.has_data_labels = True; plot.data_labels.font.size = Pt(11)
                    if ct == 'pie':
                        for i, pt in enumerate(plot.series[0].points):
                            pt.format.fill.solid(); pt.format.fill.fore_color.rgb = RGBColor.from_string(_col(cd.categories[i].label, i, PALH)[1:])
                    else:
                        for i, se in enumerate(plot.series):
                            c_ = RGBColor.from_string(_col(se.name, i, PALH)[1:])
                            if ct == 'line': se.format.line.color.rgb = c_
                            else: se.format.fill.solid(); se.format.fill.fore_color.rgb = c_
            if s.get('notes'): sl.notes_slide.notes_text_frame.text = s['notes']
        prs.save(path)

    def _png(self, path, a):
        import matplotlib; matplotlib.use('Agg')
        import matplotlib.pyplot as plt
        pal = ['#E40000', '#10B981', '#F59E0B', '#3B82F6', '#94A3B8', '#A855F7']
        cats = [str(c) for c in a['categories']]; ser = a['series']; ct = a['chart_type']
        fig, ax = plt.subplots(figsize=(11, 5.5), dpi=150)
        if ct == 'pie':
            ax.pie(ser[0]['values'], labels=cats, autopct='%1.0f%%', colors=[_col(c, i, pal) for i, c in enumerate(cats)], startangle=90, wedgeprops={'width': 0.45, 'edgecolor': 'white'})
        else:
            x = np.arange(len(cats)); n = max(1, len(ser)); bw = 0.8 / n; bottom = np.zeros(len(cats))
            for i, s in enumerate(ser):
                v = np.array([float(t or 0) for t in s['values']])
                if ct == 'line': ax.plot(cats, v, marker='o', color=_col(s['name'], i, pal), label=s['name'], linewidth=2.2)
                elif ct == 'stacked_column': ax.bar(cats, v, bottom=bottom, color=_col(s['name'], i, pal), label=s['name']); bottom += v
                elif ct == 'bar': ax.barh(x + i * bw, v, bw, color=_col(s['name'], i, pal), label=s['name'])
                else:
                    b = ax.bar(x + i * bw - 0.4 + bw / 2, v, bw, color=_col(s['name'], i, pal), label=s['name']); ax.bar_label(b, fmt='%g', fontsize=8)
            if ct == 'bar': ax.set_yticks(x + bw * (n - 1) / 2, cats)
            elif ct == 'column': ax.set_xticks(x, cats)
            for sp in ('top', 'right'): ax.spines[sp].set_visible(False)
            ax.grid(axis='y' if ct != 'bar' else 'x', alpha=.25)
            if n > 1: ax.legend(frameon=False)
            plt.setp(ax.get_xticklabels(), rotation=30 if len(cats) > 8 else 0, ha='right' if len(cats) > 8 else 'center')
        ax.set_title(a['title'], fontsize=14, fontweight='bold', loc='left')
        fig.text(0.01, 0.01, 'airtel · RFAI Tracker', color='#E40000', fontsize=9, fontweight='bold')
        fig.tight_layout(); fig.savefig(path); plt.close(fig)
