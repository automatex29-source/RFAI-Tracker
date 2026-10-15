"""
Built-in RFAI assistant that needs no API key and no internet.

It reads the question with simple language rules (circle names, statuses, ageing, dates,
"by circle", "PPT", "Excel", "chart" …), runs the same tools as the AI engines against the
loaded data, and writes a short answer with a table. Nothing leaves the server.
"""
import re
from datetime import date, timedelta

import pandas as pd

GROUP_WORDS = [  # (regex, group_by key)
    (r'circle', 'circle'), (r'status', 'status'), (r'oem|vendor', 'oem'), (r'toco\s*project|categor', 'category'),
    (r'project', 'project'), (r'toco', 'toco'), (r'partner', 'partner'), (r'tsp', 'tsp'), (r'domain|project type', 'domain'),
    (r'age?ing|aging|bucket', 'ageing_group'), (r'month', 'exit_month'),
]
ENGLISH = {'up', 'ap', 'or', 'in', 'on', 'of', 'to', 'be', 'me', 'my', 'is', 'it', 'at', 'as', 'an', 'no', 'so', 'do', 'go',
           'we', 'us', 'all', 'and', 'the', 'for', 'per', 'pri', 'top', 'new', 'old', 'one', 'two', 'any'}
MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
LABEL = {'circle': 'Circle', 'status': 'Status', 'oem': 'OEM', 'category': 'TOCO project', 'project': 'Project',
         'toco': 'ToCo', 'partner': 'Partner', 'tsp': 'TSP', 'domain': 'Project type', 'ageing_group': 'Ageing',
         'exit_month': 'Month', 'rfai_month': 'RFAI month'}
EVENT_LABEL = {'initial_rfai': 'RFAI offered', 'final_rfai': 'final RFAI', 'first_pri_start': 'PRI opened',
               'last_pri_close': 'PRI closed', 'onair': 'OnAir', 'rfai_rejection': 'RFAI rejected'}
HELP = ("I'm running in **built-in mode** (free, works offline, your data stays on this server). "
        "I understand questions like:\n"
        "- *How many pending SRs in MH?*\n- *Circle-wise open PRI*\n- *PRIs opened last week by circle*\n"
        "- *Top 10 oldest pending SRs in DL*\n- *Details of SR 1-12345*\n- *Make a PPT of RFAI status by circle*\n"
        "- *Export DL pending over 90 days to Excel*\n- *Chart OnAir per month*")


def _fmt(n):
    try:
        f = float(n)
    except (TypeError, ValueError):
        return '' if n is None else str(n)
    return f'{int(f):,}' if f == int(f) else f'{f:,.1f}'


def _md_table(cols, rows):
    out = ['| ' + ' | '.join(cols) + ' |', '|' + '---|' * len(cols)]
    for r in rows:
        out.append('| ' + ' | '.join(_fmt(v) if isinstance(v, (int, float)) else str(v if v is not None else '').replace('|', '/') for v in r) + ' |')
    return '\n'.join(out)


class LocalBrain:
    def __init__(self, assistant):
        self.ai = assistant

    # ── understanding the question ─────────────────────────────────────────
    def parse(self, text, last=None):
        t = ' ' + re.sub(r'\s+', ' ', text.lower()) + ' '
        ov = self.ai._run_tool('get_overview', {}, '')
        if ov.get('error'):
            return {'error': ov['error']}
        q = {'filters': {}, 'group_by': [], 'out': 'answer', 'top': None, 'list': False, 'sr': None, 'date_group': None}
        f = q['filters']

        # follow-ups: "make this a PPT", "same in excel"
        tt = re.sub(r'\bthis (week|month|year)\b', '', t)
        followup = bool(re.search(r'\b(this|that|same|it|above|these|those)\b', tt)) and len(t.split()) <= 9
        if followup and last:
            q.update({k: (dict(v) if isinstance(v, dict) else list(v) if isinstance(v, list) else v) for k, v in last.items()})
            q['out'] = 'answer'

        # one SR
        df = self.ai.get_df()
        srs = set(df['srNo'].astype(str).str.upper())
        for tok in re.findall(r'[A-Za-z0-9][A-Za-z0-9_\-/]{2,}', text):
            if tok.upper() in srs and any(c.isdigit() for c in tok):
                q['sr'] = tok
                break

        # circles
        cs = []
        for c in ov['circles']:
            cl = c.lower()
            m = re.search(r'(?<![a-z0-9])' + re.escape(cl) + r'(?![a-z0-9])', t)
            if not m:
                continue
            if cl in ENGLISH or len(cl) <= 2:
                if not re.search(r'(?<![A-Za-z0-9])' + re.escape(c.upper()) + r'(?![A-Za-z0-9])', text) and \
                   not re.search(r'circle\s+' + re.escape(cl) + r'\b|\b' + re.escape(cl) + r'\s+circle', t):
                    continue
            cs.append(c)
        if cs:
            f['circle'] = ','.join(cs)
        for key, src in (('oem', 'oems'), ('toco', 'tocos'), ('category', 'categories'), ('project', 'projects')):
            hit = [v for v in ov.get(src, []) if len(v) >= 3 and v.lower() not in ENGLISH and
                   re.search(r'(?<![a-z0-9])' + re.escape(v.lower()) + r'(?![a-z0-9])', t)]
            if hit and not (key == 'project' and f.get('category')):
                f[key] = ','.join(hit)

        # event + dates
        ev = None
        if re.search(r'\bpris?\b|post.?rfai', t) and re.search(r'open(ed)?|raise|start|new', t) and not re.search(r'still open|open pri ', t):
            ev = 'first_pri_start'
        elif re.search(r'\bpris?\b', t) and re.search(r'clos|resolv|clear', t):
            ev = 'last_pri_close'
        timeish = r'week|month|today|yesterday|last \d+ days|since|between|trend|per (day|month)|daily|monthly|over time'
        if ev: pass
        elif re.search(r'reject', t) and re.search(timeish, t):
            ev = 'rfai_rejection'
        elif re.search(r'on ?air|live|launched', t) and re.search(timeish, t):
            ev = 'onair'
        elif re.search(r'final rfai', t) and re.search(timeish, t):
            ev = 'final_rfai'
        elif re.search(r'rfai|offer', t) and re.search(timeish, t) and not re.search(r'pending|reject', t):
            ev = 'initial_rfai'
        d0, d1 = self._dates(t)
        if (d0 or d1) and not ev:
            ev = 'initial_rfai'
        if ev and (d0 or d1):
            f.update({'date_field': ev, 'date_from': str(d0 or date(2000, 1, 1)), 'date_to': str(d1 or date.today())})
        if ev or not (followup and last): q['event'] = ev

        # status (only when the question is not about an event in time)
        if ev not in ('first_pri_start', 'last_pri_close', 'onair', 'rfai_rejection'):
            st = []
            if re.search(r'on ?air pending|pending on ?air|oa pending', t): st.append('OnAir Pending')
            elif re.search(r'rfai pending', t): st.append('RFAI Pending')
            elif re.search(r'\bpending\b|not (yet )?on ?air|open sites|backlog', t): st.append('pending')
            if re.search(r'open pri|\bpris?\b|post.?rfai', t) and 'pending' not in st: st.append('Open PRI')
            if re.search(r'reject', t): st.append('Rejected RFAI')
            if re.search(r'\bon ?air(ed)?\b|\blive\b|completed', t) and not st: st.append('OnAir')
            if st:
                f['status'] = ','.join(st)

        # ageing
        m = re.search(r'(?:over|above|more than|>|older than|beyond|greater than)\s*(\d+)', t)
        if m:
            n = int(m.group(1)); f['aging'] = '>90' if n >= 90 else ('>30' if n >= 30 else '>30')
        elif re.search(r'60\s*-\s*90', t): f['aging'] = '60-90'
        elif re.search(r'30\s*-\s*60', t): f['aging'] = '30-60'
        elif re.search(r'0\s*-\s*30', t): f['aging'] = '0-30'

        # exit months like "Jan'26", "january 2026", "in feb"
        mm = []
        for mk in ov.get('exit_months', []):
            mt = re.match(r"([A-Za-z]{3})'?(\d{2})", mk)
            if not mt: continue
            mo, yy = mt.group(1).lower(), mt.group(2)
            if re.search(r'\b' + mo + r"[a-z]*\.?\s*'?(20)?" + yy + r'\b', t):
                mm.append(mk)
        if not mm and not (d0 or d1):
            for mo in MONTHS:
                if re.search(r'\b(in|for|of|during)\s+' + mo + r'[a-z]*\b', t):
                    mm += [mk for mk in ov.get('exit_months', []) if mk.lower().startswith(mo)][-1:]
        if mm:
            f['month'] = ','.join(mm)

        # grouping
        gb = []
        for rx, key in GROUP_WORDS:
            if re.search(r'(by|per|each|every|wise|across|split|breakdown of|which|what)\s+(\w+\s)?(' + rx + r')|(' + rx + r')[\s-]*wise', t) and key not in gb:
                if key == 'project' and 'category' in gb: continue
                gb.append(key)
        if re.search(r'\bmost|\bhighest|\bworst|\bbest|\blowest|\bleast|\bcompare|\branking|\brank\b', t) and not gb and not q['sr']:
            gb = ['circle']
        if re.search(r'\bdaily|per day|day.?wise|each day', t): q['date_group'] = 'day'
        elif re.search(r'\bweekly|per week|week.?wise|each week', t): q['date_group'] = 'week'
        elif re.search(r'\bmonthly|per month|month.?wise|each month|trend|over time', t):
            q['date_group'] = 'month' if ev else None
            if not ev and 'exit_month' not in gb: gb.append('exit_month')
        if 'exit_month' in gb and q['date_group']:
            gb.remove('exit_month')
        if gb and 'status' not in gb and len(gb) == 1 and not f.get('status') and not ev and re.search(r'status|split|mix', t):
            gb.append('status')
        q['group_by'] = gb[:2] if gb else q['group_by']

        mt = re.search(r'\btop\s*(\d+)|\b(\d+)\s*(oldest|worst|latest|newest)', t)
        if mt: q['top'] = int(mt.group(1) or mt.group(2))
        counting = bool(re.search(r'how many|count|number of|\btotal', t))
        if not counting and re.search(r'\blist\b|\bshow (me )?(the )?(all )?srs?\b|which srs|\boldest\b|\bworst srs?\b|\bsr list|\bsrs? (with|in|of|where)', t) and not gb:
            q['list'] = True

        # output
        if re.search(r'\bppt|power ?point|presentation|\bdeck\b|slides?\b', t): q['out'] = 'ppt'
        elif re.search(r'excel|xlsx|spreadsheet|\bexport|\bdownload|\bsheet\b|workbook', t): q['out'] = 'excel'
        elif re.search(r'\bchart|\bgraph|\bplot|visuali[sz]e|\bpie\b|\bbar\b', t): q['out'] = 'chart'
        q['pie'] = bool(re.search(r'\bpie|donut|share', t))
        q['help'] = bool(re.search(r'^\s*(hi|hello|hey|help|what can you do|who are you)\b', t))
        q['overview'] = not (f or gb or q['sr'] or q['list'] or q['date_group']) and \
            bool(re.search(r'summary|overview|status|overall|total|how many|count|kpi|dashboard|report|how are we|situation', t) or q['out'] != 'answer')
        return q

    @staticmethod
    def _dates(t):
        today = date.today()
        if 'yesterday' in t: return today - timedelta(1), today - timedelta(1)
        if 'today' in t: return today, today
        m = re.search(r'last\s+(\d+)\s+(day|week|month)s?', t)
        if m:
            n = int(m.group(1)) * {'day': 1, 'week': 7, 'month': 30}[m.group(2)]
            return today - timedelta(n - 1), today
        mon = today - timedelta(today.weekday())
        if 'this week' in t: return mon, today
        if re.search(r'last week|previous week|past week', t): return mon - timedelta(7), mon - timedelta(1)
        if 'this month' in t: return today.replace(day=1), today
        if re.search(r'last month|previous month|past month', t):
            e = today.replace(day=1) - timedelta(1); return e.replace(day=1), e
        if 'this year' in t: return today.replace(month=1, day=1), today
        m = re.search(r'(\d{4}-\d{2}-\d{2})\s*(?:to|-|until|till|and)\s*(\d{4}-\d{2}-\d{2})', t)
        if m: return date.fromisoformat(m.group(1)), date.fromisoformat(m.group(2))
        m = re.search(r'since\s+(\d{4}-\d{2}-\d{2})', t)
        if m: return date.fromisoformat(m.group(1)), today
        return None, None

    # ── answering ──────────────────────────────────────────────────────────
    def answer(self, text, user, last=None, ppt=None):
        """Returns (answer_markdown, files, state_for_followups, suggestions)."""
        q = self.parse(text, last)
        if q.get('error'):
            return q['error'], [], None, []
        if q['help'] and not (q['filters'] or q['group_by'] or q['sr']):
            return HELP, [], last, ['Overall RFAI status', 'Circle-wise pending over 90 days', 'Make a PPT of RFAI status by circle']
        run = lambda name, a: self.ai._run_tool(name, a, user)
        self.ppt = ppt or {}
        self.length = self.ppt.get('length', 'standard')
        f = {k: v for k, v in q['filters'].items() if v}
        scope = self._scope(f, q.get('event'))
        files = []

        if q['sr']:
            r = run('get_sr', {'sr': q['sr']})
            if r.get('error'): return f"I couldn't find SR **{q['sr']}**.", [], last, []
            m = r['matches'][0]
            keys = [('SR', 'srNo'), ('Site ID', 'siteId'), ('Circle', 'circle'), ('Status', 'srStatus'), ('Days', 'delay'),
                    ('Project', 'bPrj'), ('TOCO project', 'category'), ('ToCo', 'toco'), ('OEM', 'oem'), ('TSP', 'tsp'),
                    ('Partner', 'partner'), ('Initial RFAI', 'rfaiNotice'), ('RFAI rejected', 'rfaiRej'), ('RFAI re-offer', 'rfaiReoffer'),
                    ('Final RFAI', 'finalRfai'), ('First PRI start', 'priFirst'), ('Last PRI close', 'priLast'), ('OnAir', 'onAir')]
            rows = [(k, m.get(c)) for k, c in keys if m.get(c) not in (None, '', 'nan', 'None', '-1')]
            ans = f"**{'' if str(m.get('srNo', '')).upper().startswith('SR') else 'SR '}{m.get('srNo')}** is **{m.get('srStatus')}**" + (f" in {m.get('circle')}" if m.get('circle') else '') + '.\n\n' + _md_table(['Field', 'Value'], rows)
            return ans, [], last, []

        state = {k: q[k] for k in ('filters', 'group_by', 'top', 'list', 'date_group', 'event')}
        sug = []

        if q['date_group'] and q.get('event'):
            agg = run('aggregate', {'date_group': {'field': q['event'], 'period': q['date_group']}, 'filters': f,
                                    'group_by': [g for g in q['group_by'] if g != 'exit_month'][:1]})
            return self._finish_agg(q, agg, scope, user, state, f"{EVENT_LABEL[q['event']]} per {q['date_group']}")

        if q['out'] == 'ppt' and set(q['group_by']) <= {'circle', 'status'}:
            return self._deck(f, scope, user, state)
        if q['group_by']:
            agg = run('aggregate', {'group_by': q['group_by'], 'filters': f, 'top': q['top']})
            return self._finish_agg(q, agg, scope, user, state, ' and '.join(LABEL[g] for g in q['group_by']) + '-wise')

        if q['list'] and q['out'] == 'answer':
            n = q['top'] or 15
            r = run('list_srs', {'filters': f, 'sort_by': 'days', 'descending': True, 'limit': n})
            if not r['total_matches']: return f'No SRs match {scope}.', [], state, []
            rows = [(x.get('srNo'), x.get('circle'), x.get('srStatus'), x.get('delay'), x.get('finalRfai')) for x in r['rows']]
            ans = (f"**{r['total_matches']:,}** SRs match {scope}. Here are the {r['shown']} with the most days:\n\n" +
                   _md_table(['SR', 'Circle', 'Status', 'Days', 'Final RFAI'], rows))
            return ans, [], state, ['Export these to Excel', 'Circle-wise count of these']

        # plain totals / overview, optionally as a file
        if q['out'] == 'excel':
            r = run('make_data_excel', {'filters': f, 'title': self._title(f, 'RFAI export')})
            if r.get('error'): return r['error'], [], state, []
            return (f"Your Excel file is ready with **{r['srs']:,}** SRs ({scope}): summary with charts, Journey Table, "
                    "and pending / PRI / rejected / OnAir sheets."), [r['_file']], state, ['Make a PPT of this']
        if q['out'] == 'ppt':
            return self._deck(f, scope, user, state)
        tot = run('aggregate', {'filters': f})
        if not tot['total']: return f'No SRs match {scope}.', [], state, []
        st = tot['statuses']
        ans = f"There are **{tot['total']:,}** SRs {scope}" + (f", average **{tot['avg_days']}** days." if tot.get('avg_days') is not None else '.')
        order = ['OnAir', 'OnAir Pending', 'Open PRI', 'RFAI Pending', 'Rejected RFAI']
        rows = [(s, st.get(s, 0), f"{st.get(s, 0) / tot['total']:.0%}") for s in order if st.get(s)]
        ans += '\n\n' + _md_table(['Status', 'SRs', 'Share'], rows)
        if q['out'] == 'chart':
            r = run('make_chart_image', {'title': 'RFAI status ' + scope, 'chart_type': 'pie',
                                         'categories': [r_[0] for r_ in rows], 'series': [{'name': 'SRs', 'values': [r_[1] for r_ in rows]}]})
            files.append(r['_file'])
        sug = ['Circle-wise breakdown', 'Make a PPT of this', 'Export this to Excel']
        return ans, files, state, sug

    def _finish_agg(self, q, agg, scope, user, state, what):
        run = lambda name, a: self.ai._run_tool(name, a, user)
        rows = agg.get('rows') or []
        if not rows:
            return f'No SRs match {scope}.', [], state, []
        keys = [k for k in rows[0] if k not in ('srs', 'avg_days', 'max_days')]
        two = len(keys) == 2
        if two:   # pivot second key into columns
            p = pd.DataFrame(rows).pivot_table(index=keys[0], columns=keys[1], values='srs', aggfunc='sum', fill_value=0)
            if not q['date_group']:
                p = p.loc[p.sum(axis=1).sort_values(ascending=False).index]
            p['Total'] = p.sum(axis=1)
            cols = [LABEL.get(keys[0], keys[0].replace('_', ' ').title())] + [str(c) for c in p.columns]
            table = [[i] + [int(v) for v in r] for i, r in zip(p.index, p.values)]
        else:
            cols = [LABEL.get(keys[0], keys[0].replace('_', ' ').title()), 'SRs', 'Avg days']
            table = [[r[keys[0]], int(r['srs']), r['avg_days'] if r['avg_days'] is not None else ''] for r in rows]
        total = sum(int(r['srs']) for r in rows)
        lead = table[0] if not q['date_group'] else max(table, key=lambda r: r[-1] if two else r[1])
        lead_n = lead[-1] if two else lead[1]
        ans = (f"**{what}** {scope}: **{total:,}** SRs in **{len(table)}** group{'s' if len(table) != 1 else ''}. "
               f"Highest is **{lead[0]}** with **{_fmt(lead_n)}**" + (f" ({lead_n / total:.0%})." if total else '.'))
        shown = table[:25]
        ans += '\n\n' + _md_table(cols, shown) + (f'\n\n_Showing 25 of {len(table)}. Ask for Excel to get all._' if len(table) > 25 else '')
        files = []
        title = self._title(q['filters'], what)
        cats = [str(r[0]) for r in table[:30]]
        if two:
            series = [{'name': c, 'values': [r[j + 1] for r in table[:30]]} for j, c in enumerate(cols[1:-1])]
        else:
            series = [{'name': 'SRs', 'values': [r[1] for r in table[:30]]}]
        ctype = 'line' if q['date_group'] else ('stacked_column' if two else ('pie' if q.get('pie') else 'column'))
        if q['out'] == 'chart':
            files.append(run('make_chart_image', {'title': title, 'chart_type': ctype, 'categories': cats, 'series': series})['_file'])
        elif q['out'] == 'excel':
            files.append(run('make_custom_excel', {'title': title, 'sheets': [{'name': what[:31], 'columns': cols, 'rows': table,
                                                    'chart': {'type': 'line' if q['date_group'] else 'column', 'title': title}}]})['_file'])
            if not q['date_group']:
                r = run('make_data_excel', {'filters': q['filters'], 'title': title + ' SR list'})
                if not r.get('error'): files.append(r['_file'])
        elif q['out'] == 'ppt':
            sc = scope[0].upper() + scope[1:]
            slides = [{'kind': 'title', 'title': title, 'subtitle': f'{what} view of SRs {scope}',
                       'kpis': [{'label': 'SRs', 'value': f'{total:,}'}, {'label': 'Highest', 'value': str(lead[0])}, {'label': 'Groups', 'value': str(len(table))}]},
                      {'kind': 'summary', 'title': f'{lead[0]} leads with {_fmt(lead_n)} SRs', 'subtitle': f'{what} view, {scope}',
                       'stats': [{'value': f'{total:,}', 'label': 'SRs in scope'},
                                 {'value': str(lead[0]), 'label': f'highest, with {_fmt(lead_n)} SRs'},
                                 {'value': f'{lead_n / total:.0%}' if total else '-', 'label': f'of SRs are in {lead[0]}'}],
                       'points': [{'head': x.split(':')[0] + '.' if ':' in x else '', 'text': x.split(':', 1)[1].strip() if ':' in x else x}
                                  for x in self._points(table, cols, total, two, q)[:3]]},
                      {'kind': 'chart', 'title': f'{what}: {lead[0]} is highest' if not q['date_group'] else what,
                       'subtitle': f'Number of SRs, {scope}', 'chart_type': ctype, 'categories': cats[:20],
                       'series': [{'name': s['name'], 'values': s['values'][:20]} for s in series]}]
            trows = [list(r) for r in table]
            if two: trows.append(['Total'] + [int(sum(r[j] for r in table)) for j in range(1, len(cols))])
            elif not q['date_group']: trows.append(['Total', total, ''])
            if self.length != 'short':
                slides.append({'kind': 'table', 'title': what, 'subtitle': f'Number of SRs, {scope}', 'columns': cols, 'rows': trows})
            srt = sorted(table, key=lambda r: r[-1] if two else r[1], reverse=True)
            nx = [f'Review {srt[0][0]} first: it has the most SRs ({_fmt(srt[0][-1] if two else srt[0][1])}).']
            if len(srt) > 1: nx.append('Compare ' + ', '.join(str(r[0]) for r in srt[1:3]) + ' against the leader and agree targets.')
            if q['filters'].get('status') in ('pending', 'OnAir Pending', 'Open PRI'): nx.append('Agree OnAir dates for the oldest pending SRs.')
            nx.append('Track this view weekly to see the trend.')
            slides.append({'kind': 'closing', 'title': 'Next steps', 'subtitle': f'Suggested actions for the {what.lower()} view', 'bullets': nx})
            files.append(run('make_powerpoint', {'title': title, 'slides': slides})['_file'])
        if files:
            ans = {'ppt': 'Your **PowerPoint** is ready.', 'excel': 'Your **Excel** file is ready.', 'chart': 'Your **chart** is ready.'}[q['out']] + ' ' + ans
        sug = [] if q['out'] != 'answer' else ['Make a PPT of this', 'Export this to Excel', 'Chart this']
        return ans, files, state, sug

    @staticmethod
    def _points(table, cols, total, two, q):
        n = lambda r: r[-1] if two else r[1]
        srt = sorted(table, key=n, reverse=True)
        g = len(table)
        pts = [f'Scope: {total:,} SRs across {g} {cols[0].lower()} group{"s" if g != 1 else ""}.']
        if srt:
            top3 = srt[:3]
            pts.append('Highest: ' + ', '.join(f'{r[0]} ({_fmt(n(r))})' for r in top3) + '.')
            if total and len(srt) > 3: pts.append(f'Top 3 make up {sum(n(r) for r in top3) / total:.0%} of the total.')
        if len(srt) > 3:
            pts.append('Lowest: ' + ', '.join(f'{r[0]} ({_fmt(n(r))})' for r in srt[-3:]) + '.')
        if not two and len(table[0]) > 2:
            ad = [r for r in table if isinstance(r[2], (int, float))]
            if ad:
                w = max(ad, key=lambda r: r[2]); pts.append(f'Delay: {w[0]} has the longest average, {w[2]} days.')
        return pts

    def _deck(self, f, scope, user, state):
        run = lambda name, a: self.ai._run_tool(name, a, user)
        tot = run('aggregate', {'filters': f})
        if not tot['total']: return f'No SRs match {scope}.', [], state, []
        st = tot['statuses']; T = tot['total']
        order = ['OnAir', 'OnAir Pending', 'Open PRI', 'RFAI Pending', 'Rejected RFAI']
        pend = sum(st.get(s, 0) for s in ('OnAir Pending', 'Open PRI', 'RFAI Pending'))
        oa = st.get('OnAir', 0); pct = lambda n: f'{n / T:.0%} of total'
        title = self._title(f, 'RFAI status review')
        sc = scope[0].upper() + scope[1:]
        slides = [{'kind': 'title', 'title': title, 'subtitle': f'{sc}: RFAI to OnAir progress, open PRI and ageing',
                   'kpis': [{'label': 'Total SRs', 'value': f'{T:,}'}, {'label': 'OnAir', 'value': f'{oa / T:.0%}'},
                            {'label': 'Pending', 'value': f'{pend:,}'}]},
                  {'kind': 'kpis', 'title': f'{oa / T:.0%} of SRs are OnAir; {pend:,} still pending',
                   'subtitle': f'Snapshot of {T:,} SRs {scope}', 'kpis': [
                      {'label': 'Total SRs', 'value': f'{T:,}', 'sub': 'in scope'},
                      {'label': 'OnAir', 'value': f'{oa:,}', 'sub': pct(oa)},
                      {'label': 'OnAir Pending', 'value': f"{st.get('OnAir Pending', 0):,}", 'sub': pct(st.get('OnAir Pending', 0))},
                      {'label': 'Open PRI', 'value': f"{st.get('Open PRI', 0):,}", 'sub': pct(st.get('Open PRI', 0))},
                      {'label': 'RFAI Pending', 'value': f"{st.get('RFAI Pending', 0):,}", 'sub': pct(st.get('RFAI Pending', 0))},
                      {'label': 'Rejected RFAI', 'value': f"{st.get('Rejected RFAI', 0):,}", 'sub': pct(st.get('Rejected RFAI', 0))},
                      {'label': 'Avg days', 'value': _fmt(tot.get('avg_days')), 'sub': 'final RFAI to OnAir / today'},
                      {'label': 'Pending share', 'value': f'{pend / T:.0%}', 'sub': f'{pend:,} SRs not yet OnAir'}]}]
        summary_at = 0                        # summary goes right after the cover, filled in below
        mix = [s for s in order if st.get(s)]
        if len(mix) > 1:
            slides.append({'kind': 'chart', 'title': f'{max(mix, key=lambda m: st[m])} is the largest status group ({st[max(mix, key=lambda m: st[m])] / T:.0%})',
                           'subtitle': 'Share of SRs by current status', 'chart_type': 'pie', 'categories': mix,
                           'series': [{'name': 'SRs', 'values': [st[s] for s in mix]}],
                           'insights': [f'{s}: {st[s]:,} ({st[s] / T:.0%})' for s in sorted(mix, key=lambda m: -st[m])][:4]})
        bc = run('aggregate', {'group_by': ['circle', 'status'], 'filters': f}); p = None
        if bc.get('rows'):
            p = pd.DataFrame(bc['rows']).pivot_table(index='circle', columns='status', values='srs', aggfunc='sum', fill_value=0)
            p = p[[s for s in order if s in p.columns]]
            p = p.loc[p.sum(axis=1).sort_values(ascending=False).index]
            if len(p) > 1:
                pc = [c for c in ('OnAir Pending', 'Open PRI', 'RFAI Pending') if c in p.columns]
                ps = p[pc].sum(axis=1).sort_values(ascending=False) if pc else None
                top = p.head(15)
                ttl = (f'{ps.index[0]} has the most pending SRs ({int(ps.iloc[0]):,})' if ps is not None and len(ps) else 'Status by circle')
                slides.append({'kind': 'chart', 'title': ttl, 'subtitle': 'SRs by circle and current status' + (' (top 15 circles)' if len(p) > 15 else ''),
                               'chart_type': 'stacked_column', 'categories': [str(i) for i in top.index],
                               'series': [{'name': c, 'values': [int(v) for v in top[c]]} for c in top.columns]})
                cols = ['Circle'] + list(p.columns) + ['Total', 'OnAir %']
                rows = [[i] + [int(v) for v in r] + [int(r.sum()), f"{(r['OnAir'] if 'OnAir' in p.columns else 0) / max(r.sum(), 1):.0%}"] for i, r in zip(p.index, p.values) for r in [pd.Series(r, index=p.columns)]]
                tr = p.sum(axis=0)
                rows.append(['Total'] + [int(v) for v in tr] + [int(tr.sum()), f"{(tr['OnAir'] if 'OnAir' in p.columns else 0) / max(tr.sum(), 1):.0%}"])
                slides.append({'kind': 'table', 'title': 'Circle-wise status', 'subtitle': 'Number of SRs per circle by current status', 'columns': cols, 'rows': rows})
        tr_ = run('aggregate', {'date_group': {'field': 'onair', 'period': 'month'}, 'filters': f})
        if len(tr_.get('rows') or []) > 1:
            rr = tr_['rows'][-12:]
            best = max(rr, key=lambda r: r['srs'])
            slides.append({'kind': 'chart', 'title': f"OnAir peaked in {pd.Timestamp(best['onair_month'] + '-01').strftime('%b %Y')} with {best['srs']:,} sites",
                           'subtitle': 'Sites going OnAir per month (last 12 months)', 'chart_type': 'line',
                           'categories': [pd.Timestamp(r['onair_month'] + '-01').strftime("%b'%y") for r in rr],
                           'series': [{'name': 'OnAir', 'values': [r['srs'] for r in rr]}]})
        ff = dict(f); ff['status'] = 'pending'
        ag = run('aggregate', {'group_by': ['ageing_group'], 'filters': ff})
        ago = {'0-30': 0, '30-60': 1, '60-90': 2, '>90': 3}
        rr = sorted(ag.get('rows') or [], key=lambda r: ago.get(str(r['ageing_group']), 9))
        old_ = sum(r['srs'] for r in rr if r['ageing_group'] == '>90')
        if len(rr) == 1 and pend:
            g0 = str(rr[0]['ageing_group'])
            slides.append({'kind': 'stat', 'title': f'Every pending SR is in the {g0} days group', 'subtitle': 'Pending SRs by days since final RFAI (PRI days excluded)',
                           'value': f"{rr[0]['srs']:,}", 'label': f'pending SRs aged {g0} days', 'icon': 'hourglass',
                           'text': [f"That is 100% of the {pend:,} SRs not yet OnAir.",
                                    f"Average age of pending SRs is {_fmt(rr[0].get('avg_days'))} days." if rr[0].get('avg_days') is not None else 'Ageing is measured from final RFAI.',
                                    'Prioritise the oldest SRs (next slide) with circle teams.']})
        if len(rr) > 1:
            slides.append({'kind': 'chart', 'title': f'{old_ / max(pend, 1):.0%} of pending SRs are older than 90 days',
                           'subtitle': 'Pending SRs by days since final RFAI (PRI days excluded)', 'chart_type': 'column',
                           'categories': [str(r['ageing_group']) + ' days' for r in rr], 'series': [{'name': 'Pending SRs', 'values': [r['srs'] for r in rr]}]})
        old = run('list_srs', {'filters': ff, 'sort_by': 'days', 'descending': True, 'limit': 12})
        if old.get('rows'):
            slides.append({'kind': 'table', 'title': 'Oldest pending SRs need follow-up', 'subtitle': 'Top 12 pending SRs by days since final RFAI',
                           'columns': ['SR', 'Site ID', 'Circle', 'Status', 'Days', 'Final RFAI', 'Project'],
                           'rows': [[x.get('srNo'), x.get('siteId'), x.get('circle'), x.get('srStatus'), x.get('delay'), x.get('finalRfai'), x.get('bPrj')] for x in old['rows']]})
        summ = [{'head': 'Progress.', 'text': f'{oa:,} of {T:,} SRs are OnAir ({oa / T:.0%}); {pend:,} are still pending.', 'icon': 'check'}]
        if p is not None and len(p) > 1 and ps is not None and len(ps):
            summ.append({'head': 'Where.', 'text': f'{ps.index[0]} has the most pending SRs ({int(ps.iloc[0]):,}), then ' +
                         ', '.join(f'{i} ({int(v):,})' for i, v in ps.iloc[1:3].items()) + '.', 'icon': 'pin'})
        if rr:
            summ.append({'head': 'Ageing.', 'text': f'{old_ / max(pend, 1):.0%} of pending SRs are older than 90 days' +
                         (f"; {st['Open PRI']:,} are held by an open PRI." if st.get('Open PRI') else '.'), 'icon': 'hourglass'})
        slides.insert(summary_at + 1, {'kind': 'summary', 'title': 'Executive summary', 'subtitle': f'RFAI to OnAir, {scope}',
                                       'stats': [{'value': f'{oa / T:.0%}', 'label': 'of SRs are OnAir'},
                                                 {'value': f'{pend:,}', 'label': 'SRs still pending'},
                                                 {'value': f'{old_ / max(pend, 1):.0%}' if rr else _fmt(tot.get('avg_days')),
                                                  'label': 'of pending are over 90 days' if rr else 'average days'}],
                                       'points': summ,
                                       'notes': ' '.join(x['head'] + ' ' + x['text'] for x in summ)})
        pts = [f'{T:,} SRs in scope: {oa:,} OnAir ({oa / T:.0%}) and {pend:,} still pending ({pend / T:.0%}).']
        if p is not None and len(p) > 1 and ps is not None and len(ps):
            pts.append('Most pending: ' + ', '.join(f'{i} ({int(v):,})' for i, v in ps.head(3).items()) + '.')
        if st.get('Open PRI'): pts.append(f"{st['Open PRI']:,} SRs are held by an open PRI.")
        if st.get('Rejected RFAI'): pts.append(f"{st['Rejected RFAI']:,} RFAI offers were rejected and need a re-offer.")
        if tot.get('avg_days') is not None: pts.append(f"Average {_fmt(tot['avg_days'])} days from final RFAI to OnAir (or today if pending).")
        nxt = []
        if p is not None and len(p) > 1 and ps is not None and len(ps): nxt.append(f'Review pending SRs in {", ".join(list(ps.index[:2]))} with circle teams.')
        if old_: nxt.append(f'Escalate the {old_:,} SRs pending over 90 days and agree OnAir dates.')
        if st.get('Open PRI'): nxt.append('Close open PRIs so blocked sites can go OnAir.')
        if st.get('Rejected RFAI'): nxt.append('Follow up re-offers for rejected RFAIs.')
        nxt.append('Go through the oldest SRs in this deck site by site and set owners.')
        nxt.append('Track the weekly OnAir run-rate against plan.')
        if self.length == 'detailed':
            slides += self._deep_dive(f, ff, run)
        slides.append({'kind': 'closing', 'title': 'Next steps', 'subtitle': 'Suggested actions from this review', 'bullets': nxt})
        if self.length == 'short':
            keep = {'title', 'summary', 'closing'}
            first_chart = next((x for x in slides if x['kind'] == 'chart' and x.get('chart_type') == 'stacked_column'), None) or \
                next((x for x in slides if x['kind'] == 'chart'), None)
            first_table = next((x for x in slides if x['kind'] == 'table' and 'Oldest' in x['title']), None)
            slides = [x for x in slides if x['kind'] in keep or x is first_chart or x is first_table]
        for x in slides:
            if x['kind'] == 'chart' and 'notes' not in x:
                from ppt_builder import _auto_insights
                x['notes'] = x['title'] + '. ' + ' '.join(x.get('insights') or _auto_insights(x['categories'], x['series'], x['chart_type']))
        r = run('make_powerpoint', {'title': title, 'slides': slides})
        from ai_assistant import THEME_NAMES
        o = self.ppt
        n = sum(1 + (max(0, (len(x.get('rows') or []) - 1) // 12) if x['kind'] == 'table' else 0) for x in slides)
        pics = {'ai': ', AI pictures', 'art': ', built-in artwork'}.get(o.get('images'), '')
        if self.length == 'short':
            covers = 'Covers the summary, pending by circle, oldest SRs and next steps.'
        else:
            covers = ('Covers the executive summary, KPIs, status mix, circle-wise status, OnAir trend, ageing, oldest pending SRs'
                      + (', a deep dive by OEM, project and ageing per circle' if self.length == 'detailed' else '')
                      + ' and next steps, with speaker notes.')
        ans = (f"Your PowerPoint is ready: **{n} slides**, **{THEME_NAMES.get(o.get('theme', 'airtel'), 'Airtel Red')}** theme{pics}, "
               f"for SRs {scope}. " + covers)
        return ans, [r['_file']], state, ['Export this to Excel']

    def _deep_dive(self, f, ff, run):
        """Extra slides for a detailed deck."""
        out = [{'kind': 'section', 'title': 'Deep dive', 'subtitle': 'Pending SRs by OEM, project and circle ageing', 'icon': 'chart'}]
        for g, lab, ct in (('oem', 'OEM', 'bar'), ('project', 'project', 'column'), ('category', 'TOCO project', 'bar')):
            a = run('aggregate', {'group_by': [g], 'filters': ff, 'top': 12})
            rows = a.get('rows') or []
            if len(rows) < 2: continue
            out.append({'kind': 'chart', 'title': f"{rows[0][g]} has the most pending SRs by {lab} ({rows[0]['srs']:,})",
                        'subtitle': f'Pending SRs by {lab}' + (' (top 12)' if a.get('groups', 0) > 12 else ''), 'chart_type': ct,
                        'categories': [str(r[g]) for r in rows], 'series': [{'name': 'Pending SRs', 'values': [r['srs'] for r in rows]}]})
        a = run('aggregate', {'group_by': ['circle', 'ageing_group'], 'filters': ff})
        if a.get('rows'):
            p = pd.DataFrame(a['rows']).pivot_table(index='circle', columns='ageing_group', values='srs', aggfunc='sum', fill_value=0)
            p = p[[c for c in ('0-30', '30-60', '60-90', '>90') if c in p.columns]]
            p = p.loc[p.sum(axis=1).sort_values(ascending=False).index]
            if len(p) > 1 and len(p.columns) > 1:
                top = p.head(15)
                out.append({'kind': 'chart', 'title': 'Ageing of pending SRs by circle', 'subtitle': 'Days since final RFAI, PRI days excluded',
                            'chart_type': 'stacked_column', 'categories': [str(i) for i in top.index],
                            'series': [{'name': c + ' days', 'values': [int(v) for v in top[c]]} for c in top.columns]})
            if len(p) > 1:
                rows = [[i] + [int(v) for v in r] + [int(sum(r))] for i, r in zip(p.index, p.values)]
                tr = p.sum(axis=0); rows.append(['Total'] + [int(v) for v in tr] + [int(tr.sum())])
                out.append({'kind': 'table', 'title': 'Pending ageing by circle', 'subtitle': 'Number of pending SRs per ageing group',
                            'columns': ['Circle'] + [c + ' days' for c in p.columns] + ['Total'], 'rows': rows})
        return out

    @staticmethod
    def _title(f, base):
        bits = [f.get('circle', '').replace(',', ' ')] if f.get('circle') else []
        if f.get('status'): bits.append(f['status'].replace(',', ' ').replace('pending', 'Pending'))
        if f.get('aging'): bits.append(f['aging'] + ' days')
        if f.get('month'): bits.append(f['month'].replace(',', ' '))
        return (' '.join(bits) + ' ' + base).strip() if bits else base

    @staticmethod
    def _scope(f, ev):
        b = []
        if f.get('status'): b.append({'pending': 'pending'}.get(f['status'], f['status'].replace(',', ' / ')))
        s = ' '.join(b)
        p = []
        if f.get('circle'): p.append('in ' + f['circle'].replace(',', ', '))
        for k in ('oem', 'toco', 'category', 'project'):
            if f.get(k): p.append(f[k].replace(',', ', '))
        if f.get('aging'): p.append('over ' + f['aging'][1:] + ' days' if f['aging'].startswith('>') else f['aging'] + ' days old')
        if f.get('month'): p.append('for ' + f['month'].replace(',', ', '))
        if f.get('date_field'):
            a, z = f.get('date_from'), f.get('date_to')
            fm = lambda x: pd.Timestamp(x).strftime('%d %b %Y')
            p.append(f"with {EVENT_LABEL.get(f['date_field'], f['date_field'])} " + (f'on {fm(a)}' if a == z else f'from {fm(a)} to {fm(z)}'))
        out = ((s + ' ') if s else '') + ' '.join(p)
        return out.strip() or 'across all data'
