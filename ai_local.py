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
    def answer(self, text, user, last=None):
        """Returns (answer_markdown, files, state_for_followups, suggestions)."""
        q = self.parse(text, last)
        if q.get('error'):
            return q['error'], [], None, []
        if q['help'] and not (q['filters'] or q['group_by'] or q['sr']):
            return HELP, [], last, ['Overall RFAI status', 'Circle-wise pending over 90 days', 'Make a PPT of RFAI status by circle']
        run = lambda name, a: self.ai._run_tool(name, a, user)
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
        ans = (f"**{what}** {scope}: **{total:,}** SRs in **{len(table)}** groups. "
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
            slides = [{'kind': 'title', 'title': title, 'subtitle': scope[0].upper() + scope[1:] + ' · ' + date.today().strftime('%d %B %Y')},
                      {'kind': 'kpis', 'title': 'At a glance', 'kpis': [{'label': 'SRs', 'value': f'{total:,}'},
                                                                         {'label': 'Groups', 'value': str(len(table))},
                                                                         {'label': 'Highest', 'value': str(lead[0])},
                                                                         {'label': 'Highest count', 'value': _fmt(lead_n)}]},
                      {'kind': 'chart', 'title': what, 'chart_type': ctype if ctype != 'pie' else 'column', 'categories': cats[:20],
                       'series': [{'name': s['name'], 'values': s['values'][:20]} for s in series]}]
            for i in range(0, min(len(table), 54), 18):
                slides.append({'kind': 'table', 'title': what + (f' ({i // 18 + 1})' if len(table) > 18 else ''), 'columns': cols, 'rows': table[i:i + 18]})
            slides.append({'kind': 'bullets', 'title': 'Key points', 'bullets': self._points(table, cols, total, two, q)})
            files.append(run('make_powerpoint', {'title': title, 'slides': slides})['_file'])
        if files:
            ans = {'ppt': 'Your **PowerPoint** is ready.', 'excel': 'Your **Excel** file is ready.', 'chart': 'Your **chart** is ready.'}[q['out']] + ' ' + ans
        sug = [] if q['out'] != 'answer' else ['Make a PPT of this', 'Export this to Excel', 'Chart this']
        return ans, files, state, sug

    @staticmethod
    def _points(table, cols, total, two, q):
        n = lambda r: r[-1] if two else r[1]
        srt = sorted(table, key=n, reverse=True)
        pts = [f'{total:,} SRs in scope across {len(table)} {cols[0].lower()} groups.']
        if srt:
            top3 = srt[:3]
            pts.append('Highest: ' + ', '.join(f'{r[0]} ({_fmt(n(r))})' for r in top3) + '.')
            if total: pts.append(f'Top 3 make up {sum(n(r) for r in top3) / total:.0%} of the total.')
        if len(srt) > 3:
            pts.append('Lowest: ' + ', '.join(f'{r[0]} ({_fmt(n(r))})' for r in srt[-3:]) + '.')
        if not two and len(table[0]) > 2:
            ad = [r for r in table if isinstance(r[2], (int, float))]
            if ad:
                w = max(ad, key=lambda r: r[2]); pts.append(f'Longest average delay: {w[0]} at {w[2]} days.')
        return pts

    def _deck(self, f, scope, user, state):
        run = lambda name, a: self.ai._run_tool(name, a, user)
        tot = run('aggregate', {'filters': f})
        if not tot['total']: return f'No SRs match {scope}.', [], state, []
        st = tot['statuses']; T = tot['total']
        order = ['OnAir', 'OnAir Pending', 'Open PRI', 'RFAI Pending', 'Rejected RFAI']
        pend = sum(st.get(s, 0) for s in ('OnAir Pending', 'Open PRI', 'RFAI Pending'))
        title = self._title(f, 'RFAI status review')
        slides = [{'kind': 'title', 'title': title, 'subtitle': scope[0].upper() + scope[1:] + ' · ' + date.today().strftime('%d %B %Y')},
                  {'kind': 'kpis', 'title': 'At a glance', 'kpis': [
                      {'label': 'Total SRs', 'value': f'{T:,}'}, {'label': 'OnAir', 'value': f"{st.get('OnAir', 0):,}"},
                      {'label': 'Pending', 'value': f'{pend:,}'}, {'label': 'Open PRI', 'value': f"{st.get('Open PRI', 0):,}"},
                      {'label': 'Rejected RFAI', 'value': f"{st.get('Rejected RFAI', 0):,}"},
                      {'label': 'OnAir rate', 'value': f"{st.get('OnAir', 0) / T:.0%}"},
                      {'label': 'Avg days', 'value': _fmt(tot.get('avg_days'))}]},
                  {'kind': 'chart', 'title': 'Status mix', 'chart_type': 'pie', 'categories': [s for s in order if st.get(s)],
                   'series': [{'name': 'SRs', 'values': [st[s] for s in order if st.get(s)]}]}]
        bc = run('aggregate', {'group_by': ['circle', 'status'], 'filters': f})
        if bc.get('rows'):
            p = pd.DataFrame(bc['rows']).pivot_table(index='circle', columns='status', values='srs', aggfunc='sum', fill_value=0)
            p = p[[s for s in order if s in p.columns]]
            p = p.loc[p.sum(axis=1).sort_values(ascending=False).index]
            if len(p) > 1:
                top = p.head(15)
                slides.append({'kind': 'chart', 'title': 'Status by circle', 'chart_type': 'stacked_column', 'categories': [str(i) for i in top.index],
                               'series': [{'name': c, 'values': [int(v) for v in top[c]]} for c in top.columns]})
                cols = ['Circle'] + list(p.columns) + ['Total']
                rows = [[i] + [int(v) for v in r] + [int(r.sum())] for i, r in zip(p.index, p.values)]
                for i in range(0, min(len(rows), 36), 18):
                    slides.append({'kind': 'table', 'title': 'Circle-wise status' + (f' ({i // 18 + 1})' if len(rows) > 18 else ''), 'columns': cols, 'rows': rows[i:i + 18]})
        ff = dict(f); ff['status'] = 'pending'
        ag = run('aggregate', {'group_by': ['ageing_group'], 'filters': ff})
        if ag.get('rows'):
            rr = sorted(ag['rows'], key=lambda r: str(r['ageing_group']))
            slides.append({'kind': 'chart', 'title': 'Pending SRs by ageing', 'chart_type': 'column', 'categories': [str(r['ageing_group']) for r in rr],
                           'series': [{'name': 'Pending SRs', 'values': [r['srs'] for r in rr]}]})
        old = run('list_srs', {'filters': ff, 'sort_by': 'days', 'descending': True, 'limit': 15})
        if old.get('rows'):
            slides.append({'kind': 'table', 'title': 'Oldest pending SRs', 'columns': ['SR', 'Circle', 'Status', 'Days', 'Final RFAI'],
                           'rows': [[x.get('srNo'), x.get('circle'), x.get('srStatus'), x.get('delay'), x.get('finalRfai')] for x in old['rows']]})
        pts = [f'{T:,} SRs in scope; {st.get("OnAir", 0):,} OnAir ({st.get("OnAir", 0) / T:.0%}) and {pend:,} still pending.']
        if st.get('Open PRI'): pts.append(f"{st['Open PRI']:,} SRs are held by an open PRI.")
        if st.get('Rejected RFAI'): pts.append(f"{st['Rejected RFAI']:,} RFAI offers were rejected.")
        if bc.get('rows') and len(p) > 1:
            pc = [c for c in ('OnAir Pending', 'Open PRI', 'RFAI Pending') if c in p.columns]
            if pc:
                ps = p[pc].sum(axis=1).sort_values(ascending=False)
                pts.append('Most pending: ' + ', '.join(f'{i} ({int(v):,})' for i, v in ps.head(3).items()) + '.')
        slides.append({'kind': 'bullets', 'title': 'Key points', 'bullets': pts})
        r = run('make_powerpoint', {'title': title, 'slides': slides})
        return (f"Your PowerPoint is ready: **{len(slides)} slides** for SRs {scope}: KPIs, status mix, circle-wise status, "
                f"ageing and the oldest pending SRs."), [r['_file']], state, ['Export this to Excel']

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
