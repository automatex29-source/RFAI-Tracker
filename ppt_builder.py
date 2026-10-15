"""
Airtel-styled PowerPoint decks for the RFAI assistant.

build_deck(path, title, slides) takes a list of slide dicts:
  title    : title, subtitle, kpis (up to 3 shown on the cover)
  summary  : title, subtitle, stats [{value, label}] (up to 3 big numbers),
             points [{head, text, icon?}] (up to 3 findings)
  kpis     : title, subtitle, kpis [{label, value, sub?}]
  stat     : title, subtitle, value, label, text (one big number with context)
  chart    : title, subtitle, chart_type (column|bar|line|pie|stacked_column), categories,
             series [{name, values}], insights (auto-written when missing)
  table    : title, subtitle, columns, rows (long tables continue on more slides)
  bullets  : title, subtitle, bullets
  section  : title, subtitle
  closing  : title, subtitle, bullets (dark next-steps slide; added when missing)
Any slide may carry notes (speaker notes).
"""
from datetime import datetime
from pathlib import Path

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu

ASSETS = Path(__file__).with_name('assets')
FONT = 'Calibri'
RED = RGBColor(0xE4, 0x00, 0x00)
INK = RGBColor(0x16, 0x18, 0x1F)
INK2 = RGBColor(0x4B, 0x50, 0x5C)
MUTED = RGBColor(0x80, 0x86, 0x92)
LINE = RGBColor(0xE6, 0xE8, 0xEE)
SOFT = RGBColor(0xF5, 0xF6, 0xF9)
TINT = RGBColor(0xFF, 0xF0, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x17, 0x19, 0x21)
DARK2 = RGBColor(0x24, 0x27, 0x32)

STATUS = {'onair': '10B981', 'onair pending': 'F59E0B', 'pending': 'F59E0B', 'open pri': '3B82F6',
          'rejected rfai': 'EF4444', 'rejected': 'EF4444', 'rfai pending': '94A3B8'}
PAL = ['E40000', '1F2937', 'F59E0B', '3B82F6', '10B981', '8B5CF6', '94A3B8', 'EC4899']
ICON_FOR = [('open pri', 'alert'), ('pri', 'alert'), ('reject', 'ban'), ('rfai pending', 'clock'), ('onair pending', 'hourglass'),
            ('pending', 'hourglass'), ('onair', 'check'), ('on air', 'check'), ('day', 'calendar'), ('share', 'chart'),
            ('rate', 'trend'), ('circle', 'pin'), ('group', 'list'), ('highest', 'bullseye'), ('total', 'tower'), ('sr', 'tower')]
W, H = Inches(13.333), Inches(7.5)
M = Inches(0.6)
TOP = Inches(1.95)          # where content starts on every content slide


def _rgb(h): return RGBColor.from_string(h)


def _col(name, i):
    return STATUS.get(str(name).strip().lower(), PAL[i % len(PAL)])


def _kcol(label, i):
    """Colour for a KPI by its label (status colours), else red for the first, ink for the rest."""
    l = str(label).strip().lower()
    for k, v in STATUS.items():
        if l == k or l.startswith(k + ' '): return v
    if 'onair' in l and 'pending' not in l: return STATUS['onair']
    return 'E40000' if i == 0 else '1F2937'


def _icon(label, default='chart'):
    l = str(label).lower()
    for k, v in ICON_FOR:
        if k in l: return v
    return default


def _num(v):
    if isinstance(v, bool): return None
    if isinstance(v, (int, float)): return v
    try: return float(str(v).replace(',', '').replace('%', '').strip())
    except (TypeError, ValueError): return None


def _fmt(v):
    n = _num(v)
    if n is None or isinstance(v, str) and not str(v).replace(',', '').replace('.', '', 1).replace('-', '', 1).isdigit():
        return '' if v is None else str(v)
    return f'{int(n):,}' if float(n).is_integer() else f'{n:,.1f}'


class Deck:
    def __init__(self, title):
        self.title = title
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.prs.core_properties.title = title
        self.prs.core_properties.author = 'RFAI Tracker'
        self.n = 0

    # ── primitives ──────────────────────────────────────────────────────
    def _slide(self, bg=None):
        self.n += 1
        sl = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        if bg is not None:
            sl.background.fill.solid(); sl.background.fill.fore_color.rgb = bg
        return sl

    @staticmethod
    def box(sl, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None, shadow=False):
        s = sl.shapes.add_shape(shape, int(x), int(y), int(w), int(h))
        if fill is None: s.fill.background()
        else: s.fill.solid(); s.fill.fore_color.rgb = fill
        if line is None: s.line.fill.background()
        else: s.line.color.rgb = line; s.line.width = Pt(1)
        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE: s.adjustments[0] = radius
        if shadow:
            sp = s._element.spPr
            eff = sp.makeelement(qn('a:effectLst'), {})
            sh = eff.makeelement(qn('a:outerShdw'), {'blurRad': '190500', 'dist': '38100', 'dir': '5400000', 'algn': 't', 'rotWithShape': '0'})
            c = sh.makeelement(qn('a:srgbClr'), {'val': '1A1D24'}); a = c.makeelement(qn('a:alpha'), {'val': '12000'}); c.append(a); sh.append(c)
            eff.append(sh); sp.append(eff)
        else:
            s.shadow.inherit = False
        tf = s.text_frame; tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        return s

    @staticmethod
    def text(sl, x, y, w, h, t, size=14, bold=False, color=INK, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, italic=False, spacing=None):
        tb = sl.shapes.add_textbox(int(x), int(y), int(w), int(h)); tf = tb.text_frame
        tf.word_wrap = True; tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        lines = t if isinstance(t, list) else [t]
        for i, ln in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            if spacing: p.space_after = Pt(spacing)
            runs = ln if isinstance(ln, tuple) else ((str(ln), {}),)
            for txt, o in runs:
                r = p.add_run(); r.text = str(txt)
                f = r.font; f.size = Pt(o.get('size', size)); f.bold = o.get('bold', bold); f.italic = italic
                f.color.rgb = o.get('color', color); f.name = FONT
        return tf

    def pic(self, sl, name, x, y, size):
        p = ASSETS / 'icons' / f'{name}.png'
        if p.exists(): sl.shapes.add_picture(str(p), int(x), int(y), int(size), int(size))

    def icon_badge(self, sl, icon, x, y, d, bg, white=True):
        """Icon in a coloured circle: the deck's repeated visual motif."""
        self.box(sl, x, y, d, d, bg, shape=MSO_SHAPE.OVAL)
        pad = d * 0.27
        self.pic(sl, f'{icon}_{"w" if white else "r"}', x + pad, y + pad, d - 2 * pad)

    def _logo(self, sl, x, y, h, white=False):
        p = ASSETS / ('airtel_white.png' if white else 'airtel_red.png')
        if p.exists(): sl.shapes.add_picture(str(p), int(x), int(y), height=int(h))

    def frame(self, sl, title, subtitle=None):
        """Content slide chrome: takeaway title, context line, logo and footer at fixed positions."""
        size = 34 if len(title) <= 52 else 28
        self.text(sl, M, Inches(0.42), Inches(10.9), Inches(0.95), title, size, True, INK, anchor=MSO_ANCHOR.BOTTOM)
        if subtitle:
            self.text(sl, M, Inches(1.42), Inches(10.9), Inches(0.4), subtitle, 15, False, MUTED)
        self._logo(sl, W - M - Inches(1.25), Inches(0.5), Inches(0.43))
        self.text(sl, M, H - Inches(0.48), Inches(8), Inches(0.3), f'RFAI Tracker  ·  {self.title}', 10, color=MUTED)
        self.text(sl, W - M - Inches(3), H - Inches(0.48), Inches(3), Inches(0.3), str(self.n), 10, True, MUTED, align=PP_ALIGN.RIGHT)
        return TOP

    # ── slide kinds ─────────────────────────────────────────────────────
    def cover(self, s):
        sl = self._slide()
        bg = self.box(sl, 0, 0, W, H, RED)
        g = bg.fill; g.gradient(); g.gradient_angle = 135
        g.gradient_stops[0].color.rgb = RGBColor(0xF2, 0x14, 0x14); g.gradient_stops[1].color.rgb = RGBColor(0x96, 0x00, 0x0C)
        sw = ASSETS / 'swirl_faint.png'
        if sw.exists(): sl.shapes.add_picture(str(sw), Inches(6.9), Inches(0.4), height=Inches(7.9))
        self._logo(sl, Inches(0.8), Inches(0.75), Inches(0.62), white=True)
        self.text(sl, Inches(0.8), Inches(2.35), Inches(8.3), Inches(2.3), s.get('title') or self.title, 50, True, WHITE, anchor=MSO_ANCHOR.BOTTOM)
        self.text(sl, Inches(0.8), Inches(4.8), Inches(8.0), Inches(0.9), s.get('subtitle') or 'RFAI status review', 20, False, RGBColor(0xFF, 0xDD, 0xDF))
        self.box(sl, Inches(0.8), Inches(6.3), Inches(3.1), Inches(0.46), RGBColor(0x7E, 0x00, 0x0A), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        self.text(sl, Inches(0.8), Inches(6.3), Inches(3.1), Inches(0.46), datetime.now().strftime('%d %B %Y').lstrip('0') + '  ·  RFAI Tracker',
                  12, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        ks = (s.get('kpis') or [])[:3]
        x0, y0, cw, ch = Inches(9.3), Inches(1.45), Inches(3.45), Inches(1.4)
        for i, k in enumerate(ks):
            y = y0 + i * (ch + Inches(0.3))
            self.box(sl, x0, y, cw, ch, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.14, shadow=True)
            self.icon_badge(sl, _icon(k.get('label'), 'tower'), x0 + Inches(0.28), y + (ch - Inches(0.7)) / 2, Inches(0.7), _rgb(_kcol(k.get('label'), i)))
            self.text(sl, x0 + Inches(1.2), y + Inches(0.24), cw - Inches(1.4), Inches(0.3), str(k.get('label', '')).upper(), 11, True, MUTED)
            self.text(sl, x0 + Inches(1.2), y + Inches(0.52), cw - Inches(1.4), Inches(0.7), str(k.get('value', '')), 34, True, INK, anchor=MSO_ANCHOR.MIDDLE)

    def summary(self, s):
        sl = self._slide(); top = self.frame(sl, s.get('title', 'Executive summary'), s.get('subtitle'))
        stats = (s.get('stats') or [])[:3]
        if stats:
            gap = Inches(0.3); cw = (W - 2 * M - gap * (len(stats) - 1)) / len(stats); ch = Inches(1.75)
            for i, st in enumerate(stats):
                x = M + i * (cw + gap)
                hero = i == 0
                self.box(sl, x, top, cw, ch, RED if hero else SOFT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
                self.text(sl, x + Inches(0.4), top + Inches(0.18), cw - Inches(0.8), Inches(1.0), str(st.get('value', '')), 60, True,
                          WHITE if hero else INK, anchor=MSO_ANCHOR.MIDDLE)
                self.text(sl, x + Inches(0.4), top + Inches(1.17), cw - Inches(0.8), Inches(0.4), str(st.get('label', '')), 15, False,
                          RGBColor(0xFF, 0xE0, 0xE2) if hero else INK2)
            top += ch + Inches(0.4)
        pts = (s.get('points') or [])[:3]
        if pts:
            avail = H - top - Inches(0.75); rh = min(Inches(1.05), avail / len(pts))
            for i, p in enumerate(pts):
                if isinstance(p, str): p = {'text': p}
                y = top + i * rh
                self.icon_badge(sl, p.get('icon') or 'lightbulb', M, y + Inches(0.06), Inches(0.62), RED if i == 0 else DARK)
                runs = []
                if p.get('head'): runs.append((p['head'] + '  ', {'bold': True, 'color': INK}))
                runs.append((p.get('text', ''), {'color': INK2}))
                self.text(sl, M + Inches(0.95), y, W - 2 * M - Inches(0.95), rh - Inches(0.1), [tuple(runs)], 16, anchor=MSO_ANCHOR.MIDDLE)

    def section(self, s):
        sl = self._slide(DARK)
        self.icon_badge(sl, s.get('icon') or 'chart', Inches(0.9), Inches(2.2), Inches(0.9), RED)
        self.text(sl, Inches(0.9), Inches(3.4), Inches(11), Inches(1), s.get('title', ''), 42, True, WHITE)
        if s.get('subtitle'): self.text(sl, Inches(0.9), Inches(4.4), Inches(11), Inches(0.8), s['subtitle'], 18, color=RGBColor(0xC9, 0xCD, 0xD6))
        self._logo(sl, W - M - Inches(1.25), H - Inches(0.95), Inches(0.43), white=True)

    def kpis(self, s):
        sl = self._slide(); top = self.frame(sl, s.get('title', 'At a glance'), s.get('subtitle'))
        ks = (s.get('kpis') or [])[:8]
        if not ks: return
        per = 4 if len(ks) > 4 else len(ks); rows = (len(ks) + per - 1) // per
        gap = Inches(0.3); cw = (W - 2 * M - gap * (per - 1)) / per
        avail = H - top - Inches(0.75); ch = min(Inches(2.3), (avail - gap * (rows - 1)) / rows)
        for i, k in enumerate(ks):
            x = M + (cw + gap) * (i % per); y = top + (ch + gap) * (i // per)
            c = _rgb(_kcol(k.get('label'), i))
            self.box(sl, x, y, cw, ch, SOFT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
            self.icon_badge(sl, _icon(k.get('label')), x + Inches(0.3), y + Inches(0.3), Inches(0.56), c)
            self.text(sl, x + Inches(1.02), y + Inches(0.3), cw - Inches(1.2), Inches(0.56), str(k.get('label', '')).upper(), 12, True, INK2, anchor=MSO_ANCHOR.MIDDLE)
            self.text(sl, x + Inches(0.3), y + Inches(0.95), cw - Inches(0.5), Inches(0.75), str(k.get('value', '')), 40, True, INK, anchor=MSO_ANCHOR.MIDDLE)
            if k.get('sub'): self.text(sl, x + Inches(0.3), y + ch - Inches(0.5), cw - Inches(0.5), Inches(0.3), str(k['sub']), 12, color=MUTED)

    def stat(self, s):
        sl = self._slide(); top = self.frame(sl, s.get('title', ''), s.get('subtitle'))
        ph = H - top - Inches(0.75); pw = Inches(5.4)
        self.box(sl, M, top, pw, ph, TINT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.06)
        self.icon_badge(sl, s.get('icon') or _icon(s.get('label', ''), 'clock'), M + Inches(0.5), top + Inches(0.5), Inches(0.8), RED)
        self.text(sl, M + Inches(0.5), top + Inches(1.5), pw - Inches(1), Inches(1.7), str(s.get('value', '')), 96, True, RED, anchor=MSO_ANCHOR.MIDDLE)
        self.text(sl, M + Inches(0.5), top + Inches(3.25), pw - Inches(1), Inches(0.9), str(s.get('label', '')), 18, True, INK)
        tx = M + pw + Inches(0.6)
        body = s.get('text') or []
        body = body if isinstance(body, list) else [body]
        y = top + Inches(0.4)
        for i, t in enumerate(body[:4]):
            self.icon_badge(sl, 'lightbulb' if i == 0 else 'flag', tx, y, Inches(0.5), DARK)
            self.text(sl, tx + Inches(0.8), y - Inches(0.02), W - M - tx - Inches(0.8), Inches(1.0), str(t), 17, color=INK)
            y += Inches(1.15)

    def chart(self, s):
        sl = self._slide(); top = self.frame(sl, s.get('title', ''), s.get('subtitle'))
        cats = [str(c) for c in (s.get('categories') or [])][:30]
        ser = [x for x in (s.get('series') or []) if x.get('values')]
        if not cats or not ser: return
        ct = s.get('chart_type', 'column')
        ins = s.get('insights') or _auto_insights(cats, ser, ct)
        panel = Inches(3.75) if ins else 0
        cx, cy = M, top; cw = W - 2 * M - (panel + Inches(0.4) if panel else 0); chh = H - cy - Inches(0.75)
        cd = CategoryChartData(); cd.categories = cats
        for x in ser: cd.add_series(str(x.get('name', '')), [(_num(v) or 0) for v in x['values'][:len(cats)]])
        kind = {'column': XL_CHART_TYPE.COLUMN_CLUSTERED, 'bar': XL_CHART_TYPE.BAR_CLUSTERED, 'line': XL_CHART_TYPE.LINE_MARKERS,
                'pie': XL_CHART_TYPE.DOUGHNUT, 'stacked_column': XL_CHART_TYPE.COLUMN_STACKED}.get(ct, XL_CHART_TYPE.COLUMN_CLUSTERED)
        if ct == 'bar': cd = _reverse(cd, cats, ser)
        ch = sl.shapes.add_chart(kind, int(cx), int(cy), int(cw), int(chh), cd).chart
        ch.font.size = Pt(12); ch.font.name = FONT; ch.font.color.rgb = INK2
        ch.has_title = False
        multi = len(ser) > 1
        ch.has_legend = multi or ct == 'pie'
        if ch.has_legend:
            ch.legend.position = XL_LEGEND_POSITION.TOP if ct != 'pie' else XL_LEGEND_POSITION.RIGHT
            ch.legend.include_in_layout = False; ch.legend.font.size = Pt(12); ch.legend.font.color.rgb = INK2
        plot = ch.plots[0]
        if ct == 'pie':
            hole = plot._element.find(qn('c:holeSize'))
            if hole is None:
                hole = plot._element.makeelement(qn('c:holeSize'), {}); plot._element.append(hole)
            hole.set('val', '60')
            for i, pt in enumerate(plot.series[0].points):
                pt.format.fill.solid(); pt.format.fill.fore_color.rgb = _rgb(_col(cats[i], i))
                pt.format.line.color.rgb = WHITE; pt.format.line.width = Pt(2.5)
            plot.has_data_labels = True; dl = plot.data_labels
            dl.show_percentage = True; dl.show_value = False; dl.show_category_name = False
            dl.number_format = '0%'; dl.number_format_is_linked = False
            dl.font.size = Pt(14); dl.font.bold = True; dl.font.color.rgb = WHITE
            tot = sum((_num(v) or 0) for v in ser[0]['values'][:len(cats)])
            ring_cx = cx + (cw - Inches(2.2)) / 2
            self.text(sl, ring_cx - Inches(1.3), cy + chh / 2 - Inches(0.5), Inches(2.6), Inches(0.55), _fmt(tot), 32, True, INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            self.text(sl, ring_cx - Inches(1.3), cy + chh / 2 + Inches(0.07), Inches(2.6), Inches(0.35), 'TOTAL SRS', 11, True, MUTED, PP_ALIGN.CENTER)
        else:
            vals0 = [(_num(v) or 0) for v in ser[0]['values'][:len(cats)]]
            va, ca = ch.value_axis, ch.category_axis
            va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = LINE; va.major_gridlines.format.line.width = Pt(0.75)
            va.format.line.fill.background(); va.tick_labels.font.size = Pt(11); va.tick_labels.font.color.rgb = MUTED
            va.tick_labels.number_format = '#,##0'; va.tick_labels.number_format_is_linked = False
            if ct != 'line' or min(vals0 or [0]) >= 0: va.minimum_scale = 0      # bars always start at zero
            ca.format.line.color.rgb = LINE; ca.tick_labels.font.size = Pt(12); ca.tick_labels.font.color.rgb = INK2
            ca.has_major_gridlines = False
            if ct != 'line':
                plot.gap_width = 50 if not multi or ct == 'stacked_column' else 80
                if ct == 'stacked_column': plot.overlap = 100
            for i, se in enumerate(plot.series):
                c = _rgb(_col(se.name, i)) if multi else RED
                if ct == 'line':
                    se.format.line.color.rgb = c; se.format.line.width = Pt(3.25); se.smooth = False
                    se.marker.format.fill.solid(); se.marker.format.fill.fore_color.rgb = WHITE
                    se.marker.format.line.color.rgb = c; se.marker.size = 9
                else:
                    se.format.fill.solid(); se.format.fill.fore_color.rgb = c
                    if not multi and len(vals0) > 2:
                        # one series: grey bars, the largest one highlighted in red
                        top_i = max(range(len(vals0)), key=lambda j: vals0[j])
                        if ct == 'bar': top_i = len(vals0) - 1 - top_i
                        for j, pt in enumerate(se.points):
                            pt.format.fill.solid(); pt.format.fill.fore_color.rgb = RED if j == top_i else RGBColor(0xC7, 0xCB, 0xD3)
            if len(cats) <= 24:
                plot.has_data_labels = True; dl = plot.data_labels
                dl.number_format = '#,##0'; dl.number_format_is_linked = False; dl.font.size = Pt(12)
                if ct == 'stacked_column':
                    dl.position = XL_LABEL_POSITION.CENTER; dl.font.color.rgb = WHITE; dl.font.bold = True
                elif ct == 'line':
                    dl.position = XL_LABEL_POSITION.ABOVE; dl.font.color.rgb = INK2; dl.font.bold = True
                else:
                    dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.color.rgb = INK; dl.font.bold = True
        if ins: self._insights(sl, W - M - panel, cy, panel, chh, ins)

    def _insights(self, sl, x, y, w, h, items):
        self.box(sl, x, y, w, h, SOFT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        self.icon_badge(sl, 'lightbulb', x + Inches(0.3), y + Inches(0.3), Inches(0.5), RED)
        self.text(sl, x + Inches(0.95), y + Inches(0.3), w - Inches(1.2), Inches(0.5), 'Key insights', 16, True, INK, anchor=MSO_ANCHOR.MIDDLE)
        yy = y + Inches(1.1)
        for i, t in enumerate(items[:4]):
            self.text(sl, x + Inches(0.3), yy, Inches(0.45), Inches(0.4), f'{i + 1:02d}', 15, True, RED)
            self.text(sl, x + Inches(0.8), yy, w - Inches(1.05), Inches(1.1), str(t), 14, color=INK)
            yy += Inches(min(1.3, 0.5 + 0.23 * (len(str(t)) // 30 + 1)))

    def table(self, s):
        cols = [str(c) for c in (s.get('columns') or [])]
        rows = s.get('rows') or []
        if not cols: return
        per = 12
        chunks = [rows[i:i + per] for i in range(0, max(len(rows), 1), per)][:6]
        numeric = [any(j < len(r) and r[j] not in (None, '') for r in rows) and
                   all(_num(r[j]) is not None for r in rows if j < len(r) and r[j] not in (None, '') and str(r[0]).lower() != 'total')
                   for j in range(len(cols))]
        for k, part in enumerate(chunks):
            sl = self._slide()
            t = s.get('title', '') + (f' ({k + 1}/{len(chunks)})' if len(chunks) > 1 else '')
            top = self.frame(sl, t, s.get('subtitle'))
            lens = [max([len(cols[j])] + [len(_fmt(r[j]) if numeric[j] else str(r[j])) for r in rows[:200] if j < len(r)]) for j in range(len(cols))]
            lens = [min(max(l, 6), 34) for l in lens]
            tw = W - 2 * M
            big = len(part) <= 8
            rh = Inches(0.52) if big else Inches(0.37); fs = 15 if big else 13
            tb = sl.shapes.add_table(len(part) + 1, len(cols), int(M), int(top), int(tw), int(rh * (len(part) + 1))).table
            tblPr = tb._tbl.tblPr; tblPr.set('bandRow', '0'); tblPr.set('firstRow', '0')
            sid = tblPr.find(qn('a:tableStyleId'))
            if sid is None: sid = tblPr.makeelement(qn('a:tableStyleId'), {}); tblPr.append(sid)
            sid.text = '{5940675A-B579-460E-94D1-54222C63F5DA}'   # "No Style, Table Grid"
            tot = sum(lens)
            for j in range(len(cols)): tb.columns[j].width = int(tw * lens[j] / tot)
            for i in range(len(part) + 1): tb.rows[i].height = int(rh)
            for j, c in enumerate(cols):
                _cell(tb.cell(0, j), c.upper(), True, WHITE, DARK, PP_ALIGN.RIGHT if numeric[j] else PP_ALIGN.LEFT, fs - 2)
            # highlight the largest value of each numeric column (excluding totals)
            peak = {}
            for j in range(1, len(cols)):
                if numeric[j] and any(w in cols[j].lower() for w in ('pending', 'pri', 'reject', 'day')):
                    vals = [(_num(r[j]) or 0) for r in rows if j < len(r) and str(r[0]).lower() != 'total']
                    if len(vals) > 2: peak[j] = max(vals)
            for i, r in enumerate(part, 1):
                tr = str(r[0]).strip().lower() in ('total', 'grand total')
                fill = TINT if tr else (WHITE if i % 2 else SOFT)
                for j in range(len(cols)):
                    v = r[j] if j < len(r) else ''
                    txt = _fmt(v) if numeric[j] else ('' if v is None else str(v))
                    color = INK; bold = tr or j == 0
                    sc = STATUS.get(txt.strip().lower())
                    if sc and not numeric[j]: color = _rgb(sc); bold = True
                    if numeric[j] and 'day' in cols[j].lower() and (_num(v) or 0) > 90: color = RED; bold = True
                    if not tr and j in peak and (_num(v) or 0) == peak[j] and peak[j] > 0: color = RED; bold = True
                    _cell(tb.cell(i, j), txt, bold, color, fill, PP_ALIGN.RIGHT if numeric[j] else PP_ALIGN.LEFT, fs)

    def bullets(self, s):
        sl = self._slide(); top = self.frame(sl, s.get('title', ''), s.get('subtitle'))
        items = [str(b) for b in (s.get('bullets') or [])][:6]
        if not items: return
        two = len(items) > 3
        per_col = (len(items) + 1) // 2 if two else len(items)
        gap = Inches(0.3); cw = (W - 2 * M - (gap if two else 0)) / (2 if two else 1)
        avail = H - top - Inches(0.75); ch = min(Inches(1.4), (avail - gap * (per_col - 1)) / per_col)
        icons = ['bullseye', 'pin', 'alert', 'ban', 'calendar', 'trend']
        for i, t in enumerate(items):
            col, row = (i // per_col, i % per_col) if two else (0, i)
            x = M + col * (cw + gap); y = top + row * (ch + gap)
            self.box(sl, x, y, cw, ch, SOFT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
            d = Inches(0.62)
            self.icon_badge(sl, _icon(t, icons[i % len(icons)]) if i else 'lightbulb', x + Inches(0.3), y + (ch - d) / 2, d, RED if i == 0 else DARK)
            self.text(sl, x + Inches(1.2), y + Inches(0.1), cw - Inches(1.45), ch - Inches(0.2), t, 16, False, INK, anchor=MSO_ANCHOR.MIDDLE)

    def closing(self, s):
        sl = self._slide(DARK)
        sw = ASSETS / 'swirl_faint.png'
        if sw.exists(): sl.shapes.add_picture(str(sw), Inches(8.6), Inches(1.6), height=Inches(6.2))
        self._logo(sl, M, Inches(0.55), Inches(0.5), white=True)
        self.text(sl, M, Inches(1.35), Inches(9), Inches(0.9), s.get('title') or 'Next steps', 40, True, WHITE)
        self.text(sl, M, Inches(2.2), Inches(9), Inches(0.5), s.get('subtitle') or 'Suggested actions from this review', 16, color=RGBColor(0xB8, 0xBD, 0xC8))
        items = [str(b) for b in (s.get('bullets') or [])][:5]
        y = Inches(3.0); rh = min(Inches(0.85), (H - y - Inches(0.9)) / max(len(items), 1))
        for i, t in enumerate(items):
            self.box(sl, M, y + i * rh, Inches(8.4), rh - Inches(0.15), DARK2, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
            self.box(sl, M + Inches(0.18), y + i * rh + (rh - Inches(0.15) - Inches(0.44)) / 2, Inches(0.44), Inches(0.44), RED, shape=MSO_SHAPE.OVAL)
            self.text(sl, M + Inches(0.18), y + i * rh + (rh - Inches(0.15) - Inches(0.44)) / 2, Inches(0.44), Inches(0.44), str(i + 1), 14, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            self.text(sl, M + Inches(0.85), y + i * rh, Inches(7.4), rh - Inches(0.15), t, 16, False, WHITE, anchor=MSO_ANCHOR.MIDDLE)
        self.text(sl, M, H - Inches(0.55), Inches(9), Inches(0.3),
                  'Generated from live RFAI Tracker data on ' + datetime.now().strftime('%d %b %Y, %H:%M'), 11, color=RGBColor(0x8F, 0x95, 0xA3))


def _cell(c, txt, bold, color, fill, align, size):
    """Table cell with a single bottom hairline (border elements must precede the fill in tcPr)."""
    c.margin_left = c.margin_right = Inches(0.14); c.margin_top = c.margin_bottom = Inches(0.03)
    c.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = c.text_frame; tf.word_wrap = True; p = tf.paragraphs[0]; p.alignment = align
    for r_ in list(p.runs): r_._r.getparent().remove(r_._r)
    r_ = p.add_run(); r_.text = txt
    r_.font.size = Pt(size); r_.font.bold = bold; r_.font.color.rgb = color; r_.font.name = FONT
    tcPr = c._tc.get_or_add_tcPr()
    for ch in list(tcPr):
        if ch.tag in (qn('a:lnL'), qn('a:lnR'), qn('a:lnT'), qn('a:lnB'), qn('a:solidFill'), qn('a:noFill')): tcPr.remove(ch)
    els = []
    for edge in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        ln = tcPr.makeelement(qn(edge), {'w': '9525' if edge == 'a:lnB' else '0'})
        if edge == 'a:lnB':
            sf = ln.makeelement(qn('a:solidFill'), {}); sf.append(sf.makeelement(qn('a:srgbClr'), {'val': 'E6E8EE'})); ln.append(sf)
        else:
            ln.append(ln.makeelement(qn('a:noFill'), {}))
        els.append(ln)
    f = tcPr.makeelement(qn('a:solidFill'), {}); f.append(f.makeelement(qn('a:srgbClr'), {'val': str(fill)})); els.append(f)
    for i, e in enumerate(els): tcPr.insert(i, e)


def _reverse(cd, cats, ser):
    """Horizontal bars list from the top, so reverse the order."""
    nd = CategoryChartData(); nd.categories = list(reversed(cats))
    for x in ser: nd.add_series(str(x.get('name', '')), list(reversed([(_num(v) or 0) for v in x['values'][:len(cats)]])))
    return nd


def _auto_insights(cats, ser, ct):
    if not cats: return []
    if len(ser) == 1:
        vals = [(_num(v) or 0) for v in ser[0]['values'][:len(cats)]]
    else:
        vals = [sum((_num(x['values'][i]) or 0) if i < len(x['values']) else 0 for x in ser) for i in range(len(cats))]
    tot = sum(vals) or 1
    pairs = list(zip(cats, vals))
    out = []
    if ct == 'line' and len(pairs) > 1:
        hi = max(pairs, key=lambda p: p[1])
        out.append(f'Peak in {hi[0]} at {_fmt(hi[1])}.')
        full = pairs[:-1] if len(pairs) > 2 else pairs      # the latest period is usually still running
        a, b = full[0][1], full[-1][1]
        if a: out.append(f'{full[-1][0]} is {"up" if b >= a else "down"} {abs(b - a) / a:.0%} vs {full[0][0]}.')
        out.append(f'Average {_fmt(round(sum(vals) / len(vals)))} per period, {_fmt(sum(vals))} in total.')
        return out
    srt = sorted(pairs, key=lambda p: p[1], reverse=True)
    out.append(f'{srt[0][0]} is highest with {_fmt(srt[0][1])} ({srt[0][1] / tot:.0%} of total).')
    if len(srt) >= 4:
        out.append(f'Top 3 ({", ".join(p[0] for p in srt[:3])}) make up {sum(p[1] for p in srt[:3]) / tot:.0%}.')
    if len(srt) >= 2:
        out.append(f'{srt[-1][0]} is lowest with {_fmt(srt[-1][1])}.')
    if len(ser) > 1:
        sums = [(x['name'], sum((_num(v) or 0) for v in x['values'])) for x in ser]
        big = max(sums, key=lambda p: p[1]); out.append(f'{big[0]} is the largest part: {_fmt(big[1])} ({big[1] / tot:.0%}).')
    return out


def build_deck(path, title, slides):
    d = Deck(title)
    slides = [dict(s) for s in (slides or [])][:40]
    if not slides or slides[0].get('kind') != 'title':
        slides.insert(0, {'kind': 'title', 'title': title})
    if slides[-1].get('kind') != 'closing':
        # the last bullets slide called "next steps" becomes the dark closing slide
        if slides[-1].get('kind') == 'bullets' and 'next' in str(slides[-1].get('title', '')).lower():
            slides[-1]['kind'] = 'closing'
        else:
            slides.append({'kind': 'closing', 'title': 'Next steps', 'bullets': [
                'Review the pending SRs with circle teams', 'Escalate SRs pending over 90 days',
                'Close open PRIs so blocked sites can go OnAir', 'Track the weekly OnAir run-rate']})
    for s in slides:
        k = s.get('kind', 'bullets')
        before = len(d.prs.slides)
        {'title': d.cover, 'summary': d.summary, 'kpis': d.kpis, 'stat': d.stat, 'chart': d.chart, 'table': d.table,
         'bullets': d.bullets, 'section': d.section, 'closing': d.closing}.get(k, d.bullets)(s)
        if s.get('notes') and len(d.prs.slides) > before:
            d.prs.slides[before].notes_slide.notes_text_frame.text = str(s['notes'])
    d.prs.save(path)
