"""
Airtel-styled PowerPoint decks for the RFAI assistant.

build_deck(path, title, slides) takes a list of slide dicts:
  title    : title, subtitle, kpis (optional, up to 3 shown on the cover)
  kpis     : title, subtitle, kpis [{label, value, sub?}]
  chart    : title, subtitle, chart_type (column|bar|line|pie|stacked_column), categories,
             series [{name, values}], insights (optional; auto-written when missing)
  table    : title, subtitle, columns, rows (long tables continue on more slides)
  bullets  : title, subtitle, bullets
  section  : title, subtitle
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
RED_D = RGBColor(0xA8, 0x00, 0x0D)
INK = RGBColor(0x16, 0x18, 0x1F)
INK2 = RGBColor(0x4B, 0x50, 0x5C)
MUTED = RGBColor(0x8A, 0x8F, 0x9A)
LINE = RGBColor(0xE6, 0xE8, 0xEE)
SOFT = RGBColor(0xF6, 0xF7, 0xFA)
TINT = RGBColor(0xFF, 0xF1, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1A, 0x1D, 0x24)

STATUS = {'onair': '10B981', 'onair pending': 'F59E0B', 'pending': 'F59E0B', 'open pri': '3B82F6',
          'rejected rfai': 'EF4444', 'rejected': 'EF4444', 'rfai pending': '94A3B8',
          'total': 'E40000', 'total srs': 'E40000', 'srs': 'E40000'}
PAL = ['E40000', '1F2937', 'F59E0B', '3B82F6', '10B981', '8B5CF6', '94A3B8', 'EC4899']
W, H = Inches(13.333), Inches(7.5)
M = Inches(0.6)


def _rgb(h): return RGBColor.from_string(h)


def _col(name, i):
    return STATUS.get(str(name).strip().lower(), PAL[i % len(PAL)])


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
        self.n = 0

    # ── primitives ──────────────────────────────────────────────────────
    def _slide(self):
        self.n += 1
        return self.prs.slides.add_slide(self.prs.slide_layouts[6])

    @staticmethod
    def box(sl, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None):
        s = sl.shapes.add_shape(shape, int(x), int(y), int(w), int(h))
        if fill is None: s.fill.background()
        else: s.fill.solid(); s.fill.fore_color.rgb = fill
        if line is None: s.line.fill.background()
        else: s.line.color.rgb = line; s.line.width = Pt(0.75)
        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE: s.adjustments[0] = radius
        s.shadow.inherit = False
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
            r = p.add_run(); r.text = str(ln)
            f = r.font; f.size = Pt(size); f.bold = bold; f.italic = italic; f.color.rgb = color; f.name = FONT
        return tf

    def _logo(self, sl, x, y, h, white=False):
        p = ASSETS / ('airtel_white.png' if white else 'airtel_red.png')
        if p.exists(): sl.shapes.add_picture(str(p), int(x), int(y), height=int(h))

    def frame(self, sl, title, subtitle=None):
        """Content slide chrome: accent, action title, takeaway line, logo, footer."""
        self.box(sl, 0, 0, W, Inches(0.08), RED)
        self.box(sl, M, Inches(0.5), Inches(0.09), Inches(0.5), RED)
        self.text(sl, M + Inches(0.25), Inches(0.42), Inches(10.4), Inches(0.6), title, 26, True, INK, anchor=MSO_ANCHOR.MIDDLE)
        if subtitle:
            self.text(sl, M + Inches(0.25), Inches(1.02), Inches(10.6), Inches(0.45), subtitle, 14, False, INK2)
        self._logo(sl, W - M - Inches(1.25), Inches(0.45), Inches(0.43))
        self.box(sl, M, H - Inches(0.55), W - 2 * M, Emu(9525), LINE)
        self.text(sl, M, H - Inches(0.45), Inches(8), Inches(0.3), f'RFAI Tracker  ·  {self.title}', 10, color=MUTED)
        self.text(sl, W - M - Inches(3), H - Inches(0.45), Inches(3), Inches(0.3), f'Internal  ·  {self.n}', 10, color=MUTED, align=PP_ALIGN.RIGHT)
        return Inches(1.65) if subtitle else Inches(1.35)

    # ── slide kinds ─────────────────────────────────────────────────────
    def cover(self, s):
        sl = self._slide()
        self.box(sl, 0, 0, W, H, RED)
        g = sl.shapes[-1].fill; g.gradient(); g.gradient_angle = 135
        g.gradient_stops[0].color.rgb = RGBColor(0xF0, 0x12, 0x12); g.gradient_stops[1].color.rgb = RGBColor(0x9E, 0x00, 0x0C)
        sw = ASSETS / 'swirl_faint.png'
        if sw.exists(): sl.shapes.add_picture(str(sw), Inches(7.2), Inches(0.6), height=Inches(7.6))
        self._logo(sl, Inches(0.8), Inches(0.7), Inches(0.62), white=True)
        self.box(sl, Inches(0.8), Inches(2.55), Inches(0.9), Inches(0.07), WHITE)
        self.text(sl, Inches(0.8), Inches(2.8), Inches(8.6), Inches(1.9), s.get('title') or self.title, 44, True, WHITE, anchor=MSO_ANCHOR.TOP)
        self.text(sl, Inches(0.8), Inches(4.75), Inches(8.2), Inches(0.9), s.get('subtitle') or 'RFAI status review', 18, False, RGBColor(0xFF, 0xDD, 0xDF))
        chip = self.box(sl, Inches(0.8), Inches(6.35), Inches(2.9), Inches(0.42), RGBColor(0x8A, 0x00, 0x0B), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
        self.text(sl, Inches(0.8), Inches(6.35), Inches(2.9), Inches(0.42), datetime.now().strftime('%d %B %Y').lstrip('0') + '  ·  RFAI Tracker', 12, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
        ks = (s.get('kpis') or [])[:3]
        if ks:
            x0, y0, cw, ch = Inches(9.45), Inches(1.6), Inches(3.2), Inches(1.32)
            for i, k in enumerate(ks):
                y = y0 + i * (ch + Inches(0.22))
                self.box(sl, x0, y, cw, ch, WHITE, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
                self.box(sl, x0, y + Inches(0.28), Inches(0.08), ch - Inches(0.56), _rgb(_col(k.get('label'), i)))
                self.text(sl, x0 + Inches(0.32), y + Inches(0.2), cw - Inches(0.5), Inches(0.3), str(k.get('label', '')).upper(), 11, True, MUTED)
                self.text(sl, x0 + Inches(0.32), y + Inches(0.48), cw - Inches(0.5), Inches(0.65), str(k.get('value', '')), 32, True, INK, anchor=MSO_ANCHOR.MIDDLE)

    def section(self, s):
        sl = self._slide()
        self.box(sl, 0, 0, W, H, DARK)
        self.box(sl, 0, 0, Inches(0.25), H, RED)
        self.text(sl, Inches(1), Inches(2.9), Inches(11), Inches(1), s.get('title', ''), 40, True, WHITE)
        if s.get('subtitle'): self.text(sl, Inches(1), Inches(3.9), Inches(11), Inches(0.8), s['subtitle'], 18, color=RGBColor(0xC9, 0xCD, 0xD6))

    def kpis(self, s):
        sl = self._slide(); top = self.frame(sl, s.get('title', 'At a glance'), s.get('subtitle'))
        ks = (s.get('kpis') or [])[:8]
        if not ks: return
        per = 4 if len(ks) > 4 else len(ks); rows = (len(ks) + per - 1) // per
        gap = Inches(0.3); cw = (W - 2 * M - gap * (per - 1)) / per
        avail = H - top - Inches(0.9); ch = min(Inches(2.1), (avail - gap * (rows - 1)) / rows)
        y0 = top + Inches(0.25) + (avail - rows * ch - gap * (rows - 1)) / 2
        for i, k in enumerate(ks):
            x = M + (cw + gap) * (i % per); y = y0 + (ch + gap) * (i // per)
            c = _rgb(_col(k.get('label'), i))
            self.box(sl, x, y, cw, ch, SOFT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
            self.box(sl, x, y, cw, Inches(0.09), c)
            self.box(sl, x + Inches(0.35), y + Inches(0.45), Inches(0.16), Inches(0.16), c, shape=MSO_SHAPE.OVAL)
            self.text(sl, x + Inches(0.62), y + Inches(0.38), cw - Inches(0.8), Inches(0.32), str(k.get('label', '')).upper(), 12, True, INK2)
            self.text(sl, x + Inches(0.35), y + Inches(0.78), cw - Inches(0.6), Inches(0.8), str(k.get('value', '')), 40, True, INK, anchor=MSO_ANCHOR.MIDDLE)
            if k.get('sub'): self.text(sl, x + Inches(0.35), y + ch - Inches(0.5), cw - Inches(0.6), Inches(0.3), str(k['sub']), 12, color=MUTED)

    def chart(self, s):
        sl = self._slide(); top = self.frame(sl, s.get('title', ''), s.get('subtitle'))
        cats = [str(c) for c in (s.get('categories') or [])][:30]
        ser = [x for x in (s.get('series') or []) if x.get('values')]
        if not cats or not ser: return
        ct = s.get('chart_type', 'column')
        ins = s.get('insights') or _auto_insights(cats, ser, ct)
        panel = Inches(3.7) if ins else 0
        cx, cy = M, top + Inches(0.15); cw = W - 2 * M - (panel + Inches(0.3) if panel else 0); chh = H - cy - Inches(0.75)
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
            ch.legend.position = XL_LEGEND_POSITION.BOTTOM if ct != 'pie' else XL_LEGEND_POSITION.RIGHT
            ch.legend.include_in_layout = False; ch.legend.font.size = Pt(12)
        plot = ch.plots[0]
        if ct == 'pie':
            hole = plot._element.find(qn('c:holeSize'))
            if hole is None:
                hole = plot._element.makeelement(qn('c:holeSize'), {}); plot._element.append(hole)
            hole.set('val', '58')
            for i, pt in enumerate(plot.series[0].points):
                pt.format.fill.solid(); pt.format.fill.fore_color.rgb = _rgb(_col(cats[i], i))
                pt.format.line.color.rgb = WHITE; pt.format.line.width = Pt(2)
            plot.has_data_labels = True; dl = plot.data_labels
            dl.show_percentage = True; dl.show_value = False; dl.show_category_name = False
            dl.number_format = '0%'; dl.number_format_is_linked = False
            dl.font.size = Pt(13); dl.font.bold = True; dl.font.color.rgb = WHITE
            tot = sum((_num(v) or 0) for v in ser[0]['values'][:len(cats)])
            # centre total (legend sits on the right, so the ring is left of centre)
            ring_cx = cx + (cw - Inches(2.2)) / 2
            self.text(sl, ring_cx - Inches(1.3), cy + chh / 2 - Inches(0.5), Inches(2.6), Inches(0.55), _fmt(tot), 30, True, INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            self.text(sl, ring_cx - Inches(1.3), cy + chh / 2 + Inches(0.05), Inches(2.6), Inches(0.35), 'TOTAL', 11, True, MUTED, PP_ALIGN.CENTER)
        else:
            va, ca = ch.value_axis, ch.category_axis
            va.has_major_gridlines = True; va.major_gridlines.format.line.color.rgb = LINE; va.major_gridlines.format.line.width = Pt(0.75)
            va.format.line.fill.background(); va.tick_labels.font.size = Pt(11); va.tick_labels.font.color.rgb = MUTED
            va.tick_labels.number_format = '#,##0'; va.tick_labels.number_format_is_linked = False
            ca.format.line.color.rgb = LINE; ca.tick_labels.font.size = Pt(11 if len(cats) > 12 else 12); ca.tick_labels.font.color.rgb = INK2
            ca.has_major_gridlines = False
            if ct != 'line':
                plot.gap_width = 55 if not multi or ct == 'stacked_column' else 80
                if ct == 'stacked_column': plot.overlap = 100
            for i, se in enumerate(plot.series):
                c = _rgb(_col(se.name, i) if multi else ('E40000' if ct != 'line' else 'E40000'))
                if ct == 'line':
                    se.format.line.color.rgb = c; se.format.line.width = Pt(3); se.smooth = False
                    se.marker.format.fill.solid(); se.marker.format.fill.fore_color.rgb = WHITE
                    se.marker.format.line.color.rgb = c; se.marker.size = 8
                else:
                    se.format.fill.solid(); se.format.fill.fore_color.rgb = c
            if len(cats) <= 24:
                plot.has_data_labels = True; dl = plot.data_labels
                dl.number_format = '#,##0'; dl.number_format_is_linked = False; dl.font.size = Pt(11 if len(cats) > 12 else 12)
                if ct == 'stacked_column':
                    dl.position = XL_LABEL_POSITION.CENTER; dl.font.color.rgb = WHITE; dl.font.bold = True
                elif ct == 'line':
                    dl.position = XL_LABEL_POSITION.ABOVE; dl.font.color.rgb = INK2
                else:
                    dl.position = XL_LABEL_POSITION.OUTSIDE_END; dl.font.color.rgb = INK; dl.font.bold = True
        if ins: self._insights(sl, W - M - panel, cy, panel, chh, ins)

    def _insights(self, sl, x, y, w, h, items):
        self.box(sl, x, y, w, h, SOFT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        self.box(sl, x, y, w, Inches(0.07), RED)
        self.text(sl, x + Inches(0.3), y + Inches(0.3), w - Inches(0.6), Inches(0.35), 'KEY INSIGHTS', 12, True, RED)
        yy = y + Inches(0.8)
        for i, t in enumerate(items[:4]):
            self.box(sl, x + Inches(0.3), yy + Inches(0.02), Inches(0.34), Inches(0.34), RED if i == 0 else DARK, shape=MSO_SHAPE.OVAL)
            self.text(sl, x + Inches(0.3), yy + Inches(0.02), Inches(0.34), Inches(0.34), str(i + 1), 12, True, WHITE, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            self.text(sl, x + Inches(0.8), yy, w - Inches(1.05), Inches(1.1), str(t), 13, color=INK)
            yy += Inches(min(1.25, 0.42 + 0.21 * (len(str(t)) // 28 + 1)))

    def table(self, s):
        cols = [str(c) for c in (s.get('columns') or [])]
        rows = s.get('rows') or []
        if not cols: return
        per = 13
        chunks = [rows[i:i + per] for i in range(0, max(len(rows), 1), per)][:6]
        for k, part in enumerate(chunks):
            sl = self._slide()
            t = s.get('title', '') + (f'  ({k + 1}/{len(chunks)})' if len(chunks) > 1 else '')
            top = self.frame(sl, t, s.get('subtitle') if k == 0 else None) + Inches(0.1)
            numeric = [all(_num(r[j]) is not None for r in rows if j < len(r) and r[j] not in (None, '')) and any(j < len(r) and r[j] not in (None, '') for r in rows) for j in range(len(cols))]
            total_row = [bool(part) and str(r[0]).strip().lower() in ('total', 'grand total') for r in part]
            lens = [max([len(cols[j])] + [len(_fmt(r[j]) if numeric[j] else str(r[j])) for r in rows[:200] if j < len(r)]) for j in range(len(cols))]
            lens = [min(max(l, 5), 34) for l in lens]
            tw = W - 2 * M
            big = len(part) <= 8
            rh = Inches(0.5) if big else Inches(0.36); fs = 14 if big else 12
            gt = sl.shapes.add_table(len(part) + 1, len(cols), int(M), int(top), int(tw), int(rh * (len(part) + 1)))
            tb = gt.table
            tblPr = tb._tbl.tblPr; tblPr.set('bandRow', '0'); tblPr.set('firstRow', '0')
            sid = tblPr.find(qn('a:tableStyleId'))
            if sid is None: sid = tblPr.makeelement(qn('a:tableStyleId'), {}); tblPr.append(sid)
            sid.text = '{5940675A-B579-460E-94D1-54222C63F5DA}'   # "No Style, Table Grid": no theme colours
            tot = sum(lens)
            for j in range(len(cols)): tb.columns[j].width = int(tw * lens[j] / tot)
            for i in range(len(part) + 1): tb.rows[i].height = int(rh)

            def cell(c, txt, bold=False, color=INK, fill=WHITE, align=PP_ALIGN.LEFT, size=None):
                size = size or fs
                c.fill.solid(); c.fill.fore_color.rgb = fill
                c.margin_left = c.margin_right = Inches(0.12); c.margin_top = c.margin_bottom = Inches(0.03)
                c.vertical_anchor = MSO_ANCHOR.MIDDLE
                tf = c.text_frame; tf.word_wrap = True; p = tf.paragraphs[0]; p.alignment = align
                for r_ in list(p.runs): r_._r.getparent().remove(r_._r)
                r_ = p.add_run(); r_.text = txt
                r_.font.size = Pt(size); r_.font.bold = bold; r_.font.color.rgb = color; r_.font.name = FONT
                for edge in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
                    tcPr = c._tc.get_or_add_tcPr()
                    old = tcPr.find(qn(edge))
                    if old is not None: tcPr.remove(old)
                    ln = tcPr.makeelement(qn(edge), {'w': '6350' if edge == 'a:lnB' else '0'})
                    if edge == 'a:lnB':
                        sf = ln.makeelement(qn('a:solidFill'), {}); clr = sf.makeelement(qn('a:srgbClr'), {'val': 'E6E8EE'}); sf.append(clr); ln.append(sf)
                    else:
                        ln.append(ln.makeelement(qn('a:noFill'), {}))
                    tcPr.append(ln)

            for j, c in enumerate(cols):
                cell(tb.cell(0, j), c.upper(), True, WHITE, DARK, PP_ALIGN.RIGHT if numeric[j] else PP_ALIGN.LEFT, fs - 1)
            for i, r in enumerate(part, 1):
                tr = total_row[i - 1]
                fill = TINT if tr else (WHITE if i % 2 else SOFT)
                for j in range(len(cols)):
                    v = r[j] if j < len(r) else ''
                    txt = _fmt(v) if numeric[j] else ('' if v is None else str(v))
                    color = INK
                    sc = STATUS.get(txt.strip().lower())
                    if sc and not numeric[j] and txt.strip().lower() not in ('total', 'srs'): color = _rgb(sc)
                    if numeric[j] and 'day' in cols[j].lower() and (_num(v) or 0) > 90: color = RED
                    cell(tb.cell(i, j), txt, tr or j == 0 or bool(sc and not numeric[j]), color, fill,
                         PP_ALIGN.RIGHT if numeric[j] else PP_ALIGN.LEFT)

    def bullets(self, s):
        sl = self._slide(); top = self.frame(sl, s.get('title', ''), s.get('subtitle'))
        items = [str(b) for b in (s.get('bullets') or [])][:8]
        if not items: return
        two = len(items) > 4
        per_col = (len(items) + 1) // 2 if two else len(items)
        gap = Inches(0.3); cw = (W - 2 * M - (gap if two else 0)) / (2 if two else 1)
        avail = H - top - Inches(0.95); ch = min(Inches(1.15), (avail - gap * (per_col - 1)) / per_col)
        for i, t in enumerate(items):
            col, row = (i // per_col, i % per_col) if two else (0, i)
            x = M + col * (cw + gap); y = top + Inches(0.15) + row * (ch + Inches(0.18))
            self.box(sl, x, y, cw, ch, SOFT, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
            self.box(sl, x, y + Inches(0.18), Inches(0.07), ch - Inches(0.36), RED if i == 0 else DARK)
            self.box(sl, x + Inches(0.3), y + (ch - Inches(0.5)) / 2, Inches(0.5), Inches(0.5), RED if i == 0 else WHITE, LINE if i else None, MSO_SHAPE.OVAL)
            self.text(sl, x + Inches(0.3), y + (ch - Inches(0.5)) / 2, Inches(0.5), Inches(0.5), f'{i + 1:02d}', 13, True, WHITE if i == 0 else INK, PP_ALIGN.CENTER, MSO_ANCHOR.MIDDLE)
            self.text(sl, x + Inches(1.0), y, cw - Inches(1.25), ch, t, 16 if len(t) < 90 else 14, False, INK, anchor=MSO_ANCHOR.MIDDLE)

    def closing(self):
        sl = self._slide()
        self.box(sl, 0, 0, W, H, DARK)
        sw = ASSETS / 'swirl_faint.png'
        if sw.exists(): sl.shapes.add_picture(str(sw), Inches(8.3), Inches(1.1), height=Inches(6.4))
        self._logo(sl, Inches(0.9), Inches(0.8), Inches(0.55), white=True)
        self.box(sl, Inches(0.9), Inches(3.05), Inches(0.9), Inches(0.07), RED)
        self.text(sl, Inches(0.9), Inches(3.3), Inches(8), Inches(0.9), 'Thank you', 40, True, WHITE)
        self.text(sl, Inches(0.9), Inches(4.2), Inches(8), Inches(0.9),
                  'Generated from live RFAI Tracker data on ' + datetime.now().strftime('%d %b %Y, %H:%M'), 15, color=RGBColor(0xB8, 0xBD, 0xC8))


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
        hi = max(pairs, key=lambda p: p[1]); first, last = pairs[0][1], pairs[-1][1]
        out.append(f'Peak in {hi[0]} at {_fmt(hi[1])}.')
        if first: out.append(f'{pairs[-1][0]} is {"up" if last >= first else "down"} {abs(last - first) / first:.0%} vs {pairs[0][0]}.')
        out.append(f'Average {_fmt(round(sum(vals) / len(vals), 1))} per period; {_fmt(sum(vals))} in total.')
        return out
    srt = sorted(pairs, key=lambda p: p[1], reverse=True)
    out.append(f'{srt[0][0]} is highest with {_fmt(srt[0][1])} ({srt[0][1] / tot:.0%} of total).')
    if len(srt) >= 3:
        out.append(f'Top 3 ({", ".join(p[0] for p in srt[:3])}) make up {sum(p[1] for p in srt[:3]) / tot:.0%}.')
    if len(srt) >= 2:
        out.append(f'{srt[-1][0]} is lowest with {_fmt(srt[-1][1])}.')
    if len(ser) > 1:
        sums = [(x['name'], sum((_num(v) or 0) for v in x['values'])) for x in ser]
        big = max(sums, key=lambda p: p[1]); out.append(f'{big[0]} is the largest part: {_fmt(big[1])} ({big[1] / tot:.0%}).')
    return out


def build_deck(path, title, slides):
    d = Deck(title)
    slides = list(slides or [])[:40]
    if not slides or slides[0].get('kind') != 'title':
        slides.insert(0, {'kind': 'title', 'title': title})
    for s in slides:
        k = s.get('kind', 'bullets')
        {'title': d.cover, 'kpis': d.kpis, 'chart': d.chart, 'table': d.table,
         'bullets': d.bullets, 'section': d.section}.get(k, d.bullets)(s)
        if s.get('notes') and d.prs.slides:
            d.prs.slides[-1].notes_slide.notes_text_frame.text = s['notes']
    d.closing()
    d.prs.save(path)
