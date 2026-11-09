"""deckkit: small shared toolkit for building clean 16:9 decks with python-pptx.

Features: theme loading, text measurement (Carlito = Calibri metrics), auto-fit text
(shrink, or raise FitError so the caller can split/shorten), native charts, native tables,
RAG dots, speaker notes. No network, no external assets.
"""
import json
import os
import re
import sys

from lxml import etree
from PIL import ImageFont
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

SLIDE_W, SLIDE_H = 13.333, 7.5
MARGIN = 0.6
CONTENT_TOP = 1.75
CONTENT_BOTTOM = 6.85
CONTENT_W = SLIDE_W - 2 * MARGIN
LINE = 1.2  # line height factor used for measuring (Calibri is ~1.22)

PLACEHOLDER_RE = re.compile(r"\[(TBD|TODO|INSERT|FILL|XX|\?)[^\]]*\]|lorem ipsum|\bTBD\b|\bXXX\b", re.I)


class FitError(Exception):
    pass


# ---------------------------------------------------------------- themes
def load_theme(path, name=None):
    with open(path, encoding="utf-8") as f:
        themes = json.load(f)
    name = name or next(iter(themes))
    if name not in themes:
        raise SystemExit("Unknown theme %r. Available: %s" % (name, ", ".join(themes)))
    t = dict(themes[name])
    t["name"] = name
    return t


def rgb(hexstr):
    return RGBColor.from_string(hexstr.upper())


def mix(hex_a, hex_b, t):
    """Blend hex_a toward hex_b by t (0..1)."""
    a = [int(hex_a[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(hex_b[i:i + 2], 16) for i in (0, 2, 4)]
    return "".join("%02X" % round(a[i] + (b[i] - a[i]) * t) for i in range(3))


# ---------------------------------------------------------------- measuring
_FONT_CACHE = {}
_CANDIDATES = {
    False: ["/usr/share/fonts/truetype/crosextra/Carlito-Regular.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"],
    True: ["/usr/share/fonts/truetype/crosextra/Carlito-Bold.ttf",
           "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
           "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"],
}


def _font(bold):
    if bold not in _FONT_CACHE:
        for p in _CANDIDATES[bool(bold)]:
            if os.path.exists(p):
                _FONT_CACHE[bold] = ImageFont.truetype(p, 100)
                break
        else:
            _FONT_CACHE[bold] = ImageFont.load_default()
    return _FONT_CACHE[bold]


def text_width_pt(s, pt, bold=False):
    return _font(bold).getlength(s) * pt / 100.0


def count_lines(text, pt, width_pt, bold=False):
    """Greedy word wrap line count."""
    total = 0
    for raw in str(text).split("\n"):
        words = raw.split(" ")
        line = ""
        n = 1
        for w in words:
            trial = (line + " " + w).strip() if line else w
            if text_width_pt(trial, pt, bold) <= width_pt:
                line = trial
            else:
                if line:
                    n += 1
                # a single very long word wraps by characters
                wl = text_width_pt(w, pt, bold)
                if wl > width_pt:
                    n += int(wl // width_pt)
                    line = ""
                else:
                    line = w
        total += n
    return total


def _norm(item, default_bold, default_color):
    if isinstance(item, str):
        item = {"text": item}
    d = {"text": "", "bold": default_bold, "color": default_color, "bullet": False,
         "scale": 1.0, "gap": None, "italic": False}
    d.update(item)
    return d


BULLET_INDENT_IN = 0.2


def _height_pt(paras, size, width_in, gap_pt, spacing):
    width_pt = width_in * 72 * 0.97  # safety margin for renderer differences
    h = 0.0
    for i, p in enumerate(paras):
        w = width_pt - (BULLET_INDENT_IN * 72 if p["bullet"] else 0)
        pt = size * p["scale"]
        h += count_lines(p["text"], pt, w, p["bold"]) * pt * LINE * spacing
        if i:
            h += gap_pt if p["gap"] is None else p["gap"]
    return h


def fit_size(paras, w_in, h_in, max_pt, min_pt, gap_pt=6, spacing=1.0):
    """Largest size (0.5pt steps) at which paras fit; None if even min_pt overflows."""
    s = float(max_pt)
    while s >= min_pt - 1e-6:
        if _height_pt(paras, s, w_in, gap_pt, spacing) <= h_in * 72:
            return s
        s -= 0.5
    return None


# ---------------------------------------------------------------- primitives
def new_prs():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    return prs


def blank_slide(prs, theme, bg=None):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    fill = s.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(bg or theme["bg"])
    return s


def rect(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE, radius=None, line_w=0.75):
    sp = slide.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sp.shadow.inherit = False
    for er in sp._element.iter(qn("a:effectRef")):
        er.set("idx", "0")  # LibreOffice ignores the empty effectLst; kill theme shadow
    if fill:
        sp.fill.solid()
        sp.fill.fore_color.rgb = rgb(fill)
    else:
        sp.fill.background()
    if line:
        sp.line.color.rgb = rgb(line)
        sp.line.width = Pt(line_w)
    else:
        sp.line.fill.background()
    if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sp.adjustments[0] = radius
    return sp


def oval(slide, x, y, w, h, fill=None, line=None, line_w=1.0):
    return rect(slide, x, y, w, h, fill, line, MSO_SHAPE.OVAL, line_w=line_w)


def card(slide, x, y, w, h, theme, fill=None, line=None):
    return rect(slide, x, y, w, h, fill or theme["surface"], line, MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)


def _set_bullet(p, indent_in):
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(Inches(indent_in))))
    pPr.set("indent", str(-int(Inches(indent_in))))
    bf = etree.SubElement(pPr, qn("a:buFont"))
    bf.set("typeface", "Arial")
    bc = etree.SubElement(pPr, qn("a:buChar"))
    bc.set("char", "•")


def add_text(slide, x, y, w, h, paras, theme, max_pt=18, min_pt=10, bold=False, color=None,
             align="l", anchor="t", gap_pt=6, spacing=1.0, size=None, what="text"):
    """Add an auto-fitting text box. paras: str | list[str | dict]. Returns (shape, size_pt).
    Shrinks from max_pt down to min_pt; raises FitError if it still overflows."""
    if isinstance(paras, (str, dict)):
        paras = [paras]
    color = color or theme["ink"]
    norm = [_norm(p, bold, color) for p in paras]
    if size is None:
        size = fit_size(norm, w, h, max_pt, min_pt, gap_pt, spacing)
        if size is None:
            snippet = " | ".join(p["text"] for p in norm)[:90]
            raise FitError("%s does not fit in %.1fx%.1f in at %spt: '%s...' -> shorten the text"
                           % (what, w, h, min_pt, snippet))
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    for i, p in enumerate(norm):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[align]
        if spacing != 1.0:
            para.line_spacing = spacing
        if i:
            para.space_before = Pt(gap_pt if p["gap"] is None else p["gap"])
        if p["bullet"]:
            _set_bullet(para, BULLET_INDENT_IN)
        r = para.add_run()
        r.text = p["text"]
        f = r.font
        f.size = Pt(size * p["scale"])
        f.bold = bool(p["bold"])
        f.italic = bool(p["italic"])
        f.name = theme["font"]
        f.color.rgb = rgb(p["color"])
    return tb, size


def notes(slide, text):
    if text:
        slide.notes_slide.notes_text_frame.text = text


# ---------------------------------------------------------------- slide chrome
def title_block(slide, theme, title, kicker=None, dark=False):
    ink = theme["primary_ink"] if dark else theme["ink"]
    rect(slide, MARGIN, 0.5, 0.55, 0.07, theme["accent"])
    if kicker:
        add_text(slide, MARGIN + 0.7, 0.4, 6, 0.28, kicker.upper(), theme, max_pt=11, min_pt=9, bold=True,
                 color=theme["accent"], what="kicker")
    add_text(slide, MARGIN, 0.78, CONTENT_W, 0.85, title, theme, max_pt=30, min_pt=20, bold=True,
             color=ink, anchor="t", spacing=0.95, what="slide title")


def footer(slide, theme, left, page, dark=False):
    col = mix(theme["primary_ink"], theme["primary"], 0.35) if dark else theme["muted"]
    add_text(slide, MARGIN, 7.08, 8, 0.22, left, theme, size=9, color=col, what="footer")
    add_text(slide, SLIDE_W - MARGIN - 1.5, 7.08, 1.5, 0.22, str(page), theme, size=9, color=col, align="r")
    rect(slide, MARGIN, 7.0, CONTENT_W, 0.01, mix(theme["line"], theme["bg"], 0.0) if not dark else col)


def content_slide(prs, theme, title, page, footer_text, kicker=None, note=None):
    s = blank_slide(prs, theme)
    title_block(s, theme, title, kicker)
    footer(s, theme, footer_text, page)
    notes(s, note)
    return s


# ---------------------------------------------------------------- stat / dot helpers
def big_stat(slide, x, y, w, value, label, theme, color=None, label_color=None, value_pt=40, align="l"):
    """Big number with a small label under it. Returns height used (in)."""
    add_text(slide, x, y, w, value_pt / 72 * 1.25, value, theme, size=value_pt, bold=True,
             color=color or theme["accent"], align=align, what="stat value")
    lh = 0.55
    add_text(slide, x, y + value_pt / 72 * 1.25, w, lh, label, theme, max_pt=13, min_pt=9,
             color=label_color or theme["muted"], align=align, what="stat label")
    return value_pt / 72 * 1.25 + lh


def rag_color(theme, status):
    s = str(status).lower()
    return theme[{"g": "good", "green": "good", "a": "warn", "amber": "warn", "yellow": "warn",
                  "r": "bad", "red": "bad"}[s]]


def rag_dot(slide, x, y, d, theme, status):
    return oval(slide, x, y, d, d, fill=rag_color(theme, status))


# ---------------------------------------------------------------- tables
_NO_STYLE = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}"


def _cell_info(c, default_color, default_bold):
    if isinstance(c, dict):
        return c.get("text", ""), c.get("color", default_color), c.get("bold", default_bold), c.get("fill")
    return str(c), default_color, default_bold, None


def _table_heights(header, rows, widths_in, size, min_row_h):
    pad = 0.09 * 72  # top+bottom cell margins (pt)
    out = []
    for ri, row in enumerate([header] + rows):
        mx = 1
        for ci, c in enumerate(row):
            t, _, b, _ = _cell_info(c, None, ri == 0)
            lines = count_lines(t, size, (widths_in[ci] - 0.2) * 72 * 0.97, b or ri == 0)
            mx = max(mx, lines)
        out.append(max(min_row_h, (mx * size * LINE + pad) / 72))
    return out


def plan_table(header, rows, widths_in, avail_h, max_pt=13, min_pt=10, min_row_h=0.38):
    """Return (size, [row_groups]). Shrinks first, then splits rows over several groups."""
    s = float(max_pt)
    while s >= min_pt - 1e-6:
        if sum(_table_heights(header, rows, widths_in, s, min_row_h)) <= avail_h:
            return s, [rows]
        s -= 0.5
    groups, cur = [], []
    for r in rows:
        trial = cur + [r]
        if sum(_table_heights(header, trial, widths_in, min_pt, min_row_h)) <= avail_h:
            cur = trial
        else:
            if not cur:
                raise FitError("A single table row does not fit; shorten it: %r" % (r,))
            groups.append(cur)
            cur = [r]
    if cur:
        groups.append(cur)
    return min_pt, groups


def add_table(slide, x, y, widths_in, header, rows, theme, size, aligns=None, min_row_h=0.38,
              header_fill=None, zebra=True):
    """Native table with fixed font size (use plan_table first). Returns bottom y (in)."""
    n_rows, n_cols = len(rows) + 1, len(widths_in)
    heights = _table_heights(header, rows, widths_in, size, min_row_h)
    gf = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y), Inches(sum(widths_in)),
                                Inches(sum(heights)))
    tbl = gf.table
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    sid = tblPr.find(qn("a:tableStyleId"))
    if sid is None:
        sid = etree.SubElement(tblPr, qn("a:tableStyleId"))
    sid.text = _NO_STYLE
    for ci, wv in enumerate(widths_in):
        tbl.columns[ci].width = Inches(wv)
    aligns = aligns or ["l"] * n_cols
    for ri, row in enumerate([header] + rows):
        tbl.rows[ri].height = Inches(heights[ri])
        for ci in range(n_cols):
            cell = tbl.cell(ri, ci)
            hdr = ri == 0
            txt, col, bold, fill = _cell_info(row[ci], theme["primary_ink"] if hdr else theme["ink"], hdr)
            cell.margin_left = cell.margin_right = Inches(0.1)
            cell.margin_top = cell.margin_bottom = Inches(0.045)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if hdr:
                cell.fill.fore_color.rgb = rgb(header_fill or theme["primary"])
            elif fill:
                cell.fill.fore_color.rgb = rgb(fill)
            elif zebra and ri % 2 == 0:
                cell.fill.fore_color.rgb = rgb(theme["surface"])
            else:
                cell.fill.fore_color.rgb = rgb(theme["bg"])
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[aligns[ci]]
            r = p.add_run()
            r.text = txt
            r.font.size = Pt(size)
            r.font.bold = bool(bold)
            r.font.name = theme["font"]
            r.font.color.rgb = rgb(col)
    return y + sum(heights)


# ---------------------------------------------------------------- charts
def _style_axis_text(axis, theme, size):
    axis.tick_labels.font.size = Pt(size)
    axis.tick_labels.font.color.rgb = rgb(theme["muted"])
    axis.tick_labels.font.name = theme["font"]


def add_chart(slide, kind, x, y, w, h, categories, series, theme, number_format="General",
              legend=None, labels=True, colors=None, font_pt=11, stacked=False, gap=60,
              label_fmt=None, axis_visible=True, point_colors=None, label_pt=None, hole=58):
    """Native PowerPoint chart. kind: column | bar | line | doughnut | pie.
    series: list of (name, [values]). Returns the chart."""
    cd = CategoryChartData()
    cd.categories = categories
    for name, vals in series:
        cd.add_series(name, vals)
    ct = {
        "column": XL_CHART_TYPE.COLUMN_STACKED if stacked else XL_CHART_TYPE.COLUMN_CLUSTERED,
        "bar": XL_CHART_TYPE.BAR_STACKED if stacked else XL_CHART_TYPE.BAR_CLUSTERED,
        "line": XL_CHART_TYPE.LINE_MARKERS,
        "doughnut": XL_CHART_TYPE.DOUGHNUT,
        "pie": XL_CHART_TYPE.PIE,
    }[kind]
    gf = slide.shapes.add_chart(ct, Inches(x), Inches(y), Inches(w), Inches(h), cd)
    ch = gf.chart
    ch.font.size = Pt(font_pt)
    ch.font.name = theme["font"]
    ch.font.color.rgb = rgb(theme["muted"])
    ch.has_title = False
    colors = colors or theme["series"]
    if legend is None:
        legend = len(series) > 1 or kind in ("doughnut", "pie")
    ch.has_legend = bool(legend)
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.RIGHT if kind in ("doughnut", "pie") else XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(font_pt)
        ch.legend.font.color.rgb = rgb(theme["ink"])
    plot = ch.plots[0]
    if kind in ("column", "bar"):
        plot.gap_width = gap
        plot.overlap = 100 if stacked else -10
        va, ca = ch.value_axis, ch.category_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = rgb(theme["line"])
        va.major_gridlines.format.line.width = Pt(0.5)
        va.format.line.fill.background()
        va.tick_labels.number_format = number_format
        va.tick_labels.number_format_is_linked = False
        va.visible = axis_visible
        ca.format.line.color.rgb = rgb(theme["line"])
        ca.has_major_gridlines = False
        _style_axis_text(va, theme, font_pt)
        _style_axis_text(ca, theme, font_pt)
        for i, s in enumerate(plot.series):
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = rgb(colors[i % len(colors)])
            s.invert_if_negative = False
        if point_colors:
            for i, c in enumerate(point_colors):
                pt_ = plot.series[0].points[i]
                pt_.format.fill.solid()
                pt_.format.fill.fore_color.rgb = rgb(c)
    elif kind == "line":
        va, ca = ch.value_axis, ch.category_axis
        va.has_major_gridlines = True
        va.major_gridlines.format.line.color.rgb = rgb(theme["line"])
        va.major_gridlines.format.line.width = Pt(0.5)
        va.format.line.fill.background()
        va.tick_labels.number_format = number_format
        va.tick_labels.number_format_is_linked = False
        va.visible = axis_visible
        ca.format.line.color.rgb = rgb(theme["line"])
        _style_axis_text(va, theme, font_pt)
        _style_axis_text(ca, theme, font_pt)
        for i, s in enumerate(plot.series):
            c = rgb(colors[i % len(colors)])
            s.format.line.color.rgb = c
            s.format.line.width = Pt(2.5)
            s.smooth = False
            s.marker.style = XL_MARKER_STYLE.CIRCLE
            s.marker.size = 7
            s.marker.format.fill.solid()
            s.marker.format.fill.fore_color.rgb = c
            s.marker.format.line.color.rgb = c
    else:  # doughnut / pie
        pc = point_colors or colors
        for i in range(len(categories)):
            pt_ = plot.series[0].points[i]
            pt_.format.fill.solid()
            pt_.format.fill.fore_color.rgb = rgb(pc[i % len(pc)])
            pt_.format.line.color.rgb = rgb(theme["bg"])
            pt_.format.line.width = Pt(1.5)
        if kind == "doughnut":
            # set hole size
            dn = plot._element
            hs = dn.find(qn("c:holeSize"))
            if hs is None:
                hs = etree.SubElement(dn, qn("c:holeSize"))
            hs.set("val", str(hole))
    if labels:
        plot.has_data_labels = True
        dl = plot.data_labels
        dl.font.size = Pt(label_pt or font_pt)
        dl.font.bold = True
        dl.number_format = label_fmt or number_format
        dl.number_format_is_linked = False
        if kind in ("column", "bar"):
            dl.position = XL_LABEL_POSITION.OUTSIDE_END if not stacked else XL_LABEL_POSITION.CENTER
            dl.font.color.rgb = rgb(theme["ink"] if not stacked else theme["primary_ink"])
        elif kind == "line":
            dl.position = XL_LABEL_POSITION.ABOVE
            dl.font.color.rgb = rgb(theme["ink"])
        else:
            dl.font.color.rgb = rgb(theme["primary_ink"])
            dl.show_percentage = False
            dl.show_value = True
    return ch


# ---------------------------------------------------------------- output
def save(prs, path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    prs.save(path)
    print("saved", path, "(%d slides)" % len(prs.slides))


def find_placeholders(strings):
    return [s for s in strings if isinstance(s, str) and PLACEHOLDER_RE.search(s)]


def walk_strings(obj):
    if isinstance(obj, str):
        yield obj
    elif isinstance(obj, dict):
        for v in obj.values():
            yield from walk_strings(v)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            yield from walk_strings(v)


def money(v, decimals=None):
    """1_250_000 -> $1.25M ; 18_000 -> $18K ; 950 -> $950 ; 2.1e9 -> $2.1B"""
    a = abs(v)
    sign = "-" if v < 0 else ""

    def fmt(x, d):
        return ("%.*f" % (d, x)).rstrip("0").rstrip(".") if d else "%d" % round(x)

    if a >= 1e9:
        return "%s$%sB" % (sign, fmt(a / 1e9, 2 if decimals is None else decimals))
    if a >= 1e6:
        return "%s$%sM" % (sign, fmt(a / 1e6, 2 if decimals is None else decimals))
    if a >= 1e3:
        return "%s$%sK" % (sign, fmt(a / 1e3, 0 if decimals is None else decimals))
    return "%s$%d" % (sign, round(a))
