#!/usr/bin/env python3
"""Build a multi-page Pyramid-Principle management report (PDF) from a CSV and a notes file.

Usage:
  python build_report.py --csv data.csv --notes notes.md --out report.pdf

CSV: first column `month` (YYYY-MM), then one column per metric key (wide format, one row per month).
Notes: see SKILL.md / examples/input.md. Commentary text may use {key}, {key.prev}, {key.mom},
{key.target}, {key.vs_target}, {period_label} placeholders so numbers always match the data.
"""
import argparse
import csv
import datetime as dt
import io
import re
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (BaseDocTemplate, Flowable, Frame, Image, KeepTogether, PageBreak,
                                PageTemplate, Paragraph, Spacer, Table, TableStyle)

NAVY = colors.HexColor("#1F2A44")
TEAL = colors.HexColor("#2A7F8E")
GREY = colors.HexColor("#6B7280")
LIGHT = colors.HexColor("#F3F5F9")
RULE = colors.HexColor("#D5DAE3")
GOOD = colors.HexColor("#1E7D32")
BAD = colors.HexColor("#C0392B")
M_NAVY, M_TEAL, M_GREY, M_BAD = "#1F2A44", "#2A7F8E", "#6B7280", "#C0392B"

PAGE_W, PAGE_H = A4
MARGIN = 40
CONTENT_W = PAGE_W - 2 * MARGIN


# ------------------------------------------------------------------ parsing
class Metric:
    def __init__(self, key, label, fmt, direction, target, kind):
        self.key, self.label, self.fmt, self.dir, self.target, self.kind = key, label, fmt, direction, target, kind
        self.values = []  # aligned with months


def parse_notes(path):
    text = open(path, encoding="utf-8").read()
    head, _, rest = text.partition("\n## ")
    meta = {}
    for line in head.splitlines():
        if ":" in line and not line.startswith("#"):
            k, v = line.split(":", 1)
            meta[k.strip().lower()] = v.strip()
    sections, cur = {}, None
    for line in ("## " + rest).splitlines() if rest else []:
        if line.startswith("## "):
            cur = line[3:].strip().lower()
            sections[cur] = []
        elif cur:
            sections[cur].append(line)
    return meta, {k: "\n".join(v).strip() for k, v in sections.items()}


def parse_metrics(block):
    out = []
    for line in block.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        p = [x.strip() for x in line.split("|")]
        if len(p) < 4:
            sys.exit(f"Bad metric line (need key | label | fmt | up/down | [target] | [bar/line]): {line}")
        key, label, fmt, direction = p[:4]
        target = float(p[4]) if len(p) > 4 and p[4] else None
        kind = p[5] if len(p) > 5 and p[5] else "line"
        if fmt not in ("currency", "percent", "number") or direction not in ("up", "down"):
            sys.exit(f"Bad fmt/direction in: {line}")
        out.append(Metric(key, label, fmt, direction, target, kind))
    return out


def load_csv(path, metrics):
    with open(path, newline="", encoding="utf-8-sig") as fh:
        rd = csv.DictReader(fh)
        rows = list(rd)
    if not rows or "month" not in rows[0]:
        sys.exit("CSV needs a 'month' column (YYYY-MM).")
    rows.sort(key=lambda r: r["month"])
    months = [r["month"] for r in rows]
    for m in metrics:
        if m.key not in rows[0]:
            sys.exit(f"CSV has no column '{m.key}' (columns: {', '.join(rows[0].keys())})")
        vals = []
        for r in rows:
            try:
                vals.append(float(r[m.key]))
            except ValueError:
                sys.exit(f"Non-numeric value '{r[m.key]}' in column {m.key}, month {r['month']}")
        m.values = vals
    return months


# ------------------------------------------------------------------ formatting
def compact(v, cur):
    a = abs(v)
    sign = "-" if v < 0 else ""
    if a >= 1e6:
        return f"{sign}{cur}{a/1e6:.2f}M"
    if a >= 1e4:
        return f"{sign}{cur}{a/1e3:.0f}k"
    if a >= 1e3:
        return f"{sign}{cur}{a/1e3:.1f}k"
    return f"{sign}{cur}{a:,.0f}"


def fmt_val(m, v, cur):
    if m.fmt == "currency":
        return compact(v, cur)
    if m.fmt == "percent":
        return f"{v:.1f}%"
    return f"{v:,.0f}"


def change(m, cur_v, base_v):
    """Return (text, signed_delta). Percent metrics -> pp, others -> relative %."""
    if m.fmt == "percent":
        d = cur_v - base_v
        return f"{d:+.1f} pp", d
    if base_v == 0:
        return "n/a", 0.0
    d = cur_v / base_v - 1
    return f"{d*100:+.1f}%", d


def favourable(m, d, tol=1e-9):
    if abs(d) < tol:
        return None
    return (d > 0) == (m.dir == "up")


def month_label(ym, short=False):
    d = dt.datetime.strptime(ym, "%Y-%m")
    return d.strftime("%b %y" if short else "%B %Y")


def expand(text, ctx, metrics, idx, months, cur):
    by = {m.key: m for m in metrics}

    def sub(mo):
        tok = mo.group(1)
        if tok == "period_label":
            return month_label(months[idx])
        if tok == "period":
            return months[idx]
        key, _, attr = tok.partition(".")
        if key not in by:
            sys.exit(f"Unknown placeholder {{{tok}}} in notes (metrics: {', '.join(by)})")
        m = by[key]
        v = m.values[idx]
        pv = m.values[idx - 1] if idx > 0 else None
        if attr == "":
            return fmt_val(m, v, cur)
        if attr == "prev":
            return fmt_val(m, pv, cur) if pv is not None else "n/a"
        if attr == "mom":
            return change(m, v, pv)[0] if pv is not None else "n/a"
        if attr == "target":
            return fmt_val(m, m.target, cur) if m.target is not None else "n/a"
        if attr == "vs_target":
            return change(m, v, m.target)[0] if m.target is not None else "n/a"
        sys.exit(f"Unknown placeholder attribute in {{{tok}}}")
    return re.sub(r"\{([\w.]+)\}", sub, text)


# ------------------------------------------------------------------ charts
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 8, "axes.edgecolor": "#9CA3AF",
                     "axes.labelcolor": M_GREY, "xtick.color": M_GREY, "ytick.color": M_GREY})


def chart_png(m, months, idx, cur, w_in, h_in, show_title=True, dpi=200):
    n = min(12, idx + 1)
    xs = list(range(n))
    vals = m.values[idx - n + 1: idx + 1]
    labs = [month_label(x, True) for x in months[idx - n + 1: idx + 1]]
    fig, ax = plt.subplots(figsize=(w_in, h_in))
    if m.kind == "bar":
        ax.bar(xs, vals, color="#C5CEDD", width=0.68)
        ax.bar(xs[-1], vals[-1], color=M_NAVY, width=0.68)
    else:
        ax.plot(xs, vals, color=M_NAVY, lw=1.8, solid_capstyle="round")
        ax.fill_between(xs, vals, min(vals) - (max(vals) - min(vals)) * 0.15, color=M_NAVY, alpha=0.06)
        ax.plot(xs[-1], vals[-1], "o", color=M_NAVY, ms=5)
    lo, hi = min(vals), max(vals)
    if m.target is not None:
        lo, hi = min(lo, m.target), max(hi, m.target)
        ax.plot(xs[-1] + 0.45, m.target, marker="D", color=M_TEAL, ms=5, clip_on=False, zorder=5)
        ax.axhline(m.target, color=M_TEAL, lw=0.8, ls=(0, (4, 3)), alpha=0.8)
    pad = (hi - lo) * 0.18 or 1
    ymin = 0 if (m.kind == "bar") else lo - pad
    ax.set_ylim(ymin, hi + pad)
    ax.set_xlim(-0.6, n - 0.4)
    step = 1 if n <= 6 else (3 if w_in < 3.5 else 2)
    ax.set_xticks(xs[::step][::-1][::-1])
    ax.set_xticklabels(labs[::step], fontsize=7)
    ax.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: fmt_val(m, v, cur) if m.fmt != "percent" else f"{v:.1f}%"))
    ax.tick_params(axis="both", length=0)
    ax.grid(axis="y", color="#E5E7EB", lw=0.6)
    ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.annotate(fmt_val(m, vals[-1], cur), (xs[-1], vals[-1]), textcoords="offset points", xytext=(0, 7),
                ha="center", fontsize=8, fontweight="bold", color=M_NAVY)
    if show_title:
        ax.set_title(m.label, loc="left", fontsize=9, fontweight="bold", color=M_NAVY, pad=8)
    fig.tight_layout(pad=0.6)
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi)
    plt.close(fig)
    buf.seek(0)
    return buf


# ------------------------------------------------------------------ flowables
class Tile(Flowable):
    W, H = 165, 96

    def __init__(self, m, idx, months, cur):
        super().__init__()
        self.m, self.idx, self.months, self.cur = m, idx, months, cur
        self.width, self.height = self.W, self.H

    def draw(self):
        c, m, i = self.canv, self.m, self.idx
        v = m.values[i]
        c.setFillColor(LIGHT)
        c.setStrokeColor(RULE)
        c.roundRect(0, 0, self.W, self.H, 5, fill=1, stroke=1)
        # status: vs target if there is one, else vs last month
        tgt_ok = None
        if m.target is not None:
            _, d = change(m, v, m.target)
            tgt_ok = favourable(m, d)
        accent = GREY if tgt_ok is None else (GOOD if tgt_ok else BAD)
        c.setFillColor(accent)
        c.rect(0, 0, 4, self.H, fill=1, stroke=0)
        c.setFillColor(GREY)
        c.setFont("Helvetica-Bold", 7.2)
        c.drawString(14, self.H - 16, m.label.upper())
        c.setFillColor(NAVY)
        c.setFont("Helvetica-Bold", 23)
        c.drawString(14, self.H - 44, fmt_val(m, v, self.cur))
        # vs last month
        if i > 0:
            txt, d = change(m, v, m.values[i - 1])
            fav = favourable(m, d)
            col = GREY if fav is None else (GOOD if fav else BAD)
            self._arrow(14, self.H - 61, d, col)
            c.setFillColor(col)
            c.setFont("Helvetica-Bold", 8.5)
            c.drawString(26, self.H - 64, txt)
            c.setFillColor(GREY)
            c.setFont("Helvetica", 7.5)
            c.drawString(26 + c.stringWidth(txt, "Helvetica-Bold", 8.5) + 4, self.H - 64, "vs last month")
        # vs target
        if m.target is not None:
            txt, d = change(m, v, m.target)
            fav = favourable(m, d)
            col = GREY if fav is None else (GOOD if fav else BAD)
            word = "ahead of" if fav else ("behind" if fav is False else "on")
            c.setFillColor(col)
            c.setFont("Helvetica-Bold", 8.5)
            c.drawString(14, self.H - 82, txt)
            c.setFillColor(GREY)
            c.setFont("Helvetica", 7.5)
            c.drawString(14 + c.stringWidth(txt, "Helvetica-Bold", 8.5) + 4, self.H - 82,
                         f"{word} target {fmt_val(m, m.target, self.cur)}")
        # sparkline
        n = min(12, i + 1)
        vals = m.values[i - n + 1: i + 1]
        lo, hi = min(vals), max(vals)
        x0, x1, y0, y1 = self.W - 62, self.W - 10, self.H - 40, self.H - 16
        pts = [(x0 + (x1 - x0) * k / max(1, n - 1), y0 + (y1 - y0) * ((vv - lo) / ((hi - lo) or 1))) for k, vv in enumerate(vals)]
        c.setStrokeColor(colors.HexColor("#9CA3AF"))
        c.setLineWidth(1.1)
        for a, b in zip(pts, pts[1:]):
            c.line(a[0], a[1], b[0], b[1])
        c.setFillColor(NAVY)
        c.circle(pts[-1][0], pts[-1][1], 2, fill=1, stroke=0)

    def _arrow(self, x, y, d, col):
        c = self.canv
        c.setFillColor(col)
        p = c.beginPath()
        if abs(d) < 1e-9:
            c.rect(x, y + 2, 8, 2, fill=1, stroke=0)
            return
        if d > 0:
            p.moveTo(x, y); p.lineTo(x + 8, y); p.lineTo(x + 4, y + 7)
        else:
            p.moveTo(x, y + 7); p.lineTo(x + 8, y + 7); p.lineTo(x + 4, y)
        p.close()
        c.drawPath(p, fill=1, stroke=0)


class NumberedCanvas(rl_canvas.Canvas):
    meta = {}

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self._saved = []

    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._saved)
        for st in self._saved:
            self.__dict__.update(st)
            self._decorate(total)
            super().showPage()
        super().save()

    def _decorate(self, total):
        meta = self.meta
        self.setStrokeColor(RULE)
        self.setLineWidth(0.6)
        self.line(MARGIN, PAGE_H - 34, PAGE_W - MARGIN, PAGE_H - 34)
        self.setFont("Helvetica-Bold", 8)
        self.setFillColor(NAVY)
        self.drawString(MARGIN, PAGE_H - 28, meta["company"])
        self.setFont("Helvetica", 8)
        self.setFillColor(GREY)
        self.drawRightString(PAGE_W - MARGIN, PAGE_H - 28, f"{meta['title']}  |  {meta['period_label']}")
        self.line(MARGIN, 34, PAGE_W - MARGIN, 34)
        self.drawString(MARGIN, 22, f"Prepared by {meta['prepared_by']}  |  Confidential")
        self.drawRightString(PAGE_W - MARGIN, 22, f"Page {self._pageNumber} of {total}")


# ------------------------------------------------------------------ build
def styles():
    base = ParagraphStyle("base", fontName="Helvetica", fontSize=9.5, leading=13.5, textColor=colors.HexColor("#1F2937"), alignment=TA_LEFT)
    return {
        "base": base,
        "kicker": ParagraphStyle("kicker", parent=base, fontName="Helvetica-Bold", fontSize=8, textColor=TEAL, leading=10, spaceAfter=3),
        "title": ParagraphStyle("title", parent=base, fontName="Helvetica-Bold", fontSize=22, leading=26, textColor=NAVY, spaceAfter=4),
        "headline": ParagraphStyle("headline", parent=base, fontName="Helvetica-Bold", fontSize=15, leading=20, textColor=NAVY),
        "h2": ParagraphStyle("h2", parent=base, fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=NAVY, spaceBefore=4, spaceAfter=4),
        "h3": ParagraphStyle("h3", parent=base, fontName="Helvetica-Bold", fontSize=10.5, leading=14, textColor=NAVY, spaceAfter=3),
        "small": ParagraphStyle("small", parent=base, fontSize=8, leading=11, textColor=GREY),
        "reason": ParagraphStyle("reason", parent=base, fontSize=9.5, leading=13),
    }


def parse_points(block, ctx_fn):
    pts, cur = [], None
    for line in block.splitlines():
        if line.startswith("### "):
            cur = {"title": line[4:].strip(), "metric": None, "body": []}
            pts.append(cur)
        elif cur is not None:
            mm = re.match(r"metric:\s*(\w+)", line.strip())
            if mm:
                cur["metric"] = mm.group(1)
            elif line.strip():
                cur["body"].append(line.strip())
    return pts


def build(csv_path, notes_path, out_path):
    meta, sections = parse_notes(notes_path)
    for req in ("company", "period", "currency"):
        if req not in meta:
            sys.exit(f"Notes file is missing '{req}:' at the top.")
    metrics = parse_metrics(sections.get("metrics", ""))
    if not (4 <= len(metrics) <= 6):
        sys.exit(f"Use 4-6 headline metrics (found {len(metrics)}).")
    months = load_csv(csv_path, metrics)
    if meta["period"] not in months:
        sys.exit(f"Period {meta['period']} not in CSV months ({months[0]}..{months[-1]}).")
    idx = months.index(meta["period"])
    cur = meta["currency"]
    meta.setdefault("title", "Monthly management report")
    meta.setdefault("prepared_by", "Management")
    meta["period_label"] = month_label(meta["period"])
    NumberedCanvas.meta = meta
    ex = lambda t: expand(t, None, metrics, idx, months, cur)
    S = styles()
    by = {m.key: m for m in metrics}

    doc = BaseDocTemplate(out_path, pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN, topMargin=50, bottomMargin=48,
                          title=f"{meta['company']} – {meta['title']} – {meta['period_label']}", author=meta["prepared_by"])
    doc.addPageTemplates([PageTemplate("p", frames=[Frame(MARGIN, 48, CONTENT_W, PAGE_H - 98, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)])])
    el = []

    # ---- page 1: answer first
    el.append(Paragraph(meta["period_label"].upper() + "  |  MANAGEMENT SUMMARY", S["kicker"]))
    el.append(Paragraph(meta["title"], S["title"]))
    headline = ex(sections.get("headline", "").strip())
    if not headline:
        sys.exit("Notes need a '## Headline' section (the answer, in one or two sentences).")
    box = Table([[Paragraph(headline, S["headline"])]], colWidths=[CONTENT_W])
    box.setStyle(TableStyle([("LINEBEFORE", (0, 0), (0, 0), 3, TEAL), ("LEFTPADDING", (0, 0), (-1, -1), 12),
                             ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    el += [Spacer(1, 6), box, Spacer(1, 10)]
    points = parse_points(sections.get("key points", ""), ex)
    if not 2 <= len(points) <= 4:
        sys.exit("Use 2-4 key points (### headings under '## Key points'); three is the pyramid sweet spot.")
    el.append(Paragraph("WHY: THE THREE THINGS THAT MATTER" if len(points) == 3 else "WHY: THE KEY POINTS", S["kicker"]))
    rows = []
    for k, p in enumerate(points, 1):
        num = Paragraph(f'<font color="#2A7F8E" size="15"><b>{k}</b></font>', S["base"])
        rows.append([num, Paragraph(ex(p["title"]), S["reason"])])
    t = Table(rows, colWidths=[22, CONTENT_W - 22])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LINEBELOW", (0, 0), (-1, -2), 0.4, RULE),
                           ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    el += [t, Spacer(1, 14)]
    el.append(Paragraph("KEY METRICS: " + meta["period_label"].upper(), S["kicker"]))
    tiles = [Tile(m, idx, months, cur) for m in metrics]
    grid = [tiles[i:i + 3] for i in range(0, len(tiles), 3)]
    for g in grid:
        while len(g) < 3:
            g.append("")
    tg = Table(grid, colWidths=[Tile.W + 10] * 3, rowHeights=[Tile.H + 10] * len(grid))
    tg.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                            ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 10), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    el += [tg]
    scored = [(m, favourable(m, change(m, m.values[idx], m.target)[1])) for m in metrics if m.target is not None]
    if scored:
        ok = sum(1 for _, f in scored if f is not False)
        el.append(Paragraph(f"<b>{ok} of {len(scored)}</b> metrics are on or ahead of target; "
                            f"behind: {', '.join(m.label for m, f in scored if f is False) or 'none'}.", S["reason"]))
    nxt = sections.get("decisions and next steps", "")
    if nxt:
        el.append(Spacer(1, 4))
        el.append(Paragraph("DECISIONS NEEDED", S["kicker"]))
        for line in nxt.splitlines():
            if line.strip().startswith("-"):
                parts = [x.strip() for x in line.strip()[1:].split("|")]
                txt = ex(parts[0]) + (f'  <font color="#6B7280">({parts[1]}, {parts[2]})</font>' if len(parts) >= 3 else "")
                el.append(Paragraph("•&nbsp;&nbsp;" + txt, S["reason"]))
    el.append(PageBreak())

    # ---- page 2: trends
    el.append(Paragraph("TRENDS", S["kicker"]))
    el.append(Paragraph("Twelve-month view of every headline metric", S["h2"]))
    el.append(Paragraph("Dark bar or dot = current month. Dashed teal line and diamond = target for the month.", S["small"]))
    el.append(Spacer(1, 8))
    imgs = []
    cw = (CONTENT_W - 12) / 2
    for m in metrics:
        buf = chart_png(m, months, idx, cur, cw / 72, 2.15)
        imgs.append(Image(buf, width=cw, height=2.15 * 72))
    rows = [imgs[i:i + 2] for i in range(0, len(imgs), 2)]
    for r_ in rows:
        while len(r_) < 2:
            r_.append("")
    tt = Table(rows, colWidths=[cw + 12, cw + 12])
    tt.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 12), ("BOTTOMPADDING", (0, 0), (-1, -1), 8), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    el += [tt, PageBreak()]

    # ---- page 3: commentary (supporting arguments)
    el.append(Paragraph("COMMENTARY", S["kicker"]))
    el.append(Paragraph("What is behind the numbers", S["h2"]))
    el.append(Spacer(1, 4))
    for k, p in enumerate(points, 1):
        left = [Paragraph(f'<font color="#2A7F8E">{k}.</font>&nbsp;{ex(p["title"])}', S["h3"])]
        for b in p["body"]:
            left.append(Paragraph(ex(b), S["base"]))
            left.append(Spacer(1, 3))
        right = ""
        if p["metric"]:
            if p["metric"] not in by:
                sys.exit(f"metric: {p['metric']} in key point {k} is not a defined metric")
            w = 190
            h = 118
            right = Image(chart_png(by[p["metric"]], months, idx, cur, w / 72, h / 72, show_title=True), width=w, height=h)
        blk = Table([[left, right]], colWidths=[CONTENT_W - 205, 205])
        blk.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (0, 0), 14),
                                 ("LINEBELOW", (0, 0), (-1, -1), 0.5, RULE), ("BOTTOMPADDING", (0, 0), (-1, -1), 10), ("TOPPADDING", (0, 0), (-1, -1), 8)]))
        el.append(KeepTogether(blk))
    risks = sections.get("risks", "")
    if risks:
        el.append(Spacer(1, 10))
        el.append(Paragraph("RISKS TO WATCH", S["kicker"]))
        rrows = [[Paragraph("Risk", S["small"]), Paragraph("Owner", S["small"]), Paragraph("Mitigation / timing", S["small"])]]
        for line in risks.splitlines():
            if line.strip().startswith("-"):
                parts = [x.strip() for x in line.strip()[1:].split("|")] + ["", ""]
                rrows.append([Paragraph(ex(parts[0]), S["base"]), Paragraph(parts[1], S["base"]), Paragraph(parts[2], S["base"])])
        rt = Table(rrows, colWidths=[CONTENT_W - 220, 80, 140], repeatRows=1)
        rt.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.4, RULE), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        el.append(rt)
    if nxt:
        el.append(Spacer(1, 10))
        el.append(Paragraph("NEXT STEPS", S["kicker"]))
        nrows = [[Paragraph("Action", S["small"]), Paragraph("Owner", S["small"]), Paragraph("Due", S["small"])]]
        for line in nxt.splitlines():
            if line.strip().startswith("-"):
                parts = [x.strip() for x in line.strip()[1:].split("|")] + ["", ""]
                nrows.append([Paragraph(ex(parts[0]), S["base"]), Paragraph(parts[1], S["base"]), Paragraph(parts[2], S["base"])])
        nt = Table(nrows, colWidths=[CONTENT_W - 220, 80, 140], repeatRows=1)
        nt.setStyle(TableStyle([("LINEBELOW", (0, 0), (-1, -1), 0.4, RULE), ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        el.append(nt)
    el.append(PageBreak())

    # ---- appendix: data table
    el.append(Paragraph("APPENDIX", S["kicker"]))
    el.append(Paragraph("Data: last 12 months", S["h2"]))
    n = min(12, idx + 1)
    start = idx - n + 1
    hdr = [Paragraph("<b>Month</b>", S["small"])] + [Paragraph(f"<b>{m.label}</b>", ParagraphStyle("hh", parent=S["small"], alignment=2)) for m in metrics]
    data = [hdr]
    for i in range(start, idx + 1):
        data.append([month_label(months[i], True)] + [fmt_val(m, m.values[i], cur) for m in metrics])
    data.append(["Target"] + [fmt_val(m, m.target, cur) if m.target is not None else "–" for m in metrics])
    vt = []
    for m in metrics:
        if m.target is None:
            vt.append("–")
        else:
            vt.append(change(m, m.values[idx], m.target)[0])
    data.append(["vs target"] + vt)
    cws = [60] + [(CONTENT_W - 60) / len(metrics)] * len(metrics)
    tbl = Table(data, colWidths=cws, repeatRows=1)
    st = [("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 1), (-1, -1), 8.5), ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
          ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#1F2937")), ("LINEBELOW", (0, 0), (-1, 0), 1, NAVY),
          ("LINEBELOW", (0, 1), (-1, -4), 0.3, RULE), ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
          ("BACKGROUND", (0, n), (-1, n), LIGHT), ("FONTNAME", (0, n), (-1, n), "Helvetica-Bold"),
          ("LINEABOVE", (0, n + 1), (-1, n + 1), 0.8, NAVY), ("FONTNAME", (0, n + 1), (-1, -1), "Helvetica-Bold"), ("LEFTPADDING", (0, 0), (-1, -1), 4)]
    for j, m in enumerate(metrics, 1):
        if m.target is not None:
            _, d = change(m, m.values[idx], m.target)
            f = favourable(m, d)
            st.append(("TEXTCOLOR", (j, n + 2), (j, n + 2), GREY if f is None else (GOOD if f else BAD)))
    tbl.setStyle(TableStyle(st))
    el += [tbl, Spacer(1, 12)]
    notes = [f"Source: {csv_path.split('/')[-1]}. Month-on-month change is relative for currency and counts, in percentage points (pp) for percentages.",
             "Colour shows favourable (green) or unfavourable (red) given each metric's direction; the arrow or sign always carries the same information.",
             "Metrics where lower is better: " + (", ".join(m.label for m in metrics if m.dir == "down") or "none") + "."]
    for ntext in notes:
        el.append(Paragraph(ntext, S["small"]))
    doc.build(el, canvasmaker=NumberedCanvas)
    return len(months), idx


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--csv", required=True)
    ap.add_argument("--notes", required=True)
    ap.add_argument("--out", default="report.pdf")
    a = ap.parse_args()
    build(a.csv, a.notes, a.out)
    print(f"Wrote {a.out}")


if __name__ == "__main__":
    main()
