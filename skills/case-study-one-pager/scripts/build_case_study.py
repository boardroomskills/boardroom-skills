#!/usr/bin/env python3
"""Build a one-page customer case study: PDF (reportlab) and a matching 1-slide PPTX.

usage: build_case_study.py --spec case.json [--theme boardroom|ember|forest]
                           [--pdf out.pdf] [--pptx out.pptx]
Text auto-fits (shrinks to a floor, then exits with code 2 telling you what to shorten).
Metrics can be given as {"value": "42%"} or computed: {"before": 18, "after": 10.4, "mode": "pct_change"}
or {"before": 5, "after": 12, "mode": "multiple"}.
"""
import argparse, json, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from xml.sax.saxutils import escape
import deckkit as K

HERE = os.path.dirname(os.path.abspath(__file__))
THEMES = os.path.join(HERE, "..", "templates", "themes.json")
W, H = 612.0, 792.0
MX = 44.0


class FitError(Exception):
    pass


# ---------------------------------------------------------------- fonts
def register_fonts():
    base = "/usr/share/fonts/truetype/crosextra/"
    try:
        for n, f in (("Body", "Carlito-Regular.ttf"), ("Body-Bold", "Carlito-Bold.ttf"), ("Body-Italic", "Carlito-Italic.ttf")):
            pdfmetrics.registerFont(TTFont(n, base + f))
        return "Body", "Body-Bold", "Body-Italic"
    except Exception:
        return "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"


REG, BOLD, ITAL = register_fonts()


# ---------------------------------------------------------------- derived values
def metric_value(m):
    if "value" in m:
        return m["value"]
    b, a = float(m["before"]), float(m["after"])
    if m.get("mode") == "multiple":
        r = a / b
        return ("%dx" % r) if abs(r - round(r)) < 0.05 else "%.1fx" % r
    ch = (a - b) / b * 100
    return "%s%d%%" % ("+" if ch > 0 else "−", round(abs(ch)))


# ---------------------------------------------------------------- PDF helpers
def fit_paras(items, w, h, max_pt, min_pt, font=REG, color="#000000", bold=False, leading=1.22, gap=0.45, align=TA_LEFT):
    """items: list of str or (str, bullet:bool). Returns (list[Paragraph], size, total_height)."""
    s = max_pt
    while s >= min_pt - 1e-6:
        ps, tot = [], 0.0
        for i, it in enumerate(items):
            text, bullet = it if isinstance(it, tuple) else (it, False)
            st = ParagraphStyle("p", fontName=BOLD if bold else font, fontSize=s, leading=s * leading, textColor=HexColor(color),
                                alignment=align, bulletFontName=REG, bulletFontSize=s, leftIndent=12 if bullet else 0,
                                bulletIndent=0, spaceBefore=(s * gap if i else 0))
            p = Paragraph(escape(text), st, bulletText="•" if bullet else None)
            _, ph = p.wrap(w, 10000)
            tot += ph + (st.spaceBefore if i else 0)
            ps.append(p)
        if tot <= h:
            return ps, s, tot
        s -= 0.5
    raise FitError("text does not fit in %.0fx%.0f pt at %spt: '%s...'" % (w, h, min_pt, str(items[0])[:60]))


def draw_paras(c, ps, x, y_top, w):
    y = y_top
    for p in ps:
        _, ph = p.wrap(w, 10000)
        y -= p.style.spaceBefore
        p.drawOn(c, x, y - ph)
        y -= ph
    return y


def text_block(c, items, x, y_top, w, h, max_pt, min_pt, **kw):
    ps, s, tot = fit_paras(items, w, h, max_pt, min_pt, **kw)
    draw_paras(c, ps, x, y_top, w)
    return s, tot


def build_pdf(spec, th, d, path):
    c = canvas.Canvas(path, pagesize=(W, H))
    c.setTitle("%s customer story" % spec["customer"])
    c.setAuthor(spec["vendor"])
    P, PI, AC, A2, INK, MUT, SUR = (HexColor("#" + th[k]) for k in ("primary", "primary_ink", "accent", "accent2", "ink", "muted", "surface"))
    # header band
    c.setFillColor(P); c.rect(0, 622, W, 170, stroke=0, fill=1)
    c.setFillColor(AC); c.rect(MX, 758, 36, 4, stroke=0, fill=1)
    c.setFillColor(A2); c.setFont(BOLD, 9.5)
    c.drawString(MX + 46, 755, ("CUSTOMER STORY  |  " + spec["customer"]).upper())
    text_block(c, [spec["headline"]], MX, 742, W - 2 * MX, 70, 29, 19, font=BOLD, color="#" + th["primary_ink"], bold=True, leading=1.12)
    facts = spec.get("facts", [])
    if facts:
        fw = (W - 2 * MX) / len(facts)
        for i, f in enumerate(facts):
            x = MX + i * fw
            c.setFillColor(HexColor("#" + K.mix(th["primary_ink"], th["primary"], 0.45))); c.setFont(BOLD, 7.5)
            c.drawString(x, 652, f["label"].upper())
            text_block(c, [f["value"]], x, 646, fw - 12, 14, 11.5, 8.5, color="#" + th["primary_ink"])
    # challenge / solution
    colw, gap = 252.0, 20.0
    for i, (head, key) in enumerate((("THE CHALLENGE", "challenge"), ("THE SOLUTION", "solution"))):
        x = MX + i * (colw + gap)
        c.setFillColor(AC); c.setFont(BOLD, 9.5); c.drawString(x, 600, head)
        c.setFillColor(HexColor("#" + th["line"])); c.rect(x, 595, colw, 0.8, stroke=0, fill=1)
        items = [(t, True) for t in spec[key]] if isinstance(spec[key], list) else [spec[key]]
        try:
            text_block(c, items, x, 586, colw, 140, 12, 9, color="#" + th["ink"])
        except FitError as e:
            raise FitError("%s: %s" % (head.lower(), e))
    # metrics
    mw = (W - 2 * MX - 2 * 14) / 3
    c.setFillColor(MUT); c.setFont(BOLD, 9.5); c.drawString(MX, 437, "THE RESULTS")
    for i, m in enumerate(spec["results"]):
        x = MX + i * (mw + 14)
        c.setFillColor(SUR); c.roundRect(x, 322, mw, 100, 6, stroke=0, fill=1)
        c.setFillColor(AC); c.rect(x, 322 + 94, mw, 6, stroke=0, fill=1) if False else None
        val = d["metrics"][i]
        # shrink big value to fit width
        sz = 40
        while pdfmetrics.stringWidth(val, BOLD, sz) > mw - 24 and sz > 22:
            sz -= 1
        c.setFillColor(AC); c.setFont(BOLD, sz); c.drawCentredString(x + mw / 2, 380, val)
        ps, s, tot = fit_paras([m["label"]], mw - 28, 40, 11, 8.5, color="#" + th["muted"], align=TA_CENTER)
        draw_paras(c, ps, x + 14, 366, mw - 28)
    # quote
    q = spec["quote"]
    qy = 296
    c.setFillColor(AC); c.rect(MX, qy - 100, 4, 100, stroke=0, fill=1)
    c.setFillColor(HexColor("#" + K.mix(th["accent"], th["bg"], 0.55))); c.setFont(BOLD, 54); c.drawString(MX + 16, qy - 38, "“")
    s, tot = text_block(c, [q["text"]], MX + 52, qy - 4, W - 2 * MX - 52, 78, 16, 10.5, font=ITAL, color="#" + th["ink"], leading=1.28)
    ay = qy - 4 - tot - 16
    c.setFillColor(INK); c.setFont(BOLD, 10.5); c.drawString(MX + 52, ay, q["name"])
    c.setFillColor(MUT); c.setFont(REG, 10); c.drawString(MX + 52, ay - 13, q["role"])
    # CTA band
    c.setFillColor(P); c.rect(0, 0, W, 160, stroke=0, fill=1)
    cta = spec["cta"]
    text_block(c, [cta["headline"]], MX, 140, 340, 56, 22, 14, font=BOLD, color="#" + th["primary_ink"], bold=True, leading=1.15)
    text_block(c, [cta["body"]], MX, 76, 340, 36, 11.5, 9, color="#" + K.mix(th["primary_ink"], th["primary"], 0.2))
    bx, bw = W - MX - 190, 190
    c.setFillColor(A2 if th.get("accent2") else AC); c.roundRect(bx, 88, bw, 42, 21, stroke=0, fill=1)
    c.setFillColor(HexColor("#" + th["accent2_ink"])); c.setFont(BOLD, 13)
    bt = cta["button"]
    c.drawCentredString(bx + bw / 2, 104, bt)
    c.setFillColor(PI); c.setFont(REG, 10); c.drawCentredString(bx + bw / 2, 70, cta["url"])
    c.setFillColor(HexColor("#" + K.mix(th["primary_ink"], th["primary"], 0.45))); c.setFont(REG, 8)
    c.drawString(MX, 22, "%s  |  %s" % (spec["vendor"], spec.get("footer", "Customer results are specific to this customer and may vary.")))
    c.showPage(); c.save()


# ---------------------------------------------------------------- PPTX
def build_pptx(spec, th, d, path):
    prs = K.new_prs()
    s = K.blank_slide(prs, th)
    M = K.MARGIN
    K.rect(s, 0, 0, K.SLIDE_W, 1.9, th["primary"])
    K.rect(s, M, 0.4, 0.55, 0.07, th["accent"])
    K.add_text(s, M + 0.7, 0.3, 8, 0.28, ("Customer story | " + spec["customer"]).upper(), th, size=11, bold=True, color=th["accent2"])
    K.add_text(s, M, 0.7, K.CONTENT_W, 1.1, spec["headline"], th, max_pt=32, min_pt=20, bold=True, color=th["primary_ink"], spacing=0.95, what="headline")
    lw = 7.3
    for i, (head, key) in enumerate((("THE CHALLENGE", "challenge"), ("THE SOLUTION", "solution"))):
        y = 2.2 + i * 1.95
        K.add_text(s, M, y, 3, 0.25, head, th, size=11, bold=True, color=th["accent"])
        items = [{"text": t, "bullet": True} for t in spec[key]] if isinstance(spec[key], list) else [spec[key]]
        K.add_text(s, M, y + 0.3, lw, 1.55, items, th, max_pt=15, min_pt=10, gap_pt=5, what=head.lower())
    rx = M + lw + 0.5
    rw = K.SLIDE_W - M - rx
    for i, m in enumerate(spec["results"]):
        y = 2.2 + i * 1.3
        K.card(s, rx, y, rw, 1.18, th)
        K.add_text(s, rx + 0.2, y, 1.9, 1.18, d["metrics"][i], th, max_pt=36, min_pt=22, bold=True, color=th["accent"], anchor="m", what="metric")
        K.add_text(s, rx + 2.15, y + 0.1, rw - 2.3, 0.98, m["label"], th, max_pt=13, min_pt=9, anchor="m", what="metric label")
    q = spec["quote"]
    K.rect(s, M, 6.1, 0.06, 0.8, th["accent"])
    K.add_text(s, M + 0.25, 6.05, 8.0, 0.65, "“" + q["text"] + "”", th, max_pt=15, min_pt=10, color=th["ink"], what="quote")
    K.add_text(s, M + 0.25, 6.72, 8.0, 0.25, "%s, %s" % (q["name"], q["role"]), th, size=11, bold=True, color=th["muted"])
    K.rect(s, 8.9, 6.15, 3.83, 0.7, th["accent"], shape=K.MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.5)
    K.add_text(s, 8.9, 6.15, 3.83, 0.7, [{"text": spec["cta"]["button"], "bold": True, "gap": 0}, {"text": spec["cta"]["url"], "scale": 0.8, "gap": 0}],
               th, max_pt=15, min_pt=10, color=th["accent_ink"], align="c", anchor="m", what="cta")
    K.notes(s, spec.get("notes"))
    K.save(prs, path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--spec"); ap.add_argument("--theme"); ap.add_argument("--pdf"); ap.add_argument("--pptx")
    ap.add_argument("--list-themes", action="store_true")
    a = ap.parse_args()
    if a.list_themes:
        for k, v in json.load(open(THEMES)).items():
            print(k, "-", v["label"])
        return
    spec = json.load(open(a.spec, encoding="utf-8"))
    th = K.load_theme(THEMES, a.theme or spec.get("theme"))
    ph = K.find_placeholders(K.walk_strings(spec))
    if ph:
        print("WARNING: unresolved placeholders (ask the user or remove): %s" % ph)
    if len(spec["results"]) != 3:
        sys.exit("results must have exactly 3 metrics")
    d = {"metrics": [metric_value(m) for m in spec["results"]]}
    tok = {}
    for i, v in enumerate(d["metrics"], 1):
        tok["m%d" % i] = v
        tok["m%d_abs" % i] = v.lstrip("+\u2212-")
    spec["headline"] = spec["headline"].format(**tok)  # tokens {m1} {m1_abs} ... keep headline and callouts in sync
    errs = []
    for label, fn, path in (("pdf", build_pdf, a.pdf), ("pptx", build_pptx, a.pptx)):
        if not path:
            continue
        try:
            os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
            fn(spec, th, d, path)
            print("saved", path)
        except (FitError, K.FitError) as e:
            errs.append("%s: %s" % (label, e))
    print("metrics:", d["metrics"])
    if errs:
        print("PROBLEMS TO FIX (shorten the text):")
        for e in errs:
            print(" -", e)
        sys.exit(2)


if __name__ == "__main__":
    main()
