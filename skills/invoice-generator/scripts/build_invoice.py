#!/usr/bin/env python3
"""Build a professional invoice (PDF, optional DOCX) from JSON (+ optional CSV line items).

Usage:
  python build_invoice.py invoice.json out_dir [--items items.csv] [--name invoice]
                          [--theme classic|minimal|bold|path/to/theme.json] [--docx] [--no-pdf]
Offline and deterministic. All money maths uses Decimal with ROUND_HALF_UP; QA re-derives
every total from the rounded figures that are printed and exits with code 2 if anything differs.
"""
import argparse, csv, datetime as dt, json, os, re, sys
from decimal import Decimal, ROUND_HALF_UP

HERE = os.path.dirname(os.path.abspath(__file__))
CURRENCIES = {  # symbol, position, code
    "USD": ("$", "pre"), "EUR": ("€", "pre"), "GBP": ("£", "pre"),
}
D = lambda v: Decimal(str(v if v not in (None, "") else 0))
q2 = lambda d: d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
PLACEHOLDER = re.compile(r"\[[^\]]+\]|\bTBD\b|\bTODO\b|lorem|\bXXX+\b", re.I)


# ---------------------------------------------------------------- money formatting
def money(amount, cur, style="en", sign=False):
    """en: $1,234.50   eu: 1.234,50 EUR-symbol after."""
    amount = q2(D(amount))
    neg = amount < 0
    s = f"{abs(amount):,.2f}"
    sym = CURRENCIES.get(cur, (cur + " ", "pre"))[0]
    if style == "eu":
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        out = f"{s} {sym.strip()}"
    else:
        out = f"{sym}{s}"
    return ("-" if neg else "") + out


def fmt_date(iso):
    return dt.date.fromisoformat(iso).strftime("%d %b %Y")


def fmt_qty(q):
    q = D(q)
    return f"{q.normalize():f}" if q != q.to_integral() else f"{int(q)}"


# ---------------------------------------------------------------- calculation
def load_items(spec, csv_path):
    items = list(spec.get("items", []))
    if csv_path:
        with open(csv_path, newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                if not (r.get("description") or "").strip():
                    continue
                items.append({k: (v.strip() if isinstance(v, str) else v) for k, v in r.items() if v not in (None, "")})
    return items


def compute(spec, items):
    cur = spec.get("currency", "USD")
    default_rate = D(spec.get("default_tax_rate", 0))
    rows = []
    for i, it in enumerate(items, 1):
        qty, price = D(it.get("quantity", 1)), D(it.get("unit_price", 0))
        gross = q2(qty * price)
        dp = D(it.get("discount_percent", 0))
        disc = q2(gross * dp / 100)
        net = gross - disc
        rows.append(dict(n=i, description=it.get("description", ""), details=it.get("details", ""), unit=it.get("unit", ""),
                         qty=qty, price=price, gross=gross, disc_pct=dp, disc=disc, net=net,
                         rate=D(it.get("tax_rate", default_rate)), alloc=Decimal(0)))
    subtotal = sum((r["net"] for r in rows), Decimal(0))

    # invoice-level discount, allocated pro rata so VAT per rate stays correct
    dsc = spec.get("discount") or {}
    if dsc.get("percent"):
        inv_disc = q2(subtotal * D(dsc["percent"]) / 100)
    else:
        inv_disc = q2(D(dsc.get("amount", 0)))
    inv_disc = min(inv_disc, subtotal)
    if inv_disc and subtotal:
        rem = inv_disc
        for r in rows:
            r["alloc"] = q2(inv_disc * r["net"] / subtotal)
            rem -= r["alloc"]
        biggest = max(rows, key=lambda r: r["net"])
        biggest["alloc"] += rem  # push rounding remainder onto the largest line

    ship = spec.get("shipping") or {}
    ship_amt = q2(D(ship.get("amount", 0)))
    ship_rate = D(ship.get("tax_rate", default_rate))

    groups = {}
    for r in rows:
        g = groups.setdefault(r["rate"], Decimal(0)); groups[r["rate"]] = g + r["net"] - r["alloc"]
    if ship_amt:
        groups[ship_rate] = groups.get(ship_rate, Decimal(0)) + ship_amt
    tax_lines = []
    for rate in sorted(groups):
        base = groups[rate]
        tax_lines.append(dict(rate=rate, base=base, tax=q2(base * rate / 100)))
    net_total = subtotal - inv_disc + ship_amt
    tax_total = sum((t["tax"] for t in tax_lines), Decimal(0))
    total = net_total + tax_total
    paid = q2(D(spec.get("amount_paid", 0)))
    return dict(cur=cur, rows=rows, subtotal=subtotal, inv_disc=inv_disc, disc_label=dsc.get("label", "Discount"),
                disc_pct=D(dsc.get("percent", 0)), ship_amt=ship_amt, ship_label=ship.get("label", "Shipping"),
                tax_lines=tax_lines, net_total=net_total, tax_total=tax_total, total=total, paid=paid, due=total - paid)


def dates(spec):
    inv = spec["invoice"]
    issue = dt.date.fromisoformat(inv["issue_date"])
    due = dt.date.fromisoformat(inv["due_date"]) if inv.get("due_date") else issue + dt.timedelta(days=int(inv.get("due_days", 30)))
    return issue, due


def qa(spec, calc):
    """Independent re-derivation from the printed (rounded) numbers. Returns (errors, warnings)."""
    err, warn = [], []
    rows = calc["rows"]
    for r in rows:
        if q2(r["qty"] * r["price"]) - r["disc"] != r["net"]:
            err.append(f"line {r['n']}: qty x price - discount != amount")
        if r["qty"] <= 0 and r["net"] >= 0:
            warn.append(f"line {r['n']}: zero/negative quantity")
    if sum(r["net"] for r in rows) != calc["subtotal"]:
        err.append("sum of line amounts != subtotal")
    if sum(r["alloc"] for r in rows) != calc["inv_disc"]:
        err.append("discount allocation != invoice discount")
    if calc["subtotal"] - calc["inv_disc"] + calc["ship_amt"] != calc["net_total"]:
        err.append("subtotal - discount + shipping != net total")
    if sum(t["base"] for t in calc["tax_lines"]) != calc["net_total"]:
        err.append("tax bases do not add up to net total")
    for t in calc["tax_lines"]:
        if q2(t["base"] * t["rate"] / 100) != t["tax"]:
            err.append(f"tax at {t['rate']}% recomputed differently")
    if calc["net_total"] + sum(t["tax"] for t in calc["tax_lines"]) != calc["total"]:
        err.append("net + tax != total")
    if calc["total"] - calc["paid"] != calc["due"]:
        err.append("total - paid != balance due")
    if calc["due"] < 0:
        warn.append("balance due is negative (credit note?)")
    issue, due = dates(spec)
    if due < issue:
        err.append("due date is before issue date")
    if calc["cur"] not in CURRENCIES:
        warn.append(f"currency {calc['cur']} not in the built-in list; code will be used as prefix")
    for sect, keys in (("seller", ["name", "address"]), ("buyer", ["name"]), ("invoice", ["number", "issue_date"])):
        for k in keys:
            if not spec.get(sect, {}).get(k):
                err.append(f"missing {sect}.{k}")
    if not spec.get("payment", {}).get("iban") and not spec.get("payment", {}).get("account_number") and not spec.get("payment", {}).get("link"):
        warn.append("no payment details (iban/account_number/link)")
    if not spec.get("seller", {}).get("tax_id") and any(t["rate"] > 0 for t in calc["tax_lines"]):
        warn.append("tax charged but seller.tax_id is empty (most jurisdictions require it on the invoice)")
    def walk(o):
        if isinstance(o, str): yield o
        elif isinstance(o, dict): yield from (s for v in o.values() for s in walk(v))
        elif isinstance(o, list): yield from (s for v in o for s in walk(v))
    for sv in walk(spec):
        for m in PLACEHOLDER.findall(sv):
            warn.append(f"placeholder text left in data: {m!r}")
    return err, warn


# ---------------------------------------------------------------- theme
def load_theme(arg):
    p = arg if arg and os.path.exists(arg) else os.path.join(HERE, "..", "templates", "themes.json")
    data = json.load(open(p))
    if arg and os.path.exists(arg) and "primary" in data:
        return data
    return data[arg if arg in data else "classic"]


# ---------------------------------------------------------------- PDF
def build_pdf(spec, calc, th, path, style):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib.enums import TA_RIGHT
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, Flowable
    from reportlab.lib.fonts import addMapping

    reg, bold = "Helvetica", "Helvetica-Bold"
    base = "/usr/share/fonts/truetype/liberation/"
    if os.path.exists(base + "LiberationSans-Regular.ttf"):
        pdfmetrics.registerFont(TTFont("LS", base + "LiberationSans-Regular.ttf"))
        pdfmetrics.registerFont(TTFont("LS-B", base + "LiberationSans-Bold.ttf"))
        addMapping("LS", 0, 0, "LS"); addMapping("LS", 1, 0, "LS-B"); addMapping("LS", 0, 1, "LS"); addMapping("LS", 1, 1, "LS-B")
        reg, bold = "LS", "LS-B"
    C = lambda h: colors.HexColor("#" + h)
    prim, acc, txt, mut, line, band = (C(th[k]) for k in ("primary", "accent", "text", "muted", "line", "band"))
    cur = calc["cur"]
    M = lambda v: money(v, cur, style)

    def S(name, size=9, font=None, color=txt, lead=None, **kw):
        return ParagraphStyle(name, fontName=font or reg, fontSize=size, leading=lead or size * 1.35, textColor=color, **kw)
    body, small, lab = S("b"), S("s", 8, color=mut), S("l", 7.5, bold, mut, charSpace=0.6)
    right, rightb = S("r", 9, alignment=TA_RIGHT), S("rb", 9, bold, alignment=TA_RIGHT)
    cell_h = S("h", 8, bold, C(th["header_text"]), charSpace=0.4)
    cell_hr = S("hr", 8, bold, C(th["header_text"]), alignment=TA_RIGHT, charSpace=0.4)

    seller, buyer, inv, pay = spec["seller"], spec["buyer"], spec["invoice"], spec.get("payment", {})
    issue, due = dates(spec)
    W, H = A4
    LM = RM = 18 * mm

    class Logo(Flowable):
        def __init__(s, text, size=15 * mm):
            super().__init__(); s.t, s.size = text, size; s.width = s.height = size
        def draw(s):
            c = s.canv
            c.setFillColor(prim); c.roundRect(0, 0, s.size, s.size, 3 * mm, stroke=0, fill=1)
            c.setFillColor(colors.white if th["style"] == "bar" else colors.white)
            c.setFont(bold, 13); c.drawCentredString(s.size / 2, s.size / 2 - 4.5, s.t)

    def footer(c, d):
        c.saveState(); c.setFont(reg, 7.5); c.setFillColor(mut)
        c.drawString(LM, 10 * mm, f"{seller['name']}  |  Invoice {inv['number']}")
        c.drawRightString(W - RM, 10 * mm, f"Page {d.page}")
        c.setStrokeColor(line); c.setLineWidth(0.5); c.line(LM, 14 * mm, W - RM, 14 * mm); c.restoreState()

    doc = BaseDocTemplate(path, pagesize=A4, leftMargin=LM, rightMargin=RM, topMargin=16 * mm, bottomMargin=20 * mm,
                          title=f"Invoice {inv['number']}", author=seller["name"])
    doc.addPageTemplates([PageTemplate(id="p", frames=[Frame(LM, 20 * mm, W - LM - RM, H - 36 * mm, 0, 0, 0, 0)], onPage=footer)])
    fw = W - LM - RM
    st = []

    # header: logo + seller | INVOICE title
    logo_txt = seller.get("logo_text") or "".join(w[0] for w in seller["name"].split()[:2]).upper()
    left = Table([[Logo(logo_txt), Paragraph(f"<b>{seller['name']}</b>", S("sn", 11, bold, prim, 14))]],
                 colWidths=[19 * mm, 70 * mm], style=[("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0)])
    title = [Paragraph("INVOICE", S("t", 26, bold, prim, 28, alignment=TA_RIGHT)),
             Paragraph(f"No. {inv['number']}", S("tn", 10, reg, mut, 14, alignment=TA_RIGHT))]
    head = Table([[left, title]], colWidths=[fw * 0.55, fw * 0.45], style=[("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)])
    st += [head, Spacer(1, 3 * mm)]
    rule = Table([[""]], colWidths=[fw], rowHeights=[1.2], style=[("BACKGROUND", (0, 0), (-1, -1), acc if th["style"] == "bar" else prim)])
    st += [rule, Spacer(1, 6 * mm)]

    def block(label, name, lines):
        t = [Paragraph(label, lab), Spacer(1, 1.5 * mm), Paragraph(f"<b>{name}</b>", body)] + [Paragraph(x, body) for x in lines]
        return t
    from_lines = list(seller.get("address", [])) + [x for x in (seller.get("email"), seller.get("phone"), seller.get("tax_id")) if x]
    to_lines = ([buyer["contact"]] if buyer.get("contact") else []) + list(buyer.get("address", [])) + ([buyer["tax_id"]] if buyer.get("tax_id") else [])
    meta = [("Issue date", fmt_date(inv["issue_date"])), ("Due date", fmt_date(due.isoformat()))]
    if pay.get("terms"): meta.append(("Terms", pay["terms"]))
    if inv.get("po_number"): meta.append(("PO number", inv["po_number"]))
    mt = Table([[Paragraph(k, small), Paragraph(v, rightb)] for k, v in meta], colWidths=[22 * mm, fw * 0.28 - 22 * mm],
               style=[("LINEBELOW", (0, 0), (-1, -1), 0.4, line), ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)])
    info = Table([[block("BILL TO", buyer["name"], to_lines), block("FROM", seller["name"], from_lines), [Paragraph("INVOICE DETAILS", lab), Spacer(1, 1.5 * mm), mt]]],
                 colWidths=[fw * 0.3, fw * 0.3, fw * 0.4], style=[("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (1, -1), 8), ("RIGHTPADDING", (2, 0), (2, -1), 0)])
    st += [info, Spacer(1, 8 * mm)]

    # items table
    has_disc = any(r["disc_pct"] for r in calc["rows"])
    has_tax = any(r["rate"] for r in calc["rows"])
    cols = [("#", 7, "l"), ("DESCRIPTION", None, "l"), ("QTY", 16, "r"), ("UNIT PRICE", 24, "r")]
    if has_disc: cols.append(("DISC.", 13, "r"))
    if has_tax: cols.append(("TAX", 12, "r"))
    cols.append(("AMOUNT", 27, "r"))
    fixed = sum(c[1] for c in cols if c[1]) * mm
    widths = [(c[1] * mm if c[1] else fw - fixed) for c in cols]
    data = [[Paragraph(c[0], cell_hr if c[2] == "r" else cell_h) for c in cols]]
    for r in calc["rows"]:
        d = f"<b>{r['description']}</b>" + (f"<br/><font size=8 color='#{th['muted']}'>{r['details']}</font>" if r["details"] else "")
        row = [Paragraph(str(r["n"]), small), Paragraph(d, body),
               Paragraph(fmt_qty(r["qty"]) + (f" <font size=7.5 color='#{th['muted']}'>{r['unit']}</font>" if r["unit"] else ""), right),
               Paragraph(M(r["price"]), right)]
        if has_disc: row.append(Paragraph(f"{r['disc_pct'].normalize():f}%" if r["disc_pct"] else "-", right))
        if has_tax: row.append(Paragraph(f"{r['rate'].normalize():f}%", right))
        row.append(Paragraph(M(r["net"]), rightb))
        data.append(row)
    ts = [("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
          ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
          ("LINEBELOW", (0, 1), (-1, -1), 0.4, line)]
    if th["style"] == "bar":
        ts += [("BACKGROUND", (0, 0), (-1, 0), prim)] + [("BACKGROUND", (0, i), (-1, i), band) for i in range(2, len(data), 2)]
    else:
        ts += [("LINEBELOW", (0, 0), (-1, 0), 1.2, prim), ("LINEABOVE", (0, 0), (-1, 0), 1.2, prim)]
    st += [Table(data, colWidths=widths, repeatRows=1, style=ts), Spacer(1, 5 * mm)]

    # totals
    tl = [("Subtotal", M(calc["subtotal"]))]
    if calc["inv_disc"]:
        pct = f" ({calc['disc_pct'].normalize():f}%)" if calc["disc_pct"] else ""
        tl.append((f"{calc['disc_label']}{pct}", "-" + M(calc["inv_disc"])))
    if calc["ship_amt"]: tl.append((calc["ship_label"], M(calc["ship_amt"])))
    if len(calc["tax_lines"]) > 1 or calc["inv_disc"] or calc["ship_amt"]:
        tl.append(("Net total", M(calc["net_total"])))
    for t in calc["tax_lines"]:
        if t["rate"] or len(calc["tax_lines"]) == 1:
            tl.append((f"{spec.get('tax_label', 'VAT')} {t['rate'].normalize():f}% on {M(t['base'])}", M(t["tax"])))
    tl.append(("Total", M(calc["total"])))
    if calc["paid"]: tl.append(("Amount paid", "-" + M(calc["paid"])))
    tw = fw * 0.5
    tdata = [[Paragraph(a, S("ta", 9, reg, txt)), Paragraph(b, rightb)] for a, b in tl]
    tt = Table(tdata, colWidths=[tw * 0.64, tw * 0.36], style=[("LINEBELOW", (0, 0), (-1, -1), 0.4, line), ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 5)])
    due_box = Table([[Paragraph("BALANCE DUE", S("bd", 9, bold, C(th["header_text"]), charSpace=0.6)), Paragraph(M(calc["due"]), S("bv", 14, bold, C(th["header_text"]), 17, alignment=TA_RIGHT))]],
                    colWidths=[tw * 0.42, tw * 0.58], style=[("BACKGROUND", (0, 0), (-1, -1), prim if th["style"] == "bar" else colors.white), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, -1), 7), ("BOTTOMPADDING", (0, 0), (-1, -1), 7), ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 8)] +
                    ([] if th["style"] == "bar" else [("LINEABOVE", (0, 0), (-1, -1), 1.2, prim), ("LINEBELOW", (0, 0), (-1, -1), 1.2, prim)]))
    if th["style"] != "bar":  # minimal: dark text on white
        due_box._cellvalues[0][0] = Paragraph("BALANCE DUE", S("bd2", 9, bold, prim, charSpace=0.6))
        due_box._cellvalues[0][1] = Paragraph(M(calc["due"]), S("bv2", 14, bold, prim, 17, alignment=TA_RIGHT))
    stack = Table([[tt], [Spacer(1, 2 * mm)], [due_box]], colWidths=[tw], style=[("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 0)])
    # payment block (left of totals)
    pl = [Paragraph("PAYMENT DETAILS", lab), Spacer(1, 1.5 * mm)]
    for k, lbl in (("method", "Method"), ("bank", "Bank"), ("account_name", "Account name"), ("iban", "IBAN"), ("account_number", "Account no."),
                   ("sort_code", "Sort code"), ("bic", "BIC/SWIFT"), ("account_ref", "Account ref"), ("link", "Pay online"), ("reference", "Reference")):
        if pay.get(k): pl.append(Paragraph(f"<font color='#{th['muted']}'>{lbl}:</font>  {pay[k]}", body))
    if inv.get("notes"):
        pl += [Spacer(1, 4 * mm), Paragraph("NOTES", lab), Spacer(1, 1.5 * mm), Paragraph(inv["notes"], small)]
    tax_note = spec.get("tax_note")
    if tax_note:
        pl += [Spacer(1, 3 * mm), Paragraph(tax_note, small)]
    bottom = Table([[pl, stack]], colWidths=[fw * 0.46, fw * 0.54], style=[("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)])
    # keep totals together; allow overall split only between items and this block
    st.append(KeepTogether([bottom]))
    doc.build(st)


# ---------------------------------------------------------------- DOCX
def build_docx(spec, calc, th, path, style):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx.shared import Pt, RGBColor, Mm
    cur = calc["cur"]; M = lambda v: money(v, cur, style)
    seller, buyer, inv, pay = spec["seller"], spec["buyer"], spec["invoice"], spec.get("payment", {})
    issue, due = dates(spec)
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = sec.right_margin = Mm(18); sec.top_margin = Mm(16); sec.bottom_margin = Mm(18)
    n = doc.styles["Normal"]; n.font.name = "Calibri"; n.font.size = Pt(10)
    n.paragraph_format.space_after = Pt(0)
    rgb = RGBColor.from_string

    def shade(cell, fill):
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))

    def run(p, text, bold=False, size=None, color=None):
        r = p.add_run(text); r.bold = bold
        if size: r.font.size = Pt(size)
        if color: r.font.color.rgb = rgb(color)
        return r

    def para(cell_or_doc, text="", bold=False, size=None, color=None, align=None, first=False):
        p = cell_or_doc.paragraphs[0] if first else cell_or_doc.add_paragraph()
        if text: run(p, text, bold, size, color)
        if align == "r": p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        return p

    def borders(table, color):
        table._tbl.tblPr.append(parse_xml(f'<w:tblBorders {nsdecls("w")}><w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="{color}"/></w:tblBorders>'))

    t = doc.add_table(rows=1, cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    para(t.cell(0, 0), seller["name"], True, 14, th["primary"], first=True)
    p = para(t.cell(0, 1), "INVOICE", True, 26, th["primary"], "r", first=True)
    para(t.cell(0, 1), f"No. {inv['number']}", size=10, color=th["muted"], align="r")
    doc.add_paragraph()
    t = doc.add_table(rows=1, cols=3)
    c = t.cell(0, 0); para(c, "BILL TO", True, 7.5, th["muted"], first=True); para(c, buyer["name"], True)
    for x in ([buyer["contact"]] if buyer.get("contact") else []) + buyer.get("address", []) + ([buyer["tax_id"]] if buyer.get("tax_id") else []): para(c, x)
    c = t.cell(0, 1); para(c, "FROM", True, 7.5, th["muted"], first=True); para(c, seller["name"], True)
    for x in seller.get("address", []) + [y for y in (seller.get("email"), seller.get("phone"), seller.get("tax_id")) if y]: para(c, x)
    c = t.cell(0, 2); para(c, "INVOICE DETAILS", True, 7.5, th["muted"], first=True)
    for k, v in [("Issue date", fmt_date(inv["issue_date"])), ("Due date", fmt_date(due.isoformat()))] + ([("PO number", inv["po_number"])] if inv.get("po_number") else []) + ([("Terms", pay["terms"])] if pay.get("terms") else []):
        para(c, f"{k}: {v}")
    doc.add_paragraph()
    cols = ["#", "Description", "Qty", "Unit price", "Disc.", "Tax", "Amount"]
    t = doc.add_table(rows=1, cols=len(cols)); borders(t, th["line"])
    widths = [8, 60, 22, 25, 13, 13, 27]
    for i, h in enumerate(cols):
        cell = t.rows[0].cells[i]; shade(cell, th["primary"]); para(cell, h.upper(), True, 8, "FFFFFF", "r" if i > 1 else None, first=True)
    for r in calc["rows"]:
        cs = t.add_row().cells
        para(cs[0], str(r["n"]), size=8, color=th["muted"], first=True)
        para(cs[1], r["description"], True, first=True)
        if r["details"]: para(cs[1], r["details"], size=8, color=th["muted"])
        para(cs[2], fmt_qty(r["qty"]) + (f" {r['unit']}" if r["unit"] else ""), align="r", first=True)
        para(cs[3], M(r["price"]), align="r", first=True)
        para(cs[4], f"{r['disc_pct'].normalize():f}%" if r["disc_pct"] else "-", align="r", first=True)
        para(cs[5], f"{r['rate'].normalize():f}%", align="r", first=True)
        para(cs[6], M(r["net"]), True, align="r", first=True)
    t.autofit = False
    for i, w in enumerate(widths): t.columns[i].width = Mm(w)
    for row in t.rows:
        for i, w in enumerate(widths): row.cells[i].width = Mm(w)
    doc.add_paragraph()
    tl = [("Subtotal", M(calc["subtotal"]))]
    if calc["inv_disc"]: tl.append((calc["disc_label"], "-" + M(calc["inv_disc"])))
    if calc["ship_amt"]: tl.append((calc["ship_label"], M(calc["ship_amt"])))
    for tx in calc["tax_lines"]:
        if tx["rate"]: tl.append((f"{spec.get('tax_label', 'VAT')} {tx['rate'].normalize():f}% on {M(tx['base'])}", M(tx["tax"])))
    tl.append(("Total", M(calc["total"])))
    if calc["paid"]: tl.append(("Amount paid", "-" + M(calc["paid"])))
    tl.append(("BALANCE DUE", M(calc["due"])))
    t = doc.add_table(rows=0, cols=2); t.alignment = WD_TABLE_ALIGNMENT.RIGHT; borders(t, th["line"]); t.autofit = False
    for i, (a, b) in enumerate(tl):
        cs = t.add_row().cells; last = i == len(tl) - 1
        para(cs[0], a, last, 10, "FFFFFF" if last else None, first=True); para(cs[1], b, True, 12 if last else 10, "FFFFFF" if last else None, "r", first=True)
        if last: shade(cs[0], th["primary"]); shade(cs[1], th["primary"])
        cs[0].width, cs[1].width = Mm(80), Mm(32)
    doc.add_paragraph()
    para(doc, "PAYMENT DETAILS", True, 7.5, th["muted"])
    for k, lbl in (("method", "Method"), ("bank", "Bank"), ("account_name", "Account name"), ("iban", "IBAN"), ("account_number", "Account no."), ("sort_code", "Sort code"), ("bic", "BIC/SWIFT"), ("account_ref", "Account ref"), ("link", "Pay online"), ("reference", "Reference")):
        if pay.get(k): para(doc, f"{lbl}: {pay[k]}")
    if inv.get("notes"):
        doc.add_paragraph(); para(doc, "NOTES", True, 7.5, th["muted"]); para(doc, inv["notes"], size=9, color=th["muted"])
    if spec.get("tax_note"): doc.add_paragraph(); para(doc, spec["tax_note"], size=9, color=th["muted"])
    doc.core_properties.title = f"Invoice {inv['number']}"; doc.core_properties.author = seller["name"]
    doc.save(path)


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec"); ap.add_argument("out_dir")
    ap.add_argument("--items", help="CSV of line items appended to spec.items")
    ap.add_argument("--name"); ap.add_argument("--theme")
    ap.add_argument("--docx", action="store_true"); ap.add_argument("--no-pdf", action="store_true")
    a = ap.parse_args()
    spec = json.load(open(a.spec, encoding="utf-8"))
    items = load_items(spec, a.items)
    if not items:
        sys.exit("ERROR: no line items")
    calc = compute(spec, items)
    th = load_theme(a.theme or spec.get("theme", "classic"))
    style = spec.get("number_style", "en")
    err, warn = qa(spec, calc)
    for w in warn: print("QA WARNING:", w)
    for e in err: print("QA ERROR:", e)
    os.makedirs(a.out_dir, exist_ok=True)
    name = a.name or re.sub(r"[^A-Za-z0-9_-]+", "-", spec["invoice"]["number"])
    if not a.no_pdf:
        build_pdf(spec, calc, th, os.path.join(a.out_dir, name + ".pdf"), style)
    if a.docx:
        build_docx(spec, calc, th, os.path.join(a.out_dir, name + ".docx"), style)
    summary = {k: str(calc[k]) for k in ("subtotal", "inv_disc", "ship_amt", "net_total", "tax_total", "total", "paid", "due")}
    summary["currency"] = calc["cur"]; summary["tax"] = {str(t["rate"]): str(t["tax"]) for t in calc["tax_lines"]}
    json.dump(summary, open(os.path.join(a.out_dir, name + ".totals.json"), "w"), indent=2)
    print(f"Total {money(calc['total'], calc['cur'], style)} | Balance due {money(calc['due'], calc['cur'], style)} | QA {'FAILED' if err else 'OK'}")
    sys.exit(2 if err else 0)


if __name__ == "__main__":
    main()
