#!/usr/bin/env python3
"""Build a proposal + SOW (.docx and .pdf) from a JSON spec.

Usage: python build_proposal.py spec.json out_dir [--theme templates/theme.json] [--no-pdf]
Deterministic, offline. PDF is produced by LibreOffice (soffice) from the .docx.
"""
import argparse, json, os, re, shutil, subprocess, sys, tempfile
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Pt, RGBColor, Cm, Emu

HERE = os.path.dirname(os.path.abspath(__file__))
PLACEHOLDER = re.compile(r"\[[^\]]+\]|TBD|TODO|lorem", re.I)


def rgb(h): return RGBColor.from_string(h)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    for e in tcPr.findall(qn("w:shd")): tcPr.remove(e)
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def cell_margins(table, top=70, bottom=70, left=100, right=100):
    tblPr = table._tbl.tblPr
    m = parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/>'
                  f'<w:bottom w:w="{bottom}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tblCellMar>')
    tblPr.append(m)


def add_table_style(doc, T):
    """Real custom table style: header row, banded rows, thin horizontal rules."""
    xml = f'''<w:style {nsdecls("w")} w:type="table" w:customStyle="1" w:styleId="BoardroomTable">
<w:name w:val="Boardroom Table"/><w:basedOn w:val="TableNormal"/><w:uiPriority w:val="59"/><w:qFormat/>
<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblBorders>
<w:top w:val="single" w:sz="8" w:space="0" w:color="{T['navy']}"/>
<w:bottom w:val="single" w:sz="8" w:space="0" w:color="{T['navy']}"/>
<w:insideH w:val="single" w:sz="4" w:space="0" w:color="{T['line']}"/></w:tblBorders></w:tblPr>
<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr>
<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{T['navy']}"/></w:tcPr></w:tblStylePr>
<w:tblStylePr w:type="band2Horz"><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{T['band']}"/></w:tcPr></w:tblStylePr>
</w:style>'''
    doc.styles.element.append(parse_xml(xml))


def setup_styles(doc, T):
    f = T["font"]
    st = doc.styles

    def font(style, size, bold=False, color=None, italic=False):
        style.font.name = f; style.font.size = Pt(size); style.font.bold = bold; style.font.italic = italic
        rpr = style.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts"); rpr.append(rf)
        for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"): rf.set(qn(a), f)
        if color: style.font.color.rgb = rgb(color)

    n = st["Normal"]; font(n, 10.5, color="22262E")
    n.paragraph_format.space_after = Pt(6); n.paragraph_format.line_spacing = 1.15
    for name, size, before, after, col in (("Heading 1", 18, 20, 8, T["navy"]), ("Heading 2", 13, 12, 4, T["accent"]), ("Heading 3", 11, 8, 2, T["navy"])):
        s = st[name]; font(s, size, True, col)
        s.paragraph_format.space_before = Pt(before); s.paragraph_format.space_after = Pt(after)
        s.paragraph_format.keep_with_next = True
    h1 = st["Heading 1"].element.get_or_add_pPr()
    h1.append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="4" w:color="{T["accent"]}"/></w:pBdr>'))
    t = st["Title"]; font(t, 34, True, "FFFFFF"); t.paragraph_format.space_after = Pt(6)
    pbdr = t.element.get_or_add_pPr().find(qn("w:pBdr"))
    if pbdr is not None: t.element.pPr.remove(pbdr)
    sub = st["Subtitle"]; font(sub, 16, False, "CFE6E8"); sub.font.italic = False
    for nm, size, col, bold in (("Table Text", 9.5, "22262E", False), ("Table Header", 9.5, "FFFFFF", True),
                                ("Caption", 9, T["grey"], False), ("Cover Meta", 10.5, T["grey"], False), ("Small Note", 9, T["grey"], False)):
        s = st[nm] if nm in [x.name for x in st] else st.add_style(nm, WD_STYLE_TYPE.PARAGRAPH); s.base_style = n; font(s, size, bold, col)
        s.paragraph_format.space_after = Pt(0 if nm.startswith("Table") else 4)
        s.paragraph_format.line_spacing = 1.05
    for nm in ("List Bullet",):
        st[nm].font.name = f; st[nm].paragraph_format.space_after = Pt(3)
    add_table_style(doc, T)


def para(doc_or_cell, text, style=None, bold=False, align=None, color=None, size=None):
    p = doc_or_cell.add_paragraph(style=style)
    r = p.add_run(text); r.bold = bold or None
    if color: r.font.color.rgb = rgb(color)
    if size: r.font.size = Pt(size)
    if align: p.alignment = align
    return p


def cell_text(cell, text, style="Table Text", bold=False, align=None, color=None):
    p = cell.paragraphs[0]; p.style = style
    r = p.add_run(text); r.bold = bold or None
    if color: r.font.color.rgb = rgb(color)
    if align is not None: p.alignment = align
    return p


def mk_table(doc, header, widths_cm, T):
    tb = doc.add_table(rows=1, cols=len(header))
    tb.style = doc.styles["Boardroom Table"]; tb.alignment = WD_TABLE_ALIGNMENT.CENTER; tb.autofit = False
    tblPr = tb._tbl.tblPr
    look = tblPr.find(qn("w:tblLook"))
    if look is None:
        look = OxmlElement("w:tblLook"); tblPr.append(look)
    look.set(qn("w:val"), "0420"); look.set(qn("w:firstRow"), "1"); look.set(qn("w:noHBand"), "0"); look.set(qn("w:noVBand"), "1")
    cell_margins(tb)
    for i, h in enumerate(header):
        c = tb.rows[0].cells[i]; shade(c, T["navy"]); cell_text(c, h, "Table Header")
    trPr = tb.rows[0]._tr.get_or_add_trPr(); trPr.append(OxmlElement("w:tblHeader"))
    set_widths(tb, widths_cm)
    return tb


def set_widths(tb, widths):
    tblPr = tb._tbl.tblPr
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)
    for gc, w in zip(tb._tbl.tblGrid.findall(qn("w:gridCol")), widths): gc.set(qn("w:w"), str(int(w * 567)))
    for row in tb.rows:
        for i, w in enumerate(widths): row.cells[i].width = Cm(w)


def add_row(tb, vals, widths, T, band=None, bold_first=False):
    row = tb.add_row()
    trPr = row._tr.get_or_add_trPr(); trPr.append(OxmlElement("w:cantSplit"))
    for i, v in enumerate(vals):
        cell_text(row.cells[i], str(v), bold=(bold_first and i == 0))
        row.cells[i].width = Cm(widths[i])
        if band: shade(row.cells[i], band)
    return row


def add_field(run, instr):
    for t, txt in (("begin", None), (None, instr), ("end", None)):
        if t:
            e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), t); run._r.append(e)
        else:
            e = OxmlElement("w:instrText"); e.set(qn("xml:space"), "preserve"); e.text = txt; run._r.append(e)


def gantt(phases, path, T, total_weeks):
    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager as fm
    fam = "Carlito" if any("Carlito" in x.name for x in fm.fontManager.ttflist) else "DejaVu Sans"
    plt.rcParams["font.family"] = fam
    n = len(phases)
    fig, ax = plt.subplots(figsize=(7.0, 0.36 * n + 0.75), dpi=220)
    cols = ["#" + T["navy"], "#" + T["accent"], "#" + T["navy"], "#" + T["accent"]]
    for i, p in enumerate(phases):
        y = n - 1 - i
        ax.barh(y, p["end"] - p["start"] + 1, left=p["start"] - 1, height=0.52, color=cols[i % 4], zorder=3)
        ax.text(p["start"] - 1 + 0.15, y, f'Wk {p["start"]}-{p["end"]}', va="center", ha="left", color="white", fontsize=8, zorder=4)
        ax.scatter(p["end"], y, marker="D", s=46, color="#E8A33D", edgecolor="white", linewidth=0.8, zorder=5)
    ax.set_yticks(range(n)); ax.set_yticklabels([p["name"] for p in reversed(phases)], fontsize=9.5, color="#22262E")
    ax.set_xlim(0, total_weeks); ax.set_ylim(-0.6, n - 0.4)
    ax.set_xticks(range(0, total_weeks + 1, 2)); ax.set_xticklabels([str(x) if x else "" for x in range(0, total_weeks + 1, 2)], fontsize=8.5, color="#5B6472")
    ax.xaxis.tick_top(); ax.xaxis.set_label_position("top"); ax.set_xlabel("Week of engagement", fontsize=8.5, color="#5B6472", labelpad=6)
    ax.grid(axis="x", color="#" + T["line"], linewidth=0.7, zorder=0); ax.set_axisbelow(True)
    for s in ax.spines.values(): s.set_visible(False)
    ax.tick_params(length=0)
    ax.scatter([], [], marker="D", s=40, color="#E8A33D", label="Milestone review")
    ax.legend(loc="lower right", frameon=False, fontsize=8, bbox_to_anchor=(1.0, -0.22))
    fig.tight_layout(); fig.savefig(path, facecolor="white"); plt.close(fig)


ORDER = {
 "tblPr": "tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd tblLayout tblCellMar tblLook".split(),
 "tcPr": "cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar textDirection tcFitText vAlign hideMark".split(),
 "pPr": "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange".split(),
 "trPr": "cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight tblHeader tblCellSpacing jc hidden".split(),
}


def fix_order(root):
    """Re-sort children of property elements into OOXML schema order so Word opens the file cleanly."""
    for tag, order in ORDER.items():
        for el in root.iter(qn("w:" + tag)):
            kids = list(el)
            key = lambda k: order.index(k.tag.split("}")[1]) if k.tag.split("}")[1] in order else 999
            seen = {}
            for k in kids:
                seen[k.tag] = k  # keep last duplicate
            kids2 = sorted(seen.values(), key=key)
            for k in kids: el.remove(k)
            for k in kids2: el.append(k)


def money(v, cur): return f"{cur} {v:,.0f}"


def build(spec, out_dir, T, pdf=True, base="proposal"):
    os.makedirs(out_dir, exist_ok=True)
    m = spec["meta"]; doc = Document(); setup_styles(doc, T)
    doc.core_properties.title = f'{m["title"]} - {m["subtitle"]}'; doc.core_properties.author = m["provider"]
    sec = doc.sections[0]
    for s in (sec,):
        s.page_width, s.page_height = Cm(21.0), Cm(29.7)
        s.left_margin = s.right_margin = Cm(2.2); s.top_margin = Cm(2.2); s.bottom_margin = Cm(2.0)
    W = 16.6  # usable width cm

    # ---- Cover
    band = doc.add_table(rows=1, cols=1); band.alignment = WD_TABLE_ALIGNMENT.CENTER; band.autofit = False
    c = band.rows[0].cells[0]; c.width = Cm(W); shade(c, T["navy"]); set_widths(band, [W])
    band.rows[0].height = Cm(10.5)
    tcPr = c._tc.get_or_add_tcPr(); va = OxmlElement("w:vAlign"); va.set(qn("w:val"), "bottom"); tcPr.append(va)
    cell_margins(band, 500, 600, 600, 500)
    p = c.paragraphs[0]; p.style = "Cover Meta"; r = p.add_run(m["provider"].upper()); r.font.color.rgb = rgb("8FD0D6"); r.bold = True; r.font.size = Pt(10)
    p.paragraph_format.space_after = Pt(60)
    c.add_paragraph(m["title"], style="Title"); c.add_paragraph(m["subtitle"], style="Subtitle")
    para(doc, "", "Cover Meta")
    meta = doc.add_table(rows=0, cols=2); meta.autofit = False
    for k, v in (("Prepared for", f'{m["prepared_for"]}, {m["client"]}'), ("Prepared by", m["prepared_by"]), ("Date", m["date"]),
                 ("Valid until", m["valid_until"]), ("Reference", m["reference"])):
        row = meta.add_row(); cell_text(row.cells[0], k.upper(), "Small Note", True, color=T["accent"]); cell_text(row.cells[1], v, "Cover Meta", color="22262E")
        row.cells[0].width = Cm(3.6); row.cells[1].width = Cm(13)
    cell_margins(meta, 60, 60, 600, 100); set_widths(meta, [3.6, 13])
    para(doc, "", "Cover Meta")
    para(doc, "Commercial in confidence. This document contains proposal terms intended solely for the named recipient.", "Small Note")

    # ---- Body section with footer
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    body.footer.is_linked_to_previous = False
    fp = body.footer.paragraphs[0]; fp.style = "Small Note"
    fp.add_run(f'{m["provider"]}  |  {m["title"]}  |  {m["client"]}  |  Page ')
    add_field(fp.add_run(), "PAGE")
    fp.alignment = WD_ALIGN_PARAGRAPH.LEFT
    # make cover footer empty
    sec.footer.is_linked_to_previous = False

    counter = [0]

    def H1(t):
        counter[0] += 1; counter[1:] = []
        doc.add_heading(f"{counter[0]}.  {t}", 1)
        sub[0] = 0

    sub = [0]

    def H2(t):
        sub[0] += 1; doc.add_heading(f"{counter[0]}.{sub[0]}  {t}", 2)

    # 1 Exec summary
    H1("Executive summary")
    for t in spec["executive_summary"]: para(doc, t)
    # key-points callout (single-cell table with accent fill)
    co = doc.add_table(rows=1, cols=1); co.autofit = False; cc = co.rows[0].cells[0]; cc.width = Cm(W); shade(cc, T["accent_light"]); set_widths(co, [W])
    cell_margins(co, 120, 120, 200, 200)
    cell_text(cc, "At a glance", "Table Text", True, color=T["accent"])
    for b in spec.get("executive_bullets", []):
        q = cc.add_paragraph(style="Table Text"); q.add_run("•  " + b); q.paragraph_format.space_before = Pt(3)
    # 2 Understanding
    H1("Understanding of your needs")
    for t in spec["understanding"]["paras"]: para(doc, t)
    tb = mk_table(doc, ["Need", "What we heard"], [3.4, 13.2], T)
    for i, (a, b) in enumerate(spec["understanding"]["needs"]): add_row(tb, [a, b], [3.4, 13.2], T, bold_first=True)
    # 3 Approach
    H1("Approach and phases")
    ap = spec["approach"]; para(doc, ap["intro"])
    H2("Timeline")
    total = max(p["end"] for p in ap["phases"]); total = total + (total % 2)
    img = os.path.join(out_dir, "_gantt.png"); gantt(ap["phases"], img, T, total)
    doc.add_picture(img, width=Cm(W)); doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.paragraphs[-1].paragraph_format.keep_with_next = True
    para(doc, f"Figure 1. Indicative {total}-week plan from kick-off. Diamonds mark milestone reviews.", "Caption")
    H2("Phase detail")
    for i, ph in enumerate(ap["phases"], 1):
        doc.add_heading(f'Phase {i}: {ph["name"]} (weeks {ph["start"]}-{ph["end"]})', 3)
        para(doc, ph["description"])
        for a in ph["activities"]: doc.add_paragraph(a, style="List Bullet")
        q = para(doc, f'Milestone: {ph["milestone"]}', "Small Note", True, color=T["accent"])
    # 4 Deliverables
    H1("Deliverables")
    doc.paragraphs[-1].paragraph_format.page_break_before = True
    para(doc, "Each deliverable is accepted by the client's nominated owner at the milestone review for its phase. The Option column shows the lowest option that includes it.")
    w = [1.2, 5.2, 5.6, 2.6, 2.0]
    tb = mk_table(doc, ["Ref", "Deliverable", "Description", "Phase", "Option"], w, T)
    for d in spec["deliverables"]: add_row(tb, d, w, T, bold_first=True)
    # 5 Assumptions
    H1("Assumptions and dependencies")
    para(doc, "The scope and prices in this document rely on the following. If any proves incorrect we will raise a change request before proceeding.")
    for a in spec["assumptions"]: doc.add_paragraph(a, style="List Number")
    # 6 Pricing
    doc.paragraphs[-1].runs[-1].add_break(WD_BREAK.PAGE) if False else None
    H1("Pricing and options")
    doc.paragraphs[-1].paragraph_format.page_break_before = True
    pr = spec["pricing"]; cur = pr["currency"]; opts = pr["options"]
    para(doc, "We offer three options. All are fixed-fee for the stated scope. Better is our recommendation: it proves the benefit on live shipments before any wider commitment.")
    w = [6.7, 3.3, 3.3, 3.3]
    tb = doc.add_table(rows=2, cols=4); tb.style = doc.styles["Boardroom Table"]; tb.alignment = WD_TABLE_ALIGNMENT.CENTER; tb.autofit = False
    cell_margins(tb, 90, 90, 100, 100)
    tblPr = tb._tbl.tblPr; look = tblPr.find(qn("w:tblLook"))
    if look is None: look = OxmlElement("w:tblLook"); tblPr.append(look)
    look.set(qn("w:val"), "0420"); look.set(qn("w:firstRow"), "1"); look.set(qn("w:noVBand"), "1")
    ctr = WD_ALIGN_PARAGRAPH.CENTER
    for ci in range(4): shade(tb.rows[0].cells[ci], T["navy"])
    cell_text(tb.rows[0].cells[0], "Scope", "Table Header")
    for i, o in enumerate(opts, 1):
        lab = o["name"] + ("  (recommended)" if o.get("recommended") else "")
        cell_text(tb.rows[0].cells[i], lab, "Table Header", align=ctr)
        if o.get("recommended"): shade(tb.rows[0].cells[i], T["accent"])
    cell_text(tb.rows[1].cells[0], "Summary", "Table Text", True)
    for i, o in enumerate(opts, 1): cell_text(tb.rows[1].cells[i], o["tagline"], "Table Text", align=ctr)
    for r_i, row in enumerate(pr["rows"]):
        r = tb.add_row()
        cell_text(r.cells[0], row[0])
        for i, v in enumerate(row[1:], 1):
            cell_text(r.cells[i], "✓" if v is True else ("–" if v is False else str(v)), align=ctr, bold=v is True,
                      color=T["accent"] if v is True else "9AA3AF")
    for label, key, fmt in (("Duration", "weeks", None), ("Fixed fee", "price", "m")):
        r = tb.add_row()
        cell_text(r.cells[0], label, "Table Text", True)
        for i, o in enumerate(opts, 1):
            cell_text(r.cells[i], money(o[key], cur) if fmt else o[key], "Table Text", True, align=ctr)
        if fmt:
            for cc in r.cells: shade(cc, T["accent_light"])
    for r in tb.rows:
        r._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
    set_widths(tb, w)
    for r in tb.rows[:-1]:
        for c in r.cells:
            for q in c.paragraphs: q.paragraph_format.keep_with_next = True
    # highlight recommended column
    for r in tb.rows[1:]:
        for i, o in enumerate(opts, 1):
            if o.get("recommended") and r is not tb.rows[-1]: shade(r.cells[i], "EAF5F6")
    para(doc, "", "Small Note")
    H2("Payment schedule")
    para(doc, pr["payment_schedule"])
    para(doc, pr.get("notes", ""), "Small Note")
    # 7 Terms
    H1("Terms and conditions")
    para(doc, "The clauses below are placeholders to be completed by each party's legal counsel before signature. Square brackets mark text that must be replaced.")
    w = [4.4, 12.2]
    tb = mk_table(doc, ["Clause", "Text"], w, T)
    for a, b in spec["terms"]: add_row(tb, [a, b], w, T, bold_first=True)
    # 8 Acceptance
    H1("Acceptance")
    para(doc, f'By signing below, {m["client"]} accepts this proposal and statement of work, including the option selected, and authorises {m["provider"]} to proceed.')
    sg = spec["signature"]
    q = para(doc, "", "Normal"); r = q.add_run("Option selected:  "); r.bold = True; r.font.color.rgb = rgb(T["accent"])
    q.add_run("[  ] Good      [  ] Better      [  ] Best")
    para(doc, "", "Small Note")
    sig = doc.add_table(rows=5, cols=2); sig.autofit = False; sig.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell_margins(sig, 80, 80, 100, 200)
    heads = (m["provider"], m["client"])
    names = (sg["provider_signatory"], sg["client_signatory"]); titles = (sg["provider_title"], sg["client_title"])
    for ci in range(2):
        shade(sig.rows[0].cells[ci], T["navy"]); cell_text(sig.rows[0].cells[ci], heads[ci], "Table Header")
        for ri, (lab, val) in enumerate((("Signature", ""), ("Name", names[ci]), ("Title", titles[ci]), ("Date", "")), 1):
            c = sig.rows[ri].cells[ci]
            cell_text(c, lab.upper(), "Small Note", True, color=T["accent"])
            q = c.add_paragraph(val if val else " ", "Table Text")
            if lab == "Signature": q.paragraph_format.space_before = Pt(14); q.paragraph_format.space_after = Pt(8)
            # underline via bottom border
            tcPr = c._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:tcBorders {nsdecls("w")}><w:bottom w:val="single" w:sz="4" w:color="{T["line"]}"/></w:tcBorders>'))
    for r in sig.rows:
        r._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for i in range(2): r.cells[i].width = Cm(8.3)
    set_widths(sig, [8.3, 8.3])

    fix_order(doc.element); fix_order(doc.styles.element)
    docx_path = os.path.join(out_dir, base + ".docx"); doc.save(docx_path)
    os.remove(img)
    if pdf:
        with tempfile.TemporaryDirectory() as td:
            subprocess.run(["soffice", f"-env:UserInstallation=file://{td}/p", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
                           check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return docx_path


def qa(spec):
    """Report leftover placeholders (other than intentional ones in terms) and price sanity."""
    issues = []
    def walk(x, path=""):
        if isinstance(x, dict):
            for k, v in x.items(): walk(v, f"{path}.{k}")
        elif isinstance(x, list):
            for i, v in enumerate(x): walk(v, f"{path}[{i}]")
        elif isinstance(x, str) and PLACEHOLDER.search(x) and not path.startswith(".terms"):
            issues.append(f"placeholder in {path}: {x[:60]}")
    walk(spec)
    prices = [o["price"] for o in spec["pricing"]["options"]]
    if prices != sorted(prices): issues.append("option prices are not ascending")
    for o in spec["pricing"]["options"]:
        if not o.get("tagline"): issues.append(f"option {o['name']} has no tagline")
    return issues


if __name__ == "__main__":
    ap = argparse.ArgumentParser(); ap.add_argument("spec"); ap.add_argument("out_dir")
    ap.add_argument("--theme", default=os.path.join(HERE, "..", "templates", "theme.json")); ap.add_argument("--no-pdf", action="store_true"); ap.add_argument("--name", default="proposal", help="output file basename")
    a = ap.parse_args()
    spec = json.load(open(a.spec)); T = json.load(open(a.theme))
    for i in qa(spec): print("QA WARNING:", i, file=sys.stderr)
    print(build(spec, a.out_dir, T, not a.no_pdf, a.name))
