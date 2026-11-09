#!/usr/bin/env python3
"""Build an ATS-friendly resume/CV (.docx + .pdf) from a structured JSON.

Usage: python build_resume.py resume.json out_dir [--name resume] [--layout one-column|two-column]
                              [--theme navy|teal|graphite|file.json] [--max-pages 1|2] [--job job.txt] [--no-pdf]
The PDF is rendered from the docx by LibreOffice (soffice). The script auto-shrinks type (never below 9 pt)
to honour max_pages, then lints bullets and checks overflow / orphans in the rendered PDF.
Exit code 2 when the page limit cannot be met.
"""
import argparse, collections, json, os, re, shutil, subprocess, sys, tempfile
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, RGBColor, Mm, Emu

HERE = os.path.dirname(os.path.abspath(__file__))
rgb = RGBColor.from_string
SIZES = [10.5, 10.25, 10, 9.75, 9.5, 9.25, 9]      # body sizes tried in order (pt)
PAGE_W, PAGE_H = 210, 297                           # A4 mm; --letter switches to US Letter
WEAK = re.compile(r"^(responsible for|duties (included|included:)|worked on|helped( with)?|assisted( with)?|tasked with|in charge of|involved in|participated in)\b", re.I)
PRONOUN = re.compile(r"\b(i|my|me|we|our)\b", re.I)
BUZZ = re.compile(r"\b(team player|hard[- ]working|results[- ]driven|go-getter|synergy|detail[- ]oriented|self[- ]starter|think outside the box|dynamic)\b", re.I)
PLACEHOLDER = re.compile(r"\[[^\]]*\]|\bTBD\b|\bTODO\b|lorem|\bXX+\b|\bN/A\b", re.I)
STOP = set("the and for with that this from have has are was were will you your our their they them its into over about more most than then also such can able work working experience years year team teams role roles skills skill strong ability including include includes etc must should would could using use used based".split())


# ------------------------------------------------------------------ docx helpers
def shade(cell, fill):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/>'
                          f'<w:bottom w:w="{bottom}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'))


def bottom_border(p, color, sz=6, space=1):
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="{sz}" w:space="{space}" w:color="{color}"/></w:pBdr>'))


def fmt(p, before=0, after=0, line=1.0, keep_next=False, keep_lines=False):
    f = p.paragraph_format
    f.space_before, f.space_after, f.line_spacing = Pt(before), Pt(after), line
    f.keep_with_next, f.keep_together = keep_next, keep_lines


def run(p, text, size, color, bold=False, italic=False, font=None):
    r = p.add_run(text); r.font.size = Pt(size); r.font.color.rgb = rgb(color); r.bold = bold; r.italic = italic
    if font:
        r.font.name = font
    return r


def dates(e):
    s, en = e.get("start", ""), e.get("end", "")
    return " - ".join(x for x in (s, en) if x) if s and en else (s or en or e.get("year", ""))


class Builder:
    def __init__(self, spec, th, size, layout, letter=False):
        self.s, self.th, self.sz, self.layout = spec, th, size, layout
        self.doc = Document()
        sec = self.doc.sections[0]
        w, h = (216, 279) if letter else (PAGE_W, PAGE_H)
        sec.page_width, sec.page_height = Mm(w), Mm(h)
        m = 15 if size <= 9.75 else 17
        two = layout == "two-column"
        sec.left_margin = sec.right_margin = Mm(0 if False else m)
        sec.top_margin, sec.bottom_margin = Mm(m - 2), Mm(m - 2)
        self.text_w = Mm(w - 2 * m)
        n = self.doc.styles["Normal"]; n.font.name = th["font"]; n.font.size = Pt(size)
        n.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), th["font"])
        n.paragraph_format.space_after = Pt(0)
        b = spec["basics"]
        self.doc.core_properties.title = f"{b['name']} - Resume"; self.doc.core_properties.author = b["name"]
        bullet = self.doc.styles["List Bullet"]; bullet.font.name = th["font"]; bullet.font.size = Pt(size)

    # --- shared pieces
    def contact_items(self):
        b = self.s["basics"]
        return [x for x in [b.get("email"), b.get("phone"), b.get("location")] + list(b.get("links", [])) if x]

    def heading(self, container, text, first=False):
        p = container.add_paragraph() if container is not None else self.doc.add_paragraph()
        run(p, text.upper(), self.sz - 0.5, self.th["accent"], True).font.name = self.th["font"]
        p.runs[0]._r.get_or_add_rPr().append(parse_xml(f'<w:spacing {nsdecls("w")} w:val="20"/>'))
        fmt(p, before=0 if first else self.sz * 0.9, after=self.sz * 0.3, keep_next=True)
        bottom_border(p, self.th["accent"])
        return p

    def bullet(self, container, text, size=None):
        p = container.add_paragraph(style="List Bullet")
        run(p, text, size or self.sz, self.th["text"])
        fmt(p, after=self.sz * 0.2, line=1.04, keep_lines=True)
        pf = p.paragraph_format; pf.left_indent, pf.first_line_indent = Mm(4.2), Mm(-4.2)
        return p

    def experience(self, c, items, narrow=False):
        for i, e in enumerate(items):
            p = c.add_paragraph()
            run(p, e["role"], self.sz + 0.5, self.th["text"], True)
            if not narrow:
                p.paragraph_format.tab_stops.add_tab_stop(self.main_w if hasattr(self, "main_w") else self.text_w, WD_TAB_ALIGNMENT.RIGHT)
                run(p, "\t" + dates(e), self.sz, self.th["muted"])
            fmt(p, before=self.sz * 0.5 if i else 0, keep_next=True, keep_lines=True)
            p = c.add_paragraph()
            run(p, e["company"], self.sz, self.th["accent"], True)
            if e.get("location"): run(p, "  |  " + e["location"], self.sz, self.th["muted"])
            if narrow: run(p, "  |  " + dates(e), self.sz, self.th["muted"])
            fmt(p, after=self.sz * 0.25, keep_next=True, keep_lines=True)
            for t in e.get("bullets", []): self.bullet(c, t)

    def education(self, c, items):
        for e in items:
            p = c.add_paragraph(); run(p, e["degree"], self.sz, self.th["text"], True); fmt(p, before=self.sz * 0.2, keep_next=True)
            p = c.add_paragraph(); run(p, e["school"] + (f", {e['year']}" if e.get("year") else ""), self.sz - 0.5, self.th["muted"]); fmt(p, after=self.sz * 0.15)
            if e.get("details"):
                p = c.add_paragraph(); run(p, e["details"], self.sz - 0.5, self.th["muted"]); fmt(p, after=self.sz * 0.15)

    def projects(self, c, items):
        for pr in items:
            p = c.add_paragraph(); run(p, pr["name"], self.sz, self.th["text"], True)
            if pr.get("description"): run(p, "  " + pr["description"], self.sz, self.th["text"])
            fmt(p, after=self.sz * 0.25, line=1.04, keep_lines=True)

    def skills_lines(self, c, groups, stacked=False):
        for g in groups:
            p = c.add_paragraph()
            if stacked:
                run(p, g["group"], self.sz - 0.5, self.th["accent"], True); fmt(p, before=self.sz * 0.3, keep_next=True)
                p = c.add_paragraph(); run(p, ", ".join(g["items"]), self.sz - 0.5, self.th["text"]); fmt(p, after=self.sz * 0.15, line=1.04)
            else:
                run(p, g["group"] + ": ", self.sz, self.th["text"], True); run(p, ", ".join(g["items"]), self.sz, self.th["text"])
                fmt(p, after=self.sz * 0.2, line=1.04)

    def simple_list(self, c, items, joiner=None):
        if joiner:
            p = c.add_paragraph(); run(p, joiner.join(items), self.sz, self.th["text"]); fmt(p, line=1.04); return
        for t in items:
            p = c.add_paragraph(); run(p, t, self.sz - (0.5 if self.layout == "two-column" else 0), self.th["text"]); fmt(p, after=self.sz * 0.15, line=1.04)

    # --- layouts
    def build_one(self):
        s, b, th = self.s, self.s["basics"], self.th
        d = self.doc
        p = d.add_paragraph(); run(p, b["name"], self.sz + 14, th["accent"], True); fmt(p, after=1, keep_next=True)
        if b.get("headline"):
            p = d.add_paragraph(); run(p, b["headline"], self.sz + 1.5, th["text"]); fmt(p, after=2)
        p = d.add_paragraph(); run(p, "  |  ".join(self.contact_items()), self.sz - 0.5, th["muted"]); fmt(p, after=2)
        bottom_border(p, th["accent"], 12, 4)
        if s.get("summary"):
            self.heading(None, "Professional summary")
            p = d.add_paragraph(); run(p, s["summary"], self.sz, th["text"]); fmt(p, line=1.06, keep_lines=True)
        if s.get("experience"):
            self.heading(None, "Experience"); self.experience(d, s["experience"])
        if s.get("projects"):
            self.heading(None, "Projects"); self.projects(d, s["projects"])
        if s.get("skills"):
            self.heading(None, "Skills"); self.skills_lines(d, s["skills"])
        if s.get("education"):
            self.heading(None, "Education"); self.education(d, s["education"])
        if s.get("certifications"):
            self.heading(None, "Certifications"); self.simple_list(d, s["certifications"], joiner="  |  ")
        if s.get("languages"):
            self.heading(None, "Languages"); self.simple_list(d, s["languages"], joiner="  |  ")

    def build_two(self):
        s, b, th = self.s, self.s["basics"], self.th
        d = self.doc
        total = self.text_w
        side_w = int(total * 0.31)
        self.main_w = total - side_w - Mm(7)
        t = d.add_table(rows=1, cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.autofit = False
        tblPr = t._tbl.tblPr
        tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
        t.columns[0].width, t.columns[1].width = side_w, total - side_w
        left, right = t.cell(0, 0), t.cell(0, 1)
        page_h = self.doc.sections[0].page_height
        usable = page_h - self.doc.sections[0].top_margin - self.doc.sections[0].bottom_margin - Mm(6)
        trPr = t.rows[0]._tr.get_or_add_trPr()
        trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{int(usable / 635)}" w:hRule="atLeast"/>'))  # full-height sidebar
        left.width, right.width = side_w, total - side_w
        shade(left, th["sidebar"]); set_cell_margins(left, 160, 160, 160, 140); set_cell_margins(right, 160, 0, 300, 0)
        # name block (sidebar top)
        p = left.paragraphs[0]; run(p, b["name"], self.sz + 9, th["accent"], True); fmt(p, after=2, keep_next=True)
        if b.get("headline"):
            p = left.add_paragraph(); run(p, b["headline"], self.sz, th["text"], True); fmt(p, after=self.sz * 0.6)
        self.heading(left, "Contact", first=True)
        for x in self.contact_items():
            p = left.add_paragraph(); run(p, x, self.sz - 1, th["text"]); fmt(p, after=self.sz * 0.2)
        if s.get("skills"):
            self.heading(left, "Skills"); self.skills_lines(left, s["skills"], stacked=True)
        if s.get("education"):
            self.heading(left, "Education"); self.education(left, s["education"])
        if s.get("certifications"):
            self.heading(left, "Certifications"); self.simple_list(left, s["certifications"])
        if s.get("languages"):
            self.heading(left, "Languages"); self.simple_list(left, s["languages"])
        first = True
        if s.get("summary"):
            self.heading_in(right, "Profile", first); first = False
            p = right.add_paragraph(); run(p, s["summary"], self.sz, th["text"]); fmt(p, line=1.06, keep_lines=True)
        if s.get("experience"):
            self.heading_in(right, "Experience", first); first = False; self.experience(right, s["experience"])
        if s.get("projects"):
            self.heading_in(right, "Projects", first); self.projects(right, s["projects"])
        # remove the empty first paragraph in right cell
        fp = right.paragraphs[0]
        if not fp.text: fp._p.getparent().remove(fp._p)

    def heading_in(self, cell, text, first):
        return self.heading(cell, text, first=first)

    def build(self):
        (self.build_two if self.layout == "two-column" else self.build_one)()
        # drop trailing empty paragraph spacing after table
        if self.layout == "two-column":
            last = self.doc.paragraphs[-1] if self.doc.paragraphs else None
            if last is not None and not last.text:
                fmt(last, line=0.1); last.runs and None
                last.paragraph_format.space_before = Pt(0)
                for r in last.runs: r.font.size = Pt(1)
                pPr = last._p.get_or_add_pPr()
                rpr = parse_xml(f'<w:rPr {nsdecls("w")}><w:sz w:val="2"/></w:rPr>'); pPr.append(rpr)
        return self.doc


# ------------------------------------------------------------------ rendering and QA
def render_pdf(docx_path, out_dir):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path], capture_output=True, timeout=180)
    pdf = os.path.join(out_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
    return pdf if os.path.exists(pdf) else None


def pdf_pages(pdf):
    out = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout
    return [pg for pg in out.split("\f") if pg.strip() != "" or False]


def page_count(pdf):
    out = subprocess.run(["pdfinfo", pdf], capture_output=True, text=True).stdout
    m = re.search(r"Pages:\s+(\d+)", out)
    return int(m.group(1)) if m else len(pdf_pages(pdf))


def lint_bullets(spec):
    w = []
    starts = collections.Counter(); total = metrics = 0
    for e in spec.get("experience", []):
        for t in e.get("bullets", []):
            total += 1
            first = re.sub(r"[^A-Za-z-]", "", t.split()[0] if t.split() else "").lower()
            starts[first] += 1
            tag = f"[{e['company'].split(' (')[0]}] '{t[:48]}...'"
            if WEAK.search(t): w.append(f"weak opener (use an action verb): {tag}")
            if PRONOUN.search(t): w.append(f"first-person pronoun: {tag}")
            if BUZZ.search(t): w.append(f"buzzword/cliche: {tag}")
            if re.search(r"\d", t): metrics += 1
            if len(t.split()) > 34: w.append(f"bullet over 34 words (cut to <= 2 lines): {tag}")
            if len(t.split()) < 7: w.append(f"bullet under 7 words (add what and result): {tag}")
            if first.endswith("ing") and first not in ("marketing", "engineering"): w.append(f"starts with -ing form (use past tense for past roles): {tag}")
            if not t.rstrip().endswith("."): w.append(f"bullet has no closing full stop (be consistent): {tag}")
    for v, n in starts.items():
        if n > 2 and v: w.append(f"action verb '{v}' opens {n} bullets; vary it")
    if total and metrics / total < 0.5:
        w.append(f"only {metrics}/{total} bullets contain a number; ask the user for scale, time, money or % outcomes (never invent)")
    for e in spec.get("experience", []):
        n = len(e.get("bullets", []))
        if n > 6: w.append(f"{e['role']}: {n} bullets, keep 3-5 (fewer for older roles)")
        if n < 2: w.append(f"{e['role']}: only {n} bullet")
    return w


def lint_content(spec):
    w = []
    b = spec.get("basics", {})
    for k in ("name", "email", "phone"):
        if not b.get(k): w.append(f"missing basics.{k}")
    if b.get("email") and not re.match(r"[^@\s]+@[^@\s]+\.[^@\s]+$", b["email"]): w.append("email looks invalid")
    sm = spec.get("summary", "")
    if sm and len(sm.split()) > 70: w.append(f"summary is {len(sm.split())} words; keep to about 25-60")
    if not spec.get("experience") and not spec.get("education"): w.append("no experience or education")
    def walk(o):
        if isinstance(o, str): yield o
        elif isinstance(o, dict): yield from (s for v in o.values() for s in walk(v))
        elif isinstance(o, list): yield from (s for v in o for s in walk(v))
    for sv in walk(spec):
        for m in PLACEHOLDER.findall(sv): w.append(f"placeholder left in text: {m!r}")
    # chronology: most recent first
    ys = []
    for e in spec.get("experience", []):
        m = re.search(r"(19|20)\d\d", e.get("start", ""))
        if m: ys.append(int(m.group(0)))
    if ys != sorted(ys, reverse=True): w.append("experience is not in reverse-chronological order")
    return w


def layout_qa(pdf, max_pages, spec):
    issues, n = [], page_count(pdf)
    txt = subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True).stdout.split("\f")
    pages = [p for p in txt if p.strip()]
    if n > max_pages: issues.append(f"OVERFLOW: {n} pages, limit {max_pages}")
    heads = {h.upper() for h in ("Professional summary", "Profile", "Experience", "Projects", "Skills", "Education", "Certifications", "Languages", "Contact")}
    for i, pg in enumerate(pages, 1):
        lines = [l.strip() for l in pg.splitlines() if l.strip()]
        if lines and lines[-1] in heads: issues.append(f"stranded heading at the bottom of page {i}: {lines[-1]}")
    if len(pages) > 1:
        last = [l for l in pages[-1].splitlines() if l.strip()]
        first_n = len([l for l in pages[0].splitlines() if l.strip()])
        if len(last) < 6: issues.append(f"ORPHAN: last page holds only {len(last)} line(s); trim content to fit {len(pages)-1} page(s) or add substance")
        elif len(last) < 0.25 * first_n and len(pages) > max_pages - 1 and n <= max_pages:
            issues.append(f"note: last page is only {len(last)} of ~{first_n} lines full; looks sparse")
    # every bullet starts present in output (catches silent drops)
    flat = re.sub(r"\s+", " ", " ".join(txt))
    for e in spec.get("experience", []):
        if e["role"] not in flat: issues.append(f"role missing from PDF text: {e['role']}")
    return n, issues


def job_keywords(spec, job_path):
    job = open(job_path, encoding="utf-8").read().lower()
    words = [w for w in re.findall(r"[a-z][a-z+#.\-]{2,}", job) if w not in STOP]
    common = [w for w, c in collections.Counter(words).most_common(40) if c >= 2]
    blob = json.dumps(spec).lower()
    return [w for w in common if w not in blob][:15]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec"); ap.add_argument("out_dir")
    ap.add_argument("--name"); ap.add_argument("--layout"); ap.add_argument("--theme")
    ap.add_argument("--max-pages", type=int); ap.add_argument("--job"); ap.add_argument("--no-pdf", action="store_true")
    ap.add_argument("--letter", action="store_true", help="US Letter instead of A4")
    a = ap.parse_args()
    spec = json.load(open(a.spec, encoding="utf-8"))
    layout = a.layout or spec.get("layout", "one-column")
    max_pages = a.max_pages or spec.get("max_pages", 2)
    tn = a.theme or spec.get("theme", "navy")
    th = json.load(open(tn)) if os.path.exists(tn) else json.load(open(os.path.join(HERE, "..", "templates", "themes.json")))[tn]
    name = a.name or re.sub(r"[^A-Za-z0-9_-]+", "-", spec["basics"]["name"]) + "-resume"
    os.makedirs(a.out_dir, exist_ok=True)
    docx_path = os.path.join(a.out_dir, name + ".docx")
    pdf, used = None, None
    for sz in SIZES:
        Builder(spec, th, sz, layout, a.letter).build().save(docx_path)
        used = sz
        if a.no_pdf: break
        with tempfile.TemporaryDirectory() as td:
            tmp = render_pdf(docx_path, td)
            if not tmp: print("NOTE: soffice not found; PDF skipped and page fit unchecked"); break
            if page_count(tmp) <= max_pages or sz == SIZES[-1]:
                shutil.copy(tmp, os.path.join(a.out_dir, name + ".pdf")); pdf = os.path.join(a.out_dir, name + ".pdf"); break
    problems = 0
    for w in lint_content(spec): print("QA WARNING:", w)
    for w in lint_bullets(spec): print("BULLET WARNING:", w)
    if pdf:
        n, issues = layout_qa(pdf, max_pages, spec)
        for i in issues:
            print("LAYOUT:", i); problems += i.startswith("OVERFLOW") or i.startswith("role missing")
        print(f"Pages: {n} (limit {max_pages}), body font {used} pt, layout {layout}")
        if used < 10: print("NOTE: type was shrunk below 10 pt to fit; consider cutting content instead")
    if a.job:
        miss = job_keywords(spec, a.job)
        if miss: print("KEYWORDS from job text not in resume (add ONLY if true for the user):", ", ".join(miss))
    print("Wrote", docx_path, "and", pdf or "(no pdf)")
    sys.exit(2 if problems else 0)


if __name__ == "__main__":
    main()
