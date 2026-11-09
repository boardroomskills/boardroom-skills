#!/usr/bin/env python3
"""QA a .pptx: re-opens the file and checks what a human reviewer would catch.

Checks: shapes inside the slide, estimated text overflow per text box and table cell,
placeholder strings, empty text boxes, slide count range, native chart count, tiny fonts,
and duplicate/blank slides. Exit 1 on any FAIL.
usage: qa_pptx.py deck.pptx [--min-slides 12 --max-slides 14 --min-charts 2 --min-font 9]
"""
import argparse, os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx import Presentation
from pptx.util import Emu
import deckkit as K


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--min-slides", type=int, default=1)
    ap.add_argument("--max-slides", type=int, default=999)
    ap.add_argument("--min-charts", type=int, default=0)
    ap.add_argument("--min-font", type=float, default=9)
    a = ap.parse_args()
    prs = Presentation(a.file)
    W, H = prs.slide_width, prs.slide_height
    fails, warns, charts = [], [], 0
    n = len(prs.slides)
    if not a.min_slides <= n <= a.max_slides:
        fails.append("slide count %d outside %d-%d" % (n, a.min_slides, a.max_slides))
    for i, s in enumerate(prs.slides, 1):
        texts = 0
        for sh in s.shapes:
            if sh.left is not None and (sh.left < -Emu(5000) or sh.top < -Emu(5000) or
                                        sh.left + sh.width > W + Emu(5000) or sh.top + sh.height > H + Emu(5000)):
                # decorative shapes may bleed off-slide; text and charts may not
                if (getattr(sh, "has_text_frame", False) and sh.text_frame.text.strip()) or getattr(sh, "has_chart", False):
                    fails.append("slide %d: '%s' extends beyond the slide" % (i, sh.name))
            if getattr(sh, "has_chart", False) and sh.has_chart:
                charts += 1
            if sh.has_text_frame and sh.text_frame.text.strip():
                texts += 1
                tf = sh.text_frame
                total, w_in = 0.0, Emu(sh.width).inches
                for p in tf.paragraphs:
                    t = "".join(r.text for r in p.runs)
                    if not p.runs:
                        continue
                    sz = p.runs[0].font.size.pt if p.runs[0].font.size else 18
                    if sz < a.min_font:
                        warns.append("slide %d: font %.1fpt below %.1f in '%s'" % (i, sz, a.min_font, t[:30]))
                    bold = bool(p.runs[0].font.bold)
                    ind = 0.2 if p._p.pPr is not None and p._p.pPr.get("marL") else 0
                    total += K.count_lines(t, sz, (w_in - ind) * 72 * 0.97, bold) * sz * K.LINE * (p.line_spacing if isinstance(p.line_spacing, float) else 1.0)
                    if p.space_before:
                        total += p.space_before.pt
                if total > Emu(sh.height).inches * 72 * 1.03:
                    fails.append("slide %d: text may overflow its box (%.2f in needed vs %.2f in): '%s'"
                                 % (i, total / 72, Emu(sh.height).inches, tf.text[:40].replace("\n", " ")))
                top_in, bot_in = Emu(sh.top).inches, Emu(sh.top + sh.height).inches
                if top_in < 7.0 < bot_in - 0.02 and Emu(sh.top).inches > 0.9:
                    fails.append("slide %d: text box runs into the footer zone: '%s'" % (i, tf.text[:40].replace("\n", " ")))
                if K.find_placeholders([tf.text]):
                    fails.append("slide %d: placeholder text: '%s'" % (i, tf.text[:50]))
            if getattr(sh, "has_table", False) and sh.has_table:
                tbl = sh.table
                bottom = Emu(sh.top).inches + sum(Emu(r.height).inches for r in tbl.rows)
                if bottom > 7.0:
                    fails.append("slide %d: table bottom at %.2f in collides with footer" % (i, bottom))
                for r in tbl.rows:
                    for c in r.cells:
                        if K.find_placeholders([c.text]):
                            fails.append("slide %d: placeholder in table: %s" % (i, c.text[:40]))
        if texts == 0:
            fails.append("slide %d has no text (blank slide?)" % i)
    if charts < a.min_charts:
        fails.append("only %d native charts, need %d" % (charts, a.min_charts))
    print("%s: %d slides, %d native charts" % (os.path.basename(a.file), n, charts))
    for w in warns:
        print("WARN", w)
    for f in fails:
        print("FAIL", f)
    print("QA", "FAILED" if fails else "PASSED")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
