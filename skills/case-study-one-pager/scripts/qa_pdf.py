#!/usr/bin/env python3
"""QA for the one-page case study PDF. Needs poppler (pdfinfo, pdftotext).

Checks: exactly 1 page, every word inside the page margins, no placeholder text, and that the headline,
the 3 metric values, the quote and the CTA text are actually present. Exit 1 on failure.
usage: qa_pdf.py case.pdf --spec case.json
"""
import argparse, json, os, re, subprocess, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import deckkit as K
from build_case_study import metric_value


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("pdf"); ap.add_argument("--spec")
    a = ap.parse_args()
    fails = []
    info = subprocess.run(["pdfinfo", a.pdf], capture_output=True, text=True).stdout
    pages = int(re.search(r"Pages:\s+(\d+)", info).group(1))
    if pages != 1:
        fails.append("PDF has %d pages, expected 1" % pages)
    size = re.search(r"Page size:\s+([\d.]+) x ([\d.]+)", info)
    pw, ph = float(size.group(1)), float(size.group(2))
    bbox = subprocess.run(["pdftotext", "-bbox", a.pdf, "-"], capture_output=True, text=True).stdout
    words = re.findall(r'xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([^<]*)<', bbox)
    for x0, y0, x1, y1, w in words:
        if float(x0) < 20 or float(x1) > pw - 20 or float(y0) < 8 or float(y1) > ph - 8:
            fails.append("word outside safe margins: %r" % w)
    text = subprocess.run(["pdftotext", "-layout", a.pdf, "-"], capture_output=True, text=True).stdout
    flat = re.sub(r"\s+", " ", text)
    if K.find_placeholders([flat]):
        fails.append("placeholder text present")
    if a.spec:
        spec = json.load(open(a.spec))
        need = [spec["headline"][:25], spec["quote"]["name"], spec["cta"]["button"]] + [metric_value(m) for m in spec["results"]]
        for n in need:
            if n.replace("−", "-") not in flat.replace("−", "-"):
                fails.append("expected text missing: %r" % n)
    print("%s: %d page, %d words" % (os.path.basename(a.pdf), pages, len(words)))
    for f in fails[:20]:
        print("FAIL", f)
    print("QA", "FAILED" if fails else "PASSED")
    sys.exit(1 if fails else 0)


if __name__ == "__main__":
    main()
