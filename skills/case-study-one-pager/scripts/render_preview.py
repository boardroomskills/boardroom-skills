#!/usr/bin/env python3
"""Render a .pptx (or .pdf) to PNG pages and an optional contact sheet.

usage: render_preview.py FILE.pptx OUTDIR [--dpi 60] [--sheet sheet.png] [--cols 3] [--pages 1,5,7]
Needs LibreOffice (soffice) and poppler (pdftoppm). Keeps previews small on purpose.
"""
import argparse, glob, os, shutil, subprocess, sys, tempfile
from PIL import Image


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file"); ap.add_argument("outdir")
    ap.add_argument("--dpi", type=int, default=60)
    ap.add_argument("--sheet"); ap.add_argument("--cols", type=int, default=3)
    ap.add_argument("--pages", help="comma list of 1-based pages to keep as individual PNGs")
    a = ap.parse_args()
    os.makedirs(a.outdir, exist_ok=True)
    tmp = tempfile.mkdtemp()
    pdf = a.file
    if a.file.lower().endswith(".pptx"):
        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp, a.file],
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdf = os.path.join(tmp, os.path.splitext(os.path.basename(a.file))[0] + ".pdf")
    subprocess.run(["pdftoppm", "-png", "-r", str(a.dpi), pdf, os.path.join(tmp, "p")], check=True)
    pages = sorted(glob.glob(os.path.join(tmp, "p-*.png")))
    print("%d pages rendered" % len(pages))
    if a.pages:
        for n in [int(x) for x in a.pages.split(",")]:
            shutil.copy(pages[n - 1], os.path.join(a.outdir, "preview-%02d.png" % n))
    if a.sheet:
        ims = [Image.open(p).convert("RGB") for p in pages]
        w, h = ims[0].size
        cols = a.cols; rows = (len(ims) + cols - 1) // cols
        g = 8
        sheet = Image.new("RGB", (cols * w + (cols + 1) * g, rows * h + (rows + 1) * g), (200, 204, 210))
        for i, im in enumerate(ims):
            sheet.paste(im, (g + (i % cols) * (w + g), g + (i // cols) * (h + g)))
        sheet.save(a.sheet, optimize=True)
        print("sheet", a.sheet, sheet.size)
    shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
