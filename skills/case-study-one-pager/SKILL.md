---
name: case-study-one-pager
description: Build a one-page customer case study as a print-ready PDF (reportlab) plus a matching single-slide PowerPoint (.pptx). Layout is challenge, solution, three big result callouts, a customer quote and a call to action. Use when a marketer, founder, sales or customer-success person asks for a case study, customer story, success story or one-pager from a customer interview, notes or before/after metrics. Computes percentage changes and multiples from before/after values, auto-fits text, and offers three colour themes.
---

# Customer case study one-pager

Output: `case.pdf` (US Letter, 1 page) and `case.pptx` (1 slide, 16:9, editable) from one JSON spec, so the PDF can be sent or printed and the slide dropped into a sales deck.

## Structure (fixed, top to bottom)
Header band (customer, headline that states the headline result, 3 customer facts) -> Challenge | Solution (two columns) -> Results (exactly 3 metric callouts) -> Quote with name and role -> CTA band (headline, one sentence, button text, URL).

## Workflow
1. **Collect inputs; never invent facts, metrics or quotes.** Needed: customer name and a few facts (industry, size, time to value); the situation before; what was implemented; 3 results with before and after values; a real, approved quote with the speaker's name and role; the CTA and URL. Ask once for what is missing. Anything unknown stays a visible placeholder like `[TBD: approved customer quote]`; the build warns and QA fails on it. Remind the user that customer names, numbers and quotes need the customer's written approval before publishing. Never fabricate a testimonial.
2. **Write `case.json`** (copy `examples/case.json`). Give metrics as `{"before": 1.9, "after": 1.1, "mode": "pct_change", "label": "..."}` (renders -42%) or `{"mode": "multiple"}` (renders 20x), or a literal `{"value": "9 days", "label": "..."}` for numbers the user states directly. The headline may use `{m1} {m1_abs} {m2} {m2_abs} {m3} {m3_abs}` so headline and callouts cannot disagree. Lead the headline with the strongest result; keep it under ~100 characters. Challenge: 2-4 sentences. Solution: 3-4 bullets (or one paragraph). Labels: one short line stating what the number means.
3. **Build**: `python scripts/build_case_study.py --spec case.json --pdf case.pdf --pptx case.pptx --theme boardroom` (`--list-themes`: boardroom, ember, forest; edit `templates/themes.json` for more). Text auto-shrinks to a floor (9pt PDF body, 10pt slide); if it still does not fit the script exits 2 and names the block to shorten.
4. **QA (mandatory)**: `python scripts/qa_pdf.py case.pdf --spec case.json` (1 page, all words inside safe margins, no placeholders, headline/metrics/quote/CTA present) and `python scripts/qa_pptx.py case.pptx --min-slides 1 --max-slides 1`. Then LOOK at both: `pdftoppm -png -r 70 case.pdf prev` and `python scripts/render_preview.py case.pptx previews --dpi 60 --pages 1`; view the PNGs. Check cramped or colliding text, orphaned lines, and that the numbers in the body match the callouts. Fix and rebuild.
5. Deliver both files; state which figures were computed from before/after values and note anything awaiting customer approval.

## Design notes
Calibri-metric (Carlito) type, one accent colour, large numerals for the callouts, generous white space, no stock imagery or logos (the user can add a logo in the PPTX). The PDF embeds its fonts.

## Files
`scripts/build_case_study.py` · `scripts/qa_pdf.py` · `scripts/qa_pptx.py` · `scripts/deckkit.py` (shared slide kit) · `scripts/render_preview.py` · `templates/themes.json` · `examples/` (fictional Brightwell Foods case, output.pdf, output.pptx, previews).
