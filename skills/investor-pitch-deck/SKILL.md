---
name: investor-pitch-deck
description: Build a 13-slide investor pitch deck (.pptx) with native, editable PowerPoint charts, following the classic seed/Series A narrative (problem, solution, why now, TAM/SAM/SOM with a bottom-up calculation, product, traction, business model, competition matrix, go-to-market, team, financial summary, the ask and use of funds). Use when a founder, operator or analyst asks for a pitch deck, fundraising deck, investor presentation or seed/Series A deck from a brief, notes or metrics. Auto-fits text, offers three colour themes, computes market sizes and use-of-funds amounts so numbers stay consistent, and runs a QA pass.
---

# Investor pitch deck

Produces a 13-slide `.pptx` from a JSON content file. Charts are native python-pptx chart objects (editable in PowerPoint/Keynote/Google Slides), not images.

## Narrative (fixed order, one message per slide)
1 Title · 2 Problem · 3 Solution · 4 Why now · 5 Market (TAM/SAM/SOM + bottom-up table) · 6 Product · 7 Traction (column chart + 3 KPIs) · 8 Business model (tiers + unit economics) · 9 Competition (capability matrix) · 10 Go-to-market (3 phases + metrics) · 11 Team · 12 Financial summary (chart + table) · 13 The ask (use-of-funds doughnut). Titles are full-sentence takeaways, not labels ("MRR grew 11x in 12 months", not "Traction").

## Workflow
1. **Collect inputs. Never invent facts or numbers about the user's business.** Read the brief (see `examples/input.md` for the level of detail needed). If any of these are missing, ask the user in one batched question: company + one-line pitch, round size and use of funds, MRR/revenue series, customer count and retention, account counts and ACV for market sizing, pricing, 3-5 named competitors and capabilities, team bios, 5-year revenue/cost forecast. Anything still unknown stays as a clearly marked placeholder such as `[TBD: monthly MRR series]`; the build script warns about these and QA fails on them, so the user cannot send the deck unfinished by accident. Do not "fill in" plausible numbers. Say which numbers are user-provided vs. derived.
2. **Write the content spec** `deck.json` (copy `examples/deck.json` and replace every value). Keep titles <= ~90 characters, bullets <= ~15 words, 3 points per list. Numbers that must agree are computed by the script, so give inputs, not results: market `accounts` and `acv` (value = accounts x ACV, checked TAM >= SAM >= SOM), `use_of_funds` percentages (must sum to 100; amounts derived from `amount_m`), `revenue`/`costs` (EBITDA, break-even year and cumulative burn derived; warns if burn exceeds the raise). Titles may use tokens `{arr} {mrr_last} {growth} {months} {tam_title} {som_title} {rev_last} {breakeven} {raise}`.
3. **Pick a theme**: `python scripts/build_deck.py --list-themes` (`boardroom` navy/teal default, `ember` warm charcoal/orange, `forest` green/sage). Define more palettes by adding entries to `templates/themes.json` (keys: bg, surface, ink, muted, line, primary, accent, accent2, series[5], good/warn/bad).
4. **Build**: `python scripts/build_deck.py --spec deck.json --out pitch.pptx --theme boardroom`. Exit code 2 prints every slide whose text did not fit (even at the minimum font size) or whose numbers are inconsistent. Fix by shortening text or correcting inputs; never by shrinking fonts by hand.
5. **QA (mandatory)**: `python scripts/qa_pptx.py pitch.pptx --min-slides 12 --max-slides 14 --min-charts 3`. It re-opens the file and checks shapes inside the slide, estimated text overflow in every box, text running into the footer, placeholders, tiny fonts (<9pt), blank slides and the native chart count. Then render and LOOK: `python scripts/render_preview.py pitch.pptx previews --dpi 60 --sheet previews/sheet.png --cols 3` (needs LibreOffice + poppler) and view the PNGs. Check overlaps, cramped or orphaned text, chart label collisions, and that every number on a slide matches the other slides (ARR vs MRR x 12, round size in title vs use of funds). Fix and rebuild until clean.
6. Deliver the `.pptx`, state which numbers were supplied by the user, which are derived, and any placeholders left. Speaker notes can be added per slide through the `notes` object in the spec.

## Design system (enforced by `scripts/deckkit.py`)
- 13.33x7.5 in, 0.6 in margins, one accent bar + kicker + 2-line max title, footer with company, round and page number.
- Calibri (renders as Carlito/Calibri-metric where Calibri is absent). Title 20-30pt, body 10-22pt auto-fitted; the kit measures text with real font metrics, shrinks to a floor, then raises an error rather than overflow.
- Restrained palette: ink on white/surface cards, one accent, one highlight colour; chart series come from the theme; the latest traction bar is highlighted.
- No clip-art or stock imagery. Add a logo or screenshot manually if the user supplies one (kit has no network and ships no images).

## Rules of thumb for good content
Problem = 3 pains + one hard stat with a source; solution = what changes for the buyer; why now = 3 dated shifts; market sized bottom-up (accounts x price), never "1% of a huge number"; traction = the one curve that matters; competition = 5 capabilities that the buyer cares about, honest partial scores; ask = amount, 3-4 uses summing to 100%, milestones the money reaches.

## Files
`scripts/build_deck.py` (builder) · `scripts/deckkit.py` (themes, auto-fit, charts, tables) · `scripts/qa_pptx.py` · `scripts/render_preview.py` · `templates/themes.json` · `examples/` (fictional "Harborlane" brief, spec, output.pptx, previews).
