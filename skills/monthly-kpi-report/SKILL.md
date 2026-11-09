---
name: monthly-kpi-report
description: Generate a multi-page PDF monthly management/KPI report in Pyramid Principle structure (answer first, three supporting points, then evidence) from a CSV of monthly metrics plus a short notes file. Headline page has 4-6 KPI tiles with value, change vs last month and vs target; then trend charts, commentary with charts, risks, next steps and an appendix data table. Use for monthly business reviews, board or investor updates, ops/finance/marketing KPI packs.
---

# Monthly KPI report (PDF, Pyramid Principle)

Page 1 states the answer, the three reasons, and 4-6 KPI tiles. Page 2: 12-month trend charts. Page 3: commentary (one block per reason, with its chart), risks, next steps. Last page: appendix table. Built with reportlab and matplotlib; no network.

## Inputs

1. **CSV** (wide format): `month` column as `YYYY-MM`, then one numeric column per metric key. Percentages as plain numbers (`62.4` = 62.4%). At least the reporting month and the month before it (12+ months gives the trend pages their full effect).
2. **Notes file** (markdown-ish, see `examples/input.md`):
   - top lines `company:`, `period: YYYY-MM`, `currency:`, optional `title:`, `prepared_by:`
   - `## Metrics`: one line each `key | Label | currency/percent/number | up/down | target | bar/line` (target and chart kind optional; `down` = lower is better, e.g. churn, CAC)
   - `## Headline`: the answer in 1-2 sentences
   - `## Key points`: 2-4 `###` headings (the supporting arguments; three is ideal), each followed by `metric: <key>` and a short paragraph
   - `## Risks` and `## Decisions and next steps`: bullets `- text | owner | due or mitigation`
   - Write numbers as placeholders so text can never disagree with the data: `{revenue}`, `{revenue.prev}`, `{revenue.mom}`, `{revenue.target}`, `{revenue.vs_target}`, `{period_label}`. Unknown keys abort the build.

## Workflow

1. **Ask for what is missing.** Get the CSV, the target for each metric, and the user's reading of what happened (causes, decisions needed). Never invent causes, owners, dates, targets or numbers. If the user has no commentary, draft it from the data only, mark causes as "to be confirmed", and say so.
2. **Compute the story first.** Look at MoM and vs-target for each metric. Decide the single answer (what the reader must know or decide), then group the evidence into at most three reasons (Pyramid Principle: conclusion first, grouped, ordered by importance). The headline must be a sentence with a verdict, not a label ("Revenue is 2% behind plan because churn spiked", not "Revenue update").
3. **Run:** `python3 scripts/build_report.py --csv data.csv --notes notes.md --out report.pdf`
4. **QA (mandatory).** `pdftoppm -png -r 60 report.pdf out/p` and open every page. Check: no text clipped or overlapping chart labels, tiles show right sign/colour (lower-is-better metrics inverted correctly), every number in prose matches the appendix table, no empty sections or placeholder text, page count sensible. Fix the notes or CSV and rebuild.
5. Tell the user which numbers are theirs and which wording is a draft.

## Design rules built into the script

- A4 portrait, 40pt margins, one grid; navy/teal palette with green and red used only for favourable/unfavourable, always paired with an arrow or sign (colour-blind safe).
- Tiles: label, large value, change vs last month (relative %, or pp for percentages), gap vs target, 12-month sparkline. Accent bar shows status vs target.
- Charts: one message each, direct label on the latest value, target shown as a dashed line plus diamond, no chart junk. Page footer with page X of Y.
- Fonts: built-in Helvetica (no font files needed).

Limits: one target per metric (for the reporting month); 4-6 headline metrics; wide CSV only.
