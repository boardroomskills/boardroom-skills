---
name: proposal-sow
description: Create a client-ready proposal and statement of work as an editable Word (.docx) file plus a PDF, with cover, executive summary, understanding of needs, phased approach with a Gantt-style timeline, deliverables table, assumptions, good/better/best pricing, terms placeholders and a signature block. Use when a consultant, agency or freelancer asks for a proposal, SOW, scope of work, quote with options, or engagement letter. Not for contracts or legal drafting.
---

# Proposal + Statement of Work

Produces `<name>.docx` (real Word styles: Heading 1-3, Title, List Bullet/Number, a custom "Boardroom Table" table style, page-number footer) and `<name>.pdf` (rendered from the docx by LibreOffice), about 5-7 pages.

## Workflow

1. **Collect inputs. Ask before inventing.** Required: provider name, client name and contact, the client's problem in their own words, scope, timeline/start date, and *prices for each option*. If the user gives no prices, ask; never make up fees, durations, client metrics or dates. If the user wants to proceed without answers, use a clearly bracketed placeholder such as `[price to be confirmed]` in the text and tell them. Prices must be numbers in the JSON, so for unknown prices stop and ask.
2. **Write the spec JSON** following `examples/proposal.json` (the schema is that file; keys: `meta`, `executive_summary`, `executive_bullets`, `understanding`, `approach.phases[start,end in weeks]`, `deliverables`, `assumptions`, `pricing`, `terms`, `signature`). Use `examples/input.md` to see a brief and how it maps.
3. **Build:** `python scripts/build_proposal.py spec.json out_dir [--name client-proposal]` (needs `pip install -r requirements.txt` and LibreOffice `soffice` on PATH for the PDF; use `--no-pdf` to skip). Optional `--theme templates/theme.json` (copy and edit the colours; system fonts only, default Calibri which renders as Carlito in LibreOffice).
4. **QA** (mandatory, see below), fix the JSON, rebuild.
5. Hand the user both files and list every placeholder and assumption they must confirm.

## Content rules (follow the proposal framework)

- **Executive summary first, in the client's terms:** problem, proposed solution, outcome, price range. Two short paragraphs plus an "At a glance" box of 3 bullets. Outcomes the client stated are labelled as client targets, not guarantees.
- **Understanding of needs** reflects what the client said, with a table of need / what we heard. State that it will be validated in the first phase.
- **Approach:** 3-5 phases, each with weeks, description, 3 activities and a milestone. Phases may overlap; the Gantt is drawn from `start`/`end`.
- **Deliverables:** numbered D1..Dn, each tied to a phase and the lowest option that includes it. Every deliverable must be checkable (what, form, acceptance).
- **Assumptions:** 5-8 specific, testable statements (access, approvals, availability, exclusions, travel).
- **Pricing:** exactly three options (Good/Better/Best) as fixed fees; one marked `recommended`, with a reason. Rows are scope items with true/false per option, so options are cumulative and nothing is hidden. Prices ascend. Add the payment schedule and validity date.
- **Terms:** legal clauses stay as clearly bracketed placeholders `[... - to be inserted]`. Do not write legal terms; tell the user their lawyer should complete them.
- **Acceptance:** option checkboxes and two signature columns. Leave the client signatory blank.
- Tone: plain, specific, no filler adjectives. Short sentences. Numbers consistent everywhere (total weeks, prices, phase names).

## Design system

Navy `1F2A44` for headings and table headers, teal `0E7C86` for accents, light grey banding, one typeface. Numbered H1 sections with a teal rule; tables use a thin top/bottom rule and banded rows; a recommended-option column is highlighted. Do not add clip-art or more than two accent colours.

## QA checklist (do not skip)

1. Run the script and read its `QA WARNING` lines (leftover placeholders outside the terms table, non-ascending prices).
2. Render: `pdftoppm -png -r 60 out_dir/<name>.pdf /tmp/pv` and view every page image. Check: no clipped text or tables, no heading stranded at the bottom of a page, Gantt legible, cover not overflowing, signature block on one page, columns aligned.
3. `pdftotext out_dir/<name>.pdf -` and grep for `TBD|TODO|lorem|\[`: only the intentional bracketed term placeholders may remain.
4. Cross-check numbers: option prices in the pricing table match the exec summary bullets and payment schedule; total weeks match the Gantt; deliverable refs match the pricing rows.
5. Open the docx once with python-docx (`Document(path)`) to confirm it loads. If a layout issue appears, adjust page breaks or shorten text in the spec, rebuild and re-render.
