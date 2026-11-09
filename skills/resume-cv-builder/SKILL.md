---
name: resume-cv-builder
description: Build an ATS-friendly resume or CV as an editable Word (.docx) and a PDF by interviewing the user, then filling a structured JSON. Two layouts (one-column ATS-safe, two-column modern sidebar), three themes, consistent typography, automatic page-length control (1 or 2 pages), bullet-quality lint (action verbs, quantified results, no pronouns or cliches) and overflow/orphan checks, plus optional job-description keyword gap report. Use when someone asks to write, create, build, improve, format or tailor a resume, CV, or job application document. Never invents achievements.
---

# Resume / CV Builder

Output: `<name>.docx` (real paragraph styles, selectable text, no images, no text boxes) and `<name>.pdf` (rendered from the docx by LibreOffice, so the two always match). Default A4; `--letter` for US Letter.

## Workflow

1. **Interview, do not invent.** Ask in small batches, not all at once:
   - Target: job title, industry, country (CV vs resume norms), 1 or 2 pages, and a job posting if there is one (save as `job.txt`).
   - Contact: name, email, phone, city, 1-2 links (LinkedIn, portfolio). No photo, age, marital status or full street address (ATS and bias reasons) unless the user's country requires it and they insist.
   - Per role (newest first): title, employer, city, start and end month/year, then "What did you own? What changed because of you? Numbers: how many, how much, how fast, how often, what before and after?" Collect 3-5 bullets for recent roles, 2-3 for older ones.
   - Education, skills actually used (they should be able to talk about each), certifications, languages, optional projects.
   **If the user has no numbers, ask for scale (team size, users, budget, volume, frequency) or a before/after estimate and label it as an estimate ("about", "roughly"). If they have nothing, write a clear scope-and-action bullet without a fake metric. Never fabricate employers, dates, degrees, tools, numbers or outcomes.** If improving an existing CV, keep every fact and only tighten wording; flag anything vague as a question.
2. **Fill the JSON** following `examples/resume-one-column.json` (keys: `layout`, `theme`, `max_pages`, `basics{name,headline,email,phone,location,links[]}`, `summary`, `experience[{role,company,location,start,end,bullets[]}]`, `projects[{name,description}]`, `skills[{group,items[]}]`, `education[{degree,school,year,details}]`, `certifications[]`, `languages[]`). Dates as `Mar 2022` and `Present`. Experience newest first.
3. **Build:** `python scripts/build_resume.py resume.json out_dir [--name stem] [--layout one-column|two-column] [--theme navy|teal|graphite|file.json] [--max-pages 1|2] [--job job.txt] [--letter] [--no-pdf]`. Needs `pip install -r requirements.txt` and LibreOffice (`soffice`) plus `pdftotext/pdfinfo` (poppler) for page checks.
4. **QA** (below), fix JSON, rebuild until clean. Deliver both files and list what you assumed or the user should double-check.

## Choosing a layout

- **one-column** (default): single flow, standard headings (Professional Summary, Experience, Skills, Education), right-aligned dates via tab stop. Safest for ATS portals, recruiters' systems and email parsing. Use this when applying through a form or job board.
- **two-column**: shaded sidebar (contact, skills, education) with main column (profile, experience). Reads quickly for people, so use for direct sends to hiring managers, portfolios and design/creative roles. Tell the user that some ATS parsers read tables poorly and that they should also keep a one-column version for portals. The reading order in the file is sidebar first, then main.

## Page length rules

- 1 page: under about 10 years of experience or career changers. 2 pages: more than 10 years, academic, or senior roles. Never 1.2 pages: the script shrinks body type stepwise from 10.5 to 9 pt and tightens margins to hit `max_pages`; if it still overflows it exits with code 2. Prefer cutting content over dropping below 10 pt: delete the oldest or weakest bullets, merge roles older than 10-15 years into one line, remove the skills that are not in the job posting.
- On 2 pages, page 2 should be at least a quarter full (the QA warns about a few-lines orphan page).

## Bullet-quality checklist (apply before building)

Formula: **action verb + what you did + how/with what + measurable result.**

1. Starts with a strong past-tense verb (Led, Built, Reduced, Launched, Negotiated, Automated, Redesigned, Cut, Grew, Migrated). Current role may use present tense consistently. Avoid "Responsible for", "Worked on", "Helped with", "Assisted", "Duties included".
2. Contains a number whenever the user truly has one: %, money, time saved, volume, team size, frequency. At least half the bullets quantified; never fabricate, never round up.
3. Shows outcome, not task: "Rebuilt the dashboard, cutting reporting time from 6 hours to 45 minutes" beats "Maintained dashboard".
4. One idea, at most 2 lines (about 25-30 words); varied openers (no verb on more than 2 bullets); no first-person pronouns; consistent full stops; no cliches ("team player", "results-driven", "dynamic").
5. 3-5 bullets for the last 2 roles, 2-3 for older ones. Most relevant bullet first.
6. Specific tools and methods named once in context ("SQL", "Looker", "A/B testing") so ATS keyword matching works; use the same wording as the job posting only where it is true.
7. Summary: 2-3 sentences (25-60 words), role + years + 2 strengths + target; no adjectives without evidence.
8. Spelling: one regional variant throughout (UK or US), consistent date format, consistent tense.

The script's `BULLET WARNING` lines implement items 1, 2, 4 and 5 automatically; you handle 3, 6, 7 and 8 by reading.

## Typography and design

One font family per file (Calibri, rendered as Carlito in LibreOffice; Arial in the `graphite` theme), name at body+14 pt, section headings small-caps-style uppercase with letter spacing and an accent rule, body 10-10.5 pt, 1.0-1.06 line spacing, hanging-indent bullets, one accent colour on grey text. No icons, photos, progress bars or skill ratings (ATS cannot read them and they carry no information). Themes live in `templates/themes.json`; a custom theme is a JSON with `font, accent, text, muted, sidebar` hex colours.

## QA checklist (mandatory)

1. Script output: no `OVERFLOW`; read every `QA WARNING`, `BULLET WARNING` and `LAYOUT` line; fix or justify each.
2. Render and look: `pdftoppm -png -r 60 out.pdf /tmp/pv` then view every page. Check no clipped words, dates aligned on the right edge, no heading stranded at a page bottom, no near-empty last page, sidebar not overflowing, consistent spacing between roles.
3. ATS check: `pdftotext out.pdf -` should read in sensible order, with email and phone intact and every role, employer and date present. If it jumbles, use the one-column layout.
4. `--job job.txt` lists frequent job-posting words missing from the resume. Add a keyword only if the user has genuinely done it; otherwise leave it out and tell the user it is a gap.
5. Search for leftovers: brackets, TBD, lorem, empty fields; inspect the docx opens (`python -c "import docx; docx.Document('x.docx')"`).
6. Ask the user to confirm every date, title and number once; they own the claims.

## Notes

Fictional examples use made-up names and employers; replace everything. File name: `Firstname-Lastname-Resume.pdf` is the convention recruiters expect.
