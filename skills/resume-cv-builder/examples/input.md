# Interview notes (fictional) used to fill the JSON

Example 1, `resume-one-column.json`: Jordan Ellery, senior data analyst, Manchester. 7 years, three roles. Wants a 1-page ATS resume to apply for senior analytics roles. Metrics (6 h to 45 min, +3.1% conversion, -2.4 pt churn, etc.) were given by the user in the interview; each bullet follows action verb + what + measured result.

Example 2, `resume-two-column.json`: Sam Whitlock, product designer, Austin. 8 years, four roles, up to 2 pages, modern two-column layout for direct sends to hiring managers (not through an ATS form).

All companies, schools and people are fictional. Build:
`python ../scripts/build_resume.py resume-one-column.json . --name output-one-column`
`python ../scripts/build_resume.py resume-two-column.json . --name output-two-column --job job.txt` (job.txt optional)
