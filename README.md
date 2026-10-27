# MSP Health Dashboard

MSP & CCB Program Performance dashboard. A single self-contained HTML file with the
weekly data embedded. It is generated from `build/template.html` plus the weekly Excel
export, and distributed to internal users through Box (not through this repo).

This repo is for version control of the code only. The Excel exports and the generated
HTML are git-ignored.

## Weekly refresh

1. Create a `Source files/` folder in the repo root (it is git-ignored).
2. Put the weekly export in it, with the date in the filename, for example
   `MSP Reporting 2026-10-27.xlsx`. Also add `Glossary of Terms.xlsx` (exact name).
3. Install dependencies once: `npm install`
4. Build: `npm run build` (takes about 1 minute). Output: `Final/MSP-CCB-Performance-Dashboard.html`
5. Upload the generated HTML to the Box folder.
6. Optional: `npm run review` writes `data-review.html`, a filterable table of all records.

## Making changes

- Edit `build/template.html` (the source of truth), then rebuild. Never edit the file in `Final/`.
- Data parsing and computed fields live in `build/parse-data.js`.
- See `AGENTS.md` for the business rules (utilization, status bands, committed amount) and coding conventions.

## Working together

- Work on a branch, open a pull request, and have the other person review before merging to `main`.
- `template.html` is one large file, so keep changes focused and pull before you start to avoid merge conflicts.
