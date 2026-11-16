# MSP & CCB Performance Dashboard

A single self-contained HTML file with the weekly data embedded. It is generated from
`build/template.html` plus the weekly Excel export, and shared with users through Box.

## Folder structure

```
Source files/            <- newest weekly .xlsx + "Glossary of Terms.xlsx"
  History/               <- older exports (moved here automatically)
  Active Orgs/           <- newest V_SSABL_MSP_ACTIVE view export (.csv), same date as the .xlsx
    History/             <- older view exports (moved here automatically)
build/
  template.html          <- dashboard UI and logic (source of truth - edit this)
  parse-data.js          <- Excel -> data + computed fields, injects into the template
  make-data-review.js    <- builds data-review.html
  update.js              <- one-step weekly pipeline (macOS and Windows)
Final/
  MSP-CCB-Performance-Dashboard.html   <- generated output (upload this to Box)
  Backup/                <- timestamped copies of previous outputs
Update-Dashboard.command <- macOS: double-click to refresh
Update-Dashboard.bat     <- Windows: double-click to refresh
```

## Setup (first time, each machine)

Install Node.js (https://nodejs.org), then in this folder:

```bash
npm install
```

macOS only, once: `chmod +x Update-Dashboard.command`

## Weekly refresh

1. Drop the new export into `Source files/`, with the date in the name
   (`MSP Reporting Oct 30 2026.xlsx` or `MSP Reporting 2026-10-30.xlsx`).
   Also drop the matching active-orgs view export into `Source files/Active Orgs/`
   (`Active_Management_Orgs_30_10_26.csv`, date as dd_mm_yy).
2. Double-click `Update-Dashboard.command` (Mac) or `Update-Dashboard.bat` (Windows),
   or run `npm run update`.
3. Check "Data Refreshed" in the dashboard header shows the new date.
4. Upload `Final/MSP-CCB-Performance-Dashboard.html` to Box.

`update.js` keeps only the newest export in `Source files/` (the dashboard reads a single
snapshot), moves older ones to `History/`, backs up the previous dashboard, rebuilds the
dashboard and `data-review.html`, and opens the result. Use `npm run update -- --no-open` to skip opening.

### Active-orgs view filter

The Excel export can hold more than one contract per org that is active by date (overlaps),
which inflates CSUM. The view CSV lists one current contract per org. During the build, an
active-by-date row is kept only if its `DOMESTIC_CLIENT_ID + CONTRACT_START_DATE + CONTRACT_END_DATE`
is in the view; expired/historical rows are never touched. The build prints a per-org
reconciliation (rows dropped, CSUM before/after) and warns if the view date differs from the
export date, if an active org is missing from the view, or if a view contract has no matching rows.
With no CSV present the build warns and does not filter.

## Making changes

- Edit `build/template.html`, then run `npm run update`. Never edit the file in `Final/` by hand.
- Business rules (utilization, status bands, committed amount) and coding conventions: see `AGENTS.md`.
- Tenure, status and L12 flags are computed against the build date, so rebuild on publish day.
- The current (month-to-date) month is the newest `YEAR_MONTH` in the export when it is later than the latest completed month. It is the 12th month of the last-12-months charts and the default reporting month.
