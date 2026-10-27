# AGENTS.md

This file provides guidance to agents when working with code in this repository.

## Commands

```bash
node build/parse-data.js        # = npm run build: Excel → Final/MSP-CCB-Performance-Dashboard.html (~1 min)
node build/make-data-review.js  # = npm run review: needs the Final HTML from the step above
```

- On this Windows box PowerShell blocks `npm.ps1`; run the `node` commands directly.
- There is no test or lint setup. Validate by rebuilding, then syntax-checking the inline `<script>` blocks of the Final HTML with `new Function(src)`.

## Two-file rule

`build/template.html` is the source of truth; `Final/MSP-CCB-Performance-Dashboard.html` is generated from it by a single replace of `/*{{DATA_INJECTION}}*/`.
- Edit the template, then rebuild. This keeps both files in sync and needs `Source files/*.xlsx`.
- Without the Excel files the rebuild produces an empty dashboard, so in that case edit both files by hand.
- Don't try to `apply_diff` the `// ����` banner comments in the template: their characters are already corrupted. Match on neighbouring lines instead.

## Data and build

- `window.__MSP_DATA__ = { generatedAt, maxWeeks, snapshotDates[], dimensions{}, glossary[], weeklySnapshots[] }`. Only `weeklySnapshots[0]` is ever read.
- Inactive rows (`INACT_FLAG === 1`) are dropped in `parseRow`. The snapshot date comes from the filename (ISO, then "Month Day Year", then mtime), so renaming a file changes its slot.
- `MAX_WEEKS = 52`. The glossary is read from `Source files/Glossary of Terms.xlsx`, exact name.
- The source Excel contains mojibake (`TELEFï¿½NICA`). `cellVal()` in `build/parse-data.js` repairs it through `MOJIBAKE_FIXES`. A new corrupted name prints a build warning, so add a rule to that list.
- `name` (column `NAME`) is the org display name. `managementOrgName` is a different column.

## Computed fields (`computeFields()` in `build/parse-data.js`, build-time only)

`isLatestMonth`, `contractType` (PART_DSCR_FULL contains MONTH or ANNUM), `configuredCSum`, `committedAmount`, `isL12Months`, `tenureMonths` (DATEDIF in months, completed months only), `endConsumerName2`.
- `committedAmount` is `PART_QTY × 100000` for CCB and Annual, and `× 1000` for Monthly and for names starting with `DXC Technology`.
- A wrong multiplier shows up as a value 100× too large (a DXC annual contract showed $31.0B instead of $310M).

## Status and Utilization (`buildMgmtOrgs()` in `build/template.html`)

- Utilization % is `csumCurrent / committedAmount` for Monthly and `csumLifetime / committedAmount` for Annual.
- Monthly status uses flat bands: over 100 Over, 70 Healthy, 50 Moderate, 30 Low, above 0 At Risk.
- Annual status uses `pacingRatio = utilPct / (min(tenure, 12) / 12 × 100)`. Checks run in order and stop at the first match:
  1. No CSUM
  2. Early Stage: tenure ≤ `EARLY_STAGE_MAX_MONTHS` (2) and ratio < `EARLY_STAGE_BELOW_RATIO` (0.85)
  3. Over Utilization: ratio above 1.15
  4. Healthy: ratio 0.85 or higher
  5. Moderate: ratio 0.60 or higher
  6. Low Utilization: ratio 0.35 or higher
  7. At Risk
- Early Stage is checked before the pacing bands. Telefónica Tech UK (tenure 1, ratio 0.42) is therefore Early Stage, not Low Utilization. This was confirmed with the business.
- The utilization bar and cell tint follow the status (`UTIL_STATUS_COLORS`, `util-cell-*`), not the raw %. A new status needs a badge class, a `UTIL_STATUS_COLORS` entry, a filter `<option>`, and a `statusBadge()` branch.

## Conventions

- The whole script is one `'use strict'` IIFE of vanilla ES2020, with no bundler. Chart.js, html2canvas and jsPDF are loaded from a CDN.
- Chart instances live in `CHARTS[id]`. Program Overview charts use `makeChart(id, cfg)`. Individual Org View charts use `destroyChart(id)` and then `CHARTS[id] = new Chart(...)`.
- `iovInit()` runs once (guarded by `iovInitialized`), so wire its event listeners there and not in `iovRender()`.
- `buildMgmtOrgs()` deduplicates CSUM through the `csumCurrentSeen` and `csumLifetimeSeen` Maps. New metrics must follow the same pattern, or totals are inflated. It is called with `targetYM` from the Program Overview table and without it from the Individual Org table.
- Use CSS variables from `:root`, with no hardcoded colours except the status palette. `fmtV()` and `fmtM()` are identical copies, so don't add a third.
- Prefixes: `po-` is Program Overview, `iov-` is Individual Org View, `dd-` is Data Detail.

## Individual Org View charts and PDF

- The global rule `.chart-box canvas { max-height: 260px }` caps every canvas. `#iov-chart-csum-ec` overrides it, and its size comes from the wrapper `#iov-inner-csum-ec` via `iovApplyEcLayout()`. The constants are `EC_ROW_PX` (28), `EC_VISIBLE_ROWS` (10) and `EC_PDF_ROW_PX` (20).
- More than 10 end consumers scroll on screen. The PDF export expands the chart and caps it to one page, adds a "Showing top N of M" note, and restores the chart afterwards (`addEcChart` in `iovExportPDF`).
- PDF images are JPEG at `PDF_JPEG_QUALITY = 0.6`. PNG and one very tall capture produced a 68 MB file for DXC.
