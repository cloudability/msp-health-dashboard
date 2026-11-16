# AGENTS.md

This file provides guidance to agents when working with code in this repository.

## Commands

```bash
node build/update.js            # = npm run update: full weekly pipeline (archive old Excel, backup, build, review, open)
node build/parse-data.js        # = npm run build: Excel → Final/MSP-CCB-Performance-Dashboard.html (~1 min)
node build/make-data-review.js  # = npm run review: needs the Final HTML from the step above
```

- Works on macOS and Windows. Double-click `Update-Dashboard.command` (Mac) or `Update-Dashboard.bat` (Windows). If PowerShell blocks `npm.ps1` on Windows, run the `node` commands directly.
- The dashboard reads only `weeklySnapshots[0]`, so `Source files/` must hold one export; `update.js` moves older ones to `Source files/History/`.
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
- **Active-orgs view filter** (`applyActiveViewFilter` in `build/parse-data.js`): the newest CSV in `Source files/Active Orgs/` (V_SSABL_MSP_ACTIVE export, date `dd_mm_yy` in the name) lists one current contract per org. On the latest snapshot, a row that is active by date (`contractStart <= snapshot <= contractEnd`) is kept only if `domesticClientId|contractStart|contractEnd` is in the view; expired rows are untouched. This removes overlapping contracts that double-counted CSUM (5 orgs as of 2026-10-26). The build prints a per-org reconciliation and warns on date mismatch, active orgs missing from the view, and unmatched view rows. No CSV = warning, no filtering. `update.js` archives older view CSVs to `Active Orgs/History/`. Don't key on `sfdcId` (blank on ~25% of rows).
- The source Excel contains mojibake (`TELEFï¿½NICA`). `cellVal()` in `build/parse-data.js` repairs it through `MOJIBAKE_FIXES`. A new corrupted name prints a build warning, so add a rule to that list.
- `name` (column `NAME`) is the org display name. `managementOrgName` is a different column.
- **Sentiment** (`SENTIMENT` column, `sentiment` field): customer sentiment per org, Green / Yellow / Red, blank for orgs that are not tracked (27 of 65 orgs in the 2026-11-02 export; the value is the same on every row of an org). It is the "Sentiment" column of the Management Org Details table (between PSM and Status), shared by the Program Overview and Individual Org tables through `MGMT_COLS`. `buildMgmtOrgs()` takes it from any row of the contract line, and blank cells stay blank (no "—"). TAM and PSM follow the same rule (blank, not "—") and render inside a `.cell-scroll` div (75px wide, scrolls sideways for longer names, full name in the hover title) to keep the table compact. The Name column uses the same `.cell-scroll` pattern with `.name-scroll` (200px wide), so long org names scroll sideways instead of being cut off with "..."./

## Computed fields (`computeFields()` in `build/parse-data.js`, build-time only)

`isLatestMonth` (matches the latest completed/closed calendar month, e.g. Oct 2026 for a Nov 2 snapshot with in-flight Nov data), `isCurrentMonth` (the in-flight month-to-date month, e.g. Nov 2026; falls back to the latest completed month if the file has no in-flight month), `contractType` (PART_DSCR_FULL contains MONTH or ANNUM), `configuredCSum`, `committedAmount`, `isL12Months` (rolling 12 months ending with the current month, so the current month-to-date month is the 12th, e.g. Dec 2025 – Nov 2026), `tenureMonths` (DATEDIF in months, completed months only), `endConsumerName2`.
- `committedAmount` is `PART_QTY × 1000` for Monthly (including CCB orgs) and `× 100000` for Annual and for CCB on any other contract type.

## Status and Utilization (`buildMgmtOrgs()` in `build/template.html`)

- Utilization % is `csumCurrent / committedAmount` for Monthly and `csumLifetime / committedAmount` for Annual. For Expired rows (`contractEnd < snapshotDate`) it is whole-contract: Annual `csumLifetime / committedAmount`, Monthly `csumLifetime / (committedAmount × tenureMonths)`. Status stays Expired and the bar stays grey.
- Monthly status uses flat bands: over 100 Over, 70 Healthy, 50 Moderate, 30 Low, below 30 (including 0) At Risk. "Expired" is assigned only when `contractEnd < snapshotDate`, never as a fallback for 0%.
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
- The Program Overview and Individual Org tables default to the current month (`isCurrentMonth`), so "CSUM (Current Month)" and "Reporting Month" are month-to-date. KPIs: "Total CSUM (Last Month)" uses `isLatestMonth`, "Total CSUM (Current Month, MTD)" uses `isCurrentMonth`, "Total CSUM (Last 12 Months)" uses `isL12Months`; all three CSUM KPIs sum the `dedupeCsumLines()` output, so the Last 12 Months card equals the sum of the chart bars. "Total End Consumers (Active Contracts)" stays on `isLatestMonth` (last completed month) but counts only rows on ACTIVE contracts (`isActiveRow`: `contractStatus === 'ACTIVE'`, else date test), so orgs whose contracts have all expired drop out. The card shows the combined count with an MSP / CCB split underneath (the Management Org Details table is month-to-date, so its End Consumers column sums to a different, smaller number early in the month). KPI cards with an (i) tooltip (`tip` in the `cards` array of `renderKPIs()`): the four org/EC cards say "Last completed month: <Month Year>" (Management Orgs with Active End Consumers %, Total End Consumers) or "Orgs with an active contract as of the <snapshot date> snapshot." (Active Management / MSP / CCB Orgs, which are contract-based and ignore the month) (2026-11-02 export: 582 = MSP 467 + CCB 115; it was 597 before the active filter).
- **Program Overview charts** build on `dedupeCsumLines()` (one usage line per org + month + end consumer + vendor + amount, kept for the contract with the later start) and `commitByMonth()` (one commit per contract, counted in every month the contract is active by its dates; an org runs one contract at a time, so the later-start contract wins when two touch the same month). Do not dedupe commits by `salesOrderNum`: most rows have none, which previously dropped most contracts and produced wrong commit lines. The Annual CSUM line is cumulative over the 12 months and includes spend from contracts that have since expired.
- **Chart legends** are the same on every tab: a custom legend in the chart header, top right, next to the title (Chart.js legend is off). Bar charts use 10x10 rounded square swatches, line charts use 24x3 line swatches, 12px muted text, 16px gap. Copy the header markup of "CSUM Trending by Commit" (Individual Org View) for any new chart.
- **Configured CSUM is always green** (`#42be65`) in every chart and legend on every tab (Program Overview Monthly/Annual CSUM Trending, Individual Org View). Commit series are orange (`#ff832b`). Blue is reserved for the MSP series in "Configured CSUM — MSP vs CCB by Month" and the Early Stage status.
- **Program Overview filter bar** (`#snap-bar`): the reporting month is a "Reporting Month" dropdown (`#po-f-yearmonth`) in the "Table only" group, listing the last 12 months (`isL12Months`, newest first, no "All"), defaulting to the current month to date. It only drives `poTableFilter.yearMonth` for the table; KPIs and charts ignore it. When a past month is picked, a note under the table title (`#po-past-month-note`) says Status, Utilization % and Tenure still reflect today's contract state (only CSUM and End Consumers are cut at the selected month). Order: Org Type, Region, Contract Type, Vendor, Name (filter KPIs, charts and table), then Reporting Month, Contracts and Status in the "Table only" group. Name also filters the CSUM and End Consumer KPIs. The table keeps live-contract orgs that have no row yet in the current month, so it matches the Active Orgs KPI. The note on the right uses the same info-icon style as `#po-l12-note`.
- The Program Overview note above the first chart (`#po-l12-note`, filled by `updateL12Note()`) says all four charts show the last 12 months including the current month to date. Its date range is read from the `isL12Months` flag at render time. If a Program Overview chart stops using `isL12Months`, update or remove the note.
- Prefixes: `po-` is Program Overview, `iov-` is Individual Org View, `dd-` is Data Detail, `lm-` is the Status / Utilization % logic pop-up (opened by the (i) icon in the table headers). The rule text in the pop-up is hand-written HTML (`#lm-overlay`) and duplicates `buildMgmtOrgs()` and `STATUS-UTILIZATION-LOGIC.md`: update all three together, and the "Rules as of" date in the footer.

## Individual Org View charts and PDF

- The global rule `.chart-box canvas { max-height: 260px }` caps every canvas. `#iov-chart-csum-ec` overrides it, and its size comes from the wrapper `#iov-inner-csum-ec` via `iovApplyEcLayout()`. The constants are `EC_ROW_PX` (28), `EC_VISIBLE_ROWS` (10) and `EC_PDF_ROW_PX` (20).
- **DXC Regional Selector (`#iov-dxc-region-bar`):** Global region button group (`DXC Global`, `DXC US`, `DXC EU`, `DXC AU`) displayed below the Management Org Details table and above the charts when viewing DXC Technology Services - FinOps. Dynamically filters all 5 charts for the selected environment (US: `dxc-msp.com`, EU: `dxc-msp-eu.com`, AU: `dxc-msp-au.com`) within the active timeframe slider window.
- More than 10 end consumers scroll on screen. The PDF export expands the chart and caps it to one page, adds a "Showing top N of M" note, and restores the chart afterwards (`addEcChart` in `iovExportPDF`).
- PDF images are JPEG at `PDF_JPEG_QUALITY = 0.6`. PNG and one very tall capture produced a 68 MB file for DXC.
