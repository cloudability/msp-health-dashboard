# AGENTS.md

This file provides guidance to agents when working with code in this repository.

## Non-obvious architectural constraints

- **Data flows one way: Excel → `build/parse-data.js` → `__MSP_DATA__` JSON → `build/template.html` → Final HTML.** There is no database or server, and nothing persists. All filtering and aggregation happens in the browser against the embedded JSON.

- **`weeklySnapshots` is an array but only index `[0]` is ever used.** The array exists for a planned Snowflake migration. Don't assume multi-snapshot logic works.

- **Utilization % and Status both branch on `contractType`**, not on `type` (MSP/CCB).
  - Monthly: `csumCurrent / committedAmount`, with flat status bands.
  - Annual: `csumLifetime / committedAmount`, with a pacing-ratio status (actual % over the linear expected % for the months elapsed). Early Stage is evaluated before the pacing bands (thresholds are constants above `buildMgmtOrgs`).
  - `tenureMonths` counts completed months to the build date, while CSUM includes the current partial month. Annual pacing therefore runs slightly optimistic early in a month. This is accepted.

- **`buildMgmtOrgs(rows, targetYM)` has two call signatures.** The management table passes `poTableFilter.yearMonth`, while the IOV table passes nothing (all rows). The wrong signature gives wrong utilization numbers.

- **Business rules live in the build, not the browser.** `committedAmount` multipliers (no per-org special cases) and the encoding repairs (`MOJIBAKE_FIXES`) are in `build/parse-data.js`. Fix data problems there.

- **The PDF export is size-sensitive.** Large DXC-type orgs have 100+ end consumers. The PDF caps the end-consumer chart to one page, and images are JPEG at quality 0.6. Any new tall or dense chart needs a similar cap.

- **Chart.js is the only runtime dependency for the dashboard itself.** `html2canvas` and `jsPDF` are CDN scripts used only for PDF export.

- **Snowflake migration path:** replace `/*{{DATA_INJECTION}}*/` with a `fetch()`, and move `computeFields()` into SQL. The rendering layer needs no change.

- **`Source files/` holds customer data.** There is no `.gitignore` in the repo. A build with no Excel files yields an empty dashboard (`snapshotDates: []`) and doesn't error.
