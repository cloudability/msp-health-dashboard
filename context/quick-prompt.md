You are helping me maintain and extend the **MSP & CCB Performance Dashboard** (MSP Dashboard). Before doing anything, get oriented with the codebase and our latest progress, then wait for my instructions.

### 1. Project Overview & Rules
- Single self-contained HTML dashboard with embedded weekly snapshot data (`Final/MSP-CCB-Performance-Dashboard.html`).
- Source of truth is `MSP Dashboard/build/template.html`. Never hand-edit the `Final/` file.
- Rebuild command: `cd "MSP Dashboard" && node build/update.js --no-open` (takes ~1 minute).
- Pipeline scripts: `build/parse-data.js`, `build/update.js`, `build/make-data-review.html`.
- Validation rules:
  1. Syntax-check inline `<script>` tags in the built HTML using `new Function(code)`.
  2. For visual/UI changes, test in headless Chrome using CDP (`--remote-debugging-port=9333`) and stop it afterwards (`pkill -f remote-debugging-port=9333`).
- Git: This workspace is not a git repository. Do not run `git init`, commit, or push.

### 2. Recent Updates Completed in the Previous Session
1. **Nov 2, 2026 Snapshot & CSV Support:**
   - Ingested `Source files/MSP Reporting Nov 2 2026.csv`.
   - Updated `build/parse-data.js` and `build/update.js` to support both `.csv` and `.xlsx` files and map the `CONTRACT_STATUS` (`ACTIVE` / `EXPIRED`) column directly from the export.
2. **DXC Override Removal:**
   - Removed hardcoded overrides for DXC Technology Services - FinOps in `build/parse-data.js`, `AGENTS.md`, and `STATUS-UTILIZATION-LOGIC.md`. DXC now derives `Monthly` contract type and `$25,834,000` commit amount directly from the data.
3. **IOV Timeframe Slider & Chart Layouts:**
   - Removed the hardcoded `isL12Months` restriction from `l12Rows` in `build/template.html` so the timeframe slider fully controls the chart data window.
   - Full-width chart (**CSUM Trending Assigned vs Unassigned**) uses `autoSkip: false` to display all month labels.
   - The three 2-column smaller charts (**End Consumers Trending**, **CSUM Trending by Commit**, **CSUM Trending by Vendor**) use `autoSkip: true` on x-axis ticks to prevent label collision.
4. **Dynamic Org-Specific Timeframe Selection:**
   - Implemented `computeOrgMonths(orgName)` and `iovResetSliderForOrg(orgName)` in `build/template.html`.
   - When switching orgs (combobox or table row click), the timeframe slider track automatically bounds itself between the org's earliest contract start date and latest contract end date (capped at the current snapshot date, e.g. Nov 2026).
   - The track start/end labels and active range label update dynamically per org (e.g., AXA defaults to `Nov 2025 - Nov 2026`, Deloitte to `Jun 2025 - Nov 2026`, SVA to `Aug 2025 - Jul 2026`).

Please read `MSP Dashboard/AGENTS.md`, `MSP Dashboard/STATUS-UTILIZATION-LOGIC.md`, and `MSP Dashboard/build/template.html`, confirm you are oriented, and let me know when you are ready for the next task.