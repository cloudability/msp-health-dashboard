# AGENTS.md

This file provides guidance to agents when working with code in this repository.

## Non-obvious coding rules

- **Edit `build/template.html`, then rebuild** (`node build/parse-data.js`) so `Final/MSP-CCB-Performance-Dashboard.html` stays in sync. Edit both files by hand only if `Source files/*.xlsx` is missing. See the root `AGENTS.md`.

- **No bundler, no transpilation.** All JS is vanilla ES2020+ inside a single `(function(){ 'use strict'; … })()` IIFE. Use `const`/`let`, arrow functions, optional chaining (`?.`), and `async/await` freely.

- **Chart instances must go through `CHARTS[id]`.** `makeChart(id, cfg)` destroys any existing chart before creating a new one, so use it for program-overview charts. For IOV charts call `destroyChart(id)` and then `CHARTS[id] = new Chart(...)`.

- **`iovInit()` is guarded by `iovInitialized`.** It runs once. Wire event listeners (including the PDF button) there, not in `iovRender()`, which re-runs on every org or slider change.

- **`buildMgmtOrgs()` deduplicates CSUM** via `csumCurrentSeen` / `csumLifetimeSeen` Maps keyed on `endConsumerName|vendor|configuredCSum`. New management-table metrics must follow the same pattern or totals will be inflated.

- **`poTableFilter.yearMonth` defaults to the latest YM** (set in `populatePoFilterDropdowns`). Filter state lives in `poTableFilter`. Mutate it and call `renderMgmtTable(rows)`.

- **CSS variables only** for colours (`:root`: `--blu`, `--grn`, `--red`, `--yel`, `--pur`, `--teal`, `--mut`, `--sur`, `--sur2`, `--bdr`, `--txt`, `--bg`). The one exception is the status palette (`UTIL_STATUS_COLORS`, `.status-*`).

- **`.chart-box canvas` has `max-height: 260px`.** A chart that must be taller needs an override rule (see `#iov-chart-csum-ec`) and a sized wrapper div.

- **PDF export captures with `html2canvas` and `addElement()`**, which moves a box to the next page if it doesn't fit. Keep images as JPEG (`PDF_JPEG_QUALITY`) and never capture a very tall element, or the file grows to tens of MB.

- **`fmtV()` and `fmtM()` are identical** (global vs local to `iovRender()`). Don't add a third copy.
