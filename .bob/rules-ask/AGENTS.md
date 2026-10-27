# AGENTS.md

This file provides guidance to agents when working with code in this repository.

## Non-obvious documentation context

- **`README.md` is partially outdated** — it references `dashboard.html` as the output but the actual output is `Final/MSP-CCB-Performance-Dashboard.html`. The column mapping section is a leftover from an earlier design; the real mapping is the `COL` object in `build/parse-data.js`.

- **Computed field logic is documented as DAX/Excel equivalents** in `build/parse-data.js` comments directly above each field. When explaining what a metric means, the canonical reference is those comments, not external documentation.

- **`data-review.html`** is a separate diagnostic flat-table view built from the same `__MSP_DATA__` blob. It is not the main dashboard — it is used for data validation only.

- **The glossary** is driven by `Source files/Glossary of Terms.xlsx` (columns: Category, Term, Definition). It is separate from the metrics data files and parsed independently.

- **IOV = "Individual Org View"** — this is the second tab. All its identifiers are prefixed `iov-`.

- **PO = "Program Overview"** — the first tab. All its identifiers are prefixed `po-`.

- **DD = "Data Detail"** — the third tab. All its identifiers are prefixed `dd-`.
