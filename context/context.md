# MSP Health Dashboard: session starter prompt

Copy everything below the line into a new session as the first message. Then add the task after it.

When something in the project changes in a way a new session needs to know (new rule, new decision, new open item), update this file.

---

## Master prompt

You are helping me maintain and extend the **MSP & CCB Performance Dashboard** (also called the MSP Health Dashboard). Before doing anything else, get oriented, then wait for my task.

### 1. What this project is

- A single self-contained HTML dashboard with the weekly Excel data embedded in it. No server, no database.
- Users are sales people tracking MSP and CCB contract progress. Use plain, user-friendly language in anything they see (labels, pop-ups, tooltips).
- It is built from `build/template.html` plus a weekly Excel export, and distributed manually by uploading the output file to a Box folder.
- I work on a Mac. A colleague works on Windows. Anything that touches shared files or scripts must work on both.
- Project root: `MSP Dashboard/` (inside the workspace `MSP_Health_Dashboard/`).

### 2. Read these first (in this order)

1. `MSP Dashboard/AGENTS.md`: coding conventions and non-obvious rules.
2. `MSP Dashboard/README.md`: the weekly routine and the pipeline.
3. `MSP Dashboard/STATUS-UTILIZATION-LOGIC.md`: how Status and Utilization % are calculated, plus open business questions.
4. `MSP Dashboard/build/update.js`, `build/parse-data.js`: the pipeline and the data rules.

Do not guess about code you haven't opened. Check the files, then tell me in a few lines what you understood before you start.

### 3. How the pipeline works

- `build/parse-data.js` reads the weekly Excel in `Source files/`, drops inactive-flag rows, repairs encoding problems and computes fields such as tenure, contract type (Monthly or Annual), commit amount and the last-12-months flag. It then embeds the result as `window.__MSP_DATA__`. The dashboard reads only the newest snapshot.
- It also loads the newest **Active Management Orgs view CSV** from `Source files/Active Orgs/`. On the latest snapshot it keeps an active contract only if it appears in that view, and it prints a reconciliation report and warnings in the build log.
- `build/update.js` is the full weekly run: archive older exports and view CSVs into `History/`, back up the previous output, build, rebuild `data-review.html`, then open the result.
- Output: `Final/MSP-CCB-Performance-Dashboard.html` (about 37 MB). Backups are in `Final/Backup/`.
- The Glossary tab is driven by `Source files/Glossary of Terms.xlsx`, which is shared with my colleague.

### 4. Weekly routine (what I do when new data arrives)

1. Drop the newest `MSP Reporting <Mon> <D> <YYYY>.xlsx` into `Source files/`.
2. Drop the matching view CSV (`Active_Management_Orgs_<dd_mm_yy>.csv`) into `Source files/Active Orgs/`.
3. Run `Update-Dashboard.command` (Mac), `Update-Dashboard.bat` (Windows) or `npm run update`.
4. Check the "Data Refreshed" date and the build log warnings, then upload `Final/MSP-CCB-Performance-Dashboard.html` to Box.

Tenure, last-12-months and expired logic depend on the **build date**, so the build must be run on publish day. Who publishes weekly is not decided yet.

### 5. Rules for how you work

- **Edit `build/template.html` (the source of truth), then rebuild.** Never hand-edit `Final/MSP-CCB-Performance-Dashboard.html`.
- Rebuild with `cd "MSP Dashboard" && node build/update.js --no-open`. It takes about 1 minute, so allow up to 270 seconds.
- After each change, verify it before saying it is done:
  - Syntax-check the built file's inline `<script>` blocks with `new Function`.
  - Run the real logic against the real data when logic changed.
  - For visible changes, test in headless Chrome (`/Applications/Google Chrome.app`) with `--remote-debugging-port=9333 --user-data-dir=/tmp/chrome-prof`, driven through Node's built-in `WebSocket` and CDP. Stop it afterwards with `pkill -f remote-debugging-port=9333`.
  - Screenshots saved outside the workspace can't be read, so report measured numbers instead, or save them inside the workspace.
  - Tell me clearly what you verified and what you did not.
- Follow `AGENTS.md`: vanilla JS in an IIFE, CSS variables, prefixes `po-` (Program Overview), `iov-` (Individual Org View), `dd-` (Data Detail), `lm-` (Status and Utilization % pop-up). Do not add a third copy of `fmtV` or `fmtM`. Some banner comments in the template are corrupted; match on neighbouring lines when editing.
- Make the smallest change that solves the request. Don't refactor or add features I didn't ask for.
- Ask before making a decision I haven't made. For analysis-only requests, give the analysis and wait for my go-ahead before changing anything.
- Be direct and concise. Use plain language, not jargon.

### 6. Git and GitHub: do not touch for now

- This folder is **not** a git repository. Do not run `git init`, commit or push, and do not reconnect to any remote, unless I ask.
- An older private copy exists at `https://github.com/cloudability/msp-health-dashboard`. When I do ask for version control later:
  - Use git identity `Elvin Tahmaz <etahmaz-apptio@users.noreply.github.com>` as a local config.
  - The macOS keychain credential works (`gh` is not logged in).
  - `.gitignore` should cover `Final/`, `Source files/`, `*.xlsx`, `node_modules/`, `data-review.html`, `build/Rollback Template Version/`, `.DS_Store`, `.bob/tmp/`, `.bob/artifacts/`.
  - Add `.gitattributes` (`* text=auto eol=lf`) and `.editorconfig`.
  - The editor tool blocks writing `.gitignore`, so create it with a shell heredoc.

### 7. Current state of the dashboard (check against the code, it may have moved on)

- **Contracts filter:** table-only. It does not affect the KPIs or charts. The default is Active. Expired is not in the Status dropdown.
- **Status and Utilization %:** the rules are in `STATUS-UTILIZATION-LOGIC.md`. The (i) icon on the Status, Utilization % and Commit Amount headers opens the `lm-` pop-up in both the Program Overview and Individual Org View tables.
  - The pop-up text is hand-written HTML (`#lm-overlay`). It duplicates `buildMgmtOrgs()` and the logic doc, so update all three together.
  - The layout is locked, so don't restructure it without asking.
- **Commit amount rule:** Monthly = $1,000 per entitlement. Annual and all others = $100,000 per entitlement.
- **Active-orgs view filter:** done. Watch the build log for date mismatches, active orgs missing from the view, and unmatched view rows.

### 8. Open items (for awareness, don't start them unless I ask)

- Add "Status" and a rewritten "Utilization (%)" entry to `Source files/Glossary of Terms.xlsx`. This file is shared, so confirm with me before editing it.
- Early Stage and pacing thresholds are "pending business confirmation".
- Confirm which TIVIT contract is current (the Annual one, $34M from 2026-09-02, was dropped by the view filter, which lists the Monthly one).
- Softcat has an expired line with a $1,000 commit (qty 0.01) showing about 39,899%.
- The Program Overview table has a table-only "Reporting Month" dropdown (last 12 months, default current month to date). An earlier month changes CSUM and End Consumers but not tenure or Active/Expired status, so a note under the table says so.
- Decide who publishes weekly.
- Optional cleanup: unused `build/prune-backups.ps1`, old large files in `Final/Backup/`.
- Reconnect git and GitHub later (see section 6).

### 9. Start of session

Read the files in section 2, then reply with:

1. A 5-line summary of how the project works and its current state.
2. Anything in the code that disagrees with this prompt.
3. "Ready. What's the task?"

Do not change any files until I give you the task.

---

## Notes for me (not part of the prompt)

- Paste the section from "You are helping me…" to the end of "Do not change any files until I give you the task."
- If you give the session data to load, say which export and view CSV to use and whether to rebuild.
