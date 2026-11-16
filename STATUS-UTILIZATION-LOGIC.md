# Status & Utilization % logic

Reference for how the **Utilization %** and **Status** columns are calculated in the
Management Org Details table (Program Overview) and in the Individual Org View table, which shares the same code.

- **Source of truth:** `buildMgmtOrgs()` in `build/template.html` (the rules), and `computeFields()` in `build/parse-data.js` (tenure, contract type, commit amount, CSUM per row).
- **Replaces:** `status-utilization-logic-redesign.html`, which is the original design and is out of date (it predates Expired contracts and the removal of the "No Data" status).
- **Examples:** figures in this document come from the 2026-10-26 export (build date 2026-10-29). They will change with new data.
- **Keep in sync:** when the code changes, change this file, then the pop-up and Glossary wording that are based on it.

---

## 1. What one table row is

One row = one **contract line**: rows are grouped by *org name + contract end date*.
An org with an active contract and past expired contracts has several rows.

For each row the dashboard works out a CSUM figure, a commit amount, a tenure, a **Utilization %** and a **Status**.

| Term | Meaning |
|---|---|
| **Reporting month** | Always the current (month-to-date) month, e.g. 2026-11 for a Nov 2 snapshot. It is shown as a read-only label in the filter bar (there is no Year/Month dropdown). CSUM (current month) is therefore month-to-date. |
| **Snapshot date** | Date of the Excel export (e.g. 2026-10-26). Used to decide whether a contract is expired. |
| **Build date** | The day the dashboard was rebuilt. Used to calculate tenure. |
| **CSUM (current month)** | Configured CSUM of the reporting month only. |
| **CSUM (lifetime)** | Configured CSUM of the contract from its start up to and including the reporting month. |
| **Commit amount** | Contracted spend (see section 6). |
| **Tenure (months)** | Whole months from contract start to the build date (see section 7). |

---

## 2. Which formula applies (decision order)

Evaluated top to bottom; the first match wins.

1. **Expired** — `contractEnd < snapshot date`. Utilization is whole-contract (section 3.3). Status is always **Expired**.
2. **Annual** (live) — Utilization uses lifetime CSUM. Status uses pacing (section 4.1).
3. **Monthly** (live) — Utilization uses lifetime CSUM divided by (monthly commit × tenure). Status uses fixed bands (section 4.2).

> **Note on DXC Technology Services - FinOps:** it has no special handling. Its contract type (Monthly) and commit amount ($25,834,000 = 25,834 entitlements × $1,000) are derived from the data like any other org.

---

## 3. Utilization %

Utilization % is always shown as the raw percentage. Only the *Status* is pacing-adjusted.

### 3.1 Annual (live or expired)

```
Utilization % = CSUM (lifetime) / Commit amount × 100
```

Covers the contract so far, cumulative up to the reporting month.

### 3.2 Monthly (live or expired)

```
Utilization % = CSUM (lifetime) / (Monthly commit × Tenure months) × 100
```

Covers the whole contract life to date (cumulative up to the reporting month). Tenure for an expired contract is the months the contract covered (start to end date); for a live contract it is months from start to the build date.

A month with no new spend lowers the utilization percentage incrementally but does not reset it to 0%.

### 3.3 When it shows "—"

If the commit amount is 0, or tenure is 0 for a Monthly contract, the percentage is blank. No live contract is in this state today.

---

## 4. Status

### 4.1 Annual contracts (live): pacing-adjusted

A new annual contract is not expected to have spent 100% in month 1, so Status compares actual
spend against where the contract *should* be by now (linear burn over 12 months).

```
elapsedMonths = min(tenureMonths, 12)        // 0 counts as 1
expectedPct   = elapsedMonths / 12 × 100
actualPct     = CSUM (lifetime) / Commit amount × 100      // = the Utilization % column
pacingRatio   = actualPct / expectedPct                    // > 1 ahead of pace, < 1 behind
```

Rules, evaluated in this order:

| # | Condition | Status |
|---|---|---|
| 1 | CSUM (lifetime) = 0 | **No CSUM** |
| 2 | elapsedMonths ≤ 2 **and** pacingRatio < 0.85 | **Early Stage** |
| 3 | pacingRatio > 1.15 | **Over Utilization** |
| 4 | pacingRatio ≥ 0.85 | **Healthy** |
| 5 | pacingRatio ≥ 0.60 | **Moderate** |
| 6 | pacingRatio ≥ 0.35 | **Low Utilization** |
| 7 | otherwise | **At Risk** |

Consequences:
- **The Utilization % does not decide the Status for Annual contracts.** Two contracts can both be *Over Utilization* with very different percentages (e.g. 41.7% at month 4 and 19.1% at month 1). The same percentage can also mean different statuses at different contract ages.
- After month 12 the expected % stays at 100%, so a live annual contract older than 12 months is judged against its full commit.

### 4.2 Monthly contracts (live): fixed bands

```
Utilization % = CSUM (lifetime) / (Monthly commit × Tenure months) × 100
```

| # | Condition | Status |
|---|---|---|
| 1 | CSUM (lifetime) = 0 | **No CSUM** |
| 2 | Utilization % > 100 | **Over Utilization** |
| 3 | Utilization % ≥ 70 | **Healthy** |
| 4 | Utilization % ≥ 50 | **Moderate** |
| 5 | Utilization % ≥ 30 | **Low Utilization** |
| 6 | otherwise (including 0%) | **At Risk** |

Notes:
- Exactly 100% is **Healthy**; Over Utilization starts above 100%.
- **No CSUM tests lifetime spend.** A contract with zero lifetime spend shows No CSUM.
- A month with no new spend lowers the lifetime average incrementally; it does not reset the percentage to 0%.
- Monthly contracts never get Early Stage and are never pacing-adjusted.

### 4.3 Early Stage (Annual only)

Early Stage stops brand-new annual contracts from being flagged At Risk or Low Utilization purely because they are young.

**All of these must be true:**
1. Contract type is Annual (live).
2. Lifetime CSUM is above zero (zero shows No CSUM).
3. `elapsedMonths ≤ 2`.
4. `pacingRatio < 0.85` (below the Healthy line).

**Behaviour:**
- It *replaces* the status the contract would otherwise get. A young contract that is on pace (ratio ≥ 0.85) keeps its normal status (Healthy, Over Utilization).
- There is a cliff: when tenure reaches 3 months, Early Stage ends and the contract gets its normal status, which may be a sudden drop.
- The first month counts as a full month. A contract that started mid-month (tenure 0) still counts as 1 elapsed month, so it can look slower than it is.
- Constants in the code: `EARLY_STAGE_MAX_MONTHS = 2`, `EARLY_STAGE_BELOW_RATIO = 0.85`. Both are **pending business confirmation**. An alternative noted in the code is a ratio of 0.35 (only would-be At Risk contracts become Early Stage).

### 4.4 Expired

A contract whose end date is before the snapshot date. Status is **Expired** regardless of its utilization, and the bar stays grey. The Utilization % shown is the whole-contract figure (section 3.3), useful as a "how did it finish" number.

### 4.5 Status colours

| Status | Bar / text colour |
|---|---|
| Over Utilization | purple |
| Healthy | green |
| Moderate | yellow |
| Low Utilization | grey |
| Early Stage | blue |
| At Risk | red |
| Expired, No CSUM | grey |

The bar fill is the utilization % capped at 100%; its colour follows the Status, not the percentage.

---

## 5. How CSUM is summed

`buildMgmtOrgs()` de-duplicates rows so one usage line is not counted twice when it appears under overlapping contracts:

- **CSUM (current month):** one line per *end consumer + vendor + CSUM amount* in the reporting month. If the same line appears under two contracts, the one with the later start date wins.
- **CSUM (lifetime):** same idea, per *month + end consumer + vendor + CSUM amount*, for all months up to the reporting month.
- A row's `configuredCSum` is the adjusted cost only if the contract was active during that month; otherwise 0.

Caveat: two genuinely different lines with identical end consumer, vendor and amount in the same month would be collapsed into one.

Overlapping active contracts are also reduced at build time by the Active Orgs view filter (see `README.md`), which keeps one current contract per org.

---

## 6. Commit amount

Calculated per row in `computeFields()` from the part quantity, then the maximum across the contract line's rows is used.

| Case | Commit amount |
|---|---|
| Contract type = Monthly | quantity × 1,000 (including CCB orgs) |
| Org type = CCB, or contract type = Annual | quantity × 100,000 |
| Anything else | 0 |

Monthly is checked first, so a CCB org on a Monthly contract is quantity × 1,000, not × 100,000.

Contract type comes from the part description: contains "MONTH" → Monthly, contains "ANNUM" → Annual.

---

## 7. Tenure

- Live contract: whole months from contract start to the **build date**.
- Expired contract: whole months from start to the contract end date.
- Counted like Excel `DATEDIF(start, end, "m")`, plus one extra month when the end date falls exactly one day before the start-day anniversary (so 1 Sep → 30 Sep is 1 month, 16 Sep → 15 Sep is 12 months).

Because tenure uses the build date, it moves forward each week even if the data does not. Rebuild on publish day.

---

## 8. Worked examples (2026-10-26 export)

| Case | Inputs | Result |
|---|---|---|
| **Annual, Over Utilization** — Deloitte Global Services (2026-06-01 → 2027-05-31) | Lifetime CSUM $354.2M, commit $850M, tenure 4. actual = 41.67%, expected = 4/12 = 33.33%, pacing = 1.25 | **Over Utilization** (1.25 > 1.15). Utilization % shows 41.67%. |
| **Annual, Early Stage** — Telefónica Tech UK (2026-09-01 → 2027-08-31) | Lifetime CSUM $4.06M, commit $80.8M, tenure 1. actual = 5.02%, expected = 8.33%, pacing = 0.60 | **Early Stage** (tenure ≤ 2 and 0.60 < 0.85). Without the rule it would be Moderate. |
| **Annual, young but on pace** — HERE North America (2026-09-01) | actual 9.18%, expected 8.33%, pacing 1.10 | **Healthy**. Early Stage does not apply because pacing ≥ 0.85. |
| **Monthly** — TIVIT (2025-09-01 → 2027-08-31) | Lifetime CSUM to Oct-26 snapshot, monthly commit $2.2M, tenure (build-date dependent) | Lifetime ÷ ($2.2M × tenure). |
| **Expired Annual** — Deloitte (2025-06-01 → 2026-05-31) | Lifetime CSUM $743.3M, commit $700M | 106.18%, Status **Expired**. |
| **Expired Monthly** — DXC Technology Services - FinOps (2025-08-29 → 2026-08-28), 2026-11-02 export | Lifetime CSUM $184.0M, monthly commit $25.8M (25,834 entitlements × $1,000), tenure 12. Utilization = $184.0M ÷ ($25.8M × 12) | 59.34%, Status **Expired**. Contract type and commit come straight from the data. |

---

## 9. Known behaviours and open items

1. **Thresholds not formally confirmed:** the Early Stage limits (2 months, 0.85) and the annual pacing bands (1.15 / 0.85 / 0.60 / 0.35) come from the design document. The code comment marks Early Stage as pending business confirmation.
2. **Past reporting months are not historical for Status:** the table-only Reporting Month dropdown lets you pick any of the last 12 months, but only CSUM and End Consumers are cut at that month. Tenure, the Active/Expired flag and the commit amount stay as of the build/snapshot date, so Status and Utilization % can differ from what the dashboard showed back then. A note under the table says so when a past month is selected. Per-org history is in the Individual Org View.
3. **Tenure follows the build date, not the snapshot date**, so statuses can shift between two builds of the same export.
4. **Softcat Plc UK** has an expired contract line with commit $1,000 (part quantity 0.01), giving 39,899% utilization. This looks like a data error in the source.
5. **TIVIT:** the Oct-20 build held both an Annual ($34M, from 2026-09-02) and a Monthly ($2.2M, from 2025-09-01) contract. The Active Orgs view lists only the Monthly one, so the Annual one is filtered out. Which one is current is still to be confirmed with the view owner.
6. **Unreachable fallback:** the code can return "Expired" from the live-contract branches if commit is 0 (`utilPct === null` / `pacingRatio === null`). Today no live contract has a zero commit, so it never fires.
7. **Utilization % column for Annual vs Status:** the percentage and the Status can look inconsistent by design (section 4.1). The pop-up and Glossary should say so explicitly.
