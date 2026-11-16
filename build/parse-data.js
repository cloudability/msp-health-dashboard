/**
 * MSP Dashboard — Data ingestion script
 *
 * Usage:  node build/parse-data.js
 * Output: MSP Dashboard/MSP-CCB-Performance-Dashboard.html
 *
 * Drop weekly Excel files into Source files/
 * Naming convention (flexible — any of these work):
 *   MSP Reporting Sept 21 2026.xlsx   ← "Month Day Year" detected automatically
 *   MSP Reporting 2026-09-21.xlsx     ← ISO date preferred
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * Configuration
 * ─────────────────────────────────────────────────────────────────────────────
 */

const MAX_WEEKS = 52;          // How many weekly files to retain (oldest dropped first)
const SOURCE_DIR  = 'Source files';
const TEMPLATE    = 'build/template.html';
const OUTPUT      = 'Final/MSP-CCB-Performance-Dashboard.html';
const DATA_PLACEHOLDER = '/*{{DATA_INJECTION}}*/';

// ─────────────────────────────────────────────────────────────────────────────
// MSP Column definitions
// Every column the dashboard needs is declared here by its exact header name.
// ─────────────────────────────────────────────────────────────────────────────

const COL = {
  // Time
  YEAR_MONTH:             'YEAR_MONTH',
  FISCAL_YEAR:            'FISCAL_YEAR',
  FISCAL_QTR:             'FISCAL_QTR',
  FISCAL_MONTH:           'FISCAL_MONTH',

  // Record type
  MSP_OR_CCB:             'MSP_OR_CCB',        // "MSP" | "CCB"

  // Dimensions
  VENDOR:                 'VENDOR',
  REGION:                 'REGION',            // US | EU | AU | CA
  IOT:                    'IOT',               // Americas | EMEA | APAC | Japan
  IMT:                    'IMT',               // Market
  CLIENT_TYPE:            'CLIENT_TYPE',       // Strategic | Enterprise | Select Territory | Select Horizon  — absent in new export
  SIGNING_CATEGORY:       'SIGNING_CATEGORY',  // New Signings | Renewal | Contract Extensions
  SIGNINGS_SUB_CAT:       'SIGNINGS_SUB_CAT',

  // Customer identifiers
  MANAGEMENT_ORG_NAME:    'MANAGEMENT_ORG_NAME',
  END_CONSUMER_NAME:      'END_CONSUMER_NAME',
  CUSTOMER_NAME:          'CUSTOMER_NAME',
  CUSTOMER_NAME_FULL_MRGD:'CUSTOMER_NAME_FULL_MRGD',
  IBM_CUST_NUM_MRGD:      'IBM_CUST_NUM_MRGD',
  IBM_CLIENT_NUM:         'IBM_CLIENT_NUM',    // absent in new export
  IBM_CLIENT_NAME:        'IBM_CLIENT_NAME',   // absent in new export
  OWNER:                  'OWNER',             // absent in new export — replaced by TAM + PSM
  TAM:                    'TAM',               // Technical Account Manager (replaces OWNER in new export)
  PSM:                    'PSM',               // Partner Success Manager (new col)

  // Product
  FAMILY:                 'FAMILY',
  PART_DSCR_FULL:         'PART_DSCR_FULL',
  PART_NUM:               'PART_NUM',
  UT_LEVEL_CODE_30_DSCR:  'UT_LEVEL_CODE_30_DSCR',    // absent in new export
  OFFERING_INVESTMENT_STATUS: 'OFFERING_INVESTMENT_STATUS',

  // Contract
  // SALES_ORDER_NUM removed from new export; SFDCID is the Salesforce deal ID used for deduplication
  SALES_ORDER_NUM:        'SFDCID',            // was 'SALES_ORDER_NUM' — now mapped to SFDCID (col AC)
  SO_LI:                  'SO_LI',             // absent in new export
  COVERAGE_TERM:          'COVERAGE_TERM',
  CONTRACT_START_DATE:    'CONTRACT_START_DATE',
  CONTRACT_END_DATE:      'CONTRACT_END_DATE',
  NEXT_RENEWAL_DATE:      'NEXT_RENEWAL_DATE',

  // Metrics — PRIMARY
  SAAS_ACV_USD_AT_PLAN:   'SAAS_ACV_USD_AT_PLAN',   // primary value metric (USD)
  SAAS_TCV_USD_AT_PLAN:   'SAAS_TCV_USD_AT_PLAN',   // TCV USD
  SAAS_ACV:               'SAAS_ACV',                // local currency ACV
  SAAS_TCV:               'SAAS_TCV',                // absent in new export — local currency TCV
  FMS_AMOUNT:             'FMS_AMOUNT',              // absent in new export

  // Cloud cost (MSP-specific)
  USD_COST:               'USD_COST',
  USD_ADJUSTED_COST:      'USD_ADJUSTED_COST',
  COST:                   'COST',
  CURRENCY_CODE:          'CURRENCY_CODE',     // col 14 — now always "USD" (billing currency)
  CURRENCY:               'CURRENCY',          // col 39 — local/transaction currency (JPY, EUR, etc.)

  // Other
  COUNTRY:                'COUNTRY',
  RENEWAL_COUNTER:        'RENEWAL_COUNTER',   // absent in new export
  REPORTING_DATE:         'REPORTING_DATE',    // absent in new export
  SIGNING_DATE:           'SIGNING_DATE',      // absent in new export
  PAYING_STATUS:          'PAYING_STATUS',
  INACT_FLAG:             'INACT_FLAG',

  // Needed for computed fields
  PART_QTY:               'PART_QTY',   // quantity — used in Committed Amount
  NAME:                   'NAME',       // management org display name — used in Committed Amount

  // Additional fields for data review & dashboard filters
  CUST_NAME_FULL_RESELLER: 'CUST_NAME_FULL_RESELLER',  // absent in new export
  REVENUE_TYPE:            'REVENUE_TYPE',              // absent in new export

  // New columns present in the Oct 2026 export
  SFDCID:                 'SFDCID',            // Salesforce Deal ID (col AC)
  DOMESTIC_CLIENT_ID:     'DOMESTIC_CLIENT_ID',// col AD
  MANAGEMENT_ORG_ID:      'MANAGEMENT_ORG_ID', // col H — numeric org ID
  END_CONSUMER_ORG_ID:    'END_CONSUMER_ORG_ID',// col E
  GLBL_BUYG_GRP_NM:       'GLBL_BUYG_GRP_NM', // col AE — Global Buying Group Name
  CHANNEL:                'CHANNEL',           // col AU
  PART_NUMBER_TOP_BILL:   'PART_NUMBER_TOP_BILL',// col AV
  REVERSAL_FLAG:          'REVERSAL_FLAG',     // col BB
  // TAM / PSM defined earlier (replaced OWNER)
  // TAM and PSM are declared in the customer identifiers section above
  CONTRACT_STATUS:        'CONTRACT_STATUS',   // ACTIVE | EXPIRED
  SENTIMENT:              'SENTIMENT',         // customer sentiment: Green | Yellow | Red | blank (only some orgs are tracked)
};

// ─────────────────────────────────────────────────────────────────────────────
// Dependencies
// ─────────────────────────────────────────────────────────────────────────────

const fs      = require('fs');
const path    = require('path');
const ExcelJS = require('exceljs');

// ─────────────────────────────────────────────────────────────────────────────
// Helpers — filename → ISO date
// ─────────────────────────────────────────────────────────────────────────────

const MONTH_MAP = {
  january: '01', jan: '01',
  february: '02', feb: '02',
  march: '03', mar: '03',
  april: '04', apr: '04',
  may: '05',
  june: '06', jun: '06',
  july: '07', jul: '07',
  august: '08', aug: '08',
  september: '09', sep: '09', sept: '09',
  october: '10', oct: '10',
  november: '11', nov: '11',
  december: '12', dec: '12',
};

function extractDate(filename, mtime) {
  const base = path.basename(filename, path.extname(filename));

  // Strategy 1 — ISO date: 2026-09-21
  const isoMatch = base.match(/(\d{4}[-_]\d{2}[-_]\d{2})/);
  if (isoMatch) return isoMatch[1].replace(/_/g, '-');

  // Strategy 2 — "Month Day Year":  Sept 21 2026
  const mdy = base.match(/([A-Za-z]+)\s+(\d{1,2})\s+(\d{4})/);
  if (mdy) {
    const month = MONTH_MAP[mdy[1].toLowerCase()];
    if (month) return `${mdy[3]}-${month}-${mdy[2].padStart(2, '0')}`;
  }

  // Strategy 3 — "Month Day" (no year):  Sept 21
  const md = base.match(/([A-Za-z]+)\s+(\d{1,2})/);
  if (md) {
    const month = MONTH_MAP[md[1].toLowerCase()];
    if (month) {
      const day  = md[2].padStart(2, '0');
      let   year = new Date().getFullYear();
      const candidate = `${year}-${month}-${day}`;
      if (new Date(candidate) > new Date()) year -= 1;
      return `${year}-${month}-${day}`;
    }
  }

  // Strategy 4 — file mtime fallback
  const d = new Date(mtime);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}

// ─────────────────────────────────────────────────────────────────────────────
// Cell value helper
// ─────────────────────────────────────────────────────────────────────────────

// The source export contains mojibake: a lost non-ASCII character was written as
// U+FFFD, then mis-decoded as the 3 chars "ï¿½". The original letter is gone, so we
// restore the known names; a standalone marker between spaces is an en dash.
const MOJIBAKE = '\u00ef\u00bf\u00bd';
const MOJIBAKE_FIXES = [
  [/TELEF[\u00ef\u00bf\u00bd]{3}NICA/g, 'TELEFÓNICA'],
  [/Telef[\u00ef\u00bf\u00bd]{3}nica/g,  'Telefónica'],
  [/Baltic O[\u00ef\u00bf\u00bd]{3}\(/g, 'Baltic OÜ('],
  [/Inform[\u00ef\u00bf\u00bd]{3}tica/g, 'Informática'],
  [/M[\u00ef\u00bf\u00bd]{3}XICO/g,      'MÉXICO'],
  [/ \u00ef\u00bf\u00bd /g,              ' – '],
];

function fixMojibake(s) {
  if (!s.includes(MOJIBAKE)) return s;
  let out = s;
  for (const [re, rep] of MOJIBAKE_FIXES) out = out.replace(re, rep);
  if (out.includes(MOJIBAKE)) console.warn('  ⚠ Unresolved encoding marker in value: ' + out);
  return out;
}

function cellVal(row, colIndex) {
  if (colIndex === null || colIndex === undefined) return null;
  let v;
  if (Array.isArray(row)) {
    v = row[colIndex];
  } else if (row && typeof row.getCell === 'function') {
    const cell = row.getCell(colIndex);
    if (cell.type === ExcelJS.ValueType.Error) return null;
    v = cell.value;
    if (typeof v === 'object' && v && v.richText) v = v.richText.map(r => r.text).join('');
  } else {
    return null;
  }
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'string') return fixMojibake(v.trim());
  return v;
}

// ─────────────────────────────────────────────────────────────────────────────
// Date helpers used by computed fields
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Parse an ISO date string "YYYY-MM-DD" or "YYYY-MM" into a Date (UTC midnight).
 * Returns null if the input is null/unparseable.
 */
function parseDate(s) {
  if (!s) return null;
  const str = String(s).slice(0, 10); // accept "YYYY-MM-DD" or "YYYY-MM"
  const d = new Date(str.length === 7 ? str + '-01' : str);
  return isNaN(d.getTime()) ? null : d;
}

/** Last day of the month for a given Date (UTC). */
function endOfMonth(d) {
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0));
}

/** First day of the month for a given Date (UTC). */
function startOfMonth(d) {
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), 1));
}

/**
 * Integer months between two dates, truncated (mirrors Excel DATEDIF "m").
 * Returns null if either date is missing.
 */
function datedifMonths(startDate, endDate) {
  if (!startDate || !endDate) return null;
  const s = parseDate(startDate);
  const e = parseDate(endDate);
  if (!s || !e) return null;
  const years  = e.getUTCFullYear() - s.getUTCFullYear();
  const months = e.getUTCMonth()    - s.getUTCMonth();
  const total  = years * 12 + months;
  // Subtract 1 if the end day-of-month hasn't reached the start day-of-month yet
  return e.getUTCDate() >= s.getUTCDate() ? total : total - 1;
}

/**
 * Add N months to a Date (UTC), returning a new Date.
 * Mirrors Excel EDATE(date, -12).
 */
function addMonths(d, n) {
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + n, d.getUTCDate()));
}

// ─────────────────────────────────────────────────────────────────────────────
// Computed fields
// All 7 fields are calculated here at parse time (build step), not in the browser.
// This keeps the dashboard fast and makes the Snowflake migration straightforward
// — these become SQL computed columns.
//
// buildDate:   the date the build script runs (TODAY() equivalent)
// latestCompletedYM: the latest completed/closed YEAR_MONTH (e.g. Oct 2026 when snapshot is Nov 2026)
// ─────────────────────────────────────────────────────────────────────────────

function determineLatestCompletedYM(allYMs, snapshotIsoDate, buildDate) {
  if (!allYMs || allYMs.length === 0) return '';
  const sortedYMs = [...new Set(allYMs)].filter(Boolean).sort();
  if (sortedYMs.length === 0) return '';
  const maxYM = sortedYMs[sortedYMs.length - 1];
  const snapDateStr = snapshotIsoDate || buildDate || '';
  const snapYM = snapDateStr.slice(0, 7); // e.g. '2026-11'

  // If the max YEAR_MONTH in the file is equal to (or after) the snapshot's in-progress calendar month,
  // that month represents partial/in-flight data. The latest completed month is the prior month.
  if (snapYM && maxYM >= snapYM && sortedYMs.length > 1) {
    const priorYMs = sortedYMs.filter(ym => ym < snapYM);
    if (priorYMs.length > 0) {
      return priorYMs[priorYMs.length - 1];
    }
  }
  return maxYM;
}

// currentYM: the in-flight (month-to-date) YEAR_MONTH when the file holds one, i.e. the newest
// YEAR_MONTH if it is later than latestCompletedYM. Otherwise it equals latestCompletedYM.
function determineCurrentYM(latestYM, latestCompletedYM) {
  return (latestYM && latestCompletedYM && latestYM > latestCompletedYM) ? latestYM : latestCompletedYM;
}

function computeFields(raw, buildDate, latestCompletedYM, currentYM) {
  // ── 1. isLatestMonth ──────────────────────────────────────────────────────
  // TRUE if this row's YEAR_MONTH equals the latest completed YEAR_MONTH.
  // Equivalent to: IF(A2=TEXT(MAX_COMPLETED([YEAR_MONTH]),"yyyy-mm"),TRUE,FALSE)
  const isLatestMonth = raw.yearMonth === latestCompletedYM;

  // ── 1b. isCurrentMonth ────────────────────────────────────────────────────
  // TRUE if this row's YEAR_MONTH is the current (month-to-date) month. Falls back to the
  // latest completed month when the file has no in-flight month.
  const isCurrentMonth = raw.yearMonth === currentYM;

  // ── 2. tenureMonths ───────────────────────────────────────────────────────
  // Active contracts: DATEDIF(CONTRACT_START_DATE, TODAY(), "m")
  // Expired contracts (CONTRACT_END_DATE < buildDate): months covered by the contract.
  //
  // DATEDIF "m" counts complete calendar-month crossings and subtracts 1 when the end
  // day-of-month hasn't yet reached the start day-of-month. This under-counts contracts
  // that are commercially "N months" but end one day before the Nth-month anniversary.
  //
  // Round-up rule: add 1 when the effective end date is exactly one day before the
  // start-day anniversary, i.e. endDay === startDay − 1 (or, when startDay === 1,
  // endDay === last day of the end month — the existing end-of-month case).
  //
  // Examples that all become correct:
  //   Sep 1  → Sep 30 : endDay(30) = lastDay(30)              → +1 → 1 month  ✓
  //   Jan 1  → Mar 31 : endDay(31) = lastDay(31)              → +1 → 3 months ✓
  //   Sep 16 → Sep 15 : endDay(15) = startDay(16) − 1         → +1 → 12 months ✓
  //   Jan 2  → Dec 1  : endDay(1)  = startDay(2)  − 1         → +1 → 11 months ✓
  //   Jan 15 → Mar 14 : endDay(14) = startDay(15) − 1         → +1 → 2 months  ✓
  const contractIsExpired = raw.contractStatus ? raw.contractStatus === 'EXPIRED' : Boolean(raw.contractEnd && raw.contractEnd < buildDate);
  const tenureEndDate     = contractIsExpired ? raw.contractEnd : buildDate;
  let   tenureMonths      = datedifMonths(raw.contractStart, tenureEndDate);
  if (tenureMonths !== null && raw.contractStart) {
    const startD    = parseDate(raw.contractStart);
    const endD      = parseDate(tenureEndDate);
    if (startD && endD) {
      const sDay      = startD.getUTCDate();
      const eDay      = endD.getUTCDate();
      const eLastDay  = endOfMonth(endD).getUTCDate();
      // One-day-before rule: end lands exactly 1 day before the start-day anniversary
      const oneDayBefore = sDay === 1 ? eDay === eLastDay : eDay === sDay - 1;
      if (oneDayBefore) tenureMonths += 1;
    }
  }

  // ── 3. contractType ───────────────────────────────────────────────────────
  // Derives a human-readable contract type from PART_DSCR_FULL (column BY = col 77).
  // Equivalent to:
  //   IF(ISNUMBER(SEARCH("MONTH",BY2)),"Monthly",
  //   IF(ISNUMBER(SEARCH("ANNUM",BY2)),"Annual",""))
  //
  const pd = (raw.partDesc || '').toUpperCase();
  const contractType = pd.includes('MONTH') ? 'Monthly'
                     : pd.includes('ANNUM') ? 'Annual'
                     : '';

  // ── 4. configuredCSum ─────────────────────────────────────────────────────
  // Includes USD_ADJUSTED_COST only when the contract was active during YEAR_MONTH.
  // Equivalent to:
  //   IF(AND(CONTRACT_START_DATE <= EOMONTH(YEAR_MONTH, 0),
  //          CONTRACT_END_DATE   >= DATE(year(YEAR_MONTH), month(YEAR_MONTH), 1)),
  //      USD_ADJUSTED_COST, 0)
  let configuredCSum = 0;
  if (raw.yearMonth && raw.contractStart && raw.contractEnd) {
    const ymDate  = parseDate(raw.yearMonth);   // first day of YEAR_MONTH
    if (ymDate) {
      const ymStart = startOfMonth(ymDate);     // 1st of YEAR_MONTH
      const ymEnd   = endOfMonth(ymDate);       // last day of YEAR_MONTH
      const cStart  = parseDate(raw.contractStart);
      const cEnd    = parseDate(raw.contractEnd);
      if (cStart && cEnd && cStart <= ymEnd && cEnd >= ymStart) {
        configuredCSum = raw.usdAdjustedCost;
      }
    }
  }

  // ── 5. committedAmount ───────────────────────────────────────────────────
  // Calculates a committed spend figure based on type, contract type, and customer name.
  // Equivalent to:
  //   IF(ContractType="Monthly",    PART_QTY * 1000,
  //   IF(OR(MSP_OR_CCB="CCB", ContractType="Annual"), PART_QTY * 100000, 0))
  // A Monthly contract is always qty x 1,000, including CCB orgs (confirmed with the business).
  const qty = Number(raw.partQty) || 0;
  let committedAmount = 0;
  if (contractType === 'Monthly') {
    committedAmount = qty * 1000;
  } else if (raw.type === 'CCB' || contractType === 'Annual') {
    committedAmount = qty * 100000;
  }

  // ── 6. isL12Months ───────────────────────────────────────────────────────
  // TRUE if YEAR_MONTH falls within the 12 months ending with currentYM (the current month is the 12th).
  // (e.g. 2025-12 through 2026-11 when currentYM is 2026-11).
  let isL12Months = false;
  if (raw.yearMonth && currentYM) {
    const l12End = parseDate(currentYM);
    if (l12End) {
      const l12Start = addMonths(l12End, -11);
      const startYM = l12Start.toISOString().slice(0, 7);
      isL12Months = raw.yearMonth >= startYM && raw.yearMonth <= currentYM;
    }
  }

  // ── 7. endConsumerName2 ───────────────────────────────────────────────────
  // Pass-through of END_CONSUMER_NAME, returning null when blank.
  // Equivalent to: IF([@END_CONSUMER_NAME]="","",[@END_CONSUMER_NAME])
  const endConsumerName2 = raw.endConsumerName || null;

  return { isLatestMonth, isCurrentMonth, tenureMonths, contractType, configuredCSum,
           committedAmount, isL12Months, endConsumerName2 };
}

// ─────────────────────────────────────────────────────────────────────────────
// Parse a single MSP worksheet row into a canonical record
// ─────────────────────────────────────────────────────────────────────────────

/**
 * parseRow extracts raw column values.
 * Computed fields are added afterwards by computeFields() once latestCompletedYM and currentYM are known.
 * buildDate, latestCompletedYM and currentYM are passed in from main().
 */
function parseRow(row, hdr, buildDate, latestCompletedYM, currentYM) {
  const g = k => {
    const idx = hdr[k];
    return (idx !== undefined && idx !== null) ? cellVal(row, idx) : null;
  };

  // Skip inactive rows
  const inact = g(COL.INACT_FLAG);
  if (inact === 1 || inact === true || inact === '1' || inact === 'true') return null;

  const raw = {
    // Time
    yearMonth:    g(COL.YEAR_MONTH),
    fiscalYear:   g(COL.FISCAL_YEAR),
    fiscalQtr:    g(COL.FISCAL_QTR),
    fiscalMonth:  g(COL.FISCAL_MONTH),

    // Type
    type:         g(COL.MSP_OR_CCB),          // "MSP" | "CCB"

    // Dimensions
    vendor:       g(COL.VENDOR),
    region:       g(COL.REGION),
    iot:          g(COL.IOT),
    imt:          g(COL.IMT),
    clientType:   g(COL.CLIENT_TYPE),
    signingCategory:    g(COL.SIGNING_CATEGORY),
    signingsSubCat:     g(COL.SIGNINGS_SUB_CAT),

    // Customer
    managementOrgName:  g(COL.MANAGEMENT_ORG_NAME),
    endConsumerName:    g(COL.END_CONSUMER_NAME),
    customerName:       g(COL.CUSTOMER_NAME) || g(COL.CUSTOMER_NAME_FULL_MRGD),
    customerNameFull:   g(COL.CUSTOMER_NAME_FULL_MRGD),
    ibmCustNum:         g(COL.IBM_CUST_NUM_MRGD),
    ibmClientNum:       g(COL.IBM_CLIENT_NUM),
    ibmClientName:      g(COL.IBM_CLIENT_NAME),
    owner:              g(COL.OWNER) || g(COL.TAM),   // OWNER absent in new export; fall back to TAM
    tam:                g(COL.TAM),
    psm:                g(COL.PSM),

    // Product
    family:             g(COL.FAMILY),
    product:            g(COL.UT_LEVEL_CODE_30_DSCR),
    partDesc:           g(COL.PART_DSCR_FULL),
    partNum:            g(COL.PART_NUM),
    offeringStatus:     g(COL.OFFERING_INVESTMENT_STATUS),

    // Contract
    salesOrderNum:      g(COL.SALES_ORDER_NUM),
    soLi:               g(COL.SO_LI),
    coverageTerm:       g(COL.COVERAGE_TERM),
    contractStart:      g(COL.CONTRACT_START_DATE),
    contractEnd:        g(COL.CONTRACT_END_DATE),
    nextRenewalDate:    g(COL.NEXT_RENEWAL_DATE),
    renewalCounter:     g(COL.RENEWAL_COUNTER),
    reportingDate:      g(COL.REPORTING_DATE),
    signingDate:        g(COL.SIGNING_DATE),

    // Metrics
    acvUsd:    Number(g(COL.SAAS_ACV_USD_AT_PLAN)) || 0,   // PRIMARY metric
    tcvUsd:    Number(g(COL.SAAS_TCV_USD_AT_PLAN)) || 0,
    acvLocal:  Number(g(COL.SAAS_ACV))  || 0,
    tcvLocal:  Number(g(COL.SAAS_TCV))  || 0,
    fmsAmount: Number(g(COL.FMS_AMOUNT)) || 0,

    // Cloud cost
    usdCost:         Number(g(COL.USD_COST))          || 0,
    usdAdjustedCost: Number(g(COL.USD_ADJUSTED_COST)) || 0,
    currencyCode:    g(COL.CURRENCY) || g(COL.CURRENCY_CODE),  // CURRENCY (col 39) = local currency; CURRENCY_CODE (col 14) = billing USD

    // Meta
    country:      g(COL.COUNTRY),
    payingStatus: g(COL.PAYING_STATUS),

    // Raw columns needed only for computed fields (not surfaced directly)
    partQty: g(COL.PART_QTY),
    name:    g(COL.NAME) || g(COL.CUSTOMER_NAME_FULL_MRGD) || g(COL.CUSTOMER_NAME),

    // Additional fields (may be absent in newer exports — gracefully null)
    custNameFullReseller: g(COL.CUST_NAME_FULL_RESELLER),
    revenueType:          g(COL.REVENUE_TYPE),

    // New columns from Oct 2026 export
    sfdcId:              g(COL.SFDCID),
    domesticClientId:    g(COL.DOMESTIC_CLIENT_ID),
    managementOrgId:     g(COL.MANAGEMENT_ORG_ID),
    endConsumerOrgId:    g(COL.END_CONSUMER_ORG_ID),
    glblBuyingGroupNm:   g(COL.GLBL_BUYG_GRP_NM),
    channel:             g(COL.CHANNEL),
    partNumberTopBill:   g(COL.PART_NUMBER_TOP_BILL),
    reversalFlag:        g(COL.REVERSAL_FLAG),
    contractStatus:      g(COL.CONTRACT_STATUS),
    sentiment:           g(COL.SENTIMENT),
  };

  // Attach computed fields
  return Object.assign(raw, computeFields(raw, buildDate, latestCompletedYM, currentYM));
}

// ─────────────────────────────────────────────────────────────────────────────
// Active Management Orgs view (V_SSABL_MSP_ACTIVE) filter
//
// The Excel export can carry more than one contract line per org that is
// "active by date" (overlapping/duplicate contracts). The Snowflake view
// lists exactly one current contract per org. When a view CSV is present in
// "Source files/Active Orgs/", an active-by-date row is kept only if its
// DOMESTIC_CLIENT_ID + CONTRACT_START_DATE + CONTRACT_END_DATE is in the view.
// Expired / historical rows are never touched.
// ─────────────────────────────────────────────────────────────────────────────

const ACTIVE_VIEW_DIR = 'Active Orgs';

/** Minimal RFC-4180 CSV parser (quoted fields, escaped quotes, CRLF, BOM). */
function parseCsv(text) {
  if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
  const rows = [];
  let row = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQ) {
      if (c === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false;
      } else field += c;
    } else if (c === '"') inQ = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); field = '';
      if (row.length > 1 || row[0] !== '') rows.push(row);
      row = [];
    } else field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

/** Date of a view file: dd_mm_yy in the name (e.g. Active_Management_Orgs_26_10_26.csv), else mtime. */
function extractViewDate(filename, mtime) {
  const m = filename.match(/(\d{2})[_-](\d{2})[_-](\d{2})(?!\d)/);
  if (m) return `20${m[3]}-${m[2]}-${m[1]}`;
  return mtime.toISOString().slice(0, 10);
}

/** Newest view CSV in Source files/Active Orgs/ (subfolders ignored), or null. */
function findActiveViewFile(sourceDir) {
  const dir = path.join(sourceDir, ACTIVE_VIEW_DIR);
  if (!fs.existsSync(dir)) return null;
  const files = fs.readdirSync(dir)
    .filter(f => /\.csv$/i.test(f) && !f.startsWith('~$'))
    .map(f => {
      const full = path.join(dir, f);
      const st = fs.statSync(full);
      return { filename: f, fullPath: full, isoDate: extractViewDate(f, st.mtime), mtime: st.mtimeMs };
    })
    .sort((a, b) => b.isoDate.localeCompare(a.isoDate) || b.mtime - a.mtime);
  return files[0] || null;
}

const viewKey = (id, start, end) =>
  `${String(id || '').trim()}|${String(start || '').slice(0, 10)}|${String(end || '').slice(0, 10)}`;

/** Load the view CSV -> { file, keys:Set, rows:[{key,name,id,start,end}] } */
function loadActiveView(file) {
  const table = parseCsv(fs.readFileSync(file.fullPath, 'utf8'));
  const head = table[0].map(h => h.trim().toUpperCase());
  const ix = n => head.indexOf(n);
  const need = ['DOMESTIC_CLIENT_ID', 'CONTRACT_START_DATE', 'CONTRACT_END_DATE'];
  const missing = need.filter(n => ix(n) < 0);
  if (missing.length) throw new Error(`missing column(s): ${missing.join(', ')}`);
  const rows = table.slice(1).map(r => {
    const id = (r[ix('DOMESTIC_CLIENT_ID')] || '').trim();
    const start = (r[ix('CONTRACT_START_DATE')] || '').trim().slice(0, 10);
    const end = (r[ix('CONTRACT_END_DATE')] || '').trim().slice(0, 10);
    const name = ix('CUSTOMER_NAME_FULL_MRGD') >= 0 ? (r[ix('CUSTOMER_NAME_FULL_MRGD')] || '').trim() : id;
    return { key: viewKey(id, start, end), id, start, end, name };
  }).filter(r => r.id);
  return { file, rows, keys: new Set(rows.map(r => r.key)) };
}

/**
 * Drop active-by-date rows that are not a current contract in the view.
 * Returns { kept, dropped, report } — report is printed by the caller.
 */
function applyActiveViewFilter(records, snapshotDate, view) {
  const kept = [];
  const orgs = new Map();   // id -> { name, before, after, dropped, months:Set }
  const matchedKeys = new Set();
  const activeOrgIds = new Set();
  const getOrg = r => {
    const id = r.domesticClientId || '(no id)';
    if (!orgs.has(id)) orgs.set(id, { id, name: r.customerNameFull || r.managementOrgName || r.name || id, before: 0, after: 0, dropped: 0, months: new Set() });
    return orgs.get(id);
  };

  for (const r of records) {
    const o = getOrg(r);
    const csum = Number(r.configuredCSum) || 0;
    o.before += csum;
    const start = String(r.contractStart || '').slice(0, 10);
    const end = String(r.contractEnd || '').slice(0, 10);
    const activeByDate = start && end && start <= snapshotDate && snapshotDate <= end;
    if (activeByDate) activeOrgIds.add(r.domesticClientId || '(no id)');
    const key = viewKey(r.domesticClientId, start, end);
    if (!activeByDate || view.keys.has(key)) {
      if (activeByDate) matchedKeys.add(key);
      kept.push(r);
      o.after += csum;
    } else {
      o.dropped++;
      if (r.yearMonth) o.months.add(String(r.yearMonth).slice(0, 7));
    }
  }

  const changed = [...orgs.values()].filter(o => o.dropped > 0).sort((a, b) => (b.before - b.after) - (a.before - a.after));
  const orgsWithNoKeptActive = changed.filter(o => !kept.some(r => (r.domesticClientId || '(no id)') === o.id &&
    String(r.contractStart || '').slice(0, 10) <= snapshotDate && snapshotDate <= String(r.contractEnd || '').slice(0, 10)));
  const unmatchedView = view.rows.filter(v => !matchedKeys.has(v.key));
  return { kept, changed, orgsWithNoKeptActive, unmatchedView, activeOrgCount: activeOrgIds.size };
}

function printActiveViewReport(res, view, snapshotDate, droppedTotal) {
  const money = n => '$' + Math.round(n).toLocaleString('en-US');
  console.log(`  Active-orgs view: ${view.file.filename} (view date ${view.file.isoDate}), ${view.rows.length} contract rows`);
  if (view.file.isoDate !== snapshotDate) {
    console.log(`  WARNING: view date ${view.file.isoDate} differs from export snapshot ${snapshotDate} — make sure both were pulled together.`);
  }
  console.log(`  Active-by-date orgs in data: ${res.activeOrgCount}; rows dropped (not current in view): ${droppedTotal}`);
  if (res.changed.length) {
    console.log('  Reconciliation (orgs with dropped rows):');
    let b = 0, a = 0;
    for (const o of res.changed) {
      b += o.before; a += o.after;
      const ms = [...o.months].sort();
      console.log(`    - ${o.name}: ${o.dropped} rows dropped, CSUM ${money(o.before)} -> ${money(o.after)}, months ${ms[0]}..${ms[ms.length - 1]} (${ms.length})`);
    }
    console.log(`    Total CSUM across affected orgs: ${money(b)} -> ${money(a)} (${money(a - b)})`);
  }
  if (res.orgsWithNoKeptActive.length) {
    console.log('  WARNING: these orgs are active by date in the data but have NO contract in the view, so all their active rows were dropped:');
    res.orgsWithNoKeptActive.forEach(o => console.log(`    - ${o.name} (${o.id})`));
  }
  if (res.unmatchedView.length) {
    console.log('  WARNING: view contracts with no matching rows in the export (check IDs/dates):');
    res.unmatchedView.forEach(v => console.log(`    - ${v.name} (${v.id}) ${v.start} -> ${v.end}`));
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// Main build routine
// ─────────────────────────────────────────────────────────────────────────────

async function main() {
  // Build date = TODAY() equivalent — fixed at the moment this script runs.
  // Using ISO date string so it's consistent with parseDate().
  const today = new Date();
  const buildDate = `${today.getUTCFullYear()}-${String(today.getUTCMonth()+1).padStart(2,'0')}-${String(today.getUTCDate()).padStart(2,'0')}`;

  const baseDir      = path.join(path.dirname(__filename), '..');
  const sourceDir    = path.join(baseDir, SOURCE_DIR);
  const templatePath = path.join(baseDir, TEMPLATE);
  const outputPath   = path.join(baseDir, OUTPUT);

  // ── 1. Discover Excel files ────────────────────────────────────────────────

  if (!fs.existsSync(sourceDir)) {
    console.log(`ERROR: Source directory not found: ${sourceDir}`);
    process.exit(1);
  }

  const allFiles = fs.readdirSync(sourceDir)
    .filter(f => /\.(xlsx|xls|csv)$/i.test(f) && !/glossary/i.test(f) && !f.startsWith('~$'))
    .map(f => {
      const fullPath = path.join(sourceDir, f);
      const stat     = fs.statSync(fullPath);
      const isoDate  = extractDate(f, stat.mtime);
      return { filename: f, fullPath, isoDate };
    })
    .sort((a, b) => a.isoDate.localeCompare(b.isoDate));

  if (allFiles.length === 0) {
    console.log('WARNING: No data files (.csv, .xlsx) found in Source files/. Build will produce an empty dashboard.');
  }

  // Apply MAX_WEEKS cap
  const files = allFiles.length > MAX_WEEKS
    ? allFiles.slice(allFiles.length - MAX_WEEKS)
    : allFiles;

  if (allFiles.length > MAX_WEEKS) {
    console.log(`INFO: ${allFiles.length} files found; retaining most recent ${MAX_WEEKS} (MAX_WEEKS cap).`);
  }

  // ── 2. Parse each file ────────────────────────────────────────────────────

  /**
   * weeklySnapshots: array of { snapshotDate, filename, records[] }
   *
   * Each record is a canonical MSP row (see parseRow above).
   * YEAR_MONTH inside each record represents the billing period — it is
   * independent of snapshotDate (which is when the file was pulled).
   */
  const weeklySnapshots = [];

  // Optional: current-contract list from the V_SSABL_MSP_ACTIVE view export
  let activeView = null;
  const activeViewFile = findActiveViewFile(sourceDir);
  if (activeViewFile) {
    try {
      activeView = loadActiveView(activeViewFile);
    } catch (err) {
      console.log(`ERROR: Could not read active-orgs view ${activeViewFile.filename}: ${err.message}`);
      process.exit(1);
    }
  }
  let totalRecords = 0;
  let skippedRows  = 0;

  for (const file of files) {
    console.log(`Parsing [${file.isoDate}] ${file.filename} …`);
    const records = [];
    const isCsv = /\.csv$/i.test(file.filename);

    if (isCsv) {
      let content;
      try {
        content = fs.readFileSync(file.fullPath, 'utf8');
      } catch (err) {
        console.log(`  WARNING: Could not read CSV file — skipping. (${err.message})`);
        continue;
      }
      const rows = parseCsv(content);
      if (!rows.length) {
        console.log('  WARNING: CSV file is empty — skipping.');
        continue;
      }
      const hdr = {};
      rows[0].forEach((cell, idx) => {
        const key = String(cell ?? '').trim().replace(/^\uFEFF/, '');
        if (key) hdr[key] = idx;
      });
      const requiredCols = [COL.YEAR_MONTH, COL.SAAS_ACV_USD_AT_PLAN, COL.MSP_OR_CCB];
      const missing = requiredCols.filter(c => hdr[c] === undefined);
      if (missing.length) {
        console.log(`  WARNING: CSV missing required columns (${missing.join(', ')}) — skipping file.`);
        continue;
      }

      // Pass 1: find all distinct YMs and determine latest completed YM
      const ymCol = hdr[COL.YEAR_MONTH];
      const allYMs = [];
      let latestYM = '';
      for (let i = 1; i < rows.length; i++) {
        const v = cellVal(rows[i], ymCol);
        if (v) {
          const strV = String(v);
          allYMs.push(strV);
          if (strV > latestYM) latestYM = strV;
        }
      }
      const latestCompletedYM = determineLatestCompletedYM(allYMs, file.isoDate, buildDate);
      const currentYM = determineCurrentYM(latestYM, latestCompletedYM);
      console.log(`  Latest YEAR_MONTH in file: ${latestYM || '(none)'} (latest completed: ${latestCompletedYM || '(none)'}, current: ${currentYM || '(none)'})`);

      // Pass 2: parse rows
      for (let i = 1; i < rows.length; i++) {
        const rec = parseRow(rows[i], hdr, buildDate, latestCompletedYM, currentYM);
        if (rec) records.push(rec);
        else skippedRows++;
      }
    } else {
      const workbook = new ExcelJS.Workbook();
      try {
        await workbook.xlsx.readFile(file.fullPath);
      } catch (err) {
        console.log(`  WARNING: Could not read file — skipping. (${err.message})`);
        continue;
      }

      // Find the worksheet that contains the expected columns.
      let ws = null;
      let hdr = {};
      for (const sheet of workbook.worksheets) {
        const candidateHdr = {};
        sheet.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
          const key = String(cell.value ?? '').trim();
          if (key) candidateHdr[key] = col;
        });
        const requiredCols = [COL.YEAR_MONTH, COL.SAAS_ACV_USD_AT_PLAN, COL.MSP_OR_CCB];
        if (requiredCols.every(c => candidateHdr[c])) {
          ws = sheet;
          hdr = candidateHdr;
          console.log(`  Using worksheet: "${sheet.name}"`);
          break;
        }
      }
      if (!ws) {
        const allHdrs = {};
        workbook.worksheets.forEach(sheet =>
          sheet.getRow(1).eachCell({ includeEmpty: true }, (cell, col) => {
            const key = String(cell.value ?? '').trim();
            if (key) allHdrs[key] = col;
          })
        );
        const missing = [COL.YEAR_MONTH, COL.SAAS_ACV_USD_AT_PLAN, COL.MSP_OR_CCB].filter(c => !allHdrs[c]);
        console.log(`  WARNING: No worksheet found with required columns (${missing.join(', ')}) — skipping file.`);
        continue;
      }

      // Pass 1: determine latest completed YM
      const ymCol = hdr[COL.YEAR_MONTH];
      const allYMs = [];
      let latestYM = '';
      ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
        if (rowNum === 1) return;
        const v = cellVal(row, ymCol);
        if (v) {
          const strV = String(v);
          allYMs.push(strV);
          if (strV > latestYM) latestYM = strV;
        }
      });
      const latestCompletedYM = determineLatestCompletedYM(allYMs, file.isoDate, buildDate);
      const currentYM = determineCurrentYM(latestYM, latestCompletedYM);
      console.log(`  Latest YEAR_MONTH in file: ${latestYM || '(none)'} (latest completed: ${latestCompletedYM || '(none)'}, current: ${currentYM || '(none)'})`);

      // Pass 2: parse rows
      ws.eachRow({ includeEmpty: false }, (row, rowNum) => {
        if (rowNum === 1) return;
        const rec = parseRow(row, hdr, buildDate, latestCompletedYM, currentYM);
        if (rec) records.push(rec);
        else skippedRows++;
      });
    }

    console.log(`  Rows: ${records.length} parsed, ${skippedRows} skipped (inactive/empty)`);

    // Check if CONTRACT_STATUS column is provided in the export data
    const hasContractStatus = records.some(r => r.contractStatus);
    let finalRecords = records;

    if (hasContractStatus) {
      const activeCount = records.filter(r => r.contractStatus === 'ACTIVE').length;
      const expiredCount = records.filter(r => r.contractStatus === 'EXPIRED').length;
      console.log(`  Active tracking: Using CONTRACT_STATUS from export (${activeCount} active rows, ${expiredCount} expired rows).`);
    } else if (activeView && file === files[files.length - 1]) {
      const res = applyActiveViewFilter(records, file.isoDate, activeView);
      printActiveViewReport(res, activeView, file.isoDate, records.length - res.kept.length);
      finalRecords = res.kept;
    } else if (!activeView) {
      console.log('  WARNING: No CONTRACT_STATUS column and no active-orgs view CSV — overlapping contracts will not be filtered.');
    }
    totalRecords += finalRecords.length;

    weeklySnapshots.push({
      snapshotDate: file.isoDate,
      filename:     file.filename,
      records:      finalRecords,
    });
  }

  // ── 2b. Parse Glossary if present ──────────────────────────────────────────
  const glossary = [];
  const glossaryPath = path.join(sourceDir, 'Glossary of Terms.xlsx');
  if (fs.existsSync(glossaryPath)) {
    try {
      const gwb = new ExcelJS.Workbook();
      await gwb.xlsx.readFile(glossaryPath);
      const gws = gwb.getWorksheet(1);
      gws.eachRow({ includeEmpty: false }, (row, rowNum) => {
        if (rowNum === 1) return; // header
        const cat = String(row.getCell(1).value || '').trim();
        const term = String(row.getCell(2).value || '').trim();
        const def = String(row.getCell(3).value || '').trim();
        if (term || def) {
          glossary.push({ category: cat, term, definition: def });
        }
      });
      console.log(`Glossary parsed: ${glossary.length} terms loaded`);
    } catch (err) {
      console.log(`  WARNING: Could not parse Glossary (${err.message})`);
    }
  }

  // ── 3. Build dimension index (for filter dropdowns) ───────────────────────

  const dims = {
    regions:    new Set(),
    iots:       new Set(),
    imts:       new Set(),
    vendors:    new Set(),
    clientTypes:new Set(),
  };

  for (const snap of weeklySnapshots) {
    for (const r of snap.records) {
      if (r.region)     dims.regions.add(r.region);
      if (r.iot)        dims.iots.add(r.iot);
      if (r.imt)        dims.imts.add(r.imt);
      if (r.vendor)     dims.vendors.add(r.vendor);
      if (r.clientType) dims.clientTypes.add(r.clientType);
    }
  }

  const dimensions = {
    regions:     [...dims.regions].sort(),
    iots:        [...dims.iots].sort(),
    imts:        [...dims.imts].sort(),
    vendors:     [...dims.vendors].sort(),
    clientTypes: [...dims.clientTypes].sort(),
  };

  // ── 4. Assemble the data blob ─────────────────────────────────────────────

  /**
   * window.__MSP_DATA__ structure injected into dashboard.html:
   * {
   *   generatedAt:      ISO timestamp of when build ran
   *   maxWeeks:         MAX_WEEKS constant
   *   snapshotDates:    ["YYYY-MM-DD", ...] sorted ASC — one per file
   *   dimensions:       { regions, iots, imts, vendors, clientTypes }
   *   weeklySnapshots:  [ { snapshotDate, filename, records[] } ]
   * }
   *
   * Future Snowflake migration: replace this injection with a fetch() call
   * in template.html. All metric logic stays unchanged.
   */
  const dashboardData = {
    generatedAt:     new Date().toISOString(),
    maxWeeks:        MAX_WEEKS,
    snapshotDates:   weeklySnapshots.map(s => s.snapshotDate),
    dimensions,
    glossary,
    weeklySnapshots,
  };

  const dataScript = `window.__MSP_DATA__ = ${JSON.stringify(dashboardData)};\n`;

  // ── 5. Inject into template ────────────────────────────────────────────────

  if (!fs.existsSync(templatePath)) {
    console.log(`ERROR: Template not found: ${templatePath}`);
    process.exit(1);
  }

  const template = fs.readFileSync(templatePath, 'utf8');

  if (!template.includes(DATA_PLACEHOLDER)) {
    console.log(`ERROR: Template missing placeholder: ${DATA_PLACEHOLDER}`);
    process.exit(1);
  }

  // Backup existing output file before overwriting
  const backupDir = path.join(baseDir, 'Final', 'Backup');
  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
  }
  if (fs.existsSync(outputPath)) {
    const backupPath = path.join(backupDir, path.basename(outputPath));
    fs.copyFileSync(outputPath, backupPath);
  }

  const output = template.replace(DATA_PLACEHOLDER, dataScript);
  fs.writeFileSync(outputPath, output, 'utf8');

  // ── 6. Build summary ──────────────────────────────────────────────────────

  console.log('');
  console.log('─────────────────────────────────────────────────────');
  console.log('MSP Dashboard build complete');
  console.log(`  Files parsed  : ${weeklySnapshots.length} of ${allFiles.length} found`);
  console.log(`  Snapshots     : ${weeklySnapshots.map(s => s.snapshotDate).join(', ') || '(none)'}`);
  console.log(`  Total records : ${totalRecords}`);
  console.log(`  Skipped rows  : ${skippedRows}`);
  console.log(`  Output        : ${outputPath}`);
  console.log(`  File size     : ${(fs.statSync(outputPath).size / 1024).toFixed(0)} KB`);
  console.log('─────────────────────────────────────────────────────');
}

// Run the build only when executed directly (node build/parse-data.js);
// build/update.js requires this file just to reuse extractDate().
if (require.main === module) {
  main().catch(err => {
    console.log('Build failed:', err);
    process.exit(1);
  });
}
module.exports = { extractDate, findActiveViewFile, ACTIVE_VIEW_DIR };
