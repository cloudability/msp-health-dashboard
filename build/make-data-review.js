/**
 * MSP Dashboard — Data Review page generator
 *
 * Usage:  node build/make-data-review.js   (or: npm run review)
 *
 * Columns: Excel source order, then 7 computed fields last.
 * Filters: CUSTOMER_NAME_FULL_MRGD, YEAR_MONTH, VENDOR, Contract Type.
 * Numbers: 2 decimal places throughout.
 * Export:  CSV download button.
 */

'use strict';
const fs   = require('fs');
const path = require('path');

const baseDir    = path.join(__dirname, '..');
const srcHtml    = path.join(baseDir, 'Final', 'MSP-CCB-Performance-Dashboard.html');
const outputPath = path.join(baseDir, 'data-review.html');

if (!fs.existsSync(srcHtml)) {
  console.log('ERROR: MSP-CCB-Performance-Dashboard.html not found. Run npm run build first.');
  process.exit(1);
}

const src   = fs.readFileSync(srcHtml, 'utf8');
const start = src.indexOf('window.__MSP_DATA__ = ');
const end   = src.indexOf(';\n', start);
if (start === -1 || end === -1) {
  console.log('ERROR: Could not find __MSP_DATA__ in MSP-CCB-Performance-Dashboard.html.');
  process.exit(1);
}
const data = JSON.parse(src.slice(start + 'window.__MSP_DATA__ = '.length, end));
const recs = data.weeklySnapshots[0]?.records ?? [];
console.log(`Loaded ${recs.length} records from snapshot ${data.weeklySnapshots[0]?.snapshotDate}`);

// ── Column definitions — Excel source order, computed fields last ─────────
const COLS = [
  { key: 'yearMonth',             label: 'YEAR_MONTH' },            // 1  — displayed as "January 2026"
  { key: 'vendor',                label: 'VENDOR' },                // 2
  { key: 'endConsumerName',       label: 'END_CONSUMER_NAME' },     // 4
  { key: 'managementOrgName',     label: 'MANAGEMENT_ORG_NAME' },   // 7
  { key: 'region',                label: 'REGION' },                // 9
  { key: 'usdCost',               label: 'USD_COST',          dp:2 },// 15
  { key: 'currencyCode',          label: 'CURRENCY_CODE' },         // 14
  { key: 'usdAdjustedCost',       label: 'USD_ADJUSTED_COST', dp:2 },// 16
  { key: 'type',                  label: 'MSP_OR_CCB' },            // 163
  { key: 'payingStatus',          label: 'PAYING_STATUS' },         // 18
  { key: 'name',                  label: 'NAME' },                  // 20
  { key: 'customerName',          label: 'CUSTOMER_NAME' },         // 24
  { key: 'salesOrderNum',         label: 'SALES_ORDER_NUM' },       // 26
  { key: 'soLi',                  label: 'SO_LI' },                 // 30
  { key: 'renewalCounter',        label: 'RENEWAL_COUNTER' },       // 32
  { key: 'partQty',               label: 'PART_QTY' },              // 34
  { key: 'customerNameFull',      label: 'CUSTOMER_NAME_FULL_MRGD' },// 35
  { key: 'ibmCustNum',            label: 'IBM_CUST_NUM_MRGD' },     // 36
  { key: 'signingCategory',       label: 'SIGNING_CATEGORY' },      // 40
  { key: 'signingsSubCat',        label: 'SIGNINGS_SUB_CAT' },      // 41
  { key: 'coverageTerm',          label: 'COVERAGE_TERM' },         // 42
  { key: 'reportingDate',         label: 'REPORTING_DATE' },        // 46
  { key: 'signingDate',           label: 'SIGNING_DATE' },          // 47
  { key: 'tcvLocal',              label: 'SAAS_TCV',          dp:2 },// 49
  { key: 'acvLocal',              label: 'SAAS_ACV',          dp:2 },// 50
  { key: 'tcvUsd',                label: 'SAAS_TCV_USD_AT_PLAN',dp:2},// 51
  { key: 'acvUsd',                label: 'SAAS_ACV_USD_AT_PLAN',dp:2},// 52
  { key: 'fiscalYear',            label: 'FISCAL_YEAR' },           // 59
  { key: 'fiscalMonth',           label: 'FISCAL_MONTH' },          // 60
  { key: 'fiscalQtr',             label: 'FISCAL_QTR' },            // 61
  { key: 'partNum',               label: 'PART_NUM' },              // 76
  { key: 'partDesc',              label: 'PART_DSCR_FULL' },        // 77
  { key: 'family',                label: 'FAMILY' },                // 81
  { key: 'country',               label: 'COUNTRY' },               // 82
  { key: 'iot',                   label: 'IOT' },                   // 83
  { key: 'imt',                   label: 'IMT' },                   // 84
  { key: 'contractStart',         label: 'CONTRACT_START_DATE' },   // 85
  { key: 'contractEnd',           label: 'CONTRACT_END_DATE' },     // 86
  { key: 'nextRenewalDate',       label: 'NEXT_RENEWAL_DATE' },     // 89
  { key: 'custNameFullReseller',  label: 'CUST_NAME_FULL_RESELLER' },// 92
  { key: 'product',               label: 'UT_LEVEL_CODE_30_DSCR' }, // 104
  { key: 'offeringStatus',        label: 'OFFERING_INVESTMENT_STATUS' },// 121
  { key: 'fmsAmount',             label: 'FMS_AMOUNT',        dp:2 },// 123
  { key: 'ibmClientNum',          label: 'IBM_CLIENT_NUM' },        // 129
  { key: 'ibmClientName',         label: 'IBM_CLIENT_NAME' },       // 130
  { key: 'revenueType',           label: 'REVENUE_TYPE' },          // 135
  { key: 'clientType',            label: 'CLIENT_TYPE' },           // 158
  // ── Computed fields ───────────────────────────────────────────────────────
  { key: 'isLatestMonth',         label: 'Latest Month',          computed: true },
  { key: 'tenureMonths',          label: 'Tenure (Months)',        computed: true },
  { key: 'contractType',          label: 'Contract Type',          computed: true },
  { key: 'configuredCSum',        label: 'Configured CSum',        computed: true, dp:2 },
  { key: 'committedAmount',       label: 'Committed Amount',       computed: true, dp:2 },
  { key: 'isL12Months',           label: 'L12 Months?',            computed: true },
  { key: 'endConsumerName2',      label: 'End Consumer Name 2',    computed: true },
];

// ── YEAR_MONTH display formatter: "2026-04" → "April 2026" ───────────────
const MONTHS = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
function fmtYearMonth(v) {
  if (!v) return '';
  const [y, m] = String(v).split('-');
  const mi = parseInt(m, 10) - 1;
  return (MONTHS[mi] || m) + ' ' + y;
}

// ── Cell formatter ────────────────────────────────────────────────────────
function fmt(v, col) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (col.key === 'yearMonth') return fmtYearMonth(v);
  if (col.dp !== undefined) {
    const n = Number(v);
    return isNaN(n) ? String(v) : n.toFixed(col.dp);
  }
  return String(v);
}

// ── Build rows ────────────────────────────────────────────────────────────
const rows = recs.map(r => COLS.map(c => fmt(r[c.key], c)));

// ── Column picker groups ──────────────────────────────────────────────────
const groups = [
  { label: 'Time',        keys: ['yearMonth','fiscalYear','fiscalMonth','fiscalQtr'] },
  { label: 'Customer',    keys: ['name','customerName','customerNameFull','custNameFullReseller','managementOrgName','endConsumerName','ibmCustNum','ibmClientNum','ibmClientName'] },
  { label: 'Type & Dims', keys: ['type','vendor','region','iot','imt','clientType','signingCategory','signingsSubCat','payingStatus','revenueType'] },
  { label: 'Contract',    keys: ['salesOrderNum','soLi','coverageTerm','contractStart','contractEnd','nextRenewalDate','renewalCounter','reportingDate','signingDate'] },
  { label: 'Product',     keys: ['partNum','partDesc','family','product','offeringStatus'] },
  { label: 'Metrics',     keys: ['acvUsd','tcvUsd','acvLocal','tcvLocal','currencyCode','fmsAmount','partQty','usdCost','usdAdjustedCost','country'] },
  { label: 'Computed ✦',  keys: ['isLatestMonth','tenureMonths','contractType','configuredCSum','committedAmount','isL12Months','endConsumerName2'] },
];

// ── Distinct filter values ────────────────────────────────────────────────
function distinct(key) {
  return [...new Set(recs.map(r => r[key]).filter(Boolean))].sort();
}

// YEAR_MONTH filter — show "April 2026" labels but filter on raw "2026-04" values
const ymRaw    = [...new Set(recs.map(r => r.yearMonth).filter(Boolean))].sort();
const ymLabels = ymRaw.map(v => ({ value: v, label: fmtYearMonth(v) }));

// Parallel array of raw yearMonth values (one per row) for YEAR_MONTH filter
const rawYm = recs.map(r => r.yearMonth || '');

const pageData = {
  snapshotDate:  data.weeklySnapshots[0]?.snapshotDate,
  generatedAt:   data.generatedAt,
  cols:          COLS,
  groups,
  rows,
  rawYm,
  // Filter options
  custNameOptions:     distinct('customerNameFull'),
  vendorOptions:       distinct('vendor'),
  contractTypeOptions: ['Annual', 'Monthly'],
  ymOptions:           ymLabels,
};

// ── HTML ──────────────────────────────────────────────────────────────────
const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>MSP Data Review — ${pageData.snapshotDate}</title>
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#161616;--sur:#262626;--sur2:#393939;--bdr:#525252;--txt:#f4f4f4;--mut:#8d8d8d;
      --blu:#0f62fe;--grn:#42be65;--hi-bg:rgba(66,190,101,.07)}
html,body{height:100%;background:var(--bg);color:var(--txt);
  font-family:'IBM Plex Sans','Segoe UI',system-ui,sans-serif;font-size:13px}

/* top bar — title + actions only */
#toolbar{position:sticky;top:0;z-index:200;background:#0f0f0f;
  border-bottom:1px solid var(--bdr);padding:6px 16px;
  display:flex;align-items:center;gap:10px;flex-wrap:wrap;min-height:44px}
#toolbar h1{font-size:14px;font-weight:600;white-space:nowrap}
.ibm{font-size:10px;background:var(--blu);color:#fff;padding:2px 6px;border-radius:2px;flex-shrink:0}
.tb-sep{width:1px;height:22px;background:var(--bdr);flex-shrink:0}
.tb-btn{background:var(--sur);color:var(--txt);border:1px solid var(--bdr);
  border-radius:2px;padding:4px 10px;font-size:12px;cursor:pointer;white-space:nowrap;flex-shrink:0}
.tb-btn:hover{background:var(--sur2)}
.tb-btn.export{border-color:#4589ff;color:#4589ff}
.tb-btn.export:hover{background:#0d2b5c}
#info{font-size:11px;color:var(--mut);white-space:nowrap}
.legend{display:flex;align-items:center;gap:4px;font-size:11px;color:var(--mut)}
.leg-dot{width:9px;height:9px;background:var(--hi-bg);border:1px solid #42be65;border-radius:2px}

/* filter bar — below toolbar */
#filter-bar{background:var(--sur);border-bottom:1px solid var(--bdr);
  padding:8px 16px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.fb-ctrl{display:flex;flex-direction:column;gap:3px}
.fb-ctrl label{font-size:10px;color:var(--mut);font-weight:500;white-space:nowrap}
.fb-ctrl select{background:var(--sur2);color:var(--txt);border:1px solid var(--bdr);
  border-radius:2px;padding:4px 8px;font-size:12px;cursor:pointer}
.fb-ctrl select:focus{outline:2px solid var(--blu);outline-offset:1px}
#f-cust{min-width:200px;max-width:260px}
#f-ym{min-width:140px}
#f-vendor{min-width:120px}
#f-contracttype{min-width:110px}
#refresh-badge{margin-left:auto;font-size:11px;color:var(--mut);white-space:nowrap;
  border-left:1px solid var(--bdr);padding-left:12px}
#refresh-badge span{color:var(--txt);font-weight:500}

/* column picker */
#col-panel{position:fixed;top:88px;left:16px;z-index:300;background:var(--sur);
  border:1px solid var(--bdr);border-radius:4px;padding:12px 14px;width:260px;
  max-height:72vh;overflow-y:auto;display:none;box-shadow:0 4px 20px rgba(0,0,0,.6)}
#col-panel h3{font-size:12px;font-weight:600;margin-bottom:8px}
.cp-grp{font-size:10px;font-weight:600;color:var(--mut);text-transform:uppercase;
  letter-spacing:.05em;margin:8px 0 2px;padding-top:6px;border-top:1px solid var(--bdr)}
.cp-grp:first-of-type{border-top:none;margin-top:0;padding-top:0}
#col-panel label{display:flex;align-items:center;gap:6px;padding:2px 0;font-size:12px;cursor:pointer}
#col-panel label input{cursor:pointer;accent-color:var(--blu)}

/* table — offset for both bars */
#wrap{overflow:auto;height:calc(100vh - 44px - 44px)}
table{border-collapse:collapse;font-size:12px;white-space:nowrap}
thead th{position:sticky;top:0;z-index:10;background:var(--sur2);color:var(--mut);
  font-weight:500;text-transform:uppercase;font-size:10px;letter-spacing:.04em;
  padding:7px 10px;border-right:1px solid var(--bdr);border-bottom:2px solid var(--bdr);
  cursor:pointer;user-select:none}
thead th:hover{color:var(--txt)}
thead th.asc::after{content:' ↑'}thead th.desc::after{content:' ↓'}
thead th.computed{background:#1a2e1a;color:#6fdc8c}
thead th.computed:hover{color:var(--grn)}
tbody tr{border-bottom:1px solid var(--sur2)}
tbody tr:hover{background:var(--sur2)}
tbody td{padding:6px 10px;border-right:1px solid var(--sur2);
  max-width:300px;overflow:hidden;text-overflow:ellipsis;vertical-align:top}
tbody td.computed{background:var(--hi-bg)}
tbody td.bool-t{color:var(--grn);font-weight:600}
tbody td.bool-f{color:var(--mut)}
tbody td.num{text-align:right;font-variant-numeric:tabular-nums}
</style>
</head>
<body>

<!-- top bar: title + action controls only -->
<div id="toolbar">
  <h1>MSP Data Review</h1><span class="ibm">IBM</span>
  <span class="tb-sep"></span>
  <button class="tb-btn" id="col-toggle-btn">⚙ Columns</button>
  <span class="legend"><span class="leg-dot"></span>Computed</span>
  <span style="flex:1"></span>
  <button class="tb-btn export" id="export-btn">⬇ Export CSV</button>
  <span id="info"></span>
</div>

<!-- filter bar: all data filters + refresh date -->
<div id="filter-bar">
  <div class="fb-ctrl">
    <label>CUSTOMER</label>
    <select id="f-cust"><option value="">All Customers</option></select>
  </div>

  <div class="fb-ctrl">
    <label>MONTH</label>
    <select id="f-ym"><option value="">All Months</option></select>
  </div>

  <div class="fb-ctrl">
    <label>VENDOR</label>
    <select id="f-vendor"><option value="">All Vendors</option></select>
  </div>

  <div class="fb-ctrl">
    <label>CONTRACT TYPE</label>
    <select id="f-contracttype"><option value="">All Types</option></select>
  </div>

  <div id="refresh-badge">Data refreshed weekly &nbsp;·&nbsp; <span>${pageData.snapshotDate}</span></div>
</div>

<div id="col-panel">
  <h3>Visible columns</h3>
  <div id="col-all-row" style="display:flex;gap:8px;margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid var(--bdr)">
    <button id="col-all-btn"  style="flex:1;background:var(--sur2);color:var(--txt);border:1px solid var(--bdr);border-radius:2px;padding:3px 0;font-size:11px;cursor:pointer">Select All</button>
    <button id="col-none-btn" style="flex:1;background:var(--sur2);color:var(--txt);border:1px solid var(--bdr);border-radius:2px;padding:3px 0;font-size:11px;cursor:pointer">Select None</button>
  </div>
  <div id="col-checks"></div>
</div>

<div id="wrap">
  <table id="tbl">
    <thead id="thead"></thead>
    <tbody id="tbody"></tbody>
  </table>
</div>

<script>
const PD = ${JSON.stringify(pageData)};

const NUM_KEYS = new Set(['fiscalYear','fiscalMonth','acvUsd','tcvUsd','acvLocal','tcvLocal',
  'fmsAmount','partQty','usdCost','usdAdjustedCost','configuredCSum','committedAmount',
  'coverageTerm','renewalCounter','tenureMonths','salesOrderNum']);

// ── Column indices ────────────────────────────────────────────────────────
const CI = {};
PD.cols.forEach((c,i) => CI[c.key] = i);

// ── State ─────────────────────────────────────────────────────────────────
const colVis = {};
PD.cols.forEach(c => colVis[c.key] = true);
let sortCol = -1, sortDir = 1;
let F = { cust: '', ym: '', vendor: '', contractType: '' };
let filteredRows = PD.rows;

// ── Populate filter dropdowns ─────────────────────────────────────────────
function fillSelect(id, items) {
  const el = document.getElementById(id);
  if (Array.isArray(items) && typeof items[0] === 'string') {
    items.forEach(v => el.appendChild(new Option(v, v)));
  } else {
    // [{value, label}]
    items.forEach(o => el.appendChild(new Option(o.label, o.value)));
  }
}
fillSelect('f-cust',         PD.custNameOptions);
fillSelect('f-ym',           PD.ymOptions);
fillSelect('f-vendor',       PD.vendorOptions);
fillSelect('f-contracttype', PD.contractTypeOptions);

// ── Filter ────────────────────────────────────────────────────────────────
// YEAR_MONTH: rows[] stores "April 2026" (display); PD.rawYm[i] stores "2026-04" (raw).
const FCI = {
  cust:         CI['customerNameFull'],
  vendor:       CI['vendor'],
  contractType: CI['contractType'],
};

function applyFilter() {
  filteredRows = PD.rows.filter((row, i) => {
    if (F.cust         && row[FCI.cust]        !== F.cust)         return false;
    if (F.ym           && PD.rawYm[i]          !== F.ym)           return false;
    if (F.vendor       && row[FCI.vendor]       !== F.vendor)       return false;
    if (F.contractType && row[FCI.contractType] !== F.contractType) return false;
    return true;
  });
}

['f-cust','f-ym','f-vendor','f-contracttype'].forEach(id => {
  document.getElementById(id).addEventListener('change', e => {
    const key = { 'f-cust':'cust','f-ym':'ym','f-vendor':'vendor','f-contracttype':'contractType' }[id];
    F[key] = e.target.value;
    applyFilter();
    if (sortCol >= 0) applySort();
    rebuildTable();
  });
});

// ── Sort ─────────────────────────────────────────────────────────────────
function applySort() {
  if (sortCol < 0) return;
  const ci = sortCol;
  filteredRows.sort((a, b) => {
    const av = a[ci], bv = b[ci];
    const an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) return (an - bn) * sortDir;
    return String(av).localeCompare(String(bv)) * sortDir;
  });
}

// ── Column picker ─────────────────────────────────────────────────────────
const colPanel = document.getElementById('col-panel');
document.getElementById('col-toggle-btn').addEventListener('click', () => {
  colPanel.style.display = colPanel.style.display === 'block' ? 'none' : 'block';
});
document.addEventListener('click', e => {
  if (!colPanel.contains(e.target) && e.target.id !== 'col-toggle-btn')
    colPanel.style.display = 'none';
});
const keyToCol = {};
PD.cols.forEach(c => keyToCol[c.key] = c);
const checks = document.getElementById('col-checks');
PD.groups.forEach(g => {
  const hdr = document.createElement('div');
  hdr.className = 'cp-grp'; hdr.textContent = g.label;
  checks.appendChild(hdr);
  g.keys.forEach(k => {
    const col = keyToCol[k]; if (!col) return;
    const lbl = document.createElement('label');
    const cb  = document.createElement('input');
    cb.type = 'checkbox'; cb.checked = true; cb.dataset.key = k;
    cb.addEventListener('change', () => { colVis[k] = cb.checked; rebuildTable(); });
    lbl.appendChild(cb);
    lbl.appendChild(document.createTextNode(' ' + col.label));
    checks.appendChild(lbl);
  });
});

// ── Select All / None ─────────────────────────────────────────────────────
document.getElementById('col-all-btn').addEventListener('click', () => {
  PD.cols.forEach(c => { colVis[c.key] = true; });
  document.querySelectorAll('#col-checks input[type=checkbox]').forEach(cb => { cb.checked = true; });
  rebuildTable();
});
document.getElementById('col-none-btn').addEventListener('click', () => {
  PD.cols.forEach(c => { colVis[c.key] = false; });
  document.querySelectorAll('#col-checks input[type=checkbox]').forEach(cb => { cb.checked = false; });
  rebuildTable();
});

// ── Render ────────────────────────────────────────────────────────────────
function renderHeader() {
  const tr = document.createElement('tr');
  PD.cols.forEach((c, ci) => {
    if (!colVis[c.key]) return;
    const th = document.createElement('th');
    th.textContent = c.label;
    th.dataset.ci = ci;
    if (c.computed) th.classList.add('computed');
    if (ci === sortCol) th.classList.add(sortDir === 1 ? 'asc' : 'desc');
    th.addEventListener('click', () => {
      if (sortCol === ci) sortDir *= -1; else { sortCol = ci; sortDir = 1; }
      applySort(); rebuildTable();
    });
    tr.appendChild(th);
  });
  const thead = document.getElementById('thead');
  thead.innerHTML = ''; thead.appendChild(tr);
}

function renderBody() {
  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';
  const frag = document.createDocumentFragment();
  filteredRows.forEach(row => {
    const tr = document.createElement('tr');
    PD.cols.forEach((c, ci) => {
      if (!colVis[c.key]) return;
      const td = document.createElement('td');
      const val = row[ci];
      td.textContent = val;
      if (c.computed)            td.classList.add('computed');
      if (val === 'TRUE')        td.classList.add('bool-t');
      if (val === 'FALSE')       td.classList.add('bool-f');
      if (NUM_KEYS.has(c.key) && val !== '') td.classList.add('num');
      tr.appendChild(td);
    });
    frag.appendChild(tr);
  });
  tbody.appendChild(frag);
}

function updateInfo() {
  const vis = PD.cols.filter(c => colVis[c.key]).length;
  document.getElementById('info').textContent =
    filteredRows.length.toLocaleString() + ' of ' + PD.rows.length.toLocaleString() +
    ' rows · ' + vis + ' cols · ' + PD.snapshotDate;
}

function rebuildTable() { renderHeader(); renderBody(); updateInfo(); }

// ── Export CSV ────────────────────────────────────────────────────────────
document.getElementById('export-btn').addEventListener('click', () => {
  const visCols = PD.cols.filter(c => colVis[c.key]);
  const lines   = [];

  // Header row
  lines.push(visCols.map(c => '"' + c.label.replace(/"/g,'""') + '"').join(','));

  // Data rows (only filtered rows, only visible columns)
  filteredRows.forEach(row => {
    lines.push(visCols.map(c => {
      const v = row[CI[c.key]];
      if (v === null || v === undefined || v === '') return '';
      const s = String(v).replace(/"/g,'""');
      return s.includes(',') || s.includes('"') || s.includes('\\n') ? '"'+s+'"' : s;
    }).join(','));
  });

  const blob = new Blob([lines.join('\\n')], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = 'MSP_Data_Review_' + PD.snapshotDate + '.csv';
  a.click();
  URL.revokeObjectURL(url);
});

// ── Init ──────────────────────────────────────────────────────────────────
rebuildTable();
</script>
</body>
</html>`;

fs.writeFileSync(outputPath, html, 'utf8');
const kb = (fs.statSync(outputPath).size / 1024).toFixed(0);
console.log(`Data review written: ${outputPath}`);
console.log(`Records: ${recs.length} | Columns: ${COLS.length} (7 computed) | Size: ${kb} KB`);
