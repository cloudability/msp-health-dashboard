/**
 * MSP Dashboard — weekly update (works on macOS and Windows)
 *
 * Usage:  node build/update.js [--no-open]     (or: npm run update)
 *
 * Steps:
 *   1. Keep only the NEWEST weekly export in "Source files/"; move older ones to "Source files/History/"
 *      (the dashboard reads a single snapshot, so older files at the top level would be embedded but ignored)
 *   2. Back up the previous dashboard to Final/Backup/ with a timestamp, then prune old backups
 *   3. Build Final/MSP-CCB-Performance-Dashboard.html   (build/parse-data.js)
 *   4. Build data-review.html                           (build/make-data-review.js)
 *   5. Open the dashboard in the default browser (skip with --no-open)
 */

'use strict';
const fs   = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { extractDate, ACTIVE_VIEW_DIR } = require('./parse-data');

const ROOT       = path.join(__dirname, '..');
const SOURCE_DIR = path.join(ROOT, 'Source files');
const HISTORY    = path.join(SOURCE_DIR, 'History');
const FINAL      = path.join(ROOT, 'Final', 'MSP-CCB-Performance-Dashboard.html');
const BACKUP_DIR = path.join(ROOT, 'Final', 'Backup');
const MAX_DAYS   = 10;   // keep backups for the newest N distinct days
const PER_DAY    = 2;    // keep the newest N backups per day

const log = (m = '') => console.log(m);
const fail = m => { console.error('\nERROR: ' + m); process.exit(1); };

// ── 1. Archive older weekly exports ─────────────────────────────────────────
function archiveOldExports() {
  if (!fs.existsSync(SOURCE_DIR)) fail(`Folder not found: ${SOURCE_DIR}`);
  const files = fs.readdirSync(SOURCE_DIR)
    .filter(f => /\.(xlsx|xls|csv)$/i.test(f) && !/glossary/i.test(f) && !f.startsWith('~$'))
    .map(f => {
      const full = path.join(SOURCE_DIR, f);
      return { name: f, full, date: extractDate(f, fs.statSync(full).mtime), mtime: fs.statSync(full).mtimeMs };
    })
    // newest first; same date -> most recently modified file wins
    .sort((a, b) => b.date.localeCompare(a.date) || b.mtime - a.mtime);

  if (!files.length) fail('No data export found in "Source files/". Drop the latest MSP Reporting .csv or .xlsx there.');

  log(`Latest export: ${files[0].name}  (snapshot ${files[0].date})`);
  const older = files.slice(1);
  if (!older.length) return;

  fs.mkdirSync(HISTORY, { recursive: true });
  for (const f of older) {
    let target = path.join(HISTORY, f.name);
    if (fs.existsSync(target)) {
      const ext = path.extname(f.name);
      target = path.join(HISTORY, `${path.basename(f.name, ext)} (${Date.now()})${ext}`);
    }
    fs.renameSync(f.full, target);
    log(`  Moved older export to History/: ${f.name}`);
  }
}

// ── 1b. Archive older active-orgs view exports (Source files/Active Orgs/*.csv) ──
function archiveOldViewExports() {
  const dir = path.join(SOURCE_DIR, ACTIVE_VIEW_DIR);
  if (!fs.existsSync(dir)) {
    log('No "Source files/Active Orgs/" folder - overlapping contracts will not be filtered.');
    return;
  }
  const files = fs.readdirSync(dir)
    .filter(f => /\.csv$/i.test(f) && !f.startsWith('~$'))
    .map(f => {
      const full = path.join(dir, f);
      const st = fs.statSync(full);
      const m = f.match(/(\d{2})[_-](\d{2})[_-](\d{2})(?!\d)/);
      return { name: f, full, date: m ? `20${m[3]}-${m[2]}-${m[1]}` : new Date(st.mtimeMs).toISOString().slice(0, 10), mtime: st.mtimeMs };
    })
    .sort((a, b) => b.date.localeCompare(a.date) || b.mtime - a.mtime);

  if (!files.length) { log('No active-orgs view CSV found - overlapping contracts will not be filtered.'); return; }
  log(`Latest active-orgs view: ${files[0].name}  (view date ${files[0].date})`);
  const older = files.slice(1);
  if (!older.length) return;
  const hist = path.join(dir, 'History');
  fs.mkdirSync(hist, { recursive: true });
  for (const f of older) {
    let target = path.join(hist, f.name);
    if (fs.existsSync(target)) {
      const ext = path.extname(f.name);
      target = path.join(hist, `${path.basename(f.name, ext)} (${Date.now()})${ext}`);
    }
    fs.renameSync(f.full, target);
    log(`  Moved older view export to Active Orgs/History/: ${f.name}`);
  }
}

// ── 2. Timestamped backup + retention ───────────────────────────────────────
function backupAndPrune() {
  if (!fs.existsSync(FINAL)) return;
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  const day = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  const ts  = `${day}_${p(d.getHours())}-${p(d.getMinutes())}-${p(d.getSeconds())}`;
  const name = `MSP-CCB-Performance-Dashboard_${ts}.html`;
  fs.copyFileSync(FINAL, path.join(BACKUP_DIR, name));
  log(`Backed up previous dashboard: Final/Backup/${name}`);

  // Only touches files named MSP-CCB-Performance-Dashboard_yyyy-MM-dd_HH-mm-ss.html
  const re = /^MSP-CCB-Performance-Dashboard_(\d{4}-\d{2}-\d{2})_\d{2}-\d{2}-\d{2}\.html$/;
  const backups = fs.readdirSync(BACKUP_DIR).map(f => ({ f, m: f.match(re) })).filter(x => x.m)
    .map(x => ({ name: x.f, day: x.m[1] }));
  const keepDays = [...new Set(backups.map(b => b.day))].sort().reverse().slice(0, MAX_DAYS);
  const keep = new Set();
  keepDays.forEach(dy => backups.filter(b => b.day === dy).map(b => b.name).sort().reverse()
    .slice(0, PER_DAY).forEach(n => keep.add(n)));
  let removed = 0;
  backups.forEach(b => { if (!keep.has(b.name)) { fs.unlinkSync(path.join(BACKUP_DIR, b.name)); removed++; } });
  log(`Backup cleanup: kept ${keep.size}, removed ${removed}.`);
}

// ── 3 & 4. Run the build scripts ────────────────────────────────────────────
function run(label, script) {
  log(`\n${label}`);
  const r = spawnSync(process.execPath, [path.join(__dirname, script)], { cwd: ROOT, stdio: 'inherit' });
  if (r.status !== 0) fail(`${script} failed (see output above).`);
}

// ── 5. Open in default browser ──────────────────────────────────────────────
function openInBrowser(file) {
  const cmd  = process.platform === 'darwin' ? ['open', [file]]
             : process.platform === 'win32'  ? ['cmd', ['/c', 'start', '""', file]]
             :                                 ['xdg-open', [file]];
  spawnSync(cmd[0], cmd[1], { stdio: 'ignore', windowsVerbatimArguments: process.platform === 'win32' });
}

log('=== MSP / CCB Performance Dashboard - Update ===\n');
archiveOldExports();
archiveOldViewExports();
backupAndPrune();
run('[1/2] Building MSP-CCB-Performance-Dashboard.html ...', 'parse-data.js');
run('[2/2] Building data-review.html ...', 'make-data-review.js');
log('\nDone.');
if (!process.argv.includes('--no-open')) {
  log('Opening dashboard in your browser ...');
  openInBrowser(FINAL);
}
