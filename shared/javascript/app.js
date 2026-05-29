'use strict';

// ============================================================
// STATE
// ============================================================
const state = {
  files: {}, // { 'appSec-base': ArrayBuffer, 'appSec-tool': ArrayBuffer, ... }
  results: {}, // processed output per section
  filesCount: 0,
  renderedTables: {}, // tableId -> true (lazy year-tab rendering)
};

const HEADER_SCAN_MAX_ROWS = 120;
const TABLE_PREVIEW_ROWS = 75;

// Section configs
const SECTIONS = {
  appSec: {
    label: 'Application Security',
    trainingName: 'Introduction to Application Security #ABI',
    baseType: 'appSec',
    baseIdCol: 'Local Employee ID',
    baseEmailCol: 'Employee Email',
    toolHeaderRow: 21,
    phishingFilter: null,
    outputFile: 'AppSec_Output.xlsx',
  },
  cyberOT: {
    label: 'Cyber OT',
    trainingName: 'Cyber OT Training #ABI',
    baseType: 'cyberOT',
    baseIdCol: 'Local Employee ID',
    baseEmailCol: 'Employee Email',
    toolHeaderRow: 21,
    phishingFilter: null,
    outputFile: 'CyberOT_Output.xlsx',
  },
  growth: {
    label: 'Growth Group',
    trainingName: 'Mandatory Anti-Phishing Training — Growth Group #ABI',
    baseType: 'growth',
    baseIdCol: 'Local Employee ID',
    baseEmailCol: 'Employee Email',
    toolHeaderRow: 21,
    phishingFilter: null,
    outputFile: 'Growth_Output.xlsx',
  },
  newJoiner: {
    label: 'New Joiner',
    trainingName: 'IT Security Awareness - English #ABI',
    baseType: 'newJoiner',
    baseIdCol: null, // dynamic
    baseEmailCol: null,
    toolHeaderRow: 21,
    phishingFilter: null,
    outputFile: 'NewJoiner_Output.xlsx',
  },
  bsc: {
    label: 'BSC',
    trainingName: "Don't take the bait — Avoid Financial Impact of Phishing #ABI",
    baseType: 'bsc',
    baseIdCol: 'Emp ID',
    baseEmailCol: 'Email - Primary Work',
    toolHeaderRow: 21,
    phishingFilter: null,
    outputFile: 'BSC_Output.xlsx',
  },
  phishingNormal: {
    label: 'Phishing Normal Users',
    trainingName: 'Security & Compliance Awareness : Phishing #ABI',
    baseType: 'phishingNormal',
    baseIdCol: 'Local Employee ID',
    baseEmailCol: 'Employee Email',
    toolHeaderRow: 21,
    phishingFilter: null, // all non-band4+ users (or all)
    outputFile: 'Phishing_Normal_Output.xlsx',
  },
  band4: {
    label: 'Band 4+ Senior Management',
    trainingName: 'Security & Compliance Awareness Training (Sr. Management) #ABI',
    baseType: 'band4',
    baseIdCol: 'Local Employee ID',
    baseEmailCol: 'Employee Email',
    toolHeaderRow: 21,
    phishingFilter: 'band4plus', // filter Band 4+ = Yes
    outputFile: 'Band4Plus_Output.xlsx',
  },
};

const OUTPUT_COLS = ['Employee ID', 'Work Email Address', 'Zone', 'Transcript Status', 'Training Start Date', 'Transcript Completed Date'];

const REQUIRED_TRAINING_HEADER_SET = new Set([
  'employee id', 'transcript status', 'work email address',
]);

const MACRO_ZONE_MAP = {
  GLOBAL: 'Global Core',
  'ZONE MIDDLE AMERICAS': 'MAZ',
  'ZONE EUROPE': 'EUR',
  'ZONE ASIA PACIFIC': 'APAC',
  'ZONE AFRICA': 'AFR',
  'ZONE SOUTH AMERICA': 'SAZ',
  'ZONE NORTH AMERICA': 'NAZ',
};

const ZONE_DISPLAY_ORDER = ['MAZ', 'EUR', 'APAC', 'AFR', 'SAZ', 'NAZ', 'Global Core', 'Unknown'];

function mapMacroZone(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';
  const mapped = MACRO_ZONE_MAP[s.toUpperCase()];
  return mapped || s;
}

function rowZone(row) {
  return mapMacroZone(row['Zone'] || row['Macro Entity Level 2 (Zone)'] || '');
}

function findColumn(headers, exactNames, ...patternGroups) {
  for (const name of exactNames) {
    const hit = headers.find(h => norm(h) === norm(name));
    if (hit) return hit;
  }
  for (const patterns of patternGroups) {
    if (!patterns) continue;
    const list = Array.isArray(patterns) ? patterns : [patterns];
    for (const pat of list) {
      const hit = headers.find(h => pat.test(String(h || '')));
      if (hit) return hit;
    }
  }
  return null;
}

function resolveToolColumns(headers) {
  return {
    id: findColumn(headers, ['Employee ID'], /^employee\s*id$/i, /employee.?id|emp.?id/i),
    email: findColumn(headers, ['Work Email Address'], /^work\s*email/i, /work.?email|email.?address/i),
    status: findColumn(headers, ['Transcript Status'], /transcript.?status/i),
    start: findColumn(headers, ['Training Start Date'], /training.?start/i),
    complete: findColumn(
      headers,
      ['Transcript Completed Date', 'Transcript Completion Date', 'Completion Date'],
      /transcript.?completed|completion.?date/i
    ),
    zone: findColumn(
      headers,
      ['Macro Entity Level 2 (Zone)', 'Zone'],
      /macro.?entity.?level.?2/i,
      /^zone$/i
    ),
  };
}

function cellVal(row, col) {
  if (!col || !row) return '';
  const v = row[col];
  if (v === null || v === undefined) return '';
  return v;
}

function makeOutputRow({ empId, email, zone, status, startDate, completedDate }) {
  return {
    'Employee ID': empId != null && empId !== '' ? String(empId).trim() : '',
    'Work Email Address': email != null && email !== '' ? String(email).trim() : '',
    'Zone': mapMacroZone(zone),
    'Transcript Status': status || '',
    'Training Start Date': startDate || '',
    'Transcript Completed Date': completedDate || '',
  };
}

function deriveTranscriptStatus(rawStatus, completedDate) {
  const completed = completedDate && String(completedDate).trim() !== '';
  const raw = String(rawStatus || '').trim();
  if (completed) return 'Completed';
  if (/in progress/i.test(raw)) return 'In Progress';
  if (raw) return raw;
  return 'Not Started';
}

// ============================================================
// CLOCK
// ============================================================
function updateClock() {
  const now = new Date();
  document.getElementById('clock').textContent = now.toLocaleTimeString('en-GB');
}
setInterval(updateClock, 1000);
updateClock();

// Set default dates to today
(function setDefaultDates() {
  const today = new Date().toISOString().split('T')[0];
  ['appSec','cyberOT','growth','newJoiner','bsc','phishingNormal','band4'].forEach(s => {
    const el = document.getElementById('date-'+s);
    if (el) el.value = today;
  });
})();

// ============================================================
// TAB SWITCHING
// ============================================================
function switchTab(sectionId) {
  document.querySelectorAll('.section-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('panel-'+sectionId).classList.add('active');
  event.target.classList.add('active');
}

function switchNJTab(year) {
  document.querySelectorAll('#panel-newJoiner .sheet-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('#panel-newJoiner .sheet-content').forEach(c => c.classList.remove('active'));
  event.target.classList.add('active');
  document.getElementById('nj-sheet-'+year).classList.add('active');
  const result = state.results.newJoiner;
  if (result) renderTable('newJoiner-' + year, result[year] || []);
}

function switchYearTab(section, year) {
  document.querySelectorAll('#panel-'+section+' .sheet-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('#panel-'+section+' .sheet-content').forEach(c => c.classList.remove('active'));
  event.target.classList.add('active');
  document.getElementById(section+'-sheet-'+year).classList.add('active');
  const result = state.results[section];
  if (result) renderTable(section + '-' + year, result[year] || []);
}

/**
 * Split a flat array of rows into { '2024': [], '2025': [], '2026': [] }
 * using the Training Start Date field to determine the year.
 */
function splitRowsByYear(rows) {
  const result = { '2024': [], '2025': [], '2026': [] };
  for (const row of rows) {
    const sd = row['Training Start Date'];
    let year = new Date().getFullYear().toString();
    if (sd) {
      // SD is already formatted as DD/MM/YYYY string
      const parts = String(sd).split('/');
      if (parts.length === 3 && parts[2]) {
        year = parts[2];
      } else {
        const d = new Date(sd);
        if (!isNaN(d.getTime())) year = String(d.getFullYear());
      }
    }
    if (result[year]) result[year].push(row);
    else result['2026'].push(row);
  }
  return result;
}

// ============================================================
// FILE UPLOAD
// ============================================================
function dragOver(e, el) {
  e.preventDefault();
  el.classList.add('drag-over');
}
function dragLeave(el) {
  el.classList.remove('drag-over');
}
function dropFile(e, section, type) {
  e.preventDefault();
  const el = document.getElementById('uz-'+section+'-'+type);
  el.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) processFileUpload(file, section, type);
}
function fileSelected(e, section, type) {
  const file = e.target.files[0];
  if (file) processFileUpload(file, section, type);
}

function processFileUpload(file, section, type) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    toast('error', 'Please upload an Excel file (.xlsx or .xls)');
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    const key = section + '-' + type;
    state.files[key] = e.target.result;
    state.filesCount++;
    document.getElementById('kpi-files').textContent = state.filesCount;
    document.getElementById('fn-'+section+'-'+type).textContent = '✓ ' + file.name;
    const uz = document.getElementById('uz-'+section+'-'+type);
    uz.classList.add('has-file');
    toast('info', `Loaded: ${file.name}`);
  };
  reader.readAsArrayBuffer(file);
}

// ============================================================
// EXCEL PARSING UTILITIES
// ============================================================

/**
 * Parse an ArrayBuffer to XLSX workbook
 */
function parseWorkbook(ab) {
  return XLSX.read(new Uint8Array(ab), { type: 'array', cellDates: true });
}

function yieldToMain() {
  return new Promise(resolve => setTimeout(resolve, 0));
}

async function runHeavyWork(fn) {
  await yieldToMain();
  return fn();
}

function rowHasRequiredTrainingHeaders(rowVals) {
  const cells = new Set();
  for (let i = 0; i < rowVals.length; i++) {
    const t = String(rowVals[i] || '').trim().toLowerCase();
    if (t) cells.add(t);
  }
  for (const req of REQUIRED_TRAINING_HEADER_SET) {
    if (!cells.has(req)) return false;
  }
  return true;
}

function normalizeHeaderRow(rowVals) {
  const seen = {};
  return rowVals.map((h, i) => {
    let name = String(h ?? '').trim();
    if (!name) name = '__col_' + i;
    const key = name.toLowerCase();
    seen[key] = (seen[key] || 0) + 1;
    if (seen[key] > 1) name = name + ' (' + seen[key] + ')';
    return name;
  });
}

function isRowEmpty(row) {
  for (const k in row) {
    if (!Object.prototype.hasOwnProperty.call(row, k)) continue;
    const v = row[k];
    if (v !== '' && v != null && v !== undefined) return false;
  }
  return true;
}

/**
 * Fast header detection (first 120 rows only) + bulk row parse via sheet_to_json.
 * Returns { headers, headerRowIndex, rows }
 */
function detectHeaders(sheet, minCols = 3) {
  const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
  if (range.e.r < range.s.r) return { headers: [], headerRowIndex: 0, rows: [] };

  const scanEnd = Math.min(range.e.r, range.s.r + HEADER_SCAN_MAX_ROWS - 1);
  const scanRef = XLSX.utils.encode_range({
    s: { r: range.s.r, c: range.s.c },
    e: { r: scanEnd, c: range.e.c },
  });
  const scanAoA = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    range: scanRef,
    defval: '',
    blankrows: false,
  });

  let headerRowIndex = range.s.r;
  let headers = [];

  for (let i = 0; i < scanAoA.length; i++) {
    const rowVals = scanAoA[i].map(v => String(v ?? '').trim());
    if (rowHasRequiredTrainingHeaders(rowVals)) {
      headers = normalizeHeaderRow(rowVals);
      headerRowIndex = range.s.r + i;
      break;
    }
  }

  if (!headers.length) {
    for (let i = 0; i < scanAoA.length; i++) {
      const rowVals = scanAoA[i].map(v => String(v ?? '').trim());
      const nonEmpty = rowVals.filter(v => v !== '').length;
      if (nonEmpty >= minCols) {
        headers = normalizeHeaderRow(rowVals);
        headerRowIndex = range.s.r + i;
        break;
      }
    }
  }

  const dataRef = XLSX.utils.encode_range({
    s: { r: headerRowIndex, c: range.s.c },
    e: { r: range.e.r, c: range.e.c },
  });
  let rows = XLSX.utils.sheet_to_json(sheet, {
    range: dataRef,
    defval: '',
    blankrows: false,
    raw: false,
  });
  rows = rows.filter(r => !isRowEmpty(r));

  if (!headers.length && rows.length) {
    headers = normalizeHeaderRow(Object.keys(rows[0]));
  }

  return { headers, headerRowIndex, rows };
}

/**
 * Normalize a string for comparison (lowercase, trim)
 */
function norm(v) {
  if (v === null || v === undefined) return '';
  return String(v).toLowerCase().trim();
}

/**
 * Normalize employee ID
 */
function normId(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/\s/g, '').toLowerCase();
}

/**
 * Format a date value from XLSX
 */
function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    return v.toLocaleDateString('en-GB'); // DD/MM/YYYY
  }
  // Try parsing
  const d = new Date(v);
  if (!isNaN(d.getTime())) return d.toLocaleDateString('en-GB');
  return String(v);
}

/**
 * Build a lookup map from an array of rows
 * Keys: email (lower) and empId
 * Returns Map<key, row>
 */
function buildLookupMap(rows, emailCol, idCol) {
  const map = new Map();
  rows.forEach(row => {
    if (emailCol) {
      const e = norm(row[emailCol]);
      if (e) map.set(e, row);
    }
    if (idCol) {
      const id = normId(row[idCol]);
      if (id) map.set(id, row);
    }
  });
  return map;
}

/**
 * Find a match for a user in a lookup map
 */
function findMatch(map, email, empId) {
  if (email) {
    const r = map.get(norm(email));
    if (r) return r;
  }
  if (empId) {
    const r = map.get(normId(String(empId)));
    if (r) return r;
  }
  return null;
}

/**
 * Tool-only processor (AppSec, Growth):
 * Reads only the tool output file — no base file needed.
 * Deduplicates by email/ID, determines status from tool data.
 */
function processToolOnly(toolAB, assignDate) {
  const toolWb = parseWorkbook(toolAB);
  const toolSheet = toolWb.Sheets[toolWb.SheetNames[0]];
  const toolData = detectHeaders(toolSheet);
  const cols = resolveToolColumns(toolData.headers);

  const output = [];
  const seen = new Set();

  for (const toolRow of toolData.rows) {
    const empId = cellVal(toolRow, cols.id);
    const email = cellVal(toolRow, cols.email);

    const dedupeKey = normId(String(empId || '')) || norm(email);
    if (dedupeKey && seen.has(dedupeKey)) continue;
    if (dedupeKey) seen.add(dedupeKey);

    const rawStatus = String(cellVal(toolRow, cols.status) || '').trim();
    const completedDate = fmtDate(cellVal(toolRow, cols.complete));
    const startDate = fmtDate(cellVal(toolRow, cols.start)) ||
      (norm(rawStatus) !== 'completed' && assignDate ? assignDate : '');
    const zone = cellVal(toolRow, cols.zone);
    const status = deriveTranscriptStatus(rawStatus, completedDate);

    output.push(makeOutputRow({
      empId, email, zone, status, startDate, completedDate,
    }));
  }

  return output;
}

function processAppSec(baseAB, toolAB, assignDate) {
  return processToolOnly(toolAB, assignDate);
}

// ============================================================
// CORE PROCESSING LOGIC
// ============================================================

/**
 * Generic section processor:
 * - Reads base userbase
 * - Reads tool output (finds headers dynamically)
 * - Matches users by email/ID
 * - Returns processed rows with required output columns
 */

/**
 * Phishing Normal processor.
 * Master list = Base File filtered sequentially:
 *   1. BSC === 'no'
 *   2. Band 4+ === 'no'
 * For each master-list user, cross-reference Tool File by email/ID.
 * If found → pull Transcript Status, Training Start Date, Transcript Completed Date.
 * If not found → status = 'Not Started', start date = assignDate.
 * Deduplication by email or Emp ID.
 */
function processPhishingNormal(baseAB, toolAB, assignDate) {
  const cfg = SECTIONS['phishingNormal'];

  // --- Parse Base ---
  const baseWb = parseWorkbook(baseAB);
  const baseSheet = baseWb.Sheets[baseWb.SheetNames[0]];
  const baseData = detectHeaders(baseSheet);

  let baseIdCol = cfg.baseIdCol;
  let baseEmailCol = cfg.baseEmailCol;
  if (!baseIdCol || !baseData.headers.includes(baseIdCol)) {
    baseIdCol = baseData.headers.find(h => /employee.?id|emp.?id|local.?id/i.test(h)) || null;
  }
  if (!baseEmailCol || !baseData.headers.includes(baseEmailCol)) {
    baseEmailCol = baseData.headers.find(h => /email/i.test(h)) || null;
  }

  // Sequential filter: BSC = No → then Band 4+ = No
  const bscFiltered = baseData.rows.filter(r => {
    const bscVal = norm(r['BSC'] || r['bsc'] || '');
    return bscVal === 'no';
  });
  const masterRows = bscFiltered.filter(r => {
    const band4Val = norm(r['Band 4+'] || r['band 4+'] || r['Band4+'] || r['band4+'] || '');
    return band4Val === 'no';
  });

  // --- Parse Tool (may be null/undefined if not uploaded) ---
  let toolMap = new Map();
  let toolCols = null;

  if (toolAB) {
    const toolWb = parseWorkbook(toolAB);
    const toolSheet = toolWb.Sheets[toolWb.SheetNames[0]];
    const toolData = detectHeaders(toolSheet);
    toolCols = resolveToolColumns(toolData.headers);
    toolMap = buildLookupMap(toolData.rows, toolCols.email, toolCols.id);
  }

  const baseZoneCol = findColumn(
    baseData.headers,
    ['Zone', 'Macro Entity Level 2 (Zone)'],
    /^zone$/i,
    /macro.?entity.?level.?2/i
  );

  const output = [];
  const seen = new Set();

  // Iterate master list (filtered base rows) — this drives output, not tool rows
  for (const baseRow of masterRows) {
    const baseEmail = baseEmailCol ? String(baseRow[baseEmailCol] || '').trim() : '';
    const baseId    = baseIdCol    ? String(baseRow[baseIdCol]    || '').trim() : '';

    // Dedup by email or ID
    const dedupeKey = norm(baseEmail) || normId(baseId);
    if (!dedupeKey) continue;
    if (seen.has(dedupeKey)) continue;
    seen.add(dedupeKey);

    // Cross-reference tool file
    const toolRow = findMatch(toolMap, baseEmail, baseId);

    let status, startDate, completedDate, zone;

    if (toolRow && toolCols) {
      const rawStatus = String(cellVal(toolRow, toolCols.status) || '').trim();
      completedDate = fmtDate(cellVal(toolRow, toolCols.complete));
      startDate = fmtDate(cellVal(toolRow, toolCols.start)) || assignDate || '';
      zone = cellVal(toolRow, toolCols.zone);
      status = deriveTranscriptStatus(rawStatus, completedDate);
    } else {
      status = 'Not Started';
      startDate = assignDate || '';
      completedDate = '';
      zone = '';
    }

    const outEmail = toolRow && toolCols
      ? String(cellVal(toolRow, toolCols.email) || '').trim() || baseEmail
      : baseEmail;
    const outId = toolRow && toolCols
      ? String(cellVal(toolRow, toolCols.id) || '').trim() || baseId
      : baseId;

    if (!zone && baseZoneCol) {
      zone = cellVal(baseRow, baseZoneCol);
    }

    output.push(makeOutputRow({
      empId: outId,
      email: outEmail,
      zone,
      status,
      startDate,
      completedDate,
    }));
  }

  return output;
}

function processSection(section, baseAB, toolAB, assignDate, filterFn) {
  const cfg = SECTIONS[section];

  // Parse base
  const baseWb = parseWorkbook(baseAB);
  const baseSheet = baseWb.Sheets[baseWb.SheetNames[0]];
  const baseData = detectHeaders(baseSheet);

  // Parse tool
  const toolWb = parseWorkbook(toolAB);
  const toolSheet = toolWb.Sheets[toolWb.SheetNames[0]];
  const toolData = detectHeaders(toolSheet);

  // Determine base columns dynamically
  let baseIdCol = cfg.baseIdCol;
  let baseEmailCol = cfg.baseEmailCol;

  // Auto-detect if not found
  if (!baseIdCol || !baseData.headers.includes(baseIdCol)) {
    baseIdCol = baseData.headers.find(h => /employee.?id|emp.?id|local.?id/i.test(h)) || null;
  }
  if (!baseEmailCol || !baseData.headers.includes(baseEmailCol)) {
    baseEmailCol = baseData.headers.find(h => /email/i.test(h)) || null;
  }

  // Apply filter (e.g., Band 4+ = Yes)
  let baseRows = baseData.rows;
  if (filterFn) baseRows = baseRows.filter(filterFn);

  const cols = resolveToolColumns(toolData.headers);
  const baseZoneCol = findColumn(
    baseData.headers,
    ['Zone', 'Macro Entity Level 2 (Zone)'],
    /^zone$/i,
    /macro.?entity.?level.?2/i
  );
  const baseMap = buildLookupMap(baseRows, baseEmailCol, baseIdCol);

  const output = [];
  const seen = new Set();

  for (const toolRow of toolData.rows) {
    const empId = cellVal(toolRow, cols.id);
    const email = cellVal(toolRow, cols.email);

    const dedupeKey = norm(email) || normId(String(empId || ''));
    if (dedupeKey && seen.has(dedupeKey)) continue;
    if (dedupeKey) seen.add(dedupeKey);

    const rawStatus = String(cellVal(toolRow, cols.status) || '').trim();
    const completedDate = fmtDate(cellVal(toolRow, cols.complete));
    const startDate = fmtDate(cellVal(toolRow, cols.start)) ||
      (norm(rawStatus) !== 'completed' && assignDate ? assignDate : '');
    let zone = cellVal(toolRow, cols.zone);
    if (!zone) {
      const baseRow = findMatch(baseMap, String(email), empId);
      if (baseRow && baseZoneCol) zone = cellVal(baseRow, baseZoneCol);
    }
    const status = deriveTranscriptStatus(rawStatus, completedDate);

    output.push(makeOutputRow({
      empId, email, zone, status, startDate, completedDate,
    }));
  }

  return output;
}

/**
 * BSC-specific processor: handles Annual Training + BSC/phishing assignment logic.
 *
 * Rules:
 * 1. A user may appear multiple times in the tool file (annual training row + phishing row).
 * 2. Group all rows per user first, then apply priority:
 *    a. If annual training exists and is NOT completed → keep it as-is; skip phishing duplicate.
 *    b. If annual training IS completed → treat phishing row as new BSC assignment;
 *       use phishing assignDate (not original annual start date) as Training Start Date.
 * 3. BSC aging starts from phishing start date, not annual training date.
 * 4. No duplicate pending entries — one output row per user.
 */
function processBSC(baseAB, toolAB, assignDate) {
  const baseWb = parseWorkbook(baseAB);
  const baseSheet = baseWb.Sheets[baseWb.SheetNames[0]];
  const baseData = detectHeaders(baseSheet);

  const toolWb = parseWorkbook(toolAB);
  const toolSheet = toolWb.Sheets[toolWb.SheetNames[0]];
  const toolData = detectHeaders(toolSheet);

  const baseEmailCol = baseData.headers.find(h => /email/i.test(h)) || 'Email - Primary Work';
  const baseIdCol = baseData.headers.find(h => /emp.?id/i.test(h)) || 'Emp ID';

  const cols = resolveToolColumns(toolData.headers);

  // BSC PATCH: Group all tool rows by user key (email or empId).
  // A user can have multiple rows (annual training + phishing-triggered).
  const userRowsMap = new Map();
  for (const toolRow of toolData.rows) {
    const empId = cellVal(toolRow, cols.id);
    const email = cellVal(toolRow, cols.email);
    const key = norm(email) || normId(String(empId || ''));
    if (!key) continue;
    if (!userRowsMap.has(key)) userRowsMap.set(key, []);
    userRowsMap.get(key).push(toolRow);
  }

  const output = [];

  for (const [key, rows] of userRowsMap) {
    // Use first row for identity fields (empId, email, zone)
    const firstRow = rows[0];
    const empId = cellVal(firstRow, cols.id);
    const email = cellVal(firstRow, cols.email);
    const zone = cellVal(firstRow, cols.zone);

    let status, startDate, completedDate;

    if (rows.length === 1) {
      const existingStatus = String(cellVal(rows[0], cols.status) || '').trim();
      const existingStartDate = fmtDate(cellVal(rows[0], cols.start));
      completedDate = '';
      let needsNewAssignment = false;

      if (norm(existingStatus) === 'completed') {
        status = 'Completed';
        startDate = existingStartDate;
        completedDate = fmtDate(cellVal(rows[0], cols.complete));
        needsNewAssignment = true;
      } else if (existingStatus) {
        status = existingStatus;
        startDate = existingStartDate;
        needsNewAssignment = false;
      } else {
        status = 'Not Started';
        startDate = assignDate || '';
        needsNewAssignment = false;
      }

      if (needsNewAssignment && assignDate && existingStartDate) {
        const newD = new Date(assignDate);
        const exD = new Date(existingStartDate.split('/').reverse().join('-'));
        if (!isNaN(newD) && !isNaN(exD) && newD > exD) {
          status = 'Reassigned';
          startDate = assignDate;
          completedDate = '';
        }
      }

    } else {
      const sorted = rows.slice().sort((a, b) => {
        const da = new Date((fmtDate(cellVal(a, cols.start)) || '').split('/').reverse().join('-'));
        const db = new Date((fmtDate(cellVal(b, cols.start)) || '').split('/').reverse().join('-'));
        return (isNaN(da) ? Infinity : da) - (isNaN(db) ? Infinity : db);
      });

      const annualRow = sorted[0];
      const phishRow = sorted[sorted.length - 1];
      const annualStatus = norm(String(cellVal(annualRow, cols.status) || '').trim());

      if (annualStatus !== 'completed') {
        status = String(cellVal(annualRow, cols.status) || '').trim() || 'Not Started';
        startDate = fmtDate(cellVal(annualRow, cols.start)) || assignDate || '';
        completedDate = '';
      } else {
        const phishStartDate = fmtDate(cellVal(phishRow, cols.start)) || assignDate || '';
        const phishStatus = norm(String(cellVal(phishRow, cols.status) || '').trim());

        if (phishStatus === 'completed') {
          status = 'Completed';
          startDate = phishStartDate;
          completedDate = fmtDate(cellVal(phishRow, cols.complete));
        } else {
          status = 'Reassigned';
          startDate = assignDate || phishStartDate;
          completedDate = '';
        }
      }
    }

    output.push(makeOutputRow({
      empId, email, zone, status, startDate, completedDate,
    }));
  }

  return output;
}

/**
 * New Joiner processor: splits by year from Training Start Date.
 * Always iterates tool output rows only — users not present in the tool
 * file are excluded entirely (base file is not used as iteration source).
 */
function processNewJoiner(toolAB, assignDate) {
  const result = { '2024': [], '2025': [], '2026': [] };

  const toolWb = parseWorkbook(toolAB);
  const toolSheet = toolWb.Sheets[toolWb.SheetNames[0]];
  const toolData = detectHeaders(toolSheet);

  const cols = resolveToolColumns(toolData.headers);
  const seen = new Set();

  for (const toolRow of toolData.rows) {
    const empId = cellVal(toolRow, cols.id);
    const email = cellVal(toolRow, cols.email);

    let year = new Date().getFullYear().toString();
    const sd = cellVal(toolRow, cols.start);
    if (sd) {
      const d = sd instanceof Date ? sd : new Date(sd);
      if (!isNaN(d.getTime())) year = String(d.getFullYear());
    }

    const baseKey = norm(email) || normId(String(empId || ''));
    const dedupeKey = baseKey ? baseKey + '|' + year : null;
    if (dedupeKey && seen.has(dedupeKey)) continue;
    if (dedupeKey) seen.add(dedupeKey);

    const rawStatus = String(cellVal(toolRow, cols.status) || '').trim();
    const startDate = fmtDate(cellVal(toolRow, cols.start));
    const completedDate = fmtDate(cellVal(toolRow, cols.complete));
    const zone = cellVal(toolRow, cols.zone);
    const status = deriveTranscriptStatus(rawStatus, completedDate);

    const outRow = makeOutputRow({
      empId, email, zone, status, startDate, completedDate,
    });

    if (result[year]) result[year].push(outRow);
    else result['2026'].push(outRow);
  }

  return result;
}

// ============================================================
// CALCULATE SECTION
// ============================================================
async function calcSection(section) {
  const cfg = SECTIONS[section];
  const baseKey = section + '-base';
  const toolKey = section + '-tool';

  if (!state.files[toolKey]) {
    toast('error', 'Please upload the Tool Output Excel file first');
    return;
  }

  const btn = document.getElementById('btn-calc-' + section);
  const origText = btn.innerHTML;
  btn.innerHTML = '<span class="spinner"></span> Processing…';
  btn.disabled = true;

  await sleep(80); // allow UI to update

  try {
    const assignDate = document.getElementById('date-' + section)?.value || '';
    const baseAB = state.files[baseKey] || null;
    const toolAB = state.files[toolKey];
    state.renderedTables = {};

    let rows;
    if (section === 'newJoiner') {
      const njResult = await runHeavyWork(() => processNewJoiner(toolAB, assignDate));
      state.results[section] = njResult;
      await yieldToMain();
      renderNewJoinerTables(njResult);
      updateNJMetrics(njResult);
    } else if (section === 'appSec' || section === 'cyberOT') {
      rows = await runHeavyWork(() => processToolOnly(toolAB, assignDate));
      const yearResult = splitRowsByYear(rows);
      state.results[section] = yearResult;
      await yieldToMain();
      renderYearTables(section, yearResult);
      updateMetrics(section, rows);
    } else if (section === 'growth') {
      rows = await runHeavyWork(() => processToolOnly(toolAB, assignDate));
      const yearResult = splitRowsByYear(rows);
      state.results[section] = yearResult;
      await yieldToMain();
      renderYearTables(section, yearResult);
      updateMetrics(section, rows);
    } else if (section === 'bsc') {
      if (!baseAB) { toast('error', 'Please upload the Base Userbase Excel'); btn.innerHTML = origText; btn.disabled = false; return; }
      rows = await runHeavyWork(() => processBSC(baseAB, toolAB, assignDate));
      const yearResult = splitRowsByYear(rows);
      state.results[section] = yearResult;
      await yieldToMain();
      renderYearTables(section, yearResult);
      updateMetrics(section, rows);
    } else if (section === 'band4') {
      if (!baseAB) { toast('error', 'Please upload the Phishing Tracking Input Excel'); btn.innerHTML = origText; btn.disabled = false; return; }
      rows = await runHeavyWork(() => processSection(section, baseAB, toolAB, assignDate, r => norm(r['Band 4+'] || r['band 4+'] || '') === 'yes'));
      const yearResult = splitRowsByYear(rows);
      state.results[section] = yearResult;
      await yieldToMain();
      renderYearTables(section, yearResult);
      updateMetrics(section, rows);
    } else if (section === 'phishingNormal') {
      if (!baseAB) { toast('error', 'Please upload the Phishing Tracking Input Excel'); btn.innerHTML = origText; btn.disabled = false; return; }
      rows = await runHeavyWork(() => processPhishingNormal(baseAB, toolAB, assignDate));
      const yearResult = splitRowsByYear(rows);
      state.results[section] = yearResult;
      await yieldToMain();
      renderYearTables(section, yearResult);
      updateMetrics(section, rows);
    }

    const now = new Date();
    const ts = now.toLocaleTimeString('en-GB');
    document.getElementById('lc-' + section).textContent = 'Last calculated: ' + ts;
    document.getElementById('m-' + section + '-calc').textContent = ts;
    document.getElementById('kpi-lastrun').textContent = ts;
    setSectionDownloadButtons(section, true);

    toast('success', `${cfg.label} — calculation complete`);
    updateGlobalKPIs();
  } catch (err) {
    console.error(err);
    toast('error', 'Error: ' + err.message);
  }

  btn.innerHTML = origText;
  btn.disabled = false;
}

// ============================================================
// RENDER TABLE
// ============================================================
function getActiveYearForSection(section) {
  const panel = document.getElementById('panel-' + section);
  if (!panel) return '2026';
  const activeTab = panel.querySelector('.sheet-tab.active');
  if (!activeTab) return '2026';
  const m = (activeTab.getAttribute('onclick') || '').match(/'(\d{4})'/);
  return m ? m[1] : '2026';
}

function renderTable(section, rows) {
  const container = document.getElementById('tbl-' + section);
  if (!container) return;

  if (!rows || rows.length === 0) {
    container.innerHTML = '<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No matching users found</div></div>';
    state.renderedTables[section] = true;
    return;
  }

  const displayRows = rows.slice(0, TABLE_PREVIEW_ROWS);
  const parts = [
    '<table><thead><tr>',
    '<th>#</th><th>Employee ID</th><th>Work Email Address</th><th>Zone</th>',
    '<th>Transcript Status</th><th>Training Start Date</th><th>Transcript Completed Date</th>',
    '</tr></thead><tbody>',
  ];

  for (let i = 0; i < displayRows.length; i++) {
    const row = displayRows[i];
    const status = String(row['Transcript Status'] || '').trim();
    parts.push(
      '<tr>',
      '<td class="mono" style="color:var(--text3)">', String(i + 1), '</td>',
      '<td class="mono">', esc(row['Employee ID']), '</td>',
      '<td style="color:var(--text)">', esc(row['Work Email Address']), '</td>',
      '<td class="mono" style="color:var(--text2)">', esc(rowZone(row)), '</td>',
      '<td>', getStatusBadge(status), '</td>',
      '<td class="mono">', esc(row['Training Start Date']), '</td>',
      '<td class="mono">', esc(row['Transcript Completed Date']), '</td>',
      '</tr>'
    );
  }

  if (rows.length > TABLE_PREVIEW_ROWS) {
    parts.push(
      '<tr><td colspan="7" style="text-align:center;color:var(--text3);padding:12px;font-style:italic">',
      'Showing ', String(TABLE_PREVIEW_ROWS), ' of ', String(rows.length),
      ' rows (all rows included in Excel downloads)</td></tr>'
    );
  }
  parts.push('</tbody></table>');
  container.innerHTML = parts.join('');
  state.renderedTables[section] = true;
}

function renderNewJoinerTables(njResult) {
  document.getElementById('nj-tabs').style.display = 'flex';
  const year = getActiveYearForSection('newJoiner');
  renderTable('newJoiner-' + year, njResult[year] || []);
}

function renderYearTables(section, yearResult) {
  document.getElementById(section + '-year-tabs').style.display = 'flex';
  const year = getActiveYearForSection(section);
  renderTable(section + '-' + year, yearResult[year] || []);
}

function getStatusBadge(status) {
  const s = norm(status);
  if (s === 'completed') return `<span class="badge badge-green">Completed</span>`;
  if (s === 'in progress') return `<span class="badge badge-red">Not Completed</span>`;
  if (s === 'not started') return `<span class="badge badge-red">Not Completed</span>`;
  if (s === 'reassigned') return `<span class="badge badge-yellow">Reassigned</span>`;
  if (s === 'failed') return `<span class="badge badge-red">Failed</span>`;
  if (status) return `<span class="badge badge-red">Not Completed</span>`;
  return `<span class="badge badge-gray">—</span>`;
}

function esc(v) {
  if (v === null || v === undefined) return '';
  return String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

// ============================================================
// METRICS UPDATE
// ============================================================
function updateMetrics(section, rows) {
  if (!rows) return;
  const total = rows.length;
  let completed = 0;
  for (let i = 0; i < rows.length; i++) {
    if (norm(rows[i]['Transcript Status']) === 'completed') completed++;
  }
  const pending = total - completed;
  const pct = total > 0 ? Math.round((completed / total) * 100) : 0;

  document.getElementById('m-'+section+'-total').textContent = total;
  document.getElementById('m-'+section+'-completed').textContent = completed;
  document.getElementById('m-'+section+'-pending').textContent = pending;
  document.getElementById('pb-'+section).style.width = pct + '%';
  const pctEl = document.getElementById('m-'+section+'-pct');
  if (pctEl) pctEl.textContent = pct + '% completed';
  const pctPendingEl = document.getElementById('m-'+section+'-pct-pending');
  if (pctPendingEl) pctPendingEl.textContent = (100 - pct) + '% pending';
  const pbPending = document.getElementById('pb-pending-'+section);
  if (pbPending) pbPending.style.width = (100 - pct) + '%';
  updateZoneStatus(section, rows);
}

function updateNJMetrics(njResult) {
  const all = [...njResult['2024'], ...njResult['2025'], ...njResult['2026']];
  updateMetrics('newJoiner', all);
}

function sortZoneKeys(keys) {
  return keys.slice().sort((a, b) => {
    const ia = ZONE_DISPLAY_ORDER.indexOf(a);
    const ib = ZONE_DISPLAY_ORDER.indexOf(b);
    const ra = ia === -1 ? 999 : ia;
    const rb = ib === -1 ? 999 : ib;
    if (ra !== rb) return ra - rb;
    return a.localeCompare(b);
  });
}

function ensureZonePanel(section) {
  let el = document.getElementById('zone-status-' + section);
  if (el) return el;
  const panel = document.getElementById('panel-' + section);
  if (!panel) return null;
  const metricsRow = panel.querySelector('.metrics-row');
  if (!metricsRow) return null;
  el = document.createElement('div');
  el.id = 'zone-status-' + section;
  el.className = 'zone-status-panel';
  el.style.display = 'none';
  metricsRow.insertAdjacentElement('afterend', el);
  return el;
}

function updateZoneStatus(section, rows) {
  const el = ensureZonePanel(section);
  if (!el) return;

  if (!rows || !rows.length) {
    el.style.display = 'none';
    el.innerHTML = '';
    return;
  }

  const zones = {};
  for (const row of rows) {
    const z = rowZone(row) || 'Unknown';
    if (!zones[z]) zones[z] = { total: 0, completed: 0 };
    zones[z].total++;
    if (norm(row['Transcript Status']) === 'completed') zones[z].completed++;
  }

  const keys = sortZoneKeys(Object.keys(zones));

  let html = '<div class="zone-status-title">Zone-wise completion</div><div class="zone-status-grid">';
  for (const z of keys) {
    const { total, completed } = zones[z];
    const pending = total - completed;
    const pct = total > 0 ? Math.round((completed / total) * 100) : 0;
    html += `<div class="zone-status-card">
      <div class="zone-status-name">${esc(z)}</div>
      <div class="zone-status-stats">
        <span class="zone-stat"><strong>${completed}</strong> done</span>
        <span class="zone-stat zone-stat-pending"><strong>${pending}</strong> pending</span>
        <span class="zone-stat"><strong>${total}</strong> total</span>
      </div>
      <div class="progress-bar-wrap">
        <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${pct}%;background:var(--accent)"></div></div>
      </div>
      <div class="zone-status-pct">${pct}% completed</div>
    </div>`;
  }
  html += '</div>';
  el.innerHTML = html;
  el.style.display = 'block';
}

function updateGlobalKPIs() {
  let totalU = 0, totalC = 0, totalP = 0;
  const sections = ['appSec','cyberOT','growth','newJoiner','bsc','phishingNormal','band4'];
  sections.forEach(s => {
    const r = state.results[s];
    if (!r) return;
    const rows = [...(r['2024']||[]), ...(r['2025']||[]), ...(r['2026']||[])];
    totalU += rows.length;
    const c = rows.filter(x => norm(x['Transcript Status']) === 'completed').length;
    totalC += c;
    totalP += rows.length - c;
  });
  document.getElementById('kpi-total').textContent = totalU || '—';
  document.getElementById('kpi-completed').textContent = totalC || '—';
  document.getElementById('kpi-pending').textContent = totalP || '—';
}

// ============================================================
// DOWNLOAD EXCEL
// ============================================================
function getAllResultRows(section) {
  const result = state.results[section];
  if (!result) return [];
  return [...(result['2024'] || []), ...(result['2025'] || []), ...(result['2026'] || [])];
}

function groupRowsByZone(rows) {
  const groups = {};
  for (const row of rows) {
    const z = rowZone(row) || 'Unknown';
    if (!groups[z]) groups[z] = [];
    groups[z].push(row);
  }
  return groups;
}

function sanitizeSheetName(name) {
  return String(name || 'Unknown').replace(/[\\/*?:\[\]]/g, '_').substring(0, 31);
}

function setSectionDownloadButtons(section, enabled) {
  [
    'btn-dl-' + section,
    'btn-dl-pending-' + section,
    'btn-dl-zone-' + section,
    'btn-dl-pending-zone-' + section,
  ].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.disabled = !enabled;
  });
}

function downloadByZoneSection(section, pendingOnly) {
  const cfg = SECTIONS[section];
  const allRows = getAllResultRows(section);

  if (!allRows.length) {
    toast('error', 'No data available — please calculate first');
    return;
  }

  const rows = pendingOnly
    ? allRows.filter(r => norm(r['Transcript Status']) !== 'completed')
    : allRows;

  if (!rows.length) {
    toast('info', pendingOnly ? 'No pending users in any zone' : 'No data to download');
    return;
  }

  const byZone = groupRowsByZone(rows);
  const zoneKeys = sortZoneKeys(Object.keys(byZone));
  const wb = XLSX.utils.book_new();

  for (const z of zoneKeys) {
    const zoneRows = byZone[z];
    if (!zoneRows.length) continue;
    const ws = XLSX.utils.json_to_sheet(zoneRows, { header: OUTPUT_COLS });
    styleSheet(ws);
    let sheetName = sanitizeSheetName(z);
    let n = 1;
    while (wb.SheetNames.includes(sheetName)) {
      sheetName = sanitizeSheetName(z + ' (' + (++n) + ')');
    }
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  const suffix = pendingOnly ? '_Pending_By_Zone' : '_By_Zone';
  const filename = cfg.outputFile.replace('.xlsx', suffix + '.xlsx');
  XLSX.writeFile(wb, filename);
  const label = pendingOnly ? 'pending' : 'overall';
  toast('success', `Downloaded ${zoneKeys.length} zone sheets (${label}) — ${filename}`);
}

function downloadSection(section) {
  const result = state.results[section];
  if (!result) { toast('error', 'No data to download'); return; }
  const cfg = SECTIONS[section];

  const wb = XLSX.utils.book_new();

  const yearLabel = section === 'newJoiner' ? 'New Joiners' : cfg.label.substring(0, 20);
  ['2024','2025','2026'].forEach(year => {
    const rows = result[year] || [];
    const ws = XLSX.utils.json_to_sheet(rows.length > 0 ? rows : [{}], { header: OUTPUT_COLS });
    styleSheet(ws);
    XLSX.utils.book_append_sheet(wb, ws, yearLabel + ' ' + year);
  });

  XLSX.writeFile(wb, cfg.outputFile);
  toast('success', 'Downloaded: ' + cfg.outputFile);
}

function styleSheet(ws) {
  const wscols = [
    { wch: 16 }, { wch: 38 }, { wch: 14 }, { wch: 18 }, { wch: 20 }, { wch: 22 },
  ];
  ws['!cols'] = wscols;
}

function downloadPendingAppSec() { downloadPendingSection('appSec'); }

function downloadPendingSection(section) {
  const cfg = SECTIONS[section];
  const result = state.results[section];

  if (!result) {
    toast('error', 'No data available — please calculate first');
    return;
  }

  // All sections now store results as year-split objects
  const allRows = [...(result['2024'] || []), ...(result['2025'] || []), ...(result['2026'] || [])];

  if (!allRows || !allRows.length) {
    toast('error', 'No data available');
    return;
  }

  const wb = XLSX.utils.book_new();
  let totalPending = 0;

  ['2024', '2025', '2026'].forEach(year => {
    const yearRows = (result[year] || []).filter(r => norm(r['Transcript Status']) !== 'completed');
    if (yearRows.length > 0) {
      const ws = XLSX.utils.json_to_sheet(yearRows, { header: OUTPUT_COLS });
      styleSheet(ws);
      XLSX.utils.book_append_sheet(wb, ws, 'Pending ' + year);
      totalPending += yearRows.length;
    }
  });

  if (wb.SheetNames.length === 0) {
    toast('info', 'No pending users — everyone has completed!');
    return;
  }

  const filename = cfg.outputFile.replace('.xlsx', '_Pending_Users.xlsx');
  XLSX.writeFile(wb, filename);
  toast('success', `Downloaded ${totalPending} pending users — ${filename}`);
}

// ============================================================
// CLEAR SECTION
// ============================================================
function clearSection(section) {
  // These sections are tool-only — no base file
  const toolOnlySections = ['appSec', 'growth', 'newJoiner'];
  const typesToClear = toolOnlySections.includes(section) ? ['tool'] : ['base', 'tool'];

  delete state.files[section + '-base'];
  delete state.files[section + '-tool'];
  delete state.results[section];
  state.renderedTables = {};

  typesToClear.forEach(t => {
    const uz = document.getElementById('uz-'+section+'-'+t);
    if (uz) { uz.classList.remove('has-file'); }
    const fn = document.getElementById('fn-'+section+'-'+t);
    if (fn) fn.textContent = '';
    // Reset file inputs
    const input = uz ? uz.querySelector('input[type=file]') : null;
    if (input) input.value = '';
  });

  // Reset metrics
  ['total','completed','pending'].forEach(m => {
    const el = document.getElementById('m-'+section+'-'+m);
    if (el) el.textContent = '—';
  });
  const pb = document.getElementById('pb-'+section);
  if (pb) pb.style.width = '0%';
  const lc = document.getElementById('lc-'+section);
  if (lc) lc.textContent = 'No calculation yet';
  const mc = document.getElementById('m-'+section+'-calc');
  if (mc) mc.textContent = 'Never';

  // Reset table
  if (section === 'newJoiner') {
    ['2024','2025','2026'].forEach(y => {
      const t = document.getElementById('tbl-newJoiner-'+y);
      if (t) t.innerHTML = '<div class="empty-state"><div class="empty-icon">📂</div><div class="empty-title">Upload files and calculate</div></div>';
    });
    document.getElementById('nj-tabs').style.display = 'none';
  } else {
    ['2024','2025','2026'].forEach(y => {
      const t = document.getElementById('tbl-'+section+'-'+y);
      if (t) t.innerHTML = '<div class="empty-state"><div class="empty-icon">📂</div><div class="empty-title">Upload files and calculate</div></div>';
    });
    const yearTabsEl = document.getElementById(section+'-year-tabs');
    if (yearTabsEl) yearTabsEl.style.display = 'none';
    // Re-activate first tab
    const firstTab = document.querySelector('#panel-'+section+' .sheet-tab');
    const allTabs = document.querySelectorAll('#panel-'+section+' .sheet-tab');
    const allContents = document.querySelectorAll('#panel-'+section+' .sheet-content');
    allTabs.forEach(t => t.classList.remove('active'));
    allContents.forEach(c => c.classList.remove('active'));
    if (firstTab) firstTab.classList.add('active');
    const firstContent = document.querySelector('#panel-'+section+' .sheet-content');
    if (firstContent) firstContent.classList.add('active');
  }

  setSectionDownloadButtons(section, false);
  updateZoneStatus(section, []);
  toast('info', 'Section cleared');
  updateGlobalKPIs();
}

// ============================================================
// PROCESS ALL
// ============================================================
async function processAll() {
  const sections = ['appSec','cyberOT','growth','newJoiner','bsc','phishingNormal','band4'];
  let processed = 0;
  for (const s of sections) {
    if (state.files[s + '-tool']) {
      await calcSection(s);
      await yieldToMain();
      processed++;
    }
  }
  if (processed === 0) {
    toast('error', 'No tool output files uploaded yet');
  } else {
    toast('success', `Processed ${processed} sections`);
  }
}

// ============================================================
// TOAST
// ============================================================
function toast(type, msg) {
  const c = document.getElementById('toast-container');
  const t = document.createElement('div');
  t.className = 'toast toast-' + type;
  const icon = type === 'success' ? '✓' : type === 'error' ? '✕' : 'ℹ';
  t.innerHTML = `<span>${icon}</span><span>${msg}</span>`;
  c.appendChild(t);
  setTimeout(() => { t.remove(); }, 4000);
}

// ============================================================
// UTILS
// ============================================================
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

