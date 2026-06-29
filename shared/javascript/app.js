'use strict';

// ============================================================
// STATE
// ============================================================
const state = {
  files: {}, // { 'appSec-base': ArrayBuffer, 'appSec-tool': ArrayBuffer, ... }
  results: {}, // processed output per section
  filesCount: 0,
  renderedTables: {}, // tableId -> true (lazy year-tab rendering)
  sheetCache: {}, // cacheKey -> { headers, headerRowIndex, rows, terminatedRemoved }
  debugLogs: {}, // section -> [log lines]
};

/** LMS column headers are usually on Excel row 22 (0-based index 21). */
const DEFAULT_LMS_HEADER_ROW = 21;
const HEADER_SCAN_MAX_ROWS = 35;
const PARSE_YIELD_EVERY = 2000;
/** Bump when parse logic changes so old sheetCache entries are ignored. */
const SHEET_PARSE_VERSION = 5;
const TABLE_PREVIEW_ROWS = 75;
const YEAR_TABS = ['2024', '2025', '2026'];

/** Local server is optional (legacy). Dashboard works via file:// without a server. */
function getUploadApiUrl() {
  return null;
}

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
    /** Rows must match this Macro Entity Level 3 BU value before calculate. */
    buFilterValue: 'GLOBAL GROWTH',
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
    baseEmailCol: 'Employee Email',
    baseGlobalIdCol: 'Global Employee ID',
    baseLocalIdCol: 'Local Employee ID',
    baseIdCol: 'Emp ID',
    toolHeaderRow: 21,
    phishingFilter: null,
    outputFile: 'Phishing_Normal_Output.xlsx',
  },
  band4: {
    label: 'Band 4+ Senior Management',
    trainingName: 'Security & Compliance Awareness Training (Sr. Management) #ABI',
    baseType: 'band4',
    baseEmailCol: 'Employee Email',
    baseGlobalIdCol: 'Global Employee ID',
    baseLocalIdCol: 'Local Employee ID',
    baseIdCol: 'Emp ID',
    toolHeaderRow: 21,
    phishingFilter: 'band4plus', // filter Band 4+ = Yes
    outputFile: 'Band4Plus_Output.xlsx',
    // Training Start Date = tool LMS start date (year tabs + zone tiles bucket by it).
  },
};

const OUTPUT_COLS = ['Employee ID', 'Work Email Address', 'Zone', 'Transcript Status', 'Training Start Date', 'Transcript Completed Date'];

const BSC_APPEND_COLS = [
  'start date_extracted',
  'Training completion date',
  'training completion status',
];

/** Appended to Phishing Tracking userbase (normal + Band 4+). */
const TRACKING_USERBASE_APPEND_COLS = [
  'Training Start Date',
  'Training completion date',
  'training completion status',
];

const USERBASE_ENRICHED_SECTIONS = new Set(['bsc', 'phishingNormal', 'band4']);

const REQUIRED_TRAINING_HEADER_SET = new Set([
  'employee id', 'transcript status', 'work email address',
]);

const MACRO_ZONE_MAP = {
  GLOBAL: 'Global Core',
  'GLOBAL CORE': 'Global Core',
  'ZONE GLOBAL': 'Global Core',
  'ZONE MIDDLE AMERICAS': 'MAZ',
  'MIDDLE AMERICAS': 'MAZ',
  'MIDDLE AMERICA': 'MAZ',
  'ZONE MIDDLE AMERICA': 'MAZ',
  'ZONE MAZ': 'MAZ',
  MAZ: 'MAZ',
  'ZONE EUROPE': 'EUR',
  EUROPE: 'EUR',
  'ZONE EUR': 'EUR',
  EUR: 'EUR',
  'ZONE ASIA PACIFIC': 'APAC',
  'ASIA PACIFIC': 'APAC',
  'ZONE APAC': 'APAC',
  APAC: 'APAC',
  'ZONE AFRICA': 'AFR',
  AFRICA: 'AFR',
  'ZONE AFR': 'AFR',
  AFR: 'AFR',
  'ZONE SOUTH AMERICA': 'SAZ',
  'SOUTH AMERICA': 'SAZ',
  'ZONE SAZ': 'SAZ',
  SAZ: 'SAZ',
  'ZONE NORTH AMERICA': 'NAZ',
  'NORTH AMERICA': 'NAZ',
  'ZONE NAZ': 'NAZ',
  NAZ: 'NAZ',
};

/** Pattern fallback when exact alias is not in MACRO_ZONE_MAP (normalized uppercase key). */
const ZONE_PATTERN_RULES = [
  [/MIDDLE\s*AMERICAS?|^ZONE\s*MAZ$|^MAZ$/, 'MAZ'],
  [/^(ZONE\s*)?EUROPE$|^EUR$/, 'EUR'],
  [/^(ZONE\s*)?ASIA\s*PACIFIC$|^APAC$/, 'APAC'],
  [/^(ZONE\s*)?AFRICA$|^AFR$/, 'AFR'],
  [/^(ZONE\s*)?SOUTH\s*AMERICA$|^SAZ$/, 'SAZ'],
  [/^(ZONE\s*)?NORTH\s*AMERICA$|^NAZ$/, 'NAZ'],
  [/^(ZONE\s*)?GLOBAL(\s*CORE)?$|^GLOBAL$/, 'Global Core'],
];

const ZONE_DISPLAY_ORDER = ['MAZ', 'EUR', 'APAC', 'AFR', 'SAZ', 'NAZ', 'Global Core', 'Growth', 'Unknown'];

function normalizeZoneKey(raw) {
  return String(raw || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}

function mapMacroZone(raw) {
  const s = String(raw || '').trim();
  if (!s) return '';
  const upper = normalizeZoneKey(s);
  if (MACRO_ZONE_MAP[upper]) return MACRO_ZONE_MAP[upper];
  for (const [pat, code] of ZONE_PATTERN_RULES) {
    if (pat.test(upper)) return code;
  }
  return s;
}

function isGlobalCoreZone(mappedZone, rawZone) {
  const z = norm(mappedZone);
  const r = norm(rawZone);
  return z === 'global core' || z === 'global' || r === 'global' || r === 'global core';
}

/** New Joiner: Macro Level 2 zone + Growth subset when BU = GLOBAL GROWTH under Global. */
function resolveNewJoinerZone(toolRow, cols, buCol) {
  const rawZone = cols.zone ? cellVal(toolRow, cols.zone) : '';
  let zone = mapMacroZone(rawZone);
  if (!zone) zone = String(rawZone || '').trim();
  if (buCol && isGlobalCoreZone(zone, rawZone)) {
    const bu = normBuFilterValue(cellVal(toolRow, buCol));
    if (bu === 'GLOBAL GROWTH') return 'Growth';
  }
  return zone;
}

function rowZone(row) {
  return mapMacroZone(row['Zone'] || row['Macro Entity Level 2 (Zone)'] || '');
}

function findColumn(headers, exactNames, ...patternGroups) {
  for (const name of exactNames) {
    const want = normHeaderCell(name);
    const hit = headers.find(h => normHeaderCell(h) === want);
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

function findLastRealHeaderIndex(headers) {
  let last = -1;
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    if (h && !h.startsWith('__col_')) last = i;
  }
  return last;
}

function appendColumnsAfterBase(baseHeaders, appendCols) {
  const lastIdx = findLastRealHeaderIndex(baseHeaders);
  const headers = baseHeaders.slice(0, lastIdx + 1);
  for (const col of appendCols) {
    if (!headers.some(h => normHeaderCell(h) === normHeaderCell(col))) {
      headers.push(col);
    }
  }
  return headers;
}

function appendBscColumnsAfterBase(baseHeaders) {
  return appendColumnsAfterBase(baseHeaders, BSC_APPEND_COLS);
}

function isEnrichedUserbaseResult(result) {
  return result && Array.isArray(result.enriched) && Array.isArray(result.headers);
}

function exportHeadersFromEnriched(headers) {
  return headers.filter(h => !String(h).startsWith('_'));
}

function findBaseCenterColumn(headers) {
  return findColumn(
    headers,
    ['Center', 'Work Center', 'Cost Center', 'Personnel Area'],
    /^center$/i,
    /work\s*center/i,
    /cost\s*center/i,
    /personnel\s*area/i
  );
}

function findBaseOriginalStartDateColumn(headers) {
  return findColumn(
    headers,
    ['Start Date', 'start date', 'Original Start Date', 'Training Start Date'],
    /^start\s*date$/i,
    /original\s*start/i
  );
}

function buildToolRowsMapByUser(toolRows, cols) {
  const map = new Map();
  const add = (key, row) => {
    if (!key) return;
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
  };
  for (const toolRow of toolRows) {
    const empId = cellVal(toolRow, cols.id);
    const email = cellVal(toolRow, cols.email);
    add(norm(email), toolRow);
    add(normId(String(empId || '')), toolRow);
  }
  return map;
}

function getToolRowsForBaseUser(toolMap, baseEmail, baseId, cols) {
  const out = [];
  const pushRows = (key) => {
    const list = toolMap.get(key);
    if (!list) return;
    for (const r of list) out.push(r);
  };
  if (baseEmail) pushRows(norm(baseEmail));
  if (baseId) pushRows(normId(baseId));
  return out;
}

function resolveBscTrainingFromToolRows(toolRows, cols, assignDate) {
  const startDateExtracted = fmtDate(assignDate) || '';
  let status = 'Not Started';
  let completedDate = '';

  if (!toolRows || !toolRows.length) {
    return { status: 'Not Found', startDateExtracted, completedDate };
  }

  // Multiple matches: use the latest training-start row only.
  const selectedRow = toolRows.slice().sort((a, b) => {
    const da = new Date((fmtDate(cellVal(a, cols.start)) || '').split('/').reverse().join('-'));
    const db = new Date((fmtDate(cellVal(b, cols.start)) || '').split('/').reverse().join('-'));
    const ta = isNaN(da) ? -Infinity : da.getTime();
    const tb = isNaN(db) ? -Infinity : db.getTime();
    return tb - ta;
  })[0];

  // Year tabs + zone tiles bucket by the tool's per-row training start date;
  // fall back to the assignment date only when the matched row has no start date.
  const rowStartDate = fmtDate(cellVal(selectedRow, cols.start)) || startDateExtracted;

  // Terminated in the tool export overrides everything: mark the userbase row Terminated.
  if (isTerminatedRow(selectedRow, cols.empStatus)) {
    return { status: 'Terminated', startDateExtracted: rowStartDate, completedDate: 'Terminated' };
  }

  const rawStatus = String(cellVal(selectedRow, cols.status) || '').trim();
  completedDate = fmtDate(cellVal(selectedRow, cols.complete));
  status = deriveTranscriptStatus(rawStatus, completedDate);

  return { status, startDateExtracted: rowStartDate, completedDate };
}

function rowCompletionStatus(row, section) {
  if (section === 'bsc' || section === 'phishingNormal' || section === 'band4') {
    return norm(row['training completion status'] || row['Transcript Status'] || '');
  }
  return norm(row['Transcript Status'] || '');
}

function userbaseTrackingOutcome(row) {
  const s = norm(row['training completion status'] || '');
  if (s === 'completed') return 'completed';
  if (s === 'not found') return 'notFound';
  if (s === 'terminated') return 'terminated';
  return 'notCompleted';
}

/** BSC: a userbase row whose matched tool user was Terminated (kept, not dropped). */
function isTerminatedOutcomeRow(row, section) {
  return rowCompletionStatus(row, section) === 'terminated';
}

/** Sections that show the "Unmapped/Terminated" split (BSC only). */
function showsUnmappedTerminated(section) {
  return section === 'bsc';
}

function isUnmappedRow(row, section) {
  if (section === 'phishingNormal' || section === 'band4') {
    return userbaseTrackingOutcome(row) === 'notFound';
  }
  if (section === 'bsc') {
    return rowCompletionStatus(row, section) === 'not found';
  }
  return false;
}

function isUserbaseEnrichedSection(section) {
  return USERBASE_ENRICHED_SECTIONS.has(section);
}

const ZONE_DL_DELEGATE_BOUND = new Set();
let centerDlDelegateBound = false;

function bindZoneDownloadDelegation(section, panelEl) {
  if (!panelEl || ZONE_DL_DELEGATE_BOUND.has(section)) return;
  ZONE_DL_DELEGATE_BOUND.add(section);
  panelEl.addEventListener('click', (e) => {
    const btn = e.target.closest('.zone-dl-btn[data-zone][data-mode]');
    if (!btn || !panelEl.contains(btn)) return;
    e.preventDefault();
    downloadZoneRows(section, btn.getAttribute('data-zone'), btn.getAttribute('data-mode'));
  });
}

function bindCenterDownloadDelegation(panelEl) {
  if (!panelEl || centerDlDelegateBound) return;
  centerDlDelegateBound = true;
  panelEl.addEventListener('click', (e) => {
    const btn = e.target.closest('.zone-dl-btn[data-center][data-mode]');
    if (!btn || !panelEl.contains(btn)) return;
    e.preventDefault();
    downloadCenterRows('bsc', btn.getAttribute('data-center'), btn.getAttribute('data-mode'));
  });
}

function rowInPhishingNormalScope(row) {
  // User uploads already-filtered data for Phishing Normal.
  return true;
}

function rowInBand4Scope(row) {
  // User uploads already-filtered data for Band 4+.
  return true;
}

function rowZoneForSection(row, section) {
  if (section === 'bsc' || section === 'phishingNormal' || section === 'band4') {
    const z =
      row._zone ||
      cellVal(row, 'Zone') ||
      cellVal(row, 'Macro Entity Level 2 (Zone)') ||
      '';
    const mapped = mapMacroZone(z);
    return mapped || z || 'Unknown';
  }
  return rowZone(row);
}

function rowCenterForSection(row, section) {
  if (section === 'bsc') {
    const c = row._center || row['Center'] || row['Work Center'] || '';
    return String(c || '').trim() || 'Unknown';
  }
  return 'Unknown';
}

function getEnrichedPreviewHeaders(section, headers) {
  const pick = (patterns, exact) => {
    const h = findColumn(headers, exact, ...patterns);
    return h || null;
  };
  const ordered = [
    pick([/emp.?id/i, /local.?employee/i], ['Emp ID', 'Employee ID', 'Local Employee ID']),
    pick([/email/i], ['Email - Primary Work', 'Employee Email']),
    pick([/zone/i, /macro.?entity.?level.?2/i], ['Zone', 'Macro Entity Level 2 (Zone)']),
  ];
  if (section === 'bsc') {
    ordered.push(findBaseCenterColumn(headers), findBaseOriginalStartDateColumn(headers));
    ordered.push('start date_extracted', 'Training completion date', 'training completion status');
  } else {
    ordered.push('Training Start Date', 'Training completion date', 'training completion status');
  }
  const seen = new Set();
  return ordered.filter(Boolean).filter(h => {
    const k = normHeaderCell(h);
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });
}

function findNewJoinerOriginalHireDateColumn(headers) {
  return findColumn(
    headers,
    [
      'Original Hire Date',
      'New J orginal hire date',
      'New J original hire date',
      'New Joiner Original Hire Date',
      'First Hire Date',
    ],
    /^original\s*hire\s*date$/i,
    /^first\s*hire\s*date$/i,
    /new\s*j\s*orginal\s*hire/i,
    /new\s*j\s*original\s*hire/i,
    /new\s*joiner.*original\s*hire/i
  );
}

/** Excel serial only when decoded year is plausible (avoids 2024–2026 mistaken as serial → 1905). */
function parseExcelSerialToCalendarParts(n) {
  if (typeof n !== 'number' || isNaN(n) || n < 25000 || n > 60000) return null;
  if (typeof XLSX === 'undefined' || !XLSX.SSF) return null;
  const parts = XLSX.SSF.parse_date_code(n);
  if (!parts || !parts.y || parts.y < 1990 || parts.y > 2036) return null;
  return { day: parts.d, month: parts.m, year: parts.y };
}

/** Whole-number year in a cell (e.g. 2025), not an Excel serial. */
function parsePlainCalendarYearNumber(n) {
  if (typeof n !== 'number' || isNaN(n)) return null;
  const y = Math.round(n);
  if (y !== n || y < 1990 || y > 2036) return null;
  return { day: 1, month: 1, year: y };
}

function normalizeSlashDateYear(y) {
  if (isNaN(y)) return null;
  if (y >= 100 && y < 1900) return null;
  if (y < 100) return y < 50 ? 2000 + y : 1900 + y;
  return y;
}

/**
 * Parse LMS date cells (Training Start Date, Completed Date, Original Hire Date, etc.).
 * Handles Excel serials, plain years (2024–2026), dd/mm/yyyy, and ISO dates.
 */
function parseCalendarParts(v) {
  if (v === null || v === undefined || v === '') return null;

  let day;
  let month;
  let year;

  if (v instanceof Date && !isNaN(v.getTime())) {
    // Calendar date from Excel — use local components (not UTC) to avoid day/year shifts.
    day = v.getDate();
    month = v.getMonth() + 1;
    year = v.getFullYear();
  } else if (typeof v === 'number' && !isNaN(v)) {
    const plain = parsePlainCalendarYearNumber(v);
    if (plain) return plain;
    const serial = parseExcelSerialToCalendarParts(v);
    if (serial) return serial;
    return null;
  } else {
    const s = String(v).trim();
    if (!s) return null;

    const yearOnly = s.match(/^(20\d{2})$/);
    if (yearOnly) {
      year = parseInt(yearOnly[1], 10);
      return { day: 1, month: 1, year };
    }

    const slash = s.split('/');
    if (slash.length === 3 && slash[2]) {
      const a = parseInt(slash[0], 10);
      const b = parseInt(slash[1], 10);
      year = normalizeSlashDateYear(parseInt(slash[2], 10));
      if (isNaN(a) || isNaN(b) || year == null) return null;
      if (b > 12 && a <= 12) {
        month = a;
        day = b;
      } else if (a > 12 && b <= 12) {
        day = a;
        month = b;
      } else {
        day = a;
        month = b;
      }
    } else {
      const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (iso) {
        year = parseInt(iso[1], 10);
        month = parseInt(iso[2], 10);
        day = parseInt(iso[3], 10);
      } else {
        const n = parseFloat(s);
        if (!isNaN(n)) {
          const plain = parsePlainCalendarYearNumber(n);
          if (plain && String(Math.round(n)) === s.replace(/\.0+$/, '')) return plain;
          const serial = parseExcelSerialToCalendarParts(n);
          if (serial) return serial;
        }
        const d = new Date(s);
        if (!isNaN(d.getTime()) && d.getFullYear() > 1900) {
          day = d.getDate();
          month = d.getMonth() + 1;
          year = d.getFullYear();
        } else {
          return null;
        }
      }
    }
  }

  if (!year || !month || !day || month < 1 || month > 12 || day < 1 || day > 31) return null;
  return { day, month, year };
}

const parseHireCalendarParts = parseCalendarParts;

function formatCalendarPartsGb(parts) {
  if (!parts) return '';
  const dd = String(parts.day).padStart(2, '0');
  const mm = String(parts.month).padStart(2, '0');
  return dd + '/' + mm + '/' + parts.year;
}

/**
 * New Joiner hire cell → display text (minimal conversion; keeps Excel display when possible).
 */
function hireDateCellToString(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date && !isNaN(v.getTime())) {
    const dd = String(v.getDate()).padStart(2, '0');
    const mm = String(v.getMonth() + 1).padStart(2, '0');
    return dd + '/' + mm + '/' + v.getFullYear();
  }
  if (typeof v === 'number' && !isNaN(v)) {
    if (Number.isInteger(v) && v >= 1990 && v <= 2036) return '01/01/' + v;
    const serial = parseExcelSerialToCalendarParts(v);
    if (serial) return formatCalendarPartsGb(serial);
    return String(v);
  }
  let s = String(v).trim();
  if (!s) return '';
  if (/^\d{5}(\.\d+)?$/.test(s)) {
    const serial = parseExcelSerialToCalendarParts(parseFloat(s));
    if (serial) return formatCalendarPartsGb(serial);
  }
  return s;
}

/**
 * New Joiner year tab: find 2024, 2025, or 2026 in the hire cell as text (no strict date parse).
 */
function extractHireYearFromString(s) {
  const t = String(s || '').trim();
  if (!t) return null;

  if (/^2024$/.test(t)) return 2024;
  if (/^2025$/.test(t)) return 2025;
  if (/^2026$/.test(t)) return 2026;

  const slashEnd = t.match(/\/(2024|2025|2026)\s*$/);
  if (slashEnd) return parseInt(slashEnd[1], 10);

  const iso = t.match(/(2024|2025|2026)-\d{1,2}-\d{1,2}/);
  if (iso) return parseInt(iso[1], 10);

  const dashEnd = t.match(/(2024|2025|2026)\s*$/);
  if (dashEnd && /[-/]/.test(t)) return parseInt(dashEnd[1], 10);

  const hits = [];
  const re = /2024|2025|2026/g;
  let m;
  while ((m = re.exec(t)) !== null) hits.push(parseInt(m[0], 10));
  if (hits.length) return hits[hits.length - 1];

  return null;
}

/** New Joiner: only 2024, 2025, 2026 hire years. Year from text; display = cell string when possible. */
function normalizeOriginalHireDate(v) {
  const display = hireDateCellToString(v);
  if (!display) return null;

  const year = extractHireYearFromString(display);
  if (year !== 2024 && year !== 2025 && year !== 2026) return null;

  return {
    formatted: display,
    yearTab: String(year),
    calendarYear: year,
  };
}

function newJoinerYearTabFromHire(v) {
  const hireNorm = normalizeOriginalHireDate(v);
  return hireNorm ? hireNorm.yearTab : null;
}

function findMacroEntityLevel3BUColumn(headers) {
  return findColumn(
    headers,
    [
      'Macro entity level three BU Description',
      'Macro Entity Level 3 (BU Description)',
      'Macro Entity Level 3 BU Description',
      'Macro Entity Level Three BU Description',
    ],
    /macro.?entity.?level.?three.*bu/i,
    /macro.?entity.?level.?3.*bu/i,
    /macro.?entity.?level.?3/i,
    /macro.?entity.?level.?three/i
  );
}

function normBuFilterValue(v) {
  return String(v ?? '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, ' ');
}

function rowMatchesBuFilter(row, buCol, expected) {
  if (!buCol) return false;
  return normBuFilterValue(cellVal(row, buCol)) === normBuFilterValue(expected);
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
    empStatus: findColumn(headers, ['Employee Status'], /employee\s*status/i),
  };
}

function cellVal(row, col) {
  if (!col || !row) return '';
  if (Object.prototype.hasOwnProperty.call(row, col)) {
    const v = row[col];
    if (v === null || v === undefined) return '';
    return v;
  }
  const want = norm(col);
  for (const k in row) {
    if (!Object.prototype.hasOwnProperty.call(row, k)) continue;
    if (norm(k) === want) {
      const v = row[k];
      if (v === null || v === undefined) return '';
      return v;
    }
  }
  return '';
}

function makeOutputRow({ empId, email, zone, status, startDate, completedDate, mapZone = true }) {
  return {
    'Employee ID': empId != null && empId !== '' ? String(empId).trim() : '',
    'Work Email Address': email != null && email !== '' ? String(email).trim() : '',
    'Zone': mapZone ? mapMacroZone(zone) : String(zone || '').trim(),
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
  document.querySelectorAll('#panel-newJoiner .sheet-tab').forEach(t => {
    t.classList.toggle('active', (t.getAttribute('data-year') || '') === year);
  });
  document.querySelectorAll('#panel-newJoiner .sheet-content').forEach(c => c.classList.remove('active'));
  const sheet = document.getElementById('nj-sheet-' + year);
  if (sheet) sheet.classList.add('active');
  const result = state.results.newJoiner;
  if (result) {
    renderTable('newJoiner-' + year, result[year] || []);
    updateNJMetrics(result, year);
  }
}

function switchYearTab(section, year) {
  document.querySelectorAll('#panel-'+section+' .sheet-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('#panel-'+section+' .sheet-content').forEach(c => c.classList.remove('active'));
  event.target.classList.add('active');
  document.getElementById(section+'-sheet-'+year).classList.add('active');
  const result = state.results[section];
  if (!result) return;
  if (isEnrichedUserbaseResult(result)) {
    renderEnrichedUserbaseTable(section, section + '-' + year, result[year] || [], result.headers);
  } else {
    renderTable(section + '-' + year, result[year] || []);
  }
}

/**
 * Split a flat array of rows into { '2024': [], '2025': [], '2026': [] } using
 * Training Start Date. (BSC uses data-driven year tabs instead — see buildBscDynamicYearTabs.)
 */
/** Extract calendar year for year tabs (same parser as fmtDate / New Joiner hire dates). */
function parseDateYearForBucket(v) {
  const cal = parseCalendarParts(v);
  if (cal && cal.year) return cal.year;
  const currentYear = new Date().getFullYear();
  if (v === null || v === undefined || v === '') return currentYear;
  const s = String(v).trim();
  if (!s) return currentYear;
  const d = new Date(s);
  if (!isNaN(d.getTime()) && d.getFullYear() > 1900) return d.getFullYear();
  return currentYear;
}

function yearBucketFromStartDate(sd) {
  const year = parseDateYearForBucket(sd);
  if (isNaN(year) || year <= 2024) return '2024';
  if (year === 2025) return '2025';
  return '2026';
}

/**
 * BSC year tabs are data-driven (not the fixed 2024/2025/2026 buckets): each row
 * lands in a tab for its actual start-date calendar year (e.g. 2022, 2023, ...).
 */
function bscYearBucket(sd) {
  const year = parseDateYearForBucket(sd);
  if (isNaN(year) || year < 1900) return getCurrentDashboardYearBucket();
  return String(year);
}

/** BSC: group enriched rows by their actual start-date year; returns sorted years + map. */
function buildBscDynamicYearTabs(rows) {
  const byYear = {};
  for (const row of rows) {
    const y = bscYearBucket(row['start date_extracted'] || '');
    (byYear[y] || (byYear[y] = [])).push(row);
  }
  const years = Object.keys(byYear).sort();
  return { years, byYear };
}

function splitRowsByYear(rows) {
  const result = { '2024': [], '2025': [], '2026': [] };
  for (const row of rows) {
    const bucket = yearBucketFromStartDate(row['Training Start Date']);
    result[bucket].push(row);
  }
  return result;
}

function switchToYearWithMostRows(section, yearResult) {
  let best = '2026';
  let max = 0;
  for (const y of YEAR_TABS) {
    const n = (yearResult[y] || []).length;
    if (n > max) {
      max = n;
      best = y;
    }
  }
  if (max === 0) best = getCurrentDashboardYearBucket();
  const panel = document.getElementById('panel-' + section);
  if (!panel) return best;
  panel.querySelectorAll('.sheet-tab').forEach(t => {
    const tabYear = t.getAttribute('data-year') ||
      ((t.getAttribute('onclick') || '').match(/'(\d{4})'/) || [])[1];
    t.classList.toggle('active', tabYear === best);
  });
  panel.querySelectorAll('.sheet-content').forEach(c => c.classList.remove('active'));
  const content = document.getElementById(
    section === 'newJoiner' ? 'nj-sheet-' + best : section + '-sheet-' + best
  );
  if (content) content.classList.add('active');
  return best;
}

// ============================================================
// FILE UPLOAD
// ============================================================
function ensureDebugPanel(section) {
  let wrap = document.getElementById('debug-wrap-' + section);
  if (wrap) return wrap;
  const panel = document.getElementById('panel-' + section);
  if (!panel) return null;
  const uploadsGrid = panel.querySelector('.uploads-grid');
  if (!uploadsGrid) return null;

  wrap = document.createElement('div');
  wrap.id = 'debug-wrap-' + section;
  wrap.className = 'section-note';
  wrap.style.marginTop = '10px';
  wrap.style.padding = '10px';
  wrap.style.border = '1px solid var(--border)';
  wrap.style.borderRadius = '8px';
  wrap.style.background = 'rgba(255,255,255,0.02)';
  wrap.innerHTML = `
    <div style="font-size:12px;font-weight:600;color:var(--text2);margin-bottom:6px">Debug logs</div>
    <pre id="debug-log-${section}" style="margin:0;white-space:pre-wrap;font-size:11px;line-height:1.35;color:var(--text3);max-height:170px;overflow:auto">No logs yet.</pre>
  `;
  uploadsGrid.insertAdjacentElement('afterend', wrap);
  return wrap;
}

function ensureDebugPanels() {
  const sections = Object.keys(SECTIONS);
  for (const s of sections) ensureDebugPanel(s);
}

function appendDebugLog(section, msg) {
  ensureDebugPanel(section);
  if (!state.debugLogs[section]) state.debugLogs[section] = [];
  const now = new Date().toLocaleTimeString('en-GB');
  state.debugLogs[section].push('[' + now + '] ' + msg);
  if (state.debugLogs[section].length > 300) {
    state.debugLogs[section] = state.debugLogs[section].slice(-300);
  }
  const pre = document.getElementById('debug-log-' + section);
  if (!pre) return;
  pre.textContent = state.debugLogs[section].join('\n');
  pre.scrollTop = pre.scrollHeight;
}

function clearDebugLog(section) {
  state.debugLogs[section] = [];
  const pre = document.getElementById('debug-log-' + section);
  if (pre) pre.textContent = 'No logs yet.';
}

function formatHeaderList(headers) {
  return headers
    .filter(h => h && !String(h).startsWith('__col_'))
    .join(' | ');
}

function logSheetParseDiagnostics(section, label, data) {
  if (!data) return;
  const names = data.sheetNames || [];
  if (names.length > 1) {
    appendDebugLog(section, label + ' workbook tabs: ' + names.join(', '));
    const scores = data.sheetScores || [];
    for (let i = 0; i < scores.length; i++) {
      appendDebugLog(section, '  ' + scores[i]);
    }
  }
  if (data.sheetPickNote) appendDebugLog(section, label + ' tab selected: ' + data.sheetPickNote);
  if (data.sheetName) appendDebugLog(section, label + ' worksheet: "' + data.sheetName + '"');
  const d = data.parseDiagnostics;
  if (d) {
    appendDebugLog(
      section,
      label + ' Excel range: ' + (d.declaredRef || '(none)') +
      ' | cells loaded: ' + (d.cellCount != null ? d.cellCount.toLocaleString() : '?') +
      ' | last row with data: ' + (d.maxRowExcel || '?')
    );
    appendDebugLog(
      section,
      label + ' header row: ' + (d.headerRowExcel || '?') +
      ' | non-empty data rows in sheet: ' + (d.indexedDataRows != null ? d.indexedDataRows.toLocaleString() : '?') +
      ' | rows after parse (before terminated filter): ' + (d.parsedRows != null ? d.parsedRows.toLocaleString() : '?')
    );
    if (d.indexedDataRows != null && d.parsedRows != null && d.indexedDataRows > d.parsedRows + 50) {
      appendDebugLog(
        section,
        label + ' warning: ' + (d.indexedDataRows - d.parsedRows).toLocaleString() +
        ' indexed rows were not included in parse — check for blank rows or wrong header row'
      );
    }
  }
}

function logHireDateParseAudit(section, toolData) {
  const hireCol = findNewJoinerOriginalHireDateColumn(toolData.headers);
  if (!hireCol) {
    appendDebugLog(section, 'Parse audit: no hire-date column in parsed headers');
    return;
  }
  const stats = {
    empty: 0, number: 0, string: 0, date: 0, other: 0, parseable: 0, inScope: 0,
  };
  const samples = { parseFail: [], inScope: [] };
  for (let i = 0; i < toolData.rows.length; i++) {
    const raw = cellVal(toolData.rows[i], hireCol);
    if (raw === '' || raw == null) {
      stats.empty++;
      continue;
    }
    if (raw instanceof Date) stats.date++;
    else if (typeof raw === 'number') stats.number++;
    else if (typeof raw === 'string') stats.string++;
    else stats.other++;
    const display = hireDateCellToString(raw);
    if (display) {
      stats.parseable++;
      const y = extractHireYearFromString(display);
      const norm = normalizeOriginalHireDate(raw);
      if (norm) {
        stats.inScope++;
        if (samples.inScope.length < 3) {
          samples.inScope.push(display + ' → year ' + norm.yearTab);
        }
      } else if (y && samples.parseFail.length < 3) {
        samples.parseFail.push(display + ' (year ' + y + ' not in 2024–26)');
      }
    } else if (samples.parseFail.length < 3) {
      samples.parseFail.push(String(raw));
    }
  }
  appendDebugLog(
    section,
    'Parse audit — "' + hireCol + '" (year from text 2024/2025/2026): empty=' +
      stats.empty.toLocaleString() +
      ', number=' + stats.number.toLocaleString() +
      ', string=' + stats.string.toLocaleString() +
      ', has text=' + stats.parseable.toLocaleString() +
      ', in-scope=' + stats.inScope.toLocaleString()
  );
  if (samples.inScope.length) {
    appendDebugLog(section, '  Sample OK: ' + samples.inScope.join(' | '));
  }
  if (samples.parseFail.length) {
    appendDebugLog(section, '  Sample parse failed: ' + samples.parseFail.join(' | '));
  }
}

function logToolMappings(section, toolData) {
  const cols = resolveToolColumns(toolData.headers);
  logSheetParseDiagnostics(section, 'Tool', toolData);
  appendDebugLog(section, 'Tool file parsed; header row = Excel row ' + (toolData.headerRowIndex + 1));
  appendDebugLog(section, 'Tool headers: ' + (formatHeaderList(toolData.headers) || '(none)'));
  appendDebugLog(section, 'Mapped tool columns -> ID: ' + (cols.id || 'NOT FOUND') +
    ', Email: ' + (cols.email || 'NOT FOUND') +
    ', Status: ' + (cols.status || 'NOT FOUND') +
    ', Start: ' + (cols.start || 'NOT FOUND') +
    ', Complete: ' + (cols.complete || 'NOT FOUND') +
    ', Zone: ' + (cols.zone || 'NOT FOUND'));
}

function logBaseMappings(section, baseData, cfg) {
  logSheetParseDiagnostics(section, 'Userbase', baseData);
  appendDebugLog(section, 'Base file parsed; header row = Excel row ' + (baseData.headerRowIndex + 1));
  appendDebugLog(section, 'Base headers: ' + (formatHeaderList(baseData.headers) || '(none)'));
  const zoneCol = findColumn(
    baseData.headers,
    ['Zone', 'Macro Entity Level 2 (Zone)'],
    /^zone$/i,
    /macro.?entity.?level.?2/i
  );
  const centerCol = findBaseCenterColumn(baseData.headers);
  if (section === 'phishingNormal' || section === 'band4') {
    const mc = resolveUserbaseMatchCols(baseData, cfg);
    appendDebugLog(section, 'Mapped userbase match keys -> Employee Email: ' + (mc.emailCol || 'NOT FOUND') +
      ', Global Employee ID: ' + (mc.globalIdCol || 'NOT FOUND') +
      ', Local Employee ID: ' + (mc.localIdCol || 'NOT FOUND') +
      (mc.empIdCol && mc.empIdCol !== mc.localIdCol ? ', Emp ID: ' + mc.empIdCol : '') +
      ', Zone: ' + (zoneCol || 'NOT FOUND'));
  } else {
    const cols = resolveBaseIdEmailCols(baseData, cfg);
    appendDebugLog(section, 'Mapped base columns -> ID: ' + (cols.baseIdCol || 'NOT FOUND') +
      ', Email: ' + (cols.baseEmailCol || 'NOT FOUND') +
      ', Zone: ' + (zoneCol || 'NOT FOUND') +
      (section === 'bsc' ? ', Center: ' + (centerCol || 'NOT FOUND') : ''));
  }
}

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

async function persistUploadToDisk() {
  return null;
}

function processFileUpload(file, section, type) {
  if (!file.name.match(/\.(xlsx|xls)$/i)) {
    toast('error', 'Please upload an Excel file (.xlsx or .xls)');
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    const key = section + '-' + type;
    delete state.sheetCache[key];
    const vKey = sheetCacheKey(key);
    if (vKey) delete state.sheetCache[vKey];
    state.files[key] = e.target.result;
    state.filesCount++;
    document.getElementById('kpi-files').textContent = state.filesCount;
    document.getElementById('fn-'+section+'-'+type).textContent = '✓ ' + file.name;
    const uz = document.getElementById('uz-'+section+'-'+type);
    uz.classList.add('has-file');
    toast('info', 'Loaded: ' + file.name);
    appendDebugLog(section, 'Opened ' + type.toUpperCase() + ' Excel: ' + file.name + ' (' + file.size.toLocaleString() + ' bytes)');
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
  // Keep Excel date serials as numbers — cellDates:true turns them into JS Date objects and
  // often shifts the calendar day at parse time (timezone), which breaks hire/start columns.
  return XLSX.read(new Uint8Array(ab), {
    type: 'array',
    cellDates: false,
    cellText: true,
    dense: false,
    sheetStubs: false,
  });
}

function yieldToMain() {
  return new Promise(resolve => setTimeout(resolve, 0));
}

async function runHeavyWork(fn) {
  await yieldToMain();
  return fn();
}

function cellLooksLikeDate(cell) {
  if (!cell) return false;
  if (cell.t === 'd') return true;
  if (typeof cell.v === 'number' && !isNaN(cell.v)) {
    if (parsePlainCalendarYearNumber(cell.v) || parseExcelSerialToCalendarParts(cell.v)) return true;
  }
  if (cell.z && /[dmyhs]/i.test(String(cell.z))) return true;
  return false;
}

function getCellValue(cell) {
  if (!cell) return '';
  if (cell.t === 'd' && cell.v instanceof Date) return cell.v;
  if (typeof cell.v === 'number' && !isNaN(cell.v)) {
    if (cell.v === 0) return '';
    if (parsePlainCalendarYearNumber(cell.v)) return cell.v;
    if (parseExcelSerialToCalendarParts(cell.v)) return cell.v;
    return cell.v;
  }
  const hasV = cell.v != null && cell.v !== '';
  const hasW = cell.w != null && cell.w !== '';
  // Formula/export cells: display text only (common for date columns in LMS exports).
  if (hasW && (!hasV || cellLooksLikeDate(cell))) return cell.w;
  if (hasV) return cell.v;
  if (hasW) return cell.w;
  return '';
}

/**
 * True used range — sheet['!ref'] is often too narrow on LMS exports (row 22 has more columns).
 */
function getSheetBounds(sheet) {
  let range = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  if (sheet['!ref']) {
    try {
      range = XLSX.utils.decode_range(sheet['!ref']);
    } catch (e) {
      /* use defaults */
    }
  }
  let maxR = range.e.r;
  let maxC = range.e.c;
  let minR = range.s.r;
  let minC = range.s.c;
  let cellCount = 0;
  for (const addr in sheet) {
    if (addr.charAt(0) === '!') continue;
    let cell;
    try {
      cell = XLSX.utils.decode_cell(addr);
    } catch (e) {
      continue;
    }
    cellCount++;
    if (cell.r > maxR) maxR = cell.r;
    if (cell.c > maxC) maxC = cell.c;
    if (cell.r < minR) minR = cell.r;
    if (cell.c < minC) minC = cell.c;
  }
  return {
    s: { r: minR, c: minC },
    e: { r: maxR, c: maxC },
    cellCount,
    declaredRef: sheet['!ref'] || '',
  };
}

/** Align sheet['!ref'] with every loaded cell (helps sheet_to_json and row scans). */
function expandSheetRef(sheet) {
  const b = getSheetBounds(sheet);
  if (b.e.r < b.s.r) return b;
  sheet['!ref'] = XLSX.utils.encode_range({ s: { r: b.s.r, c: b.s.c }, e: { r: b.e.r, c: b.e.c } });
  return b;
}

/**
 * One pass over all cell addresses — avoids missing rows/columns when !ref is wrong.
 */
function buildSheetRowIndex(sheet) {
  const bounds = expandSheetRef(sheet);
  const rowCells = new Map();
  for (const addr in sheet) {
    if (addr.charAt(0) === '!') continue;
    let cell;
    try {
      cell = XLSX.utils.decode_cell(addr);
    } catch (e) {
      continue;
    }
    let arr = rowCells.get(cell.r);
    if (!arr) {
      arr = [];
      rowCells.set(cell.r, arr);
    }
    arr[cell.c] = getCellValue(sheet[addr]);
  }
  return { bounds, rowCells };
}

function readRowCellsFromIndex(rowCells, rowIndex, colStart, colEnd) {
  const rowArr = rowCells.get(rowIndex);
  const out = [];
  for (let c = colStart; c <= colEnd; c++) {
    out.push(rowArr && rowArr[c] != null ? rowArr[c] : '');
  }
  return out;
}

function readRowCells(sheet, rowIndex, colStart, colEnd, rowCells) {
  if (rowCells) return readRowCellsFromIndex(rowCells, rowIndex, colStart, colEnd);
  const out = [];
  for (let c = colStart; c <= colEnd; c++) {
    out.push(getCellValue(sheet[XLSX.utils.encode_cell({ r: rowIndex, c })]));
  }
  return out;
}

function countIndexedDataRows(rowCells, headerRowIndex, colStart, colEnd) {
  let n = 0;
  for (const r of rowCells.keys()) {
    if (r <= headerRowIndex) continue;
    if (rowValuesNonEmpty(readRowCellsFromIndex(rowCells, r, colStart, colEnd))) n++;
  }
  return n;
}

function scoreTrainingSheet(sheet, defaultHeaderRow) {
  expandSheetRef(sheet);
  const bounds = getSheetBounds(sheet);
  if (bounds.e.r < bounds.s.r) return null;
  const scanEnd = Math.min(bounds.e.r, bounds.s.r + HEADER_SCAN_MAX_ROWS - 1);
  let headerRowIndex = -1;
  for (let r = bounds.s.r; r <= scanEnd; r++) {
    if (tryHeaderRow(sheet, bounds, r)) {
      headerRowIndex = r;
      break;
    }
  }
  if (headerRowIndex < 0 && defaultHeaderRow <= bounds.e.r) {
    if (tryHeaderRow(sheet, bounds, defaultHeaderRow)) headerRowIndex = defaultHeaderRow;
  }
  if (headerRowIndex < 0 && defaultHeaderRow !== DEFAULT_LMS_HEADER_ROW) {
    if (tryHeaderRow(sheet, bounds, DEFAULT_LMS_HEADER_ROW)) headerRowIndex = DEFAULT_LMS_HEADER_ROW;
  }
  if (headerRowIndex < 0) return null;
  const estimatedDataRows = Math.max(0, bounds.e.r - headerRowIndex);
  return {
    headerRowIndex,
    estimatedDataRows,
    bounds,
    cellCount: bounds.cellCount || 0,
    maxRowExcel: bounds.e.r + 1,
  };
}

function pickTrainingWorkbookSheet(wb, defaultHeaderRow) {
  const names = wb.SheetNames || [];
  if (!names.length) return { sheet: null, sheetName: '', sheetNames: [] };
  let bestName = names[0];
  let best = null;
  let bestScore = -1;
  const scores = [];
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const sheet = wb.Sheets[name];
    if (!sheet) continue;
    const s = scoreTrainingSheet(sheet, defaultHeaderRow);
    if (!s) {
      scores.push(name + ': no LMS header');
      continue;
    }
    const score = s.estimatedDataRows * 1e6 + s.cellCount;
    scores.push(
      name + ': ~' + s.estimatedDataRows.toLocaleString() + ' data rows (max row ' +
      s.maxRowExcel + ', ' + s.cellCount.toLocaleString() + ' cells)'
    );
    if (score > bestScore) {
      bestScore = score;
      bestName = name;
      best = s;
    }
  }
  if (!best) {
    expandSheetRef(wb.Sheets[bestName]);
    return {
      sheet: wb.Sheets[bestName],
      sheetName: bestName,
      sheetNames: names,
      pickNote: 'No sheet had LMS headers; using first tab: ' + bestName,
      sheetScores: scores,
    };
  }
  return {
    sheet: wb.Sheets[bestName],
    sheetName: bestName,
    sheetNames: names,
    pickNote: names.length > 1
      ? 'Using tab "' + bestName + '" (' + best.estimatedDataRows.toLocaleString() + ' estimated data rows)'
      : 'Single worksheet: ' + bestName,
    sheetScores: scores,
    parseMeta: best,
  };
}

function scoreBaseSheet(sheet) {
  expandSheetRef(sheet);
  const bounds = getSheetBounds(sheet);
  if (bounds.e.r < bounds.s.r) return null;
  const scanEnd = Math.min(bounds.e.r, bounds.s.r + HEADER_SCAN_MAX_ROWS - 1);
  let headerRowIndex = -1;
  for (let r = bounds.s.r; r <= scanEnd; r++) {
    const rowVals = readRowCells(sheet, r, bounds.s.c, bounds.e.c).map(v => String(v ?? '').trim());
    if (rowLooksLikeBaseUserHeader(rowVals)) {
      headerRowIndex = r;
      break;
    }
  }
  if (headerRowIndex < 0) headerRowIndex = bounds.s.r;
  return {
    estimatedDataRows: Math.max(0, bounds.e.r - headerRowIndex),
    cellCount: bounds.cellCount || 0,
    maxRowExcel: bounds.e.r + 1,
  };
}

function pickBaseWorkbookSheet(wb) {
  const names = wb.SheetNames || [];
  if (!names.length) return { sheet: null, sheetName: '', sheetNames: [] };
  let bestName = names[0];
  let bestScore = -1;
  const scores = [];
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    const sheet = wb.Sheets[name];
    if (!sheet) continue;
    const s = scoreBaseSheet(sheet);
    if (!s) continue;
    const score = s.estimatedDataRows * 1e6 + s.cellCount;
    scores.push(name + ': ~' + s.estimatedDataRows.toLocaleString() + ' data rows');
    if (score > bestScore) {
      bestScore = score;
      bestName = name;
    }
  }
  return {
    sheet: wb.Sheets[bestName],
    sheetName: bestName,
    sheetNames: names,
    pickNote: names.length > 1 ? 'Using tab "' + bestName + '"' : bestName,
    sheetScores: scores,
  };
}

function normHeaderCell(v) {
  return String(v ?? '')
    .replace(/^\uFEFF/, '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function rowValuesNonEmpty(values) {
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (v !== '' && v != null && v !== undefined) return true;
  }
  return false;
}

function buildRowObject(headers, values) {
  const row = {};
  const n = Math.max(headers.length, values.length);
  for (let i = 0; i < n; i++) {
    const key = headers[i] || ('__col_' + i);
    row[key] = i < values.length ? values[i] : '';
  }
  return row;
}

function findEmployeeStatusColumn(headers) {
  return findColumn(headers, ['Employee Status'], /employee\s*status/i);
}

function isTerminatedRow(row, statusCol) {
  if (!statusCol) return false;
  const s = norm(String(cellVal(row, statusCol) ?? ''));
  return s === 'terminated' || /^terminated\b/.test(s) || /\bterminated\b/.test(s);
}

function filterOutTerminated(rows, headers) {
  const statusCol = findEmployeeStatusColumn(headers);
  if (!statusCol) return { rows, removed: 0, terminated: [] };
  const active = [];
  const terminated = [];
  for (let i = 0; i < rows.length; i++) {
    if (isTerminatedRow(rows[i], statusCol)) terminated.push(rows[i]);
    else active.push(rows[i]);
  }
  return { rows: active, removed: terminated.length, terminated };
}

function validateTrainingSheet(data, label) {
  const cols = resolveToolColumns(data.headers);
  if (!cols.id || !cols.email || !cols.status) {
    const preview = data.headers.filter(h => h && !String(h).startsWith('__col_')).slice(0, 25).join(' | ');
    throw new Error(
      'Missing required columns (Employee ID, Work Email Address, Transcript Status) in ' +
      label + '. Header at row ' + (data.headerRowIndex + 1) + '. Columns seen: ' +
      (preview || '(none)')
    );
  }
}

function sheetCacheKey(cacheKey) {
  return cacheKey ? cacheKey + '|parse' + SHEET_PARSE_VERSION : null;
}

async function loadSheetData(arrayBuffer, cacheKey, onProgress, options) {
  const effectiveKey = sheetCacheKey(cacheKey);
  if (effectiveKey && state.sheetCache[effectiveKey]) {
    return state.sheetCache[effectiveKey];
  }
  const wb = parseWorkbook(arrayBuffer);
  const requireTraining = options.requireTrainingHeaders !== false;
  const defaultHeaderRow = options.defaultHeaderRow != null
    ? options.defaultHeaderRow
    : DEFAULT_LMS_HEADER_ROW;
  const picked = requireTraining
    ? pickTrainingWorkbookSheet(wb, defaultHeaderRow)
    : pickBaseWorkbookSheet(wb);
  const sheet = picked.sheet;
  if (!sheet) throw new Error('Workbook has no worksheets');
  const data = await detectHeadersAsync(sheet, onProgress, 3, options);
  data.sheetName = picked.sheetName || '';
  data.sheetNames = picked.sheetNames || wb.SheetNames || [];
  data.sheetPickNote = picked.pickNote || '';
  data.sheetScores = picked.sheetScores || [];
  data.parseMeta = picked.parseMeta || null;
  if (effectiveKey) state.sheetCache[effectiveKey] = data;
  return data;
}

async function loadToolSheetData(arrayBuffer, cacheKey, onProgress, section) {
  const cfg = section && SECTIONS[section] ? SECTIONS[section] : null;
  const defaultHeaderRow = cfg && cfg.toolHeaderRow != null ? cfg.toolHeaderRow : DEFAULT_LMS_HEADER_ROW;
  return loadSheetData(arrayBuffer, cacheKey, onProgress, {
    requireTrainingHeaders: true,
    defaultHeaderRow,
    parseYieldEvery: section === 'newJoiner' ? 8000 : PARSE_YIELD_EVERY,
  });
}

async function loadBaseSheetData(arrayBuffer, cacheKey, onProgress) {
  return loadSheetData(arrayBuffer, cacheKey, onProgress, { requireTrainingHeaders: false });
}

function rowHasRequiredTrainingHeaders(rowVals) {
  let hasId = false;
  let hasStatus = false;
  let hasEmail = false;
  for (let i = 0; i < rowVals.length; i++) {
    const t = normHeaderCell(rowVals[i]);
    if (!t) continue;
    if (REQUIRED_TRAINING_HEADER_SET.has(t)) {
      if (t === 'employee id') hasId = true;
      if (t === 'transcript status') hasStatus = true;
      if (t === 'work email address') hasEmail = true;
    }
    if (!hasId && (/^employee\s*id$/i.test(t) || t === 'emp id' || t === 'employee number')) hasId = true;
    if (!hasStatus && /transcript\s*status/i.test(t)) hasStatus = true;
    if (!hasEmail && (/work\s*email/i.test(t) || t === 'email address' || t === 'work e-mail address')) {
      hasEmail = true;
    }
  }
  return hasId && hasStatus && hasEmail;
}

/** Same rules as processing — if resolver finds ID/email/status, row is the header. */
function rowQualifiesAsTrainingHeader(rowVals) {
  if (!rowVals || !rowVals.length) return false;
  if (rowHasRequiredTrainingHeaders(rowVals)) return true;
  const headers = normalizeHeaderRow(rowVals.map(v => String(v ?? '').trim()));
  const cols = resolveToolColumns(headers);
  return !!(cols.id && cols.email && cols.status);
}

function tryHeaderRow(sheet, bounds, rowIndex, rowCells) {
  const rowVals = readRowCells(sheet, rowIndex, bounds.s.c, bounds.e.c, rowCells)
    .map(v => String(v ?? '').trim());
  if (!rowQualifiesAsTrainingHeader(rowVals)) return null;
  return {
    headers: normalizeHeaderRow(rowVals),
    headerRowIndex: rowIndex,
  };
}

function rowLooksLikeBaseUserHeader(rowVals) {
  let hasEmail = false;
  let hasId = false;
  for (let i = 0; i < rowVals.length; i++) {
    const t = String(rowVals[i] || '').trim().toLowerCase();
    if (!t) continue;
    if (/email/.test(t)) hasEmail = true;
    if (/employee\s*id|emp\s*id|local\s*employee|global\s*employee/.test(t)) hasId = true;
  }
  return hasEmail && hasId;
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
 * Row-by-row parse (memory-safe for large LMS exports).
 * Excludes Employee Status = Terminated. Yields to keep UI responsive.
 */
async function detectHeadersAsync(sheet, onProgress, minCols = 3, options = {}) {
  const requireTraining = options.requireTrainingHeaders !== false;
  const defaultHeaderRow = options.defaultHeaderRow != null
    ? options.defaultHeaderRow
    : DEFAULT_LMS_HEADER_ROW;
  const { bounds, rowCells } = buildSheetRowIndex(sheet);
  if (bounds.e.r < bounds.s.r) {
    return {
      headers: [], headerRowIndex: 0, rows: [], terminatedRemoved: 0,
      parseDiagnostics: { cellCount: 0, maxRowExcel: 0, indexedDataRows: 0 },
    };
  }

  const scanEnd = Math.min(bounds.e.r, bounds.s.r + HEADER_SCAN_MAX_ROWS - 1);
  let headerRowIndex = bounds.s.r;
  let headers = [];

  for (let r = bounds.s.r; r <= scanEnd; r++) {
    const hit = tryHeaderRow(sheet, bounds, r, rowCells);
    if (hit) {
      headers = hit.headers;
      headerRowIndex = hit.headerRowIndex;
      break;
    }
  }

  if (!headers.length && requireTraining && defaultHeaderRow <= bounds.e.r) {
    const hit = tryHeaderRow(sheet, bounds, defaultHeaderRow, rowCells);
    if (hit) {
      headers = hit.headers;
      headerRowIndex = hit.headerRowIndex;
    }
  }

  if (!headers.length && requireTraining && defaultHeaderRow !== DEFAULT_LMS_HEADER_ROW) {
    const hit = tryHeaderRow(sheet, bounds, DEFAULT_LMS_HEADER_ROW, rowCells);
    if (hit) {
      headers = hit.headers;
      headerRowIndex = hit.headerRowIndex;
    }
  }

  if (!headers.length && !requireTraining) {
    for (let r = bounds.s.r; r <= scanEnd; r++) {
      const rowVals = readRowCells(sheet, r, bounds.s.c, bounds.e.c, rowCells)
        .map(v => String(v ?? '').trim());
      if (rowLooksLikeBaseUserHeader(rowVals)) {
        headers = normalizeHeaderRow(rowVals);
        headerRowIndex = r;
        break;
      }
    }
  }

  if (!headers.length) {
    for (let r = bounds.s.r; r <= scanEnd; r++) {
      const rowVals = readRowCells(sheet, r, bounds.s.c, bounds.e.c, rowCells)
        .map(v => String(v ?? '').trim());
      const nonEmpty = rowVals.filter(v => v !== '').length;
      if (nonEmpty >= minCols) {
        headers = normalizeHeaderRow(rowVals);
        headerRowIndex = r;
        break;
      }
    }
  }

  if (!headers.length) {
    const hint = requireTraining
      ? 'Employee ID, Transcript Status, and Work Email Address'
      : 'employee/email columns';
    const excelDefault = defaultHeaderRow + 1;
    throw new Error(
      'Could not find a header row with ' + hint + ' (scanned rows 1–' + (scanEnd + 1) +
      ', also tried row ' + excelDefault + ').'
    );
  }

  const rows = [];
  const dataStart = headerRowIndex + 1;
  const yieldEvery = options.parseYieldEvery || PARSE_YIELD_EVERY;
  const indexedDataRows = countIndexedDataRows(rowCells, headerRowIndex, bounds.s.c, bounds.e.c);

  for (let r = dataStart; r <= bounds.e.r; r++) {
    const values = readRowCells(sheet, r, bounds.s.c, bounds.e.c, rowCells);
    if (!rowValuesNonEmpty(values)) continue;
    rows.push(buildRowObject(headers, values));
    if (rows.length % yieldEvery === 0) {
      if (onProgress) onProgress(rows.length);
      await yieldToMain();
    }
  }

  const { rows: activeRows, removed, terminated } = filterOutTerminated(rows, headers);
  return {
    headers,
    headerRowIndex,
    rows: activeRows,
    terminatedRemoved: removed,
    terminatedRows: terminated,
    parseDiagnostics: {
      declaredRef: bounds.declaredRef,
      cellCount: bounds.cellCount,
      maxRowExcel: bounds.e.r + 1,
      headerRowExcel: headerRowIndex + 1,
      indexedDataRows,
      parsedRows: rows.length,
    },
  };
}

/** @deprecated use detectHeadersAsync — sync fallback for small sheets */
function detectHeaders(sheet, minCols = 3) {
  const bounds = getSheetBounds(sheet);
  if (bounds.e.r < bounds.s.r) return { headers: [], headerRowIndex: 0, rows: [], terminatedRemoved: 0 };

  const scanEnd = Math.min(bounds.e.r, bounds.s.r + HEADER_SCAN_MAX_ROWS - 1);
  let headerRowIndex = bounds.s.r;
  let headers = [];

  for (let r = bounds.s.r; r <= scanEnd; r++) {
    const hit = tryHeaderRow(sheet, bounds, r);
    if (hit) {
      headers = hit.headers;
      headerRowIndex = hit.headerRowIndex;
      break;
    }
  }

  if (!headers.length) {
    const hit = tryHeaderRow(sheet, bounds, DEFAULT_LMS_HEADER_ROW);
    if (hit) {
      headers = hit.headers;
      headerRowIndex = hit.headerRowIndex;
    }
  }

  if (!headers.length) {
    for (let r = bounds.s.r; r <= scanEnd; r++) {
      const rowVals = readRowCells(sheet, r, bounds.s.c, bounds.e.c).map(v => String(v ?? '').trim());
      if (rowVals.filter(v => v !== '').length >= minCols) {
        headers = normalizeHeaderRow(rowVals);
        headerRowIndex = r;
        break;
      }
    }
  }

  const rows = [];
  for (let r = headerRowIndex + 1; r <= bounds.e.r; r++) {
    const values = readRowCells(sheet, r, bounds.s.c, bounds.e.c);
    if (!rowValuesNonEmpty(values)) continue;
    rows.push(buildRowObject(headers, values));
  }

  const { rows: activeRows, removed, terminated } = filterOutTerminated(rows, headers);
  return { headers, headerRowIndex, rows: activeRows, terminatedRemoved: removed, terminatedRows: terminated };
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
  let t = String(v).replace(/\s/g, '').toLowerCase();
  // Excel numeric IDs often arrive as "12345.0" — strip the trailing .0 (mirrors Python _norm_id).
  if (/^\d+\.0$/.test(t)) t = t.slice(0, -2);
  return t;
}

/**
 * Format LMS date values to dd/mm/yyyy (Training Start Date, Completed Date, assignment date, etc.).
 */
function fmtDate(v) {
  if (v === null || v === undefined || v === '') return '';
  const parts = parseCalendarParts(v);
  if (parts) return formatCalendarPartsGb(parts);
  return String(v).trim();
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
 */
function processToolOnly(toolData, assignDate, options = {}) {
  const cols = resolveToolColumns(toolData.headers);
  const zoneCol = options.zoneCol || cols.zone;

  const output = [];

  for (const toolRow of toolData.rows) {
    const empId = cellVal(toolRow, cols.id);
    const email = cellVal(toolRow, cols.email);

    const rawStatus = String(cellVal(toolRow, cols.status) || '').trim();
    const completedDate = fmtDate(cellVal(toolRow, cols.complete));
    const startDate = fmtDate(cellVal(toolRow, cols.start)) ||
      (norm(rawStatus) !== 'completed' && assignDate ? assignDate : '');
    const zone = zoneCol ? cellVal(toolRow, zoneCol) : '';
    const status = deriveTranscriptStatus(rawStatus, completedDate);

    output.push(makeOutputRow({
      empId, email, zone, status, startDate, completedDate,
      mapZone: options.mapZone !== false,
    }));
  }

  return output;
}

function processAppSec(toolData, assignDate) {
  return processToolOnly(toolData, assignDate);
}

/**
 * Growth Group: only rows where Macro Entity Level 3 BU Description = GLOBAL GROWTH.
 */
function processGrowth(toolData, assignDate) {
  const cfg = SECTIONS.growth;
  const buCol = findMacroEntityLevel3BUColumn(toolData.headers);
  if (!buCol) {
    throw new Error(
      'Growth tool file is missing column "Macro entity level three BU Description".'
    );
  }

  const expected = cfg.buFilterValue || 'GLOBAL GROWTH';
  const filteredRows = toolData.rows.filter(r => rowMatchesBuFilter(r, buCol, expected));
  const excluded = toolData.rows.length - filteredRows.length;

  const output = processToolOnly(
    { headers: toolData.headers, rows: filteredRows },
    assignDate,
    { zoneCol: buCol, mapZone: false }
  );

  return { rows: output, excluded, totalInFile: toolData.rows.length };
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

function resolveBaseIdEmailCols(baseData, cfg) {
  let baseIdCol = cfg.baseIdCol;
  let baseEmailCol = cfg.baseEmailCol;
  if (!baseIdCol || !baseData.headers.includes(baseIdCol)) {
    baseIdCol = baseData.headers.find(h => /employee.?id|emp.?id|local.?id/i.test(h)) || null;
  }
  if (!baseEmailCol || !baseData.headers.includes(baseEmailCol)) {
    baseEmailCol = baseData.headers.find(h => /email/i.test(h)) || null;
  }
  return { baseIdCol, baseEmailCol };
}

/** Phished / tracking userbase: Employee Email + Global Employee ID + Local Employee ID. */
function resolveUserbaseMatchCols(baseData, cfg) {
  const emailCol = findColumn(
    baseData.headers,
    [cfg.baseEmailCol, 'Employee Email', 'Email - Primary Work'],
    /employee\s*email/i,
    /^email$/i,
    /email/i
  );
  const globalIdCol = findColumn(
    baseData.headers,
    [cfg.baseGlobalIdCol, 'Global Employee ID'],
    /global\s*employee\s*id/i
  );
  const localIdCol = findColumn(
    baseData.headers,
    [cfg.baseLocalIdCol, 'Local Employee ID'],
    /local\s*employee\s*id/i,
    /local.?id/i
  );
  let empIdCol = cfg.baseIdCol;
  if (!empIdCol || !baseData.headers.includes(empIdCol)) {
    empIdCol = findColumn(
      baseData.headers,
      ['Emp ID', 'Employee ID'],
      /emp.?id/i,
      /^employee\s*id$/i
    );
  }
  return { emailCol, globalIdCol, localIdCol, empIdCol };
}

function getUserbaseMatchValues(baseRow, matchCols) {
  return {
    email: matchCols.emailCol ? String(cellVal(baseRow, matchCols.emailCol) || '').trim() : '',
    globalId: matchCols.globalIdCol ? String(cellVal(baseRow, matchCols.globalIdCol) || '').trim() : '',
    localId: matchCols.localIdCol ? String(cellVal(baseRow, matchCols.localIdCol) || '').trim() : '',
    empId: matchCols.empIdCol ? String(cellVal(baseRow, matchCols.empIdCol) || '').trim() : '',
  };
}

function userbaseDedupeKey(values) {
  return norm(values.email) ||
    normId(values.globalId) ||
    normId(values.localId) ||
    normId(values.empId) ||
    '';
}

function findToolMatchMulti(map, values) {
  if (values.email) {
    const byEmail = map.get(norm(values.email));
    if (byEmail) return byEmail;
  }
  for (const id of [values.globalId, values.localId, values.empId]) {
    if (!id) continue;
    const hit = map.get(normId(id));
    if (hit) return hit;
  }
  return null;
}

function buildEnrichedUserbaseYearTabs(enrichedRows, startDateCol) {
  const yearTabs = { '2024': [], '2025': [], '2026': [] };
  for (const row of enrichedRows) {
    const bucket = yearBucketFromStartDate(row[startDateCol] || '');
    yearTabs[bucket].push(row);
  }
  return yearTabs;
}

/**
 * Phishing Tracking userbase: userbase is the master list — every row kept in download and metrics.
 * Training fields mapped from tool upload; no tool match → Not Found.
 */
/** Latest tool row by Training Start Date (newest wins); null if empty. */
function pickLatestToolRow(rows, cols) {
  if (!rows || !rows.length) return null;
  return rows.slice().sort((a, b) => {
    const da = new Date((fmtDate(cellVal(a, cols.start)) || '').split('/').reverse().join('-'));
    const db = new Date((fmtDate(cellVal(b, cols.start)) || '').split('/').reverse().join('-'));
    const ta = isNaN(da) ? -Infinity : da.getTime();
    const tb = isNaN(db) ? -Infinity : db.getTime();
    return tb - ta;
  })[0];
}

/** A tool row whose derived status is Completed (has a completion date or 'completed' status). */
function rowIsCompleted(row, cols) {
  const completedDate = fmtDate(cellVal(row, cols.complete));
  const rawStatus = String(cellVal(row, cols.status) || '').trim();
  return deriveTranscriptStatus(rawStatus, completedDate) === 'Completed';
}

/**
 * Completed-wins: if the user completed on ANY transcript row, return the latest such row
 * (by completion date, then start date); null if no row is completed. Keeps a user who finished
 * the training Completed even when a newer re-assignment row is Not Started/Terminated.
 */
function pickCompletedToolRow(rows, cols) {
  if (!rows || !rows.length) return null;
  const completedRows = rows.filter((r) => rowIsCompleted(r, cols));
  if (!completedRows.length) return null;
  const ts = (v) => {
    const d = new Date((fmtDate(v) || '').split('/').reverse().join('-'));
    return isNaN(d) ? -Infinity : d.getTime();
  };
  return completedRows.slice().sort((a, b) => {
    const ca = ts(cellVal(a, cols.complete));
    const cb = ts(cellVal(b, cols.complete));
    if (cb !== ca) return cb - ca;
    return ts(cellVal(b, cols.start)) - ts(cellVal(a, cols.start));
  })[0];
}

/** Collect all tool rows matching a userbase user across email + global/local/emp ids (deduped). */
function gatherToolRowsMulti(map, values) {
  const out = [];
  const seen = new Set();
  const push = (key) => {
    const list = key && map.get(key);
    if (!list) return;
    for (const r of list) { if (!seen.has(r)) { seen.add(r); out.push(r); } }
  };
  if (values.email) push(norm(values.email));
  for (const id of [values.globalId, values.localId, values.empId]) {
    if (id) push(normId(id));
  }
  return out;
}

function processUserbaseTracking(baseData, toolData, assignDate, inScopeFn, cfg) {
  const sectionCfg = cfg || SECTIONS.phishingNormal;
  const matchCols = resolveUserbaseMatchCols(baseData, sectionCfg);

  let toolCols = null;
  let toolMap = new Map();
  if (toolData && toolData.rows) {
    toolCols = resolveToolColumns(toolData.headers);
    toolMap = buildLookupMap(toolData.rows, toolCols.email, toolCols.id);
  }

  const enrichedHeaders = appendColumnsAfterBase(baseData.headers, TRACKING_USERBASE_APPEND_COLS);
  const baseZoneCol = findColumn(
    baseData.headers,
    ['Zone', 'Macro Entity Level 2 (Zone)'],
    /^zone$/i,
    /macro.?entity.?level.?2/i
  );

  const enrichedRows = [];
  const statsRows = [];
  let completed = 0;
  let notCompleted = 0;
  let notFound = 0;
  let terminated = 0;
  let matched = 0;

  for (const baseRow of baseData.rows) {
    const enriched = {};
    for (const h of enrichedHeaders) enriched[h] = '';
    for (const h of baseData.headers) {
      if (Object.prototype.hasOwnProperty.call(baseRow, h)) {
        enriched[h] = baseRow[h];
      }
    }

    if (!inScopeFn(baseRow)) {
      enrichedRows.push(enriched);
      continue;
    }

    const matchValues = getUserbaseMatchValues(baseRow, matchCols);
    const matchKey = userbaseDedupeKey(matchValues);

    if (matchKey) {
      let toolRow = null;
      if (toolCols) {
        toolRow = findToolMatchMulti(toolMap, matchValues);
      }
      let status;
      let startDate = '';
      let completedDate = '';
      let zone = baseZoneCol ? cellVal(baseRow, baseZoneCol) : '';

      if (toolRow && toolCols) {
        const rawStatus = String(cellVal(toolRow, toolCols.status) || '').trim();
        startDate = sectionCfg.yearDateFromAssignment
          ? (fmtDate(assignDate) || '')
          : (fmtDate(cellVal(toolRow, toolCols.start)) || assignDate || '');
        if (!zone) zone = cellVal(toolRow, toolCols.zone);
        completedDate = fmtDate(cellVal(toolRow, toolCols.complete));
        status = deriveTranscriptStatus(rawStatus, completedDate);
      } else {
        status = 'Not Found';
      }

      enriched['Training Start Date'] = startDate;
      enriched['Training completion date'] = completedDate;
      enriched['training completion status'] = status;
      enriched._zone = mapMacroZone(zone) || zone || 'Unknown';
    }

    enrichedRows.push(enriched);
    statsRows.push(enriched);

    const o = userbaseTrackingOutcome(enriched);
    if (o === 'completed') completed++;
    else if (o === 'notFound') notFound++;
    else if (o === 'terminated') terminated++;
    else notCompleted++;
    if (o !== 'notFound') matched++;
  }

  const yearTabs = buildEnrichedUserbaseYearTabs(statsRows, 'Training Start Date');

  return {
    headers: enrichedHeaders,
    enriched: enrichedRows,
    statsRows,
    '2024': yearTabs['2024'],
    '2025': yearTabs['2025'],
    '2026': yearTabs['2026'],
    stats: {
      total: statsRows.length,
      completed,
      notCompleted,
      notFound,
      terminated,
      matched,
      userbaseRows: baseData.rows.length,
    },
  };
}

function processPhishingNormal(baseData, toolData, assignDate) {
  return processUserbaseTracking(baseData, toolData, assignDate, rowInPhishingNormalScope);
}

function processBand4(baseData, toolData, assignDate) {
  return processUserbaseTracking(baseData, toolData, assignDate, rowInBand4Scope, SECTIONS.band4);
}

function processSection(section, baseData, toolData, assignDate, filterFn) {
  const cfg = SECTIONS[section];

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

  for (const toolRow of toolData.rows) {
    const empId = cellVal(toolRow, cols.id);
    const email = cellVal(toolRow, cols.email);

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
 * BSC: enrich userbase — match tool file by Emp ID + email; append training columns
 * after the last filled userbase column (original start date columns unchanged).
 */
function processBSC(baseData, toolData, assignDate) {
  const cfg = SECTIONS.bsc;
  let baseIdCol = cfg.baseIdCol;
  let baseEmailCol = cfg.baseEmailCol;
  if (!baseIdCol || !baseData.headers.includes(baseIdCol)) {
    baseIdCol = baseData.headers.find(h => /emp.?id/i.test(h)) || null;
  }
  if (!baseEmailCol || !baseData.headers.includes(baseEmailCol)) {
    baseEmailCol = baseData.headers.find(h => /email/i.test(h)) || null;
  }
  if (!baseIdCol && !baseEmailCol) {
    throw new Error('Userbase must include Emp ID and/or Email - Primary Work columns.');
  }

  const cols = resolveToolColumns(toolData.headers);
  // BSC keeps terminated tool rows so matched terminated users can be flagged (not dropped).
  const toolRowsAll = toolData.rows.concat(toolData.terminatedRows || []);
  const toolMap = buildToolRowsMapByUser(toolRowsAll, cols);
  const enrichedHeaders = appendBscColumnsAfterBase(baseData.headers);
  const baseZoneCol = findColumn(
    baseData.headers,
    ['Zone', 'Macro Entity Level 2 (Zone)'],
    /^zone$/i,
    /macro.?entity.?level.?2/i
  );
  const baseCenterCol = findBaseCenterColumn(baseData.headers);

  const enrichedRows = [];
  const statsRows = [];
  let matched = 0;

  for (const baseRow of baseData.rows) {
    const baseEmail = baseEmailCol ? String(cellVal(baseRow, baseEmailCol) || '').trim() : '';
    const baseId = baseIdCol ? String(cellVal(baseRow, baseIdCol) || '').trim() : '';
    const matchKey = norm(baseEmail) || normId(baseId);

    const enriched = {};
    for (const h of enrichedHeaders) enriched[h] = '';
    for (const h of baseData.headers) {
      if (Object.prototype.hasOwnProperty.call(baseRow, h)) {
        enriched[h] = baseRow[h];
      }
    }

    if (matchKey) {
      const toolRows = getToolRowsForBaseUser(toolMap, baseEmail, baseId, cols);
      const training = resolveBscTrainingFromToolRows(toolRows, cols, assignDate);

      let zone = baseZoneCol ? cellVal(baseRow, baseZoneCol) : '';
      if (!zone && toolRows.length) {
        zone = cellVal(toolRows[0], cols.zone);
      }
      const center = baseCenterCol ? String(cellVal(baseRow, baseCenterCol) || '').trim() : '';

      enriched['start date_extracted'] = training.startDateExtracted;
      enriched['Training completion date'] = training.completedDate;
      enriched['training completion status'] = training.status;
      enriched._zone = mapMacroZone(zone) || zone || 'Unknown';
      enriched._center = center || 'Unknown';
    }

    enrichedRows.push(enriched);
    statsRows.push(enriched);
    if (norm(enriched['training completion status']) !== 'not found') matched++;
  }

  const { years, byYear } = buildBscDynamicYearTabs(statsRows);

  return {
    headers: enrichedHeaders,
    enriched: enrichedRows,
    statsRows,
    years,          // sorted list of actual years present, e.g. ['2022','2024','2025','2026']
    ...byYear,      // result[year] -> rows, consumed by the generic year-tab renderers
    matched,
    total: statsRows.length,
    userbaseRows: baseData.rows.length,
  };
}

/**
 * Build lightweight rows for New Joiner (hire date normalized to dd/mm/yyyy, years 2024–2026 only).
 */
function analyzeNewJoinerRows(toolData, hireCol) {
  const audit = {
    emptyHire: 0,
    parseFailed: 0,
    outOfScopeYear: 0,
    outOfScopeByYear: {},
    inScopeByYear: { '2024': 0, '2025': 0, '2026': 0 },
  };

  for (let i = 0; i < toolData.rows.length; i++) {
    const display = hireDateCellToString(cellVal(toolData.rows[i], hireCol));
    if (!display) {
      audit.emptyHire++;
      continue;
    }
    const year = extractHireYearFromString(display);
    if (!year) {
      audit.parseFailed++;
      continue;
    }
    if (year === 2024 || year === 2025 || year === 2026) {
      audit.inScopeByYear[String(year)]++;
    } else {
      audit.outOfScopeYear++;
      const yk = String(year);
      audit.outOfScopeByYear[yk] = (audit.outOfScopeByYear[yk] || 0) + 1;
    }
  }
  return audit;
}

function prepareNewJoinerInput(toolData) {
  const cols = resolveToolColumns(toolData.headers);
  const hireCol = findNewJoinerOriginalHireDateColumn(toolData.headers);
  if (!hireCol) {
    throw new Error(
      'Tool file must include "Original Hire Date" (or "New J orginal hire date").'
    );
  }
  const buCol = findMacroEntityLevel3BUColumn(toolData.headers);
  const prepared = [];
  let excludedHireYear = 0;
  const hireAudit = analyzeNewJoinerRows(toolData, hireCol);

  for (let i = 0; i < toolData.rows.length; i++) {
    const row = toolData.rows[i];
    const hire = normalizeOriginalHireDate(cellVal(row, hireCol));
    if (!hire) {
      excludedHireYear++;
      continue;
    }

    const email = String(cellVal(row, cols.email) || '').trim();
    const empIdRaw = cellVal(row, cols.id);
    const empId = empIdRaw != null && empIdRaw !== '' ? String(empIdRaw).trim() : '';
    const rawStatus = String(cellVal(row, cols.status) || '').trim();
    const completedDate = fmtDate(cellVal(row, cols.complete));
    const status = deriveTranscriptStatus(rawStatus, completedDate);
    const zone = resolveNewJoinerZone(row, cols, buCol);

    prepared.push({
      email,
      empId,
      zone,
      status,
      startDate: hire.formatted,
      completedDate,
      hireYearTab: hire.yearTab,
    });
  }

  return {
    prepared,
    hireCol,
    excludedHireYear,
    hireAudit,
    inputRows: toolData.rows.length,
    terminatedRemoved: toolData.terminatedRemoved || 0,
    employeeStatusCol: findEmployeeStatusColumn(toolData.headers),
  };
}

function processNewJoinerFromPrepared(prepared) {
  const result = { '2024': [], '2025': [], '2026': [] };

  for (let i = 0; i < prepared.length; i++) {
    const p = prepared[i];
    const tab = p.hireYearTab;
    if (!tab || !Object.prototype.hasOwnProperty.call(result, tab)) continue;
    const outRow = {
      'Employee ID': p.empId,
      'Work Email Address': p.email,
      'Zone': p.zone,
      'Transcript Status': p.status,
      'Training Start Date': p.startDate,
      'Transcript Completed Date': p.completedDate,
      _hireDate: p.startDate,
    };
    result[tab].push(outRow);
  }
  return { result };
}

function buildNewJoinerResult(toolData) {
  const prep = prepareNewJoinerInput(toolData);
  const { result } = processNewJoinerFromPrepared(prep.prepared);
  result.stats = {
    inputRows: prep.inputRows,
    terminatedRemoved: prep.terminatedRemoved,
    employeeStatusCol: prep.employeeStatusCol,
    excludedHireYear: prep.excludedHireYear,
    hireAudit: prep.hireAudit,
    prepared: prep.prepared.length,
    hireCol: prep.hireCol,
    byYear: {
      '2024': result['2024'].length,
      '2025': result['2025'].length,
      '2026': result['2026'].length,
    },
  };
  return result;
}

/**
 * New Joiner: Original Hire Date as text; year tab if cell contains 2024, 2025, or 2026.
 * Terminated rows removed at parse. No Web Workers (unreliable on file:// / local open).
 */
async function processNewJoiner(toolData, assignDate, onProgress) {
  if (onProgress) onProgress(0);
  await yieldToMain();
  const result = await runHeavyWork(() => buildNewJoinerResult(toolData));
  if (onProgress && result.stats) onProgress(result.stats.inputRows);
  return result;
}

function logNewJoinerReconciliation(section, stats) {
  if (!stats) return;
  appendDebugLog(section, '── Reconciliation (vs Excel manual filter) ──');
  appendDebugLog(section, 'Tool rows parsed: ' + stats.inputRows.toLocaleString());
  appendDebugLog(
    section,
    'Terminated removed: ' + stats.terminatedRemoved.toLocaleString() +
      (stats.employeeStatusCol ? ' (column: ' + stats.employeeStatusCol + ')' : ' (no Employee Status column found)')
  );
  if (stats.hireAudit) {
    const a = stats.hireAudit;
    appendDebugLog(section, 'Empty Original Hire Date: ' + a.emptyHire.toLocaleString());
    appendDebugLog(section, 'No 2024/2025/2026 in hire text: ' + a.parseFailed.toLocaleString());
    appendDebugLog(section, 'Hire year outside 2024–2026: ' + a.outOfScopeYear.toLocaleString());
    const scopeKeys = Object.keys(a.outOfScopeByYear || {}).sort();
    if (scopeKeys.length) {
      appendDebugLog(
        section,
        '  Out-of-scope years: ' + scopeKeys.map(y => y + '=' + a.outOfScopeByYear[y]).join(', ')
      );
    }
    appendDebugLog(
      section,
      'Rows with in-scope hire years: 2024=' + a.inScopeByYear['2024'].toLocaleString() +
        ', 2025=' + a.inScopeByYear['2025'].toLocaleString() +
        ', 2026=' + a.inScopeByYear['2026'].toLocaleString()
    );
  }
  appendDebugLog(
    section,
    'Output rows by year: 2024=' + stats.byYear['2024'].toLocaleString() +
      ', 2025=' + stats.byYear['2025'].toLocaleString() +
      ', 2026=' + stats.byYear['2026'].toLocaleString()
  );
  appendDebugLog(section, 'Zone labels are mapped (e.g. EUR, MAZ, Growth); not raw LMS zone text.');
  appendDebugLog(section, 'Metrics / zone tiles / table use the active year tab only.');
}

// ============================================================
// CALCULATE SECTION
// ============================================================
async function calcSection(section) {
  const cfg = SECTIONS[section];
  const baseKey = section + '-base';
  const toolKey = section + '-tool';

  if (!state.files[toolKey]) {
    appendDebugLog(section, 'ERROR: Tool Output Excel not uploaded.');
    toast('error', 'Please upload the Tool Output Excel file first');
    return;
  }

  const btn = document.getElementById('btn-calc-' + section);
  const origText = btn.innerHTML;
  btn.innerHTML = '<span class="spinner"></span> Reading file…';
  btn.disabled = true;

  await sleep(80);

  const setProgress = (msg) => {
    btn.innerHTML = '<span class="spinner"></span> ' + msg;
  };

  try {
    const assignDate = document.getElementById('date-' + section)?.value || '';
    clearDebugLog(section);
    appendDebugLog(section, 'Starting calculate for section: ' + section + (assignDate ? ' | Assignment Date: ' + assignDate : ''));
    const baseAB = state.files[baseKey] || null;
    const toolAB = state.files[toolKey];
    state.renderedTables = {};

    let terminatedRemoved = 0;

    setProgress('Parsing tool export…');
    const toolData = await loadToolSheetData(toolAB, toolKey, (n) => {
      setProgress('Parsing ' + n.toLocaleString() + ' rows…');
    }, section);
    validateTrainingSheet(toolData, cfg.label + ' tool file');
    logToolMappings(section, toolData);
    appendDebugLog(section, 'Tool rows parsed: ' + toolData.rows.length.toLocaleString() +
      (toolData.terminatedRemoved ? ' | terminated removed: ' + toolData.terminatedRemoved.toLocaleString() : ''));
    terminatedRemoved += toolData.terminatedRemoved || 0;

    let baseData = null;
    if (baseAB) {
      setProgress('Parsing base file…');
      baseData = await loadBaseSheetData(baseAB, baseKey, (n) => {
        setProgress('Parsing base ' + n.toLocaleString() + ' rows…');
      });
      logBaseMappings(section, baseData, cfg);
      appendDebugLog(section, 'Userbase rows in file: ' + baseData.rows.length.toLocaleString() +
        (baseData.terminatedRemoved ? ' | terminated removed: ' + baseData.terminatedRemoved.toLocaleString() : ''));
      if (section === 'phishingNormal' || section === 'band4') {
        appendDebugLog(section, 'Userbase is master list; match tool by Employee Email, Global Employee ID, Local Employee ID');
      } else if (section === 'bsc') {
        appendDebugLog(section, 'Userbase is master list; match tool by Emp ID + email');
      }
      terminatedRemoved += baseData.terminatedRemoved || 0;
    }

    setProgress('Calculating…');
    await yieldToMain();

    let rows;
    if (section === 'newJoiner') {
      logHireDateParseAudit(section, toolData);
      const hireCol = findNewJoinerOriginalHireDateColumn(toolData.headers);
      appendDebugLog(section, 'Hire-date column: "' + (hireCol || 'NOT FOUND') + '" (required: Original Hire Date)');
      appendDebugLog(section, 'Hire year: read cell as text, match 2024 / 2025 / 2026 (no strict date conversion)');
      const buCol = findMacroEntityLevel3BUColumn(toolData.headers);
      appendDebugLog(section, 'BU column (Growth): ' + (buCol || 'NOT FOUND'));
      const njResult = await processNewJoiner(toolData, assignDate, (n) => {
        setProgress('Processing ' + n.toLocaleString() + ' rows…');
      });
      state.results[section] = njResult;
      await yieldToMain();
      const activeYear = switchToYearWithMostRows('newJoiner', njResult);
      renderNewJoinerTables(njResult);
      updateNJMetrics(njResult, activeYear);
      if (njResult.stats) {
        logNewJoinerReconciliation(section, njResult.stats);
      }
      rows = njResult[activeYear] || [];
    } else if (section === 'growth') {
      const buCol = findMacroEntityLevel3BUColumn(toolData.headers);
      appendDebugLog(section, 'Growth BU column mapped: ' + (buCol || 'NOT FOUND') + ' | expected value: GLOBAL GROWTH');
      appendDebugLog(section, 'Growth Zone output uses BU column only: ' + (buCol || 'NOT FOUND'));
      const growthResult = await runHeavyWork(() => processGrowth(toolData, assignDate));
      rows = growthResult.rows;
      const yearResult = splitRowsByYear(rows);
      state.results[section] = yearResult;
      state.growthFilterStats = {
        excluded: growthResult.excluded,
        totalInFile: growthResult.totalInFile,
      };
      await yieldToMain();
      switchToYearWithMostRows(section, yearResult);
      renderYearTables(section, yearResult);
      updateMetrics(section, rows);
    } else if (section === 'appSec' || section === 'cyberOT') {
      rows = await runHeavyWork(() => processToolOnly(toolData, assignDate));
      const yearResult = splitRowsByYear(rows);
      state.results[section] = yearResult;
      await yieldToMain();
      switchToYearWithMostRows(section, yearResult);
      renderYearTables(section, yearResult);
      updateMetrics(section, rows);
    } else if (section === 'bsc') {
      if (!baseData) {
        appendDebugLog(section, 'ERROR: Base Userbase Excel not uploaded.');
        toast('error', 'Please upload the Base Userbase Excel');
        btn.innerHTML = origText;
        btn.disabled = false;
        return;
      }
      const bscResult = await runHeavyWork(() => processBSC(baseData, toolData, assignDate));
      state.results[section] = bscResult;
      rows = bscResult.statsRows;
      await yieldToMain();
      const bscBestYear = pickYearWithMostRows(bscResult.years, bscResult);
      buildBscYearTabsDom(bscResult.years, bscBestYear);
      renderYearTables(section, bscResult);
      updateYearTabCounts(section, bscResult);
      updateMetrics(section, bscResult.statsRows);
      appendDebugLog(section, 'Userbase rows (Total KPI): ' + bscResult.total.toLocaleString() +
        ' | tool matched: ' + bscResult.matched.toLocaleString());
      state.bscMatchStats = {
        matched: bscResult.matched,
        total: bscResult.total,
        userbaseRows: bscResult.userbaseRows,
      };
    } else if (section === 'band4' || section === 'phishingNormal') {
      if (!baseData) {
        appendDebugLog(section, 'ERROR: Phishing Tracking Input Excel not uploaded.');
        toast('error', 'Please upload the Phishing Tracking Input Excel');
        btn.innerHTML = origText;
        btn.disabled = false;
        return;
      }
      const trackResult = await runHeavyWork(() =>
        section === 'band4'
          ? processBand4(baseData, toolData, assignDate)
          : processPhishingNormal(baseData, toolData, assignDate)
      );
      state.results[section] = trackResult;
      rows = trackResult.statsRows;
      if (section === 'phishingNormal') {
        state.phishingNormalStats = trackResult.stats;
      } else {
        state.band4Stats = trackResult.stats;
      }
      await yieldToMain();
      switchToYearWithMostRows(section, trackResult);
      renderYearTables(section, trackResult);
      updateUserbaseTrackingMetrics(section, trackResult.statsRows);
      appendDebugLog(section, 'Userbase rows (Total KPI): ' + trackResult.stats.total.toLocaleString() +
        ' | tool matched: ' + (trackResult.stats.matched || 0).toLocaleString());
    }

    const now = new Date();
    const ts = now.toLocaleTimeString('en-GB');
    document.getElementById('lc-' + section).textContent = 'Last calculated: ' + ts;
    document.getElementById('m-' + section + '-calc').textContent = ts;
    document.getElementById('kpi-lastrun').textContent = ts;
    setSectionDownloadButtons(section, true);

    let msg = cfg.label + ' — ' + (rows ? rows.length.toLocaleString() : '0') + ' users';
    if (terminatedRemoved > 0) {
      msg += ' (' + terminatedRemoved.toLocaleString() + ' terminated excluded)';
    }
    if (section === 'growth' && state.growthFilterStats) {
      const { excluded, totalInFile } = state.growthFilterStats;
      if (excluded > 0) {
        msg += ' (' + excluded.toLocaleString() + ' of ' + totalInFile.toLocaleString() +
          ' rows excluded — not GLOBAL GROWTH)';
      }
    }
    if (section === 'bsc' && state.bscMatchStats) {
      const { matched, total, userbaseRows } = state.bscMatchStats;
      msg += ' — userbase ' + total.toLocaleString() + ' rows';
      if (userbaseRows && userbaseRows !== total) {
        msg += ' (' + userbaseRows.toLocaleString() + ' in file)';
      }
      msg += ', tool matched ' + matched.toLocaleString();
    }
    if (
      (section === 'phishingNormal' && state.phishingNormalStats) ||
      (section === 'band4' && state.band4Stats)
    ) {
      const st = section === 'phishingNormal' ? state.phishingNormalStats : state.band4Stats;
      const { completed, notCompleted, notFound } = st;
      msg += ' — completed ' + completed.toLocaleString() +
        ', not completed ' + notCompleted.toLocaleString() +
        ', unmapped ' + notFound.toLocaleString();
      if (st.terminated) msg += ', terminated ' + st.terminated.toLocaleString();
    }
    toast(rows && rows.length ? 'success' : 'info', msg);
    appendDebugLog(section, 'Calculation finished. Output rows: ' + (rows ? rows.length.toLocaleString() : '0'));
    updateGlobalKPIs();
  } catch (err) {
    console.error(err);
    const msg = err && err.message ? err.message : String(err);
    toast('error', msg);
    appendDebugLog(section, 'ERROR: ' + msg + (err && err.stack ? '\n' + err.stack : ''));
  }

  btn.innerHTML = origText;
  btn.disabled = false;
}

// ============================================================
// RENDER TABLE
// ============================================================
function getActiveYearForSection(section) {
  const panel = document.getElementById('panel-' + section);
  if (!panel) return getCurrentDashboardYearBucket();
  const activeTab = panel.querySelector('.sheet-tab.active');
  if (!activeTab) return getCurrentDashboardYearBucket();
  const dataYear = activeTab.getAttribute('data-year');
  if (dataYear) return dataYear;
  const m = (activeTab.getAttribute('onclick') || '').match(/'(\d{4})'/);
  return m ? m[1] : getCurrentDashboardYearBucket();
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
  for (const y of YEAR_TABS) {
    renderTable('newJoiner-' + y, njResult[y] || []);
    const tab = document.querySelector('#panel-newJoiner .sheet-tab[data-year="' + y + '"]');
    if (tab) {
      const n = (njResult[y] || []).length;
      tab.textContent = y + (n ? ' (' + n.toLocaleString() + ')' : '');
    }
  }
}

function renderEnrichedUserbaseTable(section, tableId, rows, headers) {
  const container = document.getElementById('tbl-' + tableId);
  if (!container) return;

  const previewCols = getEnrichedPreviewHeaders(section, headers);
  if (!rows || !rows.length) {
    container.innerHTML = '<div class="empty-state"><div class="empty-icon">🔍</div><div class="empty-title">No matching users found</div></div>';
    state.renderedTables[tableId] = true;
    return;
  }

  const displayRows = rows.slice(0, TABLE_PREVIEW_ROWS);
  const parts = ['<table><thead><tr><th>#</th>'];
  for (const col of previewCols) {
    parts.push('<th>', esc(col), '</th>');
  }
  parts.push('</tr></thead><tbody>');

  for (let i = 0; i < displayRows.length; i++) {
    const row = displayRows[i];
    parts.push('<tr><td class="mono" style="color:var(--text3)">', String(i + 1), '</td>');
    for (const col of previewCols) {
      let val = row[col];
      if (col === 'training completion status') {
        parts.push('<td>', getStatusBadge(String(val || '')), '</td>');
      } else {
        parts.push('<td class="mono">', esc(val), '</td>');
      }
    }
    parts.push('</tr>');
  }

  if (rows.length > TABLE_PREVIEW_ROWS) {
    parts.push(
      '<tr><td colspan="', String(previewCols.length + 1),
      '" style="text-align:center;color:var(--text3);padding:12px;font-style:italic">',
      'Showing ', String(TABLE_PREVIEW_ROWS), ' of ', String(rows.length),
      ' userbase rows (full row in Excel download)</td></tr>'
    );
  }
  parts.push('</tbody></table>');
  container.innerHTML = parts.join('');
  state.renderedTables[tableId] = true;
}

function renderYearTables(section, yearResult) {
  document.getElementById(section + '-year-tabs').style.display = 'flex';
  const year = getActiveYearForSection(section);
  if (isEnrichedUserbaseResult(yearResult)) {
    renderEnrichedUserbaseTable(section, section + '-' + year, yearResult[year] || [], yearResult.headers);
  } else {
    renderTable(section + '-' + year, yearResult[year] || []);
  }
}

/** Read a year tab's year from its data-year attribute or its switchYearTab(...) onclick. */
function yearTabYear(tab) {
  return tab.getAttribute('data-year') ||
    ((tab.getAttribute('onclick') || '').match(/'(\d{4})'/) || [])[1] || '';
}

/** Show the per-year total user count on each year tab, e.g. "2024 (123)". */
function updateYearTabCounts(section, yearResult) {
  const panel = document.getElementById('panel-' + section);
  if (!panel) return;
  panel.querySelectorAll('.sheet-tab').forEach(tab => {
    const year = yearTabYear(tab);
    if (!year) return;
    const n = (yearResult[year] || []).length;
    tab.innerHTML = year +
      ' <span class="tab-count" style="pointer-events:none;opacity:.6;font-weight:600;' +
      'font-size:.82em;margin-left:3px">(' + n.toLocaleString() + ')</span>';
  });
}

/** Restore plain year labels (drop the appended count) when a section is cleared. */
function resetYearTabCounts(section) {
  const panel = document.getElementById('panel-' + section);
  if (!panel) return;
  panel.querySelectorAll('.sheet-tab').forEach(tab => {
    const year = yearTabYear(tab);
    if (year) tab.textContent = year;
  });
}

/** Pick the year (from a list) whose row array in `result` is largest; first year on a tie. */
function pickYearWithMostRows(years, result) {
  let best = years[0];
  let max = -1;
  for (const y of years) {
    const n = (result[y] || []).length;
    if (n > max) { max = n; best = y; }
  }
  return best;
}

/**
 * BSC: (re)build the year sub-tabs and their sheet-content panes from the actual
 * years present in the data. `activeYear` receives the .active class.
 */
function buildBscYearTabsDom(years, activeYear) {
  const tabsDiv = document.getElementById('bsc-year-tabs');
  const panel = document.getElementById('panel-bsc');
  if (!tabsDiv || !panel) return;
  const active = activeYear || years[0];

  tabsDiv.innerHTML = years.map(y =>
    '<button class="sheet-tab' + (y === active ? ' active' : '') + '" data-year="' + y +
    '" onclick="switchYearTab(\'bsc\',\'' + y + '\')">' + y + '</button>'
  ).join('');

  // Drop any previously built panes, then rebuild one per year right after the tabs row.
  panel.querySelectorAll('[id^="bsc-sheet-"]').forEach(el => el.remove());
  let anchor = tabsDiv;
  for (const y of years) {
    const pane = document.createElement('div');
    pane.className = 'sheet-content' + (y === active ? ' active' : '');
    pane.id = 'bsc-sheet-' + y;
    pane.innerHTML = '<div class="table-wrap" id="tbl-bsc-' + y + '">' +
      '<div class="empty-state"><div class="empty-icon">📂</div>' +
      '<div class="empty-title">Upload files and calculate</div></div></div>';
    anchor.insertAdjacentElement('afterend', pane);
    anchor = pane;
  }
}

/** BSC: restore the default 2024/2025/2026 sub-tabs when the section is cleared. */
function resetBscYearTabsDom() {
  buildBscYearTabsDom(['2024', '2025', '2026'], '2024');
  const tabsEl = document.getElementById('bsc-year-tabs');
  if (tabsEl) tabsEl.style.display = 'none';
}

function getStatusBadge(status) {
  const s = norm(status);
  if (s === 'completed') return `<span class="badge badge-green">Completed</span>`;
  if (s === 'not found') return `<span class="badge badge-gray">Not Found</span>`;
  if (s === 'terminated') return `<span class="badge badge-gray">Terminated</span>`;
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

function escAttr(v) {
  return esc(v).replace(/"/g, '&quot;');
}

function zoneDownloadButtonsHtml(section, zoneName, counts) {
  const z = escAttr(zoneName);
  const { total, completed, notCompleted, unmapped } = counts;
  const terminated = counts.terminated || 0;
  // BSC + Phishing Normal fold terminated into the unmapped download (label + enable based on both).
  const split = showsUnmappedTerminated(section);
  const unmappedLabel = split ? 'Unmapped/Terminated' : 'Unmapped';
  const unmappedTitle = split
    ? 'Download unmapped + terminated users in ' + esc(zoneName)
    : 'Download unmapped users in ' + esc(zoneName);
  const unmappedDisabled = (unmapped + terminated) === 0;
  return `<div class="zone-dl-actions">
        <button type="button" class="zone-dl-btn zone-dl-completed" data-zone="${z}" data-mode="all" ${total === 0 ? 'disabled' : ''} title="Download all users in ${esc(zoneName)}">All</button>
        <button type="button" class="zone-dl-btn zone-dl-completed" data-zone="${z}" data-mode="completed" ${completed === 0 ? 'disabled' : ''} title="Download completed users in ${esc(zoneName)}">Done</button>
        <button type="button" class="zone-dl-btn zone-dl-pending" data-zone="${z}" data-mode="pending" ${notCompleted === 0 ? 'disabled' : ''} title="Download not completed (mapped) in ${esc(zoneName)}">Open</button>
        <button type="button" class="zone-dl-btn zone-dl-unmapped" data-zone="${z}" data-mode="unmapped" ${unmappedDisabled ? 'disabled' : ''} title="${unmappedTitle}">${unmappedLabel}</button>
      </div>`;
}

function zoneStatLine(label, value, extraClass) {
  return `<div class="zone-stat${extraClass ? ' ' + extraClass : ''}"><span class="zone-stat-label">${esc(label)}</span><strong>${value}</strong></div>`;
}

function centerDownloadButtonsHtml(centerName, counts) {
  const c = escAttr(centerName);
  const { total, completed, notCompleted, unmapped } = counts;
  const terminated = counts.terminated || 0;
  return `<div class="zone-dl-actions">
        <button type="button" class="zone-dl-btn zone-dl-completed" data-center="${c}" data-mode="all" ${total === 0 ? 'disabled' : ''} title="Download all users in ${esc(centerName)}">All</button>
        <button type="button" class="zone-dl-btn zone-dl-completed" data-center="${c}" data-mode="completed" ${completed === 0 ? 'disabled' : ''} title="Download completed users in ${esc(centerName)}">Done</button>
        <button type="button" class="zone-dl-btn zone-dl-pending" data-center="${c}" data-mode="pending" ${notCompleted === 0 ? 'disabled' : ''} title="Download not completed (mapped) in ${esc(centerName)}">Open</button>
        <button type="button" class="zone-dl-btn zone-dl-unmapped" data-center="${c}" data-mode="unmapped" ${(unmapped + terminated) === 0 ? 'disabled' : ''} title="Download unmapped + terminated users in ${esc(centerName)}">Unmapped/Terminated</button>
      </div>`;
}

// ============================================================
// METRICS UPDATE
// ============================================================
function updateUserbaseTrackingMetrics(section, statsRows) {
  if (!statsRows) return;
  let completed = 0;
  let notCompleted = 0;
  let notFound = 0;
  let terminated = 0;
  for (let i = 0; i < statsRows.length; i++) {
    const o = userbaseTrackingOutcome(statsRows[i]);
    if (o === 'completed') completed++;
    else if (o === 'notFound') notFound++;
    else if (o === 'terminated') terminated++;
    else notCompleted++;
  }
  const total = statsRows.length;
  const mappedTotal = total - notFound - terminated;
  const pct = mappedTotal > 0 ? Math.round((completed / mappedTotal) * 100) : 0;
  const pctNc = mappedTotal > 0 ? Math.round((notCompleted / mappedTotal) * 100) : 0;
  const pctNf = total > 0 ? Math.round((notFound / total) * 100) : 0;

  document.getElementById('m-' + section + '-total').textContent = total;
  document.getElementById('m-' + section + '-completed').textContent = completed;
  const ncEl = document.getElementById('m-' + section + '-notCompleted');
  if (ncEl) ncEl.textContent = notCompleted;
  const nfEl = document.getElementById('m-' + section + '-notFound');
  // BSC + Phishing Normal show "unmapped / terminated"; band4 keeps the single unmapped number.
  if (nfEl) nfEl.textContent = showsUnmappedTerminated(section) ? (notFound + ' / ' + terminated) : notFound;

  const pctEl = document.getElementById('m-' + section + '-pct');
  if (pctEl) pctEl.textContent = mappedTotal > 0 ? pct + '% of mapped' : '0% of mapped';
  const pctNcEl = document.getElementById('m-' + section + '-pct-notCompleted');
  if (pctNcEl) pctNcEl.textContent = mappedTotal > 0 ? pctNc + '% of mapped' : '0% of mapped';
  const pctNfEl = document.getElementById('m-' + section + '-pct-notFound');
  if (pctNfEl) pctNfEl.textContent = pctNf + '% unmapped';

  const pb = document.getElementById('pb-' + section);
  if (pb) pb.style.width = pct + '%';
  const pbNc = document.getElementById('pb-notCompleted-' + section);
  if (pbNc) pbNc.style.width = pctNc + '%';
  const pbNf = document.getElementById('pb-notFound-' + section);
  if (pbNf) pbNf.style.width = pctNf + '%';

  updateZoneStatus(section, statsRows);
}

function updateMetrics(section, rows) {
  if (!rows) return;
  if (isUserbaseEnrichedSection(section)) {
    updateUserbaseTrackingMetrics(section, rows);
    return;
  }
  const total = rows.length;
  let completed = 0;
  for (let i = 0; i < rows.length; i++) {
    if (rowCompletionStatus(rows[i], section) === 'completed') completed++;
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

function updateNJMetrics(njResult, year) {
  year = year || getActiveYearForSection('newJoiner');
  const rows = njResult[year] || [];
  const label = document.getElementById('m-newJoiner-total-label');
  if (label) label.textContent = 'Total Users (' + year + ')';
  updateMetrics('newJoiner', rows);
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
  if (el) {
    bindZoneDownloadDelegation(section, el);
    return el;
  }
  const panel = document.getElementById('panel-' + section);
  if (!panel) return null;
  const metricsRow = panel.querySelector('.metrics-row');
  if (!metricsRow) return null;
  el = document.createElement('div');
  el.id = 'zone-status-' + section;
  el.className = 'zone-status-panel';
  el.style.display = 'none';
  metricsRow.insertAdjacentElement('afterend', el);
  bindZoneDownloadDelegation(section, el);
  return el;
}

function ensureCenterPanel(section) {
  let el = document.getElementById('center-status-' + section);
  if (el) {
    bindCenterDownloadDelegation(el);
    return el;
  }
  const panel = document.getElementById('panel-' + section);
  if (!panel) return null;
  const zonePanel = document.getElementById('zone-status-' + section);
  if (!zonePanel) return null;
  el = document.createElement('div');
  el.id = 'center-status-' + section;
  el.className = 'zone-status-panel';
  el.style.display = 'none';
  zonePanel.insertAdjacentElement('afterend', el);
  bindCenterDownloadDelegation(el);
  return el;
}

function updateZoneStatus(section, rows) {
  const el = ensureZonePanel(section);
  if (!el) return;
  const panelYear = getZonePanelYearLabel(section);

  if (!rows || !rows.length) {
    el.style.display = 'none';
    el.innerHTML = '';
    if (section === 'bsc') {
      const centerEl = document.getElementById('center-status-' + section);
      if (centerEl) {
        centerEl.style.display = 'none';
        centerEl.innerHTML = '';
      }
    }
    return;
  }

  const yearRows = section === 'newJoiner' ? rows : filterRowsToCurrentYear(section, rows);
  if (!yearRows.length) {
    el.innerHTML = '<div class="zone-status-title">Zone-wise completion (' + panelYear + ')</div>' +
      '<div class="empty-state"><div class="empty-icon">📅</div><div class="empty-title">No users in ' + panelYear + '</div></div>';
    el.style.display = 'block';
    if (section === 'bsc') {
      const centerEl = document.getElementById('center-status-' + section);
      if (centerEl) {
        centerEl.style.display = 'none';
        centerEl.innerHTML = '';
      }
    }
    return;
  }

  const zones = {};
  for (const row of yearRows) {
    const z = rowZoneForSection(row, section) || 'Unknown';
    if (!zones[z]) zones[z] = { total: 0, completed: 0, notCompleted: 0, unmapped: 0, terminated: 0 };
    zones[z].total++;
    if (isUnmappedRow(row, section)) zones[z].unmapped++;
    else if (isTerminatedOutcomeRow(row, section)) zones[z].terminated++;
    else if (rowCompletionStatus(row, section) === 'completed') zones[z].completed++;
    else zones[z].notCompleted++;
  }

  const keys = sortZoneKeys(Object.keys(zones));
  const showUnmapped = isUserbaseEnrichedSection(section);

  let html = '<div class="zone-status-title">Zone-wise completion (' + panelYear + ')</div><div class="zone-status-grid">';
  for (const z of keys) {
    const { total, completed, notCompleted, unmapped, terminated } = zones[z];
    const mappedTotal = total - unmapped - terminated;
    const pct = mappedTotal > 0 ? Math.round((completed / mappedTotal) * 100) : 0;
    // BSC + Phishing Normal merge unmapped + terminated into one "Unmapped/Terminated" line.
    const unmappedLine = showsUnmappedTerminated(section)
      ? zoneStatLine('Unmapped/Terminated', unmapped + ' / ' + terminated, 'zone-stat-unmapped')
      : zoneStatLine('Unmapped', unmapped, 'zone-stat-unmapped');
    const excludedNote = showsUnmappedTerminated(section)
      ? (unmapped + terminated ? '<br>' + unmapped + ' unmapped / ' + terminated + ' terminated (excluded from %)' : '')
      : (unmapped ? '<br>' + unmapped + ' unmapped (excluded from %)' : '');
    html += `<div class="zone-status-card">
      <div class="zone-status-name">${esc(z)}</div>
      <div class="zone-status-stats">
        ${zoneStatLine('Done', completed)}
        ${zoneStatLine('Open', notCompleted, 'zone-stat-pending')}
        ${showUnmapped ? unmappedLine : zoneStatLine('Pending', total - completed, 'zone-stat-pending')}
        ${zoneStatLine('Total', total)}
      </div>
      <div class="progress-bar-wrap">
        <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${pct}%;background:var(--accent)"></div></div>
      </div>
      <div class="zone-status-pct">${showUnmapped && mappedTotal > 0 ? pct + '% completed (mapped users)' : pct + '% completed'}${showUnmapped ? excludedNote : ''}</div>
      ${showUnmapped
        ? zoneDownloadButtonsHtml(section, z, { total, completed, notCompleted, unmapped, terminated })
        : `<div class="zone-dl-actions">
        <button type="button" class="zone-dl-btn zone-dl-completed" data-zone="${escAttr(z)}" data-mode="all" ${total === 0 ? 'disabled' : ''} title="Download all users in ${esc(z)}">All</button>
        <button type="button" class="zone-dl-btn zone-dl-completed" data-zone="${escAttr(z)}" data-mode="completed" ${completed === 0 ? 'disabled' : ''} title="Download completed users in ${esc(z)}">Done</button>
        <button type="button" class="zone-dl-btn zone-dl-pending" data-zone="${escAttr(z)}" data-mode="pending" ${(total - completed) === 0 ? 'disabled' : ''} title="Download pending users in ${esc(z)}">Open</button>
      </div>`}
    </div>`;
  }
  html += '</div>';
  el.innerHTML = html;
  el.style.display = 'block';

  if (section === 'bsc') {
    updateCenterStatus(section, yearRows);
  } else {
    const centerEl = document.getElementById('center-status-' + section);
    if (centerEl) {
      centerEl.style.display = 'none';
      centerEl.innerHTML = '';
    }
  }
}

function updateCenterStatus(section, rows) {
  const el = ensureCenterPanel(section);
  if (!el) return;
  const currentYear = getCurrentDashboardYearBucket();
  if (!rows || !rows.length) {
    el.style.display = 'none';
    el.innerHTML = '';
    return;
  }

  const yearRows = filterRowsToCurrentYear(section, rows);
  if (!yearRows.length) {
    el.innerHTML = '<div class="zone-status-title">Center-wise completion (' + currentYear + ')</div>' +
      '<div class="empty-state"><div class="empty-icon">📅</div><div class="empty-title">No users in ' + currentYear + '</div></div>';
    el.style.display = 'block';
    return;
  }

  const centers = {};
  for (const row of yearRows) {
    const c = rowCenterForSection(row, section) || 'Unknown';
    if (!centers[c]) centers[c] = { total: 0, completed: 0, notCompleted: 0, unmapped: 0, terminated: 0 };
    centers[c].total++;
    if (isUnmappedRow(row, section)) centers[c].unmapped++;
    else if (isTerminatedOutcomeRow(row, section)) centers[c].terminated++;
    else if (rowCompletionStatus(row, section) === 'completed') centers[c].completed++;
    else centers[c].notCompleted++;
  }
  const keys = Object.keys(centers).sort();
  let html = '<div class="zone-status-title">Center-wise completion (' + currentYear + ')</div><div class="zone-status-grid">';
  for (const c of keys) {
    const { total, completed, notCompleted, unmapped, terminated } = centers[c];
    const mappedTotal = total - unmapped - terminated;
    const pct = mappedTotal > 0 ? Math.round((completed / mappedTotal) * 100) : 0;
    const excludedNote = (unmapped + terminated)
      ? '<br>' + unmapped + ' unmapped / ' + terminated + ' terminated (excluded from %)' : '';
    html += `<div class="zone-status-card">
      <div class="zone-status-name">${esc(c)}</div>
      <div class="zone-status-stats">
        ${zoneStatLine('Done', completed)}
        ${zoneStatLine('Open', notCompleted, 'zone-stat-pending')}
        ${zoneStatLine('Unmapped/Terminated', unmapped + ' / ' + terminated, 'zone-stat-unmapped')}
        ${zoneStatLine('Total', total)}
      </div>
      <div class="progress-bar-wrap">
        <div class="progress-bar-bg"><div class="progress-bar-fill" style="width:${pct}%;background:var(--accent)"></div></div>
      </div>
      <div class="zone-status-pct">${mappedTotal > 0 ? pct + '% completed (mapped users)' : '0% completed'}${excludedNote}</div>
      ${centerDownloadButtonsHtml(c, { total, completed, notCompleted, unmapped, terminated })}
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
    const rows = getMetricRows(s);
    totalU += rows.length;
    const c = rows.filter(x => rowCompletionStatus(x, s) === 'completed').length;
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
  if (isEnrichedUserbaseResult(result)) return result.enriched;
  return [...(result['2024'] || []), ...(result['2025'] || []), ...(result['2026'] || [])];
}

/** Rows used for KPI / zone cards (in-scope cohort for tracking tabs). */
function getMetricRows(section) {
  const result = state.results[section];
  if (!result) return [];
  if (result.statsRows && result.statsRows.length) return result.statsRows;
  return getAllResultRows(section);
}

function getCurrentDashboardYearBucket() {
  const y = new Date().getFullYear();
  if (y <= 2024) return '2024';
  if (y === 2025) return '2025';
  return '2026';
}

function rowStartDateForSection(row, section) {
  if (section === 'bsc') {
    return row['start date_extracted'] || row['Training Start Date'] || '';
  }
  if (section === 'newJoiner') {
    return row._hireDate || row['Training Start Date'] || '';
  }
  return row['Training Start Date'] || row['start date_extracted'] || '';
}

function getNewJoinerActiveYearRows() {
  const result = state.results.newJoiner;
  if (!result) return [];
  const year = getActiveYearForSection('newJoiner');
  return result[year] || [];
}

/** Rows for zone tiles / zone downloads — New Joiner uses active hire-year tab, not calendar year. */
function getRowsForZonePanel(section) {
  if (section === 'newJoiner') return getNewJoinerActiveYearRows();
  return filterRowsToCurrentYear(section, getMetricRows(section));
}

function getZonePanelYearLabel(section) {
  if (section === 'newJoiner') return getActiveYearForSection('newJoiner');
  return getCurrentDashboardYearBucket();
}

function filterRowsToCurrentYear(section, rows) {
  const y = getCurrentDashboardYearBucket();
  return (rows || []).filter(r => yearBucketFromStartDate(rowStartDateForSection(r, section)) === y);
}

function downloadEnrichedUserbase(section, pendingOnly) {
  const result = state.results[section];
  if (!isEnrichedUserbaseResult(result)) {
    toast('error', 'No data available — please calculate first');
    return;
  }
  const cfg = SECTIONS[section];
  const headers = exportHeadersFromEnriched(result.headers);
  let rows = result.enriched;

  if (pendingOnly) {
    if (isUserbaseEnrichedSection(section)) {
      rows = (result.statsRows || []).filter(r => userbaseTrackingOutcome(r) === 'notCompleted');
    } else {
      rows = rows.filter(r => rowCompletionStatus(r, section) !== 'completed');
    }
    if (!rows.length) {
      toast('info', 'No open (mapped) users in userbase');
      return;
    }
  }

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
  styleSheet(ws);
  XLSX.utils.book_append_sheet(wb, ws, 'Userbase');

  const suffix = pendingOnly ? '_Pending_Userbase' : '_Updated_Userbase';
  const filename = cfg.outputFile.replace('.xlsx', suffix + '.xlsx');
  XLSX.writeFile(wb, filename);
  toast(
    'success',
    'Downloaded ' + rows.length.toLocaleString() + ' userbase rows — ' + filename
  );
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
  ['btn-dl-' + section, 'btn-dl-pending-' + section, 'btn-dl-unmapped-' + section, 'btn-dl-zonewise-' + section, 'btn-dl-centerwise-' + section].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.disabled = !enabled;
  });
}

function ensureExtraDownloadButtons() {
  const sections = ['appSec', 'cyberOT', 'growth', 'newJoiner', 'bsc', 'phishingNormal', 'band4'];
  for (const section of sections) {
    const panel = document.getElementById('panel-' + section);
    const btnGroup = panel ? panel.querySelector('.btns-group') : null;
    if (!btnGroup) continue;
    if (isUserbaseEnrichedSection(section)) {
      const unmappedId = 'btn-dl-unmapped-' + section;
      if (!document.getElementById(unmappedId)) {
        const ubtn = document.createElement('button');
        ubtn.type = 'button';
        ubtn.className = 'btn btn-ghost';
        ubtn.id = unmappedId;
        ubtn.disabled = true;
        ubtn.textContent = showsUnmappedTerminated(section)
          ? '⬇ Download Unmapped/Terminated Users'
          : '⬇ Download Unmapped Users';
        ubtn.onclick = () => downloadUnmappedSection(section);
        btnGroup.insertBefore(ubtn, btnGroup.lastElementChild);
      }
    }
    const zoneId = 'btn-dl-zonewise-' + section;
    if (!document.getElementById(zoneId)) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'btn btn-ghost';
      btn.id = zoneId;
      btn.disabled = true;
      btn.textContent = '🗺️ Download Zone-wise';
      btn.onclick = () => downloadZonewiseSection(section);
      btnGroup.insertBefore(btn, btnGroup.lastElementChild);
    }
    if (section === 'bsc') {
      const centerId = 'btn-dl-centerwise-bsc';
      if (!document.getElementById(centerId)) {
        const cbtn = document.createElement('button');
        cbtn.type = 'button';
        cbtn.className = 'btn btn-ghost';
        cbtn.id = centerId;
        cbtn.disabled = true;
        cbtn.textContent = '🏢 Download Centre-wise';
        cbtn.onclick = () => downloadCenterwiseSection('bsc');
        btnGroup.insertBefore(cbtn, btnGroup.lastElementChild);
      }
    }
  }
}

function downloadZoneRows(section, zoneName, mode) {
  const cfg = SECTIONS[section];
  // Band 4+ zone-tile download spans every training-date year (split into per-year sheets below);
  // the zone tile itself stays on the current dashboard year, like the other sections.
  const allRows = section === 'band4' ? getMetricRows(section) : getRowsForZonePanel(section);
  const year = getZonePanelYearLabel(section);
  const yearLabel = section === 'band4' ? 'all years' : year;

  if (!allRows.length) {
    toast('error', 'No data available — please calculate first');
    return;
  }

  const zoneRows = allRows.filter(r => (rowZoneForSection(r, section) || 'Unknown') === zoneName);
  const isUnmappedOrTerm = r => isUnmappedRow(r, section) || isTerminatedOutcomeRow(r, section);
  const rows = mode === 'completed'
    ? zoneRows.filter(r => rowCompletionStatus(r, section) === 'completed')
    : mode === 'pending'
      ? zoneRows.filter(r => !isUnmappedOrTerm(r) && rowCompletionStatus(r, section) !== 'completed')
      : mode === 'unmapped'
        ? zoneRows.filter(isUnmappedOrTerm)
        : zoneRows;

  if (!rows.length) {
    toast(
      'info',
      mode === 'completed'
        ? 'No completed users in ' + zoneName + ' (' + yearLabel + ')'
        : mode === 'pending'
          ? 'No open (mapped) users in ' + zoneName + ' (' + yearLabel + ')'
          : mode === 'unmapped'
            ? 'No ' + (showsUnmappedTerminated(section) ? 'unmapped/terminated' : 'unmapped') + ' users in ' + zoneName + ' (' + yearLabel + ')'
            : 'No users in ' + zoneName + ' (' + yearLabel + ')'
    );
    return;
  }

  const wb = XLSX.utils.book_new();
  const enriched = state.results[section];
  const headers = isEnrichedUserbaseResult(enriched)
    ? exportHeadersFromEnriched(enriched.headers)
    : OUTPUT_COLS;
  const safeZone = sanitizeSheetName(zoneName).replace(/\s+/g, '_');
  const suffix = mode === 'completed' ? '_Completed'
    : mode === 'pending' ? '_Open'
      : mode === 'unmapped' ? '_Unmapped'
        : '_All';

  let filename;
  if (section === 'band4') {
    // One worksheet per training-date year (2024/2025/2026); only years that actually have rows.
    const byYear = {};
    for (const r of rows) {
      const yb = yearBucketFromStartDate(rowStartDateForSection(r, section));
      (byYear[yb] || (byYear[yb] = [])).push(r);
    }
    for (const y of YEAR_TABS) {
      const yrRows = byYear[y];
      if (!yrRows || !yrRows.length) continue;
      const wsY = XLSX.utils.json_to_sheet(yrRows, { header: headers });
      styleSheet(wsY);
      XLSX.utils.book_append_sheet(wb, wsY, sanitizeSheetName(y));
    }
    filename = cfg.outputFile.replace('.xlsx', '_' + safeZone + '_AllYears' + suffix + '.xlsx');
  } else {
    const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
    styleSheet(ws);
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(zoneName));
    filename = cfg.outputFile.replace('.xlsx', '_' + safeZone + '_' + year + suffix + '.xlsx');
  }
  XLSX.writeFile(wb, filename);
  const label = mode === 'completed' ? 'completed'
    : mode === 'pending' ? 'open'
      : mode === 'unmapped' ? (showsUnmappedTerminated(section) ? 'unmapped/terminated' : 'unmapped')
        : 'all';
  toast('success', 'Downloaded ' + rows.length + ' ' + label + ' — ' + zoneName + ' (' + yearLabel + ')');
}

function downloadCenterRows(section, centerName, mode) {
  const cfg = SECTIONS[section];
  const allRows = filterRowsToCurrentYear(section, getMetricRows(section));
  const year = getCurrentDashboardYearBucket();
  if (!allRows.length) {
    toast('error', 'No data available — please calculate first');
    return;
  }
  const centerRows = allRows.filter(r => (rowCenterForSection(r, section) || 'Unknown') === centerName);
  const isUnmappedOrTerm = r => isUnmappedRow(r, section) || isTerminatedOutcomeRow(r, section);
  const rows = mode === 'completed'
    ? centerRows.filter(r => rowCompletionStatus(r, section) === 'completed')
    : mode === 'pending'
      ? centerRows.filter(r => !isUnmappedOrTerm(r) && rowCompletionStatus(r, section) !== 'completed')
      : mode === 'unmapped'
        ? centerRows.filter(isUnmappedOrTerm)
        : centerRows;
  if (!rows.length) {
    toast('info', mode === 'completed' ? 'No completed users in ' + centerName + ' (' + year + ')'
      : mode === 'pending' ? 'No open (mapped) users in ' + centerName + ' (' + year + ')'
        : mode === 'unmapped' ? 'No unmapped/terminated users in ' + centerName + ' (' + year + ')'
          : 'No users in ' + centerName + ' (' + year + ')');
    return;
  }
  const wb = XLSX.utils.book_new();
  const enriched = state.results[section];
  const headers = isEnrichedUserbaseResult(enriched) ? exportHeadersFromEnriched(enriched.headers) : OUTPUT_COLS;
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
  styleSheet(ws);
  XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName(centerName));
  const safeCenter = sanitizeSheetName(centerName).replace(/\s+/g, '_');
  const suffix = mode === 'completed' ? '_Completed'
    : mode === 'pending' ? '_Open'
      : mode === 'unmapped' ? '_Unmapped'
        : '_All';
  const filename = cfg.outputFile.replace('.xlsx', '_' + safeCenter + '_' + year + suffix + '.xlsx');
  XLSX.writeFile(wb, filename);
  const ctrLabel = mode === 'unmapped' ? 'unmapped/terminated' : (mode === 'all' ? 'all' : mode);
  toast('success', 'Downloaded ' + rows.length + ' ' + ctrLabel + ' — ' + centerName + ' (' + year + ')');
}

function downloadZonewiseSection(section) {
  const cfg = SECTIONS[section];
  const year = getZonePanelYearLabel(section);
  const rows = getRowsForZonePanel(section);
  if (!rows.length) {
    toast('error', 'No data available for ' + year + ' — please calculate first');
    return;
  }
  const grouped = {};
  for (const row of rows) {
    const z = rowZoneForSection(row, section) || 'Unknown';
    if (!grouped[z]) grouped[z] = [];
    grouped[z].push(row);
  }
  const wb = XLSX.utils.book_new();
  const enriched = state.results[section];
  const headers = isEnrichedUserbaseResult(enriched) ? exportHeadersFromEnriched(enriched.headers) : OUTPUT_COLS;
  const keys = sortZoneKeys(Object.keys(grouped));
  for (const z of keys) {
    const ws = XLSX.utils.json_to_sheet(grouped[z], { header: headers });
    styleSheet(ws);
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName('Z ' + z));
  }
  const filename = cfg.outputFile.replace('.xlsx', '_Zonewise_' + year + '.xlsx');
  XLSX.writeFile(wb, filename);
  toast('success', 'Downloaded zone-wise workbook (' + year + ') — ' + filename);
}

function downloadCenterwiseSection(section) {
  const cfg = SECTIONS[section];
  const year = getCurrentDashboardYearBucket();
  const rows = filterRowsToCurrentYear(section, getMetricRows(section));
  if (!rows.length) {
    toast('error', 'No data available for ' + year + ' — please calculate first');
    return;
  }
  const grouped = {};
  for (const row of rows) {
    const c = rowCenterForSection(row, section) || 'Unknown';
    if (!grouped[c]) grouped[c] = [];
    grouped[c].push(row);
  }
  const wb = XLSX.utils.book_new();
  const enriched = state.results[section];
  const headers = isEnrichedUserbaseResult(enriched) ? exportHeadersFromEnriched(enriched.headers) : OUTPUT_COLS;
  const keys = Object.keys(grouped).sort();
  for (const c of keys) {
    const ws = XLSX.utils.json_to_sheet(grouped[c], { header: headers });
    styleSheet(ws);
    XLSX.utils.book_append_sheet(wb, ws, sanitizeSheetName('C ' + c));
  }
  const filename = cfg.outputFile.replace('.xlsx', '_Centerwise_' + year + '.xlsx');
  XLSX.writeFile(wb, filename);
  toast('success', 'Downloaded center-wise workbook (' + year + ') — ' + filename);
}

function downloadSection(section) {
  const result = state.results[section];
  if (!result) { toast('error', 'No data to download'); return; }
  if (USERBASE_ENRICHED_SECTIONS.has(section) && isEnrichedUserbaseResult(result)) {
    downloadEnrichedUserbase(section, false);
    return;
  }
  const cfg = SECTIONS[section];
  const allRows = getAllResultRows(section);

  if (!allRows.length) {
    toast('error', 'No data to download — calculate first');
    return;
  }

  const wb = XLSX.utils.book_new();
  const wsAll = XLSX.utils.json_to_sheet(allRows, { header: OUTPUT_COLS });
  styleSheet(wsAll);
  XLSX.utils.book_append_sheet(wb, wsAll, 'All Users');

  const yearLabel = section === 'newJoiner' ? 'New Joiners' : cfg.label.substring(0, 20);
  YEAR_TABS.forEach(year => {
    const rows = result[year] || [];
    if (!rows.length) return;
    const ws = XLSX.utils.json_to_sheet(rows, { header: OUTPUT_COLS });
    styleSheet(ws);
    XLSX.utils.book_append_sheet(wb, ws, yearLabel + ' ' + year);
  });

  XLSX.writeFile(wb, cfg.outputFile);
  toast('success', 'Downloaded ' + allRows.length.toLocaleString() + ' rows — ' + cfg.outputFile);
}

function styleSheet(ws) {
  const wscols = [
    { wch: 16 }, { wch: 38 }, { wch: 14 }, { wch: 18 }, { wch: 20 }, { wch: 22 },
  ];
  ws['!cols'] = wscols;
}

function downloadPendingAppSec() { downloadPendingSection('appSec'); }

function downloadUnmappedSection(section) {
  const cfg = SECTIONS[section];
  const result = state.results[section];
  if (!isEnrichedUserbaseResult(result)) {
    toast('error', 'No data available — please calculate first');
    return;
  }
  const bsc = showsUnmappedTerminated(section);
  const noun = bsc ? 'unmapped/terminated' : 'unmapped';
  const rows = (result.statsRows || []).filter(r =>
    isUnmappedRow(r, section) || (bsc && isTerminatedOutcomeRow(r, section)));
  if (!rows.length) {
    toast('info', 'No ' + noun + ' users in userbase');
    return;
  }
  const wb = XLSX.utils.book_new();
  const headers = exportHeadersFromEnriched(result.headers);
  const ws = XLSX.utils.json_to_sheet(rows, { header: headers });
  styleSheet(ws);
  XLSX.utils.book_append_sheet(wb, ws, bsc ? 'Unmapped+Terminated' : 'Unmapped');
  const filename = cfg.outputFile.replace('.xlsx', bsc ? '_Unmapped_Terminated_Users.xlsx' : '_Unmapped_Users.xlsx');
  XLSX.writeFile(wb, filename);
  toast('success', 'Downloaded ' + rows.length.toLocaleString() + ' ' + noun + ' users — ' + filename);
}

function downloadPendingSection(section) {
  const cfg = SECTIONS[section];
  const result = state.results[section];

  if (!result) {
    toast('error', 'No data available — please calculate first');
    return;
  }

  if (USERBASE_ENRICHED_SECTIONS.has(section) && isEnrichedUserbaseResult(result)) {
    downloadEnrichedUserbase(section, true);
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
  clearDebugLog(section);
  appendDebugLog(section, 'Section cleared');
  // These sections are tool-only — no base file
  const toolOnlySections = ['appSec', 'growth', 'newJoiner'];
  const typesToClear = toolOnlySections.includes(section) ? ['tool'] : ['base', 'tool'];

  delete state.files[section + '-base'];
  delete state.files[section + '-tool'];
  ['-base', '-tool'].forEach(suffix => {
    const key = section + suffix;
    delete state.sheetCache[key];
    const vKey = sheetCacheKey(key);
    if (vKey) delete state.sheetCache[vKey];
  });
  delete state.results[section];
  if (section === 'growth') delete state.growthFilterStats;
  // BSC builds its year sub-tabs dynamically — restore the default tabs before the reset below.
  if (section === 'bsc') resetBscYearTabsDom();
  if (section === 'bsc') delete state.bscMatchStats;
  if (section === 'phishingNormal') delete state.phishingNormalStats;
  if (section === 'band4') delete state.band4Stats;
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
  const metricKeys = isUserbaseEnrichedSection(section)
    ? ['total', 'completed', 'notCompleted', 'notFound']
    : ['total', 'completed', 'pending'];
  metricKeys.forEach(m => {
    const el = document.getElementById('m-' + section + '-' + m);
    if (el) el.textContent = '—';
  });
  const pb = document.getElementById('pb-' + section);
  if (pb) pb.style.width = '0%';
  if (isUserbaseEnrichedSection(section)) {
    const pbNc = document.getElementById('pb-notCompleted-' + section);
    const pbNf = document.getElementById('pb-notFound-' + section);
    if (pbNc) pbNc.style.width = '0%';
    if (pbNf) pbNf.style.width = '0%';
    ['pct', 'pct-notCompleted', 'pct-notFound'].forEach(suffix => {
      const el = document.getElementById('m-' + section + '-' + suffix);
      if (el) el.textContent = '0%';
    });
  }
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
    resetYearTabCounts(section);
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

async function checkUploadServer() {
  const el = document.getElementById('server-status');
  if (el) el.style.display = 'none';
}

ensureExtraDownloadButtons();
ensureDebugPanels();
checkUploadServer();

