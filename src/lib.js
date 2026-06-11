// ============================================================
// lib.js — PURE helper functions (no SpreadsheetApp dependency)
//
// This file is shared by two runtimes:
//   • Apps Script: clasp pushes it as lib.gs; functions live in the
//     global scope alongside Code.gs (GAS concatenates all files).
//   • Node/Jest: required by test/lib.test.js via module.exports.
//
// The `typeof module` guard at the bottom is skipped in Apps Script
// (where `module` is undefined) and used only by Node.
// Keep this file free of SpreadsheetApp / HtmlService / Session calls.
// ============================================================

/**
 * Convert a 1-based column index to its spreadsheet letter.
 * 1=A, 2=B, … 26=Z, 27=AA, 28=AB …
 */
function colLetter(n) {
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/**
 * Build the list of months for a budget period.
 * Returns [{ year, month (1-based), label "Jan-26" }, …]
 */
function buildMonthList(startYear, startMonth, numMonths) {
  const ABBR = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const months = [];
  for (let i = 0; i < numMonths; i++) {
    const totalMonths = (startMonth - 1) + i;
    const y = startYear + Math.floor(totalMonths / 12);
    const m = (totalMonths % 12) + 1;
    months.push({ year: y, month: m, label: `${ABBR[m - 1]}-${String(y).slice(2)}` });
  }
  return months;
}

/**
 * Build the SUMIFS formula string used in the Budget Tracker for a given
 * category cell ref (e.g. "$A12") and a month {year, month}.
 * Pure string builder — easy to unit test, no spreadsheet calls.
 */
function sumIfsFormula(catRef, monthObj) {
  const y = monthObj.year, mn = monthObj.month;
  return `=IFERROR(SUMIFS(Transactions!$D:$D,` +
         `Transactions!$C:$C,${catRef},` +
         `Transactions!$A:$A,">="&DATE(${y},${mn},1),` +
         `Transactions!$A:$A,"<="&EOMONTH(DATE(${y},${mn},1),0)),0)`;
}

/**
 * Parse CSV text into a 2-D array of strings.
 * Handles quoted fields, embedded commas, escaped quotes (""), CRLF/LF/CR,
 * and UTF-8 BOM. Empty trailing rows are dropped.
 *
 * @param {string} text  Raw CSV text
 * @returns {string[][]}
 */
function parseCsv(text) {
  const s = String(text).replace(/^﻿/, '').replace(/\r\n|\r/g, '\n');
  const rows = [];
  let row = [], field = '', inQuotes = false;

  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (inQuotes) {
      if (ch === '"' && s[i + 1] === '"') { field += '"'; i++; }
      else if (ch === '"')                { inQuotes = false; }
      else                                { field += ch; }
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ',') {
      row.push(field); field = '';
    } else if (ch === '\n') {
      row.push(field); field = '';
      if (row.some(f => f.trim() !== '')) rows.push(row);
      row = [];
    } else {
      field += ch;
    }
  }
  // Last field / row (no trailing newline)
  row.push(field);
  if (row.some(f => f.trim() !== '')) rows.push(row);
  return rows;
}

/**
 * Auto-detect which CSV column index maps to each transaction field.
 * Returns -1 for any field that cannot be matched.
 *
 * @param {string[]} headers  First row of a parsed CSV
 * @returns {{ date, description, amount, category, account, notes }}
 */
function autoDetectMapping(headers) {
  const mapping = { date: -1, description: -1, amount: -1, category: -1, account: -1, notes: -1 };
  const patterns = {
    date:        /date|posted|trans.*date|settlement|value.*date/i,
    description: /desc|memo|narrative|detail|payee|merchant|name|label|particulars/i,
    amount:      /^amount$|^debit$|^credit$|^value$|^sum$|^total$|withdrawal|deposit/i,
    category:    /categor|type|classif/i,
    account:     /account|acct/i,
    notes:       /note|ref(erence)?$|comment|remark|tag/i,
  };
  headers.forEach((h, i) => {
    const hdr = String(h).trim();
    Object.keys(patterns).forEach(field => {
      if (mapping[field] === -1 && patterns[field].test(hdr)) mapping[field] = i;
    });
  });
  return mapping;
}

/**
 * Normalise a raw date string to "yyyy-MM-dd".
 * Supports ISO, MM/DD/YYYY (MDY), DD/MM/YYYY (DMY), yyyyMMdd,
 * and common month-name formats (e.g. "Jun 10, 2026", "10-Jun-2026").
 * Returns the original string unchanged if no format is recognised.
 *
 * @param {string} raw  Raw date value from CSV
 * @param {string} fmt  "MDY" | "DMY" | "YMD" (default MDY)
 * @returns {string}
 */
function normalizeDateString(raw, fmt) {
  const s = String(raw).trim();
  if (!s) return '';
  const pad = n => String(n).padStart(2, '0');
  const MONTHS = { jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12 };
  let m;

  // yyyy-MM-dd or yyyy/MM/dd
  m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return `${m[1]}-${pad(m[2])}-${pad(m[3])}`;

  // yyyyMMdd compact
  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;

  // d/m/y or m/d/y with 2-or-4-digit year
  m = s.match(/^(\d{1,2})[-\/\.](\d{1,2})[-\/\.](\d{2,4})$/);
  if (m) {
    const yr = m[3].length === 2 ? `20${m[3]}` : m[3];
    if (fmt === 'DMY') return `${yr}-${pad(m[2])}-${pad(m[1])}`;
    return `${yr}-${pad(m[1])}-${pad(m[2])}`; // default MDY
  }

  // d-Mon-yyyy  or  d Mon yyyy
  m = s.match(/^(\d{1,2})[-\s]([A-Za-z]{3,9})[-\s,]*(\d{2,4})$/);
  if (m) {
    const mo = MONTHS[m[2].slice(0, 3).toLowerCase()];
    if (mo) { const yr = m[3].length === 2 ? `20${m[3]}` : m[3]; return `${yr}-${pad(mo)}-${pad(m[1])}`; }
  }

  // Mon d, yyyy  or  Month d yyyy
  m = s.match(/^([A-Za-z]{3,9})\s+(\d{1,2})[,\s]+(\d{2,4})$/);
  if (m) {
    const mo = MONTHS[m[1].slice(0, 3).toLowerCase()];
    if (mo) { const yr = m[3].length === 2 ? `20${m[3]}` : m[3]; return `${yr}-${pad(mo)}-${pad(m[2])}`; }
  }

  return s; // unrecognised — return as-is
}

/**
 * Apply AutoCat rules to a transaction description.
 * Rules are checked in order; the first keyword match wins.
 * Matching is case-insensitive substring search.
 *
 * @param {string} description  Transaction description
 * @param {Array}  rules        [{ keyword: string, category: string }, …]
 * @returns {string|null}  Matched category, or null if no rule matches
 */
function applyAutoCatRules(description, rules) {
  const desc = String(description).toLowerCase();
  for (const rule of (rules || [])) {
    const kw  = String(rule.keyword  || "").trim().toLowerCase();
    const cat = String(rule.category || "").trim();
    if (!kw || !cat) continue;
    if (desc.includes(kw)) return cat;
  }
  return null;
}

// ── Node/Jest export (ignored by Apps Script) ────────────────
if (typeof module !== "undefined" && module.exports) {
  module.exports = { colLetter, buildMonthList, sumIfsFormula, applyAutoCatRules, parseCsv, autoDetectMapping, normalizeDateString };
}
