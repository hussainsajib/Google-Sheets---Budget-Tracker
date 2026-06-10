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

// ── Node/Jest export (ignored by Apps Script) ────────────────
if (typeof module !== "undefined" && module.exports) {
  module.exports = { colLetter, buildMonthList, sumIfsFormula };
}
