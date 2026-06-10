// ============================================================
// Budget Tracker — Google Apps Script
// Production-ready, multi-sheet personal budget tracker
// with setup wizard, live SUMIFS formulas, and dashboard charts.
//
// SETUP: Extensions > Apps Script > paste this file + SetupWizard.html
//        Refresh sheet → 💰 Budget Tools → Run Setup Wizard
// ============================================================

// ── Sheet names ──────────────────────────────────────────────
const SHEET_DASHBOARD    = "Dashboard";
const SHEET_INPUT        = "Input";        // transaction entry form
const SHEET_TRACKER      = "Budget Tracker";
const SHEET_TRANSACTIONS = "Transactions";
const SHEET_CONFIG       = "Config";       // visible, user-editable configuration sheet

// Fixed row positions in Config sheet (must stay in sync with createConfigSheet)
const CFG_SETTINGS_START = 3;   // first settings value row
const CFG_CAT_DATA_START = 11;  // first category data row

// ── Color palette ────────────────────────────────────────────
const CLR = {
  PRIMARY:        "#1A237E",  // deep navy  — main headers
  ACCENT:         "#3949AB",  // indigo     — sub-headers
  WHITE:          "#FFFFFF",
  GREEN:          "#2E7D32",
  DARKRED:        "#C62828",
  ORANGE:         "#E65100",
  LTGREEN:        "#E8F5E9",
  LTRED:          "#FFEBEE",
  LTORANGE:       "#FFF8E1",
  TOTAL_BG:       "#ECEFF1",
  GREY_TEXT:      "#546E7A",
  // Expense group row backgrounds
  HOUSING:        "#E8EAF6",
  TRANSPORTATION: "#FFEBEE",
  FOOD:           "#E8F5E9",
  UTILITIES:      "#F3E5F5",
  HEALTH:         "#E0F7FA",
  PERSONAL:       "#FFF8E1",
  FINANCIAL:      "#E3F2FD",
  LIFESTYLE:      "#FCE4EC",
  OTHER:          "#EFEBE9",
};

// ── Expense group header colors ───────────────────────────────
const GROUP_HDR = {
  "Housing":        "#3F51B5",
  "Transportation": "#E53935",
  "Food":           "#43A047",
  "Utilities":      "#8E24AA",
  "Health":         "#00ACC1",
  "Personal":       "#FB8C00",
  "Financial":      "#1E88E5",
  "Lifestyle":      "#D81B60",
  "Other":          "#6D4C41",
};

// ── Group background colors (keyed to GROUP_HDR keys) ─────────
const GROUP_BG = {
  "Housing":        CLR.HOUSING,
  "Transportation": CLR.TRANSPORTATION,
  "Food":           CLR.FOOD,
  "Utilities":      CLR.UTILITIES,
  "Health":         CLR.HEALTH,
  "Personal":       CLR.PERSONAL,
  "Financial":      CLR.FINANCIAL,
  "Lifestyle":      CLR.LIFESTYLE,
  "Other":          CLR.OTHER,
};

// ── Pure helpers (colLetter, buildMonthList, sumIfsFormula) live in
//    lib.js — pushed by clasp as lib.gs, shared global scope. ──────

// ============================================================
// MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("💰 Budget Tools")
    .addItem("▶ Run Setup Wizard",              "openSetupWizard")
    .addSeparator()
    .addItem("↺ Rebuild Tracker Rows",          "rebuildTrackerRows")
    .addItem("↺ Refresh Dashboard",             "refreshDashboard")
    .addSeparator()
    .addItem("⚙ Rebuild All Sheets",            "rebuildAllSheets")
    .addSeparator()
    .addItem("🔔 Set Up Trigger (run once)",    "createInstallableTrigger")
    .addToUi();
}

// ── Simple onEdit — works for the owner only ─────────────────
// For shared spreadsheets run  💰 Budget Tools → Set Up Trigger (run once)
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_INPUT) return;
  if (e.range.getA1Notation() === "B11" && e.value === "TRUE") {
    submitTransaction();
  }
}

// ── Installable trigger setup ─────────────────────────────────
// Run this once from the menu to make the submit checkbox work
// for all editors of the spreadsheet, not just the owner.
function createInstallableTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Remove existing onEdit installable triggers to avoid duplicates
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "onEdit")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    "✓ Trigger installed successfully!\n\n" +
    "The Submit checkbox on the Input sheet now works for all editors.\n" +
    "You only need to run this once."
  );
}

// ============================================================
// SETUP WIZARD
// ============================================================

/**
 * Opens the HTML setup wizard dialog (600 × 620 px).
 */
function openSetupWizard() {
  const html = HtmlService.createHtmlOutputFromFile("SetupWizard")
    .setWidth(600)
    .setHeight(620)
    .setTitle("Budget Tracker Setup Wizard");
  SpreadsheetApp.getUi().showModalDialog(html, "💰 Budget Tracker Setup Wizard");
}

/**
 * Returns preset income + expense categories to the wizard.
 * @returns {{ income: Array, expenses: Object }}
 */
function getPresets() {
  const income = [
    { label: "Salary / Wages",      on: true,  amount: 5000 },
    { label: "Freelance Income",     on: false, amount: 0    },
    { label: "Investment Income",    on: false, amount: 0    },
    { label: "Side Business",        on: false, amount: 0    },
    { label: "Rental Income",        on: false, amount: 0    },
    { label: "Government Benefits",  on: false, amount: 0    },
    { label: "Other Income",         on: true,  amount: 0    },
  ];

  const expenses = {
    "Housing": [
      { label: "Rent/Mortgage",    on: true,  amount: 1500 },
      { label: "Home Insurance",   on: true,  amount: 100  },
      { label: "Property Tax",     on: false, amount: 200  },
      { label: "Repairs",          on: false, amount: 100  },
    ],
    "Transportation": [
      { label: "Car Payment/Lease", on: false, amount: 400 },
      { label: "Car Insurance",     on: true,  amount: 150 },
      { label: "Fuel",              on: true,  amount: 200 },
      { label: "Public Transit",    on: false, amount: 150 },
      { label: "Parking",           on: false, amount: 50  },
    ],
    "Food": [
      { label: "Groceries",          on: true,  amount: 600 },
      { label: "Dining Out/Takeout", on: true,  amount: 200 },
      { label: "Coffee & Snacks",    on: false, amount: 50  },
    ],
    "Utilities": [
      { label: "Electricity/Gas",  on: true,  amount: 100 },
      { label: "Water",            on: false, amount: 50  },
      { label: "Internet",         on: true,  amount: 80  },
      { label: "Phone/Mobile",     on: true,  amount: 100 },
      { label: "Streaming",        on: false, amount: 30  },
    ],
    "Health": [
      { label: "Health Insurance", on: false, amount: 300 },
      { label: "Doctor/Dental",    on: false, amount: 100 },
      { label: "Pharmacy",         on: false, amount: 50  },
      { label: "Gym/Fitness",      on: false, amount: 60  },
    ],
    "Personal": [
      { label: "Clothing & Shopping", on: false, amount: 100 },
      { label: "Personal Care",       on: false, amount: 50  },
      { label: "Education/Courses",   on: false, amount: 100 },
      { label: "Childcare",           on: false, amount: 500 },
    ],
    "Financial": [
      { label: "Savings/Investments", on: true,  amount: 200 },
      { label: "Emergency Fund",      on: false, amount: 100 },
      { label: "Loan Payments",       on: false, amount: 300 },
      { label: "CC Payments",         on: false, amount: 200 },
    ],
    "Lifestyle": [
      { label: "Entertainment",    on: false, amount: 100 },
      { label: "Travel/Vacations", on: false, amount: 100 },
      { label: "Subscriptions",    on: false, amount: 30  },
      { label: "Gifts & Donations",on: false, amount: 50  },
    ],
    "Other": [
      { label: "Miscellaneous", on: true, amount: 100 },
    ],
  };

  return { income, expenses };
}

/**
 * Called by the wizard with the completed config object.
 * Creates the Config sheet then builds all other sheets.
 *
 * @param {Object} config  - { title, currency, startYear, startMonth,
 *                             numMonths, incomeCategories, expenseCategories }
 * @returns {{ success: boolean, error?: string }}
 */
function saveSetupConfig(config) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    createConfigSheet(ss, config);
    buildAllSheets(config);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ============================================================
// CONFIG SHEET  (visible, user-editable)
// ============================================================

/**
 * Creates (or recreates) the Config sheet — the single source of truth
 * for settings and categories. Users can edit this sheet directly at any
 * time without running the wizard again.
 *
 * Layout (fixed row positions, defined by CFG_* constants):
 *   Row 1        : Title bar
 *   Row 2        : Settings section header
 *   Rows 3–7     : Setting label (col A) | editable value (col B)
 *   Row 8        : spacer
 *   Row 9        : Categories section header
 *   Row 10       : Column headers  (Category Name | Type | Group | Monthly Budget)
 *   Row 11+      : Category data   (one row per category — add/delete freely)
 */
function createConfigSheet(ss, config) {
  let ws = ss.getSheetByName(SHEET_CONFIG);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(SHEET_CONFIG);

  ws.setColumnWidth(1, 210);  // Category / Setting label
  ws.setColumnWidth(2, 120);  // Type / Setting value
  ws.setColumnWidth(3, 170);  // Group
  ws.setColumnWidth(4, 140);  // Monthly Budget

  // ── Row 1: Title ─────────────────────────────────────────
  ws.setRowHeight(1, 36);
  ws.getRange(1, 1, 1, 4).merge()
    .setValue("⚙  Budget Configuration")
    .setBackground(CLR.PRIMARY).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── Row 2: Settings section header ───────────────────────
  ws.setRowHeight(2, 22);
  ws.getRange(2, 1, 1, 4).merge()
    .setValue("SETTINGS  —  edit the values in column B")
    .setBackground(CLR.ACCENT).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");

  // ── Rows 3–7: Settings key / value pairs ─────────────────
  // CFG_SETTINGS_START = 3
  const settingRows = [
    ["Spreadsheet Title",  config.title],
    ["Currency Symbol",    config.currency],
    ["Start Month (1–12)", config.startMonth],
    ["Start Year",         config.startYear],
    ["Number of Months",   config.numMonths],
  ];
  ws.getRange(CFG_SETTINGS_START, 1, settingRows.length, 2).setValues(settingRows);
  ws.getRange(CFG_SETTINGS_START, 1, settingRows.length, 1)
    .setBackground("#ECEFF1").setFontWeight("bold").setFontSize(10)
    .setVerticalAlignment("middle");
  ws.getRange(CFG_SETTINGS_START, 2, settingRows.length, 1)
    .setBackground("#FFFDE7").setFontSize(11).setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null, "#FB8C00", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  for (let i = 0; i < settingRows.length; i++) ws.setRowHeight(CFG_SETTINGS_START + i, 22);

  // ── Row 8: spacer ─────────────────────────────────────────
  ws.setRowHeight(8, 10);

  // ── Row 9: Categories section header ─────────────────────
  ws.setRowHeight(9, 22);
  ws.getRange(9, 1, 1, 4).merge()
    .setValue("CATEGORIES  —  add or delete rows freely. Budget changes here update the tracker instantly.")
    .setBackground(CLR.ACCENT).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");

  // ── Row 10: Category table column headers ─────────────────
  ws.setRowHeight(10, 24);
  ws.getRange(10, 1, 1, 4)
    .setValues([["Category Name", "Type", "Group", "Monthly Budget"]])
    .setBackground(CLR.PRIMARY).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── Rows 11+: Category data (CFG_CAT_DATA_START = 11) ────
  const catRows = [];
  (config.incomeCategories  || []).forEach(c => catRows.push([c.label, "Income",  "",             c.budget]));
  (config.expenseCategories || []).forEach(c => catRows.push([c.label, "Expense", c.group || "", c.budget]));

  if (catRows.length > 0) {
    ws.getRange(CFG_CAT_DATA_START, 1, catRows.length, 4).setValues(catRows);
    ws.getRange(CFG_CAT_DATA_START, 4, catRows.length, 1)
      .setNumberFormat(config.currency + "#,##0.00");
  }

  // Data-validation dropdowns (100 rows — fast, enough for any realistic category count)
  ws.getRange(CFG_CAT_DATA_START, 2, 100, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Income", "Expense"], true)
      .setAllowInvalid(false).build()
  );
  ws.getRange(CFG_CAT_DATA_START, 3, 100, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.keys(GROUP_HDR).concat([""]), true)
      .setAllowInvalid(true).build()
  );

  // ── Tip row ───────────────────────────────────────────────
  const tipRow = CFG_CAT_DATA_START + catRows.length + 2;
  ws.setRowHeight(tipRow, 48);
  ws.getRange(tipRow, 1, 1, 4).merge()
    .setValue("💡  How to use this sheet:\n" +
              "• Change a Monthly Budget → Budget Tracker updates immediately (live formula link).\n" +
              "• Add / remove category rows → run  💰 Budget Tools › Rebuild Tracker Rows  to refresh tracker layout.")
    .setBackground("#FFF8E1").setFontSize(9).setFontStyle("italic").setFontColor("#555")
    .setWrap(true).setVerticalAlignment("middle");

  ws.setFrozenRows(10);
}

/**
 * Reads the Config sheet and returns a config object, or null if not set up.
 * Budget amounts are read so rebuildAllSheets / rebuildTrackerRows stay accurate,
 * but the Budget Tracker itself references Config via VLOOKUP for live updates.
 * @returns {Object|null}
 */
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_CONFIG);
  if (!ws) return null;

  // Settings block: rows CFG_SETTINGS_START … CFG_SETTINGS_START+4, col B
  const settingVals = ws.getRange(CFG_SETTINGS_START, 2, 5, 1).getValues();
  const title      = String(settingVals[0][0] || "My Budget");
  const currency   = String(settingVals[1][0] || "$");
  const startMonth = Number(settingVals[2][0] || 1);
  const startYear  = Number(settingVals[3][0] || new Date().getFullYear());
  const numMonths  = Number(settingVals[4][0] || 12);

  if (!title) return null;

  // Category data: rows CFG_CAT_DATA_START onwards, cols A–D
  const lastRow = ws.getLastRow();
  const incomeCategories  = [];
  const expenseCategories = [];

  if (lastRow >= CFG_CAT_DATA_START) {
    const catData = ws.getRange(CFG_CAT_DATA_START, 1, lastRow - CFG_CAT_DATA_START + 1, 4).getValues();
    catData.forEach(([label, type, group, budget]) => {
      if (!label) return;
      const t = String(type).trim().toLowerCase();
      if (t === "income") {
        incomeCategories.push({ label: String(label).trim(), budget: Number(budget) || 0 });
      } else if (t === "expense") {
        expenseCategories.push({ label: String(label).trim(), group: String(group).trim() || "Other", budget: Number(budget) || 0 });
      }
    });
  }

  return { title, currency, startYear, startMonth, numMonths, incomeCategories, expenseCategories };
}

// ============================================================
// ORCHESTRATOR
// ============================================================

/**
 * Builds / rebuilds all four sheets in order.
 */
function buildAllSheets(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createInputSheet(ss, config);
  createTransactionSheet(ss, config);
  createBudgetTracker(ss, config);
  createDashboard(ss, config);

  // Tab order: Dashboard → Input → Budget Tracker → Transactions → Config
  [SHEET_DASHBOARD, SHEET_INPUT, SHEET_TRACKER, SHEET_TRANSACTIONS, SHEET_CONFIG].forEach((name, pos) => {
    const sheet = ss.getSheetByName(name);
    if (sheet) { ss.setActiveSheet(sheet); ss.moveActiveSheet(pos + 1); }
  });

  // Land on Input sheet so user can start entering right away
  const inputSheet = ss.getSheetByName(SHEET_INPUT);
  if (inputSheet) ss.setActiveSheet(inputSheet);
}

/**
 * Rebuilds only the Budget Tracker rows from the current Config sheet.
 * Run this after adding or removing categories in the Config sheet.
 * Does NOT touch Transactions or Dashboard.
 */
function rebuildTrackerRows() {
  const config = getSettings();
  if (!config) {
    SpreadsheetApp.getUi().alert(
      "Config sheet not found.\nPlease run the Setup Wizard first:\n💰 Budget Tools → Run Setup Wizard"
    );
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createBudgetTracker(ss, config);
  SpreadsheetApp.getUi().alert(
    "Budget Tracker rows rebuilt from Config sheet.\n\n" +
    "Tip: Budget amounts update live — no rebuild needed just for budget changes."
  );
}

/**
 * Re-reads Config sheet and rebuilds all sheets.
 * Called from the menu item "⚙ Rebuild All Sheets".
 */
function rebuildAllSheets() {
  const config = getSettings();
  if (!config) {
    SpreadsheetApp.getUi().alert(
      "Config sheet not found.\nPlease run the Setup Wizard first:\n💰 Budget Tools → Run Setup Wizard"
    );
    return;
  }
  buildAllSheets(config);
  SpreadsheetApp.getUi().alert("All sheets rebuilt successfully!");
}

// ============================================================
// INPUT SHEET  (transaction entry form)
// ============================================================

/**
 * Creates the Input sheet — a clean form for entering transactions.
 * Checking the Submit checkbox (B11) triggers onEdit → submitTransaction().
 *
 * Fixed cell map:
 *   B4  Date        B5  Description   B6  Category
 *   B7  Amount      B8  Account       B9  Notes
 *   B11 Submit checkbox
 *   B2  Status / feedback message
 */
function createInputSheet(ss, config) {
  let ws = ss.getSheetByName(SHEET_INPUT);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(SHEET_INPUT);

  ws.setColumnWidth(1, 160);
  ws.setColumnWidth(2, 290);
  ws.setColumnWidth(3, 190);
  // Hide unused columns to keep the form clean
  if (ws.getMaxColumns() > 3) ws.hideColumns(4, ws.getMaxColumns() - 3);

  const currency = config.currency || "$";

  // ── Row 1: Title ─────────────────────────────────────────
  ws.setRowHeight(1, 42);
  ws.getRange(1, 1, 1, 3).merge()
    .setValue("➕  Add Transaction")
    .setBackground(CLR.PRIMARY).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(15)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── Row 2: Status / feedback ──────────────────────────────
  ws.setRowHeight(2, 28);
  ws.getRange("A2").setBackground("#F5F5F5");
  ws.getRange("B2")
    .setValue("Fill in the fields below, then check the Submit box ✓")
    .setBackground("#E8EAF6").setFontColor(CLR.ACCENT)
    .setFontSize(10).setFontStyle("italic").setVerticalAlignment("middle");
  ws.getRange("C2").setBackground("#F5F5F5");

  // ── Row 3: spacer ────────────────────────────────────────
  ws.setRowHeight(3, 10);

  // ── Rows 4–9: Form fields ─────────────────────────────────
  const fields = [
    { row: 4, label: "📅  Date",         hint: "Today's date is pre-filled" },
    { row: 5, label: "📝  Description",  hint: "What was this transaction?" },
    { row: 6, label: "🏷  Category",     hint: "Pick from your budget categories" },
    { row: 7, label: "💰  Amount",       hint: `Positive number — ${currency}` },
    { row: 8, label: "💳  Account",      hint: "Optional" },
    { row: 9, label: "📌  Notes",        hint: "Optional" },
  ];
  fields.forEach(({ row, label, hint }) => {
    ws.setRowHeight(row, 34);
    ws.getRange(row, 1)
      .setValue(label)
      .setBackground("#ECEFF1").setFontWeight("bold").setFontSize(10)
      .setHorizontalAlignment("right").setVerticalAlignment("middle");
    ws.getRange(row, 2)
      .setBackground("#FFFDE7").setFontSize(11).setVerticalAlignment("middle")
      .setBorder(true, true, true, true, null, null, "#FB8C00", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    ws.getRange(row, 3)
      .setValue(hint)
      .setFontSize(9).setFontStyle("italic").setFontColor("#9E9E9E")
      .setVerticalAlignment("middle");
  });

  // Pre-fill date with today
  ws.getRange("B4").setValue(new Date()).setNumberFormat("yyyy-mm-dd");

  // Amount number format
  ws.getRange("B7").setNumberFormat(currency + "#,##0.00");

  // Category dropdown — live from Config
  const configWs = ss.getSheetByName(SHEET_CONFIG);
  if (configWs) {
    ws.getRange("B6").setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(configWs.getRange(CFG_CAT_DATA_START, 1, 100, 1), true)
        .setAllowInvalid(false).setHelpText("Select a category").build()
    );
  }

  // Account dropdown
  ws.getRange("B8").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Chequing", "Savings", "Credit Card", "Cash", "Other"], true)
      .setAllowInvalid(true).build()
  );

  // ── Row 10: spacer ────────────────────────────────────────
  ws.setRowHeight(10, 12);

  // ── Row 11: Submit checkbox ───────────────────────────────
  ws.setRowHeight(11, 38);
  ws.getRange("A11")
    .setValue("✓  Submit")
    .setBackground(CLR.GREEN).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(12)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  ws.getRange("B11")
    .insertCheckboxes()
    .setBackground("#E8F5E9").setHorizontalAlignment("left").setVerticalAlignment("middle");
  ws.getRange("C11")
    .setValue("Check this box to save the transaction and reset the form")
    .setFontSize(9).setFontStyle("italic").setFontColor("#9E9E9E").setVerticalAlignment("middle");

  // ── Row 12: spacer ────────────────────────────────────────
  ws.setRowHeight(12, 18);

  // ── Rows 13–19: Recent entries ────────────────────────────
  ws.setRowHeight(13, 24);
  ws.getRange(13, 1, 1, 3).merge()
    .setValue("RECENT ENTRIES")
    .setBackground(CLR.ACCENT).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");

  ws.setRowHeight(14, 20);
  ws.getRange(14, 1, 1, 3).setValues([["Date", "Category · Description", "Amount"]])
    .setBackground(CLR.PRIMARY).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(9)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // 5 formula rows — show the 5 most recently appended transactions
  for (let i = 1; i <= 5; i++) {
    const fRow = 14 + i;
    const idx  = `COUNTA(Transactions!A$2:A$10000)+1-${i}`;
    const bg   = i % 2 === 0 ? "#F5F5F5" : CLR.WHITE;
    ws.setRowHeight(fRow, 20);
    ws.getRange(fRow, 1)
      .setFormula(`=IFERROR(TEXT(INDEX(Transactions!A$2:A$10000,${idx}),"mmm d yyyy"),"")`)
      .setBackground(bg).setFontSize(9).setVerticalAlignment("middle");
    ws.getRange(fRow, 2)
      .setFormula(`=IFERROR(INDEX(Transactions!C$2:C$10000,${idx})&IF(INDEX(Transactions!B$2:B$10000,${idx})<>"", " · "&INDEX(Transactions!B$2:B$10000,${idx}),""),"")`)
      .setBackground(bg).setFontSize(9).setVerticalAlignment("middle");
    ws.getRange(fRow, 3)
      .setFormula(`=IFERROR(INDEX(Transactions!D$2:D$10000,${idx}),"")`)
      .setNumberFormat(currency + "#,##0.00")
      .setBackground(bg).setFontSize(9).setVerticalAlignment("middle")
      .setHorizontalAlignment("right");
  }

  ws.setRowHeight(20, 12);
  ws.getRange(20, 1, 1, 3).merge()
    .setValue("Tip: You can also log transactions directly in the Transactions sheet for bulk entry.")
    .setFontSize(8).setFontStyle("italic").setFontColor("#BDBDBD")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
}

/**
 * Reads the Input sheet form, validates, appends a row to Transactions,
 * then resets the form. Called by the onEdit trigger when B11 is checked.
 */
function submitTransaction() {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const inputWs = ss.getSheetByName(SHEET_INPUT);
  const txWs    = ss.getSheetByName(SHEET_TRANSACTIONS);
  if (!inputWs || !txWs) return;

  const statusCell = inputWs.getRange("B2");

  const dateVal = inputWs.getRange("B4").getValue();
  const desc    = String(inputWs.getRange("B5").getValue()).trim();
  const cat     = String(inputWs.getRange("B6").getValue()).trim();
  const amount  = inputWs.getRange("B7").getValue();
  const account = String(inputWs.getRange("B8").getValue()).trim();
  const notes   = String(inputWs.getRange("B9").getValue()).trim();

  // Uncheck the box immediately regardless of outcome
  inputWs.getRange("B11").setValue(false);

  // Validate required fields
  if (!dateVal || !cat || !amount) {
    statusCell
      .setValue("⚠  Please fill in Date, Category, and Amount before submitting.")
      .setBackground("#FFEBEE").setFontColor("#C62828")
      .setFontStyle("normal").setFontWeight("bold");
    return;
  }

  // Format date as yyyy-mm-dd string
  const tz      = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(new Date(dateVal), tz, "yyyy-MM-dd");

  // Append to Transactions sheet
  txWs.appendRow([dateStr, desc, cat, Number(amount), account, notes]);

  // Show success status
  const config   = getSettings();
  const currency = config ? config.currency : "$";
  const formatted = currency + Number(amount).toFixed(2);
  statusCell
    .setValue(`✓  Saved: ${cat}  ${formatted}  on ${dateStr}`)
    .setBackground("#E8F5E9").setFontColor(CLR.GREEN)
    .setFontStyle("italic").setFontWeight("normal");

  // Reset form — clear all input cells, reset date to today
  inputWs.getRange("B4").setValue(new Date()).setNumberFormat("yyyy-mm-dd");
  inputWs.getRange("B5:B9").clearContent();
}

// ============================================================
// TRANSACTIONS SHEET
// ============================================================

/**
 * Creates the Transactions input sheet.
 * Columns: Date | Description | Category | Amount | Account | Notes
 */
function createTransactionSheet(ss, config) {
  let ws = ss.getSheetByName(SHEET_TRANSACTIONS);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(SHEET_TRANSACTIONS);

  // ── Column widths ─────────────────────────────────────────
  ws.setColumnWidth(1, 120);  // Date
  ws.setColumnWidth(2, 260);  // Description
  ws.setColumnWidth(3, 210);  // Category
  ws.setColumnWidth(4, 130);  // Amount
  ws.setColumnWidth(5, 160);  // Account
  ws.setColumnWidth(6, 220);  // Notes

  // ── Header row ────────────────────────────────────────────
  ws.setRowHeight(1, 32);
  const headers = ["Date", "Description", "Category", "Amount", "Account", "Notes"];
  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(CLR.PRIMARY)
    .setFontColor(CLR.WHITE)
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  ws.setFrozenRows(1);

  // ── Number formats (limit to 200 rows — fast, expandable later) ─
  ws.getRange(2, 1, 200, 1).setNumberFormat("yyyy-mm-dd");
  ws.getRange(2, 4, 200, 1).setNumberFormat(config.currency + "#,##0.00");

  // ── Category dropdown (live reference to Config sheet) ───────
  const configWs = ss.getSheetByName(SHEET_CONFIG);
  if (configWs) {
    ws.getRange(2, 3, 200, 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(configWs.getRange(CFG_CAT_DATA_START, 1, 100, 1), true)
        .setAllowInvalid(false)
        .setHelpText("Select a budget category")
        .build()
    );
  }

  // ── Account dropdown ──────────────────────────────────────
  ws.getRange(2, 5, 200, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Chequing", "Savings", "Credit Card", "Cash", "Other"], true)
      .setAllowInvalid(true)
      .build()
  );

  // ── Alternating row colours (lightweight CF — avoids slow applyRowBanding) ─
  const dataRange = ws.getRange(2, 1, 200, headers.length);
  ws.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied("=ISEVEN(ROW())")
      .setBackground("#F5F5F5")
      .setRanges([dataRange]).build(),
  ]);

  // ── Example rows ─────────────────────────────────────────
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const firstIncome  = config.incomeCategories  && config.incomeCategories.length  > 0 ? config.incomeCategories[0].label  : "Salary / Wages";
  const firstExpense = config.expenseCategories && config.expenseCategories.length > 0 ? config.expenseCategories[0].label : "Groceries";

  ws.getRange(2, 1, 3, 6).setValues([
    [today, "EXAMPLE: Monthly paycheck",   firstIncome,  3500.00, "Chequing",    "Delete these 3 example rows"],
    [today, "EXAMPLE: Weekly grocery run", firstExpense,  185.50, "Credit Card", ""],
    [today, "EXAMPLE: Electric bill",      firstExpense,  110.00, "Chequing",    "Switch to PAD"],
  ]);
}

// ============================================================
// BUDGET TRACKER SHEET
// ============================================================

/**
 * Creates the Budget Tracker sheet with SUMIFS live formulas.
 *
 * Layout:
 *   Col A = Category label
 *   Col B = Monthly Budget
 *   Cols C … (C + numMonths - 1) = monthly actual (SUMIFS)
 *   Next col = Total
 *   Next col = Avg
 *   Last col  = vs Budget
 */
function createBudgetTracker(ss, config) {
  let ws = ss.getSheetByName(SHEET_TRACKER);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(SHEET_TRACKER);

  const months      = buildMonthList(config.startYear, config.startMonth, config.numMonths);
  const numM        = months.length;

  // Column indices (1-based)
  const COL_CAT    = 1;
  const COL_BUDGET = 2;
  const COL_FIRST  = 3;                     // first month actual
  const COL_LAST   = COL_FIRST + numM - 1;  // last month actual
  const COL_TOTAL  = COL_LAST + 1;
  const COL_AVG    = COL_LAST + 2;
  const COL_VS     = COL_LAST + 3;
  const TOTAL_COLS = COL_VS;

  // ── Column widths ─────────────────────────────────────────
  ws.setColumnWidth(COL_CAT,    200);
  ws.setColumnWidth(COL_BUDGET, 110);
  for (let c = COL_FIRST; c <= COL_LAST; c++) ws.setColumnWidth(c, 95);
  ws.setColumnWidth(COL_TOTAL,  100);
  ws.setColumnWidth(COL_AVG,    90);
  ws.setColumnWidth(COL_VS,     100);

  const lastColLetter = colLetter(COL_VS);

  // ── Row 1: Title spanning all columns ─────────────────────
  ws.setRowHeight(1, 32);
  ws.getRange(1, 1, 1, TOTAL_COLS).merge()
    .setValue(config.title)
    .setBackground(CLR.PRIMARY)
    .setFontColor(CLR.WHITE)
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // ── Row 2: Column headers ─────────────────────────────────
  ws.setRowHeight(2, 32);
  const hdrVals = ["Category", "Budget/mo"];
  months.forEach(m => hdrVals.push(m.label));
  hdrVals.push("Total", "Avg/mo", "vs Budget");
  ws.getRange(2, 1, 1, TOTAL_COLS)
    .setValues([hdrVals])
    .setBackground(CLR.ACCENT)
    .setFontColor(CLR.WHITE)
    .setFontWeight("bold")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true);

  ws.setFrozenRows(2);

  let r = 3; // current row

  // Helper: apply SUMIFS formula for a given category label + month object
  function sumIfsFormula(categoryLabel, monthObj) {
    const y  = monthObj.year;
    const mn = monthObj.month;
    return `=IFERROR(SUMIFS(Transactions!$D:$D,` +
           `Transactions!$C:$C,$A${r},` +
           `Transactions!$A:$A,">="&DATE(${y},${mn},1),` +
           `Transactions!$A:$A,"<="&EOMONTH(DATE(${y},${mn},1),0)),0)`;
  }

  // Helper: write Total / Avg / vs Budget formulas — batched into one call
  function writeSummaryFormulas(rowNum) {
    const fc = colLetter(COL_FIRST);
    const lc = colLetter(COL_LAST);
    const ac = colLetter(COL_AVG);
    ws.getRange(rowNum, COL_TOTAL, 1, 3).setFormulas([[
      `=SUM(${fc}${rowNum}:${lc}${rowNum})`,
      `=IFERROR(SUM(${fc}${rowNum}:${lc}${rowNum})/COUNTA(${fc}${rowNum}:${lc}${rowNum}),0)`,
      `=IFERROR(${ac}${rowNum}-B${rowNum},"")`
    ]]);
  }

  // Helper: format a data row — chains on same range objects to minimise API calls
  function formatDataRow(rowNum, bg, isTotalRow) {
    const fmtCurrency = `${config.currency}#,##0.00`;
    // Col A: label
    ws.getRange(rowNum, COL_CAT)
      .setBackground(bg).setHorizontalAlignment("left")
      .setVerticalAlignment("middle").setFontSize(isTotalRow ? 10 : 9);
    // Col B onward: all numeric columns in one range
    ws.getRange(rowNum, COL_BUDGET, 1, TOTAL_COLS - 1)
      .setBackground(bg).setNumberFormat(fmtCurrency)
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setFontSize(isTotalRow ? 10 : 9);
    // Bold + border on full row
    const rowRange = ws.getRange(rowNum, 1, 1, TOTAL_COLS);
    if (isTotalRow) rowRange.setFontWeight("bold");
    rowRange.setBorder(true, true, true, true, true, true, "#BDBDBD", SpreadsheetApp.BorderStyle.SOLID);
    ws.setRowHeight(rowNum, 20);
  }

  // Helper: write a section header row (merged, colored)
  function writeSectionHeader(label, bgColor, textColor) {
    ws.setRowHeight(r, 22);
    ws.getRange(r, 1, 1, TOTAL_COLS).merge()
      .setValue(label)
      .setBackground(bgColor)
      .setFontColor(textColor || CLR.WHITE)
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("left")
      .setVerticalAlignment("middle");
    r++;
  }

  // Tracks data rows that need vs-Budget CF — applied in one batch at the end
  const vsBudgetRows = [];

  // Helper: write a category row — batches all formula writes
  function writeCategoryRow(label, budget, bg, months) {
    ws.getRange(r, COL_CAT).setValue(label);
    ws.getRange(r, COL_BUDGET)
      .setFormula(`=IFERROR(VLOOKUP($A${r},Config!$A:$D,4,FALSE),0)`);
    // All monthly SUMIFS in one call instead of numM individual calls
    ws.getRange(r, COL_FIRST, 1, numM)
      .setFormulas([months.map(mObj => sumIfsFormula(label, mObj))]);
    writeSummaryFormulas(r);
    formatDataRow(r, bg, false);
    vsBudgetRows.push(r);
    r++;
  }

  // Helper: write a subtotal / total row — batches monthly formula writes
  function writeTotalRow(label, rowsToSum, bg) {
    ws.getRange(r, COL_CAT).setValue(label);

    if (rowsToSum.length > 0) {
      ws.getRange(r, COL_BUDGET)
        .setFormula(`=${rowsToSum.map(rn => `B${rn}`).join("+")}`);
      // All monthly totals in one call instead of numM individual calls
      ws.getRange(r, COL_FIRST, 1, numM).setFormulas([
        Array.from({length: numM}, (_, ci) => {
          const cLtr = colLetter(COL_FIRST + ci);
          return `=${rowsToSum.map(rn => `${cLtr}${rn}`).join("+")}`;
        })
      ]);
    }

    writeSummaryFormulas(r);
    formatDataRow(r, bg, true);
    ws.getRange(r, 1, 1, TOTAL_COLS).setBackground(bg);
    r++;
    return r - 1;
  }

  // ─────────────────────────────────────────────────────────
  // INCOME SECTION
  // ─────────────────────────────────────────────────────────
  writeSectionHeader("INCOME", CLR.GREEN, CLR.WHITE);

  const incomeRows = [];
  config.incomeCategories.forEach(cat => {
    incomeRows.push(r);
    writeCategoryRow(cat.label, cat.budget, CLR.LTGREEN, months);
  });

  const incomeTotalRow = r;
  writeTotalRow("INCOME TOTAL", incomeRows, "#C8E6C9");

  // spacer
  ws.setRowHeight(r, 8);
  ws.getRange(r, 1, 1, TOTAL_COLS).setBackground(CLR.WHITE);
  r++;

  // ─────────────────────────────────────────────────────────
  // EXPENSE SECTIONS (by group)
  // ─────────────────────────────────────────────────────────
  const allExpenseGroups  = _getExpenseGroups(config);
  const groupSubtotalRows = {}; // groupName → subtotal row number
  const allExpenseSubRows = []; // all subtotal rows for GRAND TOTAL

  allExpenseGroups.forEach(({ group, categories }) => {
    const groupBg  = GROUP_BG[group]  || CLR.OTHER;
    const groupHdr = GROUP_HDR[group] || "#555555";

    writeSectionHeader(group.toUpperCase(), groupHdr, CLR.WHITE);

    const catRows = [];
    categories.forEach(cat => {
      catRows.push(r);
      writeCategoryRow(cat.label, cat.budget, groupBg, months);
    });

    const subtotalRow = r;
    writeTotalRow(`${group.toUpperCase()} SUBTOTAL`, catRows, CLR.TOTAL_BG);
    groupSubtotalRows[group] = subtotalRow;
    allExpenseSubRows.push(subtotalRow);

    // spacer between groups
    ws.setRowHeight(r, 6);
    ws.getRange(r, 1, 1, TOTAL_COLS).setBackground(CLR.WHITE);
    r++;
  });

  // ─────────────────────────────────────────────────────────
  // GRAND TOTAL EXPENSES
  // ─────────────────────────────────────────────────────────
  const grandTotalRow = r;
  writeTotalRow("GRAND TOTAL EXPENSES", allExpenseSubRows, "#CFD8DC");
  ws.getRange(grandTotalRow, 1, 1, TOTAL_COLS)
    .setFontSize(11).setBackground("#B0BEC5");

  // spacer
  ws.setRowHeight(r, 8);
  r++;

  // ─────────────────────────────────────────────────────────
  // NET INCOME ROW
  // ─────────────────────────────────────────────────────────
  ws.setRowHeight(r, 24);
  const firstColLtr  = colLetter(COL_FIRST);
  const lastColLtr2  = colLetter(COL_LAST);
  const totalColLtr2 = colLetter(COL_TOTAL);
  const avgColLtr2   = colLetter(COL_AVG);

  ws.getRange(r, COL_CAT).setValue("NET INCOME (Income − Expenses)");
  ws.getRange(r, COL_BUDGET).setFormula(`=B${incomeTotalRow}-B${grandTotalRow}`);

  for (let ci = 0; ci < numM; ci++) {
    const cLtr = colLetter(COL_FIRST + ci);
    ws.getRange(r, COL_FIRST + ci)
      .setFormula(`=${cLtr}${incomeTotalRow}-${cLtr}${grandTotalRow}`);
  }
  ws.getRange(r, COL_TOTAL)
    .setFormula(`=${totalColLtr2}${incomeTotalRow}-${totalColLtr2}${grandTotalRow}`);
  ws.getRange(r, COL_AVG)
    .setFormula(`=${avgColLtr2}${incomeTotalRow}-${avgColLtr2}${grandTotalRow}`);
  ws.getRange(r, COL_VS).setValue(""); // not applicable

  const netRange = ws.getRange(r, 1, 1, TOTAL_COLS);
  netRange.setBackground("#1A237E")
    .setFontColor(CLR.WHITE)
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  ws.getRange(r, COL_CAT).setHorizontalAlignment("left");

  const fmtCurrency = `${config.currency}#,##0.00`;
  ws.getRange(r, COL_BUDGET, 1, TOTAL_COLS - 1).setNumberFormat(fmtCurrency);
  netRange.setBorder(true, true, true, true, true, true, "#FFF", SpreadsheetApp.BorderStyle.SOLID);
  r++;

  // ── Batch conditional formatting for all vs-Budget cells ──
  // Done once here instead of per-row to avoid timeout from repeated rule reads/writes
  if (vsBudgetRows.length > 0) {
    const vsBudgetRanges = vsBudgetRows.map(rn => ws.getRange(rn, COL_VS));
    const greenRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(0)
      .setBackground(CLR.LTGREEN).setFontColor(CLR.GREEN)
      .setRanges(vsBudgetRanges).build();
    const redRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(CLR.LTRED).setFontColor(CLR.DARKRED)
      .setRanges(vsBudgetRanges).build();
    ws.setConditionalFormatRules([greenRule, redRule]);
  }

  // ── Tip row ───────────────────────────────────────────────
  ws.setRowHeight(r, 36);
  ws.getRange(r, 1, 1, TOTAL_COLS).merge()
    .setValue("HOW TO USE: Log your transactions in the Transactions sheet. " +
              "Monthly actual columns update automatically via live formulas. " +
              "\"vs Budget\" = avg actual − monthly budget (negative = under budget = good for expenses).")
    .setFontSize(8).setFontStyle("italic").setFontColor("#777777")
    .setWrap(true).setVerticalAlignment("middle");
}

// ============================================================
// DASHBOARD SHEET
// ============================================================

/**
 * Creates the Dashboard sheet with KPIs, summary table, category table,
 * and 4 embedded charts.
 */
function createDashboard(ss, config) {
  let ws = ss.getSheetByName(SHEET_DASHBOARD);
  if (ws) ss.deleteSheet(ws);
  ws = ss.insertSheet(SHEET_DASHBOARD, 0);

  const months     = buildMonthList(config.startYear, config.startMonth, config.numMonths);
  const currency   = config.currency;
  const fmtCur     = currency + "#,##0.00";
  const fmtPct     = "0.0%";

  // Compute full-period date bounds for KPI SUMIFS
  const firstMonth = months[0];
  const lastMonth  = months[months.length - 1];
  const periodStart = `DATE(${firstMonth.year},${firstMonth.month},1)`;
  const periodEnd   = `EOMONTH(DATE(${lastMonth.year},${lastMonth.month},1),0)`;

  // Collect category labels
  const incomeLabels  = config.incomeCategories.map(c => c.label);
  const expenseLabels = config.expenseCategories.map(c => c.label);

  // ── Column widths ─────────────────────────────────────────
  ws.setColumnWidth(1, 175);  // A
  ws.setColumnWidth(2, 130);  // B
  ws.setColumnWidth(3, 130);  // C
  ws.setColumnWidth(4, 130);  // D
  ws.setColumnWidth(5, 130);  // E

  // ════════════════════════════════════════════════════════
  // SECTION 1: Title + KPI strip (rows 1–5)
  // ════════════════════════════════════════════════════════
  let r = 1;

  // Title bar
  ws.setRowHeight(r, 40);
  ws.getRange(r, 1, 1, 5).merge()
    .setValue(config.title + " — Dashboard")
    .setBackground(CLR.PRIMARY)
    .setFontColor(CLR.WHITE)
    .setFontWeight("bold")
    .setFontSize(16)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  r++;

  // Subtitle / date range
  ws.setRowHeight(r, 22);
  const startLabel = months[0].label;
  const endLabel   = months[months.length - 1].label;
  ws.getRange(r, 1, 1, 5).merge()
    .setValue(`${startLabel} – ${endLabel}  (${months.length} months)  |  Last updated: ` +
              Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy"))
    .setBackground("#283593")
    .setFontColor("#C5CAE9")
    .setFontSize(9)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  r++;

  // KPI formulas (computed for the full period)
  const totalIncomeFmla  = _sumIfsFullPeriod(incomeLabels,  periodStart, periodEnd);
  const totalExpenseFmla = _sumIfsFullPeriod(expenseLabels, periodStart, periodEnd);
  const netSavingsFmla   = `=${totalIncomeFmla.slice(1)}-(${totalExpenseFmla.slice(1)})`;
  // wrap as formulas
  const kpiIncome  = `=${totalIncomeFmla.slice(1)}`;
  const kpiExpense = `=${totalExpenseFmla.slice(1)}`;
  const kpiNet     = `=B${r + 1}-C${r + 1}`;
  const kpiRate    = `=IFERROR(D${r + 1}/B${r + 1},0)`;

  // KPI label row
  ws.setRowHeight(r, 20);
  [["A", "Total Income"], ["B", "Total Expenses"], ["C", "Net Savings"], ["D", "Savings Rate"]].forEach(([_, lbl], i) => {
    ws.getRange(r, i + 1).setValue(lbl)
      .setBackground(CLR.ACCENT).setFontColor(CLR.WHITE)
      .setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  });
  // Dummy for col E
  ws.getRange(r, 5).setBackground(CLR.ACCENT);
  r++;

  // KPI value row
  ws.setRowHeight(r, 36);
  const kpiData = [
    { fmla: `=IFERROR(${totalIncomeFmla.replace(/^=/,"")},0)`,  fmt: fmtCur, bg: CLR.LTGREEN },
    { fmla: `=IFERROR(${totalExpenseFmla.replace(/^=/,"")},0)`, fmt: fmtCur, bg: CLR.LTRED   },
    { fmla: `=IFERROR(A${r}-B${r},0)`,                          fmt: fmtCur, bg: CLR.LTORANGE },
    { fmla: `=IFERROR(C${r}/A${r},0)`,                          fmt: fmtPct, bg: "#E8EAF6"   },
  ];
  kpiData.forEach(({ fmla, fmt, bg }, i) => {
    ws.getRange(r, i + 1)
      .setFormula(fmla)
      .setNumberFormat(fmt)
      .setBackground(bg)
      .setFontWeight("bold")
      .setFontSize(14)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, null, null, "#9E9E9E", SpreadsheetApp.BorderStyle.SOLID);
  });
  ws.getRange(r, 5).setBackground(CLR.WHITE);
  const kpiValueRow = r;
  r++;

  // Spacer
  ws.setRowHeight(r, 12);
  r++;

  // ════════════════════════════════════════════════════════
  // SECTION 2: Monthly Summary Table (rows ~7–20)
  // ════════════════════════════════════════════════════════
  const monthlySummaryStartRow = r;

  // Section heading
  ws.setRowHeight(r, 24);
  ws.getRange(r, 1, 1, 5).merge()
    .setValue("MONTHLY SUMMARY")
    .setBackground(CLR.ACCENT).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  r++;

  // Column headers
  ws.setRowHeight(r, 22);
  ["Month", "Income", "Expenses", "Net", "Savings Rate"].forEach((h, i) => {
    ws.getRange(r, i + 1)
      .setValue(h)
      .setBackground(CLR.PRIMARY).setFontColor(CLR.WHITE)
      .setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  });
  const monthlyHdrRow = r;
  r++;

  // One row per month
  const monthlyDataStartRow = r;
  months.forEach(mObj => {
    ws.setRowHeight(r, 20);
    const y = mObj.year, mn = mObj.month;
    const ds = `DATE(${y},${mn},1)`;
    const de = `EOMONTH(DATE(${y},${mn},1),0)`;

    // Income for this month
    const incFmla = incomeLabels.length > 0
      ? `=IFERROR(${incomeLabels.map(l => `SUMIFS(Transactions!$D:$D,Transactions!$C:$C,"${l}",Transactions!$A:$A,">="&${ds},Transactions!$A:$A,"<="&${de})`).join("+")},0)`
      : `=0`;

    // Expenses for this month
    const expFmla = expenseLabels.length > 0
      ? `=IFERROR(${expenseLabels.map(l => `SUMIFS(Transactions!$D:$D,Transactions!$C:$C,"${l}",Transactions!$A:$A,">="&${ds},Transactions!$A:$A,"<="&${de})`).join("+")},0)`
      : `=0`;

    const rowBg = (months.indexOf(mObj) % 2 === 0) ? "#FAFAFA" : CLR.WHITE;

    ws.getRange(r, 1).setValue(mObj.label).setBackground(rowBg).setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(r, 2).setFormula(incFmla).setBackground(rowBg).setNumberFormat(fmtCur).setFontColor(CLR.GREEN).setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(r, 3).setFormula(expFmla).setBackground(rowBg).setNumberFormat(fmtCur).setFontColor(CLR.DARKRED).setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(r, 4).setFormula(`=B${r}-C${r}`).setBackground(rowBg).setNumberFormat(fmtCur).setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(r, 5).setFormula(`=IFERROR(D${r}/B${r},0)`).setBackground(rowBg).setNumberFormat(fmtPct).setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(r, 1, 1, 5).setBorder(true, true, true, true, true, true, "#E0E0E0", SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  const monthlyDataEndRow = r - 1;

  // Monthly totals row
  ws.setRowHeight(r, 22);
  ws.getRange(r, 1).setValue("TOTAL").setBackground(CLR.TOTAL_BG).setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.getRange(r, 2).setFormula(`=SUM(B${monthlyDataStartRow}:B${monthlyDataEndRow})`).setBackground(CLR.TOTAL_BG).setNumberFormat(fmtCur).setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.getRange(r, 3).setFormula(`=SUM(C${monthlyDataStartRow}:C${monthlyDataEndRow})`).setBackground(CLR.TOTAL_BG).setNumberFormat(fmtCur).setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.getRange(r, 4).setFormula(`=SUM(D${monthlyDataStartRow}:D${monthlyDataEndRow})`).setBackground(CLR.TOTAL_BG).setNumberFormat(fmtCur).setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.getRange(r, 5).setFormula(`=IFERROR(D${r}/B${r},0)`).setBackground(CLR.TOTAL_BG).setNumberFormat(fmtPct).setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle");
  ws.getRange(r, 1, 1, 5).setBorder(true, true, true, true, true, true, "#9E9E9E", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  r++;

  // Spacer
  ws.setRowHeight(r, 14);
  r++;

  // ════════════════════════════════════════════════════════
  // SECTION 3: Category Spending Table (rows ~22–40)
  // ════════════════════════════════════════════════════════
  const catTableStartRow = r;

  ws.setRowHeight(r, 24);
  ws.getRange(r, 1, 1, 5).merge()
    .setValue("EXPENSE CATEGORY BREAKDOWN")
    .setBackground(CLR.ACCENT).setFontColor(CLR.WHITE)
    .setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("left").setVerticalAlignment("middle");
  r++;

  // Column headers
  ws.setRowHeight(r, 22);
  ["Category", "Budget/mo", "Total Actual", "Avg/mo", "Over/Under"].forEach((h, i) => {
    ws.getRange(r, i + 1)
      .setValue(h)
      .setBackground(CLR.PRIMARY).setFontColor(CLR.WHITE)
      .setFontWeight("bold").setFontSize(9)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  });
  const catHdrRow = r;
  r++;

  const catDataStartRow = r;
  config.expenseCategories.forEach((cat, idx) => {
    ws.setRowHeight(r, 20);
    const y1 = firstMonth.year, m1 = firstMonth.month;
    const y2 = lastMonth.year,  m2 = lastMonth.month;
    const ds = `DATE(${y1},${m1},1)`;
    const de = `EOMONTH(DATE(${y2},${m2},1),0)`;

    const actFmla = `=IFERROR(SUMIFS(Transactions!$D:$D,Transactions!$C:$C,"${cat.label}",Transactions!$A:$A,">="&${ds},Transactions!$A:$A,"<="&${de}),0)`;
    const rowBg   = idx % 2 === 0 ? "#FAFAFA" : CLR.WHITE;

    ws.getRange(r, 1).setValue(cat.label).setBackground(rowBg).setFontSize(9).setVerticalAlignment("middle");
    ws.getRange(r, 2).setValue(cat.budget).setBackground(rowBg).setNumberFormat(fmtCur).setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(r, 3).setFormula(actFmla).setBackground(rowBg).setNumberFormat(fmtCur).setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(r, 4).setFormula(`=IFERROR(C${r}/${months.length},0)`).setBackground(rowBg).setNumberFormat(fmtCur).setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(r, 5).setFormula(`=IFERROR(D${r}-B${r},"")`).setBackground(rowBg).setNumberFormat(fmtCur).setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle");
    ws.getRange(r, 1, 1, 5).setBorder(true, true, true, true, true, true, "#E0E0E0", SpreadsheetApp.BorderStyle.SOLID);
    r++;
  });

  const catDataEndRow = r - 1;

  // ════════════════════════════════════════════════════════
  // SECTION 4: Charts — deferred to keep initial setup fast
  // Run  💰 Budget Tools › Refresh Dashboard  to generate.
  // ════════════════════════════════════════════════════════
  ws.setRowHeight(r, 32);
  ws.getRange(r, 1, 1, 5).merge()
    .setValue("📊  Charts will appear here — run  💰 Budget Tools › Refresh Dashboard  to generate them.")
    .setBackground("#E8EAF6").setFontColor(CLR.ACCENT)
    .setFontStyle("italic").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
}

/**
 * Inserts 4 charts on the Dashboard sheet.
 */
function _insertDashboardCharts(ws, months, monthlyHdrRow, monthlyDataStartRow, monthlyDataEndRow,
                                 catHdrRow, catDataStartRow, catDataEndRow, dataEndRow) {
  const PX_PER_ROW = 21; // approximate
  const chartTop   = (dataEndRow + 2) * PX_PER_ROW;
  const chartW     = 440;
  const chartH     = 280;
  const gap        = 20;

  // Chart 1: Monthly Income vs Expenses — Clustered Column
  try {
    const dataRange1 = ws.getRange(monthlyDataStartRow, 1, months.length, 3); // Month | Income | Expenses
    const chart1 = ws.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(dataRange1)
      .setPosition(dataEndRow + 2, 1, 0, 0)
      .setOption("title", "Monthly Income vs Expenses")
      .setOption("legend", { position: "bottom" })
      .setOption("hAxis", { title: "Month" })
      .setOption("vAxis", { title: "Amount (" + ws.getParent().getName() + ")" })
      .setOption("series", {
        0: { color: "#2E7D32" },  // Income - green
        1: { color: "#C62828" },  // Expenses - red
      })
      .setOption("width",  chartW)
      .setOption("height", chartH)
      .build();
    ws.insertChart(chart1);
  } catch (e) { /* charts may fail in certain script contexts — not fatal */ }

  // Chart 2: Monthly Net Savings — Line
  try {
    const dataRange2a = ws.getRange(monthlyDataStartRow, 1, months.length, 1); // Month labels
    const dataRange2b = ws.getRange(monthlyDataStartRow, 4, months.length, 1); // Net
    const chart2 = ws.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dataRange2a)
      .addRange(dataRange2b)
      .setPosition(dataEndRow + 2, 4, 0, 0)
      .setOption("title", "Monthly Net Savings")
      .setOption("legend", { position: "none" })
      .setOption("hAxis", { title: "Month" })
      .setOption("vAxis", { title: "Net Savings" })
      .setOption("colors", ["#3949AB"])
      .setOption("width",  chartW)
      .setOption("height", chartH)
      .build();
    ws.insertChart(chart2);
  } catch (e) { /* not fatal */ }

  // Chart 3: Expense by Category — Pie (top 10 by actual)
  try {
    const numCats  = Math.min(catDataEndRow - catDataStartRow + 1, 10);
    const dataRange3a = ws.getRange(catDataStartRow, 1, numCats, 1); // Category labels
    const dataRange3b = ws.getRange(catDataStartRow, 3, numCats, 1); // Total Actual
    const chart3 = ws.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dataRange3a)
      .addRange(dataRange3b)
      .setPosition(dataEndRow + 2 + Math.ceil(chartH / PX_PER_ROW) + 2, 1, 0, 0)
      .setOption("title", "Expense by Category (Top 10)")
      .setOption("legend", { position: "right" })
      .setOption("pieSliceText", "percentage")
      .setOption("width",  chartW)
      .setOption("height", chartH)
      .build();
    ws.insertChart(chart3);
  } catch (e) { /* not fatal */ }

  // Chart 4: Budget vs Actual — Bar
  try {
    const numCats = catDataEndRow - catDataStartRow + 1;
    const dataRange4a = ws.getRange(catDataStartRow, 1, numCats, 1); // Category
    const dataRange4b = ws.getRange(catDataStartRow, 2, numCats, 1); // Budget/mo
    const dataRange4c = ws.getRange(catDataStartRow, 4, numCats, 1); // Avg/mo actual
    const chart4 = ws.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dataRange4a)
      .addRange(dataRange4b)
      .addRange(dataRange4c)
      .setPosition(dataEndRow + 2 + Math.ceil(chartH / PX_PER_ROW) + 2, 4, 0, 0)
      .setOption("title", "Budget vs Actual (Avg/mo)")
      .setOption("legend", { position: "bottom" })
      .setOption("hAxis", { title: "Amount" })
      .setOption("vAxis", { title: "Category" })
      .setOption("series", {
        0: { color: "#1E88E5" },  // Budget - blue
        1: { color: "#E53935" },  // Actual - red
      })
      .setOption("width",  chartW)
      .setOption("height", chartH)
      .build();
    ws.insertChart(chart4);
  } catch (e) { /* not fatal */ }
}

// ============================================================
// REFRESH DASHBOARD
// ============================================================

/**
 * Deletes all charts on Dashboard and redraws it from saved settings.
 * Called from the menu item "↺ Refresh Dashboard".
 */
function refreshDashboard() {
  const config = getSettings();
  if (!config) {
    SpreadsheetApp.getUi().alert(
      "No settings found.\nRun the Setup Wizard first:\n💰 Budget Tools → Run Setup Wizard"
    );
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(SHEET_DASHBOARD);
  if (ws) {
    // Remove all existing charts
    ws.getCharts().forEach(chart => ws.removeChart(chart));
  }

  createDashboard(ss, config);
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(
    ss.getSheetByName(SHEET_DASHBOARD)
  );
}

// ============================================================
// PRIVATE HELPERS
// ============================================================

/**
 * Returns a flat array of all income + expense category labels.
 */
function _allCategoryLabels(config) {
  const labels = [];
  (config.incomeCategories  || []).forEach(c => labels.push(c.label));
  (config.expenseCategories || []).forEach(c => labels.push(c.label));
  return labels;
}

/**
 * Returns expense categories grouped, preserving group order.
 * Returns: [{ group, categories: [{label, budget}] }]
 */
function _getExpenseGroups(config) {
  const groupMap = new Map(); // preserves insertion order
  (config.expenseCategories || []).forEach(cat => {
    const g = cat.group || "Other";
    if (!groupMap.has(g)) groupMap.set(g, []);
    groupMap.get(g).push(cat);
  });
  const result = [];
  groupMap.forEach((cats, group) => result.push({ group, categories: cats }));
  return result;
}

/**
 * Builds a SUMIFS formula that sums all given category labels
 * over the full period (periodStart / periodEnd are DATE() formula strings).
 * Returns a formula string starting with "=".
 */
function _sumIfsFullPeriod(labels, periodStart, periodEnd) {
  if (!labels || labels.length === 0) return "=0";
  const parts = labels.map(l =>
    `SUMIFS(Transactions!$D:$D,Transactions!$C:$C,"${l}",` +
    `Transactions!$A:$A,">="&${periodStart},` +
    `Transactions!$A:$A,"<="&${periodEnd})`
  );
  return `=IFERROR(${parts.join("+")},0)`;
}

