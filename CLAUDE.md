# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# BudgetTracker — Google Sheets Budget App

A production-ready, generic personal budget tracker that runs entirely inside Google Sheets via Apps Script. Designed to be distributed/sold to the public — no hard-coded personal data. A user sets it up once through a wizard, then logs transactions and the dashboard/tracker update via live formulas.

## Commands

```bash
# Run unit tests (Jest, local — no GAS runtime needed)
npm test

# Run tests in watch mode
npm run test:watch

# Push src/ to Apps Script (requires .clasp.json with your scriptId)
npm run push

# Push and re-push on every save
npm run watch

# Pull latest from Apps Script into src/
npm run pull

# Run tests then push (CI-style local deploy)
npm run deploy

# Open the bound spreadsheet in browser
npm run open
```

To run a single test file:
```bash
npx jest test/lib.test.js
```

## clasp setup (one-time)

Copy `.clasp.json.example` to `.clasp.json` and replace `PASTE_YOUR_SCRIPT_ID_HERE` with your Apps Script project ID (found in Apps Script → Project Settings). Then `clasp login` once to authenticate.

## Files

| File | Purpose |
|------|---------|
| `src/Code.gs` | All Apps Script: menu, setup wizard backend, sheet builders, live SUMIFS formulas, dashboard + charts, Input form, installable trigger |
| `src/lib.js` | Pure helper functions shared by Apps Script and Jest — no `SpreadsheetApp` dependency |
| `src/SetupWizard.html` | 4-step HTML setup dialog (Basic Info → Income → Expenses → Review) with a live budget bar and per-group subtotals |
| `src/appsscript.json` | Apps Script manifest |
| `test/lib.test.js` | Jest unit tests for `lib.js` |

## The 5 sheets it builds

| Sheet | Role |
|-------|------|
| `Dashboard` | KPI strip + monthly summary + category breakdown + 4 charts (charts built on "Refresh Dashboard") |
| `Input` | Transaction-entry form; check the Submit box → appends to Transactions → resets the form |
| `Budget Tracker` | Monthly actuals via live SUMIFS, Total/Avg/vs-Budget columns, color-coded by group |
| `Transactions` | Raw transaction log (Date, Description, Category, Amount, Account, Notes) |
| `Config` | **Source of truth** — visible, user-editable settings + category table. Edit a budget → tracker updates instantly via VLOOKUP |

## Key architecture

- **Config sheet is the single source of truth.** Settings live in fixed rows (`CFG_SETTINGS_START = 3`), categories from `CFG_CAT_DATA_START = 11`. `getSettings()` reads it back into a config object.
- **Budget amounts link live**: Budget Tracker col B uses `VLOOKUP($A,Config!$A:$D,4,FALSE)`. Change a budget in Config → tracker updates with no rebuild.
- **Monthly actuals are live SUMIFS** against the Transactions sheet by category + month date range — never manually updated.
- **Category dropdowns use `requireValueInRange`** pointing at Config col A, so adding/removing categories updates dropdowns automatically.
- **`lib.js` dual-runtime pattern**: pure functions live in `src/lib.js` with a `typeof module` guard at the bottom. Apps Script ignores the guard; Jest imports via `module.exports`. Keep all `SpreadsheetApp` / `HtmlService` / `Session` calls out of `lib.js`.

## Menu (`onOpen`)
```
💰 Budget Tools
  ▶ Run Setup Wizard
  ↺ Rebuild Tracker Rows      (after adding/removing categories in Config)
  ↺ Refresh Dashboard         (also builds the charts)
  ⚙ Rebuild All Sheets        (after changing dates/currency/title)
  🔔 Set Up Trigger (run once) (makes Input submit work for all editors)
```

## CI / CD

GitHub Actions (`.github/workflows/ci.yml`):
- **On every push/PR**: runs `npm test`.
- **On push to `master`**: runs tests, then `clasp push --force` to Apps Script using the `CLASPRC_JSON` repo secret (store your `~/.clasprc.json` contents there).

## Known gotchas (bugs already fixed — don't reintroduce)

- **Apps Script execution timeout** (`Service Spreadsheets failed...`): caused by too many individual API calls. FIX: batch formula writes with `setFormulas([[...]])` (one call per row, not per cell); avoid `applyRowBanding` on 999 rows (use a single conditional-format rule); defer chart building out of initial setup.
- **Conditional format builder has no `setFontWeight`** — use `.setBold(true)`.
- **Can't freeze columns through a merged cell** — title/section rows are merged across all columns, so don't `setFrozenColumns` on the Budget Tracker.
- **Timezone bug in date→month mapping**: use `Utilities.formatDate(date, tz, "yyyy-M")`, not `Date.getMonth()` (UTC dates shift a day in negative-offset zones).
- **HTML wizard**: attach event listeners after rendering DOM (`attachExpenseListeners()` must be called at the end of `renderExpenses`); declare `safeGroup` before using it in template strings; the budget bar must refresh when entering Step 3.

## Deployment (end user)
Paste `Code.gs` into Apps Script, add an HTML file named `SetupWizard`, save, refresh the sheet → `💰 Budget Tools → Run Setup Wizard`. Then `Set Up Trigger (run once)` if sharing with others.

## Notes
- This folder was split out of `../personal_finance/` to keep it generic and sellable. Personal budget data stays in `personal_finance/`, NOT here.
