# BudgetTracker — Google Sheets Budget App

A production-ready, generic personal budget tracker that runs entirely inside Google Sheets via Apps Script. Designed to be distributed/sold to the public — no hard-coded personal data. A user sets it up once through a wizard, then logs transactions and the dashboard/tracker update via live formulas.

## Files

| File | Purpose |
|------|---------|
| `Code.gs` | All Apps Script: menu, setup wizard backend, sheet builders, live SUMIFS formulas, dashboard + charts, Input form, installable trigger |
| `SetupWizard.html` | 4-step HTML setup dialog (Basic Info → Income → Expenses → Review) with a live budget bar and per-group subtotals |
| `README.md` | End-user setup/deployment instructions |

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

## Menu (`onOpen`)
```
💰 Budget Tools
  ▶ Run Setup Wizard
  ↺ Rebuild Tracker Rows      (after adding/removing categories in Config)
  ↺ Refresh Dashboard         (also builds the charts)
  ⚙ Rebuild All Sheets        (after changing dates/currency/title)
  🔔 Set Up Trigger (run once) (makes Input submit work for all editors)
```

## Known gotchas (bugs already fixed — don't reintroduce)

- **Apps Script execution timeout** (`Service Spreadsheets failed...`): caused by too many individual API calls. FIX: batch formula writes with `setFormulas([[...]])` (one call per row, not per cell); avoid `applyRowBanding` on 999 rows (use a single conditional-format rule); defer chart building out of initial setup.
- **Conditional format builder has no `setFontWeight`** — use `.setBold(true)`.
- **Can't freeze columns through a merged cell** — title/section rows are merged across all columns, so don't `setFrozenColumns` on the Budget Tracker.
- **Timezone bug in date→month mapping**: use `Utilities.formatDate(date, tz, "yyyy-M")`, not `Date.getMonth()` (UTC dates shift a day in negative-offset zones).
- **HTML wizard**: attach event listeners after rendering DOM (`attachExpenseListeners()` must be called at the end of `renderExpenses`); declare `safeGroup` before using it in template strings; the budget bar must refresh when entering Step 3.

## Development & testing (planned setup — not yet scaffolded)

Recommended toolchain (see prior discussion):
1. **`clasp`** for local↔Apps Script sync + **git** for version control. `clasp push --watch` for live upload.
2. **Three test layers:**
   - Unit (Jest, local): pure functions — `colLetter`, `buildMonthList`, category normalization, the `getSettings` parser. These have no `SpreadsheetApp` dependency.
   - Integration (in-GAS): GasT / QUnitGS2 for sheet builders + SUMIFS.
   - E2E: `clasp run` to invoke functions headlessly.
3. **GitHub Actions CI**: run Jest on push/PR; `clasp push` on main (store clasp creds as a secret).
4. **Refactor for testability**: keep `SpreadsheetApp` code thin; move pure logic into a `lib.gs` that Jest can import.

## Deployment (end user)
Paste `Code.gs` into Apps Script, add an HTML file named `SetupWizard`, save, refresh the sheet → `💰 Budget Tools → Run Setup Wizard`. Then `Set Up Trigger (run once)` if sharing with others.

## Notes
- This folder was split out of `../personal_finance/` to keep it generic and sellable. Personal budget data (Sajib's actual numbers) stays in `personal_finance/`, NOT here.
