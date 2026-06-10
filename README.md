# Budget Tracker — Google Apps Script

A production-ready, zero-dependency personal budget tracker that lives entirely inside Google Sheets. Set it up once with a 4-step wizard, then log transactions and watch the dashboard update automatically via live SUMIFS formulas.

---

## Quick Setup (10 minutes)

### Step 1 — Create a new Google Spreadsheet
1. Open [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet.
2. Give it a name (e.g. "My Budget 2026").

### Step 2 — Open Apps Script
1. In the menu bar click **Extensions → Apps Script**.
2. The script editor opens in a new tab.

### Step 3 — Add the Code
1. In the script editor, click on `Code.gs` in the left-hand file list.
2. Select all the existing placeholder code (`Ctrl+A`) and delete it.
3. Paste the entire contents of **`Code.gs`** from this folder.
4. Save with `Ctrl+S`.

### Step 4 — Add the HTML Wizard file
1. In the script editor, click **File → New → HTML file**.
2. Name it exactly **`SetupWizard`** (no `.html` extension — Apps Script adds it automatically).
3. Select all the placeholder content and delete it.
4. Paste the entire contents of **`SetupWizard.html`** from this folder.
5. Save with `Ctrl+S`.

### Step 5 — Return to the spreadsheet
1. Close the Apps Script tab.
2. Go back to your spreadsheet and **refresh the page** (`F5` or `Ctrl+R`).
3. After the page reloads you will see a new menu: **💰 Budget Tools**.

> The first time you use a custom menu Google will ask you to authorize the script. Click **Review permissions → Allow**.

### Step 6 — Run the Setup Wizard
1. Click **💰 Budget Tools → ▶ Run Setup Wizard**.
2. Follow the 4-step wizard:
   - **Step 1 — Basic Info:** Enter a title, currency symbol (`$`, `€`, `£`, etc.), start month/year, and how many months to track.
   - **Step 2 — Income Sources:** Check your income categories and set monthly amounts.
   - **Step 3 — Expense Categories:** Select expense categories grouped by type and set budgets.
   - **Step 4 — Review & Create:** Confirm the summary and click **✓ Create My Budget**.
3. The wizard creates all four sheets and takes you to the Dashboard.

### Step 7 — Start tracking
Open the **Transactions** sheet and begin logging your income and expenses. The Budget Tracker and Dashboard update automatically as you type.

---

## How It Works

The app creates four sheets:

| Sheet | Purpose |
|-------|---------|
| **Dashboard** | KPI cards, monthly summary table, category breakdown, and 4 charts |
| **Budget Tracker** | All categories with monthly budget vs. actual (SUMIFS live formulas) |
| **Transactions** | Your transaction log — Date, Description, Category, Amount, Account, Notes |
| **_Settings** | Hidden config sheet; stores your setup so sheets can be rebuilt |

### Live SUMIFS Formulas
Every "actual" cell in the Budget Tracker uses a `SUMIFS` formula that reads directly from the Transactions sheet. There is no "refresh" button needed — amounts update the moment you save a transaction.

The formula pattern used is:
```
=IFERROR(SUMIFS(Transactions!$D:$D,
         Transactions!$C:$C, <category>,
         Transactions!$A:$A, ">="&DATE(year, month, 1),
         Transactions!$A:$A, "<="&EOMONTH(DATE(year, month, 1), 0)),
         0)
```

### vs Budget Column
The rightmost column on the Budget Tracker compares your average monthly actual to your budget:
- **Negative (green)** — you spent less than budgeted. Good for expenses.
- **Positive (red)** — you spent more than budgeted. Needs attention.

---

## Logging Transactions

In the **Transactions** sheet, fill in one row per transaction:

| Column | Notes |
|--------|-------|
| **Date** | Format `yyyy-mm-dd`. Type or paste — the dropdown calendar also works. |
| **Description** | Free text: merchant name, payee, etc. |
| **Category** | Choose from the dropdown (all your configured categories). |
| **Amount** | Enter a positive number for **both income and expenses**. The Category determines whether it counts as income or expense. |
| **Account** | Chequing, Savings, Credit Card, Cash, or Other. |
| **Notes** | Optional: recurring flag, merchant location, etc. |

**Delete the 3 example rows** before you start entering real data.

---

## Customizing Categories

You can add or rename categories at any time:

### Option A — Rebuild from wizard (recommended for large changes)
1. **💰 Budget Tools → ▶ Run Setup Wizard** — re-run the wizard with your updated selections.
2. This replaces all sheets. Export your transaction data first if you have existing entries.

### Option B — Add a row manually (for small additions)
1. In the **Budget Tracker** sheet, insert a row in the appropriate section.
2. Type the category name in column A and a monthly budget in column B.
3. Paste the SUMIFS formula pattern from an adjacent row into the monthly columns (adjust the category name to match).
4. In the **Transactions** sheet, add the new category to the Category dropdown:
   - Select the C column data cells → **Data → Data validation → Edit** → add the label to the list.

> **Important:** The category name in the Budget Tracker (column A) must exactly match the category name in the Transactions dropdown. Case and spacing must be identical.

---

## Understanding the Dashboard

### KPI Cards (top strip)
Four headline numbers covering the full configured period:
- **Total Income** — sum of all income transactions
- **Total Expenses** — sum of all expense transactions
- **Net Savings** — Income minus Expenses
- **Savings Rate** — Net Savings ÷ Total Income

### Monthly Summary Table
One row per month. All values are live SUMIFS. Use this to spot months where spending spiked.

### Category Breakdown Table
Every expense category with budget, total actual, average per month, and over/under status.

### Charts (4 total)
1. **Monthly Income vs Expenses** — clustered column chart; shows the income/expense gap each month.
2. **Monthly Net Savings** — line chart; highlights savings trend over time.
3. **Expense by Category (Top 10)** — pie chart; shows where money actually goes.
4. **Budget vs Actual** — bar chart; compares budgeted amount vs. average monthly actual per category.

To refresh charts after adding many transactions: **💰 Budget Tools → ↺ Refresh Dashboard**.

---

## Menu Reference

| Menu item | What it does |
|-----------|-------------|
| **▶ Run Setup Wizard** | Opens the 4-step wizard to configure or reconfigure the tracker |
| **↺ Refresh Dashboard** | Deletes and redraws all Dashboard charts from saved settings |
| **⚙ Rebuild All Sheets** | Rebuilds all four sheets using saved settings (does not change config) |

---

## Tips

- **Income transactions** should use an income category (e.g. "Salary / Wages") and a positive amount. They appear as positive in income totals.
- **Expense transactions** should use an expense category and a positive amount. They appear as positive in expense totals.
- **Savings Rate** is most meaningful once you have a full month of data.
- The **_Settings** sheet is hidden but not protected. Do not edit it manually — use the Setup Wizard instead.
- If you accidentally delete a sheet, run **⚙ Rebuild All Sheets** to restore it.
- For the most accurate "vs Budget" comparison, make sure every transaction has a category.
- The Budget Tracker's "Avg/mo" column uses `COUNTA` to count months that have any transactions, so it adapts automatically as months pass.
- You can add notes, color-code rows, or apply additional conditional formatting without breaking the formulas.
