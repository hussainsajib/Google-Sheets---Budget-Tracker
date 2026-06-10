# Next Steps — remaining manual setup

The repo is scaffolded and tests pass (7/7). These are the steps that need YOUR
Google/GitHub credentials, so they couldn't be automated. Do them in order.
(Full reference: `DEVELOPMENT.md`.)

## ☐ 1. Connect clasp to Google  (one time)
```bash
cd "C:\Works\pet projects\BudgetTracker"
npx clasp login
```
Then link a project:
```bash
# existing bound script:
npx clasp clone <SCRIPT_ID>      # ID from Sheet → Extensions → Apps Script → Project Settings
# OR new standalone project:
npx clasp create --type sheets --title "Budget Tracker" --rootDir src
```
If clasp errors about the API, enable it once at:
https://script.google.com/home/usersettings  → Apps Script API → ON

Confirm `.clasp.json` exists and contains `"rootDir": "src"`.

## ☐ 2. First push to Apps Script
```bash
npm run push
```
Refresh the Sheet → the **💰 Budget Tools** menu should appear.

## ☐ 3. Verify the daily loop works
```bash
npm test          # 7/7 should pass
npm run watch     # auto-push on save (Ctrl+C to stop)
```

## ☐ 4. Push to GitHub
```bash
git remote add origin https://github.com/<you>/budget-tracker.git
git push -u origin master
```

## ☐ 5. Enable CI auto-deploy
1. Open `C:\Users\Hussain.Sajib\.clasprc.json` (created by `clasp login`) and copy ALL its contents.
2. GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**
   - Name: `CLASPRC_JSON`
   - Value: paste the file contents
3. Make sure the real `.clasp.json` is committed (it holds the scriptId):
   ```bash
   git add .clasp.json && git commit -m "Add clasp project config" && git push
   ```
Now: merge to `main` → CI runs tests → if green, auto-deploys to Apps Script.

## ☐ 6. (Optional) Grow test coverage
Move more pure logic into `src/lib.js` and add Jest tests. Good candidates:
- `parseSettings(rows)` — split the parsing half out of `getSettings()`
- a `computeReserve()` helper for cash-flow math

---
**Reminder:** keep logic in `src/lib.js` (fast-tested, CI-protected) and keep
`src/Code.gs` thin (just SpreadsheetApp calls). That's the layer where past bugs lived.
