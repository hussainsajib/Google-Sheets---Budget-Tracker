# Development & Automation Guide

This project uses **clasp** (local ↔ Apps Script sync), **Jest** (local unit tests),
**git** (version control), and **GitHub Actions** (CI: test on every push, auto-deploy on `main`).

## Project layout
```
BudgetTracker/
├── src/                    # everything clasp pushes to Apps Script
│   ├── Code.gs             # entry points + sheet builders (uses SpreadsheetApp)
│   ├── lib.js              # PURE functions — also imported by Jest
│   ├── SetupWizard.html    # the setup dialog
│   └── appsscript.json     # Apps Script manifest
├── test/lib.test.js        # Jest unit tests (run locally + in CI)
├── .github/workflows/ci.yml
├── package.json            # npm scripts
├── jest.config.js
├── .clasp.json.example     # copy to .clasp.json after you have a scriptId
├── .claspignore            # limits what gets pushed
└── .gitignore              # ignores node_modules + .clasprc.json (the secret)
```

The split that makes this work: **pure logic lives in `src/lib.js`** with a
`module.exports` guard. Apps Script ignores the guard and uses the functions globally;
Node/Jest imports them. Keep `lib.js` free of `SpreadsheetApp`/`HtmlService`/`Session`.

---

## ONE-TIME SETUP

### 1. Install dependencies (done)
```bash
npm install
```

### 2. Connect clasp to your Apps Script project
```bash
npx clasp login                 # opens browser; authorizes clasp
```
Then either link an existing project or create a new one:
```bash
# If you already have the Sheet's bound script:
npx clasp clone <SCRIPT_ID>     # SCRIPT_ID from Apps Script → Project Settings
# OR create a new standalone project:
npx clasp create --type sheets --title "Budget Tracker" --rootDir src
```
Either command writes a real `.clasp.json`. If you cloned, make sure it has:
```json
{ "scriptId": "…", "rootDir": "src" }
```
(Copy from `.clasp.json.example` and paste your scriptId if needed.)

### 3. First push
```bash
npm run push                    # clasp push → uploads src/ to Apps Script
```
Open the Sheet, refresh → the **💰 Budget Tools** menu appears.

### 4. Enable the Apps Script API (once, if clasp errors)
Visit https://script.google.com/home/usersettings and turn **Apps Script API** ON.

---

## DAILY DEV LOOP

```bash
npm run watch     # clasp push --watch: every save auto-uploads to Apps Script
npm test          # run unit tests
npm run test:watch  # re-run tests on change
```

Typical flow:
1. Edit `src/lib.js` (pure logic) or `src/Code.gs` (sheet code).
2. `npm test` — fast feedback on the pure logic.
3. `npm run push` (or keep `watch` running) — see it live in the Sheet.
4. `git commit` — version history.

> Put as much logic as possible in `lib.js` so it's covered by fast local tests.
> `Code.gs` should be the thin layer that calls `lib.js` and talks to the Sheet.

---

## GIT + GITHUB

```bash
# already done: git init + first commit
git branch -M main
git remote add origin https://github.com/<you>/budget-tracker.git
git push -u origin main
```

Work on branches, open PRs — CI runs tests automatically on every PR.

---

## CI/CD (GitHub Actions)

`.github/workflows/ci.yml` does two things:
- **test** — runs `npm ci && npm test` on every push and PR.
- **deploy** — on push to `main` (after tests pass), runs `clasp push` to deploy.

### Enable auto-deploy
1. Locally, after `clasp login`, copy your token file contents:
   - Path: `~/.clasprc.json` (Windows: `C:\Users\<you>\.clasprc.json`)
2. GitHub repo → **Settings → Secrets and variables → Actions → New repository secret**
   - Name: `CLASPRC_JSON`
   - Value: paste the entire contents of `.clasprc.json`
3. Commit `.clasp.json` (the scriptId) so CI knows the target project.

Now: merge to `main` → tests run → if green, the code auto-deploys to Apps Script.

---

## TESTING LAYERS

| Layer | Location | Runs | Covers |
|-------|----------|------|--------|
| Unit | `test/*.test.js` (Jest) | local + CI | pure functions in `lib.js` |
| Integration (optional) | in-GAS GasT/QUnitGS2 | manually in Apps Script | sheet builders, SUMIFS |
| E2E (optional) | `clasp run <fn>` | local against real Sheet | full wizard/build |

To grow coverage: move more pure logic into `lib.js` and add tests. Good next
candidates — a pure `parseSettings(rows)` (split out of `getSettings`) and a
`computeReserve()` helper.

---

## Quick reference (npm scripts)
| Command | Does |
|---------|------|
| `npm test` | Jest unit tests |
| `npm run watch` | auto-push to Apps Script on save |
| `npm run push` / `npm run pull` | manual sync |
| `npm run deploy` | test, then push |
| `npm run open` | open the Apps Script editor |
