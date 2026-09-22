# Repository Guidelines

## Project Structure & Module Organization
This is a Google Apps Script (GAS) project managed with `@google/clasp`. All source code lives in `src/`:
- `main.js` — thin entry point exposing GAS entry functions (`onOpen`, menu handlers).
- `bonds.js`, `cashflow.js`, `coupons.js`, `ladder.js`, `analytics.js` — domain logic.
- `sheets.js`, `ui.js` — spreadsheet I/O and dialog UI.
- `config.js` — constants and configuration.
- `utils.js` — shared helpers (e.g. date math, formatting).
- `appsscript.json` — GAS project manifest.

There is no separate test or asset directory; tests do not exist yet. Keep modules in `src/` and expose only GAS-callable functions from `main.js`.

## Build, Test, and Development Commands
- `npx clasp push` — uploads `src/` to the Apps Script project (see `.clasp.json`).
- `npx eslint src/` — lints the source (equivalent of a `lint` script).
- `npm test` — no test suite is configured; do not rely on it.

This project has no build step; edit `.js` files in `src/` and `clasp push` to deploy.

## Coding Style & Naming Conventions
- Indentation: 4 spaces, no tabs.
- Quotes: single quotes; style is enforced by `eslint-config-sdm` via `eslint.config.mjs`.
- Naming: PascalCase for module objects (`Utils`, `Bonds`), camelCase for functions and variables (`addMonthsSafe`), UPPER_CASE for constants.
- Apps Script globals (`SpreadsheetApp`, `Logger`, `HtmlService`) must not be redefined; add new globals to the `globals` list in `eslint.config.mjs`.
- GAS entry points in `main.js` are intentionally exempt from `no-unused-vars`; keep that override in place.

## Testing Guidelines
There is no test framework or coverage requirement in this repo. If you add logic with meaningful branches or date math, add a small standalone test file (e.g. `test/utils.test.js`) runnable with `node --test`; keep it dependency-free.

## Commit & Pull Request Guidelines
- Commit messages: imperative, sentence case, no prefix convention in history (e.g. `Add sell yield to available-bonds coupon opportunity table`). Keep titles short and descriptive.
- PRs: describe the change, note affected sheets/flows, and include a screenshot or summary of visible UI changes when UI is touched. Link the issue being fixed if one exists.

## Security & Configuration Tips
Never commit secrets or sheet IDs beyond what `.clasp.json` already contains. Before `clasp push`, confirm `.clasp.json` points at the intended project to avoid overwriting another deployment.
