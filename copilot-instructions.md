# Team progress dashboard

This is an Armenian-language React 19 + TypeScript + Vite 8 app. Keep the existing dashboard design, responsive layout, animations, and Armenian UI unless the user explicitly requests a change.

## Main behavior

- `app/page.tsx` owns the dashboard, monthly filters, team-member records, Excel import/export, and chart.
- `lib/progress.ts` validates names, `YYYY-MM` periods, and nonnegative integer counts.
- Import columns: `Աշխատակից`, `Ամիս`, `Առաջադրանքներ`, `Դիտարկումներ` (English aliases also supported).
- A row is uniquely identified by team-member name and month. Editing a month must not erase other months.
- The existing app stores records in browser `localStorage` under `timi-arajyntac-v1`. This is local to a browser/origin. Never imply that a GitHub Pages deployment makes records shared across users or devices. Do not commit imported spreadsheets or personal team records to this public repository.
- Preserve Excel export, including the full backup button, before changing persistence. If shared storage is requested, propose a secure backend with authentication and migration from localStorage/Excel. Never place service-role keys, write tokens, or passwords in frontend code.

## Working and publishing

- Run `npm ci`, `npm run typecheck`, and `npm run build` after meaningful changes.
- `vite.config.ts` uses relative base paths so this same source works at a GitHub Pages project URL and at the root of a domain.
- `.github/workflows/pages.yml` builds and deploys `dist/` on pushes to `main`. Do not commit `node_modules/` or `dist/`.
- Explain user-facing steps in plain Armenian. The owner has no programming background.
