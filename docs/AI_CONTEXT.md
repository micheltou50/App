# Glenhaven Property Manager — AI Context

> Shared context for AI assistants (Claude, ChatGPT/Codex). Describes what is true in the code at the commit below. Anything not confirmed in the code is marked **Unverified**.

| | |
|---|---|
| Repository | micheltou50/App (public) |
| Default branch | main |
| Verified against | `d5ba06d` (2026-03-20) |
| Last reviewed | 2026-10-25 |
| Related repos | micheltou50/STRManagement (StayOps). Its code was carried over from this repo; see §11. |

## 1. Purpose

A single-property short-term-rental (STR) management web app for the **Glenhaven** property. The name comes from `index.html` `<title>Glenhaven — Property Manager</title>` and from `package.json` (`"name": "glenhaven"`, `"description": "Glenhaven property manager"`).

It covers:
- bookings, with Airbnb/VRBO/Direct platform icons
- cleaner scheduling and notifications
- monthly payout and management-fee views
- management invoices (PDF)
- property expenses with receipt OCR, maintenance and inventory
- AI pricing and expense-review helpers
- a PIN-gated "cleaner mode" for cleaners

The source of truth is a **Google Sheet / Google Apps Script** backend plus browser `localStorage`. There is **no Supabase**.

**Status: superseded.** The code shows this app is the early prototype that StayOps (`micheltou50/STRManagement`) grew out of, and it has not been developed since March 2026 (§11, §12).

## 2. Technology stack

| Layer | What | Where |
|---|---|---|
| Frontend | Static HTML + vanilla JavaScript (global functions, `onclick=` handlers), no framework, no bundler | `index.html`, `assets/js/main.js` (about 6.2k lines), `assets/js/ai-logic.js` |
| Styling | Hand-written CSS; Google Fonts (DM Serif Display, DM Sans) | `assets/css/styles.css`, `index.html` |
| PDF | jsPDF 2.5.1 + jspdf-autotable 3.8.2 from cdnjs | `index.html` `<script>` tags; `buildInvoicePDF`, `exportReportPDF` in `main.js` |
| Serverless | Netlify Functions (Node, CommonJS `exports.handler`) | `Netlify/functions/` |
| Push | Web Push (`web-push` ^3.6.7) + service worker | `Netlify/functions/send-push.js`, `sw.js` |
| Storage | Browser `localStorage` (`gh-*` keys) + Google Sheet / Apps Script web app | `assets/js/main.js` |

There is no manifest.json. The service worker handles push notifications only (no caching or offline support). There are no tests and no CI.

## 3. System architecture

```
Browser (index.html + assets/js/main.js)
 ├─ localStorage  gh-bookings, gh-cleans, gh-notes, gh-expenses, gh-cleaners, … (working copy)
 ├─ Published Google Sheet CSV (SHEET_URL) ── read bookings
 ├─ Google Apps Script web app (DEFAULT_SCRIPT_URL or gh-script-url) ── write bookings,
 │     app data (cleans/notes/cleaners/inventory/…), expenses, backups, cleaner emails
 ├─ Google OAuth (implicit flow, user-supplied client ID) ── Drive (receipts), Calendar (bookings)
 ├─ /.netlify/functions/ai        ── Anthropic Messages API (server-held key)
 ├─ /.netlify/functions/send-push ── web-push to owner/cleaner subscriptions
 └─ /.netlify/functions/send-email (Resend) ── present but no longer called by the frontend
```

The Apps Script source code is **not** in this repo, so its behaviour is **Unverified** beyond the actions the client calls.

## 4. Repository layout

```
index.html                 All markup: nav, sections, modals, settings panels; loads CSS/JS
assets/css/styles.css      Styles (green "forest/sage/moss" theme)
assets/js/main.js          All app logic (sections listed below)
assets/js/ai-logic.js      window.AIService.request(): POST to /.netlify/functions/ai
sw.js                      Service worker: push display and notification click
Netlify/functions/ai.js         Anthropic proxy (model allow-list, token cap)
Netlify/functions/send-email.js Resend email sender (unused by the current frontend)
Netlify/functions/send-push.js  Web Push sender (VAPID)
netlify.toml               Netlify config (publish ".", functions dir "Netlify/functions")
package.json               Only dependency: web-push
.gitignore                 node_modules/
.DS_Store, */.DS_Store     Committed macOS metadata (stray)
```

The section headers in `assets/js/main.js` act as a map of the file:
- config, push notifications, sheet helpers, state, sync from sheet
- nav, render, smart pricing, booking detail, actions, CSV import
- notify cleaner, settings, expenses, AI expense analyser
- backup & restore, receipt photo reader, Google Drive auth, Google Calendar sync
- maintenance, inventory, screenshot to booking, app data sync (Sheet ↔ localStorage)
- cleaner mode, PIN entry, email templates, cleaner access settings, report export

## 5. Key features and how they work

The navigation (`index.html`) has these tabs: **Dashboard, Bookings, Cleaning, Payout, Mgmt, Property, Settings**. There is also a `section-notes`. `showSection(name)` toggles `section-*` elements.

- **Bookings sync** (`syncFromSheets`)
  - Fetches the published sheet as CSV. Columns are found by header name where possible, otherwise by position (see §8).
  - Merges with local bookings by guest name + check-in date and warns about duplicate rows in the sheet.
  - Add, edit and delete go back to the sheet through the Apps Script actions `add`, `update` and `delete` (`pushToSheet`). A full push uses a hidden form POST with `replaceAll` (`syncToSheet`).
- **Manual and CSV booking entry**: `importCSV` uses the same positional columns as the sheet.
- **AI booking import** (`extractBookingFromScreenshot`): a screenshot of a booking confirmation is resized to at most 4000 px and sent to `claude-sonnet-4-20250514` through the AI proxy, which returns JSON fields.
- **Cleaning**
  - Cleans are linked to bookings, with states pending / accepted / declined / done.
  - The owner can notify a cleaner by SMS (`sms:` link built from an editable template), by email through the Apps Script `sendEmail` action using editable HTML templates (assignment and reminder), or by web push.
- **Cleaner mode**
  - The cleaner's link has the form `#cleaner/<id>/<base64 PIN>` (`cleanerLinkForId`). The cleaner enters a 4-digit PIN on a PIN pad (`verifyCleanerPin`).
  - Cleaner mode shows the cleaner's cleans and the inventory, pulls app data from the Apps Script, and offers push opt-in.
- **Payout / Mgmt**
  - Monthly host payout, cleaning fees, management fees and net payout (`renderRevenue`).
  - The management view lets you select bookings and generate an invoice PDF (`generateInvoice` → `buildInvoicePDF`), using invoice and bank details and client records kept in `localStorage`.
  - An FY (July–June) performance report can be exported as PDF (`exportReportPDF`).
- **Property**
  - Tabs: expenses, maintenance, inventory.
  - Expenses can be filled in from a receipt image or PDF read by `claude-sonnet-4-20250514`. The receipt is uploaded to Google Drive (a `drive.file` scope folder) and shared as "anyone with link: reader" (`setDriveFilePublic`). The expense is then synced to the Apps Script (`addExpense`, `updateExpense`, `deleteExpense`).
  - Inventory has a shopping-list view. Maintenance tracks contractors.
- **AI helpers** (all go through `AIService.request`, with JSON-only prompts)
  - Smart pricing (`getSmartPricing`, `claude-haiku-4-5-20251001`): seasonal rate periods, with a prompt hard-coded to the Glenhaven property and its local market.
  - Expense analyser (`analyseExpenses`, Haiku 4.5, assistant prefill `{`): duplicates, anomalies, missing receipts, recurring charges. It honours a user "AI ignore list".
- **Google Calendar**: `pushBookingToCalendar` creates all-day events on the primary calendar and deletes them when the booking is deleted.
- **Backup & restore**: `saveBackup` / `restoreBackup` through the Apps Script. The payload holds bookings, cleans, notes, expenses (photos stripped), maintenance, inventory, cleaners and selected settings.
- **UX extras**: swipe-to-go-back, action sheets, drag-to-dismiss modals, custom modals (because `confirm`/`prompt` are blocked), and a storage viewer.

## 6. Data model and Supabase integration

**Not applicable.** No Supabase client, URL or project ref appears anywhere in this repo.

Data lives in:
- **localStorage**, used as the working copy:
  - `gh-bookings`, `gh-cleans`, `gh-notes`, `gh-expenses`, `gh-maintenance`, `gh-inventory`, `gh-cleaners`, `gh-clients`
  - `gh-push-subs`, `gh-expense-cats`, `gh-sms-template`
  - `gh-inv-*` (invoice details), `gh-bank-*` (bank details for invoices)
  - `gh-script-url`, `gh-gdrive-client-id`, `gh-drive-token` / `gh-drive-token-expiry`
  - `gh-ai-ignore`, `gh-cleaner-authed-<id>`, `gh-last-*`
- **Google Sheet**: the booking rows.
- **Apps Script "app data" store**, written with `setAppData` per key: `cleans`, `notes`, `cleaners`, `inventory`, `maintenance`, `aiIgnore`, `emailTplAssignment`, `emailTplReminder`, `pushSubs`. Read with `?action=getAppData`.

Booking object fields (from `syncFromSheets`): `id`, `checkin`, `checkout`, `nights`, `name`, `guests`, `hostPayout`, `cleaningFee`, `mgmtFee`, `mgmtFeeRaw` (%), `mgmtPayout`, `netPayout`, `cleanerConfirmed`, `platform`, `confirmCode`, `status`, `gcalEventId`.

## 7. External integrations and APIs

| Service | Used for | Where | Config |
|---|---|---|---|
| Google Sheets (published CSV) | Read bookings | `SHEET_URL` constant, `syncFromSheets` | Hard-coded URL |
| Google Apps Script web app | Booking writes, app-data sync, expenses, backup/restore, sending cleaner emails | `DEFAULT_SCRIPT_URL` constant; overridable per device via `gh-script-url` | Hard-coded default URL |
| Google OAuth 2.0 (implicit, `response_type=token`) | Drive + Calendar access | `connectGoogleDrive` | User enters an OAuth client ID in Settings (`gh-gdrive-client-id`); scopes `drive.file`, `calendar.events` |
| Google Drive API v3 | Receipt upload, folder, public-link permission | `main.js` (expenses) | OAuth token |
| Google Calendar API v3 | Create/delete booking events | `pushBookingToCalendar` | OAuth token |
| Anthropic Messages API | Pricing, expense analysis, receipt and screenshot extraction | `Netlify/functions/ai.js` ← `assets/js/ai-logic.js` | `ANTHROPIC_API_KEY`. Model allow-list: `claude-haiku-4-5-20251001`, `claude-sonnet-4-20250514`. `max_tokens` capped at 8000. |
| Web Push (VAPID) | Owner and cleaner notifications | `Netlify/functions/send-push.js`, `sw.js`, `subscribeToPush` | `VAPID_PUBLIC_KEY`, `VAPID_PRIVATE_KEY`, `VAPID_SUBJECT` (the public key is also embedded client-side, as expected) |
| Resend | Cleaner emails (legacy path) | `Netlify/functions/send-email.js` | `RESEND_API_KEY`, `RESEND_FROM` |
| jsPDF / autotable (cdnjs) | Invoice and report PDFs | `index.html` | — |

## 8. Important business rules

- **Sheet columns.** Position A–N is the default; headers override where recognised:
  - check-in, nights, check-out, guest name, guests
  - host payout, cleaning fee, mgmt %, mgmt payout, net payout
  - cleaner confirmed, platform, confirmation code, status
- **Management %** may be written as `15%`, `15` or `0.15`. All three are normalised to a decimal.
  - `mgmtFee = hostPayout × mgmt%`
  - `netPayout` = the sheet value, or else `hostPayout − cleaningFee − mgmtFee`
- **Status**: `cancelled` if the sheet status contains "cancel". Otherwise `completed` if check-in is before now, else `upcoming` (`getStatus`). Cancelled bookings are excluded from revenue, pricing and report totals.
- **Booking identity on sync**: guest name + check-in date (case-insensitive name). Duplicate sheet rows prompt a warning, and the first occurrence is kept.
- **Financial year** runs July–June (`reportFY`, `fyMonths`).
- **Default expense categories** (`DEFAULT_EXPENSE_CATS`):
  - Cleaning & Garden; Maintenance & Repairs; Supplies & Consumables
  - Utilities & Rates; Insurance; Furnishings & Equipment
  - Renovation; Professional Services; Other
  - A silent migration (`migrateCats`) maps older category names onto these.
- **Clean-state merge** on pulling app data: the local "most advanced" state wins (confirmed/declined/done over the sheet copy).
- **Cache reset**: if the configured Apps Script URL changes, the local `gh-expenses` cache is cleared on load.

## 9. Deployment architecture

- `netlify.toml`: `publish = "."`, `command = "npm install"` (no build step), `functions = "Netlify/functions"`. `/sw.js` is served with `Service-Worker-Allowed: /` and `Cache-Control: no-cache`.
- Functions: `ai`, `send-email`, `send-push`. There are no scheduled functions.
- The function code uses `glenhaven21.netlify.app` as the default sender and VAPID subject domain. That site is **not** among the Netlify sites listed in the 2026-10-25 audit.
- **Unverified:** whether this repo is deployed to any Netlify site today. No site named `app` exists.
- No CI and no mobile build.

## 10. Configuration and environment variables

Netlify Function environment variables (names only):

| Name | Used by | Purpose |
|---|---|---|
| `ANTHROPIC_API_KEY` | `ai.js` | Anthropic API key for the proxy |
| `RESEND_API_KEY` | `send-email.js` | Resend API key |
| `RESEND_FROM` | `send-email.js` | Sender address (has a code default) |
| `VAPID_PUBLIC_KEY` | `send-push.js` | Web Push public key |
| `VAPID_PRIVATE_KEY` | `send-push.js` | Web Push private key |
| `VAPID_SUBJECT` | `send-push.js` | VAPID contact (has a code default) |

Client-side settings entered in the UI and kept in `localStorage`:
- the Apps Script URL override
- the Google OAuth client ID
- invoice and bank details, SMS template, cleaning-fee default, expense categories

The Settings panel has an "Anthropic API Key" field (`gh-api-key`), but nothing reads it any more because AI calls go through the proxy. `saveGeminiKey()` references an input that is not in `index.html`.

## 11. Significant design decisions

- **Google Sheets as the backend.** It started as one large `index.html` (`5b3d9fc`, 2026-03-12), edited mainly through the GitHub web editor (many "Update index.html" commits).
- **AI moved server-side** on 2026-03-13 (`7aa3145` "ai"). Before that, history shows direct browser calls with `anthropic-dangerous-direct-browser-access`. After it, all calls go through the Netlify proxy so the key is not exposed in the browser.
- **Email moved from the Resend function to the Apps Script `sendEmail` action** on 2026-03-13 (`3ad2039`). The Resend function was left in place.
- **Codex refactor** (PRs #2–#4, 2026-03-13 to 2026-03-15): CSS and JS were split out of `index.html` into `assets/`, followed by mobile UI polish, a green theme, root-relative asset URLs and fixes to the cleaning flow.
- **Lineage to StayOps (verified by code comparison):**
  - The earliest commit in the local STRManagement clone is `1c0490f` (2026-03-25; that history is shallow, so earlier commits could exist).
  - In that commit, `Netlify/functions/ai.js`, `Netlify/functions/send-email.js`, `assets/js/ai-logic.js` and `netlify.toml` are byte-identical to this repo's. `send-push.js` is an extended version.
  - Its `assets/js/main.js` has the same function names (`syncFromSheets`, `verifyCleanerPin`, `cleanerLinkForId`, `getSmartPricing`, `analyseExpenses`, `extractBookingFromScreenshot`, `buildBackupPayload`, `pushBookingToCalendar`, `sendCleanerEmail`, `exportReportPDF`) and the same `gh-` localStorage keys.
  - What it adds: `package.json` is renamed from `glenhaven` to `stayops` and gains `@supabase/supabase-js`; there is a new `assets/js/supabase.js` pointing at the `stayops` project; and `SHEET_URL`/`DEFAULT_SCRIPT_URL` are emptied.
  - So StayOps continued this codebase and moved it onto Supabase. The verified `stayops` Supabase project was created 2026-03-20, the same day as this repo's last commit.

## 12. Current development status

- **Superseded / inactive.**
  - The last code change was 2026-03-15 (`d900546`, `2731f24`, merged as PR #4).
  - The last commit, `d5ba06d` (2026-03-20, message "w"), only touches `.DS_Store` files.
  - All four non-main branches are fully merged into main. There are no open PRs.
- Active development of this product continues in `micheltou50/STRManagement` (StayOps, https://app.stayops.com.au).

## 13. Known limitations

- **Cleaner "PIN" gives no real protection.** The PIN is base64-encoded inside the cleaner's link and checked client-side against that same link (`verifyCleanerPin`), so anyone with the link has the PIN. Authorisation is stored as a `localStorage` flag.
- **Unauthenticated Netlify functions with `Access-Control-Allow-Origin: *`** (`ai`, `send-push`; `send-email` has no auth). If deployed with keys set, anyone could call them.
- **Uploaded receipts become public.** Drive files are shared as "anyone with the link".
- **Single property and single owner.** Property details and market context are hard-coded in prompts and templates.
- **Data integrity depends on localStorage plus a sheet with positional columns.** Reordering sheet columns can break imports.
- `sw.js` uses `/icon-192.png`, which does not exist in the repo.
- Leftover code: a debug CSV-column dump (`window._csvDebug`), the unused `gh-api-key` and Gemini key settings, and the unused `send-email` function.
- No tests and no build or lint tooling.

## 14. Planned future improvements

None in this repo: no TODO/FIXME markers, no plan docs, no open PRs or unmerged branches. Future work for this product lives in STRManagement.

## 15. Unverified / open questions

- **Unverified:** whether any Netlify site still deploys this repo (`glenhaven21.netlify.app` is referenced only as a default sender/VAPID domain).
- **Unverified:** whether the hard-coded published Google Sheet and Apps Script web app are still live and still hold real booking, cleaner or financial data.
- **Unverified:** the Apps Script implementation (not in the repo), including what `getAppData`, `restoreBackup` and `sendEmail` return and whether they authenticate callers.
- **Unverified:** whether STRManagement was forked from this repo or copied into a fresh repo. The code match is certain; the git mechanism is not visible in the shallow clone.
