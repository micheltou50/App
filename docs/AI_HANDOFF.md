# Glenhaven Property Manager — AI Handoff

**Date of last update:** 2026-10-25
**Based on:** `main` @ `d5ba06d` (2026-03-20)
**Updated by:** Claude (initial documentation audit)

## Latest known development status

- **Superseded and inactive.** This is the Google-Sheets-backed prototype that StayOps (`micheltou50/STRManagement`) was built from; see `docs/AI_CONTEXT.md` §11.
- No code has changed since 2026-03-15. The final commit (2026-03-20) only updates `.DS_Store` files.
- **Unverified:** whether it is still deployed anywhere.

## Recently completed features

Newest first, from `git log origin/main`. All dates are 2026.

- **03-20** `d5ba06d` "w": `.DS_Store` only.
- **03-15** PR #4 (`codex/refactor-project-structure-and-separation-196kp4`):
  - booking-detail cleaner assignment routed through the Cleaning tab flow (`31d05cf`)
  - fix for unassigned bookings and a stale confirmation state in the cleaning action tab (`3bdaaeb`)
  - imported clean-assignment states normalised and team-access info restored (`d900546`)
  - `.gitignore` for `node_modules` (`2731f24`)
- **03-14** PR #3 (`codex/refactor-project-structure-and-separation-ngtmnk`):
  - mobile dashboard polish (`3c52675`)
  - root-relative asset URLs to stop 404s on nested routes (`609bb3b`)
  - section-visibility regression fix (`bb51b04`)
  - green dashboard theme (`ee41573`)
  - banners shown only for warnings and errors (`a891aec`)
- **03-13** PR #2 (`codex/refactor-project-structure-and-separation`):
  - CSS and JS split out of `index.html` into `assets/` (`f4136a9`)
  - fixes for missing `totalHost` / `totalMgmtPayout` calculations (`21db2f9`, `221fcd7`)
- **03-13** PR #1 (`test`): `7aa3145` "ai" added the Netlify AI proxy (`Netlify/functions/ai.js`) and moved AI calls server-side.
- **03-13** About 20 "Update index.html" web edits, including `3ad2039`, which switched cleaner emails from the Resend function to the Apps Script `sendEmail` action. `c16eff9` "email" added `Netlify/functions/send-email.js` (Resend).
- **03-12** `5b3d9fc` "new": initial single-file app (`index.html`), `sw.js`, `send-push.js`, `netlify.toml`, `package.json`.

## Current work in progress

None.
- Branches `codex/refactor-project-structure-and-separation` (last commit 2026-03-13), `-ngtmnk` (2026-03-14), `-196kp4` (2026-03-15) and `test` (2026-03-13) are all fully merged into `main`, with no unmerged commits. They are stale and could be deleted.
- No open PRs.

## Outstanding bugs

Ordered by importance. Each item cites its source; inferences are marked.

- **Security: unauthenticated backend URLs in a public repo.**
  - `assets/js/main.js` (`SHEET_URL`, `DEFAULT_SCRIPT_URL`, lines 13–14) hard-codes a published Google Sheet CSV link and a Google Apps Script web-app URL. Both work without any login.
  - If they are still live, anyone can read booking rows (guest names, payouts). Through the Apps Script actions (`getAppData`, `restoreBackup`, `setAppData`, `add`/`update`/`delete`/`replaceAll`, `sendEmail`) they can also read cleaner records (including PINs) and change data.
  - Inferred from code. **Unverified** whether the endpoints are still live.
  - Action: unpublish the sheet, disable or redeploy the Apps Script, and remove the URLs. They remain in git history either way.
- **Security: open Netlify functions.**
  - `Netlify/functions/ai.js` and `send-push.js` allow any origin (`*`) and require no auth. `send-email.js` also requires no auth.
  - If deployed with `ANTHROPIC_API_KEY` / `RESEND_API_KEY` set, third parties could use the owner's Anthropic quota or send email from the configured sender. Inferred.
- **Cleaner PIN is ineffective.** The PIN is embedded base64 in the cleaner link and compared client-side (`verifyCleanerPin`, `cleanerLinkForId` in `assets/js/main.js`).
- **Receipts made public.** `setDriveFilePublic` grants "anyone with link: reader" on every uploaded receipt.
- **Missing icon.** `sw.js` uses `/icon-192.png`, which is not in the repo, so notification icons 404.
- **Dead settings and code:**
  - the `gh-api-key` "Anthropic API Key" field is unused because AI goes through the proxy
  - `saveGeminiKey()` targets a missing `#settings-gemini-key` input
  - `send-email.js` has had no caller since `3ad2039`
  - the debug `window._csvDebug` dump in `syncFromSheets` is still there
- **Pinned models.** `claude-haiku-4-5-20251001` and `claude-sonnet-4-20250514` are hard-coded in both the client and the proxy allow-list. **Unverified** whether they are still served.

No credentials (API keys, private keys, tokens) were found committed. The VAPID key in `main.js` is the public key, which is meant to be public.

## Planned changes

None recorded in this repo. Product work continues in `micheltou50/STRManagement`.

## Important decisions made

- 2026-03-12: Google Sheet + Apps Script chosen as the backend, with localStorage as the working copy.
- 2026-03-13: Anthropic calls moved behind a Netlify Function proxy with a model allow-list and token cap.
- 2026-03-13: Cleaner emails moved from Resend to the Apps Script `sendEmail` action.
- 2026-03-13 to 03-15: Codex refactor split the code into `assets/css` and `assets/js`.
- Around 2026-03-20 to 03-25 (inferred from code comparison and the `stayops` Supabase project's creation date): the codebase moved to STRManagement / StayOps on Supabase. This repo was left as is.

## Pending tests

- No test suite, CI, build or lint tooling exists.
- If the app is ever revived or kept running, check by hand:
  - sheet sync with the real column layout
  - cleaner PIN flow and push opt-in
  - invoice and report PDF generation
  - receipt OCR through the AI proxy with the current models

## Recommended next development tasks

1. **Shut down the exposed backends.** Confirm whether the published Google Sheet and the Apps Script web app named in `assets/js/main.js` are still live. If so, unpublish or redeploy them and rotate any access, because they appear to give unauthenticated access to booking and cleaner data.
2. **Confirm whether this repo is still deployed on Netlify** (the default domain in the functions is `glenhaven21.netlify.app`). If it is, delete that site or remove `ANTHROPIC_API_KEY` / `RESEND_API_KEY` / `VAPID_PRIVATE_KEY` from it, so the open functions cannot be abused.
3. **Archive the repository on GitHub** and point readers to `micheltou50/STRManagement`. Optionally delete the four merged branches.
4. If kept public, remove the hard-coded Sheet and Apps Script URLs from the working tree, and consider purging them from history.

## Update log
- 2026-10-25 — Initial AI_CONTEXT/AI_HANDOFF created from codebase audit.
