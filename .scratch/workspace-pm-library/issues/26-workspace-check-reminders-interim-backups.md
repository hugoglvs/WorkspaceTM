# 26: Full-table reads, one run per edit, Check Workspace, Monday reminder, Interim Backups sidebar

**Status:** resolved

**Blocked by:** _(none)_

**What to build:** ADR 0041.

- [x] `readRowLimit`: no read stops before the last data row (Presence Log duplicates bug).
- [x] Installable edit trigger handles every edit; simple `onEdit` steps aside (`EDIT_TRIGGER_ACTIVE`).
- [x] Check Workspace dialog with fix links and duplicate removal.
- [x] Monday 8 am reminder for Undeclared Days.
- [x] Interim Backups sidebar (add / cancel, gaps prefill dates, leave warning).

## Comments

_2026-10-12 — new `test/presenceLogScale.test.js` (fails on the old code), `test/workspaceCheck.test.js`, 3 tests in `test/syncLock.test.js`; suite 191 pass._
