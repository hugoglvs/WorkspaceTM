# ADR 0032: Per-Spreadsheet Sync Lock Between Nightly Sync and Live Edits

## Status
Accepted (resolves issue 13)

## Context
1. `nightlySyncJob` (02:00) and the two edit entry points (simple `onEdit`, installable `installedTaskEditHandler`) all write to the same tables. A Member in another timezone can edit while the nightly sync runs.
2. The races that matter are script-vs-script, not user-vs-script:
   - The nightly `syncTasks` / `syncCalendarEvents` read a snapshot, while an edit handler creates the Google Task / Calendar event for the same row: duplicate external items.
   - `compilePresence` and a Weekly Matrix View declaration both append to `tbl_presence_log` at a "next free row" computed from a snapshot.
   - `sync` (directory) and backup-conflict checks run from both paths.
3. Apps Script cannot lock cells against a human typing. It can only serialise script executions (`LockService`).
4. Google's library resource scoping: a lock created in a library "is visible to all including scripts". A **script lock** taken in the WorkspaceTM Library is therefore shared by every Consumer Sheet, and all teams' 02:00 jobs would queue behind each other and time out.

## Decision
1. **Sync Lock** = `LockService.getDocumentLock()` (one per Consumer Sheet), falling back to `getScriptLock()` only when there is no containing document. Helper: `acquireSyncLock(timeoutMs)` in `src/Code.js`.
2. `nightlySyncJob` waits up to 30 s for the Sync Lock. On timeout it runs no step and returns `{ success: false, reason: "lock_timeout" }`. The lock is released in `finally`.
3. `onEdit` and `installedTaskEditHandler` run their script-side work through `runWithEditLock`, waiting up to 10 s (simple triggers are capped at 30 s). On timeout the user's typed value stays in the cell, the follow-up writes are skipped, a "Sync in progress" toast tells the user, and the handler returns `{ skipped: true, reason: "sync_in_progress" }`. The nightly `syncTasks` / `syncCalendarEvents` push rows still missing external IDs on the next run.
4. The duplicate global `installedTaskEditHandler` in `src/TaskSyncService.js` is removed. Apps Script shares one global scope across files, so it could silently replace the canonical `Code.js` version, which also routes `tbl_events` edits and now holds the lock.
5. No Document Property mutex (issue 13, option B). `LockService` is atomic and is released automatically when an execution ends; a property flag adds a crash-stuck state and gives no extra protection.
6. No read/diff/write split (option D) for now. Revisit if the lock hold time becomes a problem.

## Write-discipline audit (issue 13, option C)
| Nightly step | Writes | Verdict |
|---|---|---|
| `sync` | Protected derived columns on existing `tbl_members` rows; new member rows and onboarding task rows at a "next free row" computed once per run | Accepted residual risk: `tbl_members.member` is protected and tenure-driven (ADR 0007), so users rarely type new rows there. The onboarding-task window is one loop over the night's new members. |
| `compilePresence` | Bulk append to `tbl_presence_log`, with the next row computed immediately before one `setValues` | Safe under the Sync Lock. Matrix declarations are now serialised with it. |
| Dashboards | `clearContent` + rewrite of view/KPI sheets only | Safe: views are derived. A Weekly Matrix declaration typed during the sync is skipped (with a toast) and redrawn away. The user re-enters it. |
| `syncTasks` / `syncCalendarEvents` | Per-cell external IDs and status | Safe under the Sync Lock (it removes the duplicate-creation race). |
| `checkProjectBackupConflicts` | Appends to `tbl_tasks` / `tbl_backups` at `getLastRow() + 1`, computed per write | Safe under the Sync Lock. |

## Consequences
- Each edit pays one lock acquisition, which is negligible outside the sync window.
- Edits made during the sync are deferred, never merged against a stale snapshot. The only user-visible cost is re-entering a Weekly Matrix declaration made during the sync.
- Teams' nightly jobs stay independent of each other.
