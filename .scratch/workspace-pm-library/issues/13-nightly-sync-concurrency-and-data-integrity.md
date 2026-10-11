# 13: Nightly Sync Concurrency & Data Integrity

**Status:** resolved

**Blocked by:** _(none)_

---

## Context

`nightlySyncJob()` ([`src/Code.js` L532–603](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/src/Code.js#L532-L603)) is a time-driven trigger that fires daily at 02:00 AM and orchestrates six sequential write operations across multiple sheets:

| Step | Function | Tables Touched |
|------|----------|----------------|
| 1 | `sync()` | `tbl_members`, `tbl_tenures` |
| 2 | `compilePresence()` | `tbl_presence_log`, `tbl_members` |
| 3a | `refreshWeeklyMatrix()` | Weekly Matrix view sheet |
| 3b | `refreshActiveResponsibles()` | Active Responsibles view sheet |
| 3c | `refreshKpis()` | KPI sheet |
| 4 | `syncTasks()` | `tbl_tasks` |
| 5 | `syncCalendarEvents()` | `tbl_events` |
| 6 | `checkProjectBackupConflicts()` | `tbl_backups`, `tbl_projects` |

The user raised the concern that a Member in a different timezone might be editing the spreadsheet at the same time as the nightly sync, potentially causing:
- Overwrites of in-progress user edits
- Stale data being merged around concurrent edits
- Partially-applied sync state leaving tables inconsistent

No implementation was done. This issue captures the requirement and design options.

---

## Requirements

1. The nightly sync job must not overwrite or corrupt live user edits.
2. Simultaneous user edits during the sync window must either be blocked gracefully or deferred, never silently corrupted.
3. The solution must not degrade real-time editing outside the sync window.
4. All existing 127 tests must continue to pass after implementation.

---

## Design Options to Evaluate

### Option A — `LockService.getScriptLock()` (Recommended starting point)

Google Apps Script's built-in `LockService` is the canonical concurrency primitive.

```js
var lock = LockService.getScriptLock();
var acquired = lock.tryLock(30000); // wait up to 30 s
if (!acquired) {
  Logger.log('[nightlySyncJob] Could not acquire script lock; aborting.');
  return { success: false, reason: 'lock_timeout' };
}
try {
  // ... all 6 sync steps ...
} finally {
  lock.releaseLock();
}
```

**Pros:** Native API, no extra storage, guaranteed exclusive access across all triggers and bound scripts.  
**Cons:** Blocks concurrent `installedTaskEditHandler` (onEdit installable) calls for the lock duration. If the nightly job takes > 6 minutes (Apps Script execution limit for time-driven triggers), the lock may not be released cleanly — `releaseLock()` in `finally` mitigates this.

**Verdict:** Try this first. Wrap the entire body of `nightlySyncJob()` and the write phases of `installedTaskEditHandler`.

---

### Option B — Document Properties Mutex (lightweight, weaker)

```js
var props = PropertiesService.getDocumentProperties();
if (props.getProperty('SYNC_LOCK') === '1') {
  Logger.log('[nightlySyncJob] Sync already in progress; aborting.');
  return;
}
props.setProperty('SYNC_LOCK', '1');
try { /* ... */ } finally { props.deleteProperty('SYNC_LOCK'); }
```

**Pros:** Simple, inspectable (an admin can clear the property if a crash leaves it set).  
**Cons:** Not atomic — two triggers can both read `null` before either writes `'1'`. Race condition window is small but real. Only use as a belt-and-suspenders layer on top of `LockService`, not as a standalone solution.

---

### Option C — Append-only / non-destructive write discipline (complementary)

Ensure every function called by `nightlySyncJob` never overwrites a user-entered cell directly. Rules:
- `compilePresence()`: only fills cells that are currently blank; never overwrites an existing AM/PM value unless it is the computed "baseline" placeholder.
- `refreshWeeklyMatrix()` / `refreshActiveResponsibles()`: read-only derivations from source tables — already safe.
- `sync()` (`syncMembersAndTenures`): check whether a row already exists before writing; skip existing rows.
- `syncTasks()` / `syncCalendarEvents()`: match by external ID (`task_id`, `calendar_event_id`) before updating; only patch delta fields.

**Verdict:** Implement this as a correctness discipline regardless of locking. The lock prevents concurrent writes; this discipline makes the writes idempotent.

---

### Option D — Split read-phase / write-phase (advanced)

1. **Read phase** (no lock needed): gather all data from all external APIs and all sheet ranges into memory.
2. **Diff phase**: compare against current sheet state in memory; compute only the minimal change set.
3. **Write phase** (locked): acquire lock, apply only the computed diff. Lock held for seconds, not minutes.

**Pros:** Minimizes lock duration dramatically.  
**Cons:** Complex to implement correctly across 6 heterogeneous sync engines. Reserve for a follow-up if Option A proves too slow.

---

## Recommended Implementation Plan

1. **Research `LockService` limits** — verify 6-minute trigger execution cap vs. lock TTL behavior.
2. **Wrap `nightlySyncJob()` body** in `LockService.getScriptLock().tryLock(30000)` / `releaseLock()`.
3. **Audit each step function** for direct cell overwrites on user-editable columns; add "skip if non-blank" guards (Option C).
4. **Add `SYNC_LOCK` Document Property** as belt-and-suspenders visibility layer (Option B).
5. **Update `test/mockGas.js`** — add `LockService` mock (`tryLock` returning `true`, `releaseLock` no-op) and a scenario where `tryLock` returns `false` to test early-exit behavior.
6. **Write tests** for the early-exit path (`nightlySyncJob` returns `{ success: false, reason: 'lock_timeout' }` when lock is not acquired).
7. **Run** `node --test test/*.js` — must remain 127 pass, 0 fail (plus any new tests added).

---

## Files to Read Before Coding

| File | Why |
|------|-----|
| [`src/Code.js` L532–603](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/src/Code.js#L532-L603) | Full `nightlySyncJob()` body — wrap this |
| [`src/PresenceEngine.js`](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/src/PresenceEngine.js) | `compilePresence()` — highest-risk write function |
| [`src/SyncEngine.js`](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/src/SyncEngine.js) | `syncMembersAndTenures()` — bidirectional member/tenure sync |
| [`src/TaskSyncService.js`](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/src/TaskSyncService.js) | `syncTasks()`, `installedTaskEditHandler()` |
| [`src/CalendarSyncService.js`](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/src/CalendarSyncService.js) | `syncCalendarEvents()` |
| [`test/mockGas.js`](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/test/mockGas.js) | Add `LockService` mock here |

---

## Test Command

```bash
node --test test/*.js
```

Expected baseline (before any changes): **127 pass, 0 fail**.

## Comments

_2026-10-11 — Issue created from user request. User said: "I am worried that people would interact with the sheet at the same time as the synchronization and I am worried that it would compromise data integrity. For now, don't implement anything." No code was changed. This issue is the handover artefact._

_2026-10-11 — Implemented, see `docs/adr/0032-per-spreadsheet-sync-lock.md`. It deviates from option A by using the **document** lock instead of the script lock: a library-created script lock is shared by every Consumer Sheet. Both `onEdit` and `installedTaskEditHandler` take the lock too (the simple `onEdit` also writes task/event IDs, presence log rows and directory sync). Option B was dropped and option D deferred. Removed the duplicate `installedTaskEditHandler` from `TaskSyncService.js`. Tests: `test/syncLock.test.js` (8 new); full suite 135 pass, 0 fail._
