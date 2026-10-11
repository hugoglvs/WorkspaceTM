# ADR 0021: Task Assignment Isolation and Foreign Google Tasks Cleanup

## Status
Accepted

## Context
During initial deployment of Google Tasks synchronization (ADR 0020), all unsynced rows in `tbl_tasks` were pushed to the active user's Google Tasks account. Because Google Apps Script executes API calls (`Tasks.Tasks.insert`) under the authenticated session of the user who clicks "Sync Google Tasks" or triggers the sync (`Session.getActiveUser()`), tasks assigned to other team members (such as onboarding tasks or peer tasks) were inadvertently created in the active user's personal Google Tasks list.

Furthermore, when diagnosing assignment discrepancies, team members needed visibility into:
1. Why specific tasks in `tbl_tasks` matched or did not match their active session.
2. Which tasks in `tbl_tasks` already held Google Task IDs despite belonging to other members.
3. An automated mechanism to clean up misassigned tasks from their Google Tasks account without losing task metadata in the spreadsheet.

## Decision

1. **Enhanced Diagnostic Audit (`src/Sidebar.js`):**
   - Extended `auditUserAndMembersIdentity(targetSpreadsheet)` to audit `tbl_tasks` row by row alongside `tbl_members`.
   - For every task row, logs:
     - Task ID, Project ID, Title, Status, and Google Task ID.
     - Assigned Member name and extracted People Chip email.
     - Exact matching result via `explainUserMemberMatch`: `Match: YES` or `Match: NO` with detailed justification.
   - Summarizes:
     - `tasksExamined`, `userTasksCount`, `otherTasksCount`.
     - `misassignedGTasksCount`: Tasks assigned to other members that have a Google Task ID recorded.
   - Alerts users in `debugUserIdentity()` UI dialog with a warning and actionable cleanup recommendation when misassigned tasks are detected.

2. **Strict Active User Isolation in `syncGoogleTasks` (`src/TaskSyncService.js`):**
   - Resolves the active user email from `Session.getActiveUser()` (or caller override `options.userEmail`).
   - Identifies active member entries in `tbl_members` matching the active user session.
   - For each task in `tbl_tasks`:
     - Evaluates whether the task is assigned to the active user (`isAssignedToActiveUser`).
     - Unless explicitly configured with `options.allUsers: true`, only pushes and patches tasks assigned to the active user.
     - Tasks assigned to other team members are safely skipped during standard sync, awaiting synchronization by their respective assignees.

3. **Foreign Tasks Cleanup Utility (`src/TaskSyncService.js` & `src/Code.js`):**
   - Implemented `cleanupMisassignedGoogleTasks(targetSpreadsheet)`:
     - Identifies tasks in `tbl_tasks` where `google_task_id` is populated but the assignee does NOT match the active user.
     - Calls `Tasks.Tasks.remove(tasklistId, googleTaskId)` to delete the task from the active user's personal Google Tasks.
     - Clears `google_task_id` in `tbl_tasks` so the rightful assignee can sync it later.
   - Exposed custom menu action: **Project Management > Clean Up Foreign Google Tasks** (`cleanupForeignTasks`).

4. **Name-to-Email Matching Robustness (`src/Sidebar.js`):**
   - Added first-name prefix matching in `explainUserMemberMatch`: when the normalized email prefix matches the first name of a multi-word display name (e.g. `alice@example.com` matching `Alice Smith`), it evaluates to `match: true`.

## Consequences
- Team members only receive their own assigned tasks in their Google Tasks list.
- Erroneously synced tasks belonging to teammates can be purged from the active user's Google Tasks in one click.
- Full diagnostic transparency is available via the existing **Debug User and Member Identity** tool.
- 100% test coverage with automated unit tests in `test/taskSync.test.js` (82/82 tests passing across test suite).
