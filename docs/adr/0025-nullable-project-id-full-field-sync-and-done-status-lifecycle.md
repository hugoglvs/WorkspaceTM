# ADR 0025: Nullable Project ID, Full Field Synchronization, and Done Status Lifecycle

## Status
Accepted

## Context
In previous iterations of the task management module:
1. `project_id` in `tbl_tasks` was defined as a required column. However, operational teams frequently create administrative, personal, or ad-hoc tasks that are not associated with any overarching project.
2. In Google Apps Script, simple triggers (`onEdit`) execute with restricted authority and are blocked from invoking OAuth services like Google Tasks API. Real-time synchronization requires an installable trigger.
3. In `syncGoogleTasks`, reconciliation for existing tasks previously only inspected task status and did not propagate changes to `title`, `description` (Google Task notes), or `due_date` to Google Tasks.
4. When a task was marked "Done", users lacked safeguards preventing accidental modification, which caused state confusion against completed Google Tasks.

## Decision

1. **Nullable `project_id` in `tbl_tasks`:**
   - Updated `SCHEMA_SPEC.tbl_tasks`:
     - `requiredColumns`: `["task_id", "title", "status"]`
     - `nullableColumns`: `["project_id", "assigned_member", "description", "due_date", "google_task_id"]`
     - `warningProtectedColumns`: `[]`
   - Tasks without a `project_id` now pass schema and row validation cleanly without triggering warning toasts.

2. **Full Field & Status Synchronization (`syncGoogleTasks` & Live Edits):**
   - In `syncGoogleTasks`, reconciliation now checks and patches all mutable fields for existing synced tasks:
     - `title`: Synchronized when modified in the sheet (or defaulting to `task_id`).
     - `status`: Synchronized bidirectionally. If marked "Done" in the sheet, Google Task is marked `completed`. If sheet status is active ("Todo", "In Progress", "Blocked") and Google Task was not, status is patched to `needsAction`.
     - `due_date`: Formatted RFC 3339 timestamps are patched to Google Tasks `due`.
     - `description`: Notes containing the description and metadata (`Task ID`, `Project`, `Assigned`) are updated in Google Tasks.
   - Provided `installTaskSyncTrigger(ss)` and `installedTaskEditHandler(e)` with custom menu item `Enable Live Task Sync (Install Trigger)` under `Project Management`, enabling authorized real-time sync on edit.

3. **"Done" Status Modification Guard & Follow-up Task Creation:**
   - In `handleTaskEdit(e)`, attempting to change the status of a task from "Done" triggers `handleDoneStatusChange`:
     - Displays an interactive warning dialog (`SpreadsheetApp.getUi().alert`) notifying the user that a "Done" task cannot be modified directly.
     - **If Cancelled:** Reverts the cell back to "Done".
     - **If Confirmed:**
       - The original row remains marked "Done" and is hidden from view (`sheet.hideRows(editedRow)`).
       - An incremented task ID is generated using `.<number>` suffixing via `getIncrementedTaskId` (e.g. `TSK-XXXX` -> `TSK-XXXX.1`, `TSK-XXXX.1` -> `TSK-XXXX.2`).
       - A new follow-up row is appended with the same metadata (`project_id`, `assigned_member` with rich text People Chip, `title`, `description`, `due_date`, and the new status).
       - A brand new Google Task is created in Google Tasks with its own new `google_task_id`.

## Consequences
- Standalone tasks can be initiated freely without a `project_id`.
- Complete field parity between Google Sheets and Google Tasks across status, titles, descriptions, and deadlines.
- Permanent audit integrity for completed tasks through automated archiving and versioned follow-up creation.
- 105/105 automated unit tests passing across the test suite (`test/*.js`).
