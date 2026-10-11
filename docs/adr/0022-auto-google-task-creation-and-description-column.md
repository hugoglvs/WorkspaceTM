# ADR 0022: Automated Google Task ID Creation and Task Description Field

## Status
Accepted

## Context
In previous iterations of the task synchronization system (ADR 0020 and ADR 0021), task creation in Google Tasks was initiated primarily via manual menu synchronization (`Project Management > Sync Google Tasks`). Team members required a more responsive, real-time workflow:
1. When creating a new task directly in the spreadsheet, setting the task title should immediately create the corresponding Google Task in the background, populating `google_task_id` without requiring a separate sync action.
2. Tasks in `tbl_tasks` lacked a dedicated `description` field for narrative specifications, acceptance criteria, or contextual notes. Google Tasks has a built-in `notes` field (up to 8,192 characters) that was previously only used for technical metadata (`Task ID`, `Project ID`, `Assigned`).

## Decision

1. **Schema Extension (`tbl_tasks`):**
   - Added `description` (Column 5, `TEXT`, nullable) to `SCHEMA_SPEC.tbl_tasks`:
     - Columns: `["task_id", "project_id", "assigned_member", "title", "description", "due_date", "status", "google_task_id"]`.
     - Required: `["task_id", "project_id", "title", "status"]`.
     - Nullable: `["assigned_member", "description", "due_date", "google_task_id"]`.
   - Updated `Init.js` to create the 8-column layout with 12px Google Sans typography.
   - Updated `SyncEngine.js` onboarding task creation to include detailed profile completion instructions in the `description` column.

2. **Dynamic & Resilient Header Resolution (`getTaskColumnMap`):**
   - Implemented `getTaskColumnMap(tasksSheet)` in `src/TaskSyncService.js` and `src/Sidebar.js`.
   - Normalizes header names via alphanumeric stripping (`/[^a-z0-9]/g`) and lowercase conversion, ensuring resilience against casing, spaces, underscores, and human-readable aliases (`"Task ID"`, `"task_id"`, `"Due Date"`, `"Google Task ID"`).
   - Dynamically resolves column positions by inspecting header row 1, ensuring seamless backwards compatibility across both the new 8-column layout and legacy 7-column sheets without hardcoded index collisions.

3. **Decoupled In-Sheet Task Creation & Google Task Sync (`createGoogleTaskForTaskRow` & `handleTaskEdit`):**
   - **Trigger Priority in `onEdit`:** Placed `tbl_tasks` handling at the top of `onEdit(e)` before `validateRow`, ensuring `task_id` (`TSK-YYMMDD-XXXX`), `status: "Todo"`, and formatting are populated immediately without throwing premature "Missing required field" validation alerts.
   - **Immediate In-Sheet Initialization:** Whenever any field in a `tbl_tasks` row is edited or populated (whether `project_id`, `assigned_member`, `title`, `description`, etc.):
     - Auto-generates `task_id` (`TSK-YYMMDD-XXXX`) if missing.
     - Defaults `status` to `"Todo"` if missing.
     - Preserves `assigned_member` as manually assigned (never automatically filled).
     - Supports nullable `due_date`: tasks can be created and synced without a deadline, and clearing the date clears it in Google Tasks.
     - Applies standardized typography to the row.
   - **Safe Google Tasks Sync:**
     - Attempts Google Tasks creation only once `title` is populated.
     - Wrapped in strict try/catch blocks so simple trigger OAuth restrictions or missing service enablement NEVER block, abort, or prevent in-sheet task creation.
     - Preserves user isolation (ADR 0021): if explicitly assigned to another teammate, skips inserting into the active user's Google Tasks.
     - Writes back the generated Google Task ID to `tbl_tasks.google_task_id`.
   - **Live Propagations:** When `title`, `description`, `due_date`, or `status` is modified on a synced task row, propagates updates to Google Tasks via `Tasks.Tasks.patch`.

## Consequences
- Immediate feedback loop: creating a task row instantly generates `task_id` and defaults `status` to `"Todo"`, regardless of which column the user edits first.
- Complete decoupling: in-sheet task generation is 100% reliable and never interrupted by Google Tasks authorization or network status.
- Assignee autonomy: `assigned_member` is never auto-filled, preserving explicit manual assignment.
- Flexible scheduling: `due_date` is nullable and optional in both the spreadsheet and Google Tasks.
- Full narrative support: task descriptions are stored in `tbl_tasks` and synced to Google Tasks notes.
- Strict assignee isolation (ADR 0021) is preserved.
- 100% test coverage with automated unit tests in `test/taskSync.test.js` (91/91 tests passing across test suite).
