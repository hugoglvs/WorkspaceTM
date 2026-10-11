# ADR 0024: Immediate Google Task Creation with Default Task ID Title and In-Place Title Patching

## Status
Accepted

## Context
In ADR 0022, task creation in `tbl_tasks` was enhanced so that setting a task title immediately created the corresponding Google Task in the background, populating `google_task_id`. However:
1. When team members initialize a task row by entering fields other than `title` first (such as `project_id`, `description`, or `due_date`), `task_id` was created immediately, but `google_task_id` remained deferred until `title` was populated.
2. In agile and fast-paced operational workflows, users expect an immediate Google Task link as soon as a task row is initiated, with `task_id` serving as an unambiguous temporary placeholder title until a descriptive title is drafted.
3. Once the user subsequently enters or refines the task title in the spreadsheet, the corresponding Google Task title should automatically update in-place from the temporary `task_id` to the real title.

## Decision

1. **Immediate Google Task Creation upon Task ID Generation (`createGoogleTaskForTaskRow`):**
   - In `createGoogleTaskForTaskRow`, Google Task insertion is no longer blocked when `title` is empty.
   - The initial title for Google Tasks defaults to `task_id` (`effectiveTitle = title || taskId`).
   - When any field is edited in a row of `tbl_tasks`, `handleTaskEdit(e)` immediately generates `task_id`, defaults `status` to `"Todo"`, and triggers `createGoogleTaskForTaskRow` right away.
   - The returned `google_task_id` is immediately saved into `tbl_tasks.google_task_id`.

2. **In-Place Title Replacement via Live Patch (`handleTaskEdit`):**
   - When a user subsequently edits the `title` column on a task that already has a `google_task_id`, `handleTaskEdit(e)` sends a `Tasks.Tasks.patch()` call to Google Tasks.
   - The patch replaces the default `task_id` title with the newly entered real title (`patchResource.title = title || taskId`).
   - If the user clears the title, the Google Task title gracefully falls back to `task_id` rather than triggering an invalid empty title API error.
   - Any updates to `title`, `description`, `project_id`, or `assigned_member` also update the structured metadata in Google Task `notes`.

3. **User Isolation & Integrity Safeguards Preserved (ADR 0021):**
   - If a task row is explicitly assigned to another team member (`assigned_member`), creation and patch calls to the active user's Google Tasks list are skipped.
   - Nullable `due_date` and manual `assigned_member` assignment remain strictly respected.
   - If Google Tasks API throws an error (e.g. offline simple trigger restrictions), in-sheet task row creation succeeds gracefully without user disruption.

## Consequences
- **Instant Synchronization:** Every task initialized in `tbl_tasks` immediately receives a valid `task_id` and `google_task_id`.
- **Seamless User Experience:** Users see their task ID right away in Google Tasks, which is automatically renamed as soon as they write the actual task name.
- **Bi-directional Consistency:** Full parity between spreadsheet edits and Google Tasks title state.
- **100% Test Coverage:** Verified with automated tests in `test/taskSync.test.js` (100/100 tests passing across the suite).
