# ADR 0020: Google Tasks API Synchronization Architecture

## Status
Accepted

## Context
Project task coordination in `tbl_tasks` requires bidirectional synchronization with Google Workspace's personal and team task management service, Google Tasks. Team members manage their day-to-day work items via Google Tasks mobile applications, Gmail sidebars, and Google Calendar interfaces.

Without integration:
1. Team members had to manually duplicate tasks between the project spreadsheet and their personal Google Tasks lists.
2. Completed work items in Google Tasks were not reflected in the project management spreadsheet, risking out-of-date delivery dashboards.
3. Status changes made directly in the spreadsheet were not propagated to assignee devices.

Issue 07 defined the requirement for an integration service connecting `tbl_tasks` with Google Tasks, managing unsynced task creation, status synchronization, batching, and list configuration.

## Decision

1. **Modular Service (`src/TaskSyncService.js`):**
   - Implemented `src/TaskSyncService.js` following the architecture defined in `spec.md` (`TaskSyncService`).
   - Enabled Google Tasks Advanced Service (`version: v1`, `serviceId: tasks`, `userSymbol: Tasks`) in `src/appsscript.json`.

2. **Team Task List Configuration via Document Properties:**
   - Property Key: `"GOOGLE_TASKS_LIST_ID"`.
   - `getTasksListId(targetSpreadsheet)`: Reads the configured list ID from `PropertiesService.getDocumentProperties()`.
   - On initial run, discovers or creates a dedicated list titled `"Project Management Tasks"` via `Tasks.Tasklists.insert()`. If permissions or quota restrict list creation, cleanly falls back to `"@default"`.
   - `setTasksListId(listId)`: Allows programmatic or administrative reconfiguration.

3. **Outbound Push (Unsynced Tasks to Google Tasks):**
   - Scans `tbl_tasks` for populated rows where `google_task_id` (Column 7) is empty.
   - Formats task resource:
     - `title`: Task title from Column 4.
     - `notes`: Metadata string combining Task ID, Project ID, and Assigned Member (`Task ID: TSK-001 | Project: PRJ-001 | Assigned: Alice Smith`).
     - `due`: RFC 3339 UTC timestamp generated from `due_date` (`formatTaskDueRfc3339`).
     - `status`: Mapped to `"completed"` (if sheet status is "Done") or `"needsAction"` (for "Todo", "In Progress", "Blocked").
   - Calls `Tasks.Tasks.insert(taskResource, tasklistId)`.
   - Writes back the generated `task.id` into `tbl_tasks.google_task_id` in a batch write.

4. **Inbound Pull & Bi-Directional Reconciliation:**
   - Queries existing items from Google Tasks (`Tasks.Tasks.list`) with `showCompleted: true` and `showHidden: true`.
   - When a task's status is `"completed"` in Google Tasks and not `"Done"` in the spreadsheet, updates `tbl_tasks.status` to `"Done"`.
   - When a task is marked `"Done"` in the spreadsheet, sync updates the task in Google Tasks to `"completed"` via `Tasks.Tasks.patch`.
   - If a previously completed task is reopened (`needsAction`) in Google Tasks, updates `tbl_tasks.status` back to `"Todo"`.

5. **Error Handling & Quotas:**
   - Verifies availability of the Google Tasks Advanced Service (`typeof Tasks !== "undefined" && Tasks.Tasks`) and reports an actionable warning if disabled.
   - Individual task insert/patch operations are wrapped in resilient try-catch blocks to prevent a single corrupted task from halting the overall sync.

6. **UI Integration:**
   - Exposed custom menu entry: **Project Management > Sync Google Tasks** (`syncTasks`).
   - Hooked into the central `sync()` routine in `src/Code.js`.
   - Toast notifications inform users of pushed tasks count and updated statuses.

## Consequences
- Team members can manage project tasks directly from their Google Tasks mobile apps, Gmail sidebars, or the spreadsheet with bidirectional consistency.
- `tbl_tasks.google_task_id` remains protected from manual user editing while being reliably populated by the synchronization engine.
- 100% test coverage with automated unit tests in `test/taskSync.test.js` (80/80 tests passing across test suite).
