# 07: Google Tasks API Synchronization

**What to build:** An integration service that connects `tbl_tasks` to Google Tasks. When executed (via custom spreadsheet menu or scheduled trigger), it scans `tbl_tasks` for new tasks without a `google_task_id`, creates corresponding items in Google Tasks assigned to the respective Member, and writes back the generated `google_task_id`. In reverse, it inspects existing tasks in Google Tasks and syncs status updates (such as marking a task "Done" when completed in Google Tasks) back to the spreadsheet.

**Blocked by:** 02: Migration Engine & Protection Manager

**Status:** closed

- [x] Pushes unsynced tasks (where `google_task_id` is empty) to Google Tasks via the `Tasks.Tasks.insert()` API.
- [x] Writes back the returned task identifier into `tbl_tasks.google_task_id` in a batch update.
- [x] Inspects modified tasks in Google Tasks and updates the `status` column in `tbl_tasks` (e.g. `needsAction` -> `Todo`, `completed` -> `Done`).
- [x] Handles batching and error handling gracefully without exceeding Apps Script execution limits.
- [x] Document Properties stores the active Google Tasks List ID configuration per team.
