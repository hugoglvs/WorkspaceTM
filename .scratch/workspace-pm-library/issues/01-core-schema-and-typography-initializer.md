# 01: Core Schema & Typography Initializer

**What to build:** An initialization engine that automatically generates the 8 core operational tables in a blank spreadsheet. When executed, it creates each tab with frozen header rows, applies 12px "Google Sans" typography across all cells, converts data ranges into native Google Sheets Named Tables with alternating row colors, applies dropdown validations for status and location fields, and records the baseline schema version (`SCHEMA_VERSION = 1`) in Document Properties.

**Blocked by:** None (can start immediately)

**Status:** resolved

- [x] `freshInit()` creates all 8 sheets: `tbl_members`, `tbl_tenures`, `tbl_presence_baselines`, `tbl_leaves`, `tbl_presence_log`, `tbl_projects`, `tbl_tasks`, and `tbl_events`.
- [x] Header rows are populated with bold text, frozen at row 1, and styled in 12px "Google Sans".
- [x] Each table is formatted as a Google Sheets Named Table with branded alternating row banding.
- [x] Dropdown validations are applied:
  - `tbl_members.status`: Active, On Leave, Departed
  - `tbl_presence_baselines.day_of_week`: All, Monday, Tuesday, Wednesday, Thursday, Friday
  - `tbl_leaves.leave_type`: Vacation, Sick Leave, Maternity/Paternity, Unpaid Leave
  - `tbl_leaves.status`: Requested, Approved, Cancelled
  - `tbl_presence_log.location_am` & `location_pm`: Remote, Main Office, Downtown Office, Branch Office, Client Site, Vacation, Sick Leave, Holiday
  - `tbl_projects.status`: Planning, Active, On Hold, Completed
  - `tbl_tasks.status`: Todo, In Progress, Done, Blocked
- [x] `PropertiesService.getDocumentProperties()` stores `SCHEMA_VERSION = "1"`.

## Answer

Implemented in `src/Schema.js`, `src/Typography.js`, `src/Init.js`, and `src/Code.js`.
The engine initializes all 8 tables, applies 12px "Google Sans" font styling, formats alternating table row banding, installs column-level dropdown validations, registers Named Ranges, and sets `SCHEMA_VERSION = "1"`. Verified with automated unit test suite in `test/init.test.js`.
