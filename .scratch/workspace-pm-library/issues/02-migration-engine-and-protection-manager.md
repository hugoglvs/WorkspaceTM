# 02: Migration Engine & Protection Manager

**What to build:** An upgrade and protection manager that inspects an existing spreadsheet's installed schema version and performs safe, non-destructive additive migrations when the Library version is incremented. In addition, it locks primary key columns and read-only calculated columns using Google Sheets `ProtectedRange` permissions so that team members can freely edit operational columns without corrupting database identifiers or formulas.

**Blocked by:** 01: Core Schema & Typography Initializer

**Status:** resolved

- [x] `setupOrUpgrade()` checks `PropertiesService.getDocumentProperties()`:
  - If version is `0`, invokes `freshInit()`.
  - If version is lower than library version, sequentially executes incremental migration steps.
  - If version matches, informs the user via `toast()` without making changes.
- [x] Migrations append missing columns to the right end of existing tables without overwriting or displacing populated row data.
- [x] Protected ranges are applied to primary keys across all tables:
  - `tbl_members`: `member_id`
  - `tbl_tenures`: `tenure_id`
  - `tbl_presence_baselines`: `baseline_id`
  - `tbl_leaves`: `leave_id`
  - `tbl_presence_log`: `log_id`, `timestamp`
  - `tbl_projects`: `project_id`
  - `tbl_tasks`: `task_id`, `google_task_id`
  - `tbl_events`: `event_id`, `google_calendar_event_id`
- [x] Operational non-key columns remain unlocked for team member data entry.
