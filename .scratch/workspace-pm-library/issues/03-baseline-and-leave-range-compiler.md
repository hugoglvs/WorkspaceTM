# 03: Baseline & Leave Range Compiler to Daily Presence Log

**What to build:** An automated presence compiler that bridges declarative range tables (`tbl_presence_baselines` and `tbl_leaves`) into concrete, queryable daily records in `tbl_presence_log`. Instead of forcing team members to enter every single day row-by-row, this engine generates daily entries for a specified forward-looking planning horizon (e.g., current month or quarter) or for complete historical data migrations from undefined start dates. It applies baseline location defaults and overrides them with approved leaves, while preserving any ad-hoc manual adjustments already entered by the member.

**Blocked by:** 02: Migration Engine & Protection Manager

**Status:** resolved

- [x] A compilation function `compilePresenceLogs(startDate, endDate)` expands working days (Mon–Fri) into `tbl_presence_log`.
- [x] For each date and member, it checks `tbl_presence_baselines` for active rules matching the day of the week or date range.
- [x] It checks `tbl_leaves` for approved absence records, overriding baseline locations with the leave type (e.g., Vacation or Sick Leave).
- [x] Existing daily rows in `tbl_presence_log` that contain manual member edits are preserved (idempotent compilation).
- [x] New records receive an auto-generated unique `log_id` and an immutable generation timestamp.
- [x] Menu item added for compiling presence logs from undefined/earliest start date for data migration:
  - `findEarliestHistoricalDate(targetSpreadsheet)` scans `tbl_tenures`, `tbl_presence_baselines`, `tbl_leaves`, and `tbl_presence_log`.
  - `compilePresenceMigration(targetSpreadsheet)` prompts with detected date or accepts custom date, and compiles historical logs up to today.

## Answer

Implemented in `src/PresenceEngine.js` and exposed via `src/Code.js`:
- `compilePresenceLogs(startDate, endDate, targetSpreadsheet)`: Expands weekdays (Monday to Friday) across active members into `tbl_presence_log`.
- Evaluates baseline defaults (`tbl_presence_baselines`), matching weekday and active date boundaries.
- Overrides defaults with approved absence records from `tbl_leaves` (`Vacation`, `Sick Leave`, etc.).
- Preserves existing rows with ad-hoc manual adjustments for full idempotence.
- Updates real-time `today_presence` in `tbl_members`.
- Reapplies column protections so newly created records are protected while keeping empty rows open.
- Added "Compile Daily Presence Logs" menu item to `onOpen` in `src/Code.js`.
- Added **Compile Presence Logs (Migration / Undefined Start Date)** menu item (`compilePresenceMigration`) with auto-discovery of earliest dates across all tables.
- Automatically refreshes `view_weekly_matrix` upon migration completion.
- Verified with comprehensive test suite in `test/presence.test.js` (62/62 tests passing).
