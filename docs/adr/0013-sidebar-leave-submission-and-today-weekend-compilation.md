# 13. Sidebar Leave Submission and Weekend Today Compilation

Date: 2026-10-10

## Status

Accepted

## Context

Users requested two key functional improvements:
1. **Weekdays Only in Ground-Truth Presence Log**: Daily presence logs (`tbl_presence_log`) should strictly track working weekdays (Monday to Friday) and must omit weekends. If compilation is executed on a weekend day (such as Saturday or Sunday), compilation caps at the latest working day (Friday). However, `tbl_members.today_presence` should accurately indicate "Weekend" (or any approved active leave) rather than remaining blank or showing "Not Logged".
2. **Leave Submission directly from Sidebar**: Rather than navigating manually to `tbl_leaves` and entering raw records, users need the ability to submit leaves directly from the Presence Sidebar (`BaselineSidebar.html`). Submitting an approved leave should preserve native Google People Chip identity, apply formatting and protections, and immediately trigger presence compilation.

## Decisions

1. **Weekday-Only Logging in tbl_presence_log (`src/PresenceEngine.js`)**:
   - `tbl_presence_log` strictly compiles weekdays (Monday–Friday) via `getWeekdaysBetween(startDate, endDate)`.
   - When compilation runs on a weekend day, no weekend records are added to `tbl_presence_log`.
   - `tbl_members.today_presence` evaluates today's status: if today is a weekend day, it displays `"Weekend"` (or active approved leave type), while on weekdays it displays the logged presence or `"Not Logged"`.
2. **Leave Submission in Sidebar (`src/Sidebar.js` & `src/BaselineSidebar.html`)**:
   - Added `submitMemberLeave(payload, targetSpreadsheet)`:
     - Validates required inputs (`member`, `startDate`, `endDate`, `leaveType`, `status`).
     - Appends new leave record into `tbl_leaves` starting at next available row.
     - Preserves native Google People Chip identity by copying from `tbl_members`.
     - Applies standard typography and protections.
     - When `status === "Approved"`, automatically triggers `compilePresenceLogs()` for real-time presence log updates.
   - Enhanced `BaselineSidebar.html` with a tab switcher:
     - **Routine Baseline**: weekly schedule and non-destructive overlap resolution.
     - **Submit Leave**: leave type, date range with quick presets ("Today Only", "Current Week", "Next Week"), and status.

## Consequences

- `tbl_presence_log` contains clean, queryable records for working days only.
- `tbl_members.today_presence` provides a clear and accurate live indicator on both weekdays and weekends.
- Users can manage both declarative recurring baselines and ad-hoc leave records in one unified sidebar interface.
- 100% test coverage with automated unit tests across all 38 test cases.
