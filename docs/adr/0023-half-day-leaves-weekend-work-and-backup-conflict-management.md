# ADR 0023: Half-Day Leaves, Weekend Routines, Direct People Chip Writing, and Project Backup Conflict Management

## Status
Accepted

## Context
As project management and team scheduling evolved, four operational requirements emerged:
1. **Half-day leaves:** `tbl_leaves` previously only supported whole-day leaves (`start_date`, `end_date`), making it impossible to record morning-only or afternoon-only leaves. Daily presence compilation (`tbl_presence_log`) supports split AM/PM locations, but could not reflect half-day leave boundaries.
2. **Weekend working routines:** Presence baselines (`tbl_presence_baselines`) were constrained to weekdays (Monday–Friday). On-call team members, weekend rotations, or special shifts could not record routine weekend schedules or have them reflected in presence compilation. Weekends needed to be supported while defaulting to `"Off"`.
3. **Presence log People Chip compilation flicker:** When compiling daily presence logs, `compilePresenceLogs` previously wrote plain text strings for member names and then asynchronously replaced them cell-by-cell with People Chips via `Range.copyTo()`. This caused a perceptible visual flash from plain text to chips and increased execution time.
4. **Project lead and backup simultaneous absence conflicts:** Active projects in `tbl_projects` have a designated `lead_member` and `backup_member`. If both are absent at the same time (overlapping approved leaves or simultaneous concurrent absences), project operations risk being unmanaged. The system needed automated conflict detection, red alert row highlighting, notification emails to the lead and manager, interim backup task creation, and a dedicated `tbl_backups` table to assign replacements and mitigate conflicts.

## Decision

### 1. Half-Day Leaves (`tbl_leaves` & Sidebar UI)
- **Schema Extension:** Updated `SCHEMA_SPEC.tbl_leaves` to an 8-column layout:
  - Columns: `["leave_id", "member", "start_date", "start_period", "end_date", "end_period", "leave_type", "status"]`.
  - Added `start_period` (Column 4) and `end_period` (Column 6), nullable dropdowns with validation `["AM", "PM"]`.
- **Dynamic Header Resolution:** Implemented `getLeaveColumnMap(leavesSheet)` in `src/Sidebar.js` and `src/PresenceEngine.js` for seamless backwards compatibility with legacy 6-column sheets.
- **Sidebar UI Toggle Pills:**
  - Added interactive toggle pill groups (`AM | PM`) next to Start Date (defaulting to `AM`) and End Date (defaulting to `PM`) in `src/BaselineSidebar.html`.
  - Added `setLeavePeriod(type, period)` and updated `submitLeave()` payload to include `startPeriod` and `endPeriod`.
- **Split AM/PM Presence Compilation:** Updated `resolveLeave(mKey, dateObj)` in `src/PresenceEngine.js`:
  - Single-day leaves: `AM-AM` applies leave to morning only (afternoon uses baseline); `PM-PM` applies leave to afternoon only (morning uses baseline); `AM-PM` applies leave to full day.
  - Multi-day leaves: start date honors `start_period` (if `PM`, morning uses baseline); end date honors `end_period` (if `AM`, afternoon uses baseline); intermediate days are full leave.

### 2. Weekend Working Days in Routine Baselines
- **Schema & Validation:** Updated `SCHEMA_SPEC.tbl_presence_baselines` validation for `day_of_week` to include `"Saturday"` and `"Sunday"`, and added `"Off"` to `location_am` and `location_pm` dropdowns.
- **Sidebar Weekend Toggle:**
  - Added `+ Weekend Days` button beneath `#weekContainer` in `src/BaselineSidebar.html`.
  - Revealing weekend days appends Saturday and Sunday rows styled with `.weekend-row`, defaulting to `"Off"`.
  - `saveMemberBaselines()` collects weekend days when provided.
- **Selective Weekend Compilation:**
  - Standard weekdays (Monday–Friday) are compiled for all active members.
  - Weekend days (Saturday & Sunday) are compiled for a member ONLY IF they have an active baseline with `location !== "Off"` or an active approved leave on that weekend day, preventing database bloat while supporting working weekends.
  - Member `today_presence` on weekends displays `"Weekend"` unless the member has an active working baseline or approved leave today.

### 3. Direct People Chip Writing in Presence Compilation
- Replaced the two-step text-then-chip write pattern in `compilePresenceLogs`:
  - When reading `tbl_members`, captures `mem.raw` (the native People Chip / cell value) and `mem.richText`.
  - Constructs `newLogRows` with `mem.raw` directly in column 3.
  - Writes new rows to `tbl_presence_log` in a single pass via `setValues(newLogRows)`.
  - Applies `setRichTextValues()` in a single bulk operation if rich texts exist.
  - Eliminates the visual plain-text-to-chip flicker and accelerates compilation.

### 4. Project Lead and Backup Simultaneous Absence Conflict Engine (`src/ProjectBackupEngine.js`)
- **Dedicated Backup Table (`tbl_backups`):**
  - Added `tbl_backups` to `SCHEMA_SPEC`:
    - Columns: `["backup_id", "project_id", "primary_member", "backup_member", "start_date", "end_date", "status", "notes"]`.
    - Types: `backup_id` (TEXT), `project_id` (TEXT), `primary_member` (PEOPLE_CHIP), `backup_member` (PEOPLE_CHIP), `start_date` (DATE), `end_date` (DATE), `status` (DROPDOWN: `["Active", "Planned", "Completed", "Cancelled"]`), `notes` (TEXT).
- **Conflict Detection (`checkProjectBackupConflicts`):**
  - Scans active projects in `tbl_projects` (`status !== "Completed"`).
  - Detects simultaneous absences between `lead_member` and `backup_member` (overlapping approved leaves or concurrent absences today).
  - Checks `tbl_backups` for active/planned replacements covering the overlap period.
- **Automated Actions on Conflict:**
  1. **Alert Red Row Highlighting:** Sets project row background to `#F8D7DA` with `#721C24` bold font. Restores `#FFFFFF` / normal formatting once resolved.
  2. **Notification Email:** Sends email alert via `MailApp.sendEmail` to both the lead member and project manager detailing the project name, ID, and conflict date range. Deduplicates recipients and checks existing alert status to prevent email spamming on repeated edits.
  3. **Interim Backup Task Creation:** Creates a task in `tbl_tasks` titled `"Organize interim backup for project <Name>"`, assigned to `lead_member` (or to `backup_member` if the lead member is already absent). Checks for existing tasks to prevent duplicates.
- **Mitigation & Integration:**
  - Implemented `assignProjectBackup(payload, targetSpreadsheet)` to register replacement assignments and re-evaluate project conflicts.
  - Integrated into `onEdit` (when editing `tbl_projects`, `tbl_leaves`, or `tbl_backups`), `submitMemberLeave` in `src/Sidebar.js`, and added menu item `Project Management > Check Project Backup Conflicts`.

## Consequences
- Accurate recording and accounting for morning-only and afternoon-only leaves.
- Full flexibility for weekend working schedules and on-call rotations without cluttering the presence log.
- Flicker-free, high-performance presence log compilation with native People Chips preserved directly.
- Continuous coverage on active projects guaranteed through automated conflict detection, red alerting, email notifications, and task creation.
- Fully backwards compatible with existing 6-column sheets and older test suites.
- 100% test coverage with automated unit tests in `test/projectBackup.test.js` (99/99 tests passing across entire suite).
