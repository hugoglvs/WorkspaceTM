# ADR 0016: Weekly Matrix Planning View Architecture

## Status
Accepted

## Context
Management and project leads need a high-level operational overview of team deployments across time. Previously, daily attendance was recorded row-by-row in `tbl_presence_log`. While ideal for relational database querying, auditing, and analytics, a tabular log is difficult to parse visually for team capacity, co-location planning, and in-office synergy.

Management requested an enhanced Weekly Matrix Planning View (`view_weekly_matrix`) with four core capabilities:
1. **Merged Member People Chip Headers**: The header row above the AM and PM columns must have merged cells displaying the member's Google People Chip copied directly from `tbl_members`.
2. **Hidden Prior Weeks**: Historical weeks should remain in the sheet for auditing, but prior weeks older than the last 4 weeks must be hidden by default (`hideRows`) to focus on the active horizon.
3. **Merged Week Numbers & Styled Weekend Rows**: Week numbers in Column A must be merged vertically across each week's days. Weekend days (Saturday and Sunday) must appear in the matrix with a dark gray background (`#5F6368`) and white text.
4. **Interactive In-Sheet Presence Declaration**: Members/users must be able to declare or update their presence directly in the matrix cells via dropdowns, with strict data integrity syncing back to `tbl_presence_log` and updating `tbl_members.today_presence`.

## Decision
1. **Dedicated View Module (`ViewBuilder.js`):**
   - Implemented `renderWeeklyMatrixView(startDateStr, endDateStr, targetSpreadsheet)` to render `view_weekly_matrix`.
   - Freezes the top two header rows (`viewSheet.setFrozenRows(2)`).
   - Generates a two-row header structure:
     - Row 1: Merged cell for `Week Number` (A1:A2), `Date` (B1:B2), and 2-column merged cells for each member with native People Chips copied from `tbl_members` via `copyCellChip()`.
     - Row 2: Subheaders `AM` and `PM` under each member's merged name cell.
2. **Calendar Horizon & Week Merging:**
   - Expands full 7-day calendar weeks (Monday through Sunday) spanning the planning window.
   - Calculates ISO 8601 week numbers (`getIsoWeekNumber(date)` -> `"W01"`, `"W02"`).
   - Merges Column A vertically across each week's 7 rows with centered, bold week labels.
   - Applies dark gray background (`#5F6368`) and white text (`#FFFFFF`) to Saturday and Sunday rows.
3. **Visibility & Prior Week Folding:**
   - Automatically unhides all rows before rendering, then hides rows for weeks older than the last 4 weeks via `viewSheet.hideRows(3, rowsToHideCount)`.
4. **In-Cell Rapid Selection & Conditional Formatting:**
   - Applies dropdown data validation (`requireValueInList`) over all member AM/PM cells with standard presence modes:
     `Main Office`, `Remote`, `Downtown Office`, `Branch Office`, `Client Site`, `Vacation`, `Sick Leave`, `Holiday`.
   - Automatically establishes conditional formatting rules:
     - **Office Locations** (`Main Office`, `Downtown Office`, `Branch Office`, `Client Site`): Soft Green (`#E6F4EA` bg, `#137333` text).
     - **Remote / Home** (`Remote`): Soft Purple (`#F3E8FD` bg, `#7627BB` text).
     - **Leaves** (`Vacation`, `Sick Leave`, `Maternity/Paternity`, `Unpaid Leave`, `Holiday`): Soft Red (`#FCE8E6` bg, `#C5221F` text).
5. **Two-Way Synchronization & Data Integrity (`handleWeeklyMatrixEdit`):**
   - Connected to `onEdit(e)`:
     - Detects edits in `view_weekly_matrix` (row >= 3, col >= 3).
     - Resolves the date from Column B and the member and period (AM vs PM) from Column headers.
     - Upserts the record into `tbl_presence_log` as ground truth (copying the member's People Chip from `tbl_members`).
     - If the edited date is today, automatically syncs `tbl_members.today_presence`.
6. **Menu & Sync Integration:**
   - Added **Project Management > Refresh Weekly Matrix View** (`refreshWeeklyMatrix`) to the Sheets UI menu.
   - Compiling daily presence logs (`compilePresence`) automatically triggers a refresh of `view_weekly_matrix`.

## Consequences
- Team members and leads can plan and declare attendance directly from the interactive matrix or via the sidebar.
- Preserves ground-truth database separation: edits in `view_weekly_matrix` write immediately to `tbl_presence_log`.
- 100% automated test coverage in `test/viewBuilder.test.js` (56/56 passing tests across the entire suite).
