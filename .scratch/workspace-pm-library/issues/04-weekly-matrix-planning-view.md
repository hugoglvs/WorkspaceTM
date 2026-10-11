# 04: Weekly Matrix Planning View

**What to build:** A dynamic weekly planning dashboard tab (`view_weekly_matrix`) that renders team deployment in the specific table format requested by management. Each row represents a calendar day with its week number and date. Columns are grouped per active Member, with each member having two side-by-side cells per day: one for their morning (AM) location and one for their afternoon (PM) location. The view reads live data from `tbl_presence_log` and updates dynamically or upon manual refresh, while also allowing direct two-way presence declarations from the grid.

**Blocked by:** 03: Baseline & Leave Range Compiler to Daily Presence Log

**Status:** resolved

- [x] Sheet `view_weekly_matrix` displays a frozen 2-row header structure:
  - Row 1: Merged `Week Number` (A1:A2), `Date` (B1:B2), and 2-column merged cells per active member displaying their Google People Chip.
  - Row 2: `AM` and `PM` subheaders per member.
- [x] Rows represent consecutive days (Monday through Sunday) for the active planning period:
  - Week numbers in Column A are merged vertically across each week group.
  - Weekend days (Saturday and Sunday) appear with a dark gray background (`#5F6368`) and white text.
- [x] Historical weeks older than the last 4 weeks are hidden by default (`hideRows`) to keep the dashboard focused.
- [x] Cells pull AM and PM work locations directly from `tbl_presence_log` via fast programmatic batch rendering.
- [x] In-sheet rapid presence declaration:
  - Dropdown data validation across member cells for rapid location selection.
  - `onEdit` handler (`handleWeeklyMatrixEdit`) syncs changes back to `tbl_presence_log` as ground truth and updates `tbl_members.today_presence`.
- [x] Conditional formatting is applied automatically to color-code presence modes:
  - Office locations (e.g. Main Office, Downtown Office, Branch Office): Soft blue/green
  - Remote / Home: Soft purple/yellow
  - Leaves (Vacation, Sick Leave): Soft red/orange
- [x] Typography follows the 12px "Google Sans" standard.

## Answer

Implemented enhanced `view_weekly_matrix` planning view in `src/ViewBuilder.js`:
1. **Dynamic Grid Structure & Merged People Chips:**
   - Freezes the top 2 header rows.
   - Merges member header cells across AM and PM columns in Row 1 and copies native People Chips from `tbl_members`.
   - Subheaders `AM` and `PM` in Row 2.
2. **Calendar Horizon, Merged Weeks & Weekend Styling:**
   - Renders 7-day calendar weeks (Monday-Sunday).
   - Merges ISO Week Number in Column A vertically across each week's 7 days.
   - Styles weekend rows (Saturday & Sunday) with `#5F6368` dark gray background and white text.
3. **Visibility & Prior Week Folding:**
   - Automatically hides prior weeks older than the last 4 weeks (`viewSheet.hideRows()`).
4. **Interactive In-Sheet Declaration & Two-Way Sync:**
   - Applies dropdown validations (`requireValueInList`) for rapid selection.
   - `handleWeeklyMatrixEdit(e)` updates or inserts records into `tbl_presence_log` and updates `tbl_members.today_presence`.
5. **Conditional Formatting:**
   - Office locations (`#E6F4EA` bg, `#137333` text).
   - Remote / Home (`#F3E8FD` bg, `#7627BB` text).
   - Leaves (`#FCE8E6` bg, `#C5221F` text).
6. **Styling & Menus:**
   - Applied 12px "Google Sans" typography (`applyTypography`).
   - Added **Project Management > Refresh Weekly Matrix View** menu command, and auto-refresh on `compilePresence()`.
7. **Testing & Documentation:**
   - 100% test coverage in `test/viewBuilder.test.js` (56/56 passing tests across the entire suite).
   - Documented in `docs/adr/0016-weekly-matrix-planning-view-architecture.md`.
