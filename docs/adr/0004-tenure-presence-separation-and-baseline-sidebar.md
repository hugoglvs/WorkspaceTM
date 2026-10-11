# 0004 Tenure-Presence Separation and Presence Baseline Sidebar

Context: Member roles, departments, and FTE allocations evolve over time across contracts and promotions, while daily operational availability changes on a day-to-day basis. Conflating employment status with daily presence creates ambiguous metrics. Additionally, typing baseline schedules into a spreadsheet table row-by-row creates high friction for team members.

Decision:
1. `role` and `department` are relocated to `tbl_tenures` as historical contract properties. `tbl_members` displays them as formula-driven read-only fields (`current_role`, `current_department`) reflecting the active tenure.
2. Status is explicitly divided into two distinct columns on `tbl_members`:
   - `tenure_status`: Contractual state (`Active`, `Upcoming`, `Departed`) evaluated from `tbl_tenures`.
   - `today_presence`: Operational availability for the current day (`Remote`, `Main Office`, `Vacation`, `Sick Leave`, etc.) evaluated from `tbl_presence_log`, `tbl_leaves`, and `tbl_events`.
3. We introduce an interactive Google Sheets Sidebar (`Presence Baseline Sidebar`) served via `HtmlService.createHtmlOutputFromFile`, allowing Members to select their recurring weekday presence pattern (AM/PM) via visual dropdowns, which commits directly to `tbl_presence_baselines`.

Considered Options:
- Static role and single status column on `tbl_members`: Rejected because role changes override historical records and status becomes ambiguous.
- Google Forms for baselines: Rejected because a built-in Apps Script sidebar keeps the user inside the spreadsheet context and directly interacts with the active spreadsheet session.

Consequences:
- `tbl_members` becomes a clean, read-only directory and live availability dashboard.
- Declaring baselines is frictionless via the integrated sidebar UI.
