# ADR 0036: Five Visible Tabs, Reverted Presence Merge, and Looker-only KPIs

## Status
Accepted (2026-10-12)

## Context
1. The user wants Members to see five tabs: **Members, Weekly Matrix, Projects, Tasks, Events**. Edits stay possible in every visible table, either directly in the grid (kept in sync) or through a sidebar.
2. The merge of Baselines and Leaves (ADR 0035) was reverted. By then the reference sheet held newer data only in `tbl_presence_ranges` (a routine from 2026-10-09 to 10-16, an end date on the previous routine, a leave ID).
3. KPIs live in Looker Studio on `fact_presence` (ADR 0034). `calc_kpis` duplicated it. `refreshKpis` was declared in both `Code.js` and `KpiEngine.js`, and Apps Script keeps only one global of each name.

## Decision
1. **Layout** (`src/Layout.js`, `applyWorkspaceLayout`, run by `setupOrUpgrade`):
   - The five tabs are shown.
   - Every other Library-managed sheet (`tbl_`, `view_`, `fact_`, `calc_`, `archive_`) is hidden, not locked.
   - Tabs the team created themselves are never touched. An empty default `Sheet1` is deleted.
2. **Presence tables:** `tbl_presence_baselines` and `tbl_leaves` again. `restorePresenceTablesFromRanges` (run in `setupOrUpgrade` and as nightly step 0) moves `tbl_presence_ranges` rows back (People Chips kept, IDs never duplicated) and archives that table as hidden `archive_tbl_presence_ranges`. The `archive_tbl_*` copies made by ADR 0035 are stale snapshots from before the merge.
3. **KPIs:** `calc_kpis` and its code are deleted. `Code.js` holds the single `refreshKpis`, which rebuilds `fact_presence`.
4. **Weekly Matrix:** edits are still saved to the Presence Log. The confirmation toast also points to *Declare Presence & Leaves* for routines and leaves.
5. `freshInit` deletes the initial default sheet only when it is empty.

## Open (tracked as issues)
- Members: Tenure creation and end dates through a sidebar that opens when a new row is started (issue 22).
- Projects: read-only backup column fed by `tbl_backups`, a backup row created with each new project, removal of `view_active_responsibles`, and the permanent-vs-interim backup model (issue 23).
