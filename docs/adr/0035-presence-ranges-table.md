# ADR 0035: One Presence Ranges Table for Baselines and Leaves

## Status
**Reverted on 2026-10-12** (see ADR 0036). It was live for about 30 minutes on "[NEW] ACN". Kept here so the reasoning stays on record.

## Context
1. `tbl_presence_baselines` and `tbl_leaves` hold the same kind of fact: a date range that decides where a Member is for each Period. Leaves override Baselines when the Presence Log is compiled.
2. The user wants fewer sheets. Members only use the Weekly Matrix View and the Presence & Leaves sidebar, never these tables.
3. A leave type (`Vacation`, `Sick Leave`…) is already a valid Work Location in `tbl_presence_log`.

## Decision
1. Replace both tables with **`tbl_presence_ranges`**: `range_id, member, range_type (Baseline | Leave), start_date, start_period, end_date, end_period, day_of_week, location_am, location_pm, status (Requested | Approved | Cancelled)`.
   - Baseline: one row per weekday (`day_of_week`), `status` Approved, no periods.
   - Leave: `day_of_week` = `All`, the leave type is written into `location_am` and `location_pm`, and periods mark half days.
2. `src/PresenceRanges.js` owns the table (column order, readers `readBaselinesByMember` / `readApprovedLeavesByMember`, `appendPresenceRanges`). No other module touches its columns. The sidebar form is unchanged.
3. **Migration** (`migrateLegacyPresenceTables`) runs in `setupOrUpgrade` and as step 0 of the nightly sync. It copies legacy rows (6- and 8-column leave layouts, People Chips included). Leaves without an ID get one, and a blank leave status becomes `Approved`, matching how it was compiled before. The old tables are renamed `archive_tbl_*` and hidden, never deleted. The migration is idempotent.
4. **Visibility**: `setupOrUpgrade` hides `tbl_presence_ranges`, `tbl_presence_log`, `fact_presence` and `calc_kpis` (tables flagged `hiddenByDefault` in `SCHEMA_SPEC`). For presence, Members see only the Weekly Matrix View. Hidden, not locked (ADR 0034): anyone can unhide.

## Consequences
- 2 tables become 1, and the presence tabs shrink to the Weekly Matrix View.
- `required` columns are shared by both range types: end dates and periods are nullable, and the sidebar enforces leave end dates.
- Rolling back = rename `archive_tbl_*` back and redeploy the previous library version (each push is tagged `deploy-*` in git).

## Why it was reverted
Hiding the two tables already met the goal (fewer visible tabs). Leaves and Baselines also have different lifecycles: Leaves go through an approval status and use half-day periods, while Baselines are per-weekday routines. Merged, half of every row's columns were empty and the required-field rules had to be loosened. The user chose two hidden tables instead.
