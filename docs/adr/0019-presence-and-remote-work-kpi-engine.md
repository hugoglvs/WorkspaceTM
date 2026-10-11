# ADR 0019: Presence & Remote Work KPI Engine Architecture (calc_kpis)

## Status
Accepted

## Context
Tracking team presence across modern hybrid work models requires accurate, multi-dimensional analytics. Team members split their working time across physical corporate offices, branch offices, client sites, and home/remote environments. Moreover, intra-day mobility is common: a member may work remotely in the morning and commute to a physical office in the afternoon, or work a half-day and take a half-day vacation.

Evaluating team capacity also requires cross-referencing contractual commitments (FTE rates from `tbl_tenures`) against working weekdays in each month, rather than assuming standard 1.0 FTE allocations for all members.

Issue 06 defined the requirements for an automated calculation and analytics sheet (`calc_kpis`) to evaluate attendance metrics across members, departments, and months with 0.5-day weighting per period.

## Decision

1. **Dedicated Module (`src/KpiEngine.js`):**
   - Created `src/KpiEngine.js` following the modular library architecture specified in `spec.md`.
   - Contains core functions:
     - `computePresenceKpis(targetSpreadsheet)`: Reads data models from `tbl_presence_log`, `tbl_members`, and `tbl_tenures` and performs multi-dimensional aggregation.
     - `renderKpiDashboard(targetSpreadsheet)`: Generates and formats the `calc_kpis` sheet tab.
     - `refreshKpis(targetSpreadsheet)`: Public refresh entry point with user toast notification.

2. **Intra-Day Mobility & Period Weighting (0.5 Working Days):**
   - Each recorded day consists of an AM period and a PM period.
   - Each period is evaluated independently with a 0.5-day weight:
     - `Worked Days = Total worked periods * 0.5` (excluding non-working leaves and holidays).
     - `Leave Days = Total leave periods * 0.5` (Vacation, Sick Leave, Holiday, etc.).
     - `Remote Days = Count of Remote periods * 0.5`.
     - `In-Office Days = Count of Physical Office periods * 0.5`.
     - `Remote Rate = Remote Days / Worked Days`.
     - `In-Office Rate = In-Office Days / Worked Days`.
   - Specific physical office breakdowns:
     - Main Office Days (`Main Office periods * 0.5`)
     - Downtown Office Days (`Downtown Office periods * 0.5`)
     - Branch Office Days (`Branch Office periods * 0.5`)
     - Client Site Days (`Client Site periods * 0.5`)

3. **Contract Capacity and FTE Utilization (`tbl_tenures`):**
   - Counts working weekdays (Monday-Friday) in each target month (`countMonthWeekdays(monthKey)`).
   - Resolves each member's active contract in that month from `tbl_tenures` to extract `department` and `fte_rate` (defaulting to 1.0 FTE if unassigned).
   - `Contract Capacity (Days) = Month Weekdays * fteRate`.
   - `Capacity Utilization = Worked Days / Contract Capacity (Days)`.

4. **Multi-Level Aggregations & Rollups:**
   - Detailed rows grouped and sorted by:
     - Primary: Month (descending)
     - Secondary: Department (alphabetical)
     - Tertiary: Member name (alphabetical)
   - Clean subtotal rows:
     - **Department Subtotals**: `Total [Department]` summarizing department FTE, capacity, worked days, remote days, and average rates with `#F1F3F4` bold styling.
     - **Month Totals**: `Total Month ([YYYY-MM])` summarizing entire organization attendance with `#E2E7EC` bold styling.

5. **Google Sheets UI & Typography Standards:**
   - 12px "Google Sans" applied across all cells via `applyTypography`.
   - Native number formatting:
     - Percentage rates: `0.0%`
     - Days and capacities: `0.0`
     - FTE rates: `0.00`
   - Numeric columns right-aligned; text left-aligned; Month centered.
   - Frozen header row with `#E8EAED` background.
   - Native People Chips copied to the `Member` column via `copyCellChip`.

6. **Lifecycle Integration:**
   - Exposed custom menu item under `Project Management > Refresh Presence & Remote KPIs (calc_kpis)`.
   - Automatically refreshed on `compilePresence()`, `compilePresenceMigration()`, and `sync()`.

## Consequences
- Operations and department leads gain clear, automated visibility into remote work distribution, office desk occupancy, and capacity utilization.
- Intra-day mobility is fully respected without distorting full-time attendance counts.
- 100% test coverage with automated unit tests in `test/kpiEngine.test.js` (72/72 passing tests across the entire test suite).
