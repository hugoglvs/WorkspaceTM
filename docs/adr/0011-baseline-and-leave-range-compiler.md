# 11. Baseline and Leave Range Compilation Engine to Daily Presence Log

Date: 2026-10-10

## Status

Accepted

## Context

Managing daily presence directly on a row-by-row basis in Google Sheets creates excessive friction for team members. Instead, members declare:
1. Default weekly recurring routines or date-range baselines in `tbl_presence_baselines` (via the Presence Baseline Sidebar).
2. Approved vacation, sick leave, and parental leaves spanning date ranges in `tbl_leaves`.

We needed an automated compilation engine that bridges these declarative range rules into concrete, queryable daily records in `tbl_presence_log` for a forward planning horizon, while preserving any ad-hoc adjustments made directly by team members.

## Decision

1. **Ground Truth Horizon Bounded by Today**:
   - `tbl_presence_log` acts as the definitive ground-truth database of actual recorded presences.
   - `compilePresenceLogs(startDate, endDate)` expands working days (Monday to Friday) strictly up to `today` (the current date). Future dates are never written into `tbl_presence_log`.
   - Forward-looking schedule projections (e.g. next week or next month) are calculated on-the-fly when rendering planning views (`view_weekly_matrix`), preserving `tbl_presence_log` as an immutable audit trail of past and present reality.

2. **Precedence Hierarchy**:
   - **Manual Ad-hoc Edits (Highest Priority)**: If an entry for `(date, member)` already exists in `tbl_presence_log`, its existing `location_am` and `location_pm` values are preserved intact. The compiler is completely idempotent.
   - **Approved Leaves**: If an approved leave record exists in `tbl_leaves` covering the date, both `location_am` and `location_pm` are set to the `leave_type` (e.g. `Vacation`, `Sick Leave`). Unapproved requests (`Requested`, `Cancelled`) do not override presence.
   - **Baselines**: If an active baseline rule in `tbl_presence_baselines` matches the member, weekday, and date boundaries, its `location_am` and `location_pm` are applied.
   - **Fallback**: Default to `"Main Office"` for AM and PM.

3. **Natural Identity & Native People Chips**:
   - Newly inserted records copy native People Chips from `tbl_members` via `copyCellChip()` to maintain typed chips in `tbl_presence_log`.

4. **Real-time Team Directory Sync**:
   - The compiler evaluates each member's availability for the current day and updates `tbl_members.today_presence` in real-time.

5. **Lifecycle Protections**:
   - `applyProtections()` is invoked at the conclusion of compilation, protecting newly created log records while leaving unallocated sheet rows open for manual entries.

## Consequences

### Positive
- Zero repetitive data entry: members' daily schedules are pre-populated automatically from their baseline preferences and approved leaves.
- Preserves member autonomy: ad-hoc in-flight schedule changes are never overwritten by recurring compilation jobs.
- Direct operational visibility: `today_presence` in `tbl_members` remains live and up to date.
- All 32 automated tests pass cleanly.
