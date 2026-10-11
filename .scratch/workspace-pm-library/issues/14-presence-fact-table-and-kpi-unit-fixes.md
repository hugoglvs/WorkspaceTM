# 14: Presence Fact Table for Looker Studio and KPI Unit Fixes

**Status:** resolved

**Blocked by:** _(none)_

**What to build:** `refreshKpis` also writes a hidden `fact_presence` sheet: one row per Member × date × Period, which Looker Studio aggregates. It fixes two live KPI bugs: `fte_rate` entered as a percentage (`100`) gave 2200 capacity days per month, and weekend `Off` was counted as leave. See ADR 0034.

- [x] `fte_rate` above 1 is read as a percentage (`100` → 1.0, `80` → 0.8) in all KPI calculations.
- [x] `Off` is not a leave and adds no worked days. A weekday `Off` keeps its capacity (the FTE already covers a part-time day off); a weekend `Off` has none.
- [x] `fact_presence` columns: `date`, `month`, `iso_week`, `weekday`, `is_weekend`, `member`, `department`, `role`, `fte_rate`, `period`, `location`, `category` (`remote` / `office` / `leave` / `off` / `not_logged`), `worked_days`, `leave_days`, `capacity_days`.
- [x] Rows run from the first logged date to today. A weekday with no log is still a row (`not_logged`, with capacity), so capacity is complete. Weekends appear only when logged, and carry no capacity.
- [x] No total rows: every aggregate is left to Looker.
- [x] The sheet is hidden and rebuilt on each `refreshKpis` (so nightly too).

## Comments

_2026-10-11 — Implemented in `src/KpiEngine.js` (`normalizeFteRate`, `isOffLocation`, `buildPresenceFacts`, `renderPresenceFacts`; `refreshKpis` now also rebuilds `fact_presence`). 6 new tests in `test/kpiEngine.test.js`; suite 141 pass. Checked against the live "[NEW] ACN" data: Hugo's October capacity is 22 days (was 2200) and leave 0 (was 1). Rows for dates outside any Tenure are kept only when logged, with capacity 0._
