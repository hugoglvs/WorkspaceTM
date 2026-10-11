# 15: Declutter Tabs (Hide Data Tables by Default)

**Status:** open

**Blocked by:** _(none)_

**What to build:** `setupOrUpgrade` hides every `tbl_*`, `calc_*` and `fact_*` sheet, orders the visible tabs (Weekly Matrix View, Active Responsibles, `tbl_tasks`), and removes an empty default `Sheet1`. Hidden only: anyone can unhide (ADR 0034).

- [ ] Data sheets hidden after `setupOrUpgrade`; `tbl_tasks` and views visible, in that order.
- [ ] An empty `Sheet1` is deleted; a non-empty one is left alone.
- [ ] Running it again is idempotent and never re-hides a sheet the user unhid on purpose (decide: only hide on first creation, or on every upgrade?).
