# ADR 0034: Flexible Grids plus Sidebars, Decluttered Tabs, and Looker Studio KPIs

## Status
Accepted

## Context
1. The reference Consumer Sheet shows 13 tabs (9 `tbl_*`, 3 views, a leftover `Sheet1`) to every Member. That is too much to take in.
2. Product principles stated by the user (2026-10-11):
   - **Maximum flexibility**: every grid stays directly editable, the Weekly Matrix View included.
   - **Full transparency**: there is no privacy model. Everyone may see and edit every profile, tenure and KPI.
   - Members work on desktop only. The Sheets mobile app (which shows no menus or sidebars) is not a target.
   - Tasks are edited both in `tbl_tasks` and in Google Tasks.
3. `calc_kpis` mixes Member rows with department/month "Total" rows, which double-counts in any BI tool. It also misreads FTE entered as a percentage (`100`) and counts weekend `Off` as leave.

## Decision
1. **Tables stay the storage layer.** All 9 `tbl_*` tables are kept: they are distinct domain concepts.
2. **Declutter, never lock.** Data tables are *hidden by default*, not restricted: anyone with edit access can unhide and edit them. Protections remain only for integrity (IDs, derived columns), not for privacy.
3. **Visible by default:** Weekly Matrix View (editable), Active Responsibles (projects and coverage), `tbl_tasks` (editable grid, synced both ways with Google Tasks).
4. **Sidebars are an additional entry path, not a replacement.** They give a guided, validated way to do common writes (My week: presence + leaves (exists); Projects & Backups; Events; Team: profiles + tenures, open to everyone). The grids remain available for anything a sidebar does not cover. The validation and protection code therefore stays.
5. **KPIs move to Looker Studio, built by the user.** The Library writes a tidy, hidden fact table `fact_presence` (one row per Member × date × Period) for Looker to aggregate. `calc_kpis` is retired once the Looker report replaces it.
6. **FTE accepts both units.** `fte_rate` values above 1 are read as percentages (`80` → `0.8`).

## Consequences
- Members see 3 tabs plus a menu of sidebars, and power users keep full grid access.
- Because grids stay editable, edits made directly in tables still need ID generation, validation and the Sync Lock (ADR 0032).
- KPIs can be filtered and charted in Looker without spreadsheet formatting code. Looker's Sheets cache (about 15 min) is acceptable for monthly KPIs.
