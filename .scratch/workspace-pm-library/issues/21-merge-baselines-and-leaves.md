# 21: Merge Baselines and Leaves into Presence Ranges

**Status:** resolved

**Blocked by:** _(none)_

**What to build:** One `tbl_presence_ranges` table replacing `tbl_presence_baselines` and `tbl_leaves`, with the sidebar form unchanged. Live data is migrated, and presence data sheets are hidden so Members see only the Weekly Matrix View. See ADR 0035.

- [x] `src/PresenceRanges.js` module: readers, appender, migration, hiding.
- [x] Presence compiler, backup conflict engine, Active Responsibles view, sidebar (baselines + leaves + diagnostics), chip sync and `onEdit` read/write the new table.
- [x] Migration in `setupOrUpgrade` and nightly step 0; old tables archived and hidden.
- [x] `tbl_presence_ranges`, `tbl_presence_log`, `fact_presence`, `calc_kpis` hidden on upgrade.

## Comments

_2026-10-11 — 4 new tests in `test/presenceRanges.test.js`. Existing presence/sidebar/backup/view tests now seed via `test/presenceFixtures.js`. Suite 146 pass._
