# 20: IDs for Presence Ranges Typed Directly in the Grid

**Status:** open

**Blocked by:** _(none)_

**What to build:** A row typed straight into `tbl_presence_ranges` gets no `range_id`. `onEdit` should assign one, as it already does for tasks and events. (Rows that existed before issue 21 got IDs during the migration, including the live 2026-10-01 → 10-05 Vacation.)
