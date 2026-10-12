# 20: IDs for Leaves Typed Directly in tbl_leaves

**Status:** closed (won't do)

**Blocked by:** _(none)_

**What to build:** A leave row typed straight into `tbl_leaves` (seen live: 2026-10-01 → 10-05 Vacation) gets no `leave_id`. `onEdit` should assign one, as it already does for tasks and events.

_2026-10-12 — Closed: the Leaves tab is hidden (ADR 0036) and leaves are declared through the Declare Presence & Leaves sidebar, which assigns IDs. Nothing reads `leave_id` for rows typed by hand._
