# 22: Members Tab with Tenure Sidebar

**Status:** resolved

**Blocked by:** _(none)_

**What to build:** On the visible Members tab, new Members and Tenure changes (new contract, end date) are declared in a sidebar that writes `tbl_tenures` and syncs `tbl_members`. Starting a new row on Members opens that sidebar.

- [x] Prototype first: can an installable `onEdit` open a sidebar for *any* editor, or only for the trigger owner? (Docs are silent.)
- [x] If not, fallback: a toast pointing to a "➕ New member" button/menu, which opens the sidebar for everyone.
- [x] Grid edits remain possible and keep syncing (ADR 0034 flexibility).

## Comments

_2026-10-12 — Implemented (ADR 0037): `src/MembersService.js`, `src/MembersSidebar.html`, a menu item and forwarders in `consumer/Code.js`, and an onEdit hook in `src/Code.js`. 10 new tests; suite 153 pass. The sidebar opens automatically, with a toast fallback. Still to confirm live: whether automatic opening works for an editor who is not the script owner (test with a second account)._

_2026-10-12 — Live test: the sidebar does **not** open by itself when a new name is typed in Members, even for the owner (the simple onEdit trigger cannot open UI). The menu works. Parked by the user until the end; options then: installable onEdit, or a clickable button/checkbox in the tab._

_2026-10-12 — Verified live: typing a new name in Members opens the Members & Contracts sidebar pre-filled (the installable edit trigger can show UI; ADR 0041/0044). No protected-column dialog first._

_2026-10-12 — Superseded (ADR 0046): the sidebar is removed. People chips written by script did not work in the enterprise domain; contracts are now typed in the Contracts tab, opened by the menu *Declare Contract* or by typing a new name in Members._

_2026-10-13 — Reworked (ADR 0048): member first (name typed in Members), then a Contracts sidebar that lists, adds and edits contracts; no email/chip writing._
