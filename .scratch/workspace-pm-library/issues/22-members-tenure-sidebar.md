# 22: Members Tab with Tenure Sidebar

**Status:** resolved

**Blocked by:** _(none)_

**What to build:** On the visible Members tab, new Members and Tenure changes (new contract, end date) are declared in a sidebar that writes `tbl_tenures` and syncs `tbl_members`. Starting a new row on Members opens that sidebar.

- [x] Prototype first: can an installable `onEdit` open a sidebar for *any* editor, or only for the trigger owner? (Docs are silent.)
- [x] If not, fallback: a toast pointing to a "➕ New member" button/menu, which opens the sidebar for everyone.
- [x] Grid edits remain possible and keep syncing (ADR 0034 flexibility).

## Comments

_2026-10-12 — Implemented (ADR 0037): `src/MembersService.js`, `src/MembersSidebar.html`, a menu item and forwarders in `consumer/Code.js`, and an onEdit hook in `src/Code.js`. 10 new tests; suite 153 pass. The sidebar opens automatically, with a toast fallback. Still to confirm live: whether automatic opening works for an editor who is not the script owner (test with a second account)._
