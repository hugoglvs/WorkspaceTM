# 22: Members Tab with Tenure Sidebar

**Status:** open

**Blocked by:** _(none)_

**What to build:** On the visible Members tab, new Members and Tenure changes (new contract, end date) are declared in a sidebar that writes `tbl_tenures` and syncs `tbl_members`. Starting a new row on Members opens that sidebar.

- [ ] Prototype first: can an installable `onEdit` open a sidebar for *any* editor, or only for the trigger owner? (Docs are silent.)
- [ ] If not, fallback: a toast pointing to a "➕ New member" button/menu, which opens the sidebar for everyone.
- [ ] Grid edits remain possible and keep syncing (ADR 0034 flexibility).
