# 24: Readable Tab, Table and Header Names

**Status:** resolved

**Blocked by:** _(none)_

**What to build:** Natural English names for tabs, native tables and visible headers, with a single source of truth and a migration on upgrade (ADR 0039).

- [x] `src/Tables.js` registry; `getTable` / `tableKeyOfSheet` used everywhere (rename-proof via stored sheet IDs).
- [x] `SCHEMA_SPEC` keyed by registry key, with readable header labels for Members, Projects, Tasks, Events.
- [x] Upgrade renames tabs, archives and native tables, and writes the labels; former names keep working until then.
- [x] User-facing messages use readable names.

## Comments

_2026-10-12 — Fixed on the way: the sync engine found Tasks columns by exact technical header and would have written due dates into Description once headers became readable. 3 + 1 new tests; suite 165 pass._
