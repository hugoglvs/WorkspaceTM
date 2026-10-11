# ADR 0039: Readable Tab, Table and Header Names through a Table Registry

## Status
Accepted (2026-10-12)

## Context
1. Tabs and native tables had technical names (`tbl_members`, `view_weekly_matrix`, `fact_presence`), and visible headers read `accenture_email`. The user wants readable, natural English names.
2. About 110 places in the code looked tabs up by literal name, and any manual tab rename silently broke the scripts.

## Decision
1. **Table registry** `src/Tables.js` is the single source of truth: a stable key per table (`members`, `tenures`, `baselines`, `leaves`, `presenceLog`, `projects`, `tasks`, `events`, `backups`, `weeklyMatrix`, `kpiData`), its readable tab name and its former technical name.
2. Code never hard-codes a tab name. `getTable(ss, key)` finds a tab by its **stored sheet ID** (Document Properties `SHEET_ID_<key>`, written on every upgrade), then by readable name, then by former name. `tableKeyOfSheet(sheet)` does the reverse. A team can therefore rename a tab without breaking anything.
3. **Names** (English, spaces allowed; Google's own API example is "Project Tracker"):

   | Key | Tab and native table |
   |---|---|
   | members | Members |
   | weeklyMatrix | Weekly Presence |
   | projects | Projects |
   | tasks | Tasks |
   | events | Events |
   | tenures | Contracts |
   | baselines | Routines |
   | leaves | Leaves |
   | presenceLog | Presence Log |
   | backups | Backups |
   | kpiData | KPI Data |

   Archives become `Archive – <name>`. Named ranges keep the technical name (`tbl_members_table`), because they cannot contain spaces.
4. **Headers.** `SCHEMA_SPEC` (now keyed by registry key) gives readable `headers` labels for the visible tables (Members, Projects, Tasks, Events), e.g. *Accenture email*, *Assigned to*, *In charge today*. Column keys stay internal; every header lookup accepts both the key and the label. Hidden tables and `KPI Data` keep `snake_case` headers (Looker-friendly).
5. **Migration on upgrade:** `renameLegacySheets` renames the tabs first. `freshInit` then writes the labels, and `createGoogleSheetsTable` renames the table it finds on each tab and widens it if needed. `registerTableIds` stores the sheet IDs. Data is untouched. Until the upgrade runs, former names still resolve.

## Consequences
- New code uses `getTable(ss, "<key>")` and `tableKeyOfSheet(sheet)`, never a literal tab name.
- User-facing messages name tabs as people see them ("the hidden Backups tab").
