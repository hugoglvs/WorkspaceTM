# 10: Column Nullability & Required Field Enforcement

**What to build:** An enforcement and validation engine across all operational tables. Schemas explicitly declare `requiredColumns` and `nullableColumns`. Required columns (e.g. `member`, `role`, `department`, `start_date`, `contract_type`, `fte_rate` in `tbl_tenures`) are strictly enforced, while nullable columns (e.g. `end_date` in `tbl_tenures` for permanent contracts, `accenture_email`/`phone_number` in `tbl_members`, `target_date` in `tbl_projects`) are permitted to be empty. Enforcement includes real-time conditional formatting warning highlights, `onEdit` user toasts, sync pipeline integrity gates, and an on-demand validation report.

**Blocked by:** 01, 02

**Status:** resolved

- [x] Declare `requiredColumns` and `nullableColumns` on every table definition in `SCHEMA_SPEC` in `src/Schema.js`.
- [x] Implement `src/ValidationManager.js` with `validateRow(sheetName, rowValues, headers)`, `validateTable(sheetName, spreadsheet)`, and `validateAllTables(spreadsheet)`.
- [x] Add conditional formatting rules in `src/Init.js` highlighting missing required fields in active rows with warning styling.
- [x] Connect `onEdit(e)` in `src/Code.js` to toast actionable warnings when required fields are cleared or omitted in active rows.
- [x] Gate `syncMembersAndTenures` in `src/SyncEngine.js` to ensure incomplete tenure rows missing required fields are flagged and reported, while valid rows with null `end_date` are correctly processed as active/ongoing.
- [x] Add "Validate Required Fields" menu item to Google Sheets UI menu.
- [x] Write unit tests verifying required/nullable validation logic and sync gating.

## Answer

Implemented in `src/Schema.js`, `src/ValidationManager.js`, `src/Init.js`, `src/SyncEngine.js`, and `src/Code.js`.
ADR-0008 recorded in `docs/adr/0008-enforcing-column-nullability-and-required-fields.md`.
All tables define explicit `requiredColumns` and `nullableColumns`. In `tbl_tenures`, `end_date` is nullable (allowing permanent and open-ended contracts) while `member`, `role`, `department`, `start_date`, `contract_type`, and `fte_rate` are enforced.
Multi-tier enforcement is installed:
1. Real-time Google Sheets conditional formatting rules highlight missing required cells in active rows.
2. `onEdit(e)` triggers an immediate toast warning when an active row has missing required columns.
3. `syncMembersAndTenures()` checks integrity and captures validation warnings.
4. Custom menu item "Validate Required Fields" runs a spreadsheet-wide health audit.
Verified with 19 passing tests in `test/validation.test.js`, `test/sync.test.js`, `test/init.test.js`, and `test/sidebar.test.js`.
