# 8. Enforcing Column Nullability and Required Field Validation

Date: 2026-10-10

## Status

Accepted

## Context

In team operational spreadsheets (such as `tbl_tenures`, `tbl_projects`, `tbl_tasks`), some columns represent mandatory attributes while others are optional:
- In `tbl_tenures`: `end_date` can naturally be null/empty for permanent or ongoing team tenures. Conversely, `member`, `role`, `department`, `start_date`, `contract_type`, and `fte_rate` are essential for tenure evaluation.
- In `tbl_members`: `accenture_email` and `phone_number` are nullable until the member completes their onboarding task.
- In `tbl_projects`: `target_date` and `backup_member` are optional.
- In `tbl_tasks`: `assigned_member`, `due_date`, and `google_task_id` are optional (e.g., backlog tasks).

Without clear schema rules and automated enforcement, users may omit critical fields (leading to broken synchronization or invalid state) or be confused about which fields are optional.

## Decision

1. **Explicit Schema Specification**:
   Declare `requiredColumns` and `nullableColumns` on every table in `SCHEMA_SPEC` ([`src/Schema.js`](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/src/Schema.js)). Every column belongs to either `requiredColumns` or `nullableColumns`.
2. **Multi-Layer Enforcement**:
   - **Table Column Placeholders**: Capitalize on Google Sheets native Table column typing (`PEOPLE_CHIP`, `DATE`, `DROPDOWN`, `DOUBLE`, `TEXT`) and the "Show placeholders" feature, presenting faint guidance chips/text (`@`, `YYYY-MM-DD`, dropdown icons) in empty cells.
   - **Prominent Visual Warning Highlighting**: Configure conditional formatting rules and audit highlighters that dynamically style missing required cells with a prominent warning color (`#F8D7DA` alert red background, bold `#721C24` text) whenever a row is populated.
   - **`onEdit` Alerting**: When a user modifies a row and leaves required columns empty, trigger a prompt toast identifying the missing required fields while confirming optional fields (e.g. `end_date`).
   - **Sync Gating**: In `syncMembersAndTenures`, validate each tenure row. If required fields are missing, skip corrupt records and log/report specific validation warnings. Valid tenures with a null `end_date` are correctly treated as ongoing/active.
   - **Validation Engine**: Expose `validateTable(sheetName, targetSpreadsheet)` and a menu command `Validate Required Fields` for on-demand integrity auditing and cell highlighting.

## Consequences

- Permanent and ongoing tenures with null `end_date` are explicitly supported and correctly calculated as `Active`.
- Incomplete rows missing critical fields cannot corrupt the Team Directory or analytics.
- Empty cells in required columns are visually unmistakable with prominent red highlighting and native table placeholder cues directly in the Google Sheets interface.
