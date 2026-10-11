# 10. Protect Rows Upon Creation (Preventing Protection Warnings During New Row Registration)

Date: 2026-10-10

## Status

Accepted

## Context

In our schema model, key business attributes (such as `member` and `start_date` in `tbl_tenures`) are configured as `warningProtectedColumns`, prompting users with a native Google Sheets warning confirmation to prevent accidental modification of existing records.

However, previously `applyProtections()` applied `ProtectedRange` locks across entire sheet columns down to `maxRows` (e.g. rows 1 to 50/1000). Consequently, whenever a team member attempted to register a new tenure on an empty row (or even clicked into the empty cell to add a People Chip or start date), Google Sheets intercepted the action with a protected range warning dialog. Furthermore, the `onEdit` trigger fired a toast warning (`Warning: You edited warning-protected column...`) even though the user was simply populating a new row.

Users require that protections be applied **only once the row is created** (i.e. once the row is fully populated and valid). Empty rows designated for entering new records must remain completely open and unprotected.

## Decision

1. **Defer Protection to Created Rows**:
   - `getLastCreatedRow(sheet, config)` dynamically scans the sheet to identify the highest row index that is populated and valid (containing all required columns via `validateRow()`).
   - `warningProtectedColumns` (e.g. `member`, `start_date` in `tbl_tenures`, `lead_member` in `tbl_projects`, `project_id` in `tbl_tasks`, etc.) are applied **strictly across created rows**: `sheet.getRange(2, colIdx, numCreatedRows, 1)` where `numCreatedRows = lastCreatedRow - 1`.
   - If a sheet has no created rows (`lastCreatedRow <= 1`), no warning protections are placed on data rows.
   - All rows below `lastCreatedRow` (including the immediate next available row for new records) remain completely unencumbered by protections.

2. **Strict Locks for System Integrity**:
   - In `tbl_members`, where users are barred from manually inserting rows, read-only columns (`member`, `current_role`, `current_department`, `tenure_status`, `today_presence`) remain strictly locked across `maxRows` to authorized admins (`ss.getOwner()`, `Session.getEffectiveUser()`, and `ADMIN_EMAILS`), leaving `accenture_email` and `phone_number` fully unlocked for member profile updates.

3. **Active Range Check in `onEdit`**:
   - `isCellProtected(sheet, row, col)` queries the sheet's active `ProtectedRange` objects to verify whether the edited cell is genuinely within an active protection boundary before displaying any warning or admin toast.
   - Edits performed on uncreated/unprotected rows during record creation pass silently without generating false-positive warnings.

4. **Lifecycle Reapplication**:
   - When a row is completed and synchronized (via `syncMembersAndTenures()`), `applyProtections()` executes, recalculating `lastCreatedRow` to include the newly committed record and dynamically expanding the protected range over it while keeping the subsequent row open.

## Consequences

### Positive
- Users can register new tenures, projects, tasks, leaves, and events without encountering any warning prompts or dialogs.
- Once a record is completed, it is safely protected against accidental tampering.
- False-positive toast alerts during data entry are completely eliminated.
- Clean separation between immutable/read-only roster sheets (`tbl_members`) and operational intake sheets (`tbl_tenures`).

### Negative
- A row that is partially typed (incomplete) is not locked until its required fields are populated. However, this is desirable as it allows uninterrupted data entry.
