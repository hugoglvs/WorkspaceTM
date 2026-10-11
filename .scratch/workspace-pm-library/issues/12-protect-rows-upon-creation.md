# Issue 12: Protect Rows Upon Creation (Eliminate New Tenure Registration Warnings)

Status: resolved
Type: task

## Description

When registering a new tenure in `tbl_tenures`, Google Sheets repeatedly warned the user that they were modifying a protected column, and `onEdit` displayed a warning toast. This occurred because `applyProtections()` applied column-wide protection locks across the entire sheet down to `maxRows` (including all empty rows).

Protections must only be applied once the row is created (populated and valid). Empty rows designated for new entries must remain completely open so users can register new records smoothly without warnings.

## Acceptance Criteria

1. Empty rows below existing data in `tbl_tenures` (and operational intake tables) must not have any ProtectedRange locks.
2. Registering a new tenure (typing People Chip, start date, etc.) must not trigger Google Sheets' native warning dialog or `onEdit` warning toasts.
3. Once a row is created and valid, protections are automatically applied to that row.
4. Subsequent edits to protected columns on an already-created row must trigger the confirmation warning.
5. In `tbl_members`, read-only computed columns remain strictly locked to admins to prevent manual row insertion, with contact columns unlocked.
6. All automated tests pass cleanly.

## Answer

Resolved by:
1. Implementing `getLastCreatedRow(sheet, config)` in `src/ProtectionManager.js` to dynamically determine the range of populated, valid rows.
2. Confining `warningProtectedColumns` to rows 2 through `lastCreatedRow`, ensuring empty rows are completely unprotected.
3. Implementing `isCellProtected(sheet, row, col)` to verify whether a cell is covered by an active ProtectedRange before `onEdit` fires any warning toasts.
4. Reapplying protections upon synchronization (`syncMembersAndTenures()`) so newly committed rows are protected while keeping subsequent rows open.
5. Adding comprehensive unit tests in `test/validation.test.js` validating that editing uncreated rows triggers no warning, while editing existing created rows triggers the appropriate warning.
