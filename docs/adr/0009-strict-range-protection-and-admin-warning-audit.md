# 9. Strict Protected Range Enforcement and Admin Edit Warnings

Date: 2026-10-10

## Status

Accepted

## Context

Previously, column protections across operational tables (`tbl_members`, `tbl_presence_baselines`, etc.) were configured with `setWarningOnly(true)`. This allowed any user to dismiss the warning dialog and override protected values, defeating the purpose of database integrity for calculated columns (e.g., `current_role`, `tenure_status`, `today_presence`).

The requirement is:
1. Prevent regular team members from changing protected columns entirely.
2. Display warning notifications only to authorized admins who are permitted to edit them.

## Decision

1. **Strict Protection Enforcement**:
   All declared `protectedColumns` in `SCHEMA_SPEC` are locked with `protection.setWarningOnly(false)`.
2. **Authorized Editor Restriction**:
   - Authorized editors include the spreadsheet owner (`ss.getOwner()`), the script executor (`Session.getEffectiveUser()`), and any emails configured in the `ADMIN_EMAILS` Document Property (`getAuthorizedAdmins()`, `setAuthorizedAdmins()`).
   - All other editors are removed via `protection.removeEditors()`, and domain editing is disabled (`protection.setDomainEdit(false)`).
   - Non-admin collaborators are strictly prevented by Google Sheets from modifying protected cells.
3. **Targeted Admin Warnings in `onEdit`**:
   Because only authorized admins can edit protected cells, any manual edit on a protected column detected in `onEdit(e)` triggers a targeted warning toast to the admin:
   `"Admin Warning: You edited protected column '[colName]' in [sheetName]. Manual edits to calculated columns may be overwritten during synchronization."`

## Consequences

- Regular team members cannot modify protected or computed database keys.
- Permitted contact columns (such as `accenture_email` and `phone_number` in `tbl_members`) remain editable by team members.
- Authorized admins receive explicit, non-blocking warning feedback whenever they modify protected columns.
