# 23: Projects with Read-Only Backup Column from tbl_backups

**Status:** open

**Blocked by:** _(none)_

**What to build:** Projects gets a read-only `active_responsible` column (who is in charge today), replacing `view_active_responsibles`. Creating a project creates a `tbl_backups` row to fill in. The daily coverage check warns whenever, on any day, a project has neither its lead nor a backup present.

**Backup model (proposed 2026-10-12; the user picked active_responsible over a read-only backup column):**
- Permanent backup = a `tbl_backups` row with no `end_date`. Changing it = end the old row and add a new one.
- Interim backup = a dated row, added when a gap is detected (the existing "Organize interim backup" task).
- Declaring a Leave never edits backup rows.
- Per day: lead if present → else an interim backup if present → else the permanent backup if present → else conflict (red row, email, task).

**Gaps found in today's `checkProjectBackupConflicts` (2026-10-12):**
- It reads the backup from `tbl_projects.backup_member` instead of `tbl_backups`.
- A conflict counts as covered only if a *single* designated backup spans the whole overlap; two consecutive interim backups are not recognised.
- `Requested` leaves are ignored until approved (decide: warn early?).
