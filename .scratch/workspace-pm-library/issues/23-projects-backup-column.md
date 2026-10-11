# 23: Projects with Read-Only Backup Column from tbl_backups

**Status:** open

**Blocked by:** user decision on the backup model

**What to build:** `tbl_projects.backup_member` becomes a read-only column showing today's backup computed from `tbl_backups`. Creating a project creates a backup row to fill in. `view_active_responsibles` is removed (optionally replaced by an `active_responsible` column). The daily coverage check warns whenever a Leave leaves a project with neither its lead nor any backup present, on any day.
