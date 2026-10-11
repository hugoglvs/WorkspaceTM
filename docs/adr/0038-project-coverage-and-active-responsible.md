# ADR 0038: Per-Day Project Coverage and the Active Responsible Column

## Status
Accepted (2026-10-12), resolves issue 23. Supersedes the conflict rules of ADR 0023 and the dashboard of ADR 0018.

## Context
1. The user wants Projects to show who is in charge today (`active_responsible`) instead of a separate `view_active_responsibles` tab. A backup should exist for every new project, and leave declarations must always be checked so that every day stays covered. Leave *requests* must warn too.
2. The old check compared only the lead's and the backup's approved leaves. It read the backup from `tbl_projects` only, and accepted an interim backup only if one row spanned the whole gap.

## Decision
1. **Backup model.** The permanent backup is `tbl_projects.backup_member` (editable, ADR 0034 flexibility), mirrored into `tbl_backups` as the one Active row with no `end_date` (`syncProjectBackups`). Changing it ends the old row yesterday and adds a new one from today. Interim backups are dated `tbl_backups` rows. Declaring a leave never edits backup rows.
2. **Per-day rule** (`checkProjectBackupConflicts`), for each working day from today to the last relevant leave (max 366 days, weekends skipped): Project Lead if present → else an interim backup active that day → else the permanent backup → else uncovered. Absence = an approved leave, a requested leave, or (today only) an absence shown in `today_presence`.
3. **Outcomes.** Uncovered days caused only by approved absences are a **conflict**: red row, one email to the lead and a manager, one "Organize interim backup" task. Days involving a requested leave are a **warning**: orange row and a warning email, with no task. Consecutive days (across weekends) are grouped into periods.
4. **`active_responsible`** (read-only column 8 of `tbl_projects`) holds today's result: `Name`, `Name (interim)`, `Name (backup)` or `⚠ Nobody available`. It is refreshed on every check (edits of projects, leaves or backups, the sidebar, compile, nightly).
5. **New projects:** `onEdit` asks with a toast for a backup when a named project has none, and the sync gives a project without an ID a `PRJ-XXXXX` ID.
6. `view_active_responsibles` and its code are removed (deleted on upgrade). The duplicate `refreshActiveResponsibles` global (`Code.js` + `ViewBuilder.js`) is gone with it.
7. **Additive table migration:** when the schema gains a column, `createGoogleSheetsTable` widens the existing native table (`updateTable` range).

## Consequences
- Past gaps are no longer reported; only today and the future matter.
- There is no guided way yet to add an interim backup (users unhide `tbl_backups`, prompted by the task). A sidebar for it is a natural follow-up.
