# ADR 0041: Full-table reads, One Run per Edit, Check Workspace, Monday Reminder and Interim Backups Sidebar

## Status
Accepted (2026-10-12)

## Context
1. Reads were capped at a fixed row count (`Math.min(getMaxRows(), 2000)` for Presence Log). Presence Log gains one row per Member per working day, so after ~2,000 rows the daily compile no longer saw existing days and appended duplicates every night, ignoring manual entries beyond the cap.
2. Both the simple `onEdit` and the installable edit trigger ran for every edit, one after the other behind the sync lock: twice the work, and Tasks / Events handlers ran twice.
3. Nothing showed, in one place, what needs attention; members who forget to declare a week were never reminded; adding an interim backup meant unhiding the Backups tab.

## Decision
1. **Reads never stop before the last row with data.** `readRowLimit(sheet, cap)` (`src/Tables.js`) returns `max(getLastRow(), min(getMaxRows(), cap))`; every former `Math.min(x.getMaxRows(), N)` uses it. The cap now only limits how far empty rows are scanned.
2. **One run per edit.** Once the installable trigger (`installedTaskEditHandler`) exists, it runs the whole edit handler (`handleSheetEdit`); the simple `onEdit` returns at once. A document property `EDIT_TRIGGER_ACTIVE` records the trigger; it is refreshed by *Initialize / Upgrade*, *Enable Live Task Sync* and the nightly job (all authorised), because the simple trigger cannot list triggers. Without the installable trigger, the simple `onEdit` works as before.
3. **Check Workspace** (menu, `src/WorkspaceCheck.js`, dialog `WorkspaceCheckDialog.html`) lists: Members without a contract, open Projects without a backup, Project Backup Conflicts and warnings (no emails, no tasks), Members not logged today, Undeclared Days this week, and duplicate Presence Log rows. Each section links to the fix: Members & Contracts, Interim Backups, Declare Presence, or *Remove duplicates* (keeps the last row of each member-day, the one Weekly Presence shows).
4. **Monday reminder.** A weekly trigger (`mondayPresenceReminder`, Mondays 8 am, installed by *Initialize / Upgrade*) emails each Active Member who has an Undeclared Day this week, listing the days and linking to Weekly Presence. Members without an email are logged, not emailed.
5. **Interim Backups sidebar** (menu, `src/BackupService.js`, `BackupSidebar.html`): pick a project, see its lead, Project Backup, uncovered periods (click one to fill the dates) and current / planned Designated Backups; add one (start and end required; the lead is refused; a leave of the chosen person is flagged) or cancel one (status `Cancelled`, row kept). Saving re-checks coverage without emails; the nightly check still notifies.

## Consequences
- The "Organize interim backup" task and conflict emails now point to *Project Management → Interim Backups*.
- `assignProjectBackup` no longer re-sends coverage emails after a save.
- Edit handling runs as the trigger owner once the installable trigger exists (it already did for Tasks and Events).
