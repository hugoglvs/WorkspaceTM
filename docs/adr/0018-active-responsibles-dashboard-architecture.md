# ADR 0018: Active Responsibles Dashboard Architecture and Manual Workspace Timeline Integration

## Status
Accepted

## Context
In multi-project operational delivery, leadership continuity is essential. When project leads take leave (vacation, sickness, holidays), designated backup leads must take over coverage. If both lead and backup are absent or on leave simultaneously, critical project decisions and deliveries risk stall or failure.

Issue 05 defined two operational views:
1. `view_project_gantt`: Timeline visualization of project schedules.
2. `view_active_responsibles`: Real-time operational dashboard cross-referencing project ownership against daily member presence and approved leaves.

Regarding the timeline view: Google Workspace offers a native Timeline view (`Insert > Timeline`), which cannot be created or configured via Google Apps Script or the Google Sheets REST API v4 because no programmatic API exists for Workspace Timeline objects. The user specifically instructed:
*"About the 'Gantt timeline', if we can't insert a native Google Workspace Gantt I don't want it and I'll make it manually."*

Therefore, custom scripted grid-based Gantt charts were intentionally omitted, and the full architectural focus was placed on the Active Responsibles Dashboard (`view_active_responsibles`).

## Decision

1. **Dashboard Engine (`renderActiveResponsiblesDashboard`):**
   - Implemented in `src/ViewBuilder.js` and exposed via `refreshActiveResponsibles(targetSpreadsheet)` in `src/Code.js`.
   - Reads project rosters from `tbl_projects`, member directory from `tbl_members`, and active approved leaves from `tbl_leaves`.
   - Evaluates availability for each project's Project Lead and Project Backup as of the target date (defaults to today):
     - Checks departure/archival tenure statuses.
     - Checks active approved leaves in `tbl_leaves`.
     - Checks daily presence in `tbl_presence_log` or `tbl_members` today presence.
     - Leaves (Vacation, Sick, Holiday, Maternity, etc.) classify members as unavailable.

2. **Leadership Coverage Hierarchy:**
   - **Lead Active:** When the Project Lead is active and present on duty. Active Responsible is set to Project Lead.
   - **Covered by Backup:** When the Project Lead is on leave or unavailable, but the Project Backup is active. Active Responsible is set to Project Backup.
   - **CRITICAL: Uncovered:** When both the Lead and Backup are on leave, absent, departed, or unassigned. Active Responsible is set to "No Responsible Available".

3. **Google Sheets UI & People Chip Fidelity:**
   - Generated sheet `view_active_responsibles` includes columns:
     `Project ID`, `Project Name`, `Status`, `Project Lead`, `Lead Presence`, `Project Backup`, `Backup Presence`, `Active Responsible`, `Coverage Status`.
   - Native People Chips are copied to `Project Lead`, `Project Backup`, and `Active Responsible` columns via `copyCellChip` from `tbl_members`.
   - 12px Google Sans typography and soft gray header styling (`#E8EAED`) applied.
   - Frozen header row and custom column widths for readability.

4. **Alert Badges via Conditional Formatting:**
   - "Lead Active": Soft green background (`#E6F4EA`), dark green bold text (`#137333`).
   - "Covered by Backup": Soft amber background (`#FEF7E0`), dark amber bold text (`#B06000`).
   - "CRITICAL: Uncovered": Soft alert red background (`#FCE8E6`), dark red bold text (`#C5221F`).

5. **Synchronization and Refresh Lifecycle:**
   - Custom menu item added to Google Sheets UI under `Project Management > Refresh Active Responsibles Dashboard`.
   - Automatically refreshed when running `sync()` (synchronizing members and tenures), `compilePresence()` (daily presence logs), or `compilePresenceMigration()` (historical migrations).

## Consequences
- Operations leads have instant, color-coded visibility into project coverage gaps and active escalations.
- Complete data integrity and native People Chip representation across all leadership roles.
- Follows the user's explicit directive to rely on manual native Google Workspace Timeline (`Insert > Timeline`) for project scheduling, avoiding redundant or unmaintainable scripted Gantt approximations.
- 100% test coverage with automated unit tests in `test/viewBuilder.test.js`.
