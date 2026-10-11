# ADR 0031: Minimal Consumer Interface and Automated Nightly Maintenance Sync

## Status
Accepted

## Context
1. Previously, the `Project Management` custom menu presented 8 distinct administrative and synchronization actions to end users, including manual buttons for compiling presence logs, syncing Google Tasks, syncing Google Calendar, and refreshing matrix views.
2. User feedback highlighted that spreadsheets should behave as modern applications: users should never need to manually invoke batch scripts or remember to click "synchronize" buttons.
3. Operations during business hours are already handled reactively (via simple and installable `onEdit` triggers for task edits, calendar edits, matrix updates, tenure declarations, and sidebar leave submissions).
4. Periodic maintenance (such as rolling forward presence logs, refreshing analytics dashboards, pulling completed statuses from Google Tasks, and reconciling backup conflicts) can be delegated completely to an automated nightly time-driven trigger.

## Decision

1. **Ultra-Minimal Consumer UI (`src/ConsumerWrapper.js`):**
   - Streamlined the spreadsheet menu to exactly two user-facing actions:
     - **Initialize / Upgrade Spreadsheet (`setupOrUpgrade`):** Used during initial setup or version upgrades. Automatically verifies schemas, applies Named Tables, installs background triggers, and executes an initial maintenance sync.
     - **Declare Presence & Leaves (`showBaselineSidebar`):** Opens the interactive visual sidebar for weekly routine submissions and leave requests.
   - Removed all manual synchronization and compilation menu items from the default UI.

2. **Automated Dual-Trigger Architecture (`src/Code.js`):**
   - Implemented `installAllTriggers(targetSpreadsheet)`:
     - **Installable `onEdit` Trigger (`installedTaskEditHandler`):** Handles real-time OAuth-authorized synchronization for Google Tasks and Google Calendar immediately upon cell modification.
     - **Daily Time-Driven Trigger (`nightlySyncJob`):** Scheduled daily at 2:00 AM (`ScriptApp.newTrigger("nightlySyncJob").timeBased().everyDays(1).atHour(2).create()`).
   - Integrated trigger installation directly into `setupOrUpgrade()`, ensuring one-click setup without manual trigger configuration.

3. **Nightly Maintenance Job (`nightlySyncJob`):**
   - Reconciles all system domains automatically at 2:00 AM:
     1. Syncs tenures and directory members (`sync`).
     2. Compiles daily presence logs from baselines and approved leaves (`compilePresence`).
     3. Refreshes analytical views: Weekly Matrix View (`refreshWeeklyMatrix`), Active Responsibles Dashboard (`refreshActiveResponsibles`), and KPIs (`refreshKpis`).
     4. Synchronizes Google Tasks (`syncTasks`): pulls external completion statuses and pushes unsynced items.
     5. Synchronizes Google Calendar (`syncCalendarEvents`).
     6. Audits simultaneous absence backup conflicts on active projects (`checkBackupConflicts`).

4. **Preserved Canonical Forwarders:**
   - Retained all canonical library action forwarders in `ConsumerWrapper.gs` to allow script editor execution, macros, or automated integrations without cluttering the spreadsheet UI.

## Consequences
- Clean, uncluttered user experience: team members interact exclusively with the Google Sheet and the Presence & Leaves sidebar.
- Zero manual synchronization overhead; background triggers handle live edits and nocturnal reconciliation automatically.
- Idempotent one-click setup: clicking "Initialize / Upgrade Spreadsheet" sets up tables, repairs triggers, and runs a complete sync.
- 127 automated unit tests passing (`node --test test/*.js`).
