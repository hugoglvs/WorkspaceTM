# WorkspacePM: Distribution & Deployment Guide

_Last updated 2026-10-12 (library state: ADR 0045)._

## Overview

WorkspacePM is a **Google Apps Script library** (`WorkspacePM`, source in `src/`). Each team spreadsheet (a **Consumer Sheet**) contains only a thin wrapper script (source in `consumer/`) that forwards menus, triggers and sidebar calls to the library.

- **One code base:** fixes and features land in the library; every linked sheet gets them.
- **Additive upgrades:** *Initialize / Upgrade* creates or widens tables and never deletes team rows.
- **Readable workspace:** tabs and native tables have readable names ("Members", "Contracts"…), found by their stored sheet ID, so renaming a tab is safe (ADR 0039).

| | Script ID |
| :--- | :--- |
| Library `WorkspacePM` | `1vyJiEQTSluHaPbh8-B9gSOj2SeciqPEzHpJ5Xug0I_NVhDvtYHdZ1b6v` |
| Reference Consumer Sheet "[NEW] ACN" (bound script) | `1IVORKENsN2r81QcA5hXa5qa-tgXub2RFFlhzD0-8fopojfAoQFixFujH` |

---

## 1. Prerequisites

- **Who sets it up:** the person who will **own the sheet and its triggers**. Every automatic job (edit handling, nightly sync, reminder and birthday emails, the Team Birthdays calendar) runs as this account and sends email from it.
- **Email volume:** a personal Gmail account can send to about **100 recipients a day** with Apps Script. The Monday reminder and birthday emails go to each Active member, so this is comfortable for a team of ~20 and tight beyond ~50.
- **Advanced services** (enabled in both the library and the consumer manifests): Google Sheets API v4, Google Tasks API v1, Google Calendar API v3.

---

## 2. Linking a New Consumer Sheet

### Option A: by hand in the Apps Script editor

1. Open the Google Sheet → **Extensions → Apps Script**. Rename the project (e.g. `Team Alpha – WorkspacePM`).
2. **Project Settings** (gear) → tick **Show "appsscript.json" manifest file in editor**.
3. Replace `appsscript.json` with the content of [`consumer/appsscript.json`](../consumer/appsscript.json). It declares the library and the three advanced services. For a production sheet, change the library entry to a numbered version (see §7):
   ```json
   { "userSymbol": "WorkspacePM", "libraryId": "1vyJiEQTSluHaPbh8-B9gSOj2SeciqPEzHpJ5Xug0I_NVhDvtYHdZ1b6v", "version": "<N>", "developmentMode": false }
   ```
4. Replace `Code.gs` with the full content of [`consumer/Code.js`](../consumer/Code.js). **Copy the file from the repo; do not retype it from an older guide.** Every menu item, sidebar button and trigger needs a top-level forwarder there (see §3).
5. Save, then reload the Google Sheet.

### Option B: with clasp (maintainers)

```bash
cd consumer
# point .clasp.json at the new sheet's bound script ID, then
~/.npm-global/bin/clasp push --force
```
Verify by pulling into `.scratch/live/consumer` and diffing (ADR 0033).

---

## 3. What the Wrapper Must Contain

`consumer/Code.js` is the source of truth. It defines:

| Group | Functions |
| :--- | :--- |
| Menu & lifecycle | `onOpen` (builds the menu), `onEdit` (forwards; steps aside when the installable trigger exists) |
| Menu actions | `setupOrUpgrade`, `showBaselineSidebar`, `showContractsTab`, `showBackupSidebar`, `showAssignSidebar`, `showTaskCalendar`, `connectMyGoogleTasks`, `showWorkspaceCheck`, `showLocationsTab`, `showHolidaysTab`, `showTaskTypesTab` |
| Trigger handlers | `installedTaskEditHandler` (edits), `nightlySyncJob` (2 am), `mondayPresenceReminder` (Mon 8 am), `dailyRemindersJob` (daily 9 am; `birthdayReminderJob` kept as an alias), `myTasksSyncJob` (hourly, one per connected person) |
| Sidebar / dialog calls (`google.script.run`) | `getSidebarInitialData`, `getMemberBaselines`, `saveMemberBaselines`, `submitMemberLeave`, `getBackupSidebarData`, `saveInterimBackup`, `cancelInterimBackup`, `getAssignSidebarData`, `assignTasks`, `stopRecurringRule`, `runWorkspaceCheck`, `removePresenceDuplicates` |

> A sidebar or dialog served by the library calls the **consumer** script. When a new library feature adds a sidebar call, a menu item or a trigger, the wrapper must be updated and pushed too, or the button fails with "function not found".

---

## 4. First Initialization

1. Reload the sheet. The **Project Management** menu appears.
2. **Project Management → Initialize / Upgrade Spreadsheet.**
3. Approve Google's authorization prompt (Sheets, Drive, Calendar, Tasks, Gmail sending, triggers). It appears again whenever a release adds a scope (e.g. the Calendar API for birthdays).
4. The run takes about 2 minutes and:
   - creates the tables with readable names, validations, typography and protections;
   - fills **Locations** with OneRoof (Office), Remote, JLL, C04, C60 (Client site), **Task Types** with the default types and **Holidays** with the French public holidays of this year and next;
   - shows the working tabs and hides the data tabs;
   - installs the four triggers (§6);
   - runs a first full nightly sync.

| Visible tabs (daily use) | Hidden tabs (data; menu items open some of them) |
| :--- | :--- |
| Members, Weekly Presence, Projects, Tasks, Calendar, Events | Locations, Holidays, Task Types, Recurring Tasks, Contracts, Routines, Leaves, Presence Log, Backups, KPI Data |

---

## 5. The Menu

| Menu item | What it does |
| :--- | :--- |
| **Initialize / Upgrade Spreadsheet** | Creates or upgrades tables, protections and triggers, then runs the nightly sync. Safe to run any time. |
| **Declare Presence & Leaves** | Sidebar: weekly routine (AM/PM per weekday) and leave requests. |
| **Declare Contract** | Unhides and opens the Contracts tab on its first empty row: one row per contract (member, role, department, start, end, type, FTE). To change contract, first set the end date of the current one. Typing a new name in Members opens it automatically. Contact details and birthday are typed in Members. |
| **Interim Backups** | Sidebar: cover a project for given dates. Shows coverage gaps (click one to fill the dates) and lets you cancel a planned backup. |
| **Assign Tasks** | Sidebar: give a task type to several people (by role and/or name), once or every week/month (each person, or a rotation that skips people on leave). Lists and stops recurring tasks. |
| **Calendar** | Opens the Calendar tab: pick a Type and a Project, then members × weeks (or months). *To do* assigns, *Done* marks done, clearing cancels. Refuses someone on approved leave. |
| **Connect my Google Tasks** | Each member clicks it once: their own tasks then sync every hour with their Google Tasks (both ways). |
| **Check Workspace** | Lists what needs attention: members without a contract, projects without a backup, people leaving, coverage gaps, not logged today, undeclared days this week, overdue tasks, tasks falling on a leave, members not connected to Google Tasks, duplicate presence rows. |
| **Edit Locations** | Shows the hidden Locations tab (Location \| Type \| Active). Untick *Active* to retire a location; history stays valid. |
| **Edit Holidays** | Shows the hidden Holidays tab. Add bridges or office closures; delete a day the team works. |
| **Edit Task Types** | Shows the hidden Task Types tab (Type, Frequency, Due rule, Mandatory, Active). |

---

## 6. Automatic Jobs

| Job | Trigger | What it does |
| :--- | :--- | :--- |
| Edit handling | Installable on-edit (`installedTaskEditHandler`) | Every grid edit: IDs, validation toasts, Google Tasks / Calendar push, Weekly Presence saves, member sync, birthday format, location warnings, coverage re-check. The simple `onEdit` steps aside once this trigger exists (one run per edit). |
| Nightly sync | Daily, 2 am (`nightlySyncJob`) | Member sync, presence compile, Weekly Presence, KPI Data, recurring tasks planned ahead (rotation turns moved off approved leaves), Calendar refresh, coverage check (emails + tasks for conflicts), Google Tasks and Calendar sync, next year's holidays, Team Birthdays calendar. Each step is timed in the log and the run stops cleanly before Apps Script's 6-minute limit. |
| Monday reminder | Mondays, 8 am (`mondayPresenceReminder`) | Emails each Active member who has a working day this week with no presence, routine or leave. |
| Daily reminders | Daily, 9 am (`dailyRemindersJob`) | Birthdays: on the working day before (Friday for Sat–Mon), every other Active member gets an "Add to my Google Calendar" link. Tasks: an email 3 working days before a typed task is due and on the first working day after if still open. |
| My Google Tasks | Hourly, one per connected member (`myTasksSyncJob`) | Syncs that person's tasks with their own Google Tasks list, both ways. |

The first four are installed (once each) by *Initialize / Upgrade*; `myTasksSyncJob` by each member's *Connect my Google Tasks*. Check them in the Apps Script editor → **Triggers**.

---

## 7. Releasing and Upgrading the Library

### Development mode (current test setup)

The reference sheet uses the library in **development mode** (`"developmentMode": true`): every `clasp push` of `src/` is live at once. Each push is tagged `deploy-YYYYMMDD-HHMM` in git so any earlier state can be restored.

### Production

1. In the library project: **Deploy → New deployment → Library**, with a description. This creates version *N*.
2. In each consumer sheet: Apps Script → **Libraries → WorkspacePM** → choose version *N* (development mode off) → **Save**. Or edit the `version` in `appsscript.json`.
3. If the release changed `consumer/`, update the wrapper (§2) as well.
4. In the sheet: **Project Management → Initialize / Upgrade Spreadsheet**. Approve any new authorization.

Recommended: keep one test copy of the sheet in development mode and move the real sheets to a numbered version.

---

## 8. Optional Settings

These are **document properties**, which the Apps Script UI does not show. Set them once by running a small function in the consumer editor (select it, **Run**), then delete it:

```javascript
function configureWorkspacePM() {
  WorkspacePM.setCalendarId("department-events@group.calendar.google.com"); // Events tab → this calendar (default: owner's "primary")
  WorkspacePM.setTasksListId("<tasks list id>");                             // Tasks tab → this list (default: "@default")
  PropertiesService.getDocumentProperties().setProperty("ADMIN_EMAILS", "lead@example.com,pm@example.com");
}
```

`ADMIN_EMAILS` adds people who may edit strictly protected columns (computed ones such as Members Role, Status, Today, and IDs), on top of the sheet owner. Note: the `ADMIN_EMAILS` line writes the consumer's own document properties; if admins are not picked up, ask a maintainer to set it through the library.

---

## 9. Troubleshooting

| Symptom | Where to look / what to do |
| :--- | :--- |
| A menu item or sidebar button says the function is not found | The wrapper is older than the library: copy the current `consumer/Code.js` (§2). |
| "WorkspacePM Library Not Found" | The library is not linked or its identifier is not exactly `WorkspacePM`. |
| Something did not update after an edit | Apps Script editor → **Executions**: open the latest `installedTaskEditHandler` run. Edits over 3 s log an `[edit timing]` line with the slow part. |
| Nightly work missing | **Executions** → `nightlySyncJob`: one line per step with its duration; `Out of time, skipped: …` lists steps cut by the 6-minute budget. |
| "The nightly sync is running…" toast on an edit | The edit happened during the nightly run; the value is saved, re-enter it a few minutes later if its follow-up (task, event, presence row) is missing. |
| Red cells in a row | A required column of that row is empty (validation highlight). |
| Duplicate presence rows | **Check Workspace → Remove duplicates** (keeps the row Weekly Presence shows). |
| Emails stop mid-day | Daily email quota of the owner account reached; they resume the next day. |

Architecture decisions are recorded in `docs/adr/` (0032–0044 cover the current design); domain terms are in `GLOSSARY.md`.
