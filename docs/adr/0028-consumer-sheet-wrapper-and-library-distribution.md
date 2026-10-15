# ADR 0028: Consumer Sheet Client Wrapper and Library Distribution Architecture

## Status
Accepted

## Context
1. The project management codebase has grown into an extensive suite of engines covering schemas, typography, protections, presence baselines, historical migrations, matrix views, active responsibles, remote work KPIs, Google Tasks synchronization, and Google Calendar event management.
2. In enterprise Google Workspace environments, duplicating this codebase across dozens of departmental spreadsheets causes code drift, maintenance debt, and security update challenges.
3. Teams require a lightweight container-bound client script (`ConsumerWrapper.gs`) that imports an immutable, versioned Google Apps Script Library (`WorkspaceTM`), delegates custom UI menus and triggers, and provides clear user alerts if dependencies or permissions are unconfigured.

## Decision

1. **Public Library Boundary Standardization (`src/Code.js`):**
   - Standardized top-level public library entry points:
     - `setupOrUpgrade(targetSpreadsheet)`: Initializes 8 core operational tables and Named Tables styling, or applies schema migrations.
     - `compilePresence()`: Compiles baseline routines and approved leaves into daily entries in `tbl_presence_log`.
     - `refreshWeeklyMatrix(targetSpreadsheet)`, `refreshActiveResponsibles(targetSpreadsheet)`, `refreshKpis(targetSpreadsheet)`: Refreshes analytical views and dashboard sheets.
     - `syncTasks(targetSpreadsheet)`: Pushes and pulls tasks against Google Tasks.
     - `syncCalendarEvents(targetSpreadsheet)` / `syncEvents(targetSpreadsheet)`: Synchronizes events with Google Calendar.
     - `showBaselineSidebar()`: Opens the presence baseline and leaves submission panel.
     - `installLiveTaskSync(targetSpreadsheet)`: Installs an authorized trigger for real-time task and calendar editing.

2. **Thin Consumer Sheet Client Wrapper (`src/ConsumerWrapper.js`):**
   - Created a clean, container-bound client template providing:
     - Custom menu `onOpen`:
       - "Initialize / Upgrade Spreadsheet"
       - "Compile Daily Presence Logs"
       - "Refresh Dashboards (Weekly Matrix & Responsibles)"
       - "Sync Google Tasks"
       - "Sync Calendar Events"
       - "Presence & Leaves Sidebar"
       - "Enable Real-Time Triggers"
     - `onEdit(e)` forwarder to `WorkspaceTM.onEdit(e)`.
     - `safeInvokeLibrary(methodName, args)` guard:
       - Displays a clear modal dialog if the `WorkspaceTM` library is unlinked.
       - Dispatches toast notifications if an action is unsupported by the linked library version.
       - Catches and logs execution errors gracefully with actionable spreadsheet toasts.

3. **Distribution & Deployment Guide (`docs/distribution-guide.md`):**
   - Published a comprehensive guide detailing:
     - How to link the library via Script ID in Google Apps Script.
     - How to enable the Google Sheets API v4 Advanced Service.
     - How to pin immutable library versions.
     - How to configure Document Properties (`GOOGLE_CALENDAR_ID`, `GOOGLE_TASKS_LIST_ID`).
     - One-time OAuth approval and trigger installation instructions.

## Consequences
- Clean separation of concerns between core library business logic and departmental consumer spreadsheets.
- Consumer spreadsheets only need to contain a ~100 line wrapper script.
- Centralized deployment allows instantaneous bug fixes and backwards-compatible improvements across all spreadsheets.
- 122/122 automated unit tests passing across the test suite (`test/*.js`).
