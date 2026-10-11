# Google Workspace Project Management - Library Distribution & Deployment Guide

## Overview

The **Google Workspace Project Management System** is architected as a centralized **Google Apps Script Library (`WorkspacePM`)**. Instead of copying thousands of lines of code into every spreadsheet, individual team spreadsheets include only a thin client wrapper (`consumer/Code.js` in the repo, `Code.gs` in Apps Script) that references the versioned library.

This design ensures:
- **Centralized Bug Fixes & Features:** Updating the library automatically benefits all teams without touching spreadsheet code.
- **Data Safety:** Library updates are purely additive and never delete or overwrite existing team rows.
- **Version Pinning:** Teams can lock their spreadsheet to an immutable numbered version (e.g., Version 1) and upgrade on their own schedule.
- **Consistent Schemas & Named Tables:** All sheets adhere to identical schemas, 12px "Google Sans" typography, and native Google Sheets Tables.

---

## 1. Prerequisites

Before linking the library to a team spreadsheet:
1. **Google Sheets API Service:** The Sheets Advanced API (v4) must be enabled in the Apps Script project to support native Google Sheets Tables.
2. **Permissions:** The Team Lead / Administrator must have permission to create Apps Script triggers and authorize Google Calendar and Google Tasks integrations.

---

## 2. Linking the Library to a Consumer Spreadsheet

Follow these steps to link the Library to a new or existing Google Sheet:

### Step 1: Open Apps Script
1. Open the target Google Sheet.
2. Click **Extensions** > **Apps Script** in the top menu.
3. Rename the Apps Script project to your team's name (e.g., `Team Alpha - Workspace PM`).

### Step 2: Enable the Google Sheets API Advanced Service
1. In the left navigation pane of the Apps Script editor, click **+** next to **Services**.
2. Select **Google Sheets API** (identifier: `Sheets`, version: `v4`).
3. Click **Add**.

### Step 3: Add the `WorkspacePM` Library
1. In the left navigation pane, click **+** next to **Libraries**.
2. In the **Script ID** field, paste the library's Script ID:
   ```text
   <WORKSPACE_PM_LIBRARY_SCRIPT_ID>
   ```
   *(Retrieve this Script ID from the Library Project Settings > Script ID)*
3. Click **Look up**.
4. Configure the settings:
   - **Version:** Select the latest stable numbered version (e.g., `1`). Do not select "HEAD" in production to prevent unexpected changes.
   - **Identifier:** `WorkspacePM` *(must be exact)*.
5. Click **Add**.

---

## 3. Adding the Client Wrapper (`ConsumerWrapper.gs`)

In the Apps Script editor of the team's spreadsheet, replace the contents of `Code.gs` (or create a new file named `ConsumerWrapper.gs`) with the following lightweight wrapper:

```javascript
/**
 * Google Workspace Project Management - Consumer Sheet Client Wrapper
 * Powered by WorkspacePM Library
 */

/**
 * Creates the Project Management menu on spreadsheet open.
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Project Management")
    .addItem("Initialize / Upgrade Spreadsheet", "setupOrUpgrade")
    .addSeparator()
    .addItem("Declare Presence & Leaves", "showBaselineSidebar")
    .addToUi();

  // Forward onOpen lifecycle to library for diagnostics and audits without rebuilding menu
  try {
    var lib = getLibraryInstance();
    if (lib) {
      if (typeof lib.onOpenAudit === "function") {
        lib.onOpenAudit(SpreadsheetApp.getActiveSpreadsheet());
      } else if (typeof lib.onOpen === "function") {
        lib.onOpen(e, true);
      }
    }
  } catch (err) {
    Logger.log("Library onOpen note: " + err);
  }
}

/**
 * Forwarder for onEdit simple trigger.
 */
function onEdit(e) {
  try {
    var lib = getLibraryInstance();
    if (lib && typeof lib.onEdit === "function") {
      lib.onEdit(e);
    }
  } catch (err) {
    Logger.log("Library onEdit error: " + err);
  }
}

/**
 * Resolves the WorkspacePM library instance.
 */
function getLibraryInstance() {
  if (typeof WorkspacePM !== "undefined") {
    return WorkspacePM;
  }
  if (typeof global !== "undefined" && typeof global.WorkspacePM !== "undefined") {
    return global.WorkspacePM;
  }
  return null;
}

/**
 * Safely invokes a library method with user-friendly error handling.
 */
function safeInvokeLibrary(methodName, args) {
  var lib = getLibraryInstance();
  if (!lib) {
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      "WorkspacePM Library Not Found",
      "The 'WorkspacePM' library is not linked to this Google Sheet.\n\n" +
      "To link the library:\n" +
      "1. In Apps Script, click '+' next to 'Libraries'.\n" +
      "2. Paste the Library Script ID and click 'Look up'.\n" +
      "3. Set Identifier to 'WorkspacePM' and select the latest version.",
      ui.ButtonSet.OK
    );
    return null;
  }

  if (typeof lib[methodName] !== "function") {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Action '" + methodName + "' is not supported in the currently linked WorkspacePM version.",
      "Unsupported Action",
      6
    );
    return null;
  }

  try {
    return lib[methodName].apply(lib, args || []);
  } catch (err) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && typeof ss.toast === "function") {
      ss.toast("Error executing " + methodName + ": " + (err.message || err), "Action Error", 6);
    }
    Logger.log("Error executing " + methodName + ": " + err);
    throw err;
  }
}

/**
 * Canonical Library Action Forwarders
 * Defines top-level functions matching WorkspacePM library methods so that any
 * menu items, buttons, or macros targeting these canonical names execute seamlessly.
 */
function showBaselineSidebar() { return safeInvokeLibrary("showBaselineSidebar"); }
function init() { return safeInvokeLibrary("init"); }
function setupOrUpgrade() { return safeInvokeLibrary("setupOrUpgrade"); }
function sync() { return safeInvokeLibrary("sync"); }
function validateAndReport() { return safeInvokeLibrary("validateAndReport"); }
function compilePresence() { return safeInvokeLibrary("compilePresence"); }
function compilePresenceMigration() { return safeInvokeLibrary("compilePresenceMigration"); }
function refreshWeeklyMatrix() { return safeInvokeLibrary("refreshWeeklyMatrix"); }
function refreshActiveResponsibles() { return safeInvokeLibrary("refreshActiveResponsibles"); }
function refreshKpis() { return safeInvokeLibrary("refreshKpis"); }
function checkBackupConflicts() { return safeInvokeLibrary("checkBackupConflicts"); }
function syncTasks() { return safeInvokeLibrary("syncTasks"); }
function syncCalendarEvents() { return safeInvokeLibrary("syncCalendarEvents"); }
function installLiveTaskSync() { return safeInvokeLibrary("installLiveTaskSync"); }
function installAllTriggers() { return safeInvokeLibrary("installAllTriggers"); }
function nightlySyncJob() { return safeInvokeLibrary("nightlySyncJob"); }
function cleanupForeignTasks() { return safeInvokeLibrary("cleanupForeignTasks"); }
function debugUserIdentity() { return safeInvokeLibrary("debugUserIdentity"); }

/**
 * Client Aliases (for existing menu references and backwards compatibility)
 */
function clientShowSidebar() { return showBaselineSidebar(); }
function clientSetupOrUpgrade() { return setupOrUpgrade(); }
function clientSyncMembersAndTenures() { return sync(); }
function clientCompilePresence() { return compilePresence(); }
function clientRefreshDashboards() {
  var r1 = refreshWeeklyMatrix();
  var r2 = refreshActiveResponsibles();
  var r3 = refreshKpis();
  return { weeklyMatrix: r1, activeResponsibles: r2, kpis: r3 };
}
function clientSyncTasks() { return syncTasks(); }
function clientSyncCalendarEvents() { return syncCalendarEvents(); }
function clientInstallTriggers() { return installAllTriggers(); }

function installedTaskEditHandler(e) {
  var lib = getLibraryInstance();
  if (lib && typeof lib.installedTaskEditHandler === "function") {
    return lib.installedTaskEditHandler(e);
  }
  return null;
}

function installedConsumerTriggerHandler(e) {
  return installedTaskEditHandler(e);
}

// Sidebar API forwarders for google.script.run
function getSidebarInitialData() { return safeInvokeLibrary("getSidebarInitialData"); }
function getMemberBaselines(member) { return safeInvokeLibrary("getMemberBaselines", [member]); }
function saveMemberBaselines(payload) { return safeInvokeLibrary("saveMemberBaselines", [payload]); }
function submitMemberLeave(payload) { return safeInvokeLibrary("submitMemberLeave", [payload]); }
```

Save the file (**Ctrl+S** / **Cmd+S**).

---

## 4. Configuration (Document Properties)

By default, the system works out of the box with personal calendars (`"primary"`) and personal Google Tasks (`"@default"`). If your department uses a shared team calendar or a specific Google Tasks list:

1. In the Apps Script editor, open **Project Settings** (gear icon on the left).
2. Scroll down to **Script Properties** / **Document Properties**.
3. Add the desired configuration keys:
   - `GOOGLE_CALENDAR_ID`: The calendar ID of the team/department calendar (e.g., `department-events@group.calendar.google.com`). Defaults to `"primary"`.
   - `GOOGLE_TASKS_LIST_ID`: The ID of a specific Google Tasks list. Defaults to `"@default"`.

---

## 5. One-Time Setup & Initialization

1. Reload the Google Sheet.
2. Notice the streamlined custom menu: **Project Management** appears in the menu bar.
3. Click **Project Management** > **Initialize / Upgrade Spreadsheet**.
4. **OAuth Authorization:**
   - Google will display an **"Authorization Required"** prompt.
   - Click **Continue**, select your corporate Google account, and click **Allow**.
5. The system automatically performs a complete setup:
   - Creates/verifies all 9 operational database tables with native Google Sheets Named Tables, 12px "Google Sans" typography, validations, and protections (`tbl_members`, `tbl_tenures`, `tbl_presence_baselines`, `tbl_leaves`, `tbl_presence_log`, `tbl_projects`, `tbl_tasks`, `tbl_events`, `tbl_backups`).
   - Automatically installs the **real-time onEdit trigger** (`installedTaskEditHandler`) for live Google Tasks and Google Calendar synchronization.
   - Automatically installs the **nightly time-driven maintenance trigger** (`nightlySyncJob`) scheduled every night at 2:00 AM.
   - Performs an immediate full initial synchronization across presence logs, matrix dashboards, project responsibles, tasks, and calendars.

---

## 6. Daily Operations & Workflows

End users never need to run technical sync scripts or batch commands. The interface is intentionally minimal:

| Menu Action | Description |
| :--- | :--- |
| **Initialize / Upgrade Spreadsheet** | Verifies table schemas, Named Tables, formatting, protections, ensures all triggers are installed, and runs a complete maintenance sync. |
| **Declare Presence & Leaves** | Opens the interactive visual sidebar to submit presence routines, view personal status checklist, and schedule leaves. |

### Fully Automated Background Operations

| Automated Process | Mechanism | Execution Frequency |
| :--- | :--- | :--- |
| **Live Task Sync** | Installable `onEdit` Trigger | Instantaneous upon typing title/description/status in `tbl_tasks`. |
| **Live Calendar Sync** | Installable `onEdit` Trigger | Instantaneous upon creating/modifying events in `tbl_events`. |
| **Live Matrix Planning** | Simple `onEdit` Trigger | Instantaneous upon modifying cells in `view_weekly_matrix`. |
| **Tenure $\to$ Member Auto-Sync** | Simple `onEdit` Trigger | Instantaneous upon adding new tenures in `tbl_tenures`. |
| **Leave & Baseline Compilation** | Sidebar Submissions | Instantaneous upon submitting leaves or weekly baselines. |
| **Nightly Reconciliation Sync** | Daily Time-Driven Trigger (2:00 AM) | Automatically rolls forward presence logs, refreshes matrix & responsibles dashboards, pulls Google Tasks completions, syncs calendars, and audits backup conflicts. |

---

## 7. Upgrading Library Versions

When the engineering team releases a new version of `WorkspacePM`:
1. Open the team's spreadsheet Apps Script editor.
2. Click **Libraries** > **WorkspacePM**.
3. Select the new **Version** from the dropdown.
4. Click **Save**.
5. Return to the Google Sheet and click **Project Management** > **Initialize / Upgrade Spreadsheet** to apply any new schema columns or enhancements.
