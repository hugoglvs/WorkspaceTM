# ADR 0030: Canonical Action Forwarders and onOpen Menu Isolation in Consumer Sheets

## Status
Accepted

## Context
When running in Library distribution mode, team spreadsheets link to the `WorkspacePM` Google Apps Script Library via a lightweight container script (`ConsumerWrapper.gs`).

Upon clicking **Project Management** > **Presence & Leaves Sidebar** in the consumer Google Sheet, Google Apps Script triggered the following system error:
```text
Script function not found: showBaselineSidebar
```

### Root Cause Analysis
1. **Google Apps Script Menu Execution Context:**
   - In Google Sheets, custom menus created via `SpreadsheetApp.getUi().createMenu().addItem(caption, functionName)` bind `functionName` strictly to global functions declared in the **active container-bound project** (`ConsumerWrapper.gs`).
   - Apps Script custom menus do not support namespaced function references (e.g., `"WorkspacePM.showBaselineSidebar"`).
   - When `WorkspacePM.onOpen()` was invoked during the consumer sheet's lifecycle, the library script executed its own `ui.createMenu("Project Management").addItem(..., "showBaselineSidebar")`. Because both menus shared the same title, the library menu replaced the consumer menu in the Google Sheets UI.
2. **Missing Canonical Forwarder in Consumer Wrapper:**
   - In `ConsumerWrapper.gs`, the sidebar function was defined under the alias `clientShowSidebar()`, while the canonical name `showBaselineSidebar()` (and other canonical names like `init()`, `sync()`, `compilePresence()`) was not defined in the consumer script's global scope.
   - When the user clicked the menu item targeting `"showBaselineSidebar"`, Apps Script searched the consumer script, found no function named `showBaselineSidebar`, and aborted with `Script function not found: showBaselineSidebar`.

## Decision

1. **Direct Canonical Forwarders in `ConsumerWrapper.gs`:**
   - Declared top-level forwarder functions for all canonical library operations directly inside `ConsumerWrapper.gs`:
     - `showBaselineSidebar()`
     - `init()`, `setupOrUpgrade()`, `sync()`, `validateAndReport()`
     - `compilePresence()`, `compilePresenceMigration()`
     - `refreshWeeklyMatrix()`, `refreshActiveResponsibles()`, `refreshKpis()`
     - `checkBackupConflicts()`
     - `syncTasks()`, `syncCalendarEvents()`
     - `installLiveTaskSync()`, `cleanupForeignTasks()`
     - `debugUserIdentity()`
   - Maintained all existing `client...` aliases (`clientShowSidebar`, `clientSetupOrUpgrade`, etc.) for backwards compatibility.

2. **Menu Overwrite Prevention & Audit Isolation in Library (`src/Code.js`):**
   - Refactored `onOpen(e, skipMenu)` in `src/Code.js` to accept a `skipMenu` parameter.
   - Extracted `onOpenAudit(targetSpreadsheet)` into a dedicated public helper so consumer sheets can execute active user diagnostic checks on sheet open without re-registering or clobbering UI menus.
   - Updated `ConsumerWrapper.gs` `onOpen` to call `lib.onOpenAudit(ss)` (or `lib.onOpen(e, true)`).

## Consequences
- Clicking **Presence & Leaves Sidebar** (or any menu item or bound spreadsheet button) immediately executes the corresponding canonical forwarder and loads the sidebar via `safeInvokeLibrary("showBaselineSidebar")`.
- Consumer sheet menus cannot be accidentally replaced by library-internal menu declarations.
- Complete backwards compatibility is preserved for both legacy client aliases and canonical function names.
- All 124 unit tests pass (`node --test test/*.js`).
