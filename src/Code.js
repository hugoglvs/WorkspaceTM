/**
 * Project Management Workspace - Public Library API
 */

/**
 * Public initialization entry point.
 * Can be called directly from consumer sheets via: WorkspacePM.init() or WorkspacePM.setupOrUpgrade()
 */
function init(targetSpreadsheet) {
  return freshInit(targetSpreadsheet);
}

/**
 * Public library boundary for spreadsheet initialization or upgrade.
 * Can be called from consumer sheets via: WorkspacePM.setupOrUpgrade()
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function setupOrUpgrade(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var curVer = getSchemaVersion();
  var res = freshInit(ss);

  // Automatically install real-time edit trigger and nightly 2:00 AM sync trigger
  try {
    res.triggers = installAllTriggers(ss);
  } catch (errTrig) {
    Logger.log("setupOrUpgrade installAllTriggers note: " + errTrig);
  }

  // Run initial full synchronization across tables, dashboards, tasks, and calendar
  try {
    res.initialSync = nightlySyncJob(ss);
  } catch (errSync) {
    Logger.log("setupOrUpgrade initial sync note: " + errSync);
  }

  if (ss && typeof ss.toast === "function") {
    var title = curVer === 0 ? "Initialization Complete" : "Upgrade Complete";
    var msg = "Operational tables verified, live triggers active, and nightly 2:00 AM sync scheduled.";
    ss.toast(msg, title, 6);
  }
  return res;
}

// Sidebar functions (showBaselineSidebar, getSidebarInitialData, getMemberBaselines,
// saveMemberBaselines, submitMemberLeave) are declared in Sidebar.js to avoid global shadowing in Apps Script.

/**
 * Public library alias for synchronizing calendar events.
 * Can be called from consumer sheets via: WorkspacePM.syncEvents()
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function syncEvents(targetSpreadsheet) {
  return syncCalendarEvents(targetSpreadsheet);
}

/**
 * Returns current installed schema version from Document Properties.
 * @return {number}
 */
function getSchemaVersion() {
  var props = PropertiesService.getDocumentProperties();
  var ver = props.getProperty("SCHEMA_VERSION");
  return ver ? parseInt(ver, 10) : 0;
}

/**
 * Returns the library's latest schema version.
 * @return {number}
 */
function getLibraryVersion() {
  return CURRENT_SCHEMA_VERSION;
}

/**
 * Bi-directionally synchronizes Team Directory (tbl_members) and tbl_tenures.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object}
 */
function sync(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var syncFn = (typeof syncMembersAndTenures === "function")
    ? syncMembersAndTenures
    : (typeof require !== "undefined" ? (function() { try { return require("./SyncEngine").syncMembersAndTenures; } catch(e) { return null; } })() : null);
  var res = syncFn ? syncFn(ss) : null;
  if (ss && typeof refreshActiveResponsibles === "function") {
    try { refreshActiveResponsibles(ss); } catch (eR) {}
  }
  if (ss && typeof refreshKpis === "function") {
    try { refreshKpis(ss); } catch (eK) {}
  }
  if (ss && typeof syncTasks === "function") {
    try { syncTasks(ss); } catch (eT) {}
  }
  if (ss && typeof ss.toast === "function" && res) {
    var toastMsg = "Directory & Tenures synchronized. Added: " + (res.addedToDirectory || 0) + ", Updated: " + (res.updatedDirectoryRows || 0) + ".";
    ss.toast(toastMsg, "Sync Complete", 4);
  }
  return res;
}

/**
 * Audits all operational tables for missing required columns and reports findings.
 * @return {Object}
 */
function validateAndReport() {
  var res = validateAllTables();
  var ui = SpreadsheetApp.getUi();

  if (res.totalIssues === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "All required columns have valid values across all tables.",
      "Validation Passed",
      5
    );
  } else {
    var details = [];
    for (var t in res.tables) {
      if (!res.tables.hasOwnProperty(t)) continue;
      var tReport = res.tables[t];
      if (tReport.invalidRows && tReport.invalidRows.length > 0) {
        if (typeof highlightInvalidCells === "function") {
          try {
            highlightInvalidCells(t);
          } catch (eH) {}
        }
        for (var r = 0; r < tReport.invalidRows.length; r++) {
          var inv = tReport.invalidRows[r];
          details.push(t + " (Row " + inv.row + "): missing " + inv.missingColumns.join(", "));
        }
      }
    }
    ui.alert(
      "Validation Issues Found (" + res.totalIssues + ")",
      details.slice(0, 10).join("\n") + (details.length > 10 ? "\n...and " + (details.length - 10) + " more" : "") +
      "\n\n(Empty required cells have been highlighted in prominent alert red)",
      ui.ButtonSet.OK
    );
  }
  return res;
}

/**
 * Compiles presence baselines and approved leaves into daily records in tbl_presence_log.
 */
function compilePresence() {
  var res = (typeof compilePresenceLogs === "function") ? compilePresenceLogs() : { success: false, message: "compiler not loaded" };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (res.success && ss) {
    if (typeof refreshWeeklyMatrix === "function") {
      try {
        refreshWeeklyMatrix(ss);
      } catch (eMatrix) {}
    }
    if (typeof refreshActiveResponsibles === "function") {
      try {
        refreshActiveResponsibles(ss);
      } catch (eResp) {}
    }
    if (typeof refreshKpis === "function") {
      try {
        refreshKpis(ss);
      } catch (eKpi) {}
    }
    ss.toast(
      "Compiled " + res.createdCount + " daily presence logs (" + res.preservedCount + " existing preserved).",
      "Presence Compiled",
      5
    );
  }
  return res;
}

/**
 * Compiles presence logs from earliest/undefined start date up to today
 * for data migration and historical population.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Compilation summary report
 */
function compilePresenceMigration(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof PresenceEngine !== "undefined" && typeof PresenceEngine.compilePresenceMigration === "function") {
    return PresenceEngine.compilePresenceMigration(ss);
  }
  if (typeof compilePresenceLogs === "function") {
    var detected = (typeof findEarliestHistoricalDate === "function") ? findEarliestHistoricalDate(ss) : null;
    var startDateToUse = detected;
    var ui = null;
    try { ui = SpreadsheetApp.getUi(); } catch (eUi) {}

    if (ui && typeof ui.prompt === "function") {
      var dStr = detected ? (typeof formatDateIso === "function" ? formatDateIso(detected) : detected.toISOString().slice(0, 10)) : "None detected (defaults to Jan 1 of current year)";
      var pRes = ui.prompt(
        "Compile Presence Logs (Migration / Undefined Start Date)",
        "Auto-detected earliest recorded date: " + dStr + "\n\nEnter start date (YYYY-MM-DD) or click OK to compile from earliest date:",
        ui.ButtonSet.OK_CANCEL
      );
      if (pRes && typeof pRes.getSelectedButton === "function") {
        var cancelBtn = (ui.Button && ui.Button.CANCEL) ? ui.Button.CANCEL : "CANCEL";
        if (pRes.getSelectedButton() === cancelBtn) {
          if (ss && typeof ss.toast === "function") ss.toast("Presence migration cancelled.", "Migration Cancelled", 3);
          return { success: false, cancelled: true, message: "Migration cancelled by user" };
        }
        var uTxt = String(pRes.getResponseText ? pRes.getResponseText() : "").trim();
        if (uTxt) {
          var match = uTxt.match(/^(\d{4})-(\d{2})-(\d{2})/);
          var parsedDate = match ? new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10), 0, 0, 0, 0) : new Date(uTxt);
          if (!isNaN(parsedDate.getTime())) {
            parsedDate.setHours(0, 0, 0, 0);
            startDateToUse = parsedDate;
          }
        }
      }
    }

    if (!startDateToUse) {
      var now = new Date();
      startDateToUse = new Date(now.getFullYear(), 0, 1);
    }

    var res = compilePresenceLogs(startDateToUse, null, ss, null, { fromEarliest: true });
    if (res.success && ss) {
      if (typeof refreshWeeklyMatrix === "function") {
        try { refreshWeeklyMatrix(ss); } catch (eM) {}
      }
      if (typeof refreshActiveResponsibles === "function") {
        try { refreshActiveResponsibles(ss); } catch (eR) {}
      }
      if (typeof refreshKpis === "function") {
        try { refreshKpis(ss); } catch (eKp) {}
      }
      if (typeof ss.toast === "function") {
        var sStr = typeof formatDateIso === "function" ? formatDateIso(startDateToUse) : startDateToUse.toISOString().slice(0, 10);
        ss.toast(
          "Compiled " + res.createdCount + " migration daily presence logs from " + sStr + " (" + res.preservedCount + " existing preserved).",
          "Migration Completed",
          6
        );
      }
    }
    return res;
  }
  return { success: false, message: "compiler not loaded" };
}

/**
 * Refreshes the Weekly Matrix planning view.
 */
function refreshWeeklyMatrix(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof viewBuilder !== "undefined" && typeof viewBuilder.refreshWeeklyMatrix === "function") {
    return viewBuilder.refreshWeeklyMatrix(ss);
  }
  if (typeof renderWeeklyMatrixView === "function") {
    var res = renderWeeklyMatrixView(null, null, ss);
    if (res.success && ss && typeof ss.toast === "function") {
      try {
        ss.toast(
          "Refreshed Weekly Matrix view with " + res.rowsRendered + " weekdays across " + res.membersCount + " members.",
          "Weekly Matrix Updated",
          5
        );
      } catch (eToast) {}
    }
    return res;
  }
  return { success: false, message: "renderWeeklyMatrixView not available" };
}

/**
 * Refreshes the Active Responsibles Dashboard view.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function refreshActiveResponsibles(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof viewBuilder !== "undefined" && typeof viewBuilder.refreshActiveResponsibles === "function") {
    return viewBuilder.refreshActiveResponsibles(ss);
  }
  if (typeof renderActiveResponsiblesDashboard === "function") {
    var res = renderActiveResponsiblesDashboard(ss);
    if (res.success && ss && typeof ss.toast === "function") {
      try {
        ss.toast(
          "Refreshed Active Responsibles: " + res.leadActiveCount + " Lead Active, " + res.coveredByBackupCount + " Backup Covered, " + res.uncoveredCount + " Uncovered.",
          "Active Responsibles Updated",
          5
        );
      } catch (eToast) {}
    }
    return res;
  }
  return { success: false, message: "renderActiveResponsiblesDashboard not available" };
}

/**
 * Refreshes the Presence & Remote Work KPI Engine dashboard (calc_kpis).
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function refreshKpis(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof KpiEngine !== "undefined" && typeof KpiEngine.refreshKpis === "function") {
    return KpiEngine.refreshKpis(ss);
  }
  if (typeof renderKpiDashboard === "function") {
    var res = renderKpiDashboard(ss);
    if (res.success && ss && typeof ss.toast === "function") {
      try {
        ss.toast(
          "Refreshed Presence & Remote KPIs: " + res.memberRowsCount + " member metrics across " + res.monthsCount + " months.",
          "KPIs Updated",
          5
        );
      } catch (eToast) {}
    }
    return res;
  }
  return { success: false, message: "renderKpiDashboard not available" };
}

/**
 * Synchronizes tbl_tasks bi-directionally with Google Tasks.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function syncTasks(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof TaskSyncService !== "undefined" && typeof TaskSyncService.syncTasks === "function") {
    return TaskSyncService.syncTasks(ss);
  }
  if (typeof syncGoogleTasks === "function") {
    var res = syncGoogleTasks(ss);
    if (ss && typeof ss.toast === "function") {
      if (res.success) {
        ss.toast(
          "Google Tasks Synced: " + res.pushedCount + " new pushed, " + res.pulledUpdatesCount + " status updates.",
          "Tasks Synchronized",
          5
        );
      } else {
        ss.toast(res.message || "Failed to synchronize tasks", "Task Sync Error", 5);
      }
    }
    return res;
  }
  return { success: false, message: "TaskSyncService not loaded" };
}

/**
 * Unlinks and deletes tasks belonging to other members from active user's Google Tasks.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Report
 */
function cleanupForeignTasks(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof TaskSyncService !== "undefined" && typeof TaskSyncService.cleanupMisassignedGoogleTasks === "function") {
    return TaskSyncService.cleanupMisassignedGoogleTasks(ss);
  }
  if (typeof cleanupMisassignedGoogleTasks === "function") {
    return cleanupMisassignedGoogleTasks(ss);
  }
  return { success: false, message: "cleanupMisassignedGoogleTasks not available" };
}

/**
 * Evaluates active projects for lead and backup simultaneous absence conflicts.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object}
 */
function checkBackupConflicts(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof checkProjectBackupConflicts === "function") {
    var res = checkProjectBackupConflicts(ss);
    if (ss && typeof ss.toast === "function") {
      var msg = res.conflictsFound > 0
        ? "⚠️ " + res.conflictsFound + " project backup conflict(s) detected! Rows highlighted in red."
        : "✅ All active projects have continuous coverage. No backup conflicts.";
      ss.toast(msg, "Project Backup Audit", 5);
    }
    return res;
  }
  return { success: false, message: "checkProjectBackupConflicts not loaded" };
}

/**
 * Synchronizes tbl_events with Google Calendar.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function syncCalendarEvents(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof CalendarSyncService !== "undefined" && typeof CalendarSyncService.syncCalendar === "function") {
    return CalendarSyncService.syncCalendar(ss);
  }
  if (typeof syncCalendar === "function") {
    return syncCalendar(ss);
  }
  return { success: false, message: "CalendarSyncService not loaded" };
}

/**
 * Installs an installable onEdit trigger for live real-time task synchronization.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function installLiveTaskSync(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof TaskSyncService !== "undefined" && typeof TaskSyncService.installTaskSyncTrigger === "function") {
    return TaskSyncService.installTaskSyncTrigger(ss);
  }
  if (typeof installTaskSyncTrigger === "function") {
    return installTaskSyncTrigger(ss);
  }
  return { success: false, message: "installTaskSyncTrigger not available" };
}

/**
 * Handler invoked by installable onEdit trigger with full OAuth authorization.
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {Object|null}
 */
function installedTaskEditHandler(e) {
  if (!e || !e.range) return null;
  return runWithEditLock(e, handleInstalledEdit);
}

/**
 * Routes an installable onEdit event to the Tasks / Calendar sync handlers.
 * Must only run while holding the sync lock.
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {Object|null}
 */
function handleInstalledEdit(e) {
  var sheet = e.range.getSheet();
  if (!sheet) return null;
  var sheetName = sheet.getName();
  if (sheetName === "tbl_tasks") {
    if (typeof TaskSyncService !== "undefined" && typeof TaskSyncService.handleTaskEdit === "function") {
      return TaskSyncService.handleTaskEdit(e);
    }
    if (typeof handleTaskEdit === "function") {
      return handleTaskEdit(e);
    }
  } else if (sheetName === "tbl_events") {
    if (typeof CalendarSyncService !== "undefined" && typeof CalendarSyncService.handleEventEdit === "function") {
      return CalendarSyncService.handleEventEdit(e);
    }
    if (typeof handleEventEdit === "function") {
      return handleEventEdit(e);
    }
  }
  return null;
}

/**
 * Installs all automated background triggers:
 * 1. Installable onEdit trigger ('installedTaskEditHandler') for real-time task, event, and matrix sync with OAuth.
 * 2. Daily time-driven trigger ('nightlySyncJob') running at 2:00 AM for automated presence compilation, dashboard refresh, and external tasks/calendar sync.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function installAllTriggers(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof ScriptApp === "undefined" || !ScriptApp.getProjectTriggers) {
    return { success: false, message: "ScriptApp not available" };
  }

  var existing = ScriptApp.getProjectTriggers();
  var hasEditTrigger = false;
  var hasNightlyTrigger = false;

  for (var i = 0; i < existing.length; i++) {
    var fnName = existing[i].getHandlerFunction();
    if (fnName === "installedTaskEditHandler") hasEditTrigger = true;
    if (fnName === "nightlySyncJob") hasNightlyTrigger = true;
  }

  var installed = [];
  var alreadyActive = [];

  if (!hasEditTrigger) {
    try {
      ScriptApp.newTrigger("installedTaskEditHandler")
        .forSpreadsheet(ss)
        .onEdit()
        .create();
      installed.push("installedTaskEditHandler (onEdit)");
    } catch (errEdit) {
      Logger.log("installAllTriggers edit trigger error: " + errEdit);
    }
  } else {
    alreadyActive.push("installedTaskEditHandler (onEdit)");
  }

  if (!hasNightlyTrigger) {
    try {
      ScriptApp.newTrigger("nightlySyncJob")
        .timeBased()
        .everyDays(1)
        .atHour(2)
        .create();
      installed.push("nightlySyncJob (daily at 2:00 AM)");
    } catch (errNightly) {
      Logger.log("installAllTriggers nightly trigger error: " + errNightly);
    }
  } else {
    alreadyActive.push("nightlySyncJob (daily at 2:00 AM)");
  }

  return {
    success: true,
    installed: installed,
    alreadyActive: alreadyActive
  };
}

var NIGHTLY_SYNC_LOCK_TIMEOUT_MS = 30000;
var EDIT_LOCK_TIMEOUT_MS = 10000; // simple onEdit triggers are capped at 30 s

/**
 * Acquires the lock that serialises sync writes on this spreadsheet.
 * Uses the document lock so each Consumer Sheet has its own lock: a script
 * lock created in a library is shared by every including script, so it would
 * make all teams' nightly jobs queue behind each other. Falls back to the
 * script lock when there is no containing document.
 *
 * @param {number} timeoutMs How long to wait for the lock
 * @return {GoogleAppsScript.Lock.Lock|null} The held lock, or null on timeout
 */
function acquireSyncLock(timeoutMs) {
  var lock = LockService.getDocumentLock() || LockService.getScriptLock();
  return lock.tryLock(timeoutMs) ? lock : null;
}

/**
 * Runs an edit handler under the sync lock. While the nightly sync holds the
 * lock, the user's typed value stays in the cell but the script-side follow-up
 * (IDs, Google Tasks / Calendar push, presence log writes) is skipped, and the
 * user is told so. The nightly sync picks up pending tasks and events.
 *
 * @param {Object} e Apps Script onEdit event object
 * @param {function(Object):*} handler
 * @return {*} The handler result, or { skipped: true, reason: "sync_in_progress" }
 */
function runWithEditLock(e, handler) {
  var lock = acquireSyncLock(EDIT_LOCK_TIMEOUT_MS);
  if (!lock) {
    try {
      var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
      ss.toast(
        "The nightly sync is running. Your change is saved, but its automatic updates were skipped. Re-enter it in a few minutes if it does not appear in Google Tasks / Calendar or the presence log.",
        "Sync in progress",
        10
      );
    } catch (tErr) {}
    return { skipped: true, reason: "sync_in_progress" };
  }
  try {
    return handler(e);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Automated nightly maintenance sync job:
 * Runs daily at 2:00 AM via time-driven trigger, or on-demand during setup/upgrade.
 * Reconciles presence logs, refreshes matrix & responsibles dashboards, pulls Google Tasks completions,
 * syncs calendar events, audits project backup conflicts, and syncs directory tenures.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function nightlySyncJob(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  Logger.log("[nightlySyncJob] Starting automated nightly maintenance job...");
  var res = {
    timestamp: new Date().toISOString(),
    success: true,
    steps: {}
  };

  var lock = acquireSyncLock(NIGHTLY_SYNC_LOCK_TIMEOUT_MS);
  if (!lock) {
    Logger.log("[nightlySyncJob] Another sync holds the spreadsheet lock; skipping this run.");
    res.success = false;
    res.reason = "lock_timeout";
    return res;
  }
  try {
    runNightlySyncSteps(ss, res);
  } finally {
    lock.releaseLock();
  }

  Logger.log("[nightlySyncJob] Completed: " + JSON.stringify(res));
  return res;
}

/**
 * The six nightly maintenance steps. Each step is isolated so one failure
 * does not stop the others. Must only run while holding the sync lock.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} res Status report to fill
 */
function runNightlySyncSteps(ss, res) {

  // 1. Sync Directory & Tenures
  try {
    if (typeof sync === "function") {
      res.steps.directorySync = sync(ss);
    }
  } catch (e1) {
    Logger.log("[nightlySyncJob] Directory sync note: " + e1);
    res.steps.directorySync = { error: String(e1) };
  }

  // 2. Compile Daily Presence Logs
  try {
    if (typeof compilePresence === "function") {
      res.steps.presenceCompile = compilePresence(ss);
    }
  } catch (e2) {
    Logger.log("[nightlySyncJob] Compile presence note: " + e2);
    res.steps.presenceCompile = { error: String(e2) };
  }

  // 3. Refresh Dashboards (Matrix, Responsibles, KPIs)
  try {
    if (typeof refreshWeeklyMatrix === "function") res.steps.weeklyMatrix = refreshWeeklyMatrix(ss);
    if (typeof refreshActiveResponsibles === "function") res.steps.activeResponsibles = refreshActiveResponsibles(ss);
    if (typeof refreshKpis === "function") res.steps.kpis = refreshKpis(ss);
  } catch (e3) {
    Logger.log("[nightlySyncJob] Dashboard refresh note: " + e3);
    res.steps.dashboards = { error: String(e3) };
  }

  // 4. Synchronize Google Tasks (pull completed status, push any pending)
  try {
    if (typeof syncTasks === "function") {
      res.steps.tasks = syncTasks(ss);
    }
  } catch (e4) {
    Logger.log("[nightlySyncJob] Tasks sync note: " + e4);
    res.steps.tasks = { error: String(e4) };
  }

  // 5. Synchronize Google Calendar Events
  try {
    if (typeof syncCalendarEvents === "function") {
      res.steps.calendar = syncCalendarEvents(ss);
    }
  } catch (e5) {
    Logger.log("[nightlySyncJob] Calendar sync note: " + e5);
    res.steps.calendar = { error: String(e5) };
  }

  // 6. Check Project Backup Absence Conflicts
  try {
    if (typeof checkProjectBackupConflicts === "function") {
      res.steps.backupConflicts = checkProjectBackupConflicts(ss);
    }
  } catch (e6) {
    Logger.log("[nightlySyncJob] Backup conflicts check note: " + e6);
    res.steps.backupConflicts = { error: String(e6) };
  }
}

/**
 * Adds custom menu items to the Google Sheets UI.
 *
 * @param {Object} [e]
 * @param {boolean} [skipMenu=false]
 */
function onOpen(e, skipMenu) {
  if (!skipMenu) {
    try {
      var ui = SpreadsheetApp.getUi();
      ui.createMenu("Project Management")
        .addItem("Initialize / Upgrade Tables", "init")
        .addItem("Sync Directory & Tenures", "sync")
        .addItem("Validate Required Fields", "validateAndReport")
        .addItem("Compile Daily Presence Logs", "compilePresence")
        .addItem("Compile Presence Logs (Migration / Undefined Start Date)", "compilePresenceMigration")
        .addItem("Refresh Weekly Matrix View", "refreshWeeklyMatrix")
        .addItem("Refresh Active Responsibles Dashboard", "refreshActiveResponsibles")
        .addItem("Refresh Presence & Remote KPIs (calc_kpis)", "refreshKpis")
        .addItem("Check Project Backup Conflicts", "checkBackupConflicts")
        .addItem("Sync Google Tasks", "syncTasks")
        .addItem("Sync Calendar Events", "syncCalendarEvents")
        .addItem("Enable Live Task Sync (Install Trigger)", "installLiveTaskSync")
        .addItem("Clean Up Foreign Google Tasks", "cleanupForeignTasks")
        .addSeparator()
        .addItem("Presence & Leaves Sidebar", "showBaselineSidebar")
        .addItem("Debug User & Member Identity", "debugUserIdentity")
        .addToUi();
    } catch (eUi) {
      Logger.log("[onOpen] Menu creation note: " + eUi);
    }
  }

  // Audit active user status and log identity diagnostics onOpen
  onOpenAudit();
}

/**
 * Runs user identity diagnostics onOpen without modifying UI menus.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object|null}
 */
function onOpenAudit(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  try {
    if (typeof checkActiveUserStatus === "function") {
      return checkActiveUserStatus(ss, true);
    }
  } catch (eAudit) {
    Logger.log("[onOpenAudit] Error in checkActiveUserStatus: " + eAudit);
  }
  return null;
}

/**
 * Runs active user and tbl_members diagnostic audit and displays popup alert in Sheets.
 */
function debugUserIdentity() {
  if (typeof auditUserAndMembersIdentity === "function") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var audit = auditUserAndMembersIdentity(ss);
    try {
      var ui = SpreadsheetApp.getUi();
      if (ui) {
        var summary = audit.logLines.join("\n");
        if (summary.length > 1800) {
          summary = summary.substring(0, 1800) + "\n\n... (truncated, check Extensions > Apps Script > Executions for full logs)";
        }
        ui.alert("User & Member Identity Diagnostics", summary, ui.ButtonSet.OK);
      }
    } catch (eUi) {}
    return audit;
  }
}

/**
 * Automatically triggers validation warnings and synchronization on edit.
 */
function onEdit(e) {
  if (!e || !e.range) return null;
  return runWithEditLock(e, handleSheetEdit);
}

/**
 * Script-side processing of a user edit (IDs, validation, protections,
 * directory sync, presence declarations, backup conflicts).
 * Must only run while holding the sync lock.
 *
 * @param {Object} e Apps Script onEdit event object
 */
function handleSheetEdit(e) {
  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  var editedRow = e.range.getRow();

  // Handle in-sheet presence declaration in view_weekly_matrix
  if (sheetName === "view_weekly_matrix") {
    if (typeof handleWeeklyMatrixEdit === "function") {
      try {
        handleWeeklyMatrixEdit(e);
      } catch (errM) {
        Logger.log("handleWeeklyMatrixEdit error: " + errM);
      }
    }
    return;
  }

  // Handle task row creation / editing in tbl_tasks immediately so task_id, status, and formatting are applied first!
  if (sheetName === "tbl_tasks" && editedRow > 1) {
    var taskEditHandler = (typeof handleTaskEdit === "function")
      ? handleTaskEdit
      : (typeof TaskSyncService !== "undefined" && typeof TaskSyncService.handleTaskEdit === "function")
        ? TaskSyncService.handleTaskEdit
        : null;
    if (taskEditHandler) {
      try {
        taskEditHandler(e);
      } catch (errT) {
        Logger.log("onEdit handleTaskEdit error: " + errT);
      }
    }
  }

  // Handle event row creation / editing in tbl_events immediately so event_id, calendar sync, and formatting are applied
  if (sheetName === "tbl_events" && editedRow > 1) {
    var eventEditHandler = (typeof handleEventEdit === "function")
      ? handleEventEdit
      : (typeof CalendarSyncService !== "undefined" && typeof CalendarSyncService.handleEventEdit === "function")
        ? CalendarSyncService.handleEventEdit
        : null;
    if (eventEditHandler) {
      try {
        eventEditHandler(e);
      } catch (errE) {
        Logger.log("onEdit handleEventEdit error: " + errE);
      }
    }
  }

  // Validate edited row if within an operational table
  if (editedRow > 1 && typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC[sheetName] && typeof validateRow === "function") {
    var numCols = Math.max(typeof sheet.getLastColumn === "function" ? (sheet.getLastColumn() || 0) : 0, SCHEMA_SPEC[sheetName].columns.length);
    var rowVals = sheet.getRange(editedRow, 1, 1, numCols).getValues()[0];
    var vRes = validateRow(sheetName, rowVals);
    if (vRes.populated && !vRes.valid) {
      try {
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "Missing required column(s): " + vRes.missingColumns.join(", ") + " in " + sheetName + " (Row " + editedRow + ").",
          "Required Field Missing",
          5
        );
      } catch (tErr) {}
    }
  }

  // Check if edited cell falls into an actively protected range
  if (typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC[sheetName]) {
    var tblConfig = SCHEMA_SPEC[sheetName];
    var editedCol = e.range.getColumn();
    var colName = (tblConfig.columns && editedCol <= tblConfig.columns.length) ? tblConfig.columns[editedCol - 1] : "";

    var protInfo = (typeof isCellProtected === "function") ? isCellProtected(sheet, editedRow, editedCol) : null;

    if (protInfo && protInfo.isProtected) {
      var protHandler = (typeof handleProtectedCellEdit === "function")
        ? handleProtectedCellEdit
        : (typeof require !== "undefined" ? (function() { try { return require("./ProtectionManager").handleProtectedCellEdit; } catch(e) { return null; } })() : null);

      if (protHandler) {
        try {
          protHandler(sheet, editedRow, editedCol, colName, protInfo);
        } catch (hProtErr) {
          Logger.log("handleProtectedCellEdit note: " + hProtErr);
        }
      }
    }
  }

  if (sheetName === "tbl_members" || sheetName === "tbl_tenures") {
    try {
      var syncTargetSs = (sheet.getParent && typeof sheet.getParent === "function") ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet();
      var syncFn = (typeof syncMembersAndTenures === "function")
        ? syncMembersAndTenures
        : (typeof require !== "undefined" ? (function() { try { return require("./SyncEngine").syncMembersAndTenures; } catch(e) { return null; } })() : null);

      if (syncFn) {
        syncFn(syncTargetSs);
      }
    } catch (err) {
      Logger.log("onEdit sync error: " + err);
    }
  }

  // Real-time simultaneous absence conflict re-evaluation
  if (sheetName === "tbl_projects" || sheetName === "tbl_leaves" || sheetName === "tbl_backups") {
    if (typeof checkProjectBackupConflicts === "function") {
      try {
        checkProjectBackupConflicts(SpreadsheetApp.getActiveSpreadsheet());
      } catch (errConf) {
        Logger.log("onEdit checkProjectBackupConflicts error: " + errConf);
      }
    }
  }
}

if (typeof module !== 'undefined' && module.exports) {
  var _SidebarMod = typeof require !== 'undefined' ? (function() { try { return require('./Sidebar'); } catch(e) { return null; } })() : null;
  var _ProtMod = typeof require !== 'undefined' ? (function() { try { return require('./ProtectionManager'); } catch(e) { return null; } })() : null;
  module.exports = {
    init: init,
    setupOrUpgrade: setupOrUpgrade,
    sync: sync,
    syncEvents: syncEvents,
    validateAndReport: validateAndReport,
    getSchemaVersion: getSchemaVersion,
    getLibraryVersion: getLibraryVersion,
    onOpen: onOpen,
    onOpenAudit: onOpenAudit,
    onEdit: onEdit,
    compilePresence: compilePresence,
    compilePresenceMigration: compilePresenceMigration,
    debugUserIdentity: debugUserIdentity,
    refreshWeeklyMatrix: refreshWeeklyMatrix,
    refreshActiveResponsibles: refreshActiveResponsibles,
    refreshKpis: refreshKpis,
    syncTasks: syncTasks,
    syncCalendarEvents: syncCalendarEvents,
    cleanupForeignTasks: cleanupForeignTasks,
    checkBackupConflicts: checkBackupConflicts,
    installLiveTaskSync: installLiveTaskSync,
    installAllTriggers: installAllTriggers,
    nightlySyncJob: nightlySyncJob,
    installedTaskEditHandler: installedTaskEditHandler,
    showBaselineSidebar: function() {
      if (typeof showBaselineSidebar === "function") return showBaselineSidebar();
      if (_SidebarMod && typeof _SidebarMod.showBaselineSidebar === "function") return _SidebarMod.showBaselineSidebar();
    },
    getSidebarInitialData: function() {
      if (typeof getSidebarInitialData === "function") return getSidebarInitialData();
      if (_SidebarMod && typeof _SidebarMod.getSidebarInitialData === "function") return _SidebarMod.getSidebarInitialData();
      return {};
    },
    getMemberBaselines: function(m) {
      if (typeof getMemberBaselines === "function") return getMemberBaselines(m);
      if (_SidebarMod && typeof _SidebarMod.getMemberBaselines === "function") return _SidebarMod.getMemberBaselines(m);
      return {};
    },
    saveMemberBaselines: function(p) {
      if (typeof saveMemberBaselines === "function") return saveMemberBaselines(p);
      if (_SidebarMod && typeof _SidebarMod.saveMemberBaselines === "function") return _SidebarMod.saveMemberBaselines(p);
      return { success: false };
    },
    submitMemberLeave: function(p, t) {
      if (typeof submitMemberLeave === "function") return submitMemberLeave(p, t);
      if (_SidebarMod && typeof _SidebarMod.submitMemberLeave === "function") return _SidebarMod.submitMemberLeave(p, t);
      return { success: false };
    },
    isProtectionWarningSnoozed: function() {
      if (typeof isProtectionWarningSnoozed === "function") return isProtectionWarningSnoozed();
      if (_ProtMod && typeof _ProtMod.isProtectionWarningSnoozed === "function") return _ProtMod.isProtectionWarningSnoozed();
      return false;
    },
    snoozeProtectionWarnings: function(d) {
      if (typeof snoozeProtectionWarnings === "function") return snoozeProtectionWarnings(d);
      if (_ProtMod && typeof _ProtMod.snoozeProtectionWarnings === "function") return _ProtMod.snoozeProtectionWarnings(d);
    },
    clearProtectionWarningSnooze: function() {
      if (typeof clearProtectionWarningSnooze === "function") return clearProtectionWarningSnooze();
      if (_ProtMod && typeof _ProtMod.clearProtectionWarningSnooze === "function") return _ProtMod.clearProtectionWarningSnooze();
    },
    handleProtectedCellEdit: function(s, r, c, n, p) {
      if (typeof handleProtectedCellEdit === "function") return handleProtectedCellEdit(s, r, c, n, p);
      if (_ProtMod && typeof _ProtMod.handleProtectedCellEdit === "function") return _ProtMod.handleProtectedCellEdit(s, r, c, n, p);
    }
  };
}
