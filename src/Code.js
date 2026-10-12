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

  // Readable tab names (ADR 0039): rename former technical tabs before tables are checked
  var renamedTabs = [];
  try {
    renamedTabs = renameLegacySheets(ss);
  } catch (errRename) {
    Logger.log("setupOrUpgrade rename note: " + errRename);
  }

  var res = freshInit(ss);
  res.renamedTabs = renamedTabs;

  // One-time restore of the reverted merged presence table (ADR 0036)
  try {
    res.presenceRestore = restorePresenceTablesFromRanges(ss);
  } catch (errRestore) {
    Logger.log("setupOrUpgrade presence restore note: " + errRestore);
  }

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

  // Remember each tab's ID so the team can rename tabs safely (ADR 0039)
  try {
    registerTableIds(ss);
  } catch (errIds) {
    Logger.log("setupOrUpgrade table IDs note: " + errIds);
  }

  // Members see Members, Weekly Presence, Projects, Tasks and Events (ADR 0036)
  try {
    res.layout = applyWorkspaceLayout(ss);
  } catch (errLayout) {
    Logger.log("setupOrUpgrade layout note: " + errLayout);
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
  if (ss && typeof checkProjectBackupConflicts === "function") {
    try { checkProjectBackupConflicts(ss); } catch (eR) {}
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
    if (typeof checkProjectBackupConflicts === "function") {
      try {
        checkProjectBackupConflicts(ss);
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
      if (typeof checkProjectBackupConflicts === "function") {
        try { checkProjectBackupConflicts(ss); } catch (eR) {}
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
 * Rebuilds the hidden fact_presence table that feeds the Looker Studio KPIs.
 * The single public refreshKpis: KpiEngine.js must not declare another one,
 * because Apps Script shares one global scope across files.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function refreshKpis(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  return renderPresenceFacts(ss);
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
    var r1 = TaskSyncService.installTaskSyncTrigger(ss);
    refreshEditTriggerFlag();
    return r1;
  }
  if (typeof installTaskSyncTrigger === "function") {
    var r2 = installTaskSyncTrigger(ss);
    refreshEditTriggerFlag();
    return r2;
  }
  return { success: false, message: "installTaskSyncTrigger not available" };
}

/**
 * Handler invoked by the installable onEdit trigger, with full OAuth
 * authorization. Once that trigger exists it handles every edit (ADR 0041), so
 * the simple onEdit steps aside and an edit runs once.
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {Object|null}
 */
function installedTaskEditHandler(e) {
  if (!e || !e.range) return null;
  return runWithEditLock(e, handleSheetEdit);
}

var EDIT_TRIGGER_FLAG = "EDIT_TRIGGER_ACTIVE";

function editTriggerProps() {
  try {
    return (typeof PropertiesService !== "undefined") ? PropertiesService.getDocumentProperties() : null;
  } catch (e) {
    return null;
  }
}

/** True when the installable onEdit trigger is known to exist (document property). */
function editTriggerActive() {
  var props = editTriggerProps();
  return !!(props && props.getProperty(EDIT_TRIGGER_FLAG) === "1");
}

/**
 * Records whether the installable onEdit trigger exists. Needs authorization, so
 * it runs from the installers and the nightly job, never from the simple onEdit.
 *
 * @return {boolean} Whether the trigger exists
 */
function refreshEditTriggerFlag() {
  var props = editTriggerProps();
  if (!props || typeof ScriptApp === "undefined" || !ScriptApp.getProjectTriggers) return false;
  var active = ScriptApp.getProjectTriggers().some(function(t) {
    var fn = t.getHandlerFunction();
    return fn === "installedTaskEditHandler" || fn === "installedConsumerTriggerHandler";
  });
  if (active) props.setProperty(EDIT_TRIGGER_FLAG, "1");
  else props.deleteProperty(EDIT_TRIGGER_FLAG);
  return active;
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
  var hasReminderTrigger = false;
  var hasBirthdayTrigger = false;

  for (var i = 0; i < existing.length; i++) {
    var fnName = existing[i].getHandlerFunction();
    if (fnName === "installedTaskEditHandler") hasEditTrigger = true;
    if (fnName === "nightlySyncJob") hasNightlyTrigger = true;
    if (fnName === "mondayPresenceReminder") hasReminderTrigger = true;
    if (fnName === "birthdayReminderJob") hasBirthdayTrigger = true;
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

  if (!hasReminderTrigger) {
    try {
      ScriptApp.newTrigger("mondayPresenceReminder")
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(8)
        .create();
      installed.push("mondayPresenceReminder (Mondays at 8:00 AM)");
    } catch (errRem) {
      Logger.log("installAllTriggers reminder trigger error: " + errRem);
    }
  } else {
    alreadyActive.push("mondayPresenceReminder (Mondays at 8:00 AM)");
  }

  if (!hasBirthdayTrigger) {
    try {
      ScriptApp.newTrigger("birthdayReminderJob")
        .timeBased()
        .everyDays(1)
        .atHour(9)
        .create();
      installed.push("birthdayReminderJob (daily at 9:00 AM)");
    } catch (errBd) {
      Logger.log("installAllTriggers birthday trigger error: " + errBd);
    }
  } else {
    alreadyActive.push("birthdayReminderJob (daily at 9:00 AM)");
  }

  try { refreshEditTriggerFlag(); } catch (eFlag) { Logger.log("Edit trigger flag note: " + eFlag); }

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
    try { refreshEditTriggerFlag(); } catch (eFlag) {}
    runNightlySyncSteps(ss, res);
  } finally {
    lock.releaseLock();
  }

  Logger.log("[nightlySyncJob] Completed: " + JSON.stringify(res));
  return res;
}

var NIGHTLY_TIME_BUDGET_MS = 300000; // Apps Script stops a run at 6 minutes

/**
 * The nightly maintenance steps. Each step is isolated so one failure does not
 * stop the others, timed in the log, and skipped once the run nears the
 * 6-minute limit (so a slow step never kills the whole run silently).
 * Must only run while holding the sync lock.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} res Status report to fill
 */
function runNightlySyncSteps(ss, res) {
  var started = Date.now();
  var steps = [
    ["presenceRestore", function() { return (typeof restorePresenceTablesFromRanges === "function") ? restorePresenceTablesFromRanges(ss) : null; }],
    ["directorySync", function() { return (typeof sync === "function") ? sync(ss) : null; }],
    ["presenceCompile", function() { return (typeof compilePresence === "function") ? compilePresence(ss) : null; }],
    ["weeklyMatrix", function() { return (typeof refreshWeeklyMatrix === "function") ? refreshWeeklyMatrix(ss) : null; }],
    ["kpis", function() { return (typeof refreshKpis === "function") ? refreshKpis(ss) : null; }],
    ["backupConflicts", function() { return (typeof checkProjectBackupConflicts === "function") ? checkProjectBackupConflicts(ss) : null; }],
    ["tasks", function() { return (typeof syncTasks === "function") ? syncTasks(ss) : null; }],
    ["calendar", function() { return (typeof syncCalendarEvents === "function") ? syncCalendarEvents(ss) : null; }],
    ["holidays", function() { return (typeof seedPublicHolidays === "function") ? seedPublicHolidays(ss) : null; }],
    ["birthdayCalendar", function() { return (typeof syncBirthdayCalendar === "function") ? syncBirthdayCalendar(ss) : null; }]
  ];
  res.timings = {};
  res.skipped = [];
  for (var i = 0; i < steps.length; i++) {
    var name = steps[i][0];
    if (Date.now() - started > NIGHTLY_TIME_BUDGET_MS) {
      res.skipped.push(name);
      continue;
    }
    var t0 = Date.now();
    try {
      res.steps[name] = steps[i][1]();
    } catch (err) {
      Logger.log("[nightlySyncJob] " + name + " note: " + err);
      res.steps[name] = { error: String(err) };
    }
    res.timings[name] = Date.now() - t0;
    Logger.log("[nightlySyncJob] " + name + ": " + res.timings[name] + " ms");
  }
  if (res.skipped.length) Logger.log("[nightlySyncJob] Out of time, skipped: " + res.skipped.join(", "));
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
        .addItem("Refresh KPI Data for Looker", "refreshKpis")
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
  // The installable trigger handles the edit when it exists (ADR 0041)
  if (editTriggerActive()) return { skipped: true, reason: "handled_by_installed_trigger" };
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
  var tableKey = tableKeyOfSheet(sheet);
  var editedRow = e.range.getRow();
  var editStart = Date.now();
  var editMarks = [];
  var mark = function(label) { editMarks.push(label + " " + (Date.now() - editStart) + "ms"); };

  // Handle in-sheet presence declaration in view_weekly_matrix
  if (tableKey === "weeklyMatrix") {
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
  if (tableKey === "tasks" && editedRow > 1) {
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
  if (tableKey === "events" && editedRow > 1) {
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
  if (editedRow > 1 && typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC[tableKey] && typeof validateRow === "function") {
    var numCols = Math.max(typeof sheet.getLastColumn === "function" ? (sheet.getLastColumn() || 0) : 0, SCHEMA_SPEC[tableKey].columns.length);
    var rowVals = sheet.getRange(editedRow, 1, 1, numCols).getValues()[0];
    var vRes = validateRow(tableKey, rowVals);
    if (vRes.populated && !vRes.valid) {
      try {
        var missingLabels = vRes.missingColumns.map(function(c) { return schemaColumnLabel(SCHEMA_SPEC[tableKey], c); });
        SpreadsheetApp.getActiveSpreadsheet().toast(
          "Missing required column(s): " + missingLabels.join(", ") + " in " + sheetName + " (Row " + editedRow + ").",
          "Required Field Missing",
          5
        );
      } catch (tErr) {}
    }
  }

  mark("validate");

  // Check if edited cell falls into an actively protected range
  if (typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC[tableKey]) {
    var tblConfig = SCHEMA_SPEC[tableKey];
    var editedCol = e.range.getColumn();
    var colName = (tblConfig.columns && editedCol <= tblConfig.columns.length) ? tblConfig.columns[editedCol - 1] : "";

    // Typing a new name in an empty Members row is the way to add someone: no
    // warning dialog there (ADR 0044)
    var addingMember = tableKey === "members" && colName === "member" &&
      (e.oldValue === undefined || e.oldValue === "") && e.value !== undefined && e.value !== "";
    var protInfo = (!addingMember && typeof isCellProtected === "function") ? isCellProtected(sheet, editedRow, editedCol) : null;

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

  mark("protections");

  // A name deleted from Members removes that member's row content (ADR 0044)
  if (tableKey === "members" && editedRow > 1 && e.range.getColumn() === 1 &&
      e.oldValue !== undefined && e.oldValue !== "" && (e.value === undefined || e.value === "") &&
      e.range.getNumRows() === 1 && e.range.getNumColumns() === 1 && SCHEMA_SPEC.members) {
    try {
      sheet.getRange(editedRow, 2, 1, SCHEMA_SPEC.members.columns.length - 1).clearContent();
    } catch (errClr) { Logger.log("onEdit member row clear note: " + errClr); }
  }

  // Birthday typed in Members: store it as "14 Mar" (ADR 0042)
  if (tableKey === "members" && typeof normalizeBirthdayEdit === "function") {
    try { normalizeBirthdayEdit(e); } catch (errBd) { Logger.log("onEdit birthday note: " + errBd); }
  }

  // A new name in Members without a contract: open the Members & Contracts sidebar (ADR 0037)
  if (tableKey === "members" && typeof promptContractForNewMember === "function") {
    try {
      promptContractForNewMember(e);
    } catch (errPrompt) {
      Logger.log("onEdit promptContractForNewMember note: " + errPrompt);
    }
  }

  if (tableKey === "members" || tableKey === "tenures") {
    try {
      var syncTargetSs = (sheet.getParent && typeof sheet.getParent === "function") ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet();
      var syncFn = (typeof syncMembersAndTenures === "function")
        ? syncMembersAndTenures
        : (typeof require !== "undefined" ? (function() { try { return require("./SyncEngine").syncMembersAndTenures; } catch(e) { return null; } })() : null);

      if (syncFn) {
        syncFn(syncTargetSs, { fromEdit: true });
      }
    } catch (err) {
      Logger.log("onEdit sync error: " + err);
    }
  }

  mark("members");

  // Locations tab edited: refresh the location dropdowns and Weekly Presence (ADR 0040)
  if (tableKey === "locations" && editedRow > 1 && typeof refreshLocationDropdowns === "function") {
    try {
      refreshLocationDropdowns(SpreadsheetApp.getActiveSpreadsheet() || sheet.getParent());
      if (typeof refreshWeeklyMatrix === "function") refreshWeeklyMatrix(sheet.getParent ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet());
      // Last, so its toast is not replaced by the refresh toast
      if (typeof warnRetiredLocation === "function") warnRetiredLocation(e);
    } catch (errLoc) {
      Logger.log("onEdit locations refresh note: " + errLoc);
    }
  }

  // A project without a backup: ask for one (ADR 0038)
  if (tableKey === "projects" && typeof promptProjectBackup === "function") {
    try {
      promptProjectBackup(e);
    } catch (errBk) {
      Logger.log("onEdit promptProjectBackup note: " + errBk);
    }
  }

  mark("locations+projects");

  // Real-time simultaneous absence conflict re-evaluation
  if (tableKey === "projects" || tableKey === "leaves" || tableKey === "backups") {
    if (typeof checkProjectBackupConflicts === "function") {
      try {
        checkProjectBackupConflicts(SpreadsheetApp.getActiveSpreadsheet());
      } catch (errConf) {
        Logger.log("onEdit checkProjectBackupConflicts error: " + errConf);
      }
    }
  }
  mark("coverage");
  if (Date.now() - editStart > 3000) Logger.log("[edit timing] " + (tableKey || sheetName) + ": " + editMarks.join(", "));
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
    refreshKpis: refreshKpis,
    syncTasks: syncTasks,
    syncCalendarEvents: syncCalendarEvents,
    cleanupForeignTasks: cleanupForeignTasks,
    checkBackupConflicts: checkBackupConflicts,
    installLiveTaskSync: installLiveTaskSync,
    installAllTriggers: installAllTriggers,
    nightlySyncJob: nightlySyncJob,
    installedTaskEditHandler: installedTaskEditHandler,
    editTriggerActive: editTriggerActive,
    refreshEditTriggerFlag: refreshEditTriggerFlag,
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
