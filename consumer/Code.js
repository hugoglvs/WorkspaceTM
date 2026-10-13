/**
 * Google Workspace Project Management - Consumer Sheet Client Wrapper
 * 
 * Instructions for Team Leads:
 * 1. Open your Google Sheet, navigate to Extensions > Apps Script.
 * 2. In the Apps Script editor, click '+' next to 'Libraries'.
 * 3. Enter the Project Management Library Script ID and click 'Look up'.
 * 4. Set the Identifier to 'WorkspacePM' and select the latest version.
 * 5. Paste the contents of this file into your project's Code.gs.
 * 6. Reload the spreadsheet or run onOpen() to activate the 'Project Management' menu.
 */

var LIBRARY_ID = "WorkspacePM";

/**
 * Creates the Project Management menu on spreadsheet open.
 * Clean, minimal interface:
 * 1. Initialize / Upgrade Spreadsheet (automatically configures tables & installs all background triggers)
 * 2. Declare Presence & Leaves (opens the presence & leaves sidebar)
 *
 * @param {Object} [e]
 */
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Project Management")
    .addItem("Initialize / Upgrade Spreadsheet", "setupOrUpgrade")
    .addSeparator()
    .addItem("Declare Presence & Leaves", "showBaselineSidebar")
    .addItem("Declare Contract", "showContractsSidebar")
    .addItem("Interim Backups", "showBackupSidebar")
    .addItem("Assign Tasks", "showAssignSidebar")
    .addItem("Calendar", "showTaskCalendar")
    .addItem("Connect my Google Tasks", "connectMyGoogleTasks")
    .addSeparator()
    .addItem("Check Workspace", "showWorkspaceCheck")
    .addItem("Edit Locations", "showLocationsTab")
    .addItem("Edit Holidays", "showHolidaysTab")
    .addItem("Edit Task Types", "showTaskTypesTab")
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
 *
 * @param {Object} e
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
 *
 * @return {Object|null}
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
 * Safely invokes a library method with user-friendly error handling and toast notifications.
 *
 * @param {string} methodName
 * @param {Array} [args]
 * @return {*}
 */
function safeInvokeLibrary(methodName, args) {
  var lib = getLibraryInstance();
  if (!lib) {
    try {
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        "WorkspacePM Library Not Found",
        "The 'WorkspacePM' library is not linked to this Google Sheet.\n\n" +
        "To link the library:\n" +
        "1. In Apps Script, click '+' next to 'Libraries' in the left navigation.\n" +
        "2. Paste the Library Script ID and click 'Look up'.\n" +
        "3. Set the Identifier to 'WorkspacePM'.\n" +
        "4. Select the latest version and click 'Add'.",
        ui.ButtonSet.OK
      );
    } catch (eUi) {
      Logger.log("Library not found alert error: " + eUi);
    }
    return { success: false, error: "LIBRARY_NOT_LINKED" };
  }

  if (typeof lib[methodName] !== "function") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss && typeof ss.toast === "function") {
      ss.toast(
        "Action '" + methodName + "' is not supported in the currently linked WorkspacePM library version.",
        "Unsupported Library Action",
        6
      );
    }
    return { success: false, error: "METHOD_NOT_SUPPORTED", method: methodName };
  }

  try {
    return lib[methodName].apply(lib, args || []);
  } catch (err) {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (activeSs && typeof activeSs.toast === "function") {
      activeSs.toast("Error executing " + methodName + ": " + (err.message || err), "Action Error", 6);
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
function showBaselineSidebar() {
  return safeInvokeLibrary("showBaselineSidebar");
}

function init() {
  return safeInvokeLibrary("init");
}

function setupOrUpgrade() {
  return safeInvokeLibrary("setupOrUpgrade");
}

function sync() {
  return safeInvokeLibrary("sync");
}

function validateAndReport() {
  return safeInvokeLibrary("validateAndReport");
}

function compilePresence() {
  return safeInvokeLibrary("compilePresence");
}

function compilePresenceMigration() {
  return safeInvokeLibrary("compilePresenceMigration");
}

function refreshWeeklyMatrix() {
  return safeInvokeLibrary("refreshWeeklyMatrix");
}

function refreshKpis() {
  return safeInvokeLibrary("refreshKpis");
}

function checkBackupConflicts() {
  return safeInvokeLibrary("checkBackupConflicts");
}

function syncTasks() {
  return safeInvokeLibrary("syncTasks");
}

function syncCalendarEvents() {
  return safeInvokeLibrary("syncCalendarEvents");
}

function installLiveTaskSync() {
  return safeInvokeLibrary("installLiveTaskSync");
}

function installAllTriggers() {
  return safeInvokeLibrary("installAllTriggers");
}

function nightlySyncJob() {
  return safeInvokeLibrary("nightlySyncJob");
}

function cleanupForeignTasks() {
  return safeInvokeLibrary("cleanupForeignTasks");
}

function debugUserIdentity() {
  return safeInvokeLibrary("debugUserIdentity");
}

/**
 * Client Aliases (for backwards compatibility)
 */
function clientShowSidebar() {
  return showBaselineSidebar();
}

function clientSetupOrUpgrade() {
  return setupOrUpgrade();
}

function clientSyncMembersAndTenures() {
  return sync();
}

function clientCompilePresence() {
  return compilePresence();
}

function clientRefreshDashboards() {
  var r1 = refreshWeeklyMatrix();
  var r2 = checkBackupConflicts();
  var r3 = refreshKpis();
  return { weeklyMatrix: r1, backupCoverage: r2, kpis: r3 };
}

function clientSyncTasks() {
  return syncTasks();
}

function clientSyncCalendarEvents() {
  return syncCalendarEvents();
}

function clientInstallTriggers() {
  return installAllTriggers();
}

/**
 * Handler invoked by installable onEdit trigger installed in the consumer spreadsheet.
 */
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

/**
 * Sidebar API forwarders for google.script.run calls from BaselineSidebar.html
 */
function showLocationsTab() {
  return safeInvokeLibrary("showLocationsTab");
}

function showHolidaysTab() {
  return safeInvokeLibrary("showHolidaysTab");
}

function showContractsSidebar() {
  return safeInvokeLibrary("showContractsSidebar");
}

function getContractsSidebarData() {
  return safeInvokeLibrary("getContractsSidebarData");
}

function saveContract(payload) {
  return safeInvokeLibrary("saveContract", [payload]);
}

function restoreMemberFormulas() {
  return safeInvokeLibrary("restoreMemberFormulas");
}

function showBackupSidebar() {
  return safeInvokeLibrary("showBackupSidebar");
}

function getBackupSidebarData() {
  return safeInvokeLibrary("getBackupSidebarData");
}

function saveInterimBackup(payload) {
  return safeInvokeLibrary("saveInterimBackup", [payload]);
}

function cancelInterimBackup(backupId) {
  return safeInvokeLibrary("cancelInterimBackup", [backupId]);
}

function showWorkspaceCheck() {
  return safeInvokeLibrary("showWorkspaceCheck");
}

function runWorkspaceCheck() {
  return safeInvokeLibrary("runWorkspaceCheck");
}

function removePresenceDuplicates() {
  return safeInvokeLibrary("removePresenceDuplicates");
}

/** Time trigger: Monday presence reminder (installed by Initialize / Upgrade). */
function mondayPresenceReminder() {
  return safeInvokeLibrary("mondayPresenceReminder");
}

/** Time trigger: birthday reminders and the Team Birthdays calendar (daily, 9 am). */
function birthdayReminderJob() {
  return safeInvokeLibrary("birthdayReminderJob");
}

/** Time trigger: birthdays, Team Birthdays calendar and task reminders (daily, 9 am). */
function dailyRemindersJob() {
  return safeInvokeLibrary("dailyRemindersJob");
}

function connectMyGoogleTasks() {
  return safeInvokeLibrary("connectMyGoogleTasks");
}

/** Time trigger (one per connected person, hourly): their Google Tasks sync. */
function myTasksSyncJob() {
  return safeInvokeLibrary("myTasksSyncJob");
}

function showAssignSidebar() {
  return safeInvokeLibrary("showAssignSidebar");
}

function showTaskCalendar() {
  return safeInvokeLibrary("showTaskCalendar");
}

function showTaskTypesTab() {
  return safeInvokeLibrary("showTaskTypesTab");
}

function getAssignSidebarData() {
  return safeInvokeLibrary("getAssignSidebarData");
}

function assignTasks(payload) {
  return safeInvokeLibrary("assignTasks", [payload]);
}

function stopRecurringRule(ruleId) {
  return safeInvokeLibrary("stopRecurringRule", [ruleId]);
}

function getSidebarInitialData() {
  return safeInvokeLibrary("getSidebarInitialData");
}

function getMemberBaselines(member) {
  return safeInvokeLibrary("getMemberBaselines", [member]);
}

function saveMemberBaselines(payload) {
  return safeInvokeLibrary("saveMemberBaselines", [payload]);
}

function submitMemberLeave(payload) {
  return safeInvokeLibrary("submitMemberLeave", [payload]);
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    onOpen: onOpen,
    onEdit: onEdit,
    getLibraryInstance: getLibraryInstance,
    safeInvokeLibrary: safeInvokeLibrary,
    showBaselineSidebar: showBaselineSidebar,
    init: init,
    setupOrUpgrade: setupOrUpgrade,
    sync: sync,
    validateAndReport: validateAndReport,
    compilePresence: compilePresence,
    compilePresenceMigration: compilePresenceMigration,
    refreshWeeklyMatrix: refreshWeeklyMatrix,
    refreshKpis: refreshKpis,
    checkBackupConflicts: checkBackupConflicts,
    syncTasks: syncTasks,
    syncCalendarEvents: syncCalendarEvents,
    installLiveTaskSync: installLiveTaskSync,
    installAllTriggers: installAllTriggers,
    nightlySyncJob: nightlySyncJob,
    cleanupForeignTasks: cleanupForeignTasks,
    debugUserIdentity: debugUserIdentity,
    clientSetupOrUpgrade: clientSetupOrUpgrade,
    clientSyncMembersAndTenures: clientSyncMembersAndTenures,
    clientCompilePresence: clientCompilePresence,
    clientRefreshDashboards: clientRefreshDashboards,
    clientSyncTasks: clientSyncTasks,
    clientSyncCalendarEvents: clientSyncCalendarEvents,
    clientShowSidebar: clientShowSidebar,
    clientInstallTriggers: clientInstallTriggers,
    installedTaskEditHandler: installedTaskEditHandler,
    installedConsumerTriggerHandler: installedConsumerTriggerHandler,
    showLocationsTab: showLocationsTab,
    showContractsSidebar: showContractsSidebar,
    getContractsSidebarData: getContractsSidebarData,
    saveContract: saveContract,
    restoreMemberFormulas: restoreMemberFormulas,
    showHolidaysTab: showHolidaysTab,
    showBackupSidebar: showBackupSidebar,
    getBackupSidebarData: getBackupSidebarData,
    saveInterimBackup: saveInterimBackup,
    cancelInterimBackup: cancelInterimBackup,
    showWorkspaceCheck: showWorkspaceCheck,
    runWorkspaceCheck: runWorkspaceCheck,
    removePresenceDuplicates: removePresenceDuplicates,
    mondayPresenceReminder: mondayPresenceReminder,
    birthdayReminderJob: birthdayReminderJob,
    dailyRemindersJob: dailyRemindersJob,
    connectMyGoogleTasks: connectMyGoogleTasks,
    myTasksSyncJob: myTasksSyncJob,
    showAssignSidebar: showAssignSidebar,
    showTaskCalendar: showTaskCalendar,
    showTaskTypesTab: showTaskTypesTab,
    getAssignSidebarData: getAssignSidebarData,
    assignTasks: assignTasks,
    stopRecurringRule: stopRecurringRule,
    getSidebarInitialData: getSidebarInitialData,
    getMemberBaselines: getMemberBaselines,
    saveMemberBaselines: saveMemberBaselines,
    submitMemberLeave: submitMemberLeave
  };
}
