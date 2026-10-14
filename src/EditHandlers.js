/**
 * Edit handlers, one entry per table (ADR 0049).
 *
 * handleSheetEdit (src/Code.js) looks up the edited tab's key in EDIT_HANDLERS:
 *   only   - the tab is a view (Weekly Presence, Calendar): its handler does
 *            everything and the common checks are skipped;
 *   before - run before the common checks (row IDs, task status);
 *   after  - run after them (members sync, refreshes, coverage).
 * The common checks for every schema table are: a toast for missing required
 * cells, then the protected-column warning.
 * A handler gets (e, ctx) with ctx = {ss, sheet, tableKey, row, col, colName};
 * returning "stop" ends the processing of this edit.
 */

/** A function by name: the Apps Script global, else (Node tests) from its module. */
function edFn(name, modulePath) {
  var g = (typeof globalThis !== "undefined") ? globalThis : this;
  if (g && typeof g[name] === "function") return g[name];
  if (typeof require !== "undefined" && modulePath) {
    try { var mod = require(modulePath); if (mod && typeof mod[name] === "function") return mod[name]; } catch (e) {}
  }
  return null;
}

/** Calls a function by name if it exists; returns its result. */
function edCall(name, modulePath, args) {
  var fn = edFn(name, modulePath);
  return fn ? fn.apply(null, args || []) : undefined;
}

// ---- Views -----------------------------------------------------------------

function editWeeklyMatrix(e) { edCall("handleWeeklyMatrixEdit", "./ViewBuilder", [e]); }

function editTaskCalendar(e) { edCall("handleTaskCalendarEdit", "./TaskCalendar", [e]); }

// ---- Tasks, task types, recurring tasks, events ------------------------------

/** Status column: Done on / Done by, closed-month confirmation (ADR 0045). */
function editTaskStatus(e, ctx) {
  if (ctx.row < 2) return;
  if (edCall("handleTaskStatusEdit", "./TaskCalendar", [e]) === "kept") return "stop";
}

/** New or edited task: ID, defaults, Google Tasks. */
function editTaskRow(e, ctx) {
  if (ctx.row < 2) return;
  var fn = edFn("handleTaskEdit", "./TaskSyncService") ||
    ((typeof TaskSyncService !== "undefined" && typeof TaskSyncService.handleTaskEdit === "function") ? TaskSyncService.handleTaskEdit : null);
  if (fn) fn(e);
}

/** Task Types edited: the type dropdowns follow (ADR 0045). */
function editTaskTypes(e, ctx) {
  if (ctx.row > 1) edCall("refreshLocationDropdowns", "./Locations", [ctx.ss]);
}

/** Recurring Tasks edited: plan the rules. */
function editRecurringTasks(e, ctx) {
  if (ctx.row > 1) edCall("generateRecurringTasks", "./RecurringTasks", [ctx.ss]);
}

/** New or edited event: ID, calendar sync. */
function editEventRow(e, ctx) {
  if (ctx.row < 2) return;
  var fn = edFn("handleEventEdit", "./CalendarSyncService") ||
    ((typeof CalendarSyncService !== "undefined" && typeof CalendarSyncService.handleEventEdit === "function") ? CalendarSyncService.handleEventEdit : null);
  if (fn) fn(e);
}

// ---- Members and contracts ---------------------------------------------------

/** A name deleted from Members removes the rest of that row (ADR 0044). */
function editMemberNameDeleted(e, ctx) {
  if (ctx.row < 2 || ctx.col !== 1 || e.oldValue === undefined || e.oldValue === "" ||
      (e.value !== undefined && e.value !== "") || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return;
  ctx.sheet.getRange(ctx.row, 2, 1, SCHEMA_SPEC.members.columns.length - 1).clearContent();
}

/** Birthday typed in Members (ADR 0042, 0049). */
function editMemberBirthday(e) { edCall("normalizeBirthdayEdit", "./Birthdays", [e]); }

/** A new name without a contract opens the Contracts sidebar (ADR 0048). */
function editMemberPromptContract(e) { edCall("promptContractForNewMember", "./Contracts", [e]); }

/** Members or Contracts edited: directory sync (new people, onboarding, formulas). */
function editSyncMembers(e, ctx) { edCall("syncMembersAndTenures", "./SyncEngine", [ctx.ss, { fromEdit: true }]); }

// ---- Settings, locations, projects, coverage --------------------------------

function editSettings(e) { edCall("handleSettingsEdit", "./Settings", [e]); }

/** Locations edited: dropdowns and Weekly Presence follow (ADR 0040). */
function editLocations(e, ctx) {
  if (ctx.row < 2) return;
  edCall("refreshLocationDropdowns", "./Locations", [ctx.ss]);
  edCall("refreshWeeklyMatrix", "./ViewBuilder", [ctx.ss]);
  // Last, so its toast is not replaced by the refresh toast
  edCall("warnRetiredLocation", "./Locations", [e]);
}

/** A project renamed: Tasks, Backups and Recurring Tasks follow; the dropdowns list the names (ADR 0053). */
function editProjectName(e, ctx) { edCall("handleProjectNameEdit", "./Projects", [e, ctx]); }

/** A project without a backup: ask for one (ADR 0038). */
function editProjectBackupPrompt(e) { edCall("promptProjectBackup", "./ProjectBackupEngine", [e]); }

/** Projects, leaves or backups edited: re-check coverage. */
function editCoverage(e, ctx) { edCall("checkProjectBackupConflicts", "./ProjectBackupEngine", [ctx.ss]); }

var EDIT_HANDLERS = {
  weeklyMatrix:   { only: editWeeklyMatrix },
  taskCalendar:   { only: editTaskCalendar },
  tasks:          { before: [editTaskStatus, editTaskRow] },
  taskTypes:      { before: [editTaskTypes] },
  recurringTasks: { before: [editRecurringTasks] },
  events:         { before: [editEventRow] },
  members:        { after: [editMemberNameDeleted, editMemberBirthday, editMemberPromptContract, editSyncMembers] },
  tenures:        { after: [editSyncMembers] },
  settings:       { after: [editSettings] },
  locations:      { after: [editLocations] },
  projects:       { after: [editProjectName, editProjectBackupPrompt, editCoverage] },
  leaves:         { after: [editCoverage] },
  backups:        { after: [editCoverage] }
};

// ---- Common checks ------------------------------------------------------------

/** Toast when the edited row misses a required value. */
function editCheckRequired(e, ctx) {
  if (ctx.row < 2 || !SCHEMA_SPEC[ctx.tableKey]) return;
  var validate = edFn("validateRow", "./ValidationManager");
  if (!validate) return;
  var config = SCHEMA_SPEC[ctx.tableKey];
  var numCols = Math.max((typeof ctx.sheet.getLastColumn === "function" ? ctx.sheet.getLastColumn() : 0) || 0, config.columns.length);
  var res = validate(ctx.tableKey, ctx.sheet.getRange(ctx.row, 1, 1, numCols).getValues()[0]);
  if (!res.populated || res.valid) return;
  var labels = res.missingColumns.map(function(c) { return schemaColumnLabel(config, c); });
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Missing required column(s): " + labels.join(", ") + " in " + ctx.sheet.getName() + " (Row " + ctx.row + ").",
      "Required Field Missing", 5);
  } catch (eToast) {}
}

/** Warning when a protected column is edited; typing a new name in an empty Members row is exempt (ADR 0044). */
function editCheckProtected(e, ctx) {
  if (!SCHEMA_SPEC[ctx.tableKey]) return;
  var addingMember = ctx.tableKey === "members" && ctx.colName === "member" &&
    (e.oldValue === undefined || e.oldValue === "") && e.value !== undefined && e.value !== "";
  if (addingMember) return;
  var isProtected = edFn("isCellProtected", "./ProtectionManager");
  var info = isProtected ? isProtected(ctx.sheet, ctx.row, ctx.col) : null;
  if (!info || !info.isProtected) return;
  var handler = edFn("handleProtectedCellEdit", "./ProtectionManager");
  if (handler) handler(ctx.sheet, ctx.row, ctx.col, ctx.colName, info);
}

var EDIT_COMMON_CHECKS = [editCheckRequired, editCheckProtected];

/**
 * Runs the handlers of one edit. Each handler runs on its own: an error is
 * logged and the next one still runs.
 *
 * @param {Object} e Apps Script edit event
 * @return {{tableKey: (string|null), ran: string[], stopped: boolean}}
 */
function dispatchSheetEdit(e) {
  var sheet = e.range.getSheet();
  var tableKey = (typeof tableKeyOfSheet === "function") ? tableKeyOfSheet(sheet) : null;
  var col = e.range.getColumn();
  var config = (typeof SCHEMA_SPEC !== "undefined" && tableKey) ? SCHEMA_SPEC[tableKey] : null;
  var ctx = {
    ss: (sheet.getParent && sheet.getParent()) || SpreadsheetApp.getActiveSpreadsheet(),
    sheet: sheet, tableKey: tableKey, row: e.range.getRow(), col: col,
    colName: (config && config.columns && col <= config.columns.length) ? config.columns[col - 1] : ""
  };
  var entry = EDIT_HANDLERS[tableKey] || {};
  var started = Date.now();
  var report = { tableKey: tableKey, ran: [], stopped: false };
  var run = function(list) {
    for (var i = 0; i < list.length && !report.stopped; i++) {
      var name = list[i].name || "handler";
      var t0 = Date.now();
      try {
        if (list[i](e, ctx) === "stop") report.stopped = true;
      } catch (err) {
        Logger.log("Edit handler " + name + " (" + tableKey + ") note: " + err);
      }
      report.ran.push(name + " " + (Date.now() - t0) + "ms");
    }
  };
  if (entry.only) {
    run([entry.only]);
  } else {
    run(entry.before || []);
    run(EDIT_COMMON_CHECKS);
    run(entry.after || []);
  }
  if (Date.now() - started > 3000) Logger.log("[edit timing] " + (tableKey || sheet.getName()) + ": " + report.ran.join(", "));
  return report;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    EDIT_HANDLERS: EDIT_HANDLERS,
    dispatchSheetEdit: dispatchSheetEdit
  };
}
