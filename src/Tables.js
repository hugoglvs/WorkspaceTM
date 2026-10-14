/**
 * Table registry (ADR 0039): the single source of truth for every sheet the
 * Library manages. Code refers to a stable key ("members"); people see a readable
 * tab name ("Members"). Tabs are found by their stored Google sheet ID first, so a
 * team can rename a tab without breaking the scripts.
 */

var TABLE_REGISTRY = {
  locations:    { sheetName: "Locations",       legacyName: "tbl_locations" },
  members:      { sheetName: "Members",         legacyName: "tbl_members" },
  tenures:      { sheetName: "Contracts",       legacyName: "tbl_tenures" },
  baselines:    { sheetName: "Routines",        legacyName: "tbl_presence_baselines" },
  leaves:       { sheetName: "Leaves",          legacyName: "tbl_leaves" },
  presenceLog:  { sheetName: "Presence Log",    legacyName: "tbl_presence_log" },
  projects:     { sheetName: "Projects",        legacyName: "tbl_projects" },
  tasks:        { sheetName: "Tasks",           legacyName: "tbl_tasks" },
  events:       { sheetName: "Events",          legacyName: "tbl_events" },
  backups:      { sheetName: "Backups",         legacyName: "tbl_backups" },
  weeklyMatrix: { sheetName: "Weekly Presence", legacyName: "view_weekly_matrix" },
  kpiData:      { sheetName: "KPI Data",        legacyName: "fact_presence" },
  holidays:     { sheetName: "Holidays",        legacyName: "tbl_holidays" },
  taskTypes:    { sheetName: "Task Templates",  legacyName: "tbl_task_types", formerNames: ["Task Types"] },
  recurringTasks: { sheetName: "Recurring Tasks", legacyName: "tbl_recurring_tasks" },
  taskCalendar: { sheetName: "Calendar",        legacyName: "view_task_calendar" },
  settings:     { sheetName: "Settings",        legacyName: "tbl_settings" }
};

var ARCHIVE_PREFIX = "Archive – ";
var LEGACY_ARCHIVE_NAMES = {
  presence_ranges: "Presence Ranges"
};

function tableIdProperty(key) {
  return "SHEET_ID_" + key;
}

function tableDocProps() {
  try {
    return (typeof PropertiesService !== "undefined") ? PropertiesService.getDocumentProperties() : null;
  } catch (e) {
    return null;
  }
}

/**
 * How many rows to read from a sheet: never fewer than the last row holding data,
 * so a growing table (Presence Log gets a row per Member per day) is never cut
 * short. The cap only limits how far empty rows are scanned.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} cap Rows to scan when the sheet is mostly empty
 * @return {number} Last row number to read (1 = header only)
 */
function readRowLimit(sheet, cap) {
  if (!sheet) return 1;
  var last = (typeof sheet.getLastRow === "function") ? (sheet.getLastRow() || 1) : 1;
  return Math.max(last, Math.min(sheet.getMaxRows(), cap));
}

/** Readable tab name of a table key ("tenures" -> "Contracts"). */
function tableSheetName(key) {
  return TABLE_REGISTRY[key] ? TABLE_REGISTRY[key].sheetName : key;
}

/** Key of a tab name, display or former ("Presence Log" / "tbl_presence_log" -> "presenceLog"), else null. */
function tableKeyOfName(name) {
  for (var key in TABLE_REGISTRY) {
    if (!TABLE_REGISTRY.hasOwnProperty(key)) continue;
    var e = TABLE_REGISTRY[key];
    if (e.sheetName === name || e.legacyName === name || key === name || (e.formerNames || []).indexOf(name) !== -1) return key;
  }
  return null;
}

/**
 * Finds a managed sheet: stored sheet ID, then readable name, then former name.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} key Registry key, e.g. "members"
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function getTable(ss, key) {
  if (!ss) return null;
  var entry = TABLE_REGISTRY[key];
  if (!entry) return ss.getSheetByName(key);

  var props = tableDocProps();
  var storedId = props ? props.getProperty(tableIdProperty(key)) : null;
  if (storedId) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (String(sheets[i].getSheetId()) === String(storedId)) return sheets[i];
    }
  }
  var found = ss.getSheetByName(entry.sheetName) || ss.getSheetByName(entry.legacyName);
  for (var f = 0; !found && f < (entry.formerNames || []).length; f++) found = ss.getSheetByName(entry.formerNames[f]);
  return found;
}

/** Registry key of a sheet (stored ID, readable or former name), else null. */
function tableKeyOfSheet(sheet) {
  if (!sheet) return null;
  var props = tableDocProps();
  if (props) {
    for (var key in TABLE_REGISTRY) {
      if (!TABLE_REGISTRY.hasOwnProperty(key)) continue;
      var id = props.getProperty(tableIdProperty(key));
      if (id && String(id) === String(sheet.getSheetId())) return key;
    }
  }
  return tableKeyOfName(sheet.getName());
}

/** Stores the sheet ID of every managed sheet found, so later renames are harmless. */
function registerTableIds(ss) {
  var props = tableDocProps();
  if (!props) return 0;
  var count = 0;
  for (var key in TABLE_REGISTRY) {
    if (!TABLE_REGISTRY.hasOwnProperty(key)) continue;
    var sheet = getTable(ss, key);
    if (sheet) {
      props.setProperty(tableIdProperty(key), String(sheet.getSheetId()));
      count++;
    }
  }
  return count;
}

function readableArchiveName(oldName) {
  var rest = oldName.replace(/^archive_/, "");
  var key = tableKeyOfName(rest);
  if (key) return ARCHIVE_PREFIX + TABLE_REGISTRY[key].sheetName;
  var bare = rest.replace(/^tbl_/, "").replace(/_\d+$/, "");
  if (LEGACY_ARCHIVE_NAMES[bare]) return ARCHIVE_PREFIX + LEGACY_ARCHIVE_NAMES[bare];
  return ARCHIVE_PREFIX + bare.replace(/_/g, " ").replace(/^\w/, function(c) { return c.toUpperCase(); });
}

/**
 * Renames former technical tabs (tbl_*, view_weekly_matrix, fact_presence,
 * archive_*) to their readable names. Idempotent; never overwrites an existing tab.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Array<{from: string, to: string}>}
 */
function renameLegacySheets(ss) {
  var renamed = [];
  for (var key in TABLE_REGISTRY) {
    if (!TABLE_REGISTRY.hasOwnProperty(key)) continue;
    var entry = TABLE_REGISTRY[key];
    if (ss.getSheetByName(entry.sheetName)) continue;
    [entry.legacyName].concat(entry.formerNames || []).forEach(function(old) {
      var legacy = ss.getSheetByName(old);
      if (legacy && !ss.getSheetByName(entry.sheetName)) {
        legacy.setName(entry.sheetName);
        renamed.push({ from: old, to: entry.sheetName });
      }
    });
  }
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var name = sheets[s].getName();
    if (name.indexOf("archive_") !== 0) continue;
    var target = readableArchiveName(name), n = 2, candidate = target;
    while (ss.getSheetByName(candidate)) candidate = target + " " + (n++);
    sheets[s].setName(candidate);
    renamed.push({ from: name, to: candidate });
  }
  return renamed;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    readRowLimit: readRowLimit,
    TABLE_REGISTRY: TABLE_REGISTRY,
    ARCHIVE_PREFIX: ARCHIVE_PREFIX,
    tableSheetName: tableSheetName,
    tableKeyOfName: tableKeyOfName,
    getTable: getTable,
    tableKeyOfSheet: tableKeyOfSheet,
    registerTableIds: registerTableIds,
    renameLegacySheets: renameLegacySheets
  };
}
