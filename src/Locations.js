/**
 * Locations (ADR 0040): the configurable list of Work Locations, kept in the
 * hidden "Locations" tab (Location | Type | Active). Type drives the KPIs
 * (Office, Client site, Remote). Inactive locations stay valid in history but are
 * no longer offered for new entries. "Off" and leave types are built in.
 */

var LOCATION_TYPES = ["Office", "Client site", "Remote"];
var DEFAULT_LOCATIONS = [
  ["OneRoof", "Office", true],
  ["Remote", "Remote", true],
  ["JLL", "Client site", true],
  ["C04", "Client site", true],
  ["C60", "Client site", true]
];
var OFF_LOCATION = "Off";
var LEAVE_LOCATIONS = ["Vacation", "Sick Leave", "Maternity/Paternity", "Unpaid Leave", "Holiday"];

/**
 * Presence colours, shared by Weekly Presence and the Members "Today" column:
 * Office purple (Accenture), Client site blue, Remote green, leaves red, Off grey.
 * Each entry is [background, font].
 */
var PRESENCE_STYLES = {
  office: ["#F3E8FD", "#7500C0"],
  client_site: ["#E8F0FE", "#1967D2"],
  remote: ["#E6F4EA", "#137333"],
  leave: ["#FCE8E6", "#C5221F"],
  off: ["#F1F3F4", "#5F6368"],
  not_logged: ["#FEF7E0", "#B06000"]
};

/**
 * The value → style list for presence cells: every listed location by its type,
 * leave types and Off.
 * @return {Array<{text: string, style: string[]}>}
 */
function presenceStyleEntries(ss) {
  var list = (typeof readLocations === "function" && ss) ? readLocations(ss) : [];
  var entries = list.map(function(l) {
    return { text: l.name, style: PRESENCE_STYLES[l.type === "Remote" ? "remote" : (l.type === "Client site" ? "client_site" : "office")] };
  });
  LEAVE_LOCATIONS.forEach(function(t) { entries.push({ text: t, style: PRESENCE_STYLES.leave }); });
  entries.push({ text: OFF_LOCATION, style: PRESENCE_STYLES.off });
  return entries;
}

function isPresenceTextRule(rule) {
  var cond = (rule && typeof rule.getBooleanCondition === "function") ? rule.getBooleanCondition() : null;
  if (!cond) return false;
  var type = String(cond.getCriteriaType());
  return type === "TEXT_EQUAL_TO" || type === "TEXT_STARTS_WITH";
}

/**
 * Colours the Members "Today" column like Weekly Presence. "Weekend" is grey,
 * "Not Logged" amber, and a split day ("C04 / Remote") takes its morning colour.
 * Other rules on the tab (such as the required-cell warnings) are kept.
 */
function applyTodayPresenceColours(ss) {
  if (!ss || typeof SpreadsheetApp === "undefined" || !SpreadsheetApp.newConditionalFormatRule) return false;
  var sheet = getTable(ss, "members");
  var config = SCHEMA_SPEC.members;
  if (!sheet || !config) return false;
  var col = config.columns.indexOf("today_presence") + 1;
  if (col < 1) return false;
  var range = sheet.getRange(2, col, Math.max(sheet.getMaxRows(), 50) - 1, 1);

  var entries = presenceStyleEntries(ss);
  var make = function(style) {
    return SpreadsheetApp.newConditionalFormatRule().setBackground(style[0]).setFontColor(style[1]).setRanges([range]);
  };
  var rules = [];
  entries.forEach(function(en) { rules.push(make(en.style).whenTextEqualTo(en.text).build()); });
  rules.push(make(PRESENCE_STYLES.off).whenTextEqualTo("Weekend").build());
  rules.push(make(PRESENCE_STYLES.off).whenTextEqualTo("Public holiday").build());
  rules.push(make(PRESENCE_STYLES.not_logged).whenTextEqualTo("Not Logged").build());
  entries.forEach(function(en) { rules.push(make(en.style).whenTextStartsWith(en.text + " /").build()); });

  var kept = (sheet.getConditionalFormatRules() || []).filter(function(r) { return !isPresenceTextRule(r); });
  sheet.setConditionalFormatRules(rules.concat(kept));
  return true;
}

function isTruthyCell(v) {
  if (v === true) return true;
  if (v === false) return false;
  var s = String(v === undefined || v === null ? "" : v).trim().toLowerCase();
  if (s === "") return true; // a new row without the box ticked counts as active
  return s === "true" || s === "yes" || s === "y" || s === "1" || s === "x" || s === "✓";
}

function normalizeLocationType(t) {
  var s = String(t || "").trim().toLowerCase();
  if (s.indexOf("remote") !== -1 || s.indexOf("home") !== -1) return "Remote";
  if (s.indexOf("client") !== -1) return "Client site";
  return "Office";
}

/**
 * Every configured location, in tab order. Falls back to the defaults when the
 * tab is missing or empty.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Array<{name: string, type: string, active: boolean}>}
 */
function readLocations(ss) {
  var sheet = (typeof getTable === "function") ? getTable(ss, "locations") : null;
  var list = [];
  if (sheet && sheet.getMaxRows() > 1) {
    var vals = sheet.getRange(2, 1, readRowLimit(sheet, 200) - 1, 3).getValues();
    for (var r = 0; r < vals.length; r++) {
      var name = String(vals[r][0] || "").trim();
      if (!name) continue;
      list.push({ name: name, type: normalizeLocationType(vals[r][1]), active: isTruthyCell(vals[r][2]) });
    }
  }
  if (!list.length) {
    list = DEFAULT_LOCATIONS.map(function(d) { return { name: d[0], type: d[1], active: d[2] }; });
  }
  return list;
}

function activeLocationNames(ss) {
  return readLocations(ss).filter(function(l) { return l.active; }).map(function(l) { return l.name; });
}

function allLocationNames(ss) {
  return readLocations(ss).map(function(l) { return l.name; });
}

/** First active location of a type, e.g. the default office. */
function firstActiveLocationOfType(ss, type) {
  var list = readLocations(ss);
  for (var i = 0; i < list.length; i++) if (list[i].active && list[i].type === type) return list[i].name;
  return null;
}

/** Default Work Location for a weekday without any routine: the first active office. */
function defaultOfficeLocation(ss) {
  return firstActiveLocationOfType(ss, "Office") || (activeLocationNames(ss)[0] || "Office");
}

/**
 * KPI category of a Work Location: "office", "client_site" or "remote".
 * Unlisted names (older data) are guessed from their wording.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet|Array} ssOrList Spreadsheet or readLocations() result
 * @param {string} name
 * @return {string}
 */
function locationCategory(ssOrList, name) {
  var list = Array.isArray(ssOrList) ? ssOrList : readLocations(ssOrList);
  var key = String(name || "").trim().toLowerCase();
  for (var i = 0; i < list.length; i++) {
    if (list[i].name.toLowerCase() === key) {
      return list[i].type === "Remote" ? "remote" : (list[i].type === "Client site" ? "client_site" : "office");
    }
  }
  var t = normalizeLocationType(key);
  return t === "Remote" ? "remote" : (t === "Client site" ? "client_site" : "office");
}

/** Writes the default list into an empty Locations tab. */
function seedDefaultLocations(ss) {
  var sheet = getTable(ss, "locations");
  if (!sheet) return false;
  var existing = sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), 1).getValues().some(function(r) { return String(r[0] || "").trim(); });
  if (existing) return false;
  sheet.getRange(2, 1, DEFAULT_LOCATIONS.length, 3).setValues(DEFAULT_LOCATIONS.map(function(d) { return d.slice(); }));
  return true;
}

/**
 * Dropdown lists of a table, with location columns filled from the Locations tab.
 * Data tables accept every location (active or not) so history stays valid.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} config A SCHEMA_SPEC entry
 * @return {Object} validations
 */
function schemaValidationsFor(ss, config) {
  var validations = {};
  var base = config.validations || {};
  for (var k in base) if (base.hasOwnProperty(k)) validations[k] = base[k];
  var dyn = config.locationColumns || {};
  for (var col in dyn) {
    if (!dyn.hasOwnProperty(col)) continue;
    var values = allLocationNames(ss).concat([OFF_LOCATION]);
    if (dyn[col] === "withLeaves") values = values.concat(LEAVE_LOCATIONS);
    validations[col] = values;
  }
  // Task type dropdowns follow the Task Templates tab (ADR 0045)
  var tt = config.taskTypeColumns || {};
  for (var tcol in tt) {
    if (!tt.hasOwnProperty(tcol)) continue;
    validations[tcol] = (typeof taskTypeNames === "function") ? taskTypeNames(ss) : [];
  }
  // Project dropdowns list the Projects tab's names (ADR 0053)
  var pc = config.projectColumns || {};
  for (var pcol in pc) {
    if (!pc.hasOwnProperty(pcol)) continue;
    var pnFn = (typeof projectNames === "function") ? projectNames
      : (typeof require !== "undefined" ? require("./Projects").projectNames : null);
    validations[pcol] = pnFn ? pnFn(ss) : [];
  }
  // "Created when" lists the task triggers (ADR 0051)
  var trig = config.triggerColumns || {};
  for (var gcol in trig) {
    if (!trig.hasOwnProperty(gcol)) continue;
    var labels = (typeof taskTriggerLabels === "function") ? taskTriggerLabels()
      : (typeof require !== "undefined" ? require("./TaskTriggers").taskTriggerLabels() : []);
    validations[gcol] = labels;
  }
  return validations;
}

/** Re-applies the dropdowns that follow other tabs (locations, task types, projects) to the data tables. */
function refreshLocationDropdowns(ss) {
  var updated = [];
  for (var key in SCHEMA_SPEC) {
    if (!SCHEMA_SPEC.hasOwnProperty(key) || !(SCHEMA_SPEC[key].locationColumns || SCHEMA_SPEC[key].taskTypeColumns || SCHEMA_SPEC[key].projectColumns)) continue;
    var config = SCHEMA_SPEC[key];
    var sheet = getTable(ss, key);
    if (!sheet) continue;
    createGoogleSheetsTable(ss.getId(), sheet.getSheetId(), config.sheetName, config.columns, config.columnTypes,
      schemaValidationsFor(ss, config), Math.min(Math.max(sheet.getMaxRows(), 50), 25), schemaHeaderRow(config));
    updated.push(config.sheetName);
  }
  try { applyTodayPresenceColours(ss); } catch (eCol) { Logger.log("Today colours note: " + eCol); }
  return updated;
}


/**
 * Where a location is still planned: ongoing or future routines, and Presence Log
 * days from today on. Past days are history and are not counted.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @param {Date} [asOf] Defaults to today
 * @return {{routineMembers: string[], routineCount: number, futureDays: number}}
 */
function findLocationUsage(ss, name, asOf) {
  var todayIso = isoDate(asOf || new Date());
  var key = String(name || "").trim().toLowerCase();
  var res = { routineMembers: [], routineCount: 0, futureDays: 0 };
  var read = function(sheet, numCols) {
    if (!sheet || sheet.getMaxRows() < 2) return { vals: [], rich: [] };
    var range = sheet.getRange(2, 1, readRowLimit(sheet, 2000) - 1, numCols);
    return { vals: range.getValues(), rich: (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [] };
  };
  var memberOf = function(data, r, c) {
    var raw = data.vals[r][c];
    return (typeof extractCellValue === "function") ? extractCellValue(raw, data.rich[r] ? data.rich[r][c] : null) : String(raw || "").trim();
  };

  // Routines: baseline_id, member, start_date, end_date, day_of_week, location_am, location_pm
  var routines = read(getTable(ss, "baselines"), 7);
  for (var r = 0; r < routines.vals.length; r++) {
    var row = routines.vals[r];
    var member = memberOf(routines, r, 1);
    if (!member) continue;
    var end = isoDate(row[3]);
    if (end && end < todayIso) continue;
    if (String(row[5]).trim().toLowerCase() === key || String(row[6]).trim().toLowerCase() === key) {
      res.routineCount++;
      if (res.routineMembers.indexOf(member) === -1) res.routineMembers.push(member);
    }
  }

  // Presence Log: log_id, date, member, location_am, location_pm, timestamp
  var log = read(getTable(ss, "presenceLog"), 5);
  for (var l = 0; l < log.vals.length; l++) {
    var date = isoDate(log.vals[l][1]);
    if (!date || date < todayIso) continue;
    if (String(log.vals[l][3]).trim().toLowerCase() === key || String(log.vals[l][4]).trim().toLowerCase() === key) res.futureDays++;
  }
  return res;
}

/**
 * onEdit hook on the Locations tab: when a location is retired (Active unticked),
 * a toast lists the routines and future days that still use it.
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {Object|null} The usage found, or null when nothing to report
 */
function warnRetiredLocation(e) {
  if (!e || !e.range || e.range.getRow() < 2) return null;
  var sheet = e.range.getSheet();
  var row = sheet.getRange(e.range.getRow(), 1, 1, 3).getValues()[0];
  var name = String(row[0] || "").trim();
  if (!name || isTruthyCell(row[2])) return null;

  var ss = e.source || sheet.getParent();
  var usage = findLocationUsage(ss, name);
  if (!usage.routineCount && !usage.futureDays) return null;

  var parts = [];
  if (usage.routineCount) {
    parts.push(usage.routineCount + " routine" + (usage.routineCount > 1 ? "s" : "") + " (" + usage.routineMembers.join(", ") + ")");
  }
  if (usage.futureDays) {
    parts.push(usage.futureDays + " future day" + (usage.futureDays > 1 ? "s" : "") + " in Weekly Presence");
  }
  try {
    ss.toast(name + " is now inactive but is still used by " + parts.join(" and ") +
      ". Update them with Declare Presence & Leaves, or in Weekly Presence.", "Location still in use", 12);
  } catch (eToast) {}
  return usage;
}

/** Menu action: shows the hidden Locations tab so it can be edited. */
function showLocationsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getTable(ss, "locations");
  if (!sheet) return { success: false, message: "The Locations tab was not found. Run Initialize / Upgrade first." };
  if (sheet.isSheetHidden()) sheet.showSheet();
  sheet.activate();
  try {
    ss.toast("Edit the list, then hide the tab again (right-click → Hide sheet). Untick Active to retire a location.", "Locations", 8);
  } catch (eToast) {}
  return { success: true };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    LOCATION_TYPES: LOCATION_TYPES,
    DEFAULT_LOCATIONS: DEFAULT_LOCATIONS,
    LEAVE_LOCATIONS: LEAVE_LOCATIONS,
    readLocations: readLocations,
    activeLocationNames: activeLocationNames,
    allLocationNames: allLocationNames,
    firstActiveLocationOfType: firstActiveLocationOfType,
    defaultOfficeLocation: defaultOfficeLocation,
    locationCategory: locationCategory,
    seedDefaultLocations: seedDefaultLocations,
    schemaValidationsFor: schemaValidationsFor,
    refreshLocationDropdowns: refreshLocationDropdowns,
    showLocationsTab: showLocationsTab,
    findLocationUsage: findLocationUsage,
    warnRetiredLocation: warnRetiredLocation,
    PRESENCE_STYLES: PRESENCE_STYLES,
    presenceStyleEntries: presenceStyleEntries,
    applyTodayPresenceColours: applyTodayPresenceColours
  };
}
