/**
 * Project Management Workspace - Presence KPI Engine
 *
 * Builds the hidden fact_presence table (one row per Member x date x Period)
 * that Looker Studio aggregates into presence, remote-work and capacity KPIs
 * (ADR 0034, 0036). AM and PM each weigh 0.5 day; capacity follows the Tenure FTE.
 */

/**
 * Normalizes and parses raw date cell into Date object at midnight.
 *
 * @param {Date|string|number} rawDate
 * @return {Date|null}
 */
function parseKpiDate(rawDate) {
  if (!rawDate) return null;
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    var d = new Date(rawDate.getTime());
    d.setHours(0, 0, 0, 0);
    return d;
  }
  var str = String(rawDate).trim();
  var isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (isoMatch) {
    return new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10), 0, 0, 0, 0);
  }
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    parsed.setHours(0, 0, 0, 0);
    return parsed;
  }
  return null;
}

/**
 * Returns month key in YYYY-MM format.
 *
 * @param {Date} date
 * @return {string}
 */
function formatMonthKey(date) {
  if (!date || isNaN(date.getTime())) return "";
  var y = date.getFullYear();
  var m = String(date.getMonth() + 1);
  if (m.length < 2) m = "0" + m;
  return y + "-" + m;
}

/**
 * Returns total number of weekdays (Mon-Fri) in a given YYYY-MM month.
 *
 * @param {string} monthKey - E.g. "2026-01"
 * @return {number} Count of weekdays
 */
function countMonthWeekdays(monthKey) {
  if (!monthKey || typeof monthKey !== "string") return 0;
  var parts = monthKey.split("-");
  if (parts.length !== 2) return 0;
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10); // 1-12
  if (isNaN(year) || isNaN(month)) return 0;

  var startDate = new Date(year, month - 1, 1, 0, 0, 0, 0);
  var endDate = new Date(year, month, 0, 0, 0, 0, 0); // Last day of month

  var count = 0;
  var curr = new Date(startDate.getTime());
  while (curr <= endDate) {
    var day = curr.getDay();
    if (day >= 1 && day <= 5) {
      count++;
    }
    curr.setDate(curr.getDate() + 1);
  }
  return count;
}

/**
 * Determines whether a location string represents a non-working leave or holiday.
 *
 * @param {string} location
 * @return {boolean}
 */
function isLeaveLocation(location) {
  if (!location) return false;
  var loc = String(location).trim().toLowerCase();
  return (
    loc.indexOf("vacation") !== -1 ||
    loc.indexOf("sick") !== -1 ||
    loc.indexOf("holiday") !== -1 ||
    loc.indexOf("unpaid") !== -1 ||
    loc.indexOf("maternity") !== -1 ||
    loc.indexOf("paternity") !== -1 ||
    loc === "leave" ||
    loc === "absence"
  );
}

/**
 * Determines whether a location means a planned non-working Period that is not
 * a Leave (e.g. a weekend afternoon "Off"). It counts as neither worked nor leave.
 *
 * @param {string} location
 * @return {boolean}
 */
function isOffLocation(location) {
  if (!location) return false;
  return String(location).trim().toLowerCase() === "off";
}

/**
 * Normalises an fte_rate cell to a fraction of full time.
 * Values above 1 are read as percentages (80 -> 0.8); blank or invalid -> 1.0.
 *
 * @param {*} raw
 * @return {number}
 */
function normalizeFteRate(raw) {
  if (raw === "" || raw === null || raw === undefined) return 1.0;
  var n = parseFloat(String(raw).replace("%", "").replace(",", "."));
  if (isNaN(n) || n < 0) return 1.0;
  return n > 1 ? n / 100 : n;
}

/**
 * Categorizes a work location into one of standard buckets:
 * "remote", "main_office", "downtown_office", "branch_office", "client_site", "other_office", or "none" (if leave or empty).
 *
 * @param {string} location
 * @return {string}
 */
function categorizeLocation(location) {
  if (!location) return "none";
  var loc = String(location).trim();
  if (!loc) return "none";
  if (isLeaveLocation(loc) || isOffLocation(loc)) return "none";

  var lower = loc.toLowerCase();
  if (lower === "remote" || lower.indexOf("home") !== -1) {
    return "remote";
  }
  if (lower.indexOf("main office") !== -1) {
    return "main_office";
  }
  if (lower.indexOf("downtown") !== -1) {
    return "downtown_office";
  }
  if (lower.indexOf("branch") !== -1) {
    return "branch_office";
  }
  if (lower.indexOf("client") !== -1) {
    return "client_site";
  }
  return "other_office";
}

var FACT_PRESENCE_SHEET = "fact_presence";
var FACT_PRESENCE_HEADERS = [
  "date", "month", "iso_week", "weekday", "is_weekend", "member", "department", "role",
  "fte_rate", "period", "location", "category", "worked_days", "leave_days", "capacity_days"
];
var WEEKDAY_LABELS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

function formatIsoDate(d) {
  var m = String(d.getMonth() + 1);
  var day = String(d.getDate());
  return d.getFullYear() + "-" + (m.length < 2 ? "0" + m : m) + "-" + (day.length < 2 ? "0" + day : day);
}

function formatIsoWeek(d) {
  var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  var dayNum = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - dayNum); // Thursday of this ISO week
  var yearStart = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  var week = Math.ceil(((t - yearStart) / 86400000 + 1) / 7);
  return t.getUTCFullYear() + "-W" + (week < 10 ? "0" + week : week);
}

/**
 * Classifies one Period's Work Location for the fact table.
 *
 * @param {string} location
 * @return {string} remote | office | leave | off | not_logged
 */
function classifyPeriod(location) {
  if (!location || !String(location).trim()) return "not_logged";
  if (isLeaveLocation(location)) return "leave";
  if (isOffLocation(location)) return "off";
  return categorizeLocation(location) === "remote" ? "remote" : "office";
}

function readCellName(vals, rich, r, c) {
  var raw = vals[r][c];
  var rt = (rich && rich[r]) ? rich[r][c] : null;
  return (typeof extractCellValue === "function") ? extractCellValue(raw, rt) : String(raw || "").trim();
}

function readSheetRows(sheet, numCols, maxRows) {
  if (!sheet || sheet.getMaxRows() < 2) return { vals: [], rich: [] };
  var n = Math.min(sheet.getMaxRows(), maxRows) - 1;
  var range = sheet.getRange(2, 1, n, Math.max(sheet.getLastColumn(), numCols));
  return {
    vals: range.getValues(),
    rich: (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : []
  };
}

/**
 * Builds the tidy presence fact table for Looker Studio: one row per
 * Member x date x Period, from the first logged date to asOfDate (default today).
 * Weekdays inside a Tenure always get rows (unlogged = "not_logged") so capacity
 * is complete; other dates get rows only when logged. No totals: Looker aggregates.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {{asOfDate: (Date|string)}} [options]
 * @return {{success: boolean, headers: string[], rows: Array[]}}
 */
function buildPresenceFacts(targetSpreadsheet, options) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var asOf = parseKpiDate((options && options.asOfDate) || new Date());

  // Members: department fallback from tbl_members
  var memberDept = {};
  var mRows = readSheetRows(ss.getSheetByName("tbl_members"), 7, 500);
  for (var m = 0; m < mRows.vals.length; m++) {
    var mName = readCellName(mRows.vals, mRows.rich, m, 0);
    if (mName) memberDept[mName.toLowerCase()] = String(mRows.vals[m][4] || "").trim();
  }

  // Tenures, grouped by Member
  var members = {}; // key -> { name, tenures: [], logs: {} }
  function memberEntry(name) {
    var key = name.toLowerCase();
    if (!members[key]) members[key] = { name: name, tenures: [], logs: {} };
    return members[key];
  }
  var tRows = readSheetRows(ss.getSheetByName("tbl_tenures"), 7, 500);
  for (var t = 0; t < tRows.vals.length; t++) {
    var tName = readCellName(tRows.vals, tRows.rich, t, 0);
    if (!tName) continue;
    memberEntry(tName).tenures.push({
      role: String(tRows.vals[t][1] || "").trim(),
      department: String(tRows.vals[t][2] || "").trim(),
      start: parseKpiDate(tRows.vals[t][3]),
      end: parseKpiDate(tRows.vals[t][4]),
      fte: normalizeFteRate(tRows.vals[t][6])
    });
  }

  // Presence Log
  var firstDate = null;
  var lRows = readSheetRows(ss.getSheetByName("tbl_presence_log"), 6, 20000);
  for (var r = 0; r < lRows.vals.length; r++) {
    var d = parseKpiDate(lRows.vals[r][1]);
    var lName = readCellName(lRows.vals, lRows.rich, r, 2);
    if (!d || !lName || d > asOf) continue;
    memberEntry(lName).logs[formatIsoDate(d)] = { am: lRows.vals[r][3], pm: lRows.vals[r][4] };
    if (!firstDate || d < firstDate) firstDate = d;
  }

  var rows = [];
  if (!firstDate) return { success: true, headers: FACT_PRESENCE_HEADERS, rows: rows };

  var keys = Object.keys(members).sort();
  for (var k = 0; k < keys.length; k++) {
    var mem = members[keys[k]];
    for (var cur = new Date(firstDate.getTime()); cur <= asOf; cur.setDate(cur.getDate() + 1)) {
      var iso = formatIsoDate(cur);
      var dow = cur.getDay();
      var isWeekend = (dow === 0 || dow === 6);
      var tenure = null;
      for (var ti = 0; ti < mem.tenures.length; ti++) {
        var ten = mem.tenures[ti];
        if ((!ten.start || ten.start <= cur) && (!ten.end || ten.end >= cur)) { tenure = ten; break; }
      }
      var log = mem.logs[iso];
      if (!log && (isWeekend || !tenure)) continue;

      var fte = tenure ? tenure.fte : 0;
      var dept = (tenure && tenure.department) || memberDept[keys[k]] || "General";
      var periods = [["AM", log ? log.am : ""], ["PM", log ? log.pm : ""]];
      for (var p = 0; p < periods.length; p++) {
        var location = String(periods[p][1] || "").trim();
        var category = classifyPeriod(location);
        rows.push([
          iso, formatMonthKey(cur), formatIsoWeek(cur), WEEKDAY_LABELS[dow], isWeekend,
          mem.name, dept, tenure ? tenure.role : "", fte,
          periods[p][0], location, category,
          (category === "remote" || category === "office") ? 0.5 : 0,
          category === "leave" ? 0.5 : 0,
          (!isWeekend && tenure) ? 0.5 * fte : 0
        ]);
      }
    }
  }
  return { success: true, headers: FACT_PRESENCE_HEADERS, rows: rows };
}

/**
 * Rewrites the hidden fact_presence sheet that feeds Looker Studio.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {{asOfDate: (Date|string)}} [options]
 * @return {{success: boolean, rowsCount: number}}
 */
function renderPresenceFacts(targetSpreadsheet, options) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var facts = buildPresenceFacts(ss, options);
  var sheet = ss.getSheetByName(FACT_PRESENCE_SHEET) || ss.insertSheet(FACT_PRESENCE_SHEET);
  var numCols = FACT_PRESENCE_HEADERS.length;
  var needed = facts.rows.length + 1;

  if (sheet.getMaxRows() < needed) sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());
  if (sheet.getMaxColumns() < numCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  sheet.getRange(1, 1, 1, numCols).setValues([FACT_PRESENCE_HEADERS]);
  if (facts.rows.length > 0) sheet.getRange(2, 1, facts.rows.length, numCols).setValues(facts.rows);
  sheet.setFrozenRows(1);
  if (typeof sheet.isSheetHidden === "function" && !sheet.isSheetHidden()) sheet.hideSheet();

  return { success: true, rowsCount: facts.rows.length };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseKpiDate: parseKpiDate,
    formatMonthKey: formatMonthKey,
    countMonthWeekdays: countMonthWeekdays,
    isLeaveLocation: isLeaveLocation,
    isOffLocation: isOffLocation,
    normalizeFteRate: normalizeFteRate,
    categorizeLocation: categorizeLocation,
    buildPresenceFacts: buildPresenceFacts,
    renderPresenceFacts: renderPresenceFacts
  };
}
