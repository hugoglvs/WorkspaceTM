/**
 * Presence Ranges: the single table (tbl_presence_ranges) holding every
 * date range that shapes a Member's presence (ADR 0035):
 *   - Baseline: a recurring default Work Location per weekday.
 *   - Leave: an absence; its leave type is stored as the Work Location.
 * Every other module reads and writes this table only through this file.
 */

var PRESENCE_RANGES_SHEET = "tbl_presence_ranges";
var PRESENCE_RANGE_COLUMNS = ["range_id", "member", "range_type", "start_date", "start_period", "end_date", "end_period", "day_of_week", "location_am", "location_pm", "status"];
var LEGACY_PRESENCE_SHEETS = ["tbl_presence_baselines", "tbl_leaves"];
var PRESENCE_HIDDEN_EXTRA_SHEETS = ["fact_presence", "calc_kpis"];

/** 1-based column index of a tbl_presence_ranges column. */
function presenceRangeColumn(name) {
  return PRESENCE_RANGE_COLUMNS.indexOf(name) + 1;
}

/** Returns a cell as a "YYYY-MM-DD" string ("" when blank). */
function toIsoDateString(val) {
  if (val === "" || val === null || val === undefined) return "";
  if (val instanceof Date && !isNaN(val.getTime())) {
    var m = String(val.getMonth() + 1), d = String(val.getDate());
    return val.getFullYear() + "-" + (m.length < 2 ? "0" + m : m) + "-" + (d.length < 2 ? "0" + d : d);
  }
  var str = String(val).trim();
  var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[0];
  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? str : toIsoDateString(parsed);
}

/** Parses a cell into a Date at local midnight, or null. */
function parseRangeDate(val) {
  var iso = toIsoDateString(val);
  var m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)) : null;
}

function cellText(val, rich) {
  return (typeof extractCellValue === "function") ? extractCellValue(val, rich) : String(val || "").trim();
}

function generateRangeId(rangeType) {
  return (rangeType === "Leave" ? "leave_" : "base_") + Math.random().toString(36).substring(2, 9);
}

/**
 * Reads every populated row of tbl_presence_ranges.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Array<Object>} { row, rangeId, member, memberRaw, memberRich, rangeType, startDate, startIso,
 *   startPeriod, endDate, endIso, endPeriod, dayOfWeek, locationAm, locationPm, status }
 */
function readPresenceRanges(ss) {
  var sheet = ss ? ss.getSheetByName(PRESENCE_RANGES_SHEET) : null;
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var n = Math.min(sheet.getMaxRows(), 2000) - 1;
  var range = sheet.getRange(2, 1, n, PRESENCE_RANGE_COLUMNS.length);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  var c = {};
  PRESENCE_RANGE_COLUMNS.forEach(function(name, i) { c[name] = i; });

  var out = [];
  for (var r = 0; r < vals.length; r++) {
    var row = vals[r];
    var memberRich = (rich && rich[r]) ? rich[r][c.member] : null;
    var member = cellText(row[c.member], memberRich);
    if (!member && !row[c.range_id]) continue;
    out.push({
      row: r + 2,
      rangeId: cellText(row[c.range_id]),
      member: member,
      memberRaw: row[c.member],
      memberRich: memberRich,
      rangeType: cellText(row[c.range_type]) || "Baseline",
      startDate: parseRangeDate(row[c.start_date]),
      startIso: toIsoDateString(row[c.start_date]),
      startPeriod: String(row[c.start_period] || "").trim().toUpperCase(),
      endDate: parseRangeDate(row[c.end_date]),
      endIso: toIsoDateString(row[c.end_date]),
      endPeriod: String(row[c.end_period] || "").trim().toUpperCase(),
      dayOfWeek: cellText(row[c.day_of_week]),
      locationAm: cellText(row[c.location_am]),
      locationPm: cellText(row[c.location_pm]),
      status: cellText(row[c.status])
    });
  }
  return out;
}

/**
 * Baselines grouped by lower-cased Member name (Cancelled ones excluded).
 * @return {Object<string, Array<{startDate, endDate, dayOfWeek, locationAm, locationPm}>>}
 */
function readBaselinesByMember(ss) {
  var byMember = {};
  readPresenceRanges(ss).forEach(function(r) {
    if (r.rangeType !== "Baseline" || !r.member || r.status.toLowerCase() === "cancelled") return;
    var key = r.member.toLowerCase();
    (byMember[key] = byMember[key] || []).push({
      startDate: r.startDate,
      endDate: r.endDate,
      dayOfWeek: r.dayOfWeek,
      locationAm: r.locationAm || "Main Office",
      locationPm: r.locationPm || "Main Office"
    });
  });
  return byMember;
}

/**
 * Approved Leaves grouped by lower-cased Member name. Missing periods default to a
 * full day (AM start, PM end); a missing end date means a single-day Leave.
 * @return {Object<string, Array<{startDate, startPeriod, endDate, endPeriod, leaveType}>>}
 */
function readApprovedLeavesByMember(ss) {
  var byMember = {};
  readPresenceRanges(ss).forEach(function(r) {
    if (r.rangeType !== "Leave" || !r.member || !r.startDate || r.status.toLowerCase() !== "approved") return;
    var key = r.member.toLowerCase();
    (byMember[key] = byMember[key] || []).push({
      startDate: r.startDate,
      startPeriod: r.startPeriod || "AM",
      endDate: r.endDate || new Date(r.startDate.getTime()),
      endPeriod: r.endPeriod || "PM",
      leaveType: r.locationAm || r.locationPm || "Vacation"
    });
  });
  return byMember;
}

function nextFreeRangeRow(sheet) {
  var last = (typeof sheet.getLastRow === "function") ? (sheet.getLastRow() || 1) : 1;
  if (last < 2) return 2;
  var vals = sheet.getRange(2, 1, last - 1, 3).getValues();
  for (var r = vals.length - 1; r >= 0; r--) {
    if (vals[r][0] || vals[r][1] || vals[r][2]) return r + 3;
  }
  return 2;
}

/**
 * Appends Presence Ranges. Each item is keyed by column name; range_id and
 * status get defaults when missing.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Array<Object>} items
 * @return {{firstRow: number, count: number}}
 */
function appendPresenceRanges(ss, items) {
  var sheet = ss.getSheetByName(PRESENCE_RANGES_SHEET);
  if (!sheet) throw new Error("Sheet " + PRESENCE_RANGES_SHEET + " not found. Please run Initialize / Upgrade first.");
  if (!items || items.length === 0) return { firstRow: 0, count: 0 };

  var rows = items.map(function(item) {
    var type = item.range_type || "Baseline";
    return PRESENCE_RANGE_COLUMNS.map(function(col) {
      if (col === "range_id") return item.range_id || generateRangeId(type);
      if (col === "range_type") return type;
      if (col === "status") return item.status || "Approved";
      var v = item[col];
      return (v === undefined || v === null) ? "" : v;
    });
  });
  var firstRow = nextFreeRangeRow(sheet);
  var needed = firstRow + rows.length - 1;
  if (sheet.getMaxRows() < needed && typeof sheet.insertRowsAfter === "function") {
    sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());
  }
  sheet.getRange(firstRow, 1, rows.length, PRESENCE_RANGE_COLUMNS.length).setValues(rows);
  if (typeof applyTypography === "function") {
    applyTypography(sheet.getRange(firstRow, 1, rows.length, PRESENCE_RANGE_COLUMNS.length));
  }
  return { firstRow: firstRow, count: rows.length };
}

function legacyHeaderMap(sheet) {
  var lastCol = Math.max(sheet.getLastColumn ? sheet.getLastColumn() : 0, 1);
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  header.forEach(function(h, i) { if (h) map[String(h).trim().toLowerCase()] = i; });
  return { map: map, numCols: lastCol };
}

function legacyRowsToRanges(sheet, legacyName) {
  var hm = legacyHeaderMap(sheet);
  var last = (typeof sheet.getLastRow === "function") ? (sheet.getLastRow() || 1) : 1;
  if (last < 2) return [];
  var range = sheet.getRange(2, 1, last - 1, hm.numCols);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  var col = function(row, name) { return hm.map[name] === undefined ? "" : row[hm.map[name]]; };
  var out = [];

  for (var r = 0; r < vals.length; r++) {
    var row = vals[r];
    var memberIdx = hm.map.member;
    var member = cellText(row[memberIdx], (rich && rich[r]) ? rich[r][memberIdx] : null);
    if (!member) continue;
    var item;
    if (legacyName === "tbl_leaves") {
      var leaveType = cellText(col(row, "leave_type")) || "Vacation";
      var status = cellText(col(row, "status")) || "Approved";
      item = {
        range_id: cellText(col(row, "leave_id")) || generateRangeId("Leave"),
        member: member, range_type: "Leave",
        start_date: toIsoDateString(col(row, "start_date")),
        start_period: String(col(row, "start_period") || "AM").trim().toUpperCase(),
        end_date: toIsoDateString(col(row, "end_date")),
        end_period: String(col(row, "end_period") || "PM").trim().toUpperCase(),
        day_of_week: "All", location_am: leaveType, location_pm: leaveType, status: status
      };
    } else {
      item = {
        range_id: cellText(col(row, "baseline_id")) || generateRangeId("Baseline"),
        member: member, range_type: "Baseline",
        start_date: toIsoDateString(col(row, "start_date")),
        start_period: "", end_date: toIsoDateString(col(row, "end_date")), end_period: "",
        day_of_week: cellText(col(row, "day_of_week")) || "All",
        location_am: cellText(col(row, "location_am")), location_pm: cellText(col(row, "location_pm")),
        status: "Approved"
      };
    }
    out.push({ item: item, sourceRow: r + 2, memberCol: memberIdx + 1 });
  }
  return out;
}

/**
 * One-time, idempotent migration of the legacy tbl_presence_baselines and
 * tbl_leaves tables into tbl_presence_ranges. Old tables are renamed
 * archive_<name> and hidden, never deleted.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {{migratedCount: number, archived: string[]}}
 */
function migrateLegacyPresenceTables(ss) {
  var res = { migratedCount: 0, archived: [] };
  var legacy = LEGACY_PRESENCE_SHEETS.filter(function(n) { return !!ss.getSheetByName(n); });
  if (legacy.length === 0) return res;
  if (!ss.getSheetByName(PRESENCE_RANGES_SHEET) && typeof freshInit === "function") freshInit(ss);

  var target = ss.getSheetByName(PRESENCE_RANGES_SHEET);
  var memberCol = presenceRangeColumn("member");
  for (var i = 0; i < legacy.length; i++) {
    var name = legacy[i];
    var src = ss.getSheetByName(name);
    var mapped = legacyRowsToRanges(src, name);
    if (mapped.length > 0) {
      var written = appendPresenceRanges(ss, mapped.map(function(m) { return m.item; }));
      if (typeof copyCellChip === "function") {
        for (var k = 0; k < mapped.length; k++) {
          try {
            copyCellChip(src, mapped[k].sourceRow, mapped[k].memberCol, target, written.firstRow + k, memberCol,
              mapped[k].item.member, ss.getId ? ss.getId() : "");
          } catch (eChip) {}
        }
      }
      res.migratedCount += mapped.length;
    }
    var archiveName = "archive_" + name;
    var suffix = 2;
    while (ss.getSheetByName(archiveName)) archiveName = "archive_" + name + "_" + (suffix++);
    src.setName(archiveName);
    try { src.hideSheet(); } catch (eHide) {}
    res.archived.push(archiveName);
  }
  return res;
}

/**
 * Hides the presence data sheets (tables marked hiddenByDefault, the fact table
 * and calc_kpis) so Members see the Weekly Matrix View instead. Hidden, not
 * locked: anyone can unhide them.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {string[]} Names of sheets hidden by this call
 */
function hidePresenceDataSheets(ss) {
  var names = PRESENCE_HIDDEN_EXTRA_SHEETS.slice();
  if (typeof SCHEMA_SPEC !== "undefined") {
    for (var key in SCHEMA_SPEC) {
      if (SCHEMA_SPEC.hasOwnProperty(key) && SCHEMA_SPEC[key].hiddenByDefault) names.push(SCHEMA_SPEC[key].sheetName);
    }
  }
  var hidden = [];
  names.forEach(function(n) {
    var sheet = ss.getSheetByName(n);
    if (sheet && typeof sheet.isSheetHidden === "function" && !sheet.isSheetHidden()) {
      try { sheet.hideSheet(); hidden.push(n); } catch (e) {}
    }
  });
  return hidden;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    PRESENCE_RANGES_SHEET: PRESENCE_RANGES_SHEET,
    PRESENCE_RANGE_COLUMNS: PRESENCE_RANGE_COLUMNS,
    presenceRangeColumn: presenceRangeColumn,
    toIsoDateString: toIsoDateString,
    readPresenceRanges: readPresenceRanges,
    readBaselinesByMember: readBaselinesByMember,
    readApprovedLeavesByMember: readApprovedLeavesByMember,
    appendPresenceRanges: appendPresenceRanges,
    migrateLegacyPresenceTables: migrateLegacyPresenceTables,
    hidePresenceDataSheets: hidePresenceDataSheets
  };
}
