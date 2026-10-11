/**
 * Workspace layout (ADR 0036): which tabs Members see, plus the one-time
 * restore of the short-lived merged presence table (ADR 0035, reverted).
 */

var VISIBLE_WORKSPACE_TABLES = ["members", "weeklyMatrix", "projects", "tasks", "events"];
var MANAGED_SHEET_PREFIXES = ["tbl_", "view_", "fact_", "calc_", "archive_", "Archive – "];
var RETIRED_SHEETS = ["calc_kpis", "view_active_responsibles"];

function isManagedSheet(sheet) {
  if (typeof tableKeyOfSheet === "function" && tableKeyOfSheet(sheet)) return true;
  return isManagedSheetName(sheet.getName());
}

function isManagedSheetName(name) {
  for (var i = 0; i < MANAGED_SHEET_PREFIXES.length; i++) {
    if (name.indexOf(MANAGED_SHEET_PREFIXES[i]) === 0) return true;
  }
  return false;
}

function isSheetEmpty(sheet) {
  var lastRow = (typeof sheet.getLastRow === "function") ? sheet.getLastRow() : 0;
  var lastCol = (typeof sheet.getLastColumn === "function") ? sheet.getLastColumn() : 0;
  return !lastRow || !lastCol;
}

/**
 * Shows the five working tabs and hides every other Library-managed sheet
 * (registered tables, archives and former technical names; hidden, not locked). Deletes retired sheets (calc_kpis)
 * and an empty default "Sheet1". Tabs the team created themselves are left alone.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {{hidden: string[], shown: string[], deleted: string[]}}
 */
function applyWorkspaceLayout(ss) {
  var res = { hidden: [], shown: [], deleted: [] };

  for (var r = 0; r < RETIRED_SHEETS.length; r++) {
    var retired = ss.getSheetByName(RETIRED_SHEETS[r]);
    if (retired) { ss.deleteSheet(retired); res.deleted.push(RETIRED_SHEETS[r]); }
  }

  // Show first, so hiding never leaves the spreadsheet without a visible tab
  var visibleIds = {};
  for (var v = 0; v < VISIBLE_WORKSPACE_TABLES.length; v++) {
    var keep = getTable(ss, VISIBLE_WORKSPACE_TABLES[v]);
    if (!keep) continue;
    visibleIds[keep.getSheetId()] = true;
    if (keep.isSheetHidden()) { keep.showSheet(); res.shown.push(keep.getName()); }
  }

  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var name = sheets[s].getName();
    if (visibleIds[sheets[s].getSheetId()]) continue;
    if (name === "Sheet1" && isSheetEmpty(sheets[s]) && ss.getSheets().length > 1) {
      ss.deleteSheet(sheets[s]);
      res.deleted.push(name);
      continue;
    }
    if (isManagedSheet(sheets[s]) && !sheets[s].isSheetHidden()) {
      try { sheets[s].hideSheet(); res.hidden.push(name); } catch (eHide) {}
    }
  }
  return res;
}

function presenceIso(val) {
  if (val === "" || val === null || val === undefined) return "";
  if (val instanceof Date && !isNaN(val.getTime())) {
    var m = String(val.getMonth() + 1), d = String(val.getDate());
    return val.getFullYear() + "-" + (m.length < 2 ? "0" + m : m) + "-" + (d.length < 2 ? "0" + d : d);
  }
  var str = String(val).trim();
  var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[0];
  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? str : presenceIso(parsed);
}

function appendRowsAfterLast(sheet, rows) {
  var last = (typeof sheet.getLastRow === "function") ? (sheet.getLastRow() || 1) : 1;
  var first = 2;
  if (last >= 2) {
    var vals = sheet.getRange(2, 1, last - 1, 2).getValues();
    for (var i = vals.length - 1; i >= 0; i--) { if (vals[i][0] || vals[i][1]) { first = i + 3; break; } }
  }
  var needed = first + rows.length - 1;
  if (sheet.getMaxRows() < needed && typeof sheet.insertRowsAfter === "function") {
    sheet.insertRowsAfter(sheet.getMaxRows(), needed - sheet.getMaxRows());
  }
  sheet.getRange(first, 1, rows.length, rows[0].length).setValues(rows);
  return first;
}

function existingIds(sheet) {
  var ids = {};
  var last = (typeof sheet.getLastRow === "function") ? (sheet.getLastRow() || 1) : 1;
  if (last < 2) return ids;
  var vals = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) if (vals[i][0]) ids[String(vals[i][0]).trim()] = true;
  return ids;
}

/**
 * Moves rows of the merged tbl_presence_ranges (ADR 0035, reverted) back into
 * tbl_presence_baselines and tbl_leaves, keeping People Chips, then renames the
 * merged table archive_tbl_presence_ranges and hides it. Idempotent; rows whose
 * ID already exists in the target table are skipped.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {{restoredCount: number}}
 */
function restorePresenceTablesFromRanges(ss) {
  var res = { restoredCount: 0 };
  var src = ss.getSheetByName("tbl_presence_ranges");
  if (!src) return res;
  if ((!getTable(ss, "baselines") || !getTable(ss, "leaves")) && typeof freshInit === "function") {
    freshInit(ss);
  }
  var baselines = getTable(ss, "baselines");
  var leaves = getTable(ss, "leaves");

  var last = (typeof src.getLastRow === "function") ? (src.getLastRow() || 1) : 1;
  if (last >= 2) {
    var range = src.getRange(2, 1, last - 1, 11);
    var vals = range.getValues();
    var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
    var baseIds = existingIds(baselines), leaveIds = existingIds(leaves);
    var baseRows = [], baseSrc = [], leaveRows = [], leaveSrc = [];

    for (var r = 0; r < vals.length; r++) {
      var v = vals[r];
      var member = (typeof extractCellValue === "function") ? extractCellValue(v[1], rich[r] ? rich[r][1] : null) : String(v[1] || "").trim();
      if (!member) continue;
      var id = String(v[0] || "").trim();
      if (String(v[2]).trim() === "Leave") {
        if (id && leaveIds[id]) continue;
        leaveRows.push([id, member, presenceIso(v[3]), String(v[4] || "AM"), presenceIso(v[5]), String(v[6] || "PM"), v[8] || v[9] || "Vacation", v[10] || "Approved"]);
        leaveSrc.push(r + 2);
      } else {
        if (id && baseIds[id]) continue;
        baseRows.push([id, member, presenceIso(v[3]), presenceIso(v[5]), v[7], v[8], v[9]]);
        baseSrc.push(r + 2);
      }
    }

    var copyChips = function(target, firstRow, srcRows, memberName) {
      if (typeof copyCellChip !== "function") return;
      for (var k = 0; k < srcRows.length; k++) {
        try { copyCellChip(src, srcRows[k], 2, target, firstRow + k, 2, memberName(k), ss.getId ? ss.getId() : ""); } catch (eChip) {}
      }
    };
    if (baseRows.length) {
      var bFirst = appendRowsAfterLast(baselines, baseRows);
      copyChips(baselines, bFirst, baseSrc, function(k) { return baseRows[k][1]; });
    }
    if (leaveRows.length) {
      var lFirst = appendRowsAfterLast(leaves, leaveRows);
      copyChips(leaves, lFirst, leaveSrc, function(k) { return leaveRows[k][1]; });
    }
    res.restoredCount = baseRows.length + leaveRows.length;
  }

  var archiveName = "Archive – Presence Ranges", n = 2;
  while (ss.getSheetByName(archiveName)) archiveName = "Archive – Presence Ranges " + (n++);
  src.setName(archiveName);
  try { src.hideSheet(); } catch (eHide) {}
  return res;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    VISIBLE_WORKSPACE_TABLES: VISIBLE_WORKSPACE_TABLES,
    applyWorkspaceLayout: applyWorkspaceLayout,
    restorePresenceTablesFromRanges: restorePresenceTablesFromRanges
  };
}
