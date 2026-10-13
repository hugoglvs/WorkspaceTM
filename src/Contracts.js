/**
 * Contracts (ADR 0046): contracts are declared in the Contracts tab itself,
 * like Locations and Holidays. The menu item "Declare Contract" and a new name
 * typed in Members unhide and open that tab. Role, Department and Status in
 * Members are formulas reading Contracts (ADR 0047).
 */

function ctName(val, rich) {
  return (typeof extractCellValue === "function") ? extractCellValue(val, rich) : String(val || "").trim();
}

/** True when `name` has at least one row in Contracts. */
function memberHasContract(ss, name) {
  var sheet = getTable(ss, "tenures");
  var key = String(name || "").trim().toLowerCase();
  if (!sheet || !key || sheet.getMaxRows() < 2) return false;
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 500) - 1, 1);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  for (var r = 0; r < vals.length; r++) {
    if (ctName(vals[r][0], rich[r] ? rich[r][0] : null).toLowerCase() === key) return true;
  }
  return false;
}

/**
 * Unhides and opens the Contracts tab on its first empty row, with a toast
 * saying what to fill in.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} [forMember] name of a new Member who has no contract yet
 * @return {{success: boolean, message: string, row: (number|undefined)}}
 */
function openContractsTab(ss, forMember) {
  var sheet = getTable(ss, "tenures");
  if (!sheet) return { success: false, message: "The Contracts tab was not found. Run Initialize / Upgrade first." };
  if (sheet.isSheetHidden()) sheet.showSheet();
  var row = (typeof findNextAvailableRow === "function") ? findNextAvailableRow(sheet) : sheet.getLastRow() + 1;
  try {
    sheet.activate();
    sheet.getRange(row, 1).activate();
  } catch (eAct) {}
  var message = (forMember ? forMember + " has no contract yet. " : "") +
    "Fill one row per contract: member, role, department, start date, contract type and FTE. " +
    "A new contract for someone: first set the end date of the current one. Hide the tab again when done.";
  try { ss.toast(message, "Contracts", 12); } catch (eToast) {}
  return { success: true, message: message, row: row };
}

/** Menu: Declare Contract. */
function showContractsTab() {
  return openContractsTab(SpreadsheetApp.getActiveSpreadsheet());
}

/**
 * onEdit hook for Members: a name typed in the member column without any
 * contract opens the Contracts tab.
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {string|null} "contracts" or null
 */
function promptContractForNewMember(e) {
  if (!e || !e.range || e.range.getColumn() !== 1 || e.range.getRow() < 2) return null;
  var sheet = e.range.getSheet();
  var ss = e.source || (sheet.getParent ? sheet.getParent() : null) || SpreadsheetApp.getActiveSpreadsheet();
  var name = ctName(e.range.getValue(), typeof e.range.getRichTextValue === "function" ? e.range.getRichTextValue() : null);
  if (!name || memberHasContract(ss, name)) return null;
  return openContractsTab(ss, name).success ? "contracts" : null;
}

/**
 * Reference rule for Role, Department and Status in Members (ADR 0047). The
 * sheet computes them with the formulas below; this function states the same
 * rule for tests and documentation.
 *   Active: a contract has started (or has no start) and has not ended (end blank or >= today).
 *   Upcoming: no active contract, one starts later. Departed: only past contracts.
 *   Role/Department come from the earliest active contract, else the earliest
 *   upcoming one, else the contract that started last ("Team Member"/"Delivery" if blank).
 *
 * @param {Array<{role: string, department: string, start: (Date|null), end: (Date|null)}>} records
 * @param {Date} today
 */
function contractProfile(records, today) {
  if (!records || !records.length) return { role: "", department: "", status: "No Contract" };
  var t = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
  var at = function(d) { return d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime() : null; };
  var active = records.filter(function(r) { return (at(r.start) === null || at(r.start) <= t) && (at(r.end) === null || at(r.end) >= t); });
  var upcoming = records.filter(function(r) { return at(r.start) !== null && at(r.start) > t; });
  // Like SORT in Sheets: blank start dates go last in both directions
  var byStart = function(list, asc) {
    return list.slice().sort(function(a, b) {
      var x = at(a.start), y = at(b.start);
      if (x === null) return y === null ? 0 : 1;
      if (y === null) return -1;
      return asc ? x - y : y - x;
    })[0];
  };
  var status = active.length ? "Active" : (upcoming.length ? "Upcoming" : "Departed");
  var pick = active.length ? byStart(active, true) : (upcoming.length ? byStart(upcoming, true) : byStart(records, false));
  return { role: pick.role || "Team Member", department: pick.department || "Delivery", status: status };
}

/** Members columns computed by formula from Contracts (ADR 0047). */
var MEMBER_FORMULA_COLUMNS = ["current_role", "current_department", "tenure_status"];

/**
 * The three formulas of one Members row.
 *
 * @param {string} contractsSheetName the Contracts tab's current name
 * @param {number} row Members row
 * @return {Object.<string, string>} keyed by Members column
 */
function memberFormulas(contractsSheetName, row) {
  var q = "'" + String(contractsSheetName).replace(/'/g, "''") + "'!";
  var rng = function(key) {
    var c = String.fromCharCode(65 + SCHEMA_SPEC.tenures.columns.indexOf(key));
    return q + "$" + c + "$2:$" + c;
  };
  var a = "$A" + row;
  var base = "k," + rng("member") + "=" + a + ",s," + rng("start_date") + ",e," + rng("end_date") +
    ',a,k*((s="")+(s<=TODAY()))*((e="")+(e>=TODAY())),u,k*(s>TODAY())';
  var status = "=IF(" + a + '="","",ARRAYFORMULA(LET(' + base +
    ',IF(SUM(k*1)=0,"No Contract",IF(SUM(a)>0,"Active",IF(SUM(u)>0,"Upcoming","Departed"))))))';
  var pick = function(key, fallback) {
    return "=IF(" + a + '="","",ARRAYFORMULA(LET(' + base + ",v," + rng(key) +
      ",p,IF(SUM(a)>0,a,IF(SUM(u)>0,u,k*1))," +
      'IF(SUM(p)=0,"",LET(x,SORT(FILTER(HSTACK(v,s),p>0),2,SUM(a)+SUM(u)>0),y,INDEX(x,1,1),IF(y="","' + fallback + '",y))))))';
  };
  return { current_role: pick("role", "Team Member"), current_department: pick("department", "Delivery"), tenure_status: status };
}

/** Last row of the Members table (its native table range), at least the last row with content. */
function memberTableEndRow(ss, sheet) {
  var end = Math.max(sheet.getLastRow(), 2);
  try {
    var listFn = (typeof listSheetTables === "function") ? listSheetTables
      : (typeof require !== "undefined" ? require("./Typography").listSheetTables : null);
    if (listFn && ss.getId) {
      listFn(ss.getId()).forEach(function(t) {
        if (t.sheetId === sheet.getSheetId() && t.range && t.range.endRowIndex) end = Math.max(end, t.range.endRowIndex);
      });
    }
  } catch (eTbl) {}
  return Math.min(end, sheet.getMaxRows());
}

/**
 * Writes the Role, Department and Status formulas into every row of the
 * Members table, only where a cell does not already hold the right formula
 * (typed over, new row, renamed Contracts tab). Run by the members sync.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {{rows: number, written: number}} rows checked, cells rewritten
 */
function applyMemberFormulas(ss) {
  var members = getTable(ss, "members");
  var contracts = getTable(ss, "tenures");
  if (!members || !contracts) return { rows: 0, written: 0 };
  var cols = SCHEMA_SPEC.members.columns;
  var first = cols.indexOf(MEMBER_FORMULA_COLUMNS[0]) + 1;
  var width = MEMBER_FORMULA_COLUMNS.length;
  var end = memberTableEndRow(ss, members);
  if (end < 2) return { rows: 0, written: 0 };
  var range = members.getRange(2, first, end - 1, width);
  var current = range.getFormulas();
  var wanted = [];
  var written = 0;
  for (var r = 0; r < end - 1; r++) {
    var f = memberFormulas(contracts.getName(), r + 2);
    var row = MEMBER_FORMULA_COLUMNS.map(function(c) { return f[c]; });
    for (var c = 0; c < width; c++) if ((current[r] && current[r][c]) !== row[c]) written++;
    wanted.push(row);
  }
  if (written) range.setFormulas(wanted);
  return { rows: end - 1, written: written };
}

/**
 * Members rows where someone typed over a Role, Department or Status formula.
 *
 * @return {string[]} e.g. "Alice: Status"
 */
function findTypedOverMemberCells(ss) {
  var members = getTable(ss, "members");
  if (!members || members.getLastRow() < 2) return [];
  var cols = SCHEMA_SPEC.members.columns;
  var first = cols.indexOf(MEMBER_FORMULA_COLUMNS[0]) + 1;
  var n = members.getLastRow() - 1;
  var names = members.getRange(2, 1, n, 1);
  var nameVals = names.getValues();
  var nameRich = (typeof names.getRichTextValues === "function") ? names.getRichTextValues() : [];
  var formulas = members.getRange(2, first, n, MEMBER_FORMULA_COLUMNS.length).getFormulas();
  var out = [];
  for (var r = 0; r < n; r++) {
    var name = ctName(nameVals[r][0], nameRich[r] ? nameRich[r][0] : null);
    if (!name) continue;
    var typed = MEMBER_FORMULA_COLUMNS.filter(function(c, i) { return !formulas[r][i]; })
      .map(function(c) { return SCHEMA_SPEC.members.headers[c] || c; });
    if (typed.length) out.push(name + ": " + typed.join(", "));
  }
  return out;
}

/** Check Workspace action: puts the Members formulas back. */
function restoreMemberFormulas(targetSpreadsheet) {
  var res = applyMemberFormulas(targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet());
  return { success: true, message: res.written ? res.written + " cell(s) restored." : "Nothing to restore." };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    memberHasContract: memberHasContract,
    contractProfile: contractProfile,
    memberFormulas: memberFormulas,
    MEMBER_FORMULA_COLUMNS: MEMBER_FORMULA_COLUMNS,
    applyMemberFormulas: applyMemberFormulas,
    findTypedOverMemberCells: findTypedOverMemberCells,
    restoreMemberFormulas: restoreMemberFormulas,
    openContractsTab: openContractsTab,
    showContractsTab: showContractsTab,
    promptContractForNewMember: promptContractForNewMember
  };
}
