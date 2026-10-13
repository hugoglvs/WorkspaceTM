/**
 * Contracts (ADR 0048): a person starts in Members; their contracts are
 * declared in the Contracts sidebar (menu "Declare Contract", or opened by a
 * new name typed in Members), which lists them, adds new ones and edits old
 * ones. The Contracts tab stays hidden. Role, Department and Status in Members
 * are formulas reading Contracts (ADR 0047).
 */

function ctName(val, rich) {
  return (typeof extractCellValue === "function") ? extractCellValue(val, rich) : String(val || "").trim();
}

var CT_COLS = { member: 1, role: 2, department: 3, start_date: 4, end_date: 5, contract_type: 6, fte_rate: 7 };

function ctIso(val) {
  if (val === "" || val === null || val === undefined) return "";
  if (val instanceof Date && !isNaN(val.getTime())) {
    var m = String(val.getMonth() + 1), d = String(val.getDate());
    return val.getFullYear() + "-" + (m.length < 2 ? "0" + m : m) + "-" + (d.length < 2 ? "0" + d : d);
  }
  var str = String(val).trim();
  var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[0];
  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? "" : ctIso(parsed);
}

function ctShift(iso, days) {
  var p = iso.split("-");
  var d = new Date(+p[0], +p[1] - 1, +p[2]);
  d.setDate(d.getDate() + days);
  return ctIso(d);
}

function ctFtePercent(raw) {
  var n = parseFloat(String(raw === undefined ? "" : raw).replace("%", "").replace(",", "."));
  if (isNaN(n)) return NaN;
  return n <= 1 ? Math.round(n * 1000) / 10 : n;
}

/** Members rows with a name: [{row, name, key, status}]. */
function ctMembers(ss) {
  var sheet = getTable(ss, "members");
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 500) - 1, SCHEMA_SPEC.members.columns.length);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  var statusIdx = SCHEMA_SPEC.members.columns.indexOf("tenure_status");
  var out = [];
  for (var r = 0; r < vals.length; r++) {
    var name = ctName(vals[r][0], rich[r] ? rich[r][0] : null);
    if (name) out.push({ row: r + 2, name: name, key: name.toLowerCase(), status: String(vals[r][statusIdx] || "").trim() });
  }
  return out;
}

/** Contract rows: [{row, key, role, department, startDate, endDate, contractType, ftePercent}] (ISO dates). */
function ctContracts(ss) {
  var sheet = getTable(ss, "tenures");
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 500) - 1, 7);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  var out = [];
  for (var r = 0; r < vals.length; r++) {
    var name = ctName(vals[r][0], rich[r] ? rich[r][0] : null);
    if (!name) continue;
    out.push({
      row: r + 2, key: name.toLowerCase(),
      role: String(vals[r][1] || "").trim(), department: String(vals[r][2] || "").trim(),
      startDate: ctIso(vals[r][3]), endDate: ctIso(vals[r][4]),
      contractType: String(vals[r][5] || "").trim(), ftePercent: ctFtePercent(vals[r][6])
    });
  }
  return out;
}

/** True when `name` has at least one row in Contracts. */
function memberHasContract(ss, name) {
  var key = String(name || "").trim().toLowerCase();
  return !!key && ctContracts(ss).some(function(c) { return c.key === key; });
}

function ctContractTypes() {
  var spec = SCHEMA_SPEC.tenures && SCHEMA_SPEC.tenures.validations ? SCHEMA_SPEC.tenures.validations.contract_type : null;
  return spec || ["Permanent", "Contractor", "Apprentice", "Intern"];
}

/**
 * Data for the Contracts sidebar: every Member with their contracts (oldest
 * first), plus the known roles, departments and contract types.
 */
function getContractsSidebarData(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var contracts = ctContracts(ss);
  var roles = {}, departments = {};
  contracts.forEach(function(c) {
    if (c.role) roles[c.role] = true;
    if (c.department) departments[c.department] = true;
  });
  return {
    members: ctMembers(ss).map(function(m) {
      return {
        name: m.name, status: m.status,
        contracts: contracts.filter(function(c) { return c.key === m.key; })
          .sort(function(a, b) { return a.startDate < b.startDate ? -1 : (a.startDate > b.startDate ? 1 : 0); })
          .map(function(c) {
            return { row: c.row, role: c.role, department: c.department, startDate: c.startDate, endDate: c.endDate,
              contractType: c.contractType, ftePercent: c.ftePercent };
          })
      };
    }),
    roles: Object.keys(roles).sort(),
    departments: Object.keys(departments).sort(),
    contractTypes: ctContractTypes(),
    today: ctIso(new Date())
  };
}

function ctOverlaps(a, b) {
  var aEnd = a.endDate || "9999-12-31", bEnd = b.endDate || "9999-12-31";
  return a.startDate <= bEnd && b.startDate <= aEnd;
}

/**
 * Adds a contract (payload.row empty) or edits one (payload.row = its Contracts
 * row). A new contract ends the current one on payload.previousEndDate, by
 * default the day before it starts. Contracts of one person may not overlap.
 *
 * @param {{member: string, row: (number|undefined), role: string, department: string, startDate: string,
 *          endDate: (string|undefined), contractType: string, ftePercent: (number|string),
 *          previousEndDate: (string|undefined)}} payload
 * @return {{success: boolean, message: string}}
 */
function saveContract(payload, targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var p = payload || {};
  var key = String(p.member || "").trim().toLowerCase();
  var c = {
    role: String(p.role || "").trim(), department: String(p.department || "").trim(),
    startDate: ctIso(p.startDate), endDate: ctIso(p.endDate),
    contractType: String(p.contractType || "").trim(), ftePercent: ctFtePercent(p.ftePercent)
  };
  if (!key || !c.role || !c.department || !c.startDate || !c.contractType) {
    return { success: false, message: "Role, department, start date and contract type are required." };
  }
  if (isNaN(c.ftePercent) || c.ftePercent <= 0 || c.ftePercent > 100) return { success: false, message: "FTE must be between 1 and 100 %." };
  if (c.endDate && c.endDate < c.startDate) return { success: false, message: "The end date cannot be before the start date." };

  var member = ctMembers(ss).filter(function(m) { return m.key === key; })[0];
  if (!member) return { success: false, message: "Add the person in Members first." };
  var sheet = getTable(ss, "tenures");
  if (!sheet) return { success: false, message: "The Contracts tab was not found. Run Initialize / Upgrade first." };

  var editRow = parseInt(p.row, 10) || 0;
  var others = ctContracts(ss).filter(function(x) { return x.key === key && x.row !== editRow; });
  if (editRow && !ctContracts(ss).some(function(x) { return x.key === key && x.row === editRow; })) {
    return { success: false, message: "This contract was not found any more. Reopen the sidebar." };
  }
  var describe = function(x) { return x.role + " from " + x.startDate + (x.endDate ? " to " + x.endDate : ""); };
  var toClose = [];
  var previousEnd = "";

  if (!editRow) {
    var later = others.filter(function(x) { return x.startDate >= c.startDate; })[0];
    if (later) return { success: false, message: member.name + " already has a contract starting " + later.startDate + ". Edit that one, or choose a later start date." };
    toClose = others.filter(function(x) { return !x.endDate || x.endDate >= c.startDate; });
    if (toClose.length) {
      previousEnd = ctIso(p.previousEndDate) || ctShift(c.startDate, -1);
      if (previousEnd >= c.startDate) return { success: false, message: "The current contract must end before the new one starts (" + c.startDate + ")." };
      for (var i = 0; i < toClose.length; i++) {
        if (previousEnd < toClose[i].startDate) return { success: false, message: "The current contract cannot end before it started (" + toClose[i].startDate + ")." };
      }
    }
  } else {
    var clash = others.filter(function(x) { return ctOverlaps(c, x); })[0];
    if (clash) return { success: false, message: "This overlaps the contract " + describe(clash) + ". Change the dates of one of them." };
  }

  toClose.forEach(function(x) { sheet.getRange(x.row, CT_COLS.end_date).setValue(previousEnd); });
  var row = editRow || ((typeof findNextAvailableRow === "function") ? findNextAvailableRow(sheet) : sheet.getLastRow() + 1);
  sheet.getRange(row, 2, 1, 6).setValues([[c.role, c.department, c.startDate, c.endDate, c.contractType, c.ftePercent]]);
  if (!editRow) {
    // The name cell is copied from Members, so it is the same People chip
    var membersSheet = getTable(ss, "members");
    if (typeof copyCellChip === "function") {
      try { copyCellChip(membersSheet, member.row, 1, sheet, row, 1, member.name, ss.getId ? ss.getId() : ""); } catch (eChip) {}
    }
    if (!ctName(sheet.getRange(row, 1).getValue(), null)) sheet.getRange(row, 1).setValue(member.name);
    if (typeof applyTypography === "function") applyTypography(sheet.getRange(row, 1, 1, 7));
  }

  var message = editRow ? "Contract updated." : "Contract saved for " + member.name + "." +
    (toClose.length ? " The previous contract now ends " + previousEnd + "." : "");
  var all = others.map(function(x) { return toClose.indexOf(x) !== -1 ? { startDate: x.startDate, endDate: previousEnd } : x; }).concat([c])
    .sort(function(a, b) { return a.startDate < b.startDate ? -1 : 1; });
  for (var g = 1; g < all.length; g++) {
    var prevEnd = all[g - 1].endDate;
    if (prevEnd && prevEnd < ctShift(all[g].startDate, -1)) {
      message += " Note: no contract between " + ctShift(prevEnd, 1) + " and " + ctShift(all[g].startDate, -1) +
        ", so " + member.name + " shows as Departed then and counts no capacity.";
    }
  }
  return { success: true, message: message };
}

/**
 * Opens the Contracts sidebar, optionally on one Member.
 *
 * @param {string} [prefillMember]
 */
function showContractsSidebar(prefillMember) {
  var template = HtmlService.createTemplateFromFile("ContractsSidebar");
  template.prefillMember = (typeof prefillMember === "string") ? prefillMember : "";
  SpreadsheetApp.getUi().showSidebar(template.evaluate().setTitle("Contracts").setWidth(350));
}

/**
 * onEdit hook for Members: a name typed in the member column without any
 * contract opens the Contracts sidebar on that person; if the sidebar cannot
 * open (e.g. for someone else than the trigger owner), a toast points to the menu.
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {string|null} "sidebar", "toast" or null
 */
function promptContractForNewMember(e) {
  if (!e || !e.range || e.range.getColumn() !== 1 || e.range.getRow() < 2) return null;
  var sheet = e.range.getSheet();
  var ss = e.source || (sheet.getParent ? sheet.getParent() : null) || SpreadsheetApp.getActiveSpreadsheet();
  var name = ctName(e.range.getValue(), typeof e.range.getRichTextValue === "function" ? e.range.getRichTextValue() : null);
  if (!name || memberHasContract(ss, name)) return null;
  try {
    showContractsSidebar(name);
    return "sidebar";
  } catch (eUi) {
    try { ss.toast(name + " has no contract yet: add it with Project Management → Declare Contract.", "New member", 10); } catch (eToast) {}
    return "toast";
  }
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
    getContractsSidebarData: getContractsSidebarData,
    saveContract: saveContract,
    showContractsSidebar: showContractsSidebar,
    promptContractForNewMember: promptContractForNewMember
  };
}
