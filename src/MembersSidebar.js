/**
 * Members & Contracts sidebar backend (ADR 0037, issue 22).
 * Declares new Members and Tenure changes (new contract, end date) into
 * tbl_tenures, then lets syncMembersAndTenures update tbl_members.
 * Grid edits stay possible: this is the guided path, not the only one.
 */

var TENURE_COLS = { member: 1, role: 2, department: 3, start_date: 4, end_date: 5, contract_type: 6, fte_rate: 7 };
var MEMBER_COLS = { member: 1, email: 2, phone: 3, role: 4, department: 5, status: 6 };

function msIsoDate(val) {
  if (val === "" || val === null || val === undefined) return "";
  if (val instanceof Date && !isNaN(val.getTime())) {
    var m = String(val.getMonth() + 1), d = String(val.getDate());
    return val.getFullYear() + "-" + (m.length < 2 ? "0" + m : m) + "-" + (d.length < 2 ? "0" + d : d);
  }
  var str = String(val).trim();
  var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[0];
  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? "" : msIsoDate(parsed);
}

function msDayBefore(iso) {
  var p = iso.split("-");
  var d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
  d.setDate(d.getDate() - 1);
  return msIsoDate(d);
}

function msName(val, rich) {
  return (typeof extractCellValue === "function") ? extractCellValue(val, rich) : String(val || "").trim();
}

function msReadRows(sheet, numCols) {
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var range = sheet.getRange(2, 1, Math.min(sheet.getMaxRows(), 500) - 1, numCols);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  var out = [];
  for (var r = 0; r < vals.length; r++) {
    var name = msName(vals[r][0], rich[r] ? rich[r][0] : null);
    if (name) out.push({ row: r + 2, name: name, key: name.toLowerCase(), values: vals[r] });
  }
  return out;
}

function msTenuresOf(ss, memberName) {
  var key = String(memberName || "").trim().toLowerCase();
  return msReadRows(ss.getSheetByName("tbl_tenures"), 7).filter(function(t) { return t.key === key; });
}

function msContractTypes() {
  var spec = (typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC.tbl_tenures && SCHEMA_SPEC.tbl_tenures.validations) ? SCHEMA_SPEC.tbl_tenures.validations.contract_type : null;
  return spec || ["Permanent", "Contractor", "Apprentice", "Intern"];
}

function msToFtePercent(raw) {
  var n = parseFloat(String(raw === undefined ? "" : raw).replace("%", "").replace(",", "."));
  if (isNaN(n)) return NaN;
  return n <= 1 ? Math.round(n * 1000) / 10 : n;
}

function msSync(ss) {
  if (typeof syncMembersAndTenures === "function") {
    try { syncMembersAndTenures(ss); } catch (eSync) { Logger.log("Members sidebar sync note: " + eSync); }
  }
}

/**
 * Data for the Members & Contracts sidebar.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {{members: Object[], roles: string[], departments: string[], contractTypes: string[], prefillMember: string}}
 */
function getMembersSidebarData(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var tenureRows = msReadRows(ss.getSheetByName("tbl_tenures"), 7);
  var roles = {}, departments = {};
  var byMember = {};

  tenureRows.forEach(function(t) {
    var v = t.values;
    var tenure = {
      startDate: msIsoDate(v[3]), endDate: msIsoDate(v[4]),
      role: String(v[1] || "").trim(), department: String(v[2] || "").trim(),
      contractType: String(v[5] || "").trim(), ftePercent: msToFtePercent(v[6])
    };
    if (tenure.role) roles[tenure.role] = true;
    if (tenure.department) departments[tenure.department] = true;
    (byMember[t.key] = byMember[t.key] || []).push(tenure);
  });

  var members = msReadRows(ss.getSheetByName("tbl_members"), 7).map(function(m) {
    var list = (byMember[m.key] || []).sort(function(a, b) { return a.startDate < b.startDate ? -1 : 1; });
    return {
      name: m.name,
      email: String(m.values[1] || "").trim(),
      phone: String(m.values[2] || "").trim(),
      role: String(m.values[3] || "").trim(),
      department: String(m.values[4] || "").trim(),
      status: String(m.values[5] || "").trim(),
      tenures: list
    };
  });

  return {
    members: members,
    roles: Object.keys(roles).sort(),
    departments: Object.keys(departments).sort(),
    contractTypes: msContractTypes(),
    prefillMember: ""
  };
}

/**
 * Declares a Tenure: a new Member's first contract, or a new contract for an
 * existing Member (role, department, FTE change). An open or overlapping earlier
 * Tenure is ended the day before the new one starts.
 *
 * @param {{member: string, role: string, department: string, startDate: string, endDate: (string|undefined),
 *          contractType: string, ftePercent: number, email: (string|undefined), phone: (string|undefined)}} payload
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {{success: boolean, message: string, closedPrevious: (boolean|undefined)}}
 */
function saveMemberTenure(payload, targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var p = payload || {};
  var member = String(p.member || "").trim();
  var role = String(p.role || "").trim();
  var department = String(p.department || "").trim();
  var startDate = msIsoDate(p.startDate);
  var endDate = msIsoDate(p.endDate);
  var contractType = String(p.contractType || "").trim();
  var fte = msToFtePercent(p.ftePercent);

  if (!member || !role || !department || !startDate || !contractType) {
    return { success: false, message: "Member, role, department, start date and contract type are required." };
  }
  if (isNaN(fte) || fte <= 0 || fte > 100) {
    return { success: false, message: "FTE must be between 1 and 100 %." };
  }
  if (endDate && endDate < startDate) {
    return { success: false, message: "The end date cannot be before the start date." };
  }

  var tenuresSheet = ss.getSheetByName("tbl_tenures");
  if (!tenuresSheet) return { success: false, message: "tbl_tenures not found. Run Initialize / Upgrade first." };

  var existing = msTenuresOf(ss, member);
  for (var i = 0; i < existing.length; i++) {
    var exStart = msIsoDate(existing[i].values[3]);
    if (exStart && exStart >= startDate) {
      return { success: false, message: member + " already has a contract starting " + exStart + ". Choose a later start date, or edit that contract." };
    }
  }

  var closedPrevious = false;
  for (var j = 0; j < existing.length; j++) {
    var exEnd = msIsoDate(existing[j].values[4]);
    if (!exEnd || exEnd >= startDate) {
      tenuresSheet.getRange(existing[j].row, TENURE_COLS.end_date).setValue(msDayBefore(startDate));
      closedPrevious = true;
    }
  }

  var targetRow = (typeof findNextAvailableRow === "function") ? findNextAvailableRow(tenuresSheet) : tenuresSheet.getLastRow() + 1;
  tenuresSheet.getRange(targetRow, 1, 1, 7).setValues([[member, role, department, startDate, endDate, contractType, fte]]);

  var membersSheet = ss.getSheetByName("tbl_members");
  var memberRow = msReadRows(membersSheet, 7).filter(function(m) { return m.key === member.toLowerCase(); })[0];
  if (memberRow && typeof copyCellChip === "function") {
    try { copyCellChip(membersSheet, memberRow.row, 1, tenuresSheet, targetRow, 1, member, ss.getId ? ss.getId() : ""); } catch (eChip) {}
  }
  if (typeof applyTypography === "function") applyTypography(tenuresSheet.getRange(targetRow, 1, 1, 7));

  msSync(ss);
  if (p.email || p.phone) saveMemberProfile({ member: member, email: p.email, phone: p.phone }, ss);

  return {
    success: true,
    closedPrevious: closedPrevious,
    message: closedPrevious
      ? "Contract saved. The previous contract now ends " + msDayBefore(startDate) + "."
      : "Contract saved for " + member + "."
  };
}

/**
 * Sets the end date of one Tenure (identified by Member and start date).
 *
 * @param {{member: string, startDate: string, endDate: string}} payload
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {{success: boolean, message: string}}
 */
function endMemberTenure(payload, targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var p = payload || {};
  var startDate = msIsoDate(p.startDate);
  var endDate = msIsoDate(p.endDate);
  if (!p.member || !startDate || !endDate) return { success: false, message: "Member, contract start and end date are required." };
  if (endDate < startDate) return { success: false, message: "The end date cannot be before the contract start (" + startDate + ")." };

  var match = msTenuresOf(ss, p.member).filter(function(t) { return msIsoDate(t.values[3]) === startDate; })[0];
  if (!match) return { success: false, message: "No contract found for " + p.member + " starting " + startDate + "." };

  ss.getSheetByName("tbl_tenures").getRange(match.row, TENURE_COLS.end_date).setValue(endDate);
  msSync(ss);
  return { success: true, message: "Contract end date set to " + endDate + "." };
}

/**
 * Updates a Member's contact details in tbl_members.
 *
 * @param {{member: string, email: (string|undefined), phone: (string|undefined)}} payload
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {{success: boolean, message: string}}
 */
function saveMemberProfile(payload, targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var p = payload || {};
  var key = String(p.member || "").trim().toLowerCase();
  var sheet = ss.getSheetByName("tbl_members");
  var row = msReadRows(sheet, 7).filter(function(m) { return m.key === key; })[0];
  if (!row) return { success: false, message: "Member not found in the directory." };
  if (p.email !== undefined) sheet.getRange(row.row, MEMBER_COLS.email).setValue(String(p.email || "").trim());
  if (p.phone !== undefined) sheet.getRange(row.row, MEMBER_COLS.phone).setValue(String(p.phone || "").trim());
  return { success: true, message: "Contact details saved." };
}

/**
 * Opens the Members & Contracts sidebar, optionally pre-filled for one person.
 *
 * @param {string} [prefillMember]
 */
function showMembersSidebar(prefillMember) {
  var template = HtmlService.createTemplateFromFile("MembersSidebar");
  template.prefillMember = (typeof prefillMember === "string") ? prefillMember : "";
  var html = template.evaluate().setTitle("Members & Contracts").setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * onEdit hook for tbl_members: a name typed in the member column without any
 * Tenure opens the sidebar pre-filled for that person. If the sidebar cannot
 * open in this trigger context, a toast points to the menu instead.
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {string|null} "sidebar", "toast" or null
 */
function promptContractForNewMember(e) {
  if (!e || !e.range || e.range.getColumn() !== MEMBER_COLS.member || e.range.getRow() < 2) return null;
  var sheet = e.range.getSheet();
  var ss = e.source || (sheet.getParent ? sheet.getParent() : null) || SpreadsheetApp.getActiveSpreadsheet();
  var name = msName(e.range.getValue(), typeof e.range.getRichTextValue === "function" ? e.range.getRichTextValue() : null);
  if (!name || msTenuresOf(ss, name).length > 0) return null;

  try {
    showMembersSidebar(name);
    return "sidebar";
  } catch (eUi) {
    try {
      ss.toast("Add " + name + "'s contract with Project Management → Members & Contracts.", "New member", 8);
    } catch (eToast) {}
    return "toast";
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getMembersSidebarData: getMembersSidebarData,
    saveMemberTenure: saveMemberTenure,
    endMemberTenure: endMemberTenure,
    saveMemberProfile: saveMemberProfile,
    showMembersSidebar: showMembersSidebar,
    promptContractForNewMember: promptContractForNewMember
  };
}
