/**
 * Contracts (ADR 0046): contracts are declared in the Contracts tab itself,
 * like Locations and Holidays. The menu item "Declare Contract" and a new name
 * typed in Members unhide and open that tab; the sync on edit then updates
 * Members (status, role, department) as before.
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

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    memberHasContract: memberHasContract,
    openContractsTab: openContractsTab,
    showContractsTab: showContractsTab,
    promptContractForNewMember: promptContractForNewMember
  };
}
