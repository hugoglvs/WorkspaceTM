/**
 * Dropdowns (ADR 0054): the hidden "Dropdowns" tab holds one model cell per
 * list. Code cannot choose the chip style or the option colours (no API), so a
 * person sets them once by hand on the model cell (Data → Data validation);
 * the code then copies that dropdown onto the cells that use the list, on
 * every upgrade and every rebuild of a view.
 *
 * Each model cell is a "Dropdown (from a range)" pointing at the list's
 * options column in the same tab. The code keeps that column up to date (new
 * locations, statuses…) and never rewrites a model's dropdown once it exists,
 * so the chosen style and colours survive. A new option starts without colour.
 *
 * Tab layout: A List | B Dropdown (the model cell) | C Used in, then one
 * options column per list from column E (header = list name, options below).
 */

var DROPDOWN_FIRST_OPTIONS_COL = 5;

/** A function by name: the Apps Script global, else (Node tests) from its module. */
function ddFn(name, modulePath) {
  var g = (typeof globalThis !== "undefined") ? globalThis : this;
  if (g && typeof g[name] === "function") return g[name];
  if (typeof require !== "undefined" && modulePath) {
    try { var mod = require(modulePath); if (mod && typeof mod[name] === "function") return mod[name]; } catch (e) {}
  }
  return null;
}

/**
 * The lists, in tab order. Views apply their model when they rebuild; a list
 * with {table, column} would be copied onto that table column by
 * applyTableDropdowns, but tables ignore copied dropdowns (verified live).
 */
var DROPDOWN_LISTS = [
  { key: "presence", label: "Presence", usedIn: "Weekly Presence",
    options: function(ss) {
      var active = ddFn("activeLocationNames", "./Locations");
      return (active ? active(ss) : ["Remote"]).concat(["Off", "Vacation", "Sick Leave", "Holiday"]);
    } }
  // Table columns cannot take a copied dropdown: a Google Sheets table keeps
  // its own column dropdown, so their colours are set in the table itself
  // (column menu → Edit column type) and kept by createGoogleSheetsTable,
  // which leaves unchanged columns alone (ADR 0054).
];

/** Options of a list: its own function, else the schema list of its table column. */
function dropdownOptions(ss, list) {
  if (typeof list.options === "function") return list.options(ss);
  var cfg = SCHEMA_SPEC[list.table];
  return (cfg && cfg.validations && cfg.validations[list.column]) ? cfg.validations[list.column].slice() : [];
}

function dropdownList(key) {
  return DROPDOWN_LISTS.filter(function(l) { return l.key === key; })[0] || null;
}

/**
 * Creates the Dropdowns tab when missing, refreshes every options column, and
 * gives a plain model dropdown to a list that has none yet. Existing model
 * dropdowns (style and colours) are left alone.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet|null}
 */
function ensureDropdownsTab(ss) {
  if (!ss) return null;
  var sheet = getTable(ss, "dropdowns");
  if (!sheet) {
    sheet = ss.insertSheet(tableSheetName("dropdowns"));
    try { sheet.hideSheet(); } catch (eHide) {}
  }
  var width = DROPDOWN_FIRST_OPTIONS_COL - 1 + DROPDOWN_LISTS.length;
  if (sheet.getMaxColumns() < width && typeof sheet.insertColumnsAfter === "function") {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), width - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, 3).setValues([["List", "Dropdown (set its chip style and colours here)", "Used in"]]);
  // Lists removed by a release: their rows and options columns are cleared
  var extraRows = sheet.getMaxRows() - (DROPDOWN_LISTS.length + 1);
  if (extraRows > 0) {
    sheet.getRange(DROPDOWN_LISTS.length + 2, 1, extraRows, 3).clearContent();
    try { sheet.getRange(DROPDOWN_LISTS.length + 2, 2, extraRows, 1).clearDataValidations(); } catch (eDv) {}
  }
  var extraCols = sheet.getMaxColumns() - width;
  if (extraCols > 0) sheet.getRange(1, width + 1, sheet.getMaxRows(), extraCols).clearContent();
  try { sheet.getRange(1, 1, 1, width).setFontWeight("bold"); sheet.setFrozenRows(1); } catch (eFmt) {}

  DROPDOWN_LISTS.forEach(function(list, i) {
    var row = i + 2;
    var optCol = DROPDOWN_FIRST_OPTIONS_COL + i;
    var options = dropdownOptions(ss, list);
    sheet.getRange(row, 1).setValue(list.label);
    sheet.getRange(row, 3).setValue(list.usedIn);
    // Options column: rewritten each time, so new values reach the dropdowns
    var maxRows = Math.max(sheet.getMaxRows(), options.length + 1);
    if (sheet.getMaxRows() < maxRows && typeof sheet.insertRowsAfter === "function") {
      sheet.insertRowsAfter(sheet.getMaxRows(), maxRows - sheet.getMaxRows());
    }
    sheet.getRange(1, optCol).setValue(list.label);
    if (sheet.getMaxRows() > 1) sheet.getRange(2, optCol, sheet.getMaxRows() - 1, 1).clearContent();
    if (options.length) sheet.getRange(2, optCol, options.length, 1).setValues(options.map(function(o) { return [o]; }));
    // Model cell: a dropdown from the options column, created once
    var model = sheet.getRange(row, 2);
    if (!model.getDataValidation()) {
      var colLetter = String.fromCharCode(64 + optCol);
      model.setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInRange(sheet.getRange(colLetter + "2:" + colLetter), true)
        .setAllowInvalid(true).build());
    }
  });
  return sheet;
}

/** The model cell of a list, or null. */
function dropdownModel(ss, key) {
  var idx = DROPDOWN_LISTS.indexOf(dropdownList(key));
  var sheet = idx >= 0 ? getTable(ss, "dropdowns") : null;
  if (!sheet) return null;
  var cell = sheet.getRange(idx + 2, 2);
  return cell.getDataValidation() ? cell : null;
}

/**
 * Copies a list's model dropdown onto a range (keeps chip style and colours).
 * @return {boolean} false when there is no model, so the caller sets its own rule
 */
function applyDropdown(ss, key, range) {
  var model = dropdownModel(ss, key);
  if (!model || !range) return false;
  try {
    model.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    return true;
  } catch (e) {
    Logger.log("Dropdown " + key + " copy note: " + e);
    return false;
  }
}

/**
 * Copies every table list's model onto its table column (rows 2 to the end).
 * @return {string[]} lists applied
 */
function applyTableDropdowns(ss) {
  var done = [];
  DROPDOWN_LISTS.forEach(function(list) {
    if (!list.table) return;
    var sheet = getTable(ss, list.table);
    var col = SCHEMA_SPEC[list.table].columns.indexOf(list.column) + 1;
    if (!sheet || col < 1 || sheet.getMaxRows() < 2) return;
    if (applyDropdown(ss, list.key, sheet.getRange(2, col, sheet.getMaxRows() - 1, 1))) done.push(list.key);
  });
  return done;
}

/** Menu: Edit Dropdowns. */
function showDropdownsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureDropdownsTab(ss);
  if (!sheet) return { success: false, message: "The Dropdowns tab could not be created." };
  if (sheet.isSheetHidden()) sheet.showSheet();
  sheet.activate();
  try {
    ss.toast("Click a cell in column B, then Data → Data validation: pick the Chip style and a colour per option. Run Initialize / Upgrade (or edit Weekly Presence's week) to copy it everywhere, then hide the tab again.", "Dropdowns", 15);
  } catch (eToast) {}
  return { success: true };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    DROPDOWN_LISTS: DROPDOWN_LISTS,
    dropdownOptions: dropdownOptions,
    ensureDropdownsTab: ensureDropdownsTab,
    dropdownModel: dropdownModel,
    applyDropdown: applyDropdown,
    applyTableDropdowns: applyTableDropdowns,
    showDropdownsTab: showDropdownsTab
  };
}
