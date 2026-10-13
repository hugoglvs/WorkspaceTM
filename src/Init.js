/**
 * Project Management Workspace - Initialization Engine
 * Ticket 01: Core Schema & Typography Initializer
 */

function initNorm(s) {
  return String(s === null || s === undefined ? "" : s).trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

/**
 * True when a table holding data has a known column (by key, readable label or
 * former label) at another position than the schema says, i.e. a release
 * reordered or removed columns. A column added at the end, or a renamed label
 * in place, is not a move.
 */
function columnsMoved(sheet, config) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return false;
  var data = sheet.getRange(2, 1, Math.min(sheet.getLastRow(), 500) - 1, sheet.getLastColumn()).getValues();
  var hasData = data.some(function(row) { return row.some(function(v) { return v !== "" && v !== null && v !== false; }); });
  if (!hasData) return false;
  var header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var known = {};
  config.columns.forEach(function(key, i) {
    known[initNorm(key)] = i;
    if (config.headers && config.headers[key]) known[initNorm(config.headers[key])] = i;
    ((config.formerHeaders || {})[key] || []).forEach(function(h) { known[initNorm(h)] = i; });
  });
  ((config.removedHeaders) || []).forEach(function(h) { known[initNorm(h)] = -1; });
  for (var c = 0; c < header.length; c++) {
    var n = initNorm(header[c]);
    if (n && known.hasOwnProperty(n) && known[n] !== c) return true;
  }
  return false;
}

/** Renames a table whose columns moved to "Archive – <tab> <date>", hidden, and forgets its sheet ID. */
function archiveMovedTable(ss, tableKey, config, sheet) {
  var d = new Date();
  var stamp = d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2);
  var base = (typeof ARCHIVE_PREFIX !== "undefined" ? ARCHIVE_PREFIX : "Archive – ") + config.sheetName + " " + stamp;
  var name = base, n = 2;
  while (ss.getSheetByName(name)) name = base + " (" + (n++) + ")";
  sheet.setName(name);
  try { sheet.hideSheet(); } catch (eHide) {}
  var props = (typeof tableDocProps === "function") ? tableDocProps() : null;
  if (props && typeof tableIdProperty === "function") props.deleteProperty(tableIdProperty(tableKey));
  Logger.log("Columns of '" + config.sheetName + "' changed: old tab kept as '" + name + "', table rebuilt.");
  return name;
}

/**
 * Initializes a Google Spreadsheet from scratch with the 8 core operational tables.
 * Applies 12px "Google Sans", frozen header rows, table row banding,
 * dropdown validations, and sets SCHEMA_VERSION = "1".
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet] - Optional target spreadsheet (defaults to active).
 * @return {Object} Status report of created sheets.
 */
function freshInit(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("No active spreadsheet found. Please run within a Google Sheet.");
  }

  var createdSheets = [];
  var initialSheets = ss.getSheets();

  for (var tableKey in SCHEMA_SPEC) {
    if (!SCHEMA_SPEC.hasOwnProperty(tableKey)) continue;

    var config = SCHEMA_SPEC[tableKey];
    var sheet = getTable(ss, tableKey);
    var isNew = false;

    // Columns reordered or removed by a release: keep the old tab as an archive
    // and build the table again (ADR 0051), rather than shifting data under new headers
    if (sheet && columnsMoved(sheet, config)) {
      archiveMovedTable(ss, tableKey, config, sheet);
      sheet = null;
    }

    if (!sheet) {
      sheet = ss.insertSheet(config.sheetName);
      isNew = true;
    }

    // Ensure at least 50 rows and exact number of columns
    var colCount = config.columns.length;
    var currentCols = sheet.getMaxColumns();
    if (currentCols < colCount) {
      sheet.insertColumnsAfter(currentCols, colCount - currentCols);
    } else if (currentCols > colCount && isNew) {
      sheet.deleteColumns(colCount + 1, currentCols - colCount);
    }

    var rowCount = Math.max(sheet.getMaxRows(), 50);


    // 1. Write Header Row
    var headerRange = sheet.getRange(1, 1, 1, colCount);
    headerRange.setValues([schemaHeaderRow(config)]);
    sheet.setFrozenRows(1);

    // 2. Set Typography across data area (12px Google Sans)
    var dataArea = sheet.getRange(1, 1, rowCount, colCount);
    applyTypography(dataArea);

    // 3. Remove any existing alternating row bandings to prevent Sheets API conflict
    var existingBandings = sheet.getBandings();
    for (var b = 0; b < existingBandings.length; b++) {
      existingBandings[b].remove();
    }

    // 4. Create Native Google Sheets Table (with typed columns: PEOPLE, DATE, DROPDOWN, NUMERIC, TEXT)
    createGoogleSheetsTable(
      ss.getId(),
      sheet.getSheetId(),
      config.sheetName,
      config.columns,
      config.columnTypes,
      (typeof schemaValidationsFor === "function") ? schemaValidationsFor(ss, config) : config.validations,
      Math.min(rowCount, 25),
      schemaHeaderRow(config)
    );


    // Columns beyond the schema that belonged to it (e.g. a column moved or removed by a
    // release, on a table without data): cleared once the table was narrowed above, so
    // no stray header remains (ADR 0051)
    if (!isNew && sheet.getLastColumn() > colCount) {
      var extra = sheet.getRange(1, colCount + 1, 1, sheet.getLastColumn() - colCount).getValues()[0];
      var knownHeaders = {};
      config.columns.forEach(function(k) {
        knownHeaders[initNorm(k)] = true;
        if (config.headers && config.headers[k]) knownHeaders[initNorm(config.headers[k])] = true;
        ((config.formerHeaders || {})[k] || []).forEach(function(h) { knownHeaders[initNorm(h)] = true; });
      });
      (config.removedHeaders || []).forEach(function(h) { knownHeaders[initNorm(h)] = true; });
      extra.forEach(function(h, i) {
        if (!knownHeaders[initNorm(h)]) return;
        try { sheet.getRange(1, colCount + 1 + i, Math.max(sheet.getLastRow(), 1), 1).clearContent(); }
        catch (eClr) { Logger.log("Former column " + h + " of " + config.sheetName + " not cleared: " + eClr); }
      });
    }

    // 5. Technical columns stay in the table, hidden (ADR 0051)
    (config.hiddenColumns || []).forEach(function(key) {
      var idx = config.columns.indexOf(key) + 1;
      if (idx > 0 && typeof sheet.hideColumns === "function") {
        try { sheet.hideColumns(idx); } catch (eHide) { Logger.log("Hide column note: " + eHide); }
      }
    });

    // 6. Register Named Range for table
    // Named ranges cannot contain spaces: keep the stable technical name
    var namedRangeName = (config.legacyName || tableKey) + "_table";
    var existingNamedRange = ss.getRangeByName(namedRangeName);
    if (existingNamedRange) {
      ss.removeNamedRange(namedRangeName);
    }
    ss.setNamedRange(namedRangeName, sheet.getRange(1, 1, rowCount, colCount));

    // 7. Adjust column widths
    for (var c = 1; c <= colCount; c++) {
      sheet.setColumnWidth(c, 160);
    }

    // 8. Clear any static cell backgrounds and apply dynamic conditional formatting
    if (typeof clearStaticCellBackgrounds === "function") {
      try {
        clearStaticCellBackgrounds(sheet, config);
      } catch (errClr) {
        Logger.log("Clear static backgrounds note: " + errClr);
      }
    }
    if (typeof applyRequiredValidationFormatting === "function") {
      try {
        applyRequiredValidationFormatting(sheet, config);
      } catch (errVal) {
        Logger.log("Validation formatting note: " + errVal);
      }
    }

    createdSheets.push(config.sheetName);
  }

  // Organisation settings, before anything that reads them (ADR 0049)
  if (typeof seedDefaultSettings === "function") {
    try { seedDefaultSettings(ss); } catch (eSet) { Logger.log("Settings seed note: " + eSet); }
  }

  // Default Work Locations in a new, empty Locations tab (ADR 0040)
  if (typeof seedDefaultLocations === "function") {
    try { seedDefaultLocations(ss); } catch (eSeed) { Logger.log("Locations seed note: " + eSeed); }
  }

  // Default task types in a new Task Types tab (ADR 0045)
  if (typeof seedDefaultTaskTypes === "function") {
    try { seedDefaultTaskTypes(ss); } catch (eTT) { Logger.log("Task types seed note: " + eTT); }
  }

  // French public holidays of this year and next in a new Holidays tab (ADR 0043)
  if (typeof seedPublicHolidays === "function") {
    try { seedPublicHolidays(ss); } catch (eHol) { Logger.log("Holidays seed note: " + eHol); }
  }

  // Safely remove the initial blank sheet (e.g. "Sheet1" or "Feuille 1") if created from scratch and still empty
  if (initialSheets.length === 1) {
    var firstSheet = initialSheets[0];
    var firstName = firstSheet.getName();
    var firstIsEmpty = !firstSheet.getLastRow() || !firstSheet.getLastColumn();
    if (!tableKeyOfName(firstName) && createdSheets.length > 0 && firstIsEmpty) {
      try {
        ss.deleteSheet(firstSheet);
      } catch (e) {
        // Can't delete if it's the only remaining sheet, which won't happen here
      }
    }
  }

  // Apply column protections across all tables
  if (typeof applyProtections === "function") {
    try {
      applyProtections(ss);
    } catch (e) {
      Logger.log("Protections init note: " + e);
    }
  }

  // Record baseline schema version in Document Properties
  var props = PropertiesService.getDocumentProperties();
  props.setProperty("SCHEMA_VERSION", String(CURRENT_SCHEMA_VERSION));

  return {
    success: true,
    schemaVersion: CURRENT_SCHEMA_VERSION,
    sheetsCreated: createdSheets
  };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    freshInit: freshInit
  };
}
