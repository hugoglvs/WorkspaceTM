/**
 * Project Management Workspace - Initialization Engine
 * Ticket 01: Core Schema & Typography Initializer
 */

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

  // Default Work Locations in a new, empty Locations tab (ADR 0040)
  if (typeof seedDefaultLocations === "function") {
    try { seedDefaultLocations(ss); } catch (eSeed) { Logger.log("Locations seed note: " + eSeed); }
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
