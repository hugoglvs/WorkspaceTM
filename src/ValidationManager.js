/**
 * Project Management Workspace - Validation Manager
 * Enforces column nullability, required field validation, and integrity auditing.
 */

var PROMINENT_WARNING_BG = "#F8D7DA"; // Prominent Alert Coral/Red
var PROMINENT_WARNING_FONT = "#721C24"; // Bold Danger Maroon

/**
 * Checks whether an array of cell values contains at least one non-empty value.
 *
 * @param {Array} rowValues
 * @return {boolean}
 */
function isRowPopulated(rowValues) {
  if (!rowValues || !rowValues.length) return false;
  for (var i = 0; i < rowValues.length; i++) {
    var v = rowValues[i];
    if (v !== "" && v !== null && v !== undefined) {
      if (typeof v === "string" && v.trim() !== "") return true;
      if (typeof v !== "string") return true;
    }
  }
  return false;
}

/**
 * Validates a single row against its schema's required and nullable columns.
 *
 * @param {string} sheetName
 * @param {Array} rowValues - Array of cell values for the row.
 * @param {Array} [headers] - Array of column header names (defaults to schema).
 * @return {Object} { valid: boolean, populated: boolean, missingColumns: Array<string> }
 */
function validateRow(sheetName, rowValues, headers) {
  var config = (typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC[sheetName]) ? SCHEMA_SPEC[sheetName] : null;
  if (!config) {
    return { valid: true, populated: isRowPopulated(rowValues), missingColumns: [] };
  }

  var cols = headers || config.columns;
  var required = config.requiredColumns || [];

  if (!isRowPopulated(rowValues)) {
    return { valid: true, populated: false, missingColumns: [] };
  }

  // Build map of colName -> value
  var valueMap = {};
  for (var c = 0; c < cols.length; c++) {
    var colName = String(cols[c] || "").trim().toLowerCase();
    valueMap[colName] = c < rowValues.length ? rowValues[c] : "";
  }

  var missingColumns = [];
  for (var r = 0; r < required.length; r++) {
    var reqCol = required[r].toLowerCase();
    var val = valueMap[reqCol];

    var isEmpty = false;
    if (val === null || val === undefined) {
      isEmpty = true;
    } else if (typeof val === "string") {
      isEmpty = val.trim() === "";
    } else if (val instanceof Date) {
      isEmpty = isNaN(val.getTime());
    } else if (typeof val === "object" && typeof val.getEmail === "function") {
      isEmpty = !val.getEmail() && !val.getName();
    }

    if (isEmpty) {
      missingColumns.push(required[r]);
    }
  }

  return {
    valid: missingColumns.length === 0,
    populated: true,
    missingColumns: missingColumns
  };
}

/**
 * Validates all populated rows in a given table sheet.
 *
 * @param {string} sheetName
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Report of populated, valid, and invalid rows.
 */
function validateTable(sheetName, targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { sheetName: sheetName, error: "No active spreadsheet" };

  var sheet = ss.getSheetByName(sheetName);
  var config = (typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC[sheetName]) ? SCHEMA_SPEC[sheetName] : null;

  if (!sheet || !config) {
    return { sheetName: sheetName, error: "Sheet or schema configuration not found" };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(typeof sheet.getLastColumn === "function" ? (sheet.getLastColumn() || 0) : 0, config.columns.length);

  if (lastRow <= 1) {
    return {
      sheetName: sheetName,
      totalRows: 0,
      populatedRows: 0,
      validRows: 0,
      invalidRows: []
    };
  }

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var dataValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  var populatedCount = 0;
  var validCount = 0;
  var invalidRows = [];

  for (var i = 0; i < dataValues.length; i++) {
    var rowNum = i + 2;
    var result = validateRow(sheetName, dataValues[i], headers);

    if (result.populated) {
      populatedCount++;
      if (result.valid) {
        validCount++;
      } else {
        invalidRows.push({
          row: rowNum,
          missingColumns: result.missingColumns
        });
      }
    }
  }

  return {
    sheetName: sheetName,
    totalRows: dataValues.length,
    populatedRows: populatedCount,
    validRows: validCount,
    invalidRows: invalidRows
  };
}

/**
 * Validates all operational tables across the spreadsheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Global validation summary.
 */
function validateAllTables(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet" };

  var results = {};
  var totalIssues = 0;

  for (var tableKey in SCHEMA_SPEC) {
    if (!SCHEMA_SPEC.hasOwnProperty(tableKey)) continue;
    var report = validateTable(tableKey, ss);
    results[tableKey] = report;
    if (report.invalidRows && report.invalidRows.length > 0) {
      totalIssues += report.invalidRows.length;
    }
  }

  return {
    success: totalIssues === 0,
    totalIssues: totalIssues,
    tables: results
  };
}

/**
 * Clears hardcoded static backgrounds from data cells so dynamic conditional formatting
 * operates without interference.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} [config]
 */
function clearStaticCellBackgrounds(sheet, config) {
  if (!sheet) return;
  var maxRows = Math.max(sheet.getMaxRows(), 50);
  var colCount = config ? config.columns.length : Math.max(typeof sheet.getLastColumn === "function" ? (sheet.getLastColumn() || 0) : 0, 7);
  if (maxRows > 1 && colCount > 0) {
    sheet.getRange(2, 1, maxRows - 1, colCount).setBackground(null).setFontColor(null).setFontWeight("normal");
  }
}

/**
 * Directly sets the background color of empty required cells in populated rows
 * to prominent alert red (#F8D7DA) during explicit health audits.
 *
 * @param {string} sheetName
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {number} Count of highlighted cells
 */
function highlightInvalidCells(sheetName, targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return 0;

  var sheet = ss.getSheetByName(sheetName);
  var config = (typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC[sheetName]) ? SCHEMA_SPEC[sheetName] : null;
  if (!sheet || !config) return 0;

  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(typeof sheet.getLastColumn === "function" ? (sheet.getLastColumn() || 0) : 0, config.columns.length);
  if (lastRow <= 1) return 0;

  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
  var dataValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var highlightedCount = 0;

  for (var i = 0; i < dataValues.length; i++) {
    var rowNum = i + 2;
    var rowResult = validateRow(sheetName, dataValues[i], headers);

    if (rowResult.populated) {
      for (var c = 0; c < headers.length; c++) {
        var colName = headers[c];
        var isMissing = rowResult.missingColumns.indexOf(colName) !== -1;
        var cellRange = sheet.getRange(rowNum, c + 1);

        if (isMissing) {
          cellRange.setBackground(PROMINENT_WARNING_BG);
          cellRange.setFontColor(PROMINENT_WARNING_FONT);
          cellRange.setFontWeight("bold");
          highlightedCount++;
        } else if (config.requiredColumns.indexOf(colName) !== -1) {
          cellRange.setBackground(null);
          cellRange.setFontColor(null);
          cellRange.setFontWeight("normal");
        }
      }
    }
  }

  return highlightedCount;
}

/**
 * Applies prominent warning conditional formatting to highlight empty required cells
 * in real-time whenever a row is populated.
 *
 * Formula: =AND(COUNTA($A2:$G2)>0, LEN(TRIM(A2))=0)
 * Evaluates dynamically and universally in the browser:
 * - Empty rows: zero highlights.
 * - Active rows: any empty required cell turns prominent alert red (#F8D7DA).
 * - Once set: instantly reverts back to clean table banding without delay.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} config
 */
function applyRequiredValidationFormatting(sheet, config) {
  if (!sheet || !config || !config.requiredColumns) return;
  if (typeof SpreadsheetApp === "undefined" || !SpreadsheetApp.newConditionalFormatRule) return;

  try {
    var rules = sheet.getConditionalFormatRules ? sheet.getConditionalFormatRules() : [];
    var maxRows = Math.max(sheet.getMaxRows(), 50);
    var cols = config.columns;
    var required = config.requiredColumns;
    var endColLetter = String.fromCharCode(64 + cols.length);

    for (var i = 0; i < required.length; i++) {
      var colName = required[i];
      var colIdx = cols.indexOf(colName) + 1;
      if (colIdx < 1) continue;

      var colLetter = String.fromCharCode(64 + colIdx);
      var range = sheet.getRange(2, colIdx, maxRows - 1, 1);

      // Universal real-time formula: row has content anywhere, but this cell is blank
      var formula = '=AND(COUNTA($A2:$' + endColLetter + '2)>0, LEN(TRIM(' + colLetter + '2))=0)';

      var rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground(PROMINENT_WARNING_BG)
        .setFontColor(PROMINENT_WARNING_FONT)
        .setBold(true)
        .setRanges([range])
        .build();

      rules.push(rule);
    }

    if (sheet.setConditionalFormatRules) {
      sheet.setConditionalFormatRules(rules);
    }
  } catch (err) {
    Logger.log("applyRequiredValidationFormatting note: " + err);
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    PROMINENT_WARNING_BG: PROMINENT_WARNING_BG,
    PROMINENT_WARNING_FONT: PROMINENT_WARNING_FONT,
    isRowPopulated: isRowPopulated,
    validateRow: validateRow,
    validateTable: validateTable,
    validateAllTables: validateAllTables,
    clearStaticCellBackgrounds: clearStaticCellBackgrounds,
    highlightInvalidCells: highlightInvalidCells,
    applyRequiredValidationFormatting: applyRequiredValidationFormatting
  };
}
