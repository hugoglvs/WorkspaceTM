/**
 * Project Management Workspace - Typography & Table Styling
 */

/**
 * Applies 12px "Google Sans" font to a given range.
 * @param {GoogleAppsScript.Spreadsheet.Range} range
 */
function applyTypography(range) {
  range.setFontFamily(FONT_FAMILY);
  range.setFontSize(FONT_SIZE);
}

/**
 * Creates a native Google Sheets Table using the Sheets API v4 (addTable).
 * Casts columns to native types: PEOPLE (smart chips), DATE, DROPDOWN, NUMERIC, TEXT.
 *
 * @param {string} spreadsheetId - The active spreadsheet ID.
 * @param {number} sheetId - The specific sheet/tab ID.
 * @param {string} tableName - The canonical table name (e.g., "tbl_members").
 * @param {string[]} columns - Array of column header names.
 * @param {Object} columnTypes - Map of column name to data type (PEOPLE, DATE, DROPDOWN, NUMERIC, TEXT).
 * @param {Object} [validations] - Map of column name to allowed dropdown options.
 * @param {number} [initialRows=20] - Number of initial rows in the table (including header).
 * @return {boolean}
 */
function createGoogleSheetsTable(spreadsheetId, sheetId, tableName, columns, columnTypes, validations, initialRows) {
  if (typeof Sheets === "undefined" || !Sheets.Spreadsheets) {
    throw new Error(
      "The 'Google Sheets API' advanced service is not enabled in your Apps Script project.\n\n" +
      "To create native Google Sheets Tables, please enable it:\n" +
      "1. In the Apps Script left sidebar, click '+' next to Services.\n" +
      "2. Select 'Google Sheets API' (v4, identifier 'Sheets').\n" +
      "3. Click 'Add'.\n" +
      "4. Run init() again."
    );
  }

  var numRows = initialRows || 20;
  var numCols = columns.length;

  var columnProperties = [];
  for (var i = 0; i < columns.length; i++) {
    var colName = columns[i];
    var declaredType = (columnTypes && columnTypes[colName]) ? columnTypes[colName] : "TEXT";

    var prop = {
      columnIndex: i,
      columnName: colName
    };

    if (declaredType === "DROPDOWN" && validations && validations[colName]) {
      prop.columnType = "DROPDOWN";
      prop.dataValidationRule = {
        condition: {
          type: "ONE_OF_LIST",
          values: validations[colName].map(function(val) {
            return { userEnteredValue: String(val) };
          })
        }
      };
    } else if (declaredType === "PEOPLE" || declaredType === "PEOPLE_CHIP") {
      prop.columnType = "PEOPLE_CHIP";
    } else if (declaredType === "DATE" || declaredType === "DATETIME") {
      prop.columnType = "DATE";
    } else if (declaredType === "DOUBLE" || declaredType === "NUMBER" || declaredType === "NUMERIC" || declaredType === "INTEGER") {
      prop.columnType = "DOUBLE";
    } else if (declaredType === "PERCENT") {
      prop.columnType = "PERCENT";
    } else if (declaredType === "CURRENCY") {
      prop.columnType = "CURRENCY";
    } else {
      prop.columnType = "TEXT";
    }

    columnProperties.push(prop);
  }

  var request = {
    addTable: {
      table: {
        name: tableName,
        range: {
          sheetId: sheetId,
          startRowIndex: 0,
          endRowIndex: numRows,
          startColumnIndex: 0,
          endColumnIndex: numCols
        },
        columnProperties: columnProperties
      }
    }
  };

  // Table names are unique per spreadsheet and survive tab renames, so an archived tab
  // (e.g. archive_tbl_leaves) can still hold "tbl_leaves". Check before adding.
  var requests = [request];
  var existing = findTablesByName(spreadsheetId, tableName);
  for (var e = 0; e < existing.length; e++) {
    if (existing[e].sheetId === sheetId) {
      Logger.log("Table '" + tableName + "' already exists on its sheet, preserving it.");
      return true;
    }
  }
  for (var x = 0; x < existing.length; x++) {
    requests.unshift({
      updateTable: {
        table: { tableId: existing[x].tableId, name: "archived_" + tableName + "_" + existing[x].sheetId },
        fields: "name"
      }
    });
  }

  try {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
    return true;
  } catch (err) {
    var msg = (err && err.message) ? err.message : String(err);
    if (msg.indexOf("already exists") !== -1 || msg.indexOf("already in use") !== -1) {
      Logger.log("Table '" + tableName + "' already exists, preserving existing table.");
      return true;
    }
    Logger.log("Sheets API addTable error on '" + tableName + "': " + msg);
    throw new Error("Failed to create Google Sheets Table '" + tableName + "': " + msg);
  }
}

/**
 * Lists the native tables named tableName across the whole spreadsheet.
 * Returns [] if the lookup fails, so table creation falls back to its old behaviour.
 *
 * @param {string} spreadsheetId
 * @param {string} tableName
 * @return {Array<{tableId: string, sheetId: number}>}
 */
function findTablesByName(spreadsheetId, tableName) {
  var found = [];
  try {
    var meta = Sheets.Spreadsheets.get(spreadsheetId, { fields: "sheets(properties(sheetId,title),tables(tableId,name))" });
    var sheets = (meta && meta.sheets) || [];
    for (var i = 0; i < sheets.length; i++) {
      var tables = sheets[i].tables || [];
      for (var t = 0; t < tables.length; t++) {
        if (tables[t].name === tableName) found.push({ tableId: tables[t].tableId, sheetId: sheets[i].properties.sheetId });
      }
    }
  } catch (eGet) {
    Logger.log("findTablesByName note: " + eGet);
  }
  return found;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    applyTypography: applyTypography,
    createGoogleSheetsTable: createGoogleSheetsTable
  };
}
