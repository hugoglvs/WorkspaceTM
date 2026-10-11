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
 * @param {string} tableName - The readable table name, same as its tab (e.g. "Members").
 * @param {string[]} columns - Array of column header names.
 * @param {Object} columnTypes - Map of column name to data type (PEOPLE, DATE, DROPDOWN, NUMERIC, TEXT).
 * @param {Object} [validations] - Map of column name to allowed dropdown options.
 * @param {number} [initialRows=20] - Number of initial rows in the table (including header).
 * @return {boolean}
 */
function createGoogleSheetsTable(spreadsheetId, sheetId, tableName, columns, columnTypes, validations, initialRows, columnLabels) {
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
      columnName: (columnLabels && columnLabels[i]) ? columnLabels[i] : colName
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
    } else if (declaredType === "BOOLEAN" || declaredType === "CHECKBOX") {
      prop.columnType = "BOOLEAN";
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
  // can still hold a name. A table already on this tab (any name) is kept, renamed to
  // tableName and widened if the schema gained columns; another tab's table holding
  // tableName is renamed out of the way.
  var requests = [request];
  var all = listSheetTables(spreadsheetId);
  var own = all.filter(function(t) { return t.sheetId === sheetId; })[0];
  var updates = [];
  all.forEach(function(t) {
    if (t.name === tableName && t.sheetId !== sheetId) {
      updates.push({ updateTable: { table: { tableId: t.tableId, name: "archived_" + tableName.replace(/\s+/g, "_") + "_" + t.sheetId }, fields: "name" } });
    }
  });
  if (own) {
    var fields = [];
    var table = { tableId: own.tableId };
    if (own.name !== tableName) { table.name = tableName; fields.push("name"); }
    var r = own.range;
    if (r && typeof r.endColumnIndex === "number" && r.endColumnIndex < numCols) {
      table.range = {
        sheetId: sheetId,
        startRowIndex: r.startRowIndex || 0,
        endRowIndex: r.endRowIndex || numRows,
        startColumnIndex: r.startColumnIndex || 0,
        endColumnIndex: numCols
      };
      fields.push("range");
    }
    if (fields.length) updates.push({ updateTable: { table: table, fields: fields.join(",") } });
    if (updates.length) {
      try {
        Sheets.Spreadsheets.batchUpdate({ requests: updates }, spreadsheetId);
      } catch (eUpd) {
        Logger.log("Updating table '" + tableName + "' note: " + eUpd);
      }
    }
    // Column names, types and dropdown lists follow the schema (e.g. the Locations list).
    // Sent separately so a refusal here never blocks the rename above.
    try {
      Sheets.Spreadsheets.batchUpdate({ requests: [{
        updateTable: { table: { tableId: own.tableId, columnProperties: columnProperties }, fields: "columnProperties" }
      }] }, spreadsheetId);
    } catch (eCols) {
      Logger.log("Updating columns of table '" + tableName + "' note: " + eCols);
    }
    return true;
  }
  requests = updates.concat(requests);

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
 * Lists every native table in the spreadsheet. Returns [] if the lookup fails,
 * so table creation falls back to adding the table.
 *
 * @param {string} spreadsheetId
 * @return {Array<{tableId: string, name: string, sheetId: number, range: Object}>}
 */
function listSheetTables(spreadsheetId) {
  var found = [];
  try {
    var meta = Sheets.Spreadsheets.get(spreadsheetId, { fields: "sheets(properties(sheetId,title),tables(tableId,name,range))" });
    var sheets = (meta && meta.sheets) || [];
    for (var i = 0; i < sheets.length; i++) {
      var tables = sheets[i].tables || [];
      for (var t = 0; t < tables.length; t++) {
        found.push({ tableId: tables[t].tableId, name: tables[t].name, sheetId: sheets[i].properties.sheetId, range: tables[t].range });
      }
    }
  } catch (eGet) {
    Logger.log("listSheetTables note: " + eGet);
  }
  return found;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    applyTypography: applyTypography,
    createGoogleSheetsTable: createGoogleSheetsTable
  };
}
