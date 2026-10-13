/**
 * Mock Google Apps Script environment for local unit testing
 */

/**
 * The mock does not run Sheets formulas, except the Members formulas of
 * ADR 0047: those are evaluated with the reference rule contractProfile.
 */
function mockCellText(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "object" && !(v instanceof Date)) {
    if (typeof v.getName === "function") return String(v.getName()).trim();
    return String(v.name || v.email || "").trim();
  }
  return String(v).trim();
}

function mockDate(v) {
  if (v instanceof Date) return v;
  var str = String(v === undefined || v === null ? "" : v).trim();
  if (!str) return null;
  var iso = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3]);
  var d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function evaluateMockFormula(sheet, row, col, formula) {
  var m = /^=IF\(\$A(\d+)="",""/.exec(formula);
  var sn = /'((?:[^']|'')+)'!/.exec(formula);
  if (!m || !sn) return "";
  var ss = sheet.getParent ? sheet.getParent() : null;
  var contracts = ss && ss.getSheetByName ? ss.getSheetByName(sn[1].replace(/''/g, "'")) : null;
  var name = mockCellText(sheet.grid[+m[1]] && sheet.grid[+m[1]][1]).toLowerCase();
  if (!name || !contracts) return "";
  var records = [];
  Object.keys(contracts.grid).forEach(function(k) {
    var r = +k, g = contracts.grid[k];
    if (r < 2 || !g || mockCellText(g[1]).toLowerCase() !== name) return;
    records.push({ role: mockCellText(g[2]), department: mockCellText(g[3]), start: mockDate(g[4]), end: mockDate(g[5]) });
  });
  var prof = require("../src/Contracts").contractProfile(records, new Date());
  if (formula.indexOf('"No Contract"') !== -1) return prof.status;
  if (formula.indexOf('"Team Member"') !== -1) return prof.role;
  if (formula.indexOf('"Delivery"') !== -1) return prof.department;
  return "";
}

function mockClearFormulas(sheet, startRow, startCol, numRows, numCols) {
  if (!sheet.formulas) return;
  for (var r = 0; r < numRows; r++) {
    var fr = sheet.formulas[startRow + r];
    if (fr) for (var c = 0; c < numCols; c++) delete fr[startCol + c];
  }
}

function createMockRange(sheet, startRow, startCol, numRows, numCols) {
  return {
    sheet: sheet,
    startRow: startRow,
    startCol: startCol,
    numRows: numRows,
    numCols: numCols,
    values: [],
    fontFamily: null,
    fontSize: null,
    fontWeight: null,
    background: null,
    fontColor: null,
    horizontalAlignment: null,
    dataValidation: null,
    getSheet: function() { return this.sheet; },
    getRow: function() { return this.startRow; },
    getColumn: function() { return this.startCol; },
    getNumRows: function() { return this.numRows; },
    getNumColumns: function() { return this.numCols; },
    getLastRow: function() { return this.startRow + this.numRows - 1; },
    getLastColumn: function() { return this.startCol + this.numCols - 1; },

    setFormula: function(f) { return this.setFormulas([[f]]); },
    setFormulas: function(arr) {
      if (!this.sheet.formulas) this.sheet.formulas = {};
      for (var r = 0; r < arr.length; r++) {
        for (var c = 0; c < arr[r].length; c++) {
          var rowIdx = startRow + r, colIdx = startCol + c;
          if (!this.sheet.formulas[rowIdx]) this.sheet.formulas[rowIdx] = {};
          this.sheet.formulas[rowIdx][colIdx] = arr[r][c];
          if (this.sheet.grid[rowIdx]) delete this.sheet.grid[rowIdx][colIdx];
        }
      }
      return this;
    },
    getFormula: function() { return this.getFormulas()[0][0]; },
    getFormulas: function() {
      var out = [];
      for (var r = 0; r < numRows; r++) {
        var row = [];
        for (var c = 0; c < numCols; c++) {
          var fr = this.sheet.formulas && this.sheet.formulas[startRow + r];
          row.push((fr && fr[startCol + c]) || "");
        }
        out.push(row);
      }
      return out;
    },
    setValue: function(val) {
      mockClearFormulas(this.sheet, startRow, startCol, 1, 1);
      if (!this.sheet.grid[startRow]) this.sheet.grid[startRow] = {};
      this.sheet.grid[startRow][startCol] = val;
      this.values = [[val]];
      return this;
    },
    clearContent: function() {
      mockClearFormulas(this.sheet, startRow, startCol, numRows, numCols);
      for (var r = 0; r < numRows; r++) {
        var rowIdx = startRow + r;
        if (this.sheet.grid[rowIdx]) {
          for (var c = 0; c < numCols; c++) {
            delete this.sheet.grid[rowIdx][startCol + c];
          }
        }
      }
      this.values = [];
      return this;
    },
    setValues: function(valArray) {
      mockClearFormulas(this.sheet, startRow, startCol, valArray.length, valArray.length ? valArray[0].length : 0);
      this.values = valArray;
      for (var r = 0; r < valArray.length; r++) {
        var rowIdx = startRow + r;
        if (!this.sheet.grid[rowIdx]) this.sheet.grid[rowIdx] = {};
        for (var c = 0; c < valArray[r].length; c++) {
          this.sheet.grid[rowIdx][startCol + c] = valArray[r][c];
        }
      }
      return this;
    },
    getValue: function() {
      var vals = this.getValues();
      return (vals.length > 0 && vals[0].length > 0) ? vals[0][0] : "";
    },
    copyTo: function(destinationRange) {
      var vals = this.getValues();
      destinationRange.setValues(vals);
      return destinationRange;
    },
    getValues: function() {
      var result = [];
      for (var r = 0; r < numRows; r++) {
        var row = [];
        var rowIdx = startRow + r;
        for (var c = 0; c < numCols; c++) {
          var colIdx = startCol + c;
          var formula = this.sheet.formulas && this.sheet.formulas[rowIdx] && this.sheet.formulas[rowIdx][colIdx];
          var val = formula ? evaluateMockFormula(this.sheet, rowIdx, colIdx, formula)
            : (this.sheet.grid[rowIdx] && this.sheet.grid[rowIdx][colIdx] !== undefined)
            ? this.sheet.grid[rowIdx][colIdx]
            : "";
          row.push(val);
        }
        result.push(row);
      }
      return result;
    },
    getRichTextValue: function() {
      var rtv = this.getRichTextValues();
      return (rtv.length > 0 && rtv[0].length > 0) ? rtv[0][0] : null;
    },
    getRichTextValues: function() {
      function makeCell(rawVal) {
        var link = "";
        if (rawVal && typeof rawVal === "object") {
          if (rawVal.linkUrl) link = rawVal.linkUrl;
          else if (typeof rawVal.getEmail === "function") link = "mailto:" + rawVal.getEmail();
          else if (rawVal.email) link = "mailto:" + rawVal.email;
        } else if (typeof rawVal === "string" && rawVal.indexOf("@") !== -1 && rawVal.indexOf(" ") === -1) {
          link = "mailto:" + rawVal;
        }
        return {
          getText: function() {
            return (typeof rawVal === "object" && typeof rawVal.getName === "function") ? rawVal.getName() : String(rawVal || "");
          },
          getLinkUrl: function() { return link; },
          getRuns: function() {
            return link ? [{ getLinkUrl: function() { return link; } }] : [];
          }
        };
      }

      var result = [];
      for (var r = 0; r < numRows; r++) {
        var row = [];
        var rowIdx = startRow + r;
        for (var c = 0; c < numCols; c++) {
          var colIdx = startCol + c;
          var raw = (this.sheet.grid[rowIdx] && this.sheet.grid[rowIdx][colIdx] !== undefined)
            ? this.sheet.grid[rowIdx][colIdx]
            : "";
          row.push(makeCell(raw));
        }
        result.push(row);
      }
      return result;
    },
    setNumberFormat: function(fmt) {
      this.numberFormat = fmt;
      return this;
    },
    setNumberFormats: function(formats) {
      this.numberFormats = formats;
      return this;
    },
    setFontFamily: function(font) {
      this.fontFamily = font;
      return this;
    },
    setFontSize: function(size) {
      this.fontSize = size;
      return this;
    },
    setFontWeight: function(weight) {
      this.fontWeight = weight;
      return this;
    },
    setBackground: function(bg) {
      this.background = bg;
      return this;
    },
    getBackground: function() {
      return this.background || "";
    },
    setFontColor: function(color) {
      this.fontColor = color;
      return this;
    },
    getFontColor: function() {
      return this.fontColor || "";
    },
    getFontWeight: function() {
      return this.fontWeight || "normal";
    },
    setHorizontalAlignment: function(align) {
      this.horizontalAlignment = align;
      return this;
    },
    setVerticalAlignment: function(align) {
      this.verticalAlignment = align;
      return this;
    },
    merge: function() {
      this.merged = true;
      return this;
    },
    setBackgrounds: function(m) { this.backgrounds = m; return this; },
    setNotes: function(m) { this.notes = m; return this; },
    setNote: function(n) { this.note = n; return this; },
    setDataValidation: function(rule) {
      this.dataValidation = rule;
      return this;
    },
    clearDataValidations: function() {
      this.dataValidation = null;
      this.validationCleared = true;
      return this;
    },
    protect: function() {
      var protection = {
        description: "",
        warningOnly: false,
        domainEdit: false,
        editors: [],
        range: this,
        setDescription: function(d) { this.description = d; return this; },
        getDescription: function() { return this.description; },
        getRange: function() { return this.range; },
        setWarningOnly: function(w) { this.warningOnly = w; return this; },
        isWarningOnly: function() { return this.warningOnly; },
        getEditors: function() { return this.editors.slice(); },
        addEditor: function(emailOrUser) {
          var email = (typeof emailOrUser === 'object' && emailOrUser.getEmail) ? emailOrUser.getEmail() : String(emailOrUser);
          this.editors.push({ getEmail: function() { return email; } });
          return this;
        },
        addEditors: function(emails) {
          for (var i = 0; i < emails.length; i++) this.addEditor(emails[i]);
          return this;
        },
        removeEditors: function(usersToRemove) {
          var emailsToRemove = {};
          for (var i = 0; i < usersToRemove.length; i++) {
            var em = (typeof usersToRemove[i] === 'object' && usersToRemove[i].getEmail) ? usersToRemove[i].getEmail() : String(usersToRemove[i]);
            emailsToRemove[em.toLowerCase()] = true;
          }
          this.editors = this.editors.filter(function(ed) {
            return !emailsToRemove[ed.getEmail().toLowerCase()];
          });
          return this;
        },
        removeEditor: function(userToRemove) {
          return this.removeEditors([userToRemove]);
        },
        canDomainEdit: function() { return this.domainEdit; },
        setDomainEdit: function(allow) { this.domainEdit = allow; return this; },
        remove: function() {
          var idx = this.range.sheet.protections.indexOf(this);
          if (idx !== -1) this.range.sheet.protections.splice(idx, 1);
        }
      };
      this.sheet.protections.push(protection);
      return protection;
    },
    getBandings: function() {
      return this.bandings || [];
    },
    applyRowBanding: function(theme, showHeader, showFooter) {
      var banding = {
        theme: theme,
        showHeader: showHeader,
        showFooter: showFooter,
        removed: false,
        remove: function() { this.removed = true; }
      };
      if (!this.bandings) this.bandings = [];
      this.bandings.push(banding);
      return banding;
    }
  };
}

function createMockSheet(name, parent) {
  return {
    name: name,
    parent: parent || null,
    frozenRows: 0,
    maxRows: 100,
    maxCols: 26,
    ranges: {},
    grid: {},
    columnWidths: {},
    protections: [],

    sheetId: Math.floor(Math.random() * 100000),
    getName: function() { return this.name; },
    setName: function(n) { this.name = n; return this; },
    getSheetId: function() { return this.sheetId; },
    getParent: function() {
      return this.parent || (typeof SpreadsheetApp !== "undefined" && SpreadsheetApp.getActiveSpreadsheet ? SpreadsheetApp.getActiveSpreadsheet() : null);
    },
    getFrozenRows: function() { return this.frozenRows; },
    setFrozenRows: function(rows) { this.frozenRows = rows; },
    setFrozenColumns: function(cols) { this.frozenColumns = cols; },
    clear: function() { this.grid = {}; return this; },
    getMaxRows: function() { return this.maxRows; },
    getMaxColumns: function() { return this.maxCols; },
    insertRowsAfter: function(after, count) { this.maxRows += count; },
    hidden: false,
    hideSheet: function() { this.hidden = true; return this; },
    showSheet: function() { this.hidden = false; return this; },
    isSheetHidden: function() { return this.hidden; },
    insertColumnsAfter: function(after, count) { this.maxCols += count; },
    deleteColumns: function(col, count) { this.maxCols -= count; },
    setColumnWidth: function(col, width) { this.columnWidths[col] = width; },
    getBandings: function() { return []; },
    getProtections: function(type) { return this.protections.slice(); },
    hiddenColumns: {},
    hideColumns: function(colIndex, numCols) {
      numCols = numCols || 1;
      for (var c = 0; c < numCols; c++) this.hiddenColumns[colIndex + c] = true;
    },
    showColumns: function(colIndex, numCols) {
      numCols = numCols || 1;
      for (var c = 0; c < numCols; c++) delete this.hiddenColumns[colIndex + c];
    },
    isColumnHiddenByUser: function(col) { return !!this.hiddenColumns[col]; },
    activated: false,
    activate: function() { this.activated = true; return this; },
    hiddenRows: {},
    hideRows: function(rowIndex, numRows) {
      numRows = numRows || 1;
      for (var r = 0; r < numRows; r++) {
        this.hiddenRows[rowIndex + r] = true;
      }
    },
    showRows: function(rowIndex, numRows) {
      numRows = numRows || 1;
      for (var r = 0; r < numRows; r++) {
        delete this.hiddenRows[rowIndex + r];
      }
    },
    isRowHiddenByUser: function(rowIndex) {
      return !!this.hiddenRows[rowIndex];
    },

    lastRow: 0,
    getLastColumn: function() {
      var maxC = 0;
      for (var r in this.grid) {
        var rowObj = this.grid[r];
        for (var c in rowObj) {
          var colNum = parseInt(c, 10);
          if (rowObj[c] !== "" && rowObj[c] !== null && rowObj[c] !== undefined) {
            if (colNum > maxC) maxC = colNum;
          }
        }
      }
      return maxC;
    },
    getLastRow: function() {
      var maxR = 0;
      for (var r in this.grid) {
        var rowNum = parseInt(r, 10);
        if (rowNum > maxR) {
          // Check if at least one cell in row has content
          var rowObj = this.grid[r];
          var hasContent = false;
          for (var c in rowObj) {
            if (rowObj[c] !== "" && rowObj[c] !== null && rowObj[c] !== undefined) {
              hasContent = true;
              break;
            }
          }
          if (hasContent) maxR = rowNum;
        }
      }
      return maxR;
    },
    deleteRow: function(rowIndex) {
      delete this.grid[rowIndex];
      // Shift rows up
      var newGrid = {};
      for (var r in this.grid) {
        var rowNum = parseInt(r, 10);
        if (rowNum < rowIndex) {
          newGrid[rowNum] = this.grid[rowNum];
        } else if (rowNum > rowIndex) {
          newGrid[rowNum - 1] = this.grid[rowNum];
        }
      }
      this.grid = newGrid;
    },
    getRange: function(row, col, numRows, numCols) {
      numRows = numRows || 1;
      numCols = numCols || 1;
      var key = row + ":" + col + ":" + numRows + ":" + numCols;
      if (!this.ranges[key]) {
        this.ranges[key] = createMockRange(this, row, col, numRows, numCols);
      }
      return this.ranges[key];
    },
    conditionalFormatRules: [],
    getConditionalFormatRules: function() { return this.conditionalFormatRules; },
    setConditionalFormatRules: function(rules) { this.conditionalFormatRules = rules; }
  };
}

function createMockSpreadsheet() {
  var namedRanges = {};
  var ss = {
    id: "mock-spreadsheet-12345",
    sheets: [],
    namedRanges: namedRanges,
    lastToast: null,
    toast: function(msg, title, timeout) {
      this.lastToast = { msg: msg, title: title };
    },

    getId: function() { return this.id; },
    getUrl: function() { return "https://docs.google.com/spreadsheets/d/" + this.id + "/edit"; },
    getOwner: function() {
      return {
        getEmail: function() { return "owner@example.com"; }
      };
    },
    getSheets: function() { return this.sheets.slice(); },
    getSheetByName: function(name) {
      for (var i = 0; i < this.sheets.length; i++) {
        if (this.sheets[i].getName() === name) return this.sheets[i];
      }
      return null;
    },
    insertSheet: function(name) {
      var s = createMockSheet(name, this);
      this.sheets.push(s);
      return s;
    },
    deleteSheet: function(sheet) {
      var idx = this.sheets.indexOf(sheet);
      if (idx !== -1) {
        this.sheets.splice(idx, 1);
      }
    },
    setNamedRange: function(name, range) {
      this.namedRanges[name] = range;
    },
    getRangeByName: function(name) {
      return this.namedRanges[name] || null;
    },
    removeNamedRange: function(name) {
      delete this.namedRanges[name];
    }
  };
  ss.sheets.push(createMockSheet("Sheet1", ss));
  return ss;
}

function setupGasGlobals(mockSpreadsheet) {
  var docProps = {};
  var batchUpdateCalls = [];
  var mockTables = [];

  var mockUi = {
    shownSidebar: null,
    menu: null,
    createMenu: function(name) {
      var items = [];
      this.menu = {
        name: name,
        items: items,
        addItem: function(caption, fn) { items.push({ caption: caption, fn: fn }); return this; },
        addSeparator: function() { items.push({ separator: true }); return this; },
        addToUi: function() { return this; }
      };
      return this.menu;
    },
    showSidebar: function(htmlOutput) {
      this.shownSidebar = htmlOutput;
    },
    showModalDialog: function(htmlOutput, title) {
      this.shownDialog = { html: htmlOutput, title: title };
    },
    lastAlert: null,
    alertResponse: null,
    alert: function(title, msg, buttons) {
      this.lastAlert = { title: title, msg: msg, buttons: buttons };
      if (this.alertResponse) return this.alertResponse;
      if (buttons === this.ButtonSet.YES_NO) return this.Button.NO;
      return "OK";
    },
    lastPrompt: null,
    promptResponse: null,
    prompt: function(title, msg, buttons) {
      this.lastPrompt = { title: title, msg: msg, buttons: buttons };
      if (this.promptResponse) return this.promptResponse;
      return {
        getSelectedButton: function() { return "OK"; },
        getResponseText: function() { return ""; }
      };
    },
    Button: { OK: "OK", CANCEL: "CANCEL", YES: "YES", NO: "NO", CLOSE: "CLOSE" },
    ButtonSet: { OK: "OK", OK_CANCEL: "OK_CANCEL", YES_NO: "YES_NO" }
  };

  global.SpreadsheetApp = {
    getActiveSpreadsheet: function() { return mockSpreadsheet; },
    getUi: function() { return mockUi; },
    BandingTheme: { LIGHT_GREY: "LIGHT_GREY" },
    ProtectionType: { RANGE: "RANGE", SHEET: "SHEET" },
    newConditionalFormatRule: function() {
      var rule = {
        formula: null,
        bg: null,
        color: null,
        bold: false,
        ranges: [],
        whenFormulaSatisfied: function(f) { this.formula = f; return this; },
        whenTextEqualTo: function(text) { this.textEqualTo = text; return this; },
        whenTextContains: function(text) { this.textContains = text; return this; },
        whenTextStartsWith: function(text) { this.textStartsWith = text; return this; },
        getBooleanCondition: function() {
          var self = this;
          var type = self.formula ? "CUSTOM_FORMULA" : (self.textEqualTo !== undefined ? "TEXT_EQUAL_TO" : (self.textStartsWith !== undefined ? "TEXT_STARTS_WITH" : (self.textContains !== undefined ? "TEXT_CONTAINS" : null)));
          return type ? {
            getCriteriaType: function() { return type; },
            getCriteriaValues: function() { return [self.formula || self.textEqualTo || self.textStartsWith || self.textContains]; }
          } : null;
        },
        getRanges: function() { return this.ranges; },
        setBackground: function(bg) { this.bg = bg; return this; },
        setFontColor: function(c) { this.color = c; return this; },
        setBold: function(b) { this.bold = b; return this; },
        setRanges: function(r) { this.ranges = r; return this; },
        build: function() { return this; }
      };
      return rule;
    },
    newDataValidation: function() {
      var rule = {
        allowedValues: [],
        showDropdown: false,
        allowInvalid: true,
        requireValueInList: function(values, showDropdown) {
          this.allowedValues = values;
          this.showDropdown = showDropdown;
          return this;
        },
        setAllowInvalid: function(allow) {
          this.allowInvalid = allow;
          return this;
        },
        build: function() { return this; }
      };
      return rule;
    }
  };

  global.HtmlService = {
    createTemplateFromFile: function(filename) {
      return {
        filename: filename,
        evaluate: function() {
          var tpl = this;
          return {
            filename: filename,
            template: tpl,
            title: "",
            width: 300,
            setTitle: function(t) { this.title = t; return this; },
            setWidth: function(w) { this.width = w; return this; },
            setHeight: function(h) { this.height = h; return this; }
          };
        }
      };
    },
    createHtmlOutputFromFile: function(filename) {
      return {
        filename: filename,
        title: "",
        width: 300,
        setTitle: function(t) { this.title = t; return this; },
        setWidth: function(w) { this.width = w; return this; }
      };
    }
  };

  global.Session = {
    getActiveUser: function() {
      return {
        getEmail: function() { return "alice@example.com"; }
      };
    },
    getEffectiveUser: function() {
      return {
        getEmail: function() { return "owner@example.com"; }
      };
    }
  };

  global.Logger = {
    logs: [],
    log: function(msg) {
      if (!this.logs) this.logs = [];
      this.logs.push(String(msg));
    },
    getLogs: function() { return this.logs || []; },
    clear: function() { this.logs = []; }
  };

  global.Sheets = {
    batchUpdateCalls: batchUpdateCalls,
    Spreadsheets: {
      tables: mockTables,
      batchUpdate: function(resource, spreadsheetId) {
        batchUpdateCalls.push({ resource: resource, spreadsheetId: spreadsheetId });
        if (resource && resource.requests) {
          for (var i = 0; i < resource.requests.length; i++) {
            var req = resource.requests[i];
            if (req.updateTable && req.updateTable.table) {
              for (var ut = 0; ut < mockTables.length; ut++) {
                if (mockTables[ut].tableId !== req.updateTable.table.tableId) continue;
                if (req.updateTable.table.name) mockTables[ut].name = req.updateTable.table.name;
                if (req.updateTable.table.range) mockTables[ut].range = req.updateTable.table.range;
                if (req.updateTable.table.columnProperties) mockTables[ut].columnProperties = req.updateTable.table.columnProperties;
              }
            } else if (req.addTable && req.addTable.table) {
              var newTable = req.addTable.table;
              for (var at = 0; at < mockTables.length; at++) {
                if (mockTables[at].name === newTable.name) {
                  throw new Error("API call to sheets.spreadsheets.batchUpdate failed with error: Internal error encountered.");
                }
              }
              mockTables.push({ tableId: "tbl" + (mockTables.length + 1), name: newTable.name, sheetId: newTable.range.sheetId, range: newTable.range, columnProperties: newTable.columnProperties });
            } else if (req.copyPaste) {
              var cp = req.copyPaste;
              var srcSheet = null;
              var dstSheet = null;
              for (var s = 0; s < mockSpreadsheet.sheets.length; s++) {
                if (mockSpreadsheet.sheets[s].getSheetId() === cp.source.sheetId) srcSheet = mockSpreadsheet.sheets[s];
                if (mockSpreadsheet.sheets[s].getSheetId() === cp.destination.sheetId) dstSheet = mockSpreadsheet.sheets[s];
              }
              if (srcSheet && dstSheet) {
                var sRow = cp.source.startRowIndex + 1;
                var sCol = cp.source.startColumnIndex + 1;
                var dRow = cp.destination.startRowIndex + 1;
                var dCol = cp.destination.startColumnIndex + 1;
                var val = (srcSheet.grid[sRow] && srcSheet.grid[sRow][sCol] !== undefined) ? srcSheet.grid[sRow][sCol] : "";
                if (!dstSheet.grid[dRow]) dstSheet.grid[dRow] = {};
                dstSheet.grid[dRow][dCol] = val;
              }
            } else if (req.updateCells && req.updateCells.rows) {
              var uc = req.updateCells;
              var targetSheet = null;
              for (var ts = 0; ts < mockSpreadsheet.sheets.length; ts++) {
                if (mockSpreadsheet.sheets[ts].getSheetId() === uc.range.sheetId) targetSheet = mockSpreadsheet.sheets[ts];
              }
              if (targetSheet && uc.rows[0] && uc.rows[0].values && uc.rows[0].values[0]) {
                var cellObj = uc.rows[0].values[0];
                var cellVal = (cellObj.chipRuns && cellObj.chipRuns[0] && cellObj.chipRuns[0].chip && cellObj.chipRuns[0].chip.personProperties)
                  ? cellObj.chipRuns[0].chip.personProperties.email
                  : (cellObj.userEnteredValue ? cellObj.userEnteredValue.stringValue : "");
                var rIdx = uc.range.startRowIndex + 1;
                var cIdx = uc.range.startColumnIndex + 1;
                if (!targetSheet.grid[rIdx]) targetSheet.grid[rIdx] = {};
                targetSheet.grid[rIdx][cIdx] = cellVal;
              }
            }
          }
        }
        return { replies: [{ addTable: { table: {} } }] };
      },
      get: function(spreadsheetId, options) {
        if (options && options.fields && String(options.fields).indexOf("tables") !== -1) {
          return {
            sheets: mockSpreadsheet.sheets.map(function(sh) {
              return {
                properties: { sheetId: sh.getSheetId(), title: sh.getName() },
                tables: mockTables.filter(function(t) { return t.sheetId === sh.getSheetId(); }).map(function(t) {
                  return { tableId: t.tableId, name: t.name, range: t.range };
                })
              };
            })
          };
        }
        var ranges = (options && options.ranges) ? options.ranges : [];
        var resultSheets = [];
        for (var i = 0; i < ranges.length; i++) {
          var rStr = ranges[i];
          var parts = rStr.split("!");
          var sheetName = parts[0].replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
          var a1 = parts[1] || "";
          var sheet = null;
          for (var s = 0; s < mockSpreadsheet.sheets.length; s++) {
            if (mockSpreadsheet.sheets[s].getName() === sheetName) {
              sheet = mockSpreadsheet.sheets[s];
              break;
            }
          }
          var rowData = [];
          if (sheet) {
            var match = a1.match(/([A-Z]+)(\d+):([A-Z]+)(\d*)/);
            var startRow = match ? parseInt(match[2], 10) : 2;
            var endRow = (match && match[4]) ? parseInt(match[4], 10) : Math.max(sheet.getLastRow(), startRow);
            for (var r = startRow; r <= endRow; r++) {
              var val = (sheet.grid[r] && sheet.grid[r][1] !== undefined) ? sheet.grid[r][1] : "";
              var cellObj = {
                formattedValue: (typeof val === "object" && val.name) ? val.name : String(val),
                chipRuns: []
              };
              if (val && typeof val === "object" && val.personProperties) {
                cellObj.chipRuns.push({
                  startIndex: 0,
                  endIndex: cellObj.formattedValue.length,
                  chip: { personProperties: val.personProperties }
                });
              } else if (val && typeof val === "object" && val.chipRuns) {
                cellObj.chipRuns = val.chipRuns;
              } else if (typeof val === "string" && val.indexOf("@") !== -1) {
                cellObj.chipRuns.push({
                  startIndex: 0,
                  endIndex: val.length,
                  chip: { personProperties: { email: val, name: val } }
                });
              }
              rowData.push({ values: [cellObj] });
            }
          }
          resultSheets.push({ data: [{ rowData: rowData }] });
        }
        return { sheets: resultSheets };
      }
    }
  };

  var userProps = {};
  var userCache = {};

  global.PropertiesService = {
    getDocumentProperties: function() {
      return {
        getProperty: function(key) { return docProps[key] || null; },
        setProperty: function(key, val) { docProps[key] = String(val); },
        deleteProperty: function(key) { delete docProps[key]; },
        getProperties: function() { return Object.assign({}, docProps); }
      };
    },
    getUserProperties: function() {
      return {
        getProperty: function(key) { return userProps[key] || null; },
        setProperty: function(key, val) { userProps[key] = String(val); },
        deleteProperty: function(key) { delete userProps[key]; },
        getProperties: function() { return Object.assign({}, userProps); }
      };
    }
  };

  // LockService mock: one shared slot per lock scope, like Apps Script.
  // Tests simulate another execution by setting LockService.__state.<scope>.held = true.
  var lockState = {
    document: { held: false },
    script: { held: false },
    documentLockAvailable: true
  };
  function createMockLock(slot) {
    var mine = false;
    return {
      tryLock: function(timeoutMs) {
        if (slot.held) return false;
        slot.held = true;
        mine = true;
        return true;
      },
      waitLock: function(timeoutMs) {
        if (!this.tryLock(timeoutMs)) throw new Error("Lock timeout: another process was holding the lock for too long.");
      },
      releaseLock: function() {
        if (mine) { slot.held = false; mine = false; }
      },
      hasLock: function() { return mine; }
    };
  }
  global.LockService = {
    getDocumentLock: function() { return lockState.documentLockAvailable ? createMockLock(lockState.document) : null; },
    getScriptLock: function() { return createMockLock(lockState.script); },
    __state: lockState
  };

  global.CacheService = {
    getUserCache: function() {
      return {
        get: function(key) { return userCache[key] || null; },
        put: function(key, val, ttl) { userCache[key] = String(val); },
        remove: function(key) { delete userCache[key]; }
      };
    }
  };

  var mockTasklists = [{ id: "@default", title: "My Tasks" }];
  var mockTasksStore = { "@default": [] };

  global.Tasks = {
    Tasklists: {
      list: function() {
        return { items: mockTasklists.slice() };
      },
      get: function(tasklistId) {
        for (var i = 0; i < mockTasklists.length; i++) {
          if (mockTasklists[i].id === tasklistId) return mockTasklists[i];
        }
        return { id: tasklistId, title: "Custom Tasklist" };
      },
      insert: function(resource) {
        var id = "tl_" + Math.random().toString(36).substring(2, 9);
        var created = Object.assign({ id: id }, resource);
        mockTasklists.push(created);
        mockTasksStore[id] = [];
        return created;
      }
    },
    Tasks: {
      list: function(tasklistId, optionalArgs) {
        var list = mockTasksStore[tasklistId] || [];
        return { items: list.slice() };
      },
      get: function(tasklistId, taskId) {
        var list = mockTasksStore[tasklistId] || [];
        for (var i = 0; i < list.length; i++) {
          if (list[i].id === taskId) return list[i];
        }
        throw new Error("Task not found: " + taskId);
      },
      insert: function(taskResource, tasklistId) {
        var listId = tasklistId || "@default";
        if (!mockTasksStore[listId]) mockTasksStore[listId] = [];
        var id = "gtask_" + Math.random().toString(36).substring(2, 9);
        var created = Object.assign({
          id: id,
          status: "needsAction",
          title: "",
          updated: new Date().toISOString()
        }, taskResource);
        mockTasksStore[listId].push(created);
        return created;
      },
      patch: function(taskResource, tasklistId, taskId) {
        var listId = tasklistId || "@default";
        var list = mockTasksStore[listId] || [];
        for (var i = 0; i < list.length; i++) {
          if (list[i].id === taskId) {
            Object.assign(list[i], taskResource, { updated: new Date().toISOString() });
            return list[i];
          }
        }
        throw new Error("Task not found to patch: " + taskId);
      },
      update: function(taskResource, tasklistId, taskId) {
        return this.patch(taskResource, tasklistId, taskId);
      },
      remove: function(tasklistId, taskId) {
        var listId = tasklistId || "@default";
        var list = mockTasksStore[listId] || [];
        for (var i = 0; i < list.length; i++) {
          if (list[i].id === taskId) {
            list.splice(i, 1);
            return;
          }
        }
      },
      delete: function(tasklistId, taskId) {
        return this.remove(tasklistId, taskId);
      }
    }
  };

  var mockSentEmails = [];
  global.MailApp = {
    sendEmail: function(opt1, opt2, opt3) {
      if (typeof opt1 === "object") {
        mockSentEmails.push(opt1);
      } else {
        mockSentEmails.push({ to: opt1, subject: opt2, body: opt3 });
      }
    },
    getSentEmails: function() { return mockSentEmails.slice(); },
    clearSentEmails: function() { mockSentEmails = []; }
  };

  var mockTriggers = [];
  global.ScriptApp = {
    WeekDay: { MONDAY: "MONDAY", TUESDAY: "TUESDAY", WEDNESDAY: "WEDNESDAY", THURSDAY: "THURSDAY", FRIDAY: "FRIDAY", SATURDAY: "SATURDAY", SUNDAY: "SUNDAY" },
    getProjectTriggers: function() { return mockTriggers.slice(); },
    deleteTrigger: function(trigger) {
      var idx = mockTriggers.indexOf(trigger);
      if (idx !== -1) mockTriggers.splice(idx, 1);
    },
    newTrigger: function(handlerFunction) {
      var trig = {
        handler: handlerFunction,
        getHandlerFunction: function() { return this.handler; },
        forSpreadsheet: function(ss) { return this; },
        onEdit: function() { return this; },
        timeBased: function() { return this; },
        everyDays: function(days) { return this; },
        everyHours: function(h) { this.everyHoursValue = h; return this; },
        everyWeeks: function(weeks) { return this; },
        onWeekDay: function(day) { this.weekDay = day; return this; },
        atHour: function(hour) { this.hour = hour; return this; },
        create: function() {
          mockTriggers.push(this);
          return this;
        }
      };
      return trig;
    }
  };

  var mockCalendars = {};
  function createMockCalendar(calId, calName) {
    var eventsStore = {};
    return {
      id: calId,
      name: calName,
      getId: function() { return this.id; },
      getName: function() { return this.name; },
      createAllDayEvent: function(title, date, options) {
        options = options || {};
        var evtId = "cal_evt_" + Math.random().toString(36).substring(2, 9);
        var guestsList = [];
        if (options.guests) {
          var emails = String(options.guests).split(",");
          for (var i = 0; i < emails.length; i++) {
            (function(guestEmail) {
              var em = String(guestEmail || "").trim();
              if (em) guestsList.push({ getEmail: function() { return em; } });
            })(emails[i]);
          }
        }
        var evt = {
          id: evtId,
          title: title,
          date: date,
          endDate: (options && options.endTime) ? options.endTime : null,
          isAllDay: options && options.isAllDay !== undefined ? options.isAllDay : true,
          location: options.location || "",
          description: options.description || "",
          guests: guestsList,
          getId: function() { return this.id; },
          getTitle: function() { return this.title; },
          setTitle: function(t) { this.title = t; return this; },
          getLocation: function() { return this.location; },
          setLocation: function(l) { this.location = l; return this; },
          getAllDayStartDate: function() { return this.date; },
          setAllDayDate: function(d) { this.date = d; this.isAllDay = true; return this; },
          getStartTime: function() { return this.date; },
          getEndTime: function() { return this.endDate || this.date; },
          setTime: function(start, end) { this.date = start; this.endDate = end; this.isAllDay = false; return this; },
          isAllDayEvent: function() { return this.isAllDay; },
          getDescription: function() { return this.description; },
          setDescription: function(d) { this.description = d; return this; },
          getGuestList: function() { return this.guests.slice(); },
          addGuest: function(email) {
            this.guests.push({ getEmail: function() { return email; } });
            return this;
          },
          deleteEvent: function() {
            delete eventsStore[this.id];
          }
        };
        eventsStore[evtId] = evt;
        return evt;
      },
      createEvent: function(title, startTime, endTime, options) {
        options = Object.assign({}, options || {}, { isAllDay: false, endTime: endTime });
        var created = this.createAllDayEvent(title, startTime, options);
        created.isAllDay = false;
        created.endDate = endTime;
        return created;
      },
      getEventById: function(evtId) {
        return eventsStore[evtId] || null;
      },
      createAllDayEventSeries: function(title, date, recurrence, options) {
        var series = this.createAllDayEvent(title, date, options);
        series.recurrence = recurrence;
        series.isSeries = true;
        series.deleteEventSeries = function() { delete eventsStore[this.id]; };
        return series;
      },
      getEventSeriesById: function(evtId) {
        var e = eventsStore[evtId];
        return e && e.isSeries ? e : null;
      },
      getEvents: function() {
        var res = [];
        for (var k in eventsStore) res.push(eventsStore[k]);
        return res;
      }
    };
  }

  var defaultCal = createMockCalendar("primary", "Team Calendar");
  mockCalendars["primary"] = defaultCal;

  var mockAcl = {};
  global.Calendar = {
    Acl: {
      insert: function(rule, calendarId, optionalArgs) {
        mockAcl[calendarId] = mockAcl[calendarId] || {};
        mockAcl[calendarId][rule.scope.value] = { rule: rule, args: optionalArgs || {} };
        return rule;
      },
      remove: function(calendarId, ruleId) {
        if (mockAcl[calendarId]) delete mockAcl[calendarId][String(ruleId).replace(/^user:/, "")];
      },
      __entries: function(calendarId) { return mockAcl[calendarId] || {}; }
    }
  };

  global.CalendarApp = {
    getDefaultCalendar: function() { return defaultCal; },
    getCalendarById: function(id) {
      var activeEmail = (typeof Session !== "undefined" && Session.getActiveUser) ? Session.getActiveUser().getEmail() : "";
      if (!id || id === "primary" || id === "@default" || (activeEmail && id.toLowerCase() === activeEmail.toLowerCase())) {
        return defaultCal;
      }
      if (!mockCalendars[id]) {
        mockCalendars[id] = createMockCalendar(id, "Custom Calendar " + id);
      }
      return mockCalendars[id];
    },
    newRecurrence: function() {
      return { rules: [], addYearlyRule: function() { this.rules.push("YEARLY"); return this; } };
    },
    createCalendar: function(name) {
      var id = "cal_" + Math.random().toString(36).substring(2, 9);
      var c = createMockCalendar(id, name);
      mockCalendars[id] = c;
      return c;
    }
  };
}

module.exports = {
  createMockSpreadsheet: createMockSpreadsheet,
  setupGasGlobals: setupGasGlobals,
  get MailApp() { return global.MailApp; }
};
