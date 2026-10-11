const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');

// Set up globals before requiring source modules
const mockSpreadsheet = createMockSpreadsheet();
setupGasGlobals(mockSpreadsheet);

// Load source modules into global scope (as Apps Script does)
const schema = require('../src/Schema');
global.CURRENT_SCHEMA_VERSION = schema.CURRENT_SCHEMA_VERSION;
global.FONT_FAMILY = schema.FONT_FAMILY;
global.FONT_SIZE = schema.FONT_SIZE;
global.SCHEMA_SPEC = schema.SCHEMA_SPEC;
global.DATE_FORMAT = schema.DATE_FORMAT;

const typography = require('../src/Typography');
global.applyTypography = typography.applyTypography;
global.styleHeaderRow = typography.styleHeaderRow;
global.applyTableBanding = typography.applyTableBanding;
global.createGoogleSheetsTable = typography.createGoogleSheetsTable;

const { freshInit } = require('../src/Init');
const { init, getSchemaVersion, getLibraryVersion } = require('../src/Code');

test('freshInit should create all 8 operational tables and set SCHEMA_VERSION = 1', () => {
  const result = freshInit(mockSpreadsheet);

  assert.strictEqual(result.success, true);
  assert.strictEqual(result.schemaVersion, 1);

  const expectedSheets = [
    'tbl_members',
    'tbl_tenures',
    'tbl_presence_baselines',
    'tbl_leaves',
    'tbl_presence_log',
    'tbl_projects',
    'tbl_tasks',
    'tbl_events',
    'tbl_backups'
  ];

  assert.deepStrictEqual(result.sheetsCreated, expectedSheets);

  // Check all sheets exist on spreadsheet
  for (const sheetName of expectedSheets) {
    const sheet = mockSpreadsheet.getSheetByName(sheetName);
    assert.ok(sheet, `Sheet ${sheetName} must exist`);
    assert.strictEqual(sheet.getFrozenRows(), 1, `Sheet ${sheetName} must have Row 1 frozen`);

    const config = schema.SCHEMA_SPEC[sheetName];
    const headerRange = sheet.getRange(1, 1, 1, config.columns.length);

    // Verify header text and typography
    assert.deepStrictEqual(headerRange.values, [config.columns]);
    assert.strictEqual(sheet.getRange(1, 1, sheet.getMaxRows(), config.columns.length).fontFamily, 'Google Sans');
    assert.strictEqual(sheet.getRange(1, 1, sheet.getMaxRows(), config.columns.length).fontSize, 12);
  }

  // Verify Document Properties
  assert.strictEqual(getSchemaVersion(), 1);
  assert.strictEqual(getLibraryVersion(), 1);

  // Verify Sheets API v4 addTable batchUpdate was invoked for all operational tables
  assert.strictEqual(global.Sheets.batchUpdateCalls.length, expectedSheets.length);
  for (let i = 0; i < expectedSheets.length; i++) {
    const call = global.Sheets.batchUpdateCalls[i];
    assert.strictEqual(call.spreadsheetId, 'mock-spreadsheet-12345');
    const tableReq = call.resource.requests[0].addTable.table;
    assert.strictEqual(tableReq.name, expectedSheets[i]);
    assert.strictEqual(tableReq.range.startRowIndex, 0);
    assert.strictEqual(tableReq.range.startColumnIndex, 0);
    assert.ok(Array.isArray(tableReq.columnProperties), 'Must have columnProperties array');
    assert.strictEqual(tableReq.columnProperties.length, schema.SCHEMA_SPEC[expectedSheets[i]].columns.length);

    // Verify that PEOPLE, DATE, DROPDOWN, NUMERIC types are cast in columnProperties
    const spec = schema.SCHEMA_SPEC[expectedSheets[i]];
    for (let p = 0; p < tableReq.columnProperties.length; p++) {
      const colProp = tableReq.columnProperties[p];
      const expectedType = spec.columnTypes[colProp.columnName];
      assert.strictEqual(colProp.columnType, expectedType, `Column ${colProp.columnName} must have type ${expectedType}`);
    }
  }
});

test('freshInit should throw actionable error if Sheets Advanced API is not present', () => {
  const fallbackSpreadsheet = createMockSpreadsheet();
  const savedSheets = global.Sheets;
  delete global.Sheets; // simulate environment without Sheets advanced service

  try {
    assert.throws(
      () => freshInit(fallbackSpreadsheet),
      /Google Sheets API/
    );
  } finally {
    global.Sheets = savedSheets;
  }
});
