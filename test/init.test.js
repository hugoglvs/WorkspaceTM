const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');
Object.assign(global, require('../src/Tables'));
Object.assign(global, require('../src/Locations'));
global.schemaHeaderRow = require('../src/Schema').schemaHeaderRow;
global.schemaColumnLabel = require('../src/Schema').schemaColumnLabel;

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

test('freshInit should create all operational tables and set SCHEMA_VERSION = 1', () => {
  const result = freshInit(mockSpreadsheet);

  assert.strictEqual(result.success, true);
  assert.strictEqual(result.schemaVersion, 1);

  const expectedSheets = [
    'Locations',
    'Members',
    'Contracts',
    'Routines',
    'Leaves',
    'Presence Log',
    'Projects',
    'Tasks',
    'Events',
    'Backups',
    'Holidays',
    'Task Templates',
    'Recurring Tasks',
    'Settings'
  ];

  assert.deepStrictEqual(result.sheetsCreated, expectedSheets);

  // Check all sheets exist on spreadsheet
  for (const sheetName of expectedSheets) {
    const sheet = mockSpreadsheet.getSheetByName(sheetName);
    assert.ok(sheet, `Sheet ${sheetName} must exist`);
    assert.strictEqual(sheet.getFrozenRows(), 1, `Sheet ${sheetName} must have Row 1 frozen`);

    const config = schema.SCHEMA_SPEC[tableKeyOfName(sheetName)];
    const headerRange = sheet.getRange(1, 1, 1, config.columns.length);

    // Verify header text (readable labels where defined) and typography
    assert.deepStrictEqual(headerRange.values, [schema.schemaHeaderRow(config)]);
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
    const spec = schema.SCHEMA_SPEC[tableKeyOfName(expectedSheets[i])];
    assert.strictEqual(tableReq.columnProperties.length, spec.columns.length);

    // Verify that PEOPLE, DATE, DROPDOWN, NUMERIC types are cast in columnProperties,
    // and that table column names use the readable header labels
    for (let p = 0; p < tableReq.columnProperties.length; p++) {
      const colProp = tableReq.columnProperties[p];
      const key = spec.columns[colProp.columnIndex];
      assert.strictEqual(colProp.columnName, schema.schemaColumnLabel(spec, key));
      assert.strictEqual(colProp.columnType, spec.columnTypes[key], `Column ${key} must have type ${spec.columnTypes[key]}`);
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

test('Tasks: technical columns hidden; a table whose columns moved is archived and rebuilt', () => {
  freshInit(mockSpreadsheet);
  var tasks = mockSpreadsheet.getSheetByName('Tasks');
  var cols = schema.SCHEMA_SPEC.tasks.columns;
  ['task_id', 'google_task_id', 'origin'].forEach(function(k) {
    assert.ok(tasks.isColumnHiddenByUser(cols.indexOf(k) + 1), k + ' hidden');
  });
  assert.ok(!tasks.isColumnHiddenByUser(cols.indexOf('title') + 1));

  // Former layout (ID first, Done by, Recurring) with data
  var old = ['ID', 'Project ID', 'Assigned to', 'Title', 'Description', 'Due date', 'Status', 'Google Task ID', 'Type', 'Done on', 'Done by', 'Recurring'];
  tasks.getRange(1, 1, 1, old.length).setValues([old]);
  tasks.getRange(2, 1, 1, old.length).setValues([['TSK-1', 'P', 'Alice', 'Old task', '', '2026-01-01', 'Todo', '', '', '', '', '']]);

  freshInit(mockSpreadsheet);

  var archived = mockSpreadsheet.getSheets().filter(s => /^Archive – Tasks \d{4}-\d{2}-\d{2}$/.test(s.getName()))[0];
  assert.ok(archived, 'old tab archived');
  assert.strictEqual(archived.getRange(2, 4).getValue(), 'Old task', 'old data kept in the archive');
  var fresh = mockSpreadsheet.getSheetByName('Tasks');
  assert.notStrictEqual(fresh, archived);
  assert.deepStrictEqual(fresh.getRange(1, 1, 1, cols.length).getValues()[0], schema.schemaHeaderRow(schema.SCHEMA_SPEC.tasks));
  var col = k => cols.indexOf(k) + 1;
  assert.strictEqual(fresh.getRange(2, col('title')).getValue(), 'Old task', 'rows copied into the new table (ADR 0053)');
  assert.strictEqual(fresh.getRange(2, col('task_id')).getValue(), 'TSK-1');
  assert.strictEqual(fresh.getRange(2, col('project')).getValue(), 'P');
  assert.strictEqual(fresh.getRange(2, col('due_date')).getValue(), '2026-01-01');

  // Upgrading again does not archive anything more
  var count = mockSpreadsheet.getSheets().length;
  fresh.getRange(2, 2).setValue('Kept');
  freshInit(mockSpreadsheet);
  assert.strictEqual(mockSpreadsheet.getSheets().length, count);
  assert.strictEqual(mockSpreadsheet.getSheetByName('Tasks').getRange(2, 2).getValue(), 'Kept');
});

test('Projects with the former ID column: archived, rows copied without the ID (ADR 0053)', () => {
  var own = createMockSpreadsheet();
  setupGasGlobals(own);
  freshInit(own);
  var projects = own.getSheetByName('Projects');
  var old = ['ID', 'Project', 'Status', 'Start', 'Target date', 'Lead', 'Backup', 'In charge today'];
  projects.getRange(1, 1, 1, old.length).setValues([old]);
  own.getSheetByName('Tasks').getRange(2, 1, 1, 4).setValues([['', 'Old task', '', 'PRJ-2']]);
  projects.getRange(2, 1, 2, old.length).setValues([
    ['PRJ-1', 'Apollo', 'Active', '2026-01-01', '', 'Alice', 'Bob', 'Alice'],
    ['PRJ-2', 'Borealis', 'Planning', '2026-02-01', '', 'Carol', '', '']
  ]);
  freshInit(own);
  var fresh = own.getSheetByName('Projects');
  assert.deepStrictEqual(fresh.getRange(1, 1, 1, 7).getValues()[0], schema.schemaHeaderRow(schema.SCHEMA_SPEC.projects));
  assert.deepStrictEqual(fresh.getRange(2, 1, 2, 6).getValues(), [
    ['Apollo', 'Active', '2026-01-01', '', 'Alice', 'Bob'],
    ['Borealis', 'Planning', '2026-02-01', '', 'Carol', '']
  ]);
  assert.ok(own.getSheets().some(s => /^Archive – Projects /.test(s.getName())));
  assert.strictEqual(own.getSheetByName('Tasks').getRange(2, 4).getValue(), 'Borealis', 'former ID replaced by the name');
});

test('Tasks without data in the former 13-column layout: headers rewritten in place, the extra column cleared', () => {
  var own = createMockSpreadsheet();
  setupGasGlobals(own);
  freshInit(own);
  var tasks = own.getSheetByName('Tasks');
  var old = ['ID', 'Project ID', 'Assigned to', 'Title', 'Description', 'Deadline', 'Status', 'Google Task ID', 'Type', 'Done on', 'Done by', 'Recurring', 'Planned date'];
  tasks.getRange(1, 1, 1, old.length).setValues([old]);
  freshInit(own);
  var same = own.getSheetByName('Tasks');
  assert.strictEqual(same, tasks, 'no archive without data');
  assert.strictEqual(same.getRange(1, 1).getValue(), 'Type');
  assert.strictEqual(same.getRange(1, 13).getValue(), '', 'former 13th column cleared');
});
