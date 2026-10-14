const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');
Object.assign(global, require('../src/Tables'));
Object.assign(global, require('../src/Locations'));
global.schemaHeaderRow = require('../src/Schema').schemaHeaderRow;
global.schemaColumnLabel = require('../src/Schema').schemaColumnLabel;

let mockSpreadsheet;

test.beforeEach(() => {
  mockSpreadsheet = createMockSpreadsheet();
  setupGasGlobals(mockSpreadsheet);
});

const schema = require('../src/Schema');
global.CURRENT_SCHEMA_VERSION = schema.CURRENT_SCHEMA_VERSION;
global.FONT_FAMILY = schema.FONT_FAMILY;
global.FONT_SIZE = schema.FONT_SIZE;
global.SCHEMA_SPEC = schema.SCHEMA_SPEC;
global.DATE_FORMAT = schema.DATE_FORMAT;

const typography = require('../src/Typography');
global.applyTypography = typography.applyTypography;
global.createGoogleSheetsTable = typography.createGoogleSheetsTable;
const syncEngine = require('../src/SyncEngine');
global.copyCellChip = syncEngine.copyCellChip;
const { freshInit } = require('../src/Init');
global.freshInit = freshInit;
const sidebar = require('../src/Sidebar');
global.extractCellValue = sidebar.extractCellValue;
global.isUserMemberMatch = sidebar.isUserMemberMatch;

Object.assign(global, require('../src/Layout'));
Object.assign(global, require('../src/ViewBuilder'));
Object.assign(global, require('../src/PresenceEngine'));
Object.assign(global, require('../src/KpiEngine'));
const code = require('../src/Code');

const RANGES_HEADER = ['range_id', 'member', 'range_type', 'start_date', 'start_period', 'end_date', 'end_period', 'day_of_week', 'location_am', 'location_pm', 'status'];

// The live "[NEW] ACN" tbl_presence_ranges on 2026-10-12, trimmed to 3 rows
function addMergedRangesTable(spreadsheet) {
  var sheet = spreadsheet.insertSheet('tbl_presence_ranges');
  sheet.getRange(1, 1, 1, 11).setValues([RANGES_HEADER]);
  sheet.getRange(2, 1, 3, 11).setValues([
    ['base_tm8krbd', 'Hugo Gonçalves', 'Baseline', '2026-09-02', '', '2026-10-09', '', 'Monday', 'Remote', 'Remote', 'Approved'],
    ['leave_q45eloj', 'Hugo Gonçalves', 'Leave', '2026-10-01', 'AM', '2026-10-05', 'PM', 'All', 'Vacation', 'Vacation', 'Requested'],
    ['base_egk5kvt', 'Hugo Gonçalves', 'Baseline', '2026-10-09', '', '2026-10-16', '', 'Monday', 'Remote', 'Remote', 'Approved']
  ]);
  return sheet;
}

function rows(spreadsheet, name, cols) {
  return spreadsheet.getSheetByName(name).getRange(2, 1, 20, cols).getValues().filter(r => r[0] || r[1]);
}

test('restorePresenceTablesFromRanges: moves merged ranges back into baselines and leaves, then archives the merged table', () => {
  freshInit(mockSpreadsheet);
  addMergedRangesTable(mockSpreadsheet);

  var res = restorePresenceTablesFromRanges(mockSpreadsheet);

  assert.strictEqual(res.restoredCount, 3);
  assert.deepStrictEqual(rows(mockSpreadsheet, 'Routines', 7), [
    ['base_tm8krbd', 'Hugo Gonçalves', '2026-09-02', '2026-10-09', 'Monday', 'Remote', 'Remote'],
    ['base_egk5kvt', 'Hugo Gonçalves', '2026-10-09', '2026-10-16', 'Monday', 'Remote', 'Remote']
  ]);
  assert.deepStrictEqual(rows(mockSpreadsheet, 'Leaves', 8), [
    ['leave_q45eloj', 'Hugo Gonçalves', '2026-10-01', 'AM', '2026-10-05', 'PM', 'Vacation', 'Requested']
  ]);
  assert.strictEqual(mockSpreadsheet.getSheetByName('tbl_presence_ranges'), null);
  assert.ok(mockSpreadsheet.getSheetByName('Archive – Presence Ranges').isSheetHidden());

  assert.strictEqual(restorePresenceTablesFromRanges(mockSpreadsheet).restoredCount, 0, 'idempotent');
});

test('restorePresenceTablesFromRanges: never duplicates a row whose ID already exists', () => {
  freshInit(mockSpreadsheet);
  mockSpreadsheet.getSheetByName('Routines').getRange(2, 1, 1, 7).setValues([
    ['base_tm8krbd', 'Hugo Gonçalves', '2026-09-02', '', 'Monday', 'Remote', 'Remote']
  ]);
  addMergedRangesTable(mockSpreadsheet);

  restorePresenceTablesFromRanges(mockSpreadsheet);

  var ids = rows(mockSpreadsheet, 'Routines', 7).map(r => r[0]);
  assert.deepStrictEqual(ids, ['base_tm8krbd', 'base_egk5kvt']);
});

test('setupOrUpgrade: shows only Members, Weekly Matrix, Projects, Tasks and Events', () => {
  mockSpreadsheet.insertSheet('calc_kpis');
  freshInit(mockSpreadsheet);
  mockSpreadsheet.getSheetByName('Members').getRange(2, 1, 1, 7).setValues([
    ['Hugo Gonçalves', '', '', 'Data Analyst', 'FAL ME', 'Active', '']
  ]);
  addMergedRangesTable(mockSpreadsheet);
  mockSpreadsheet.insertSheet('Sheet1');

  code.setupOrUpgrade(mockSpreadsheet);

  var visible = mockSpreadsheet.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName()).sort();
  assert.deepStrictEqual(visible, ['Events', 'Members', 'Projects', 'Tasks', 'Weekly Presence']);
  assert.strictEqual(mockSpreadsheet.getSheetByName('calc_kpis'), null, 'calc_kpis deleted');
  assert.strictEqual(mockSpreadsheet.getSheetByName('Sheet1'), null, 'empty Sheet1 deleted');
  assert.strictEqual(rows(mockSpreadsheet, 'Routines', 7).length, 2, 'merged data restored');
});

test('setupOrUpgrade: keeps a Sheet1 that has content', () => {
  mockSpreadsheet.getSheetByName('Sheet1').getRange(1, 1).setValue('my notes');

  code.setupOrUpgrade(mockSpreadsheet);

  assert.ok(mockSpreadsheet.getSheetByName('Sheet1'));
});

test('freshInit: recreates a table whose name is still held by an archived tab', () => {
  freshInit(mockSpreadsheet);
  mockSpreadsheet.getSheetByName('Leaves').setName('Archive – Leaves');

  assert.doesNotThrow(() => freshInit(mockSpreadsheet));

  var bySheet = {};
  Sheets.Spreadsheets.tables.forEach(t => {
    var sheet = mockSpreadsheet.getSheets().find(s => s.getSheetId() === t.sheetId);
    bySheet[sheet.getName()] = t.name;
  });
  assert.strictEqual(bySheet['Leaves'], 'Leaves', 'new table takes the canonical name');
  assert.notStrictEqual(bySheet['Archive – Leaves'], 'Leaves', 'archived table was renamed out of the way');
});

function makeLegacyWorkbook(spreadsheet) {
  // The live sheet before ADR 0039: technical tab names and technical table names
  freshInit(spreadsheet);
  Object.keys(TABLE_REGISTRY).forEach(function(key) {
    var sheet = spreadsheet.getSheetByName(TABLE_REGISTRY[key].sheetName);
    if (!sheet) return;
    sheet.setName(TABLE_REGISTRY[key].legacyName);
    Sheets.Spreadsheets.tables.filter(t => t.sheetId === sheet.getSheetId()).forEach(t => { t.name = TABLE_REGISTRY[key].legacyName; });
    if (SCHEMA_SPEC[key]) sheet.getRange(1, 1, 1, SCHEMA_SPEC[key].columns.length).setValues([SCHEMA_SPEC[key].columns]);
  });
  PropertiesService.getDocumentProperties().getProperties && Object.keys(PropertiesService.getDocumentProperties().getProperties())
    .filter(k => k.indexOf('SHEET_ID_') === 0).forEach(k => PropertiesService.getDocumentProperties().deleteProperty(k));
}

test('setupOrUpgrade: recovers the half-upgraded "[NEW] ACN" state of 2026-10-12', () => {
  makeLegacyWorkbook(mockSpreadsheet);
  mockSpreadsheet.getSheetByName('tbl_members').getRange(2, 1, 1, 7).setValues([
    ['Hugo Gonçalves', '', '', 'Data Analyst', 'FAL ME', 'Active', '']
  ]);
  mockSpreadsheet.getSheetByName('tbl_presence_baselines').setName('archive_tbl_presence_baselines');
  mockSpreadsheet.getSheetByName('tbl_leaves').setName('archive_tbl_leaves');
  addMergedRangesTable(mockSpreadsheet);
  mockSpreadsheet.insertSheet('tbl_presence_baselines'); // empty tab left by the failed upgrade
  mockSpreadsheet.insertSheet('calc_kpis');

  var res = code.setupOrUpgrade(mockSpreadsheet);

  assert.strictEqual(res.success, true);
  assert.strictEqual(rows(mockSpreadsheet, 'Routines', 7).length, 2);
  assert.strictEqual(rows(mockSpreadsheet, 'Leaves', 8).length, 1);
  assert.strictEqual(mockSpreadsheet.getSheetByName('tbl_presence_ranges'), null);
  ['Archive – Routines', 'Archive – Leaves', 'Archive – Presence Ranges', 'Routines', 'Leaves']
    .forEach(n => assert.ok(mockSpreadsheet.getSheetByName(n).isSheetHidden(), n + ' hidden'));
});

test('setupOrUpgrade: gives the live workbook readable tab, table and header names, keeping data', () => {
  makeLegacyWorkbook(mockSpreadsheet);
  mockSpreadsheet.getSheetByName('tbl_members').getRange(2, 1, 1, 7).setValues([
    ['Hugo Gonçalves', 'hugo@accenture.com', '', 'Data Analyst', 'FAL ME', 'Active', '']
  ]);

  code.setupOrUpgrade(mockSpreadsheet);

  var names = mockSpreadsheet.getSheets().map(s => s.getName());
  ['Members', 'Contracts', 'Routines', 'Leaves', 'Presence Log', 'Projects', 'Tasks', 'Events', 'Backups']
    .forEach(n => assert.ok(names.includes(n), n + ' tab'));
  assert.ok(!names.some(n => n.indexOf('tbl_') === 0), 'no technical tab left');
  var tableNames = Sheets.Spreadsheets.tables.map(t => t.name);
  ['Members', 'Projects', 'Tasks', 'Events', 'Presence Log'].forEach(n => assert.ok(tableNames.includes(n), n + ' table'));
  var members = mockSpreadsheet.getSheetByName('Members');
  assert.deepStrictEqual(members.getRange(1, 1, 1, 7).getValues()[0],
    ['Member', 'Accenture email', 'Phone', 'Role', 'Department', 'Status', 'Today']);
  assert.strictEqual(members.getRange(2, 2).getValue(), 'hugo@accenture.com', 'data kept');
  assert.deepStrictEqual(mockSpreadsheet.getSheetByName('Contracts').getRange(1, 1, 1, 2).getValues()[0], ['member', 'role'], 'hidden tables keep technical headers');
});

test('freshInit: widens an existing table when the schema gains a column', () => {
  freshInit(mockSpreadsheet);
  var projects = mockSpreadsheet.getSheetByName('Projects');
  var table = Sheets.Spreadsheets.tables.find(t => t.name === 'Projects');
  table.range = Object.assign({}, table.range, { endColumnIndex: 6 }); // as created before active_responsible existed

  freshInit(mockSpreadsheet);

  assert.strictEqual(Sheets.Spreadsheets.tables.find(t => t.name === 'Projects').range.endColumnIndex, 7);
  assert.strictEqual(projects.getRange(1, 7).getValue(), 'In charge today');
});

test('setupOrUpgrade: removes the retired Active Responsibles view', () => {
  mockSpreadsheet.insertSheet('view_active_responsibles');

  code.setupOrUpgrade(mockSpreadsheet);

  assert.strictEqual(mockSpreadsheet.getSheetByName('view_active_responsibles'), null);
});
