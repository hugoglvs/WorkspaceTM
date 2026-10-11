const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');

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
  assert.deepStrictEqual(rows(mockSpreadsheet, 'tbl_presence_baselines', 7), [
    ['base_tm8krbd', 'Hugo Gonçalves', '2026-09-02', '2026-10-09', 'Monday', 'Remote', 'Remote'],
    ['base_egk5kvt', 'Hugo Gonçalves', '2026-10-09', '2026-10-16', 'Monday', 'Remote', 'Remote']
  ]);
  assert.deepStrictEqual(rows(mockSpreadsheet, 'tbl_leaves', 8), [
    ['leave_q45eloj', 'Hugo Gonçalves', '2026-10-01', 'AM', '2026-10-05', 'PM', 'Vacation', 'Requested']
  ]);
  assert.strictEqual(mockSpreadsheet.getSheetByName('tbl_presence_ranges'), null);
  assert.ok(mockSpreadsheet.getSheetByName('archive_tbl_presence_ranges').isSheetHidden());

  assert.strictEqual(restorePresenceTablesFromRanges(mockSpreadsheet).restoredCount, 0, 'idempotent');
});

test('restorePresenceTablesFromRanges: never duplicates a row whose ID already exists', () => {
  freshInit(mockSpreadsheet);
  mockSpreadsheet.getSheetByName('tbl_presence_baselines').getRange(2, 1, 1, 7).setValues([
    ['base_tm8krbd', 'Hugo Gonçalves', '2026-09-02', '', 'Monday', 'Remote', 'Remote']
  ]);
  addMergedRangesTable(mockSpreadsheet);

  restorePresenceTablesFromRanges(mockSpreadsheet);

  var ids = rows(mockSpreadsheet, 'tbl_presence_baselines', 7).map(r => r[0]);
  assert.deepStrictEqual(ids, ['base_tm8krbd', 'base_egk5kvt']);
});

test('setupOrUpgrade: shows only Members, Weekly Matrix, Projects, Tasks and Events', () => {
  mockSpreadsheet.insertSheet('calc_kpis');
  freshInit(mockSpreadsheet);
  mockSpreadsheet.getSheetByName('tbl_members').getRange(2, 1, 1, 7).setValues([
    ['Hugo Gonçalves', '', '', 'Data Analyst', 'FAL ME', 'Active', '']
  ]);
  addMergedRangesTable(mockSpreadsheet);
  mockSpreadsheet.insertSheet('Sheet1');

  code.setupOrUpgrade(mockSpreadsheet);

  var visible = mockSpreadsheet.getSheets().filter(s => !s.isSheetHidden()).map(s => s.getName()).sort();
  assert.deepStrictEqual(visible, ['tbl_events', 'tbl_members', 'tbl_projects', 'tbl_tasks', 'view_weekly_matrix']);
  assert.strictEqual(mockSpreadsheet.getSheetByName('calc_kpis'), null, 'calc_kpis deleted');
  assert.strictEqual(mockSpreadsheet.getSheetByName('Sheet1'), null, 'empty Sheet1 deleted');
  assert.strictEqual(rows(mockSpreadsheet, 'tbl_presence_baselines', 7).length, 2, 'merged data restored');
});

test('setupOrUpgrade: keeps a Sheet1 that has content', () => {
  mockSpreadsheet.getSheetByName('Sheet1').getRange(1, 1).setValue('my notes');

  code.setupOrUpgrade(mockSpreadsheet);

  assert.ok(mockSpreadsheet.getSheetByName('Sheet1'));
});
