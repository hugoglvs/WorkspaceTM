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
global.formatDateIso = sidebar.formatDateIso;

const presenceRanges = require('../src/PresenceRanges');
Object.assign(global, presenceRanges);
const presenceEngine = require('../src/PresenceEngine');
Object.assign(global, presenceEngine);
const code = require('../src/Code');

function rangesTable(spreadsheet) {
  var sheet = spreadsheet.getSheetByName('tbl_presence_ranges');
  var header = sheet.getRange(1, 1, 1, 11).getValues()[0];
  return sheet.getRange(2, 1, 20, 11).getValues()
    .filter(r => r[0] || r[1])
    .map(r => { var o = {}; header.forEach((h, i) => { o[h] = r[i]; }); return o; });
}

function addLegacySheet(spreadsheet, name, header, rows) {
  var sheet = spreadsheet.insertSheet(name);
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
  return sheet;
}

test('freshInit: creates one tbl_presence_ranges table instead of baselines and leaves', () => {
  freshInit(mockSpreadsheet);

  assert.ok(mockSpreadsheet.getSheetByName('tbl_presence_ranges'));
  assert.strictEqual(mockSpreadsheet.getSheetByName('tbl_presence_baselines'), null);
  assert.strictEqual(mockSpreadsheet.getSheetByName('tbl_leaves'), null);
  assert.deepStrictEqual(
    mockSpreadsheet.getSheetByName('tbl_presence_ranges').getRange(1, 1, 1, 11).getValues()[0],
    ['range_id', 'member', 'range_type', 'start_date', 'start_period', 'end_date', 'end_period', 'day_of_week', 'location_am', 'location_pm', 'status']
  );
});

test('migrateLegacyPresenceTables: moves baselines and leaves into presence ranges and archives the old tables', () => {
  freshInit(mockSpreadsheet);
  addLegacySheet(mockSpreadsheet, 'tbl_presence_baselines',
    ['baseline_id', 'member', 'start_date', 'end_date', 'day_of_week', 'location_am', 'location_pm'],
    [['base_1', 'Hugo Gonçalves', '2026-09-02', '', 'Monday', 'Remote', 'Remote']]);
  addLegacySheet(mockSpreadsheet, 'tbl_leaves',
    ['leave_id', 'member', 'start_date', 'start_period', 'end_date', 'end_period', 'leave_type', 'status'],
    [['', 'Hugo Gonçalves', '2026-10-01', '', '2026-10-05', 'AM', 'Vacation', 'Requested']]);

  var res = migrateLegacyPresenceTables(mockSpreadsheet);

  assert.strictEqual(res.migratedCount, 2);
  var rows = rangesTable(mockSpreadsheet);
  assert.deepStrictEqual(rows[0], {
    range_id: 'base_1', member: 'Hugo Gonçalves', range_type: 'Baseline',
    start_date: '2026-09-02', start_period: '', end_date: '', end_period: '',
    day_of_week: 'Monday', location_am: 'Remote', location_pm: 'Remote', status: 'Approved'
  });
  assert.ok(String(rows[1].range_id).startsWith('leave_'), 'a leave typed without an ID gets one');
  assert.deepStrictEqual(Object.assign({}, rows[1], { range_id: 'x' }), {
    range_id: 'x', member: 'Hugo Gonçalves', range_type: 'Leave',
    start_date: '2026-10-01', start_period: 'AM', end_date: '2026-10-05', end_period: 'AM',
    day_of_week: 'All', location_am: 'Vacation', location_pm: 'Vacation', status: 'Requested'
  });

  var archived = mockSpreadsheet.getSheetByName('archive_tbl_leaves');
  assert.ok(archived && archived.isSheetHidden(), 'old leaves table archived and hidden, not deleted');
  assert.ok(mockSpreadsheet.getSheetByName('archive_tbl_presence_baselines').isSheetHidden());
  assert.strictEqual(mockSpreadsheet.getSheetByName('tbl_leaves'), null);

  assert.strictEqual(migrateLegacyPresenceTables(mockSpreadsheet).migratedCount, 0, 'idempotent');
  assert.strictEqual(rangesTable(mockSpreadsheet).length, 2);
});

test('migrateLegacyPresenceTables: reads the legacy 6-column leaves layout', () => {
  freshInit(mockSpreadsheet);
  addLegacySheet(mockSpreadsheet, 'tbl_leaves',
    ['leave_id', 'member', 'start_date', 'end_date', 'leave_type', 'status'],
    [['LV-001', 'Bob Jones', '2026-01-05', '2026-01-09', 'Sick Leave', 'Approved']]);

  migrateLegacyPresenceTables(mockSpreadsheet);

  var row = rangesTable(mockSpreadsheet)[0];
  assert.strictEqual(row.range_id, 'LV-001');
  assert.strictEqual(row.start_period, 'AM');
  assert.strictEqual(row.end_period, 'PM');
  assert.strictEqual(row.location_am, 'Sick Leave');
  assert.strictEqual(row.status, 'Approved');
});

test('setupOrUpgrade: migrates legacy tables and hides presence data so only the Weekly Matrix shows', () => {
  addLegacySheet(mockSpreadsheet, 'tbl_leaves',
    ['leave_id', 'member', 'start_date', 'start_period', 'end_date', 'end_period', 'leave_type', 'status'],
    [['LV-1', 'Hugo Gonçalves', '2026-10-01', 'AM', '2026-10-05', 'PM', 'Vacation', 'Approved']]);

  code.setupOrUpgrade(mockSpreadsheet);

  assert.strictEqual(rangesTable(mockSpreadsheet).length, 1);
  ['tbl_presence_ranges', 'tbl_presence_log'].forEach(function(name) {
    assert.strictEqual(mockSpreadsheet.getSheetByName(name).isSheetHidden(), true, name + ' hidden');
  });
  assert.strictEqual(mockSpreadsheet.getSheetByName('tbl_tasks').isSheetHidden(), false, 'other tables untouched');
});
