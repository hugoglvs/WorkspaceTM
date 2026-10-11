const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');

let mockSpreadsheet;

test.beforeEach(() => {
  mockSpreadsheet = createMockSpreadsheet();
  setupGasGlobals(mockSpreadsheet);
});

const tables = require('../src/Tables');

test('getTable: finds a table by its display name, its former name, or its stored ID after a rename', () => {
  var legacy = mockSpreadsheet.insertSheet('Members');
  assert.strictEqual(tables.getTable(mockSpreadsheet, 'members'), legacy, 'former name still works before the upgrade');

  legacy.setName('Members');
  tables.registerTableIds(mockSpreadsheet);
  assert.strictEqual(tables.getTable(mockSpreadsheet, 'members'), legacy, 'display name');

  legacy.setName('Team');
  assert.strictEqual(tables.getTable(mockSpreadsheet, 'members'), legacy, 'a tab renamed by the team is still found');
  assert.strictEqual(tables.tableKeyOfSheet(legacy), 'members');
});

test('tableSheetName and tableKeyOfName map keys, display names and former names', () => {
  assert.strictEqual(tables.tableSheetName('tenures'), 'Contracts');
  assert.strictEqual(tables.tableSheetName('weeklyMatrix'), 'Weekly Presence');
  assert.strictEqual(tables.tableKeyOfName('tbl_presence_baselines'), 'baselines');
  assert.strictEqual(tables.tableKeyOfName('Presence Log'), 'presenceLog');
  assert.strictEqual(tables.tableKeyOfName('My own tab'), null);
});

test('renameLegacySheets: gives former tabs their readable names and archives readable names too', () => {
  ['tbl_members', 'tbl_tenures', 'view_weekly_matrix', 'fact_presence', 'archive_tbl_leaves'].forEach(n => mockSpreadsheet.insertSheet(n));

  var renamed = tables.renameLegacySheets(mockSpreadsheet);

  var names = mockSpreadsheet.getSheets().map(s => s.getName());
  ['Members', 'Contracts', 'Weekly Presence', 'KPI Data', 'Archive – Leaves'].forEach(n => assert.ok(names.includes(n), n));
  assert.strictEqual(renamed.length, 5);
  assert.strictEqual(tables.renameLegacySheets(mockSpreadsheet).length, 0, 'idempotent');
});
