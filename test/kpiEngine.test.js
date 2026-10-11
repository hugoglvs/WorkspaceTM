const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');
Object.assign(global, require('../src/Tables'));
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
const sidebar = require('../src/Sidebar');
global.extractCellValue = sidebar.extractCellValue;
global.extractEmailFromChip = sidebar.extractEmailFromChip;
global.isUserMemberMatch = sidebar.isUserMemberMatch;
global.checkActiveUserStatus = sidebar.checkActiveUserStatus;

const viewBuilder = require('../src/ViewBuilder');
global.renderWeeklyMatrixView = viewBuilder.renderWeeklyMatrixView;
global.refreshWeeklyMatrix = viewBuilder.refreshWeeklyMatrix;

const kpiEngine = require('../src/KpiEngine');
global.KpiEngine = kpiEngine;
global.renderPresenceFacts = kpiEngine.renderPresenceFacts;

const code = require('../src/Code');

test('countMonthWeekdays: correctly calculates working weekdays for month', () => {
  // January 2026: Jan 1 is Thursday, Jan 31 is Saturday -> 22 weekdays
  assert.strictEqual(kpiEngine.countMonthWeekdays('2026-01'), 22);
  // February 2026: 28 days -> exactly 20 weekdays
  assert.strictEqual(kpiEngine.countMonthWeekdays('2026-02'), 20);
});

test('categorizeLocation & isLeaveLocation: accurately identifies locations', () => {
  assert.strictEqual(kpiEngine.isLeaveLocation('Vacation'), true);
  assert.strictEqual(kpiEngine.isLeaveLocation('Sick Leave'), true);
  assert.strictEqual(kpiEngine.isLeaveLocation('Holiday'), true);
  assert.strictEqual(kpiEngine.isLeaveLocation('Remote'), false);
  assert.strictEqual(kpiEngine.isLeaveLocation('Main Office'), false);

  assert.strictEqual(kpiEngine.categorizeLocation('Remote'), 'remote');
  assert.strictEqual(kpiEngine.categorizeLocation('Main Office'), 'main_office');
  assert.strictEqual(kpiEngine.categorizeLocation('Downtown Office'), 'downtown_office');
  assert.strictEqual(kpiEngine.categorizeLocation('Branch Office'), 'branch_office');
  assert.strictEqual(kpiEngine.categorizeLocation('Client Site'), 'client_site');
  assert.strictEqual(kpiEngine.categorizeLocation('Vacation'), 'none');
});

test('onOpen: registers the KPI data refresh menu item', () => {
  code.onOpen();
  var ui = SpreadsheetApp.getUi();
  assert.ok(ui.menu);
  var itemCaptions = ui.menu.items.map(function(item) { return item.caption; });
  assert.ok(itemCaptions.includes('Refresh KPI Data for Looker'));
});

function seedTeam(spreadsheet, tenureRows, logRows) {
  freshInit(spreadsheet);
  spreadsheet.getSheetByName('Members').getRange(2, 1, tenureRows.length, 7).setValues(
    tenureRows.map(t => [t[0], '', '', t[1], t[2], 'Active', ''])
  );
  spreadsheet.getSheetByName('Contracts').getRange(2, 1, tenureRows.length, 7).setValues(tenureRows);
  if (logRows.length) {
    spreadsheet.getSheetByName('Presence Log').getRange(2, 1, logRows.length, 6).setValues(logRows);
  }
}

function factsByKey(result) {
  var map = {};
  result.rows.forEach(function(row) {
    var obj = {};
    result.headers.forEach(function(h, i) { obj[h] = row[i]; });
    map[obj.member + '|' + obj.date + '|' + obj.period] = obj;
  });
  return map;
}

test('buildPresenceFacts: one row per Member, date and Period with worked, leave and capacity days', () => {
  seedTeam(mockSpreadsheet, [
    ['Bob Jones', 'Lead Designer', 'Design', '2026-01-01', '', 'Contractor', 80]
  ], [
    ['log-1', '2026-01-05', 'Bob Jones', 'Remote', 'Main Office', '2026-01-05 09:00:00'],   // Monday
    ['log-2', '2026-01-07', 'Bob Jones', 'Main Office', 'Vacation', '2026-01-07 09:00:00'], // Wednesday
    ['log-3', '2026-01-10', 'Bob Jones', 'Client Site', 'Off', '2026-01-10 09:00:00']       // Saturday
  ]);

  var facts = kpiEngine.buildPresenceFacts(mockSpreadsheet, { asOfDate: '2026-01-11' });
  var f = factsByKey(facts);

  assert.deepStrictEqual(facts.headers, ['date', 'month', 'iso_week', 'weekday', 'is_weekend', 'member', 'department', 'role', 'fte_rate', 'period', 'location', 'category', 'worked_days', 'leave_days', 'capacity_days']);
  // Mon 5 → Fri 9 = 5 weekdays × 2 Periods, plus logged Saturday 10; unlogged Sunday 11 has no rows
  assert.strictEqual(facts.rows.length, 12);

  assert.deepStrictEqual(f['Bob Jones|2026-01-05|AM'], {
    date: '2026-01-05', month: '2026-01', iso_week: '2026-W02', weekday: 'Mon', is_weekend: false,
    member: 'Bob Jones', department: 'Design', role: 'Lead Designer', fte_rate: 0.8,
    period: 'AM', location: 'Remote', category: 'remote',
    worked_days: 0.5, leave_days: 0, capacity_days: 0.4
  });
  assert.strictEqual(f['Bob Jones|2026-01-05|PM'].category, 'office');
  assert.strictEqual(f['Bob Jones|2026-01-06|AM'].category, 'not_logged');
  assert.strictEqual(f['Bob Jones|2026-01-06|AM'].capacity_days, 0.4);
  assert.strictEqual(f['Bob Jones|2026-01-07|PM'].category, 'leave');
  assert.strictEqual(f['Bob Jones|2026-01-07|PM'].leave_days, 0.5);
  assert.strictEqual(f['Bob Jones|2026-01-10|AM'].worked_days, 0.5);
  assert.strictEqual(f['Bob Jones|2026-01-10|AM'].capacity_days, 0);
  assert.strictEqual(f['Bob Jones|2026-01-10|PM'].category, 'off');
  assert.strictEqual(f['Bob Jones|2026-01-10|PM'].worked_days + f['Bob Jones|2026-01-10|PM'].leave_days + f['Bob Jones|2026-01-10|PM'].capacity_days, 0);
});

test('buildPresenceFacts: capacity only within each Member tenure', () => {
  seedTeam(mockSpreadsheet, [
    ['Alice Smith', 'Senior Eng', 'Core Engineering', '2026-01-07', '', 'Permanent', 1.0],
    ['Bob Jones', 'Lead Designer', 'Design', '2026-01-01', '', 'Contractor', 80]
  ], [
    ['log-1', '2026-01-05', 'Bob Jones', 'Remote', 'Remote', '2026-01-05 09:00:00']
  ]);

  var facts = kpiEngine.buildPresenceFacts(mockSpreadsheet, { asOfDate: '2026-01-08' });
  var aliceDates = facts.rows.filter(r => r[5] === 'Alice Smith').map(r => r[0]);

  assert.deepStrictEqual(Array.from(new Set(aliceDates)), ['2026-01-07', '2026-01-08']);
});

test('refreshKpis: writes the hidden fact_presence sheet with no total rows', () => {
  seedTeam(mockSpreadsheet, [
    ['Bob Jones', 'Lead Designer', 'Design', '2026-01-01', '', 'Contractor', 80]
  ], [
    ['log-1', '2026-01-05', 'Bob Jones', 'Remote', 'Remote', '2026-01-05 09:00:00']
  ]);

  code.refreshKpis(mockSpreadsheet);

  var sheet = mockSpreadsheet.getSheetByName('KPI Data');
  assert.ok(sheet, 'fact_presence exists');
  assert.strictEqual(sheet.isSheetHidden(), true);
  var values = sheet.getRange(1, 1, sheet.getLastRow(), 15).getValues();
  assert.strictEqual(values[0][0], 'date');
  assert.ok(values.length > 1);
  assert.ok(values.slice(1).every(r => r[5] === 'Bob Jones'), 'only Member rows, no totals');
});

test('buildPresenceFacts: a weekday Off keeps capacity because the FTE already covers it', () => {
  seedTeam(mockSpreadsheet, [
    ['Bob Jones', 'Lead Designer', 'Design', '2026-01-01', '', 'Contractor', 80]
  ], [
    ['log-1', '2026-01-09', 'Bob Jones', 'Off', 'Off', '2026-01-09 09:00:00'] // Friday
  ]);

  var f = factsByKey(kpiEngine.buildPresenceFacts(mockSpreadsheet, { asOfDate: '2026-01-09' }));

  assert.strictEqual(f['Bob Jones|2026-01-09|AM'].category, 'off');
  assert.strictEqual(f['Bob Jones|2026-01-09|AM'].capacity_days, 0.4);
});
