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
const sidebar = require('../src/Sidebar');
global.extractCellValue = sidebar.extractCellValue;
global.extractEmailFromChip = sidebar.extractEmailFromChip;
global.isUserMemberMatch = sidebar.isUserMemberMatch;
global.checkActiveUserStatus = sidebar.checkActiveUserStatus;

const viewBuilder = require('../src/ViewBuilder');
global.renderWeeklyMatrixView = viewBuilder.renderWeeklyMatrixView;
global.refreshWeeklyMatrix = viewBuilder.refreshWeeklyMatrix;
global.renderActiveResponsiblesDashboard = viewBuilder.renderActiveResponsiblesDashboard;
global.refreshActiveResponsibles = viewBuilder.refreshActiveResponsibles;

const kpiEngine = require('../src/KpiEngine');
global.KpiEngine = kpiEngine;
global.renderKpiDashboard = kpiEngine.renderKpiDashboard;
global.refreshKpis = kpiEngine.refreshKpis;
global.computePresenceKpis = kpiEngine.computePresenceKpis;

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

test('computePresenceKpis: accounts for intra-day mobility with 0.5-day weighting and tenure FTE capacity', () => {
  freshInit(mockSpreadsheet);

  // Setup tbl_members
  // Alice: Core Engineering (FTE 1.0)
  // Bob: Design (FTE 0.8)
  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 2, 7).setValues([
    ['Alice Smith', '', '', 'Senior Eng', 'Core Engineering', 'Active', 'Main Office'],
    ['Bob Jones', '', '', 'Lead Designer', 'Design', 'Active', 'Remote']
  ]);

  // Setup tbl_tenures
  var tenuresSheet = mockSpreadsheet.getSheetByName('tbl_tenures');
  tenuresSheet.getRange(2, 1, 2, 7).setValues([
    ['Alice Smith', 'Senior Eng', 'Core Engineering', '2026-01-01', '', 'Permanent', 1.0],
    ['Bob Jones', 'Lead Designer', 'Design', '2026-01-01', '', 'Permanent', 0.8]
  ]);

  // Setup tbl_presence_log for January 2026:
  // Alice Day 1: AM Remote, PM Main Office (Split intra-day mobility: 0.5 Remote, 0.5 Main Office = 1.0 Worked day)
  // Alice Day 2: AM Main Office, PM Vacation (0.5 Worked day, 0.5 Leave day)
  // Bob Day 1: AM Remote, PM Remote (1.0 Remote = 1.0 Worked day)
  // Bob Day 2: AM Downtown Office, PM Branch Office (0.5 Downtown, 0.5 Branch = 1.0 In-Office day)
  var logSheet = mockSpreadsheet.getSheetByName('tbl_presence_log');
  logSheet.getRange(2, 1, 4, 6).setValues([
    ['log-1', '2026-01-05', 'Alice Smith', 'Remote', 'Main Office', '2026-01-05 09:00:00'],
    ['log-2', '2026-01-06', 'Alice Smith', 'Main Office', 'Vacation', '2026-01-06 09:00:00'],
    ['log-3', '2026-01-05', 'Bob Jones', 'Remote', 'Remote', '2026-01-05 09:00:00'],
    ['log-4', '2026-01-06', 'Bob Jones', 'Downtown Office', 'Branch Office', '2026-01-06 09:00:00']
  ]);

  var kpis = kpiEngine.computePresenceKpis(mockSpreadsheet);
  assert.strictEqual(kpis.success, true);
  assert.strictEqual(kpis.memberRows.length, 2);

  // Find Alice
  var alice = kpis.memberRows.find(r => r.member === 'Alice Smith');
  assert.ok(alice);
  assert.strictEqual(alice.month, '2026-01');
  assert.strictEqual(alice.department, 'Core Engineering');
  assert.strictEqual(alice.fteRate, 1.0);
  assert.strictEqual(alice.monthWeekdays, 22);
  assert.strictEqual(alice.contractCapacityDays, 22.0); // 22 * 1.0
  assert.strictEqual(alice.workedDays, 1.5); // Day 1 (1.0) + Day 2 (0.5)
  assert.strictEqual(alice.leaveDays, 0.5);  // Day 2 PM Vacation (0.5)
  assert.strictEqual(alice.remoteDays, 0.5); // Day 1 AM Remote (0.5)
  assert.strictEqual(alice.inOfficeDays, 1.0); // Day 1 PM (0.5) + Day 2 AM (0.5)
  assert.strictEqual(alice.mainOfficeDays, 1.0);
  assert.strictEqual(alice.remoteRate, 0.5 / 1.5); // 33.33%
  assert.strictEqual(alice.inOfficeRate, 1.0 / 1.5); // 66.67%
  assert.strictEqual(alice.capacityUtilization, 1.5 / 22.0);

  // Find Bob
  var bob = kpis.memberRows.find(r => r.member === 'Bob Jones');
  assert.ok(bob);
  assert.strictEqual(bob.department, 'Design');
  assert.strictEqual(bob.fteRate, 0.8);
  assert.strictEqual(bob.contractCapacityDays, 17.6); // 22 * 0.8
  assert.strictEqual(bob.workedDays, 2.0); // Day 1 (1.0) + Day 2 (1.0)
  assert.strictEqual(bob.leaveDays, 0.0);
  assert.strictEqual(bob.remoteDays, 1.0);
  assert.strictEqual(bob.inOfficeDays, 1.0);
  assert.strictEqual(bob.downtownOfficeDays, 0.5);
  assert.strictEqual(bob.branchOfficeDays, 0.5);
  assert.strictEqual(bob.remoteRate, 0.5); // 1.0 / 2.0 = 50%
  assert.strictEqual(bob.inOfficeRate, 0.5); // 1.0 / 2.0 = 50%
  assert.strictEqual(bob.capacityUtilization, 2.0 / 17.6);
});

test('renderKpiDashboard: builds calc_kpis sheet with formatted columns, rollups, and People Chips', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 2, 7).setValues([
    ['Alice Smith', '', '', 'Senior Eng', 'Core Engineering', 'Active', 'Main Office'],
    ['Bob Jones', '', '', 'Lead Designer', 'Design', 'Active', 'Remote']
  ]);

  var logSheet = mockSpreadsheet.getSheetByName('tbl_presence_log');
  logSheet.getRange(2, 1, 2, 6).setValues([
    ['log-1', '2026-01-05', 'Alice Smith', 'Remote', 'Main Office', '2026-01-05 09:00:00'],
    ['log-2', '2026-01-05', 'Bob Jones', 'Remote', 'Remote', '2026-01-05 09:00:00']
  ]);

  var res = kpiEngine.renderKpiDashboard(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.sheetName, 'calc_kpis');

  var kpiSheet = mockSpreadsheet.getSheetByName('calc_kpis');
  assert.ok(kpiSheet);
  assert.strictEqual(kpiSheet.getFrozenRows(), 1);

  // Verify Headers (16 columns)
  var headers = kpiSheet.getRange(1, 1, 1, 16).getValues()[0];
  assert.strictEqual(headers[0], 'Month');
  assert.strictEqual(headers[1], 'Department');
  assert.strictEqual(headers[2], 'Member');
  assert.strictEqual(headers[3], 'FTE Rate');
  assert.strictEqual(headers[4], 'Contract Capacity (Days)');
  assert.strictEqual(headers[5], 'Worked Days');
  assert.strictEqual(headers[6], 'Leave Days');
  assert.strictEqual(headers[7], 'Capacity Utilization');
  assert.strictEqual(headers[8], 'Remote Days');
  assert.strictEqual(headers[9], 'Remote Rate');
  assert.strictEqual(headers[10], 'In-Office Days');
  assert.strictEqual(headers[11], 'In-Office Rate');
  assert.strictEqual(headers[12], 'Main Office (Days)');
  assert.strictEqual(headers[13], 'Downtown Office (Days)');
  assert.strictEqual(headers[14], 'Branch Office (Days)');
  assert.strictEqual(headers[15], 'Client Site (Days)');

  // Verify column widths
  assert.strictEqual(kpiSheet.columnWidths[1], 110);
  assert.strictEqual(kpiSheet.columnWidths[2], 160);
  assert.strictEqual(kpiSheet.columnWidths[3], 190);

  // Check that row values contain valid data
  var rowValues = kpiSheet.getRange(2, 1, res.totalTableRows, 16).getValues();
  assert.ok(rowValues.length >= 2);
});

test('refreshKpis: toasts update and renders dashboard successfully', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Charlie Brown', '', '', 'Dev', 'Delivery', 'Active', 'Main Office']
  ]);

  var res = code.refreshKpis(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.ok(mockSpreadsheet.lastToast);
  assert.ok(mockSpreadsheet.lastToast.title.includes('KPIs Updated'));
});

test('onOpen: registers Refresh Presence & Remote KPIs menu item', () => {
  code.onOpen();
  var ui = SpreadsheetApp.getUi();
  assert.ok(ui.menu);
  var itemCaptions = ui.menu.items.map(function(item) { return item.caption; });
  assert.ok(itemCaptions.includes('Refresh Presence & Remote KPIs (calc_kpis)'));
});

function seedTeam(spreadsheet, tenureRows, logRows) {
  freshInit(spreadsheet);
  spreadsheet.getSheetByName('tbl_members').getRange(2, 1, tenureRows.length, 7).setValues(
    tenureRows.map(t => [t[0], '', '', t[1], t[2], 'Active', ''])
  );
  spreadsheet.getSheetByName('tbl_tenures').getRange(2, 1, tenureRows.length, 7).setValues(tenureRows);
  if (logRows.length) {
    spreadsheet.getSheetByName('tbl_presence_log').getRange(2, 1, logRows.length, 6).setValues(logRows);
  }
}

test('computePresenceKpis: reads fte_rate entered as a percentage', () => {
  seedTeam(mockSpreadsheet, [
    ['Alice Smith', 'Senior Eng', 'Core Engineering', '2026-01-01', '', 'Permanent', 100],
    ['Bob Jones', 'Lead Designer', 'Design', '2026-01-01', '', 'Contractor', 80]
  ], [
    ['log-1', '2026-01-05', 'Alice Smith', 'Remote', 'Remote', '2026-01-05 09:00:00'],
    ['log-2', '2026-01-05', 'Bob Jones', 'Remote', 'Remote', '2026-01-05 09:00:00']
  ]);

  var kpis = kpiEngine.computePresenceKpis(mockSpreadsheet);

  var alice = kpis.memberRows.find(r => r.member === 'Alice Smith');
  var bob = kpis.memberRows.find(r => r.member === 'Bob Jones');
  assert.strictEqual(alice.fteRate, 1.0);
  assert.strictEqual(alice.contractCapacityDays, 22);
  assert.strictEqual(bob.fteRate, 0.8);
  assert.strictEqual(bob.contractCapacityDays, 17.6);
});

test('computePresenceKpis: weekend Off is not counted as leave', () => {
  seedTeam(mockSpreadsheet, [
    ['Alice Smith', 'Senior Eng', 'Core Engineering', '2026-01-01', '', 'Permanent', 1.0]
  ], [
    // Saturday 2026-01-10: client site in the morning, off in the afternoon
    ['log-1', '2026-01-10', 'Alice Smith', 'Client Site', 'Off', '2026-01-10 09:00:00']
  ]);

  var alice = kpiEngine.computePresenceKpis(mockSpreadsheet).memberRows.find(r => r.member === 'Alice Smith');

  assert.strictEqual(alice.workedDays, 0.5);
  assert.strictEqual(alice.leaveDays, 0);
});

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

  kpiEngine.refreshKpis(mockSpreadsheet);

  var sheet = mockSpreadsheet.getSheetByName('fact_presence');
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
