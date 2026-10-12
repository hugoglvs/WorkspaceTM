const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');
Object.assign(global, require('../src/Tables'));

let mockSpreadsheet;

test.beforeEach(() => {
  mockSpreadsheet = createMockSpreadsheet();
  setupGasGlobals(mockSpreadsheet);
});

const schema = require('../src/Schema');
Object.assign(global, schema);
const typography = require('../src/Typography');
Object.assign(global, typography);
Object.assign(global, require('../src/SyncEngine'));
const { freshInit } = require('../src/Init');
global.freshInit = freshInit;
const sidebar = require('../src/Sidebar');
Object.assign(global, sidebar);
Object.assign(global, require('../src/PresenceEngine'));
Object.assign(global, require('../src/ViewBuilder'));
const kpi = require('../src/KpiEngine');
Object.assign(global, kpi);
const locations = require('../src/Locations');
Object.assign(global, locations);
Object.assign(global, require('../src/Layout'));
const code = require('../src/Code');

function locationRows() {
  return getTable(mockSpreadsheet, 'locations').getRange(2, 1, 10, 3).getValues().filter(r => r[0]);
}

function seedMembers(rows) {
  getTable(mockSpreadsheet, 'members').getRange(2, 1, rows.length, 7).setValues(rows);
}

test('freshInit: creates the Locations table with the default list', () => {
  freshInit(mockSpreadsheet);

  var sheet = getTable(mockSpreadsheet, 'locations');
  assert.strictEqual(sheet.getName(), 'Locations');
  assert.deepStrictEqual(sheet.getRange(1, 1, 1, 3).getValues()[0], ['Location', 'Type', 'Active']);
  assert.deepStrictEqual(locationRows(), [
    ['OneRoof', 'Office', true],
    ['Remote', 'Remote', true],
    ['JLL', 'Client site', true],
    ['C04', 'Client site', true],
    ['C60', 'Client site', true]
  ]);

  freshInit(mockSpreadsheet);
  assert.strictEqual(locationRows().length, 5, 'defaults are written once');
});

test('readLocations: active list for new entries, full list for history, type per location', () => {
  freshInit(mockSpreadsheet);
  getTable(mockSpreadsheet, 'locations').getRange(5, 3).setValue(false); // C04 inactive

  assert.deepStrictEqual(locations.activeLocationNames(mockSpreadsheet), ['OneRoof', 'Remote', 'JLL', 'C60']);
  assert.deepStrictEqual(locations.allLocationNames(mockSpreadsheet), ['OneRoof', 'Remote', 'JLL', 'C04', 'C60']);
  assert.strictEqual(locations.locationCategory(mockSpreadsheet, 'JLL'), 'client_site');
  assert.strictEqual(locations.locationCategory(mockSpreadsheet, 'OneRoof'), 'office');
  assert.strictEqual(locations.defaultOfficeLocation(mockSpreadsheet), 'OneRoof');
});

test('Presence & Leaves sidebar: offers active locations plus Off, and the default office', () => {
  freshInit(mockSpreadsheet);
  getTable(mockSpreadsheet, 'locations').getRange(5, 3).setValue(false);

  var data = sidebar.getSidebarInitialData();

  assert.deepStrictEqual(data.locationOptions, ['OneRoof', 'Remote', 'JLL', 'C60', 'Off']);
  assert.strictEqual(data.defaultOffice, 'OneRoof');
  assert.strictEqual(data.defaultRemote, 'Remote');
});

test('hidden data tables accept every location, active or not, so history stays valid', () => {
  freshInit(mockSpreadsheet);
  getTable(mockSpreadsheet, 'locations').getRange(5, 3).setValue(false);

  locations.refreshLocationDropdowns(mockSpreadsheet);

  var routinesSheetId = getTable(mockSpreadsheet, 'baselines').getSheetId();
  var routines = Sheets.Spreadsheets.tables.find(t => t.sheetId === routinesSheetId);
  var amColumn = routines.columnProperties.find(c => c.columnName === 'location_am');
  var values = amColumn.dataValidationRule.condition.values.map(v => v.userEnteredValue);
  assert.ok(values.includes('C04'), 'inactive location still valid');
  assert.ok(values.includes('Off'));
});

test('Weekly Presence: dropdown of active locations from the current week, none on past weeks', () => {
  freshInit(mockSpreadsheet);
  seedMembers([['Alice', '', '', 'Dev', 'Eng', 'Active', '']]);
  getTable(mockSpreadsheet, 'locations').getRange(5, 3).setValue(false);

  renderWeeklyMatrixView('2026-09-28', '2026-10-18', mockSpreadsheet, '2026-10-12');

  var sheet = getTable(mockSpreadsheet, 'weeklyMatrix');
  var withRule = Object.values(sheet.ranges).filter(r => r.dataValidation && r.startRow >= 3);
  assert.strictEqual(withRule.length, 1);
  assert.strictEqual(withRule[0].startRow, 3 + 14, 'starts on Monday 12 October (row 17)');
  assert.deepStrictEqual(withRule[0].dataValidation.allowedValues.slice(0, 5), ['OneRoof', 'Remote', 'JLL', 'C60', 'Off']);
  assert.ok(!withRule[0].dataValidation.allowedValues.includes('C04'));
  assert.ok(Object.values(sheet.ranges).some(r => r.validationCleared && r.startRow === 3), 'past weeks cleared');
});

test('Weekly Presence: departed members are hidden, upcoming and active members shown', () => {
  freshInit(mockSpreadsheet);
  seedMembers([
    ['Alice', '', '', 'Dev', 'Eng', 'Active', ''],
    ['Bob', '', '', 'Dev', 'Eng', 'Departed', ''],
    ['Carol', '', '', 'Dev', 'Eng', 'Upcoming', '']
  ]);

  renderWeeklyMatrixView('2026-10-05', '2026-10-11', mockSpreadsheet);

  var sheet = getTable(mockSpreadsheet, 'weeklyMatrix');
  assert.strictEqual(sheet.getRange(1, 5).getValue(), 'Bob');
  assert.deepStrictEqual([3, 4, 5, 6, 7, 8].map(c => sheet.isColumnHiddenByUser(c)), [false, false, true, true, false, false]);
});

test('KPI Data: client-site days are worked days of their own category', () => {
  freshInit(mockSpreadsheet);
  seedMembers([['Alice', '', '', 'Dev', 'Eng', 'Active', '']]);
  getTable(mockSpreadsheet, 'tenures').getRange(2, 1, 1, 7).setValues([['Alice', 'Dev', 'Eng', '2026-01-01', '', 'Permanent', 100]]);
  getTable(mockSpreadsheet, 'presenceLog').getRange(2, 1, 1, 6).setValues([['l1', '2026-10-05', 'Alice', 'JLL', 'OneRoof', '']]);

  var facts = kpi.buildPresenceFacts(mockSpreadsheet, { asOfDate: '2026-10-05' });
  var cat = facts.headers.indexOf('category'), worked = facts.headers.indexOf('worked_days');

  assert.deepStrictEqual(facts.rows.map(r => [r[cat], r[worked]]), [['client_site', 0.5], ['office', 0.5]]);
});

test('compile: a weekday without any routine defaults to the first active office', () => {
  freshInit(mockSpreadsheet);
  seedMembers([['Alice', '', '', 'Dev', 'Eng', 'Active', '']]);

  compilePresenceLogs('2026-10-05', '2026-10-05', mockSpreadsheet, '2026-10-05');

  var log = getTable(mockSpreadsheet, 'presenceLog').getRange(2, 1, 1, 6).getValues()[0];
  assert.deepStrictEqual([log[3], log[4]], ['OneRoof', 'OneRoof']);
});

test('editing Locations refreshes the dropdowns, and the menu can show the hidden tab', () => {
  freshInit(mockSpreadsheet);
  seedMembers([['Alice', '', '', 'Dev', 'Eng', 'Active', '']]);
  var sheet = getTable(mockSpreadsheet, 'locations');
  sheet.getRange(7, 1, 1, 3).setValues([['Station F', 'Client site', true]]);

  code.onEdit({ range: sheet.getRange(7, 1), value: 'Station F', source: mockSpreadsheet });

  var sidebarData = sidebar.getSidebarInitialData();
  assert.ok(sidebarData.locationOptions.includes('Station F'));

  sheet.hideSheet();
  locations.showLocationsTab();
  assert.strictEqual(sheet.isSheetHidden(), false);
  assert.strictEqual(sheet.activated, true);
});

test('Members: Status is a dropdown of the four statuses the sync writes', () => {
  freshInit(mockSpreadsheet);

  var members = Sheets.Spreadsheets.tables.find(t => t.sheetId === getTable(mockSpreadsheet, 'members').getSheetId());
  var status = members.columnProperties.find(c => c.columnName === 'Status');
  assert.strictEqual(status.columnType, 'DROPDOWN');
  assert.deepStrictEqual(status.dataValidationRule.condition.values.map(v => v.userEnteredValue), ['Active', 'Upcoming', 'Departed', 'No Contract']);
});

function retire(name) {
  var sheet = getTable(mockSpreadsheet, 'locations');
  var row = sheet.getRange(2, 1, 10, 1).getValues().findIndex(r => r[0] === name) + 2;
  sheet.getRange(row, 3).setValue(false);
  code.onEdit({ range: sheet.getRange(row, 3), value: 'FALSE', source: mockSpreadsheet });
}

test('retiring a location warns about routines and future days that still use it', () => {
  freshInit(mockSpreadsheet);
  seedMembers([['Alice', '', '', 'Dev', 'Eng', 'Active', ''], ['Bob', '', '', 'Dev', 'Eng', 'Active', '']]);
  getTable(mockSpreadsheet, 'baselines').getRange(2, 1, 3, 7).setValues([
    ['b1', 'Alice', '2026-01-01', '', 'Monday', 'C04', 'C04'],           // ongoing: counts
    ['b2', 'Bob', '2026-01-01', '2026-02-01', 'Monday', 'C04', 'C04'],   // ended: ignored
    ['b3', 'Bob', '2026-01-01', '', 'Tuesday', 'OneRoof', 'OneRoof']
  ]);
  getTable(mockSpreadsheet, 'presenceLog').getRange(2, 1, 2, 6).setValues([
    ['l1', '2099-01-05', 'Bob', 'C04', 'OneRoof', ''],    // future: counts
    ['l2', '2026-01-05', 'Alice', 'C04', 'C04', '']       // past: history, ignored
  ]);

  retire('C04');

  var toast = mockSpreadsheet.lastToast;
  assert.ok(toast, 'a toast is shown');
  assert.ok(toast.msg.includes('C04'));
  assert.ok(toast.msg.includes('1 routine') && toast.msg.includes('Alice'), toast.msg);
  assert.ok(toast.msg.includes('1 future day'), toast.msg);
});

test('retiring an unused location shows no warning', () => {
  freshInit(mockSpreadsheet);
  seedMembers([['Alice', '', '', 'Dev', 'Eng', 'Active', '']]);
  mockSpreadsheet.lastToast = null;

  retire('C60');

  assert.ok(!mockSpreadsheet.lastToast || !mockSpreadsheet.lastToast.msg.includes('C60'));
});
