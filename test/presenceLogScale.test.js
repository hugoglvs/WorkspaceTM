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

// A year of history for 10 members is ~2,500 rows: more than the old 2,000-row read cap.
function fillLog(rows) {
  var sheet = getTable(mockSpreadsheet, 'presenceLog');
  sheet.maxRows = rows.length + 200;
  sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  return sheet;
}

test('Presence Log beyond 2,000 rows: compile finds existing days and adds no duplicates', () => {
  freshInit(mockSpreadsheet);
  getTable(mockSpreadsheet, 'members').getRange(2, 1, 1, 7).setValues([['Alice', '', '', 'Dev', 'Eng', 'Active', '']]);
  var rows = [];
  for (var i = 0; i < 2500; i++) rows.push(['old' + i, '2020-01-01', 'Ghost' + i, 'OneRoof', 'OneRoof', '']);
  rows.push(['mine', '2026-10-05', 'Alice', 'C04', 'Remote', '']); // row 2502, a manual entry
  var sheet = fillLog(rows);

  compilePresenceLogs('2026-10-05', '2026-10-05', mockSpreadsheet, '2026-10-05');

  var all = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  var alice = all.filter(r => r[2] === 'Alice');
  assert.strictEqual(alice.length, 1, 'no duplicate row for Alice');
  assert.deepStrictEqual([alice[0][3], alice[0][4]], ['C04', 'Remote'], 'manual entry kept');
});

test('readRowLimit never stops before the last row with data', () => {
  freshInit(mockSpreadsheet);
  var sheet = fillLog(Array.from({ length: 3000 }, (_, i) => ['id' + i, '2020-01-01', 'M', 'OneRoof', 'OneRoof', '']));
  assert.ok(readRowLimit(sheet, 2000) >= 3001);
  assert.strictEqual(readRowLimit(getTable(mockSpreadsheet, 'leaves'), 500), Math.min(getTable(mockSpreadsheet, 'leaves').getMaxRows(), 500));
});
