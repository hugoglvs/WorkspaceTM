const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');
Object.assign(global, require('../src/Tables'));

let ss;

test.beforeEach(() => {
  ss = createMockSpreadsheet();
  setupGasGlobals(ss);
});

const schema = require('../src/Schema');
Object.assign(global, schema);
Object.assign(global, require('../src/Typography'));
Object.assign(global, require('../src/SyncEngine'));
const { freshInit } = require('../src/Init');
Object.assign(global, require('../src/Sidebar'));
Object.assign(global, require('../src/PresenceEngine'));
Object.assign(global, require('../src/ViewBuilder'));
Object.assign(global, require('../src/Locations'));
Object.assign(global, require('../src/Dropdowns'));

function tab() { return getTable(ss, 'dropdowns'); }

test('Dropdowns tab: one model per list, a dropdown from its options column, kept once customised', () => {
  freshInit(ss);
  ensureDropdownsTab(ss);
  var sheet = tab();
  assert.strictEqual(sheet.getName(), 'Dropdowns');
  assert.ok(sheet.isSheetHidden(), 'hidden');
  assert.deepStrictEqual(sheet.getRange(2, 1, DROPDOWN_LISTS.length, 1).getValues().map(r => r[0]), ['Presence']);
  // Presence options: active locations, then Off and leaves
  assert.deepStrictEqual(sheet.getRange(2, 5, 4, 1).getValues().map(r => r[0]), ['OneRoof', 'Remote', 'JLL', 'C04']);
  var rule = sheet.getRange(2, 2).getDataValidation();
  assert.ok(rule && rule.sourceRange, 'model dropdown reads its options column');

  // Someone styles the model (chip, colours): an upgrade does not replace it, options still refresh
  var styled = { styled: true };
  sheet.getRange(2, 2).setDataValidation(styled);
  getTable(ss, 'locations').getRange(7, 1, 1, 3).setValues([['Lyon', 'Office', true]]);
  ensureDropdownsTab(ss);
  assert.strictEqual(sheet.getRange(2, 2).getDataValidation(), styled);
  assert.ok(sheet.getRange(2, 5, 10, 1).getValues().map(r => r[0]).includes('Lyon'));
});

test('Dropdowns: the Presence model is copied onto Weekly Presence', () => {
  freshInit(ss);
  ensureDropdownsTab(ss);
  var presence = { styled: 'presence' };
  tab().getRange(2, 2).setDataValidation(presence);
  getTable(ss, 'members').getRange(2, 1, 1, 7).setValues([['Alice', '', '', 'Dev', 'Eng', 'Active', '']]);
  renderWeeklyMatrixView('2026-10-12', '2026-10-18', ss, '2026-10-12');
  assert.strictEqual(getTable(ss, 'weeklyMatrix').getRange(3, 3).getDataValidation(), presence);
  assert.deepStrictEqual(applyTableDropdowns(ss), [], 'no table list: tables keep their own dropdowns');
});

test('Weekly Presence without a Dropdowns tab keeps its own list', () => {
  freshInit(ss);
  getTable(ss, 'members').getRange(2, 1, 1, 7).setValues([['Alice', '', '', 'Dev', 'Eng', 'Active', '']]);
  renderWeeklyMatrixView('2026-10-12', '2026-10-18', ss, '2026-10-12');
  var rule = getTable(ss, 'weeklyMatrix').getRange(3, 3).getDataValidation();
  assert.ok(rule && rule.allowedValues.includes('Remote'));
});

test('tables: an upgrade leaves unchanged columns alone, so colours set in the table stay', () => {
  freshInit(ss);
  var t = Sheets.Spreadsheets.tables.find(x => x.name === 'Tasks');
  t.columnProperties[SCHEMA_SPEC.tasks.columns.indexOf('status')].dataValidationRule.colour = 'green';
  freshInit(ss);
  t = Sheets.Spreadsheets.tables.find(x => x.name === 'Tasks');
  assert.strictEqual(t.columnProperties[SCHEMA_SPEC.tasks.columns.indexOf('status')].dataValidationRule.colour, 'green');
  assert.ok(tableColumnsSignature(t.columnProperties).length > 0);
});
