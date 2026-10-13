const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals, setLegacyTasks, getLegacyTasks, legacyTaskCol } = require('./mockGas');
Object.assign(global, require('../src/Tables'));
Object.assign(global, require('../src/Locations'));
const schema = require('../src/Schema');
global.schemaHeaderRow = schema.schemaHeaderRow;
global.schemaColumnLabel = schema.schemaColumnLabel;
global.CURRENT_SCHEMA_VERSION = schema.CURRENT_SCHEMA_VERSION;
global.FONT_FAMILY = schema.FONT_FAMILY;
global.FONT_SIZE = schema.FONT_SIZE;
global.SCHEMA_SPEC = schema.SCHEMA_SPEC;
global.DATE_FORMAT = schema.DATE_FORMAT;
Object.assign(global, require('../src/Typography'));
Object.assign(global, require('../src/SyncEngine'));
const sidebar = require('../src/Sidebar');
global.extractCellValue = sidebar.extractCellValue;
global.formatDateIso = sidebar.formatDateIso;
Object.assign(global, require('../src/Contracts'));
const settings = require('../src/Settings');
Object.assign(global, settings);
Object.assign(global, require('../src/Holidays'));
const { freshInit } = require('../src/Init');

let ss;
test.beforeEach(() => {
  ss = createMockSpreadsheet();
  setupGasGlobals(ss);
  freshInit(ss);
});

function settingRow(label) {
  var sheet = ss.getSheetByName('Settings');
  for (var r = 2; r < 20; r++) if (sheet.getRange(r, 1).getValue() === label) return r;
  return 0;
}

test('Settings tab: created with every default once, existing values kept on upgrade', () => {
  var sheet = ss.getSheetByName('Settings');
  assert.ok(sheet, 'Settings tab exists');
  assert.strictEqual(sheet.getRange(1, 1, 1, 3).getValues()[0].join('|'), 'Setting|Value|Description');
  assert.strictEqual(settings.getSetting(ss, 'default_role'), 'Team Member');
  sheet.getRange(settingRow('Default role'), 2).setValue('Consultant');

  assert.strictEqual(settings.seedDefaultSettings(ss), 0, 'nothing added twice');
  assert.strictEqual(settings.getSetting(ss, 'default_role'), 'Consultant');
});

test('Settings: an empty value falls back to the default', () => {
  ss.getSheetByName('Settings').getRange(settingRow('Default department'), 2).setValue('');
  assert.strictEqual(settings.getSetting(ss, 'default_department'), 'Delivery');
});

test('Settings: the default role is used by the Members formulas, and an edit re-applies them', () => {
  ss.getSheetByName('Contracts').getRange(2, 1, 1, 7).setValues([['Ines', '', 'FAL ME', '2025-01-01', '', 'Permanent', 100]]);
  syncMembersAndTenures(ss);
  var members = ss.getSheetByName('Members');
  assert.strictEqual(members.getRange(2, 4).getValue(), 'Team Member');

  ss.getSheetByName('Settings').getRange(settingRow('Default role'), 2).setValue('Analyst "junior"');
  settings.handleSettingsEdit({ source: ss });

  assert.strictEqual(members.getRange(2, 4).getValue(), 'Analyst "junior"');
  assert.ok(members.getRange(2, 4).getFormula().includes('"Analyst ""junior"""'), 'quotes escaped in the formula');
});

test('Settings: the former onboarding rows are cleared on upgrade (onboarding is a Task Type now)', () => {
  var sheet = ss.getSheetByName('Settings');
  sheet.getRange(10, 1, 1, 3).setValues([['Onboarding task title', 'Say hello', 'old']]);
  settings.seedDefaultSettings(ss);
  assert.strictEqual(settingRow('Onboarding task title'), 0);
});

test('Settings: Public holidays = None stops adding French holidays', () => {
  ss.getSheetByName('Settings').getRange(settingRow('Public holidays'), 2).setValue('None');
  PropertiesService.getDocumentProperties().deleteProperty('HOLIDAYS_SEEDED_YEARS');
  var res = seedPublicHolidays(ss, '2030-06-01');
  assert.strictEqual(res.added, 0);
  assert.strictEqual(res.skipped, 'none');
});

test('Edit Settings: shows the tab', () => {
  var sheet = ss.getSheetByName('Settings');
  sheet.hideSheet();
  settings.showSettingsTab();
  assert.strictEqual(sheet.isSheetHidden(), false);
});
