const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');
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
Object.assign(global, require('../src/TaskTypes'));
Object.assign(global, require('../src/RecurringTasks'));
const triggers = require('../src/TaskTriggers');
Object.assign(global, triggers);
const { freshInit } = require('../src/Init');

let ss;
test.beforeEach(() => {
  ss = createMockSpreadsheet();
  setupGasGlobals(ss);
  freshInit(ss);
});

function addPerson(name, role, start) {
  var m = ss.getSheetByName('Members');
  var r = findNextAvailableRow(m);
  m.getRange(r, 1).setValue(name);
  var c = ss.getSheetByName('Contracts');
  c.getRange(findNextAvailableRow(c), 1, 1, 7).setValues([[name, role, 'FAL ME', start, '', 'Permanent', 100]]);
}

function taskTypesRow(name) {
  var sheet = ss.getSheetByName('Task Types');
  for (var r = 2; r < 30; r++) if (sheet.getRange(r, 1).getValue() === name) return r;
  return 0;
}

test('Task Types: Onboarding is created by "New member", 7 days after, for the person concerned', () => {
  var t = findTaskType(ss, 'Onboarding');
  assert.strictEqual(t.createdWhen, 'New member');
  assert.strictEqual(t.daysAfter, 7);
  assert.strictEqual(t.assignTo, 'Person concerned');
  assert.ok(t.instructions.includes('Members tab'));
  assert.deepStrictEqual(triggers.taskTriggerLabels(), ['New member']);
});

test('New member: one onboarding task per new person, due 7 days after the contract start, never twice', () => {
  addPerson('Ines', 'Consultant', '2026-11-02');
  addPerson('Old timer', 'Consultant', '2024-01-01');

  assert.strictEqual(triggers.fireNewMemberTriggers(ss, '2026-10-13'), 1, 'not for someone who started long ago');
  var rows = readTaskRows(ss);
  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0].member, 'Ines');
  assert.strictEqual(rows[0].type, 'Onboarding');
  assert.strictEqual(isoDate(rows[0].due), '2026-11-09');
  assert.strictEqual(rows[0].origin, 'trigger:newMember:Ines:Onboarding');

  assert.strictEqual(triggers.fireNewMemberTriggers(ss, '2026-10-14'), 0, 'created once');
});

test('Assign to: roles and names besides the person concerned; Days after can be negative', () => {
  var sheet = ss.getSheetByName('Task Types');
  var r = findNextAvailableRow(sheet);
  sheet.getRange(r, 1, 1, 9).setValues([['Prepare workstation', 'One-off', '', false, true, 'New member', -5, 'Role: Manager', 'Laptop and badge']]);
  addPerson('Boss', 'Manager', '2020-01-01');
  addPerson('Ines', 'Consultant', '2026-11-02');
  applyMemberFormulas(ss); // Role comes from Contracts by formula (done by the members sync)

  triggers.fireNewMemberTriggers(ss, '2026-10-13');

  var prep = readTaskRows(ss).filter(t => t.type === 'Prepare workstation');
  assert.strictEqual(prep.length, 1);
  assert.strictEqual(prep[0].member, 'Boss');
  assert.strictEqual(isoDate(prep[0].due), '2026-10-28');
});

test('A type whose "Created when" is cleared, or inactive, creates nothing', () => {
  ss.getSheetByName('Task Types').getRange(taskTypesRow('Onboarding'), 6).setValue('');
  addPerson('Ines', 'Consultant', '2026-11-02');
  assert.strictEqual(triggers.fireNewMemberTriggers(ss, '2026-10-13'), 0);
});

test('Upgrade fills the trigger columns of an existing Onboarding row once', () => {
  var sheet = ss.getSheetByName('Task Types');
  var r = taskTypesRow('Onboarding');
  sheet.getRange(r, 6, 1, 4).setValues([['', '', '', '']]);
  PropertiesService.getDocumentProperties().deleteProperty('TASK_TRIGGERS_SEEDED');

  seedDefaultTaskTypes(ss);
  assert.strictEqual(sheet.getRange(r, 6).getValue(), 'New member');

  sheet.getRange(r, 6).setValue('');
  seedDefaultTaskTypes(ss);
  assert.strictEqual(sheet.getRange(r, 6).getValue(), '', 'cleared by the team: left alone');
});

test('Saving a first contract in the sidebar creates the onboarding task at once', () => {
  ss.getSheetByName('Members').getRange(2, 1).setValue('Ines');
  var start = isoDate(new Date());
  var res = saveContract({ member: 'Ines', role: 'Consultant', department: 'FAL ME', startDate: start, contractType: 'Permanent', ftePercent: 100 }, ss);
  assert.strictEqual(res.success, true, res.message);
  var onboarding = readTaskRows(ss).filter(t => t.type === 'Onboarding');
  assert.strictEqual(onboarding.length, 1);
  assert.strictEqual(isoDate(onboarding[0].due), shiftIso(start, 7));
});
