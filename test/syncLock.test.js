const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');
Object.assign(global, require('../src/Tables'));
Object.assign(global, require('../src/Locations'));
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

global.TaskSyncService = require('../src/TaskSyncService');

const code = require('../src/Code');

function typeTaskTitle(spreadsheet, title) {
  freshInit(spreadsheet);
  spreadsheet.getSheetByName('Members').getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Senior Eng', 'Core', 'Active', 'Main Office']
  ]);
  var titleRange = spreadsheet.getSheetByName('Tasks').getRange(2, 4);
  titleRange.setValue(title);
  return { range: titleRange, value: title, source: spreadsheet };
}

function taskRow(spreadsheet) {
  return spreadsheet.getSheetByName('Tasks').getRange(2, 1, 1, 8).getValues()[0];
}

test('nightlySyncJob: exits with lock_timeout and runs no step while another sync holds the spreadsheet lock', () => {
  LockService.__state.document.held = true;

  var res = code.nightlySyncJob(mockSpreadsheet);

  assert.strictEqual(res.success, false);
  assert.strictEqual(res.reason, 'lock_timeout');
  assert.deepStrictEqual(res.steps, {});
});

test('nightlySyncJob: releases the lock so the next run can proceed', () => {
  var first = code.nightlySyncJob(mockSpreadsheet);
  var second = code.nightlySyncJob(mockSpreadsheet);

  assert.strictEqual(first.success, true);
  assert.strictEqual(second.success, true);
  assert.strictEqual(LockService.__state.document.held, false);
});

test('nightlySyncJob: is not blocked by another Consumer Sheet running the shared library', () => {
  // A library-created script lock is shared by every including script.
  LockService.__state.script.held = true;

  var res = code.nightlySyncJob(mockSpreadsheet);

  assert.strictEqual(res.success, true);
  assert.ok(Object.keys(res.steps).length > 0);
});

test('nightlySyncJob: falls back to the script lock outside a containing document', () => {
  LockService.__state.documentLockAvailable = false;
  LockService.__state.script.held = true;

  var res = code.nightlySyncJob(mockSpreadsheet);

  assert.strictEqual(res.reason, 'lock_timeout');
});

test('onEdit: during a nightly sync, keeps the user edit, skips script writes and tells the user', () => {
  var editEvent = typeTaskTitle(mockSpreadsheet, 'Build authentication modal');
  LockService.__state.document.held = true;

  var res = code.onEdit(editEvent);

  var row = taskRow(mockSpreadsheet);
  assert.strictEqual(res.reason, 'sync_in_progress');
  assert.strictEqual(row[3], 'Build authentication modal');
  assert.strictEqual(row[0], '', 'No task_id is written while the sync holds the lock');
  assert.ok(mockSpreadsheet.lastToast.title.includes('Sync in progress'));
});

test('onEdit: processes the edit normally and releases the lock when no sync is running', () => {
  var editEvent = typeTaskTitle(mockSpreadsheet, 'Build authentication modal');

  code.onEdit(editEvent);

  assert.ok(taskRow(mockSpreadsheet)[0].startsWith('TSK-'));
  assert.strictEqual(LockService.__state.document.held, false);
});

test('installedTaskEditHandler: defers Google Task creation while a nightly sync is running', () => {
  var editEvent = typeTaskTitle(mockSpreadsheet, 'Build authentication modal');
  LockService.__state.document.held = true;

  var res = code.installedTaskEditHandler(editEvent);

  assert.strictEqual(res.reason, 'sync_in_progress');
  assert.strictEqual(taskRow(mockSpreadsheet)[7], '', 'No Google Task is created');
});

test('installedTaskEditHandler: releases the lock after handling the edit', () => {
  var editEvent = typeTaskTitle(mockSpreadsheet, 'Build authentication modal');

  code.installedTaskEditHandler(editEvent);

  assert.ok(taskRow(mockSpreadsheet)[7], 'Google Task is created');
  assert.strictEqual(LockService.__state.document.held, false);
});

test('one run per edit: once the installable trigger exists, the simple onEdit steps aside', () => {
  var editEvent = typeTaskTitle(mockSpreadsheet, 'Build authentication modal');
  assert.strictEqual(code.editTriggerActive(), false);

  code.installAllTriggers(mockSpreadsheet);
  assert.strictEqual(code.editTriggerActive(), true, 'installing records the trigger');

  var simple = code.onEdit(editEvent);
  assert.strictEqual(simple.reason, 'handled_by_installed_trigger');
  assert.strictEqual(taskRow(mockSpreadsheet)[0], '', 'the simple trigger did nothing');

  code.installedTaskEditHandler(editEvent);
  assert.ok(taskRow(mockSpreadsheet)[0].startsWith('TSK-'), 'the installed trigger handled it');
});

test('one run per edit: the flag clears when the trigger is removed, so the simple onEdit takes over again', () => {
  code.installAllTriggers(mockSpreadsheet);
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  code.refreshEditTriggerFlag();
  assert.strictEqual(code.editTriggerActive(), false);
});

test('the installed trigger also handles non-task edits (Locations warning)', () => {
  freshInit(mockSpreadsheet);
  var sheet = getTable(mockSpreadsheet, 'locations');
  sheet.getRange(4, 3).setValue(false); // JLL unused
  var res = code.installedTaskEditHandler({ range: sheet.getRange(4, 3), value: 'FALSE', source: mockSpreadsheet });
  assert.notStrictEqual(res && res.reason, 'sync_in_progress');
});
