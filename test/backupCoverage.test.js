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
  global.MailApp.clearSentEmails();
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
Object.assign(global, require('../src/SyncEngine'));
const { freshInit } = require('../src/Init');
const sidebar = require('../src/Sidebar');
global.extractCellValue = sidebar.extractCellValue;
global.getLeaveColumnMap = sidebar.getLeaveColumnMap;
const engine = require('../src/ProjectBackupEngine');
Object.assign(global, engine);
const code = require('../src/Code');

// Monday 2026-06-08 is "today" in these tests
const TODAY = '2026-06-08';

function setup(project, leaves, backups) {
  freshInit(mockSpreadsheet);
  mockSpreadsheet.getSheetByName('Members').getRange(2, 1, 5, 7).setValues([
    ['Alice', 'alice@accenture.com', '', 'Tech Lead', 'Eng', 'Active', ''],
    ['Bob', 'bob@accenture.com', '', 'Dev', 'Eng', 'Active', ''],
    ['Carol', 'carol@accenture.com', '', 'Dev', 'Eng', 'Active', ''],
    ['Dan', 'dan@accenture.com', '', 'Dev', 'Eng', 'Active', ''],
    ['Maya', 'maya@accenture.com', '', 'Engineering Manager', 'Eng', 'Active', '']
  ]);
  mockSpreadsheet.getSheetByName('Projects').getRange(2, 1, 1, 7).setValues([project]);
  if (leaves.length) {
    mockSpreadsheet.getSheetByName('Leaves').getRange(2, 1, leaves.length, 8).setValues(
      leaves.map((l, i) => ['lv' + i, l[0], l[1], 'AM', l[2], 'PM', 'Vacation', l[3]])
    );
  }
  if (backups && backups.length) {
    mockSpreadsheet.getSheetByName('Backups').getRange(2, 1, backups.length, 8).setValues(backups);
  }
}

function project() { return mockSpreadsheet.getSheetByName('Projects').getRange(2, 1, 1, 8).getValues()[0]; }
function backupRows() { return mockSpreadsheet.getSheetByName('Backups').getRange(2, 1, 20, 8).getValues().filter(r => r[0]); }
function check(asOf) { return engine.checkProjectBackupConflicts(mockSpreadsheet, { asOfDateParam: asOf || TODAY }); }

const PHOENIX = ['PRJ-1', 'Phoenix', 'Active', '2026-01-01', '', 'Alice', 'Bob'];

test('active_responsible: the lead when present, the backup when the lead is on leave', () => {
  setup(PHOENIX, [['Alice', '2026-06-08', '2026-06-12', 'Approved']]);

  check('2026-06-15');
  assert.strictEqual(project()[7], 'Alice');

  var res = check(TODAY);
  assert.strictEqual(project()[7], 'Bob (backup)');
  assert.strictEqual(res.conflictsFound, 0);
});

test('coverage: consecutive interim backups together cover a gap', () => {
  setup(PHOENIX, [
    ['Alice', '2026-06-08', '2026-06-12', 'Approved'],
    ['Bob', '2026-06-08', '2026-06-12', 'Approved']
  ], [
    ['b1', 'PRJ-1', 'Alice', 'Carol', '2026-06-08', '2026-06-10', 'Active', ''],
    ['b2', 'PRJ-1', 'Alice', 'Dan', '2026-06-11', '2026-06-12', 'Active', '']
  ]);

  var res = check();

  assert.strictEqual(res.conflictsFound, 0);
  assert.strictEqual(res.warningsFound, 0);
  assert.strictEqual(project()[7], 'Carol (interim)');
});

test('coverage: an uncovered day is a conflict with red row, email and task', () => {
  setup(PHOENIX, [
    ['Alice', '2026-06-08', '2026-06-12', 'Approved'],
    ['Bob', '2026-06-10', '2026-06-15', 'Approved']
  ]);

  var res = check();

  assert.strictEqual(res.conflictsFound, 1);
  assert.deepStrictEqual(res.conflicts[0].periods, [{ kind: 'conflict', start: '2026-06-10', end: '2026-06-12' }]);
  assert.strictEqual(mockSpreadsheet.getSheetByName('Projects').getRange(2, 1, 1, 8).getBackground(), '#F8D7DA');
  assert.ok(global.MailApp.getSentEmails()[0].subject.startsWith('[ACTION REQUIRED]'));
  var task = mockSpreadsheet.getSheetByName('Tasks').getRange(2, 1, 1, 8).getValues()[0];
  assert.strictEqual(task[3], 'Organize interim backup for project Phoenix');
  assert.strictEqual(task[2], 'Bob', 'lead already away today, so the backup organises it');
});

test('coverage: a leave request that would leave nobody available is a warning, without task', () => {
  setup(PHOENIX, [
    ['Alice', '2026-06-08', '2026-06-12', 'Approved'],
    ['Bob', '2026-06-11', '2026-06-11', 'Requested']
  ]);

  var res = check();

  assert.strictEqual(res.conflictsFound, 0);
  assert.strictEqual(res.warningsFound, 1);
  assert.deepStrictEqual(res.warnings[0].periods, [{ kind: 'warning', start: '2026-06-11', end: '2026-06-11' }]);
  assert.strictEqual(mockSpreadsheet.getSheetByName('Projects').getRange(2, 1, 1, 8).getBackground(), '#FCE8B2');
  assert.ok(global.MailApp.getSentEmails()[0].subject.startsWith('[WARNING]'));
  assert.strictEqual(mockSpreadsheet.getSheetByName('Tasks').getRange(2, 1).getValue(), '', 'no task for a warning');
});

test('permanent backup: mirrored into tbl_backups, and a change keeps the history', () => {
  setup(PHOENIX, []);

  check();
  assert.deepStrictEqual(backupRows().map(r => [r[1], r[2], r[3], r[4], r[5], r[6]]), [
    ['PRJ-1', 'Alice', 'Bob', '2026-01-01', '', 'Active']
  ]);

  mockSpreadsheet.getSheetByName('Projects').getRange(2, 7).setValue('Carol');
  check();
  assert.deepStrictEqual(backupRows().map(r => [r[3], r[4], r[5]]), [
    ['Bob', '2026-01-01', '2026-06-07'],
    ['Carol', '2026-06-08', '']
  ]);
});

test('new project: gets an ID, and a toast asks for a backup when none is set', () => {
  setup(['', 'Atlas', 'Planning', '2026-06-01', '', 'Alice', ''], []);
  var cell = mockSpreadsheet.getSheetByName('Projects').getRange(2, 2);

  code.onEdit({ range: cell, value: 'Atlas', source: mockSpreadsheet });

  assert.match(String(project()[0]), /^PRJ-[A-Z0-9]{5}$/);
  assert.ok(mockSpreadsheet.lastToast.msg.includes('Atlas') && mockSpreadsheet.lastToast.msg.includes('backup'));
});
