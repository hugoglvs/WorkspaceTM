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
Object.assign(global, require('../src/Holidays'));
Object.assign(global, require('../src/PresenceEngine'));
Object.assign(global, require('../src/WorkspaceCheck'));
Object.assign(global, require('../src/Birthdays'));
const kpi = require('../src/KpiEngine');

function seed(opts) {
  freshInit(mockSpreadsheet);
  seedPublicHolidays(mockSpreadsheet, '2026-10-12');
  getTable(mockSpreadsheet, 'members').getRange(2, 1, 3, 7).setValues([
    ['Alice', 'alice@accenture.com', '', 'Tech Lead', 'Eng', 'Active', ''],
    ['Bob', 'bob@accenture.com', '', 'Dev', 'Eng', 'Active', ''],
    ['Dan', 'dan@accenture.com', '', 'Dev', 'Eng', 'Departed', '']
  ]);
  getTable(mockSpreadsheet, 'tenures').getRange(2, 1, 3, 7).setValues([
    ['Alice', 'Tech Lead', 'Eng', '2025-01-01', (opts && opts.aliceEnd) || '', 'Permanent', 100],
    ['Bob', 'Dev', 'Eng', '2025-01-01', '', 'Permanent', 100],
    ['Dan', 'Dev', 'Eng', '2025-01-01', '2026-06-30', 'Permanent', 100]
  ]);
  getTable(mockSpreadsheet, 'projects').getRange(2, 1, 1, 7).setValues([['P1', 'Apollo', 'Active', '2026-01-01', '', 'Alice', '']]);
}

test('French public holidays: 11 a year, Easter-based ones computed', () => {
  var h = frenchPublicHolidays(2026);
  assert.strictEqual(h.length, 11);
  var byName = Object.fromEntries(h.map(x => [x[1], x[0]]));
  assert.strictEqual(byName['Easter Monday'], '2026-04-06');
  assert.strictEqual(byName['Ascension Day'], '2026-05-14');
  assert.strictEqual(byName['Whit Monday'], '2026-05-25');
  assert.strictEqual(byName['Armistice Day'], '2026-11-11');
  assert.strictEqual(frenchPublicHolidays(2027).filter(x => x[1] === 'Easter Monday')[0][0], '2027-03-29');
});

test('Holidays tab: seeded once per year, a deleted day stays deleted', () => {
  seed();
  var sheet = getTable(mockSpreadsheet, 'holidays');
  assert.strictEqual(sheet.getRange(1, 1, 1, 2).getValues()[0].join('|'), 'Date|Holiday');
  assert.strictEqual(Object.keys(readHolidays(mockSpreadsheet)).filter(d => d.startsWith('2026')).length, 11);
  assert.strictEqual(Object.keys(readHolidays(mockSpreadsheet)).filter(d => d.startsWith('2027')).length, 11);
  var rows = sheet.getRange(2, 1, 30, 2).getValues();
  var whit = rows.findIndex(r => r[1] === 'Whit Monday') + 2;
  sheet.getRange(whit, 1, 1, 2).setValues([['', '']]);
  assert.strictEqual(seedPublicHolidays(mockSpreadsheet, '2026-10-12').added, 0);
  assert.ok(!readHolidays(mockSpreadsheet)['2026-05-25']);
});

test('presence: nothing is compiled on a public holiday and Today says so', () => {
  seed();
  compilePresenceLogs('2026-11-11', '2026-11-11', mockSpreadsheet, '2026-11-11');
  var log = getTable(mockSpreadsheet, 'presenceLog').getRange(2, 1, 5, 6).getValues().filter(r => r[0]);
  assert.strictEqual(log.length, 0);
  assert.strictEqual(getTable(mockSpreadsheet, 'members').getRange(2, 7).getValue(), 'Public holiday');
});

test('KPI: no capacity on a public holiday', () => {
  seed();
  getTable(mockSpreadsheet, 'presenceLog').getRange(2, 1, 1, 6).setValues([['l1', '2026-11-10', 'Bob', 'Remote', 'Remote', '']]);
  var facts = kpi.buildPresenceFacts(mockSpreadsheet, { asOfDate: '2026-11-12' });
  var bobHoliday = facts.rows.filter(r => r[5] === 'Bob' && r[0] === '2026-11-11');
  var bobThu = facts.rows.filter(r => r[5] === 'Bob' && r[0] === '2026-11-12');
  assert.strictEqual(bobHoliday.length, 0, 'unlogged holiday has no rows');
  assert.strictEqual(bobThu[0][14], 0.5);
});

test('coverage: a lead whose contract has ended no longer covers the project', () => {
  seed({ aliceEnd: '2026-10-30' });
  var rep = checkProjectBackupConflicts(mockSpreadsheet, { sendNotifications: false, createTasks: false, asOfDateParam: '2026-10-12' });
  assert.strictEqual(rep.conflictsFound, 1);
  assert.strictEqual(rep.conflicts[0].periods[0].start, '2026-11-02', 'Sat 31, Sun 1 Nov (also a holiday) are skipped');
  assert.ok(rep.conflicts[0].reason.includes("Alice's contract ends"));
});

test('coverage: a leave on a public holiday is not a gap', () => {
  seed();
  getTable(mockSpreadsheet, 'leaves').getRange(2, 1, 1, 8).setValues([['lv1', 'Alice', '2026-11-11', 'AM', '2026-11-11', 'PM', 'Vacation', 'Approved']]);
  var rep = checkProjectBackupConflicts(mockSpreadsheet, { sendNotifications: false, createTasks: false, asOfDateParam: '2026-10-12' });
  assert.strictEqual(rep.conflictsFound, 0);
});

test('Check Workspace: hand-overs for people leaving, open tasks of people who left', () => {
  seed({ aliceEnd: '2026-10-30' });
  getTable(mockSpreadsheet, 'tasks').getRange(2, 1, 2, 7).setValues([
    ['T1', 'P1', 'Dan', 'Write the runbook', '', '', 'In Progress'],
    ['T2', 'P1', 'Dan', 'Old task', '', '', 'Done']
  ]);
  var items = runWorkspaceCheck(mockSpreadsheet, '2026-10-12').sections.filter(s => s.key === 'leavers')[0].items;
  assert.deepStrictEqual(items, ['Apollo: lead Alice leaves on 2026-10-30', 'Task "Write the runbook": Dan left on 2026-06-30']);
});

test('undeclared days and birthday announcements skip public holidays', () => {
  seed();
  var undeclared = findUndeclaredDays(mockSpreadsheet, '2026-11-09').filter(u => u.name === 'Bob')[0];
  assert.deepStrictEqual(undeclared.days, ['2026-11-09', '2026-11-10', '2026-11-12', '2026-11-13']);
  var hol = readHolidays(mockSpreadsheet);
  const iso = d => `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}`;
  assert.deepStrictEqual(birthdayAnnouncementDays('2026-11-10', hol).map(iso), ['2026-11-11', '2026-11-12']);
  assert.deepStrictEqual(birthdayAnnouncementDays('2026-11-11', hol), []);
});
