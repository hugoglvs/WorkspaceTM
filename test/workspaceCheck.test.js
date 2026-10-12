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
Object.assign(global, require('../src/WorkspaceCheck'));
Object.assign(global, require('../src/BackupService'));

// Monday 2026-10-12 is "today" in these tests
const TODAY = '2026-10-12';

function seed() {
  freshInit(mockSpreadsheet);
  getTable(mockSpreadsheet, 'members').getRange(2, 1, 3, 7).setValues([
    ['Alice', 'alice@accenture.com', '', 'Tech Lead', 'Eng', 'Active', 'OneRoof'],
    ['Bob', 'bob@accenture.com', '', 'Dev', 'Eng', 'Active', 'Not Logged'],
    ['Carl', '', '', 'Dev', 'Eng', 'No Contract', '']
  ]);
  getTable(mockSpreadsheet, 'projects').getRange(2, 1, 2, 7).setValues([
    ['P1', 'Apollo', 'Active', '2026-01-01', '', 'Alice', ''],
    ['P2', 'Borealis', 'Active', '2026-01-01', '', 'Alice', 'Bob']
  ]);
  getTable(mockSpreadsheet, 'leaves').getRange(2, 1, 1, 8).setValues([
    ['lv1', 'Alice', '2026-10-14', 'AM', '2026-10-16', 'PM', 'Vacation', 'Approved']
  ]);
  getTable(mockSpreadsheet, 'baselines').getRange(2, 1, 1, 7).setValues([
    ['b1', 'Alice', '2026-01-01', '', 'All', 'OneRoof', 'OneRoof']
  ]);
  getTable(mockSpreadsheet, 'presenceLog').getRange(2, 1, 3, 6).setValues([
    ['l1', '2026-10-05', 'Alice', 'OneRoof', 'OneRoof', ''],
    ['l2', '2026-10-05', 'Alice', 'Remote', 'Remote', ''],   // duplicate: this later row wins
    ['l3', '2026-10-13', 'Bob', 'C04', 'C04', '']
  ]);
}

function section(report, key) { return report.sections.filter(s => s.key === key)[0]; }

test('Check Workspace lists contracts, backups, coverage, not logged, undeclared days and duplicates', () => {
  seed();
  var report = runWorkspaceCheck(mockSpreadsheet, TODAY);

  assert.deepStrictEqual(section(report, 'contracts').items, ['Carl']);
  assert.deepStrictEqual(section(report, 'backups').items, ['Apollo']);
  assert.strictEqual(section(report, 'coverage').items.length, 1);
  assert.ok(section(report, 'coverage').items[0].includes('Apollo') && section(report, 'coverage').items[0].includes('2026-10-14'));
  assert.deepStrictEqual(section(report, 'notLogged').items, ['Bob']);
  assert.deepStrictEqual(section(report, 'undeclared').items, ['Bob: Mon 12, Wed 14, Thu 15, Fri 16'], 'Alice has a routine; Bob logged Tuesday');
  assert.strictEqual(section(report, 'duplicates').action, 'dedupe');
  assert.strictEqual(MailApp.getSentEmails().length, 0, 'checking sends no email');
});

test('Check Workspace: removing duplicates keeps the last row of each day', () => {
  seed();
  var res = removePresenceDuplicates(mockSpreadsheet);
  assert.strictEqual(res.removed, 1);
  var rows = getTable(mockSpreadsheet, 'presenceLog').getRange(2, 1, 5, 6).getValues().filter(r => r[0]);
  assert.deepStrictEqual(rows.map(r => r[0]), ['l2', 'l3']);
  assert.strictEqual(findPresenceDuplicates(mockSpreadsheet).count, 0);
});

test('Check Workspace opens as a dialog', () => {
  seed();
  showWorkspaceCheck();
  var dlg = SpreadsheetApp.getUi().shownDialog;
  assert.strictEqual(dlg.title, 'Check Workspace');
  assert.ok(JSON.parse(dlg.html.template.reportJson).sections.length >= 6);
});

test('Monday reminder emails members with undeclared days, linking to Weekly Presence', () => {
  seed();
  var res = sendPresenceReminders(mockSpreadsheet, TODAY);
  assert.deepStrictEqual(res.sent, ['Bob']);
  var mail = MailApp.getSentEmails()[0];
  assert.strictEqual(mail.to, 'bob@accenture.com');
  assert.ok(mail.body.includes('Mon 12, Wed 14, Thu 15, Fri 16'));
  assert.ok(mail.body.includes(mockSpreadsheet.getUrl()));
});

test('Monday reminder: nobody to remind when everyone is declared', () => {
  seed();
  getTable(mockSpreadsheet, 'baselines').getRange(3, 1, 1, 7).setValues([['b2', 'Bob', '2026-01-01', '', 'All', 'Remote', 'Remote']]);
  assert.deepStrictEqual(sendPresenceReminders(mockSpreadsheet, TODAY).sent, []);
});

test('Interim Backups sidebar: shows gaps, adds a backup that closes them, cancel reopens them', () => {
  seed();
  var data = getBackupSidebarData(mockSpreadsheet, TODAY);
  var apollo = data.projects.filter(p => p.id === 'P1')[0];
  assert.strictEqual(data.projects[0].id, 'P1', 'projects with gaps come first');
  assert.deepStrictEqual(apollo.gaps.map(g => [g.start, g.end]), [['2026-10-14', '2026-10-16']]);
  assert.deepStrictEqual(data.leaves.alice, [{ start: '2026-10-14', end: '2026-10-16', status: 'approved' }]);
  assert.ok(data.members.includes('Bob'));

  var res = saveInterimBackup({ projectId: 'P1', backupMember: 'Bob', startDate: '2026-10-14', endDate: '2026-10-16' }, mockSpreadsheet, TODAY);
  assert.ok(res.success, res.message);
  assert.strictEqual(res.status, 'Planned');
  assert.strictEqual(MailApp.getSentEmails().length, 0, 'saving sends no email');

  data = getBackupSidebarData(mockSpreadsheet, TODAY);
  apollo = data.projects.filter(p => p.id === 'P1')[0];
  assert.strictEqual(apollo.gaps.length, 0);
  assert.strictEqual(apollo.interims.length, 1);

  cancelInterimBackup(apollo.interims[0].id, mockSpreadsheet);
  apollo = getBackupSidebarData(mockSpreadsheet, TODAY).projects.filter(p => p.id === 'P1')[0];
  assert.strictEqual(apollo.interims.length, 0);
  assert.strictEqual(apollo.gaps.length, 1);
});

test('Interim Backups sidebar: refuses the lead, a missing end date, or an end before the start', () => {
  seed();
  assert.strictEqual(saveInterimBackup({ projectId: 'P1', backupMember: 'Alice', startDate: '2026-10-14', endDate: '2026-10-16' }, mockSpreadsheet, TODAY).success, false);
  assert.strictEqual(saveInterimBackup({ projectId: 'P1', backupMember: 'Bob', startDate: '2026-10-14', endDate: '' }, mockSpreadsheet, TODAY).success, false);
  assert.strictEqual(saveInterimBackup({ projectId: 'P1', backupMember: 'Bob', startDate: '2026-10-16', endDate: '2026-10-14' }, mockSpreadsheet, TODAY).success, false);
});
