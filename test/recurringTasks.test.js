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
Object.assign(global, require('../src/TaskTypes'));
Object.assign(global, require('../src/RecurringTasks'));
Object.assign(global, require('../src/TaskCalendar'));
Object.assign(global, require('../src/AssignService'));
Object.assign(global, require('../src/TaskSyncService'));

const MON = '2026-10-12';

function seed() {
  freshInit(mockSpreadsheet);
  seedPublicHolidays(mockSpreadsheet, MON);
  getTable(mockSpreadsheet, 'members').getRange(2, 1, 4, 7).setValues([
    ['Alice', 'alice@accenture.com', '', 'Consultant', 'Eng', 'Active', ''],
    ['Bob', 'bob@accenture.com', '', 'Consultant', 'Eng', 'Active', ''],
    ['Carol', 'carol@accenture.com', '', 'Manager', 'Eng', 'Active', ''],
    ['Dan', 'dan@accenture.com', '', 'Consultant', 'Eng', 'Departed', '']
  ]);
  getTable(mockSpreadsheet, 'projects').getRange(2, 1, 1, 7).setValues([['ACN', 'ACN team', 'Active', '2026-01-01', '', 'Carol', 'Alice']]);
  MailApp.clearSentEmails();
}
function leave(member, from, to, status) {
  var sheet = getTable(mockSpreadsheet, 'leaves');
  var row = Math.max(sheet.getLastRow(), 1) + 1;
  sheet.getRange(row, 1, 1, 8).setValues([['lv' + row, member, from, 'AM', to, 'PM', 'Vacation', status || 'Approved']]);
}
const iso = d => isoDate(d);
function tasksOf(type) { return readTaskRows(mockSpreadsheet).filter(t => t.type === type && t.status !== 'Cancelled').sort((a, b) => iso(a.due) < iso(b.due) ? -1 : 1); }

test('task types: defaults, periods and due dates (weekends and holidays move earlier)', () => {
  seed();
  assert.ok(taskTypeNames(mockSpreadsheet).includes('Time & expenses'));
  var hol = readHolidays(mockSpreadsheet);
  assert.strictEqual(periodOf('Weekly', MON).key, '2026-W42');
  assert.strictEqual(periodFromKey('2026-W42').start.getDate(), 12);
  assert.strictEqual(iso(dueDateFor({ frequency: 'Monthly', due: 'Last working day' }, periodOf('Monthly', MON), hol)), '2026-10-30');
  assert.strictEqual(iso(dueDateFor({ frequency: 'Monthly', due: '20' }, periodOf('Monthly', MON), hol)), '2026-10-20');
  assert.strictEqual(iso(dueDateFor({ frequency: 'Weekly', due: 'Wednesday' }, periodOf('Weekly', '2026-11-09'), hol)), '2026-11-10', 'Armistice Day');
  assert.strictEqual(iso(dueDateFor({ frequency: 'Monthly', due: '1' }, periodOf('Monthly', '2026-11-01'), hol)), '2026-11-02', 'Sunday 1 Nov (also a holiday) → Monday 2');
});

test('rotation: one person per week in turn, skipping approved leave; the skipped person goes first next', () => {
  seed();
  leave('Bob', '2026-10-23', '2026-10-23');
  leave('Alice', '2026-10-30', '2026-10-30', 'Requested'); // a request does not skip
  addRecurringRule(mockSpreadsheet, { type: 'Quiz', projectId: 'ACN', mode: 'Rotation', appliesTo: 'Alice; Bob; Carol', start: MON });
  var res = generateRecurringTasks(mockSpreadsheet, MON);
  assert.strictEqual(res.created, 8);
  assert.deepStrictEqual(tasksOf('Quiz').slice(0, 4).map(t => [iso(t.due), t.member]),
    [['2026-10-16', 'Alice'], ['2026-10-23', 'Carol'], ['2026-10-30', 'Bob'], ['2026-11-06', 'Alice']]);
  assert.strictEqual(generateRecurringTasks(mockSpreadsheet, MON).created, 0, 'no duplicates');
});

test('rotation: a leave approved later moves the turn and emails both people', () => {
  seed();
  addRecurringRule(mockSpreadsheet, { type: 'Quiz', projectId: 'ACN', mode: 'Rotation', appliesTo: 'Alice; Bob; Carol', start: MON });
  generateRecurringTasks(mockSpreadsheet, MON);
  leave('Alice', '2026-10-16', '2026-10-16');
  var res = generateRecurringTasks(mockSpreadsheet, MON);
  assert.strictEqual(res.reassigned, 1);
  var first = tasksOf('Quiz')[0];
  assert.notStrictEqual(first.member, 'Alice');
  assert.strictEqual(MailApp.getSentEmails().length, 2);
});

test('each person: monthly task for a role, three months ahead, departed members left out', () => {
  seed();
  addRecurringRule(mockSpreadsheet, { type: 'Time & expenses', projectId: 'ACN', mode: 'Each person', appliesTo: 'Role: Consultant', start: MON });
  generateRecurringTasks(mockSpreadsheet, MON);
  var t = tasksOf('Time & expenses');
  assert.strictEqual(t.length, 6);
  assert.deepStrictEqual([...new Set(t.map(x => x.member))].sort(), ['Alice', 'Bob']);
  assert.strictEqual(iso(t[0].due), '2026-10-30');
  assert.strictEqual(generateRecurringTasks(mockSpreadsheet, MON).created, 0);
});

test('calendar: members x weeks of the chosen type; assign, refuse on leave, done (stamped), cancel', () => {
  seed();
  leave('Bob', '2026-10-23', '2026-10-23');
  renderTaskCalendar(mockSpreadsheet, MON);
  var cal = getTable(mockSpreadsheet, 'taskCalendar');
  cal.getRange(1, 2).setValue('Quiz');
  handleTaskCalendarEdit({ range: cal.getRange(1, 2), source: mockSpreadsheet }, MON);
  assert.strictEqual(cal.getRange(1, 2).getValue(), 'Quiz');
  assert.deepStrictEqual(cal.getRange(3, 1, 3, 1).getValues().map(r => r[0]), ['Alice', 'Bob', 'Carol'], 'departed hidden');
  var col = 2 + 4; // current week (4 weeks of history first)
  var edit = (row, c, value, old) => { cal.getRange(row, c).setValue(value); return handleTaskCalendarEdit({ range: cal.getRange(row, c), oldValue: old, source: mockSpreadsheet }, MON); };

  assert.strictEqual(edit(3, col, 'To do', undefined), 'assigned');
  assert.strictEqual(iso(tasksOf('Quiz')[0].due), '2026-10-16');
  assert.strictEqual(cal.getRange(3, col).getValue(), 'To do');

  assert.strictEqual(edit(4, col + 1, 'To do', undefined), 'refused_leave', 'Bob on leave Fri 23 Oct');
  assert.strictEqual(cal.getRange(4, col + 1).getValue(), '');

  assert.strictEqual(edit(3, col, 'Done', 'To do'), 'done');
  var t = tasksOf('Quiz')[0];
  assert.strictEqual(t.status, 'Done');
  assert.strictEqual(iso(toDate(t.doneOn)), MON);

  assert.strictEqual(edit(3, col, '', 'Done'), 'cancelled');
  assert.strictEqual(tasksOf('Quiz').length, 0);
});

test('calendar: a closed month asks first; "No" keeps the old value', () => {
  seed();
  renderTaskCalendar(mockSpreadsheet, MON);
  var cal = getTable(mockSpreadsheet, 'taskCalendar');
  SpreadsheetApp.getUi().alertResponse = 'NO';
  cal.getRange(3, 2).setValue('Done'); // 4 weeks back = September
  var out = handleTaskCalendarEdit({ range: cal.getRange(3, 2), oldValue: '', source: mockSpreadsheet }, MON);
  assert.strictEqual(out, 'kept');
  assert.strictEqual(cal.getRange(3, 2).getValue(), '');
  SpreadsheetApp.getUi().alertResponse = null;
});

test('tasks grid: setting Done stamps Done on, reopening clears it', () => {
  seed();
  var made = createTaskRow(mockSpreadsheet, { projectId: 'ACN', member: rtMembers(mockSpreadsheet)[0], title: 'Forecast', due: toDate('2026-10-20'), type: 'Forecast' });
  var tasks = getTable(mockSpreadsheet, 'tasks');
  var cell = tasks.getRange(made.row, taskCol('status'));
  cell.setValue('Done');
  handleTaskStatusEdit({ range: cell, oldValue: 'Todo', source: mockSpreadsheet }, MON);
  assert.strictEqual(tasks.getRange(made.row, taskCol('done_on')).getValue(), MON);
  cell.setValue('Todo');
  handleTaskStatusEdit({ range: cell, oldValue: 'Done', source: mockSpreadsheet }, MON);
  assert.strictEqual(tasks.getRange(made.row, taskCol('done_on')).getValue(), '');
});

test('reminders: due in 3 working days and overdue since the last working day, once a day', () => {
  seed();
  var m = rtMembers(mockSpreadsheet);
  createTaskRow(mockSpreadsheet, { member: m[0], title: 'Forecast', due: toDate('2026-10-15'), type: 'Forecast' }); // Thu = Mon + 3 working days
  createTaskRow(mockSpreadsheet, { member: m[1], title: 'SoW', due: toDate('2026-10-09'), type: 'SoW' });           // previous working day (Fri)
  createTaskRow(mockSpreadsheet, { member: m[2], title: 'PRM', due: toDate('2026-10-14'), type: 'PRM' });           // not a reminder day
  var res = sendTaskReminders(mockSpreadsheet, MON);
  assert.deepStrictEqual(res.sent.sort(), ['Alice', 'Bob']);
  assert.ok(MailApp.getSentEmails().some(e => e.subject === 'Overdue task reminder' && e.to === 'bob@accenture.com'));
  assert.strictEqual(sendTaskReminders(mockSpreadsheet, MON).skipped, 'already_sent_today');
});

test('assign sidebar: once for a role, recurring rotation, stop cancels the future', () => {
  seed();
  var data = getAssignSidebarData(mockSpreadsheet);
  assert.deepStrictEqual(data.roles, ['Consultant', 'Manager']);
  var once = assignTasks({ type: 'Peer review', projectId: 'ACN', mode: 'Each person', appliesTo: 'Role: Consultant', repeat: false, date: MON }, mockSpreadsheet, MON);
  assert.ok(once.success, once.message);
  assert.deepStrictEqual(tasksOf('Peer review').map(t => t.member).sort(), ['Alice', 'Bob']);

  var rot = assignTasks({ type: 'Team presentation', projectId: 'ACN', mode: 'Rotation', appliesTo: '', start: MON }, mockSpreadsheet, MON);
  assert.ok(rot.success, rot.message);
  assert.strictEqual(tasksOf('Team presentation').length, 8);
  var rule = readRecurringRules(mockSpreadsheet)[0];
  var stop = stopRecurringRule(rule.id, mockSpreadsheet, MON);
  assert.ok(stop.success);
  assert.strictEqual(tasksOf('Team presentation').length, 0);
  assert.strictEqual(readRecurringRules(mockSpreadsheet)[0].active, false);
});

function asUser(email, listId) {
  global.Session.getEffectiveUser = () => ({ getEmail: () => email });
  PropertiesService.getUserProperties().setProperty('GOOGLE_TASKS_LIST_ID', listId);
}
function gList(listId) { return Tasks.Tasks.list(listId).items; }

test('Google Tasks per person: each syncs only their own tasks into their own list', () => {
  seed();
  var m = rtMembers(mockSpreadsheet);
  createTaskRow(mockSpreadsheet, { member: m[0], title: 'Forecast', due: toDate('2026-10-20'), type: 'Forecast' });
  createTaskRow(mockSpreadsheet, { member: m[1], title: 'SoW', due: toDate('2026-10-20'), type: 'SoW' });
  asUser('alice@accenture.com', 'listA');
  myTasksSyncJob();
  asUser('bob@accenture.com', 'listB');
  myTasksSyncJob();
  assert.deepStrictEqual(gList('listA').map(t => t.title), ['Forecast']);
  assert.deepStrictEqual(gList('listB').map(t => t.title), ['SoW']);
});

test('Google Tasks per person: a reassigned task leaves the old list and reaches the new one; cancelled ones go', () => {
  seed();
  var m = rtMembers(mockSpreadsheet);
  var a = createTaskRow(mockSpreadsheet, { member: m[0], title: 'Quiz', due: toDate('2026-10-16'), type: 'Quiz' });
  var c = createTaskRow(mockSpreadsheet, { member: m[0], title: 'PRM', due: toDate('2026-10-30'), type: 'PRM' });
  asUser('alice@accenture.com', 'listA2');
  myTasksSyncJob();
  assert.strictEqual(gList('listA2').length, 2);

  setTaskMember(mockSpreadsheet, a.row, m[1]);       // Quiz handed to Bob
  setTaskStatus(mockSpreadsheet, c.row, 'Cancelled');  // PRM cancelled
  myTasksSyncJob();                                     // Alice
  assert.strictEqual(gList('listA2').length, 0);
  asUser('bob@accenture.com', 'listB2');
  myTasksSyncJob();                                     // Bob
  assert.deepStrictEqual(gList('listB2').map(t => t.title), ['Quiz']);
});

test('Google Tasks per person: ticked in Google Tasks marks Done with Done on', () => {
  seed();
  var m = rtMembers(mockSpreadsheet);
  var t = createTaskRow(mockSpreadsheet, { member: m[0], title: 'Time & expenses', due: toDate('2026-10-30'), type: 'Time & expenses' });
  asUser('alice@accenture.com', 'listA3');
  myTasksSyncJob();
  var g = gList('listA3')[0];
  Tasks.Tasks.patch({ status: 'completed' }, 'listA3', g.id);
  myTasksSyncJob();
  var row = readTaskRows(mockSpreadsheet).filter(x => x.row === t.row)[0];
  assert.strictEqual(row.status, 'Done');
  assert.ok(row.doneOn, 'Done on stamped');
});

test('Connect my Google Tasks: installs one hourly job per person and records them', () => {
  seed();
  asUser('carol@accenture.com', 'listC');
  connectMyGoogleTasks();
  connectMyGoogleTasks();
  var jobs = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'myTasksSyncJob');
  assert.strictEqual(jobs.length, 1);
  assert.strictEqual(jobs[0].everyHoursValue, 1);
  assert.deepStrictEqual(connectedTaskUsers(), ['carol@accenture.com']);
});
