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
  getTable(mockSpreadsheet, 'projects').getRange(2, 1, 1, 6).setValues([['ACN', 'Active', '2026-01-01', '', 'Carol', 'Alice']]);
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

test('twice a month: half-month periods, the 15th and the last day, weekends move earlier', () => {
  seed();
  var hol = readHolidays(mockSpreadsheet);
  var te = { frequency: 'Twice a month', due: '15; Last day' };
  assert.strictEqual(periodOf('Twice a month', MON).key, '2026-10-H1');
  assert.strictEqual(periodOf('Twice a month', '2026-10-16').label, '16–31 Oct 2026');
  assert.strictEqual(periodShift('Twice a month', MON, 4).key, '2026-12-H1');
  assert.strictEqual(periodShift('Twice a month', MON, -1).key, '2026-09-H2');
  assert.strictEqual(periodFromKey('2026-10-H2').start.getDate(), 16);
  assert.strictEqual(iso(dueDateFor(te, periodOf(te.frequency, MON), hol)), '2026-10-15');
  assert.strictEqual(iso(dueDateFor(te, periodOf(te.frequency, '2026-10-20'), hol)), '2026-10-30', 'Sat 31 → Fri 30');
  assert.strictEqual(iso(dueDateFor(te, periodOf(te.frequency, '2026-11-02'), hol)), '2026-11-13', 'Sun 15 → Fri 13');
  assert.strictEqual(iso(dueDateFor({ frequency: 'Twice a month', due: '' }, periodOf('Twice a month', '2026-11-20'), hol)), '2026-11-30');
  assert.deepStrictEqual(parseRemindDays('3; 1; 0'), [3, 1, 0]);
  assert.deepStrictEqual(parseRemindDays(''), [3]);
  var t = findTaskType(mockSpreadsheet, 'Time & expenses');
  assert.strictEqual(t.frequency, 'Twice a month');
  assert.deepStrictEqual(t.remind, [3, 1]);
  assert.strictEqual(findTaskType(mockSpreadsheet, 'Peer review').frequency, 'Weekly');
});

test('rotation: one person per week in turn, skipping approved leave; the skipped person goes first next', () => {
  seed();
  leave('Bob', '2026-10-23', '2026-10-23');
  leave('Alice', '2026-10-30', '2026-10-30', 'Requested'); // a request does not skip
  addRecurringRule(mockSpreadsheet, { type: 'Quiz', project: 'ACN', mode: 'Rotation', appliesTo: 'Alice; Bob; Carol', start: MON });
  var res = generateRecurringTasks(mockSpreadsheet, MON);
  assert.strictEqual(res.created, 8);
  assert.deepStrictEqual(tasksOf('Quiz').slice(0, 4).map(t => [iso(t.due), t.member]),
    [['2026-10-16', 'Alice'], ['2026-10-23', 'Carol'], ['2026-10-30', 'Bob'], ['2026-11-06', 'Alice']]);
  assert.strictEqual(generateRecurringTasks(mockSpreadsheet, MON).created, 0, 'no duplicates');
});

test('rotation: a leave approved later moves the turn and emails both people', () => {
  seed();
  addRecurringRule(mockSpreadsheet, { type: 'Quiz', project: 'ACN', mode: 'Rotation', appliesTo: 'Alice; Bob; Carol', start: MON });
  generateRecurringTasks(mockSpreadsheet, MON);
  leave('Alice', '2026-10-16', '2026-10-16');
  var res = generateRecurringTasks(mockSpreadsheet, MON);
  assert.strictEqual(res.reassigned, 1);
  var first = tasksOf('Quiz')[0];
  assert.notStrictEqual(first.member, 'Alice');
  assert.strictEqual(MailApp.getSentEmails().length, 2);
});

test('each person: twice-a-month task for a role, six half months ahead, departed members left out', () => {
  seed();
  addRecurringRule(mockSpreadsheet, { type: 'Time & expenses', project: 'ACN', mode: 'Each person', appliesTo: 'Role: Consultant', start: MON });
  generateRecurringTasks(mockSpreadsheet, MON);
  var t = tasksOf('Time & expenses');
  assert.strictEqual(t.length, 12);
  assert.deepStrictEqual([...new Set(t.map(x => x.member))].sort(), ['Alice', 'Bob']);
  assert.deepStrictEqual([...new Set(t.map(x => iso(x.due)))].slice(0, 3), ['2026-10-15', '2026-10-30', '2026-11-13']);
  assert.ok(t[0].origin.endsWith('|2026-10-H1'));
  assert.strictEqual(generateRecurringTasks(mockSpreadsheet, MON).created, 0);
});

test('a template schedule changed from the sidebar applies to everyone and moves the coming tasks', () => {
  seed();
  var first = assignTasks({ type: 'Forecast', project: 'ACN', mode: 'Each person', appliesTo: 'Alice', repeat: true, start: MON }, mockSpreadsheet, MON);
  assert.ok(first.success, first.message);
  assert.deepStrictEqual(tasksOf('Forecast').map(t => iso(t.due)), ['2026-10-30', '2026-11-30', '2026-12-31']);
  // Same frequency, other due day: deadlines move
  var moved = assignTasks({ type: 'Forecast', project: 'ACN', mode: 'Each person', appliesTo: 'Bob', repeat: true, start: MON,
    frequency: 'Monthly', due: '20', remind: '2' }, mockSpreadsheet, MON);
  assert.ok(/changed for everyone/.test(moved.message), moved.message);
  assert.deepStrictEqual(tasksOf('Forecast').filter(t => t.member === 'Alice').map(t => iso(t.due)), ['2026-10-20', '2026-11-20', '2026-12-18']);
  assert.deepStrictEqual(findTaskType(mockSpreadsheet, 'Forecast').remind, [2]);
  // Other frequency: coming tasks are cancelled and the new periods planned
  assignTasks({ type: 'Forecast', project: 'ACN', mode: 'Each person', appliesTo: 'Carol', repeat: true, start: MON,
    frequency: 'Twice a month', due: '15; Last day' }, mockSpreadsheet, MON);
  var alice = tasksOf('Forecast').filter(t => t.member === 'Alice');
  assert.strictEqual(alice.length, 6);
  assert.ok(alice.every(t => /-H[12]$/.test(t.origin)));
});

test('calendar: checkboxes; inserting one assigns, ticking marks done, deleting cancels; done count per period', () => {
  seed();
  leave('Bob', '2026-10-23', '2026-10-23');
  renderTaskCalendar(mockSpreadsheet, MON);
  var cal = getTable(mockSpreadsheet, 'taskCalendar');
  cal.getRange(1, 2).setValue('Quiz');
  handleTaskCalendarEdit({ range: cal.getRange(1, 2), source: mockSpreadsheet }, MON);
  assert.strictEqual(cal.getRange(1, 2).getValue(), 'Quiz');
  assert.deepStrictEqual(cal.getRange(5, 1, 3, 1).getValues().map(r => r[0]), ['Alice', 'Bob', 'Carol'], 'departed hidden');
  assert.strictEqual(cal.getRange(2, 2).getValue(), '2026', 'year above the weeks');
  assert.strictEqual(cal.getRange(3, 1).getValue(), 'Member');
  var col = 2 + 4; // current week (4 weeks of history first)
  var edit = (row, c, value, old) => { cal.getRange(row, c).setValue(value); return handleTaskCalendarEdit({ range: cal.getRange(row, c), oldValue: old, source: mockSpreadsheet }, MON); };

  assert.strictEqual(edit(5, col, false, undefined), 'assigned', 'checkbox inserted');
  assert.strictEqual(iso(tasksOf('Quiz')[0].due), '2026-10-16');
  assert.strictEqual(cal.getRange(5, col).getValue(), false);
  assert.strictEqual(cal.getRange(4, col).getValue(), '0/1');

  assert.strictEqual(edit(6, col + 1, false, undefined), 'refused_leave', 'Bob on leave Fri 23 Oct');
  assert.strictEqual(cal.getRange(6, col + 1).getValue(), '');

  assert.strictEqual(edit(5, col, true, 'FALSE'), 'done');
  var t = tasksOf('Quiz')[0];
  assert.strictEqual(t.status, 'Done');
  assert.strictEqual(iso(toDate(t.doneOn)), MON);
  assert.strictEqual(cal.getRange(5, col).getValue(), true);
  assert.strictEqual(cal.getRange(4, col).getValue(), '1/1');

  assert.strictEqual(edit(5, col, false, 'TRUE'), 'reopened');
  assert.strictEqual(tasksOf('Quiz')[0].status, 'Todo');

  assert.strictEqual(edit(5, col, '', 'FALSE'), 'cancelled', 'checkbox deleted');
  assert.strictEqual(tasksOf('Quiz').length, 0);
  assert.strictEqual(cal.getRange(4, col).getValue(), '');
});

test('calendar: a closed month asks first; "No" keeps the old value', () => {
  seed();
  renderTaskCalendar(mockSpreadsheet, MON);
  var cal = getTable(mockSpreadsheet, 'taskCalendar');
  SpreadsheetApp.getUi().alertResponse = 'NO';
  cal.getRange(5, 2).setValue(true); // 4 weeks back = September
  var out = handleTaskCalendarEdit({ range: cal.getRange(5, 2), oldValue: undefined, source: mockSpreadsheet }, MON);
  assert.strictEqual(out, 'kept');
  assert.strictEqual(cal.getRange(5, 2).getValue(), '');
  SpreadsheetApp.getUi().alertResponse = null;
});

test('calendar: a year row above the periods, one label per year, months without their year', () => {
  seed();
  var st = { type: 'Forecast' };
  PropertiesService.getDocumentProperties().setProperty('TASK_CALENDAR_STATE', JSON.stringify(st));
  renderTaskCalendar(mockSpreadsheet, MON);
  var cal = getTable(mockSpreadsheet, 'taskCalendar');
  // Monthly: Jul 2026 to Jun 2027 (3 back, 8 ahead)
  var years = cal.getRange(2, 2, 1, 12).getValues()[0];
  var labels = cal.getRange(3, 2, 1, 12).getValues()[0];
  assert.strictEqual(years[0], '2026');
  assert.strictEqual(years[1], '', 'one label per year');
  assert.strictEqual(years[labels.indexOf('Jan')], '2027');
  assert.strictEqual(labels[0], 'Jul');
});

test('tasks grid: setting Done stamps Done on, reopening clears it', () => {
  seed();
  var made = createTaskRow(mockSpreadsheet, { project: 'ACN', member: rtMembers(mockSpreadsheet)[0], title: 'Forecast', due: toDate('2026-10-20'), type: 'Forecast' });
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

test('reminders follow the template Remind days; the Manager role gets who has not done a mandatory task', () => {
  seed();
  var m = rtMembers(mockSpreadsheet);
  // Time & expenses: Remind "3; 1". Wed 14 Oct: due Thu 15 = 1 working day
  createTaskRow(mockSpreadsheet, { member: m[0], title: 'Time & expenses', due: toDate('2026-10-15'), type: 'Time & expenses' });
  createTaskRow(mockSpreadsheet, { member: m[1], title: 'Time & expenses', due: toDate('2026-10-15'), type: 'Time & expenses' });
  var res = sendTaskReminders(mockSpreadsheet, '2026-10-14');
  assert.deepStrictEqual(res.sent.sort(), ['Alice', 'Bob']);
  assert.ok(MailApp.getSentEmails().every(e => e.subject === 'Tasks due soon'));
  assert.deepStrictEqual(res.digest, []);
  // Fri 16: Alice done, Bob not → overdue to Bob, digest to Carol (Manager)
  setTaskStatus(mockSpreadsheet, readTaskRows(mockSpreadsheet).filter(t => t.member === 'Alice')[0].row, 'Done', '2026-10-15');
  MailApp.clearSentEmails();
  res = sendTaskReminders(mockSpreadsheet, '2026-10-16');
  assert.deepStrictEqual(res.sent, ['Bob']);
  assert.deepStrictEqual(res.digest, ['Carol']);
  var digest = MailApp.getSentEmails().filter(e => e.to === 'carol@accenture.com')[0];
  assert.ok(/Time & expenses \(due 2026-10-15\): Bob/.test(digest.body), digest.body);
});

test('assign sidebar: once for a role, recurring rotation, stop cancels the future', () => {
  seed();
  var data = getAssignSidebarData(mockSpreadsheet);
  assert.deepStrictEqual(data.roles, ['Consultant', 'Manager']);
  var once = assignTasks({ type: 'Peer review', project: 'ACN', mode: 'Each person', appliesTo: 'Role: Consultant', repeat: false, date: MON }, mockSpreadsheet, MON);
  assert.ok(once.success, once.message);
  assert.deepStrictEqual(tasksOf('Peer review').map(t => t.member).sort(), ['Alice', 'Bob']);

  var rot = assignTasks({ type: 'Team presentation', project: 'ACN', mode: 'Rotation', appliesTo: '', start: MON }, mockSpreadsheet, MON);
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

test('projects: the Project columns are a dropdown of names, and a renamed project follows everywhere (ADR 0053)', () => {
  seed();
  assert.deepStrictEqual(schemaValidationsFor(mockSpreadsheet, SCHEMA_SPEC.tasks).project, ['ACN']);
  assert.deepStrictEqual(schemaValidationsFor(mockSpreadsheet, SCHEMA_SPEC.recurringTasks).project, ['ACN']);
  assignTasks({ type: 'Quiz', project: 'ACN', mode: 'Rotation', appliesTo: 'Alice; Bob', start: MON }, mockSpreadsheet, MON);
  checkProjectBackupConflicts(mockSpreadsheet, { asOfDateParam: MON, sendNotifications: false, createTasks: false });
  var projects = getTable(mockSpreadsheet, 'projects');
  projects.getRange(2, 1).setValue('ACN Paris');
  var out = handleProjectNameEdit({ range: projects.getRange(2, 1), oldValue: 'ACN', source: mockSpreadsheet },
    { ss: mockSpreadsheet, sheet: projects, tableKey: 'projects', row: 2, col: 1, colName: 'name' });
  assert.strictEqual(out, 'renamed');
  assert.ok(tasksOf('Quiz').length && tasksOf('Quiz').every(t => t.project === 'ACN Paris'));
  assert.strictEqual(readRecurringRules(mockSpreadsheet)[0].project, 'ACN Paris');
  assert.strictEqual(getTable(mockSpreadsheet, 'backups').getRange(2, 2).getValue(), 'ACN Paris', 'permanent backup row');
  assert.deepStrictEqual(projectNames(mockSpreadsheet), ['ACN Paris']);
});
