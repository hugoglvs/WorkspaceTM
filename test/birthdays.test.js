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
Object.assign(global, require('../src/Birthdays'));

function seed(rows) {
  freshInit(mockSpreadsheet);
  MailApp.clearSentEmails();
  getTable(mockSpreadsheet, 'members').getRange(2, 1, rows.length, 8).setValues(rows);
}
const TEAM = [
  ['Alice', 'alice@accenture.com', '', 'Dev', 'Eng', 'Active', '', '14 Oct'],
  ['Bob', 'bob@accenture.com', '', 'Dev', 'Eng', 'Active', '', ''],
  ['Carol', 'carol@accenture.com', '', 'Dev', 'Eng', 'Active', '', '18 Oct'],  // Saturday 2026
  ['Dan', 'dan@accenture.com', '', 'Dev', 'Eng', 'Departed', '', '13 Oct']
];

test('Members has a Birthday column at the end', () => {
  freshInit(mockSpreadsheet);
  var header = getTable(mockSpreadsheet, 'members').getRange(1, 1, 1, 8).getValues()[0];
  assert.strictEqual(header[7], 'Birthday');
  assert.strictEqual(header[6], 'Today', 'existing columns did not move');
});

test('parseBirthday reads common formats, drops the year, rejects nonsense', () => {
  const f = v => formatBirthday(parseBirthday(v));
  assert.strictEqual(f('14 Mar'), '14 Mar');
  assert.strictEqual(f('14 march'), '14 Mar');
  assert.strictEqual(f('March 14'), '14 Mar');
  assert.strictEqual(f('14/03'), '14 Mar');
  assert.strictEqual(f('14.03.1990'), '14 Mar');
  assert.strictEqual(f('1990-03-14'), '14 Mar');
  assert.strictEqual(f(new Date(1990, 2, 14)), '14 Mar');
  assert.strictEqual(f('29 Feb'), '29 Feb');
  assert.strictEqual(parseBirthday('31/02'), null);
  assert.strictEqual(parseBirthday('soon'), null);
});

test('Birthday is a full date: a date is kept, text with a year becomes a date, the rest gets a toast', () => {
  seed(TEAM);
  var sheet = getTable(mockSpreadsheet, 'members');
  var cell = sheet.getRange(3, 8);
  var typed = new Date(1990, 6, 3);
  cell.setValue(typed);
  code.onEdit({ range: cell, value: typed, source: mockSpreadsheet });
  assert.strictEqual(cell.getValue(), typed);
  cell.setValue('1988-07-03');
  code.onEdit({ range: cell, value: '1988-07-03', source: mockSpreadsheet });
  assert.strictEqual(cell.getValue().getTime(), new Date(1988, 6, 3).getTime());
  cell.setValue('3 Jul');
  code.onEdit({ range: cell, value: '3 Jul', source: mockSpreadsheet });
  assert.strictEqual(mockSpreadsheet.lastToast.title, 'Birthday');
  assert.match(mockSpreadsheet.lastToast.msg, /full birth date/);
  assert.strictEqual(parseBirthDate('29/02/2001'), null, 'no 29 Feb in 2001');
  assert.strictEqual(parseBirthDate('14/03'), null, 'no year');
});

test('announcement days: the next day, and Saturday to Monday on a Friday, nothing at weekends', () => {
  const iso = d => d.toISOString ? `${d.getFullYear()}-${d.getMonth() + 1}-${d.getDate()}` : d;
  assert.deepStrictEqual(birthdayAnnouncementDays('2026-10-13').map(iso), ['2026-10-14']);
  assert.deepStrictEqual(birthdayAnnouncementDays('2026-10-16').map(iso), ['2026-10-17', '2026-10-18', '2026-10-19']);
  assert.deepStrictEqual(birthdayAnnouncementDays('2026-10-17'), []);
});

test('reminder: every other Active member is emailed the day before, with a yearly calendar link', () => {
  seed(TEAM);
  var res = sendBirthdayReminders(mockSpreadsheet, '2026-10-13');
  assert.deepStrictEqual(res.celebrants, ['Alice']);
  assert.deepStrictEqual(res.sent.sort(), ['Bob', 'Carol'], 'not Alice herself, not departed Dan');
  var mail = MailApp.getSentEmails()[0];
  assert.ok(mail.subject.includes("Alice's birthday is tomorrow"));
  assert.ok(mail.htmlBody.includes('action=TEMPLATE'));
  assert.ok(mail.htmlBody.includes('dates=20261014/20261015'));
  assert.ok(mail.htmlBody.includes(encodeURIComponent('RRULE:FREQ=YEARLY')));
});

test('reminder: a Saturday birthday is announced on Friday; departed members are not celebrated', () => {
  seed(TEAM);
  assert.deepStrictEqual(sendBirthdayReminders(mockSpreadsheet, '2026-10-16').celebrants, ['Carol']);
  MailApp.clearSentEmails();
  assert.deepStrictEqual(sendBirthdayReminders(mockSpreadsheet, '2026-10-12').celebrants, [], 'Dan (13 Oct) has left');
});

test('reminder: sent once a day even if the job runs twice', () => {
  seed(TEAM);
  sendBirthdayReminders(mockSpreadsheet, '2026-10-13');
  var again = sendBirthdayReminders(mockSpreadsheet, '2026-10-13');
  assert.strictEqual(again.skipped, 'already_sent_today');
  assert.strictEqual(MailApp.getSentEmails().length, 2);
});

test('29 Feb is celebrated on 28 Feb in other years', () => {
  seed([['Lea', 'lea@x.com', '', 'Dev', 'Eng', 'Active', '', '29 Feb'], ['Bob', 'bob@x.com', '', 'Dev', 'Eng', 'Active', '', '']]);
  assert.deepStrictEqual(sendBirthdayReminders(mockSpreadsheet, '2027-02-25').celebrants, [], 'Thursday announces Friday 26 only');
  assert.deepStrictEqual(sendBirthdayReminders(mockSpreadsheet, '2027-02-26').celebrants, ['Lea'], 'Friday announces Sat 27 to Mon 1 Mar, so Sunday 28');
  MailApp.clearSentEmails();
  seed([['Lea', 'lea@x.com', '', 'Dev', 'Eng', 'Active', '', '29 Feb'], ['Bob', 'bob@x.com', '', 'Dev', 'Eng', 'Active', '', '']]);
  assert.deepStrictEqual(sendBirthdayReminders(mockSpreadsheet, '2026-02-27').celebrants, ['Lea'], 'Friday 27 Feb 2026 announces Sat 28');
});

test('Team Birthdays calendar: one yearly event per active birthday, shared with active members, kept in step', () => {
  seed(TEAM);
  var res = syncBirthdayCalendar(mockSpreadsheet, '2026-10-12');
  assert.strictEqual(res.created, 2);
  var calId = PropertiesService.getDocumentProperties().getProperty('BIRTHDAY_CALENDAR_ID');
  var cal = CalendarApp.getCalendarById(calId);
  var titles = cal.getEvents().map(e => e.getTitle()).sort();
  assert.deepStrictEqual(titles, ["🎂 Alice's birthday", "🎂 Carol's birthday"]);
  assert.deepStrictEqual(cal.getEvents()[0].recurrence.rules, ['YEARLY']);
  var acl = Calendar.Acl.__entries(calId);
  assert.deepStrictEqual(Object.keys(acl).sort(), ['alice@accenture.com', 'bob@accenture.com', 'carol@accenture.com']);
  assert.strictEqual(acl['bob@accenture.com'].args.sendNotifications, false);

  // Alice's birthday changes, Carol leaves
  getTable(mockSpreadsheet, 'members').getRange(2, 8).setValue('15 Oct');
  getTable(mockSpreadsheet, 'members').getRange(4, 6).setValue('Departed');
  res = syncBirthdayCalendar(mockSpreadsheet, '2026-10-12');
  assert.deepStrictEqual([res.created, res.removed, res.unshared], [1, 2, 1]);
  assert.deepStrictEqual(cal.getEvents().map(e => e.getTitle()), ["🎂 Alice's birthday"]);
  assert.strictEqual(cal.getEvents()[0].getAllDayStartDate().getDate(), 15);

  // A second run changes nothing
  res = syncBirthdayCalendar(mockSpreadsheet, '2026-10-12');
  assert.deepStrictEqual([res.created, res.removed, res.shared, res.unshared], [0, 0, 0, 0]);
});

test('the reminder email links to the Team Birthdays calendar once it exists', () => {
  seed(TEAM);
  syncBirthdayCalendar(mockSpreadsheet, '2026-10-12');
  sendBirthdayReminders(mockSpreadsheet, '2026-10-13');
  assert.ok(MailApp.getSentEmails()[0].htmlBody.includes('calendar/r?cid='));
});

