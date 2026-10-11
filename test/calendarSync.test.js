const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals } = require('./mockGas');

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

const { freshInit } = require('../src/Init');
const sidebar = require('../src/Sidebar');
global.extractCellValue = sidebar.extractCellValue;
global.extractEmailFromChip = sidebar.extractEmailFromChip;

const calendarSyncService = require('../src/CalendarSyncService');
global.CalendarSyncService = calendarSyncService;
global.syncGoogleCalendar = calendarSyncService.syncGoogleCalendar;
global.syncCalendar = calendarSyncService.syncCalendar;
global.handleEventEdit = calendarSyncService.handleEventEdit;

const code = require('../src/Code');

test('getCalendarId & setCalendarId: manages Google Calendar ID via Document Properties', () => {
  var initialId = calendarSyncService.getCalendarId(mockSpreadsheet);
  assert.strictEqual(initialId, 'primary');

  calendarSyncService.setCalendarId('custom-team-calendar@group.calendar.google.com');
  var updatedId = calendarSyncService.getCalendarId(mockSpreadsheet);
  assert.strictEqual(updatedId, 'custom-team-calendar@group.calendar.google.com');
});

test('generateEventId: produces valid EVT-YYMMDD-XXXX format', () => {
  var id = calendarSyncService.generateEventId();
  assert.ok(id.startsWith('EVT-'));
  assert.match(id, /^EVT-\d{6}-\d{4}$/);
});

test('parseEventDate & formatEventDateIso: correctly handles dates and datetimes', () => {
  var parsedDate = calendarSyncService.parseEventDate('2026-09-15');
  assert.ok(parsedDate instanceof Date);
  assert.strictEqual(calendarSyncService.formatEventDateIso(parsedDate), '2026-09-15');

  var parsedDateTime = calendarSyncService.parseEventDate('2026-10-15 14:30');
  assert.ok(parsedDateTime instanceof Date);
  assert.strictEqual(parsedDateTime.getFullYear(), 2026);
  assert.strictEqual(parsedDateTime.getMonth(), 9);
  assert.strictEqual(parsedDateTime.getDate(), 15);
  assert.strictEqual(parsedDateTime.getHours(), 14);
  assert.strictEqual(parsedDateTime.getMinutes(), 30);

  var dObj = new Date(2026, 8, 20);
  assert.strictEqual(calendarSyncService.formatEventDateIso(dObj), '2026-09-20');
});

test('parseEventDuration: parses numbers, strings, hours, and handles null for whole day', () => {
  assert.strictEqual(calendarSyncService.parseEventDuration(null), null);
  assert.strictEqual(calendarSyncService.parseEventDuration(''), null);
  assert.strictEqual(calendarSyncService.parseEventDuration(60), 60);
  assert.strictEqual(calendarSyncService.parseEventDuration('60'), 60);
  assert.strictEqual(calendarSyncService.parseEventDuration('90m'), 90);
  assert.strictEqual(calendarSyncService.parseEventDuration('45 mins'), 45);
  assert.strictEqual(calendarSyncService.parseEventDuration('1h'), 60);
  assert.strictEqual(calendarSyncService.parseEventDuration('1.5h'), 90);
  assert.strictEqual(calendarSyncService.parseEventDuration('2 hours'), 120);
});

test('getActiveTeamMemberEmails: retrieves active members and filters departed members', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  // Columns: member, accenture_email, phone_number, current_role, current_department, tenure_status, today_presence
  membersSheet.getRange(2, 1, 3, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Senior Eng', 'Core', 'Active', 'Main Office'],
    ['Bob Jones', 'bob@accenture.com', '', 'Designer', 'UX', 'Active', 'Remote'],
    ['Charlie Brown', 'charlie@accenture.com', '', 'Analyst', 'Ops', 'Departed', 'Off']
  ]);

  var emails = calendarSyncService.getActiveTeamMemberEmails(mockSpreadsheet);
  assert.strictEqual(emails.length, 2);
  assert.ok(emails.includes('alice@accenture.com'));
  assert.ok(emails.includes('bob@accenture.com'));
  assert.strictEqual(emails.includes('charlie@accenture.com'), false, 'Departed members should not be invited');
});

test('syncGoogleCalendar: creates whole day event when duration is null and does not auto-invite', () => {
  freshInit(mockSpreadsheet);

  var eventsSheet = mockSpreadsheet.getSheetByName('tbl_events');
  // Columns: event_id, title, date, duration, location, organizer, google_calendar_event_id
  eventsSheet.getRange(2, 1, 1, 7).setValues([
    ['EVT-001', 'Q3 Team Social', '2026-09-25', '', 'Rooftop Lounge', 'alice@accenture.com', '']
  ]);

  var res = calendarSyncService.syncGoogleCalendar(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.createdCount, 1);

  var row2 = eventsSheet.getRange(2, 1, 1, 7).getValues()[0];
  var gEvtId = row2[6];
  assert.ok(gEvtId, 'Row 2 should have google_calendar_event_id in column 7');

  // Verify in organizer calendar
  var aliceCal = CalendarApp.getCalendarById('alice@accenture.com');
  var evt = aliceCal.getEventById(gEvtId);
  assert.ok(evt, 'Event must exist in organizer calendar');
  assert.strictEqual(evt.getTitle(), 'Q3 Team Social');
  assert.strictEqual(evt.getLocation(), 'Rooftop Lounge');
  assert.strictEqual(evt.isAllDayEvent(), true, 'Event with null duration must be all-day (whole day)');
  assert.strictEqual(evt.getGuestList().length, 0, 'Organizer invites members directly from Google Calendar; not auto-invited');
});

test('syncGoogleCalendar: creates timed event when duration is provided', () => {
  freshInit(mockSpreadsheet);

  var eventsSheet = mockSpreadsheet.getSheetByName('tbl_events');
  // Columns: event_id, title, date, duration, location, organizer, google_calendar_event_id
  eventsSheet.getRange(2, 1, 1, 7).setValues([
    ['EVT-002', 'Sprint Planning', '2026-10-15 14:00', 60, 'Room 101', 'bob@accenture.com', '']
  ]);

  var res = calendarSyncService.syncGoogleCalendar(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.createdCount, 1);

  var row2 = eventsSheet.getRange(2, 1, 1, 7).getValues()[0];
  var gEvtId = row2[6];
  assert.ok(gEvtId);

  var bobCal = CalendarApp.getCalendarById('bob@accenture.com');
  var evt = bobCal.getEventById(gEvtId);
  assert.ok(evt);
  assert.strictEqual(evt.getTitle(), 'Sprint Planning');
  assert.strictEqual(evt.isAllDayEvent(), false, 'Event with duration takes a time slot in the agenda');
  assert.strictEqual(evt.getStartTime().getHours(), 14);
  assert.strictEqual(evt.getEndTime().getHours(), 15);
});

test('syncGoogleCalendar: updates existing events when title, date, duration, or location is edited in sheet', () => {
  freshInit(mockSpreadsheet);

  var eventsSheet = mockSpreadsheet.getSheetByName('tbl_events');
  eventsSheet.getRange(2, 1, 1, 7).setValues([
    ['EVT-100', 'Initial Title', '2026-09-10', '', 'Room A', 'alice@accenture.com', '']
  ]);

  // Initial push: whole day event
  calendarSyncService.syncGoogleCalendar(mockSpreadsheet);
  var gId = eventsSheet.getRange(2, 7).getValue();
  assert.ok(gId);

  var aliceCal = CalendarApp.getCalendarById('alice@accenture.com');
  var evt = aliceCal.getEventById(gId);
  assert.strictEqual(evt.isAllDayEvent(), true);

  // User edits title, changes date to datetime with 90 min duration, and changes location
  eventsSheet.getRange(2, 2).setValue('Updated Architecture Review');
  eventsSheet.getRange(2, 3).setValue('2026-09-12 10:00');
  eventsSheet.getRange(2, 4).setValue(90);
  eventsSheet.getRange(2, 5).setValue('Room B');

  // Sync again
  var updateRes = calendarSyncService.syncGoogleCalendar(mockSpreadsheet);
  assert.strictEqual(updateRes.success, true);
  assert.strictEqual(updateRes.updatedCount, 1);
  assert.strictEqual(updateRes.createdCount, 0);

  var updatedEvt = aliceCal.getEventById(gId);
  assert.strictEqual(updatedEvt.getTitle(), 'Updated Architecture Review');
  assert.strictEqual(updatedEvt.getLocation(), 'Room B');
  assert.strictEqual(updatedEvt.isAllDayEvent(), false, 'Event should be updated to timed event');
  assert.strictEqual(updatedEvt.getStartTime().getHours(), 10);
  assert.strictEqual(updatedEvt.getEndTime().getHours(), 11);
  assert.strictEqual(updatedEvt.getEndTime().getMinutes(), 30);
});

test('syncCalendar: public entry point toasts sync summary', () => {
  freshInit(mockSpreadsheet);

  var eventsSheet = mockSpreadsheet.getSheetByName('tbl_events');
  eventsSheet.getRange(2, 1, 1, 7).setValues([
    ['EVT-200', 'Annual Gala', '2026-12-18', '', 'Grand Ballroom', '', '']
  ]);

  var res = code.syncCalendarEvents(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.createdCount, 1);
  assert.ok(mockSpreadsheet.lastToast);
  assert.ok(mockSpreadsheet.lastToast.title.includes('Calendar Synchronized'));
});

test('handleEventEdit: auto-generates event_id and creates Calendar event on live edit', () => {
  freshInit(mockSpreadsheet);

  var eventsSheet = mockSpreadsheet.getSheetByName('tbl_events');
  // User enters title and date on row 2
  eventsSheet.getRange(2, 2).setValue('Hackathon 2026');
  var dateRange = eventsSheet.getRange(2, 3);
  dateRange.setValue('2026-11-20');

  // Trigger onEdit
  code.onEdit({ range: dateRange, value: '2026-11-20' });

  var row = eventsSheet.getRange(2, 1, 1, 7).getValues()[0];
  var eventId = row[0];
  var title = row[1];
  var gId = row[6];

  assert.ok(eventId.startsWith('EVT-'), 'Event ID must be auto-generated');
  assert.strictEqual(title, 'Hackathon 2026');
  assert.ok(gId, 'Google Calendar event ID must be created automatically in column 7');

  var cal = CalendarApp.getDefaultCalendar();
  var evt = cal.getEventById(gId);
  assert.strictEqual(evt.getTitle(), 'Hackathon 2026');
  assert.strictEqual(evt.isAllDayEvent(), true, 'Default duration null is all-day event');
});

test('handleEventEdit: updates existing Calendar event on live title and duration edit', () => {
  freshInit(mockSpreadsheet);

  var eventsSheet = mockSpreadsheet.getSheetByName('tbl_events');
  // Seed synced event
  eventsSheet.getRange(2, 1, 1, 7).setValues([
    ['EVT-300', 'Sprint Review', '2026-09-30 15:00', 45, 'Online', 'charlie@accenture.com', 'cal_sprint_300']
  ]);

  var charlieCal = CalendarApp.getCalendarById('charlie@accenture.com');
  var mockEvt = charlieCal.createEvent('Sprint Review', new Date(2026, 8, 30, 15, 0), new Date(2026, 8, 30, 15, 45));
  mockEvt.id = 'cal_sprint_300';
  charlieCal.getEventById = function(id) {
    if (id === 'cal_sprint_300') return mockEvt;
    return null;
  };

  // User edits title
  var titleRange = eventsSheet.getRange(2, 2);
  titleRange.setValue('Sprint 15 Review & Demo');
  code.onEdit({ range: titleRange });

  assert.strictEqual(mockEvt.getTitle(), 'Sprint 15 Review & Demo', 'Calendar event title must be updated live');
});

test('onOpen: includes Sync Calendar Events in custom menu', () => {
  code.onOpen();
  var ui = SpreadsheetApp.getUi();
  assert.ok(ui.menu);
  var captions = ui.menu.items.map(function(item) { return item.caption; });
  assert.ok(captions.includes('Sync Calendar Events'));
});
