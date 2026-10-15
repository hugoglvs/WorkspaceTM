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
global.checkActiveUserStatus = sidebar.checkActiveUserStatus;

const viewBuilder = require('../src/ViewBuilder');
global.renderWeeklyMatrixView = viewBuilder.renderWeeklyMatrixView;
global.refreshWeeklyMatrix = viewBuilder.refreshWeeklyMatrix;
global.handleWeeklyMatrixEdit = viewBuilder.handleWeeklyMatrixEdit;

const { compilePresenceLogs } = require('../src/PresenceEngine');
global.compilePresenceLogs = compilePresenceLogs;
const code = require('../src/Code');
const edits = require('../src/EditHandlers');

// Weekly Presence is created and kept up to date by the actions that change it (ADR 0055)

function seedAlice() {
  mockSpreadsheet.getSheetByName('Members').getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Engineer', 'Core', 'Active', 'Not Logged']
  ]);
}

test('a new tab is created by the member and contract edit handlers, without a toast', () => {
  freshInit(mockSpreadsheet);
  seedAlice();
  assert.ok(!mockSpreadsheet.getSheetByName('Weekly Presence'), 'no tab before the first member action');
  mockSpreadsheet.lastToast = null;

  assert.ok(edits.EDIT_HANDLERS.members.after.includes(edits.editRefreshWeeklyPresence));
  assert.ok(edits.EDIT_HANDLERS.tenures.after.includes(edits.editRefreshWeeklyPresence));

  edits.editRefreshWeeklyPresence({}, { ss: mockSpreadsheet });
  var tab = mockSpreadsheet.getSheetByName('Weekly Presence');
  assert.ok(tab, 'Weekly Presence exists');
  assert.strictEqual(tab.getRange(1, 3).getValue(), 'Alice Smith');
  assert.strictEqual(mockSpreadsheet.lastToast, null, 'quiet refresh');
});

test('saving a routine compiles the log and creates Weekly Presence', () => {
  freshInit(mockSpreadsheet);
  seedAlice();
  var res = sidebar.saveMemberBaselines({
    member: 'Alice Smith',
    startDate: '2026-06-01',
    days: { Monday: { location_am: 'Remote', location_pm: 'Remote' } }
  });
  assert.strictEqual(res.success, true);
  assert.ok(mockSpreadsheet.getSheetByName('Weekly Presence'), 'Weekly Presence exists');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');
  assert.ok(logSheet.getLastRow() > 1, 'the daily log was compiled');
});

test('submitting a leave creates Weekly Presence, even before it is approved', () => {
  freshInit(mockSpreadsheet);
  seedAlice();
  var res = sidebar.submitMemberLeave({
    member: 'Alice Smith', startDate: '2026-06-01', endDate: '2026-06-03',
    leaveType: 'Sick Leave', status: 'Requested'
  }, mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.ok(mockSpreadsheet.getSheetByName('Weekly Presence'), 'Weekly Presence exists');
});

test('the Project Management menu no longer offers a manual Weekly Presence refresh', () => {
  code.onOpen();
  var labels = SpreadsheetApp.getUi().menu.items.map(function(i) { return String(i.caption || ''); }).join('|');
  assert.ok(!/Weekly Matrix/i.test(labels));
});
