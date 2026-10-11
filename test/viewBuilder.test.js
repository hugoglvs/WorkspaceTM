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

test('getIsoWeekNumber: correctly computes ISO week number', () => {
  assert.strictEqual(viewBuilder.getIsoWeekNumber(new Date(2026, 0, 1)), 1);
  assert.strictEqual(viewBuilder.getIsoWeekNumber(new Date(2026, 0, 5)), 2);
  assert.strictEqual(viewBuilder.formatIsoWeekString(new Date(2026, 0, 5)), 'W02');
});

test('renderWeeklyMatrixView: should build view_weekly_matrix with merged People Chip headers, weekends, and merged week numbers', () => {
  freshInit(mockSpreadsheet);

  // Seed tbl_members with 2 active members (Row 2 Alice, Row 3 Bob)
  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 2, 7).setValues([
    ['Alice Smith', '', '', 'Engineer', 'Core', 'Active', 'Main Office'],
    ['Bob Jones', '', '', 'Designer', 'UX', 'Active', 'Remote']
  ]);

  // Seed tbl_presence_log with sample records
  var logSheet = mockSpreadsheet.getSheetByName('tbl_presence_log');
  logSheet.getRange(2, 1, 3, 5).setValues([
    ['log-1', '2026-01-05', 'Alice Smith', 'Main Office', 'Main Office'],
    ['log-2', '2026-01-05', 'Bob Jones', 'Remote', 'Remote'],
    ['log-3', '2026-01-06', 'Alice Smith', 'Remote', 'Main Office']
  ]);

  // Render 1 week: Monday 2026-01-05 to Sunday 2026-01-11 (7 days)
  var res = viewBuilder.renderWeeklyMatrixView('2026-01-05', '2026-01-11', mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.sheetName, 'view_weekly_matrix');
  assert.strictEqual(res.rowsRendered, 7); // Mon to Sun = 7 days
  assert.strictEqual(res.membersCount, 2);

  var matrixSheet = mockSpreadsheet.getSheetByName('view_weekly_matrix');
  assert.ok(matrixSheet);
  assert.strictEqual(matrixSheet.getFrozenRows(), 2); // 2 frozen header rows

  // Check Row 1 (Merged member headers)
  var row1Vals = matrixSheet.getRange(1, 1, 1, 6).getValues()[0];
  assert.strictEqual(row1Vals[0], 'Week Number');
  assert.strictEqual(row1Vals[1], 'Date');
  assert.strictEqual(row1Vals[2], 'Alice Smith');
  assert.strictEqual(row1Vals[4], 'Bob Jones');
  assert.strictEqual(matrixSheet.getRange(1, 3, 1, 2).merged, true, 'Alice header should be merged across AM and PM');
  assert.strictEqual(matrixSheet.getRange(1, 5, 1, 2).merged, true, 'Bob header should be merged across AM and PM');

  // Check Row 2 (AM and PM subheaders)
  var row2Vals = matrixSheet.getRange(2, 1, 1, 6).getValues()[0];
  assert.strictEqual(row2Vals[2], 'AM');
  assert.strictEqual(row2Vals[3], 'PM');
  assert.strictEqual(row2Vals[4], 'AM');
  assert.strictEqual(row2Vals[5], 'PM');

  // Check Data Row 1 (Monday 2026-01-05, at sheet row 3)
  var dataRow1 = matrixSheet.getRange(3, 1, 1, 6).getValues()[0];
  assert.strictEqual(dataRow1[0], 'W02');
  assert.strictEqual(dataRow1[1], '2026-01-05');
  assert.strictEqual(dataRow1[2], 'Main Office');
  assert.strictEqual(dataRow1[3], 'Main Office');
  assert.strictEqual(dataRow1[4], 'Remote');
  assert.strictEqual(dataRow1[5], 'Remote');

  // Check Weekend styling: Saturday 2026-01-10 is at sheet row 8
  var satRange = matrixSheet.getRange(8, 1, 1, 6);
  assert.strictEqual(satRange.background, '#5F6368', 'Weekend row should have dark gray background');
  assert.strictEqual(satRange.fontColor, '#FFFFFF', 'Weekend row text should be white');

  // Check conditional formatting rules exist
  var rules = matrixSheet.getConditionalFormatRules();
  assert.ok(rules.length >= 3, 'Should have conditional formatting rules for Remote, Office, and Leaves');

  // Check Column widths
  assert.strictEqual(matrixSheet.columnWidths[1], 110);
  assert.strictEqual(matrixSheet.columnWidths[2], 120);
  assert.strictEqual(matrixSheet.columnWidths[3], 110);
});

test('renderWeeklyMatrixView: should hide prior weeks and only show last 4 weeks', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', '', '', 'Engineer', 'Core', 'Active', 'Main Office']
  ]);

  // 6 weeks of days (42 days) from 2026-01-05 (W02) to 2026-02-15 (W07)
  var res = viewBuilder.renderWeeklyMatrixView('2026-01-05', '2026-02-15', mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.totalWeeks, 6);

  var matrixSheet = mockSpreadsheet.getSheetByName('view_weekly_matrix');
  // First 2 weeks (14 days, rows 3 to 16) should be hidden
  assert.strictEqual(matrixSheet.isRowHiddenByUser(3), true, 'Row 3 should be hidden');
  assert.strictEqual(matrixSheet.isRowHiddenByUser(16), true, 'Row 16 should be hidden');
  // Row 17 (start of last 4 weeks) should NOT be hidden
  assert.strictEqual(matrixSheet.isRowHiddenByUser(17), false, 'Row 17 should be visible');
});

test('handleWeeklyMatrixEdit: should sync in-sheet presence declaration to tbl_presence_log and tbl_members', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', '', '', 'Engineer', 'Core', 'Active', 'Not Logged']
  ]);

  // Render matrix view
  viewBuilder.renderWeeklyMatrixView('2026-01-05', '2026-01-11', mockSpreadsheet);
  var matrixSheet = mockSpreadsheet.getSheetByName('view_weekly_matrix');

  // User edits cell C3 (Row 3, Col 3 = Monday 2026-01-05 Alice AM) to "Remote"
  var editRange = matrixSheet.getRange(3, 3, 1, 1);
  editRange.setValue('Remote');

  var mockEvent = {
    range: editRange
  };

  // Trigger onEdit
  code.onEdit(mockEvent);

  // Check tbl_presence_log
  var logSheet = mockSpreadsheet.getSheetByName('tbl_presence_log');
  var logRows = logSheet.getRange(2, 1, 1, 6).getValues();
  assert.strictEqual(logRows.length, 1);
  assert.strictEqual(logRows[0][1], '2026-01-05');
  assert.strictEqual(logRows[0][2], 'Alice Smith');
  assert.strictEqual(logRows[0][3], 'Remote'); // AM location set!
  assert.strictEqual(logRows[0][4], ''); // PM empty
  assert.match(String(logRows[0][5]), /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/, 'Timestamp should be formatted as YYYY-MM-DD HH:mm:ss');

  // Now user edits cell D3 (Row 3, Col 4 = Monday 2026-01-05 Alice PM) to "Main Office"
  var editRangePm = matrixSheet.getRange(3, 4, 1, 1);
  editRangePm.setValue('Main Office');
  code.onEdit({ range: editRangePm });

  var updatedLog = logSheet.getRange(2, 1, 1, 6).getValues();
  assert.strictEqual(updatedLog[0][3], 'Remote');
  assert.strictEqual(updatedLog[0][4], 'Main Office');
  assert.match(String(updatedLog[0][5]), /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/, 'Updated timestamp should also be YYYY-MM-DD HH:mm:ss');
});

test('compilePresence: should automatically trigger weekly matrix refresh', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', '', '', 'Engineer', 'Core', 'Active', 'Main Office']
  ]);

  var baselinesSheet = mockSpreadsheet.getSheetByName('tbl_presence_baselines');
  baselinesSheet.getRange(2, 1, 1, 7).setValues([
    ['b-1', 'Alice Smith', '2026-01-05', '2026-01-09', 'Monday', 'Main Office', 'Remote']
  ]);

  // Public compilePresence entry point
  code.compilePresence();

  var matrixSheet = mockSpreadsheet.getSheetByName('view_weekly_matrix');
  assert.ok(matrixSheet, 'compilePresence should create view_weekly_matrix');
  var headersRow1 = matrixSheet.getRange(1, 1, 1, 4).getValues()[0];
  assert.strictEqual(headersRow1[0], 'Week Number');
  assert.strictEqual(headersRow1[1], 'Date');
  assert.strictEqual(headersRow1[2], 'Alice Smith');

  var headersRow2 = matrixSheet.getRange(2, 1, 1, 4).getValues()[0];
  assert.strictEqual(headersRow2[2], 'AM');
  assert.strictEqual(headersRow2[3], 'PM');
});

test('refreshWeeklyMatrix: should toast update and refresh view', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Charlie Brown', '', '', 'Lead', 'Delivery', 'Active', 'Main Office']
  ]);

  var res = code.refreshWeeklyMatrix(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.membersCount, 1);
  assert.ok(mockSpreadsheet.lastToast);
  assert.ok(mockSpreadsheet.lastToast.title.includes('Weekly Matrix Updated'));
});

test('handleWeeklyMatrixEdit: keeps the edit and suggests the sidebar for routines and leaves', () => {
  freshInit(mockSpreadsheet);
  mockSpreadsheet.getSheetByName('tbl_members').getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', '', '', 'Engineer', 'Core', 'Active', 'Not Logged']
  ]);
  viewBuilder.renderWeeklyMatrixView('2026-01-05', '2026-01-11', mockSpreadsheet);
  var editRange = mockSpreadsheet.getSheetByName('view_weekly_matrix').getRange(3, 3, 1, 1);
  editRange.setValue('Remote');

  code.onEdit({ range: editRange });

  assert.strictEqual(mockSpreadsheet.getSheetByName('tbl_presence_log').getRange(2, 4).getValue(), 'Remote', 'edit saved');
  assert.ok(mockSpreadsheet.lastToast, 'a toast is shown');
  assert.ok(mockSpreadsheet.lastToast.msg.includes('Declare Presence & Leaves'), 'it points to the sidebar');
});
