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
global.renderActiveResponsiblesDashboard = viewBuilder.renderActiveResponsiblesDashboard;
global.refreshActiveResponsibles = viewBuilder.refreshActiveResponsibles;

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

test('renderActiveResponsiblesDashboard: should accurately compute Lead Active, Covered by Backup, and CRITICAL Uncovered statuses', () => {
  freshInit(mockSpreadsheet);

  // Setup members in tbl_members
  // Row 2: Alice Smith (Lead on Proj 1, active in office)
  // Row 3: Bob Jones (Lead on Proj 2, on vacation today; Backup on Proj 1)
  // Row 4: Charlie Brown (Backup on Proj 2, active remote)
  // Row 5: Dave Wilson (Lead on Proj 3, on sick leave)
  // Row 6: Eve Adams (Backup on Proj 3, departed tenure)
  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 5, 7).setValues([
    ['Alice Smith', '', '', 'Senior Eng', 'Core', 'Active', 'Main Office'],
    ['Bob Jones', '', '', 'Lead Eng', 'Core', 'Active', 'Main Office'],
    ['Charlie Brown', '', '', 'Architect', 'Core', 'Active', 'Remote'],
    ['Dave Wilson', '', '', 'PM', 'Delivery', 'Active', 'Main Office'],
    ['Eve Adams', '', '', 'Dev', 'Delivery', 'Departed', 'Main Office']
  ]);

  // Setup tbl_projects
  // Proj 1: Lead = Alice Smith, Backup = Bob Jones
  // Proj 2: Lead = Bob Jones, Backup = Charlie Brown
  // Proj 3: Lead = Dave Wilson, Backup = Eve Adams
  var projectsSheet = mockSpreadsheet.getSheetByName('tbl_projects');
  projectsSheet.getRange(2, 1, 3, 7).setValues([
    ['PRJ-001', 'Core Platform Revamp', 'Active', '2026-01-01', '2026-06-30', 'Alice Smith', 'Bob Jones'],
    ['PRJ-002', 'Mobile Redesign', 'Active', '2026-02-01', '2026-08-31', 'Bob Jones', 'Charlie Brown'],
    ['PRJ-003', 'Legacy Migration', 'Planned', '2026-03-01', '2026-12-31', 'Dave Wilson', 'Eve Adams']
  ]);

  // Setup tbl_leaves: Bob on Vacation today (2026-01-05), Dave on Sick Leave
  var leavesSheet = mockSpreadsheet.getSheetByName('tbl_leaves');
  leavesSheet.getRange(2, 1, 2, 6).setValues([
    ['LV-001', 'Bob Jones', '2026-01-05', '2026-01-09', 'Vacation', 'Approved'],
    ['LV-002', 'Dave Wilson', '2026-01-05', '2026-01-06', 'Sick Leave', 'Approved']
  ]);

  // Render dashboard as of 2026-01-05
  var res = viewBuilder.renderActiveResponsiblesDashboard(mockSpreadsheet, '2026-01-05');
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.projectsCount, 3);
  assert.strictEqual(res.leadActiveCount, 1, 'PRJ-001 should be Lead Active');
  assert.strictEqual(res.coveredByBackupCount, 1, 'PRJ-002 should be Covered by Backup');
  assert.strictEqual(res.uncoveredCount, 1, 'PRJ-003 should be CRITICAL: Uncovered');

  var dashboardSheet = mockSpreadsheet.getSheetByName('view_active_responsibles');
  assert.ok(dashboardSheet);
  assert.strictEqual(dashboardSheet.getFrozenRows(), 1);

  // Check Headers
  var headers = dashboardSheet.getRange(1, 1, 1, 9).getValues()[0];
  assert.strictEqual(headers[0], 'Project ID');
  assert.strictEqual(headers[1], 'Project Name');
  assert.strictEqual(headers[2], 'Status');
  assert.strictEqual(headers[3], 'Project Lead');
  assert.strictEqual(headers[4], 'Lead Presence');
  assert.strictEqual(headers[5], 'Project Backup');
  assert.strictEqual(headers[6], 'Backup Presence');
  assert.strictEqual(headers[7], 'Active Responsible');
  assert.strictEqual(headers[8], 'Coverage Status');

  // Check Row 2 (PRJ-001): Alice active -> Lead Active
  var row1 = dashboardSheet.getRange(2, 1, 1, 9).getValues()[0];
  assert.strictEqual(row1[0], 'PRJ-001');
  assert.strictEqual(row1[3], 'Alice Smith');
  assert.strictEqual(row1[4], 'Main Office');
  assert.strictEqual(row1[5], 'Bob Jones');
  assert.strictEqual(row1[6], 'Vacation');
  assert.strictEqual(row1[7], 'Alice Smith');
  assert.strictEqual(row1[8], 'Lead Active');

  // Check Row 3 (PRJ-002): Bob on vacation, Charlie active -> Covered by Backup
  var row2 = dashboardSheet.getRange(3, 1, 1, 9).getValues()[0];
  assert.strictEqual(row2[0], 'PRJ-002');
  assert.strictEqual(row2[3], 'Bob Jones');
  assert.strictEqual(row2[4], 'Vacation');
  assert.strictEqual(row2[5], 'Charlie Brown');
  assert.strictEqual(row2[6], 'Remote');
  assert.strictEqual(row2[7], 'Charlie Brown');
  assert.strictEqual(row2[8], 'Covered by Backup');

  // Check Row 4 (PRJ-003): Dave sick, Eve departed -> CRITICAL: Uncovered
  var row3 = dashboardSheet.getRange(4, 1, 1, 9).getValues()[0];
  assert.strictEqual(row3[0], 'PRJ-003');
  assert.strictEqual(row3[3], 'Dave Wilson');
  assert.strictEqual(row3[4], 'Sick Leave');
  assert.strictEqual(row3[5], 'Eve Adams');
  assert.strictEqual(row3[6], 'Departed');
  assert.strictEqual(row3[7], 'No Responsible Available');
  assert.strictEqual(row3[8], 'CRITICAL: Uncovered');

  // Check conditional formatting rules
  var rules = dashboardSheet.getConditionalFormatRules();
  assert.strictEqual(rules.length, 3, 'Should have rules for Lead Active, Covered by Backup, and CRITICAL: Uncovered');

  // Check column widths
  assert.strictEqual(dashboardSheet.columnWidths[1], 120);
  assert.strictEqual(dashboardSheet.columnWidths[2], 220);
  assert.strictEqual(dashboardSheet.columnWidths[9], 180);
});

test('refreshActiveResponsibles: should toast update and refresh view', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('tbl_members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Frank Castle', '', '', 'Security Lead', 'Core', 'Active', 'Main Office']
  ]);

  var projectsSheet = mockSpreadsheet.getSheetByName('tbl_projects');
  projectsSheet.getRange(2, 1, 1, 7).setValues([
    ['PRJ-100', 'Security Hardening', 'Active', '2026-01-01', '2026-12-31', 'Frank Castle', '']
  ]);

  var res = code.refreshActiveResponsibles(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.projectsCount, 1);
  assert.strictEqual(res.leadActiveCount, 1);
  assert.ok(mockSpreadsheet.lastToast);
  assert.ok(mockSpreadsheet.lastToast.title.includes('Active Responsibles Updated'));
});

test('onOpen: should include Refresh Active Responsibles Dashboard in custom menu', () => {
  code.onOpen();
  var ui = SpreadsheetApp.getUi();
  assert.ok(ui.menu);
  var itemCaptions = ui.menu.items.map(function(item) { return item.caption; });
  assert.ok(itemCaptions.includes('Refresh Active Responsibles Dashboard'), 'Menu should have Refresh Active Responsibles Dashboard');
  assert.ok(itemCaptions.includes('Refresh Weekly Matrix View'), 'Menu should have Refresh Weekly Matrix View');
});

