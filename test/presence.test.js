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

const {
  applyProtections,
  setAuthorizedAdmins,
  getAuthorizedAdmins,
  getLastCreatedRow,
  isCellProtected
} = require('../src/ProtectionManager');
global.applyProtections = applyProtections;
global.setAuthorizedAdmins = setAuthorizedAdmins;
global.getAuthorizedAdmins = getAuthorizedAdmins;
global.getLastCreatedRow = getLastCreatedRow;
global.isCellProtected = isCellProtected;

const {
  isRowPopulated,
  validateRow,
  validateTable,
  validateAllTables,
  clearStaticCellBackgrounds,
  applyRequiredValidationFormatting
} = require('../src/ValidationManager');
global.isRowPopulated = isRowPopulated;
global.validateRow = validateRow;
global.validateTable = validateTable;
global.validateAllTables = validateAllTables;
global.clearStaticCellBackgrounds = clearStaticCellBackgrounds;
global.applyRequiredValidationFormatting = applyRequiredValidationFormatting;

const { freshInit } = require('../src/Init');
const { extractCellValue, formatDateIso, saveMemberBaselines } = require('../src/Sidebar');
global.extractCellValue = extractCellValue;
global.formatDateIso = formatDateIso;
global.saveMemberBaselines = saveMemberBaselines;

const { syncMembersAndTenures, copyCellChip } = require('../src/SyncEngine');
global.syncMembersAndTenures = syncMembersAndTenures;
global.copyCellChip = copyCellChip;

const {
  compilePresenceLogs,
  compilePresenceMigration,
  findEarliestHistoricalDate,
  getWeekdaysBetween,
  getWeekdayName
} = require('../src/PresenceEngine');
global.compilePresenceLogs = compilePresenceLogs;
global.compilePresenceMigration = compilePresenceMigration;
global.findEarliestHistoricalDate = findEarliestHistoricalDate;
global.getWeekdaysBetween = getWeekdaysBetween;
global.getWeekdayName = getWeekdayName;

const { compilePresence, compilePresenceMigration: codeCompileMigration, onOpen } = require('../src/Code');

test('getWeekdaysBetween: should only return Monday to Friday dates excluding weekends', () => {
  // Friday 2026-05-01 to Tuesday 2026-05-05: Fri (1), Mon (4), Tue (5) -> 3 weekdays
  var d1 = new Date(2026, 4, 1); // May 1, 2026 (Friday)
  var d2 = new Date(2026, 4, 5); // May 5, 2026 (Tuesday)

  var days = getWeekdaysBetween(d1, d2);
  assert.strictEqual(days.length, 3);
  assert.strictEqual(getWeekdayName(days[0]), 'Friday');
  assert.strictEqual(getWeekdayName(days[1]), 'Monday');
  assert.strictEqual(getWeekdayName(days[2]), 'Tuesday');
});

test('compilePresenceLogs: should expand working days for active members with baseline defaults', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var baselinesSheet = mockSpreadsheet.getSheetByName('Routines');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  // 1. Add active member in tbl_members
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '+33123456789', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);

  // 2. Set baseline routine for Alice (Mon=Remote, Tue-Fri=Main Office)
  saveMemberBaselines({
    member: 'alice@accenture.com',
    startDate: '2026-06-01',
    days: {
      Monday: { location_am: 'Remote', location_pm: 'Remote' },
      Tuesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Wednesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Thursday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Friday: { location_am: 'Remote', location_pm: 'Remote' }
    }
  });

  // 3. Compile for 1 week: Monday June 1, 2026 to Friday June 5, 2026 (5 weekdays)
  var res = compilePresenceLogs(new Date(2026, 5, 1), new Date(2026, 5, 5), mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.workingDaysCount, 5);
  assert.strictEqual(res.activeMembersCount, 1);
  assert.strictEqual(res.createdCount, 5);

  // 4. Verify tbl_presence_log rows
  var logRows = logSheet.getRange(2, 1, 5, 6).getValues();
  // Monday: Remote / Remote
  assert.strictEqual(logRows[0][1], '2026-06-01');
  assert.strictEqual(logRows[0][2], 'alice@accenture.com');
  assert.strictEqual(logRows[0][3], 'Remote');
  assert.strictEqual(logRows[0][4], 'Remote');
  assert.ok(String(logRows[0][0]).startsWith('log_'));

  // Tuesday: Main Office / Main Office
  assert.strictEqual(logRows[1][1], '2026-06-02');
  assert.strictEqual(logRows[1][3], 'Main Office');
  assert.strictEqual(logRows[1][4], 'Main Office');

  // Friday: Remote / Remote
  assert.strictEqual(logRows[4][1], '2026-06-05');
  assert.strictEqual(logRows[4][3], 'Remote');
});

test('compilePresenceLogs: should override baseline locations with approved leave records', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var leavesSheet = mockSpreadsheet.getSheetByName('Leaves');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['bob@accenture.com', 'bob@accenture.com', '+33123456780', 'Dev', 'Delivery', 'Active', 'Not Logged']
  ]);

  // Set baseline: All days Main Office
  saveMemberBaselines({
    member: 'bob@accenture.com',
    startDate: '2026-06-01',
    days: {
      Monday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Tuesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Wednesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Thursday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Friday: { location_am: 'Main Office', location_pm: 'Main Office' }
    }
  });

  // Approved Vacation on Wednesday June 3 to Thursday June 4
  // Pending Sick Leave on Friday June 5 (should NOT override because not Approved)
  leavesSheet.getRange(2, 1, 2, 6).setValues([
    ['leave_01', 'bob@accenture.com', '2026-06-03', '2026-06-04', 'Vacation', 'Approved'],
    ['leave_02', 'bob@accenture.com', '2026-06-05', '2026-06-05', 'Sick Leave', 'Requested']
  ]);

  var res = compilePresenceLogs(new Date(2026, 5, 1), new Date(2026, 5, 5), mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.createdCount, 5);

  var logRows = logSheet.getRange(2, 1, 5, 6).getValues();
  // Monday: Main Office
  assert.strictEqual(logRows[0][3], 'Main Office');
  // Wednesday: Overridden by Approved Vacation!
  assert.strictEqual(logRows[2][1], '2026-06-03');
  assert.strictEqual(logRows[2][3], 'Vacation');
  assert.strictEqual(logRows[2][4], 'Vacation');
  // Thursday: Overridden by Approved Vacation!
  assert.strictEqual(logRows[3][1], '2026-06-04');
  assert.strictEqual(logRows[3][3], 'Vacation');
  assert.strictEqual(logRows[3][4], 'Vacation');
  // Friday: Requested leave did NOT override, remains Main Office!
  assert.strictEqual(logRows[4][1], '2026-06-05');
  assert.strictEqual(logRows[4][3], 'Main Office');
});

test('compilePresenceLogs: should be idempotent and preserve existing manual edits', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);

  // First compilation
  var res1 = compilePresenceLogs(new Date(2026, 5, 1), new Date(2026, 5, 3), mockSpreadsheet);
  assert.strictEqual(res1.createdCount, 3);
  assert.strictEqual(res1.preservedCount, 0);

  // User manually edits Tuesday June 2 to "Client Site" in AM and "Branch Office" in PM
  logSheet.getRange(3, 4, 1, 2).setValues([['Client Site', 'Branch Office']]);

  // Re-run compilation on same date range
  var res2 = compilePresenceLogs(new Date(2026, 5, 1), new Date(2026, 5, 3), mockSpreadsheet);
  assert.strictEqual(res2.createdCount, 0); // No new rows created
  assert.strictEqual(res2.preservedCount, 3); // All 3 preserved

  // Check that manual edit was preserved intact
  var tuesdayRow = logSheet.getRange(3, 1, 1, 6).getValues()[0];
  assert.strictEqual(tuesdayRow[1], '2026-06-02');
  assert.strictEqual(tuesdayRow[3], 'Client Site');
  assert.strictEqual(tuesdayRow[4], 'Branch Office');
});

test('compilePresence: public entry point should toast compilation status', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', '', '', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);

  var report = compilePresence();
  assert.strictEqual(report.success, true);
  assert.ok(mockSpreadsheet.lastToast);
  assert.strictEqual(mockSpreadsheet.lastToast.title, 'Presence Compiled');
  assert.ok(mockSpreadsheet.lastToast.msg.includes('Compiled'));
});

test('compilePresenceLogs: should strictly cap compilation at today as ground truth database and reject future dates', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);

  // Suppose asOfDate is Wednesday June 3, 2026.
  // The caller requests compiling until Friday June 5, 2026 (future).
  var asOfWednesday = new Date(2026, 5, 3); // June 3, 2026
  var res = compilePresenceLogs(new Date(2026, 5, 1), new Date(2026, 5, 5), mockSpreadsheet, asOfWednesday);

  assert.strictEqual(res.success, true);
  // Must ONLY compile Monday June 1, Tuesday June 2, and Wednesday June 3 (3 days)!
  // Thursday June 4 and Friday June 5 must NOT be compiled into tbl_presence_log!
  assert.strictEqual(res.workingDaysCount, 3);
  assert.strictEqual(res.createdCount, 3);

  var logRows = logSheet.getRange(2, 1, 3, 6).getValues();
  assert.strictEqual(logRows[0][1], '2026-06-01');
  assert.strictEqual(logRows[1][1], '2026-06-02');
  assert.strictEqual(logRows[2][1], '2026-06-03');

  // Verify row 5 (which would be June 4) is completely empty!
  var row4Val = logSheet.getRange(5, 1, 1, 1).getValue();
  assert.strictEqual(row4Val, '');
});

test('compilePresenceLogs: should only compile weekdays and omit weekend logs when run on a weekend', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);

  // Saturday June 6, 2026
  var asOfSaturday = new Date(2026, 5, 6);
  // Request compilation from Monday June 1 to Saturday June 6
  var res = compilePresenceLogs(new Date(2026, 5, 1), new Date(2026, 5, 6), mockSpreadsheet, asOfSaturday);

  assert.strictEqual(res.success, true);
  // Mon (1), Tue (2), Wed (3), Thu (4), Fri (5) -> strictly 5 working days (Saturday excluded!)
  assert.strictEqual(res.workingDaysCount, 5);
  assert.strictEqual(res.createdCount, 5);

  var logRows = logSheet.getRange(2, 1, 5, 6).getValues();
  assert.strictEqual(logRows[4][1], '2026-06-05'); // Latest row is Friday June 5
  // Row 7 (which would be Saturday) is empty!
  var row6Val = logSheet.getRange(7, 1, 1, 1).getValue();
  assert.strictEqual(row6Val, '');

  // But tbl_members today_presence column displays 'Weekend' for today
  var todayPresenceVal = membersSheet.getRange(2, 7).getValue();
  assert.strictEqual(todayPresenceVal, 'Weekend');
});

test('findEarliestHistoricalDate: should detect earliest start date across tenures, baselines, leaves, and logs', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');
  var baselinesSheet = mockSpreadsheet.getSheetByName('Routines');
  var leavesSheet = mockSpreadsheet.getSheetByName('Leaves');

  // No dates present yet
  var early1 = findEarliestHistoricalDate(mockSpreadsheet);
  assert.strictEqual(early1, null);

  // Add tenure with start date 2025-09-01
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'Engineer', 'Core', '2025-09-01', '', 'Permanent', 1.0]
  ]);
  var early2 = findEarliestHistoricalDate(mockSpreadsheet);
  assert.strictEqual(formatDateIso(early2), '2025-09-01');

  // Add baseline with earlier date 2025-03-01
  baselinesSheet.getRange(2, 1, 1, 7).setValues([
    ['b-1', 'alice@accenture.com', '2025-03-01', '', 'All', 'Main Office', 'Main Office']
  ]);
  var early3 = findEarliestHistoricalDate(mockSpreadsheet);
  assert.strictEqual(formatDateIso(early3), '2025-03-01');

  // Add leave with even earlier date 2025-01-15
  leavesSheet.getRange(2, 1, 1, 6).setValues([
    ['lv-1', 'alice@accenture.com', '2025-01-15', '2025-01-20', 'Vacation', 'Approved']
  ]);
  var early4 = findEarliestHistoricalDate(mockSpreadsheet);
  assert.strictEqual(formatDateIso(early4), '2025-01-15');
});

test('compilePresenceMigration: should compile presence logs starting from earliest detected historical date', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');
  var baselinesSheet = mockSpreadsheet.getSheetByName('Routines');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);

  // Tenure start date in early 2026: 2026-01-05 (Monday)
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'Lead', 'Delivery', '2026-01-05', '', 'Permanent', 1.0]
  ]);

  // Baseline from 2026-01-05
  baselinesSheet.getRange(2, 1, 1, 7).setValues([
    ['b-1', 'alice@accenture.com', '2026-01-05', '', 'All', 'Remote', 'Remote']
  ]);

  var res = codeCompileMigration(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.ok(res.createdCount > 0, 'Should compile logs from historical start date');

  // Verify first compiled row is 2026-01-05
  var firstLogRow = logSheet.getRange(2, 1, 1, 5).getValues();
  assert.strictEqual(firstLogRow[0][1], '2026-01-05');
  assert.strictEqual(firstLogRow[0][2], 'alice@accenture.com');
  assert.strictEqual(firstLogRow[0][3], 'Remote');
});

test('compilePresenceMigration: should compile presence logs when baseline has undefined/empty start date', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');
  var baselinesSheet = mockSpreadsheet.getSheetByName('Routines');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);

  // Baseline has NO start date (empty string / undefined in sheet)
  baselinesSheet.getRange(2, 1, 1, 7).setValues([
    ['b-1', 'alice@accenture.com', '', '', 'All', 'Main Office', 'Remote']
  ]);

  // Tenure provides the historical anchor date
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'Lead', 'Delivery', '2026-02-02', '', 'Permanent', 1.0]
  ]);

  var res = codeCompileMigration(mockSpreadsheet);
  assert.strictEqual(res.success, true);

  var firstLogRow = logSheet.getRange(2, 1, 1, 5).getValues();
  assert.strictEqual(firstLogRow[0][1], '2026-02-02');
  assert.strictEqual(firstLogRow[0][3], 'Main Office');
  assert.strictEqual(firstLogRow[0][4], 'Remote');
});

test('compilePresenceMigration: should prompt user and respect custom entered start date', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var baselinesSheet = mockSpreadsheet.getSheetByName('Routines');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);
  baselinesSheet.getRange(2, 1, 1, 7).setValues([
    ['b-1', 'alice@accenture.com', '', '', 'All', 'Downtown Office', 'Downtown Office']
  ]);

  // Mock UI prompt response to return a user-entered date
  var ui = SpreadsheetApp.getUi();
  ui.promptResponse = {
    getSelectedButton: function() { return 'OK'; },
    getResponseText: function() { return '2026-03-02'; } // Monday March 2, 2026
  };

  var res = codeCompileMigration(mockSpreadsheet);
  assert.strictEqual(res.success, true);

  var firstLogRow = logSheet.getRange(2, 1, 1, 5).getValues();
  assert.strictEqual(firstLogRow[0][1], '2026-03-02');
  assert.strictEqual(firstLogRow[0][3], 'Downtown Office');

  // Reset promptResponse
  ui.promptResponse = null;
});

test('compilePresenceMigration: should cancel migration gracefully if cancelled by user', () => {
  freshInit(mockSpreadsheet);

  var ui = SpreadsheetApp.getUi();
  ui.promptResponse = {
    getSelectedButton: function() { return 'CANCEL'; },
    getResponseText: function() { return ''; }
  };

  var res = codeCompileMigration(mockSpreadsheet);
  assert.strictEqual(res.cancelled, true);
  assert.strictEqual(res.success, false);

  ui.promptResponse = null;
});

test('onOpen should include Compile Presence Logs (Migration / Undefined Start Date) in menu', () => {
  onOpen();
  var ui = SpreadsheetApp.getUi();
  assert.ok(ui.menu, 'Menu should be created');
  assert.strictEqual(ui.menu.name, 'Project Management');

  var foundMigrationItem = false;
  for (var i = 0; i < ui.menu.items.length; i++) {
    var item = ui.menu.items[i];
    if (item.caption && item.caption.indexOf('Migration') !== -1) {
      foundMigrationItem = true;
      assert.strictEqual(item.fn, 'compilePresenceMigration');
    }
  }
  assert.strictEqual(foundMigrationItem, true, 'Menu should include migration compilation item');
});

