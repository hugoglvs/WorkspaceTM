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

const {
  applyProtections,
  setAuthorizedAdmins,
  getAuthorizedAdmins,
  getLastCreatedRow,
  isCellProtected,
  handleProtectedCellEdit,
  isProtectionWarningSnoozed,
  snoozeProtectionWarnings,
  clearProtectionWarningSnooze
} = require('../src/ProtectionManager');
global.applyProtections = applyProtections;
global.setAuthorizedAdmins = setAuthorizedAdmins;
global.getAuthorizedAdmins = getAuthorizedAdmins;
global.getLastCreatedRow = getLastCreatedRow;
global.isCellProtected = isCellProtected;
global.handleProtectedCellEdit = handleProtectedCellEdit;
global.isProtectionWarningSnoozed = isProtectionWarningSnoozed;
global.snoozeProtectionWarnings = snoozeProtectionWarnings;
global.clearProtectionWarningSnooze = clearProtectionWarningSnooze;

const {
  PROMINENT_WARNING_BG,
  PROMINENT_WARNING_FONT,
  isRowPopulated,
  validateRow,
  validateTable,
  validateAllTables,
  highlightInvalidCells,
  applyRequiredValidationFormatting
} = require('../src/ValidationManager');
global.PROMINENT_WARNING_BG = PROMINENT_WARNING_BG;
global.PROMINENT_WARNING_FONT = PROMINENT_WARNING_FONT;
global.isRowPopulated = isRowPopulated;
global.validateRow = validateRow;
global.validateTable = validateTable;
global.validateAllTables = validateAllTables;
global.highlightInvalidCells = highlightInvalidCells;
global.applyRequiredValidationFormatting = applyRequiredValidationFormatting;

const { freshInit } = require('../src/Init');
const { extractCellValue, formatDateIso } = require('../src/Sidebar');
global.extractCellValue = extractCellValue;
global.formatDateIso = formatDateIso;

const { syncMembersAndTenures } = require('../src/SyncEngine');
const { onEdit, validateAndReport } = require('../src/Code');

test('validateRow: should recognize completely empty rows as not populated', () => {
  var emptyRow = ['', '', '', '', '', '', ''];
  var res = validateRow('tbl_tenures', emptyRow);
  assert.strictEqual(res.valid, true);
  assert.strictEqual(res.populated, false);
  assert.deepStrictEqual(res.missingColumns, []);
});

test('validateRow: should pass tbl_tenures row with empty end_date (nullable)', () => {
  // Columns: member, role, department, start_date, end_date, contract_type, fte_rate
  var validTenureOngoing = [
    'alice@accenture.com',
    'Tech Lead',
    'Cloud First',
    '2026-01-01',
    '', // end_date is null / empty for ongoing tenure
    'Permanent',
    1.0
  ];
  var res = validateRow('tbl_tenures', validTenureOngoing);
  assert.strictEqual(res.valid, true);
  assert.strictEqual(res.populated, true);
  assert.deepStrictEqual(res.missingColumns, []);
});

test('validateRow: should fail tbl_tenures row when required columns are missing', () => {
  // Missing role and start_date
  var invalidTenure = [
    'bob@accenture.com',
    '', // missing role
    'Cloud First',
    '', // missing start_date
    '', // nullable end_date
    'Permanent',
    1.0
  ];
  var res = validateRow('tbl_tenures', invalidTenure);
  assert.strictEqual(res.valid, false);
  assert.strictEqual(res.populated, true);
  assert.ok(res.missingColumns.includes('role'));
  assert.ok(res.missingColumns.includes('start_date'));
  assert.ok(!res.missingColumns.includes('end_date')); // end_date must NOT be in missingColumns
});

test('validateRow: should pass tbl_members with empty accenture_email and phone_number (nullable)', () => {
  // Columns: member, accenture_email, phone_number, current_role, current_department, tenure_status, today_presence
  var newMemberRow = [
    'charlie@accenture.com',
    '', // nullable accenture_email
    '', // nullable phone_number
    'Developer',
    'Delivery',
    'Active',
    'Not Logged'
  ];
  var res = validateRow('tbl_members', newMemberRow);
  assert.strictEqual(res.valid, true);
  assert.strictEqual(res.populated, true);
  assert.deepStrictEqual(res.missingColumns, []);
});

test('validateTable: should audit populated sheet and flag invalid rows', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');

  // Row 2: valid ongoing tenure (end_date empty)
  // Row 3: invalid tenure (missing role and contract_type)
  tenuresSheet.getRange(2, 1, 2, 7).setValues([
    ['alice@accenture.com', 'Manager', 'Strategy', '2026-01-01', '', 'Permanent', 1.0],
    ['bad@accenture.com', '', 'Strategy', '2026-01-01', '', '', 1.0]
  ]);

  var report = validateTable('tbl_tenures', mockSpreadsheet);
  assert.strictEqual(report.populatedRows, 2);
  assert.strictEqual(report.validRows, 1);
  assert.strictEqual(report.invalidRows.length, 1);
  assert.strictEqual(report.invalidRows[0].row, 3);
  assert.ok(report.invalidRows[0].missingColumns.includes('role'));
  assert.ok(report.invalidRows[0].missingColumns.includes('contract_type'));
});

test('validateAllTables: should summarize health across all tables', () => {
  freshInit(mockSpreadsheet);
  var res = validateAllTables(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.totalIssues, 0);
  assert.ok(res.tables.tenures);
  assert.ok(res.tables.members);
});

test('syncMembersAndTenures: should report validation warnings on invalid tenure rows while supporting null end_date', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');
  var membersSheet = mockSpreadsheet.getSheetByName('Members');

  // Row 2: Valid ongoing tenure with null end_date -> Active
  // Row 3: Incomplete tenure missing department & contract_type
  tenuresSheet.getRange(2, 1, 2, 7).setValues([
    ['alice@accenture.com', 'Senior Dev', 'Engineering', '2026-01-01', '', 'Permanent', 1.0],
    ['incomplete@accenture.com', 'Lead', '', '2026-01-01', '', '', 1.0]
  ]);

  var res = syncMembersAndTenures(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.validationWarnings.length, 1);
  assert.strictEqual(res.validationWarnings[0].row, 3);
  assert.ok(res.validationWarnings[0].missingColumns.includes('department'));
  assert.ok(res.validationWarnings[0].missingColumns.includes('contract_type'));

  // Alice's ongoing tenure must be Active
  var mRows = membersSheet.getRange(2, 1, 2, 7).getValues();
  assert.strictEqual(mRows[0][0], 'alice@accenture.com');
  assert.strictEqual(mRows[0][5], 'Active');
});

test('onEdit: should trigger toast warning when editing an incomplete row', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');

  // Seed an incomplete row (Row 2 missing role & start_date)
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['bob@accenture.com', '', 'Engineering', '', '', 'Permanent', 1.0]
  ]);

  var mockEvent = {
    range: tenuresSheet.getRange(2, 2, 1, 1)
  };

  onEdit(mockEvent);

  assert.ok(mockSpreadsheet.lastToast);
  assert.strictEqual(mockSpreadsheet.lastToast.title, 'Required Field Missing');
  assert.ok(mockSpreadsheet.lastToast.msg.includes('role'));
});

test('validateAndReport: should alert when issues exist and toast when all valid', () => {
  freshInit(mockSpreadsheet);
  // Initially empty sheets -> all valid
  validateAndReport();
  assert.ok(mockSpreadsheet.lastToast);
  assert.strictEqual(mockSpreadsheet.lastToast.title, 'Validation Passed');

  // Add invalid row
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['broken@accenture.com', '', '', '', '', '', '']
  ]);

  var ui = SpreadsheetApp.getUi();
  validateAndReport();
  assert.ok(ui.lastAlert);
  assert.ok(ui.lastAlert.title.includes('Validation Issues Found'));
});

test('highlightInvalidCells: should color missing required cells with prominent warning background', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');

  // Populate row 2 with member and start_date, but missing role, department, contract_type, fte_rate (end_date nullable)
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', '', '', '2026-01-01', '', '', '']
  ]);

  var count = highlightInvalidCells('tbl_tenures', mockSpreadsheet);
  assert.strictEqual(count, 4); // role, department, contract_type, fte_rate

  // Check background of role cell (Row 2, Col 2) -> PROMINENT_WARNING_BG (#F8D7DA)
  var roleCell = tenuresSheet.getRange(2, 2, 1, 1);
  assert.strictEqual(roleCell.background, '#F8D7DA');
  assert.strictEqual(roleCell.fontColor, '#721C24');
  assert.strictEqual(roleCell.fontWeight, 'bold');

  // Check end_date cell (Row 2, Col 5) -> nullable, so background must remain null / uncolored!
  var endDateCell = tenuresSheet.getRange(2, 5, 1, 1);
  assert.strictEqual(endDateCell.background, null);
});

test('applyRequiredValidationFormatting: should configure prominent alert color and bold font', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');
  var rules = tenuresSheet.getConditionalFormatRules();

  // Rules applied for each required column in tbl_tenures (6 required columns)
  assert.strictEqual(rules.length, 6);
  for (var i = 0; i < rules.length; i++) {
    assert.strictEqual(rules[i].bg, '#F8D7DA');
    assert.strictEqual(rules[i].color, '#721C24');
    assert.strictEqual(rules[i].bold, true);
    assert.ok(rules[i].formula.includes('COUNTA($A2:$G2)>0'));
    assert.ok(rules[i].formula.includes('LEN(TRIM('));
  }
});

test('onEdit: should trigger admin warning toast when editing a strictly protected column', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');

  // tbl_members Column 4 is 'current_role' (strictly protected)
  var mockEvent = {
    range: membersSheet.getRange(2, 4, 1, 1)
  };

  onEdit(mockEvent);

  assert.ok(mockSpreadsheet.lastToast);
  assert.strictEqual(mockSpreadsheet.lastToast.title, 'Strictly Protected Column Edited');
  assert.ok(mockSpreadsheet.lastToast.msg.includes('current_role'));
  assert.ok(mockSpreadsheet.lastToast.msg.includes('Admin Warning'));
});

test('onEdit: should trigger warning toast when editing a warning-protected column on a created row', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');

  // Populate row 2 as a valid created tenure
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'Manager', 'Strategy', '2026-01-01', '', 'Permanent', 1.0]
  ]);
  applyProtections(mockSpreadsheet);

  // tbl_tenures Column 1 is 'member' (warning protected on created row 2)
  var mockEvent = {
    range: tenuresSheet.getRange(2, 1, 1, 1)
  };

  onEdit(mockEvent);

  assert.ok(mockSpreadsheet.lastToast);
  assert.strictEqual(mockSpreadsheet.lastToast.title, 'Warning-Protected Column Edited');
  assert.ok(mockSpreadsheet.lastToast.msg.includes('member'));
});

test('onEdit: should NOT trigger warning toast when registering a new tenure on an uncreated row', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');

  // Populate only row 2
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'Manager', 'Strategy', '2026-01-01', '', 'Permanent', 1.0]
  ]);
  applyProtections(mockSpreadsheet);

  mockSpreadsheet.lastToast = null;

  // Row 3 is a new, uncreated row where user is registering a new tenure
  var newRowEvent = {
    range: tenuresSheet.getRange(3, 1, 1, 1)
  };

  onEdit(newRowEvent);

  // Must NOT show warning toast!
  assert.strictEqual(mockSpreadsheet.lastToast, null);
});

test('applyProtections: should only protect created rows in tbl_tenures, leaving empty rows unprotected', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');

  // Populate row 2
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'Manager', 'Strategy', '2026-01-01', '', 'Permanent', 1.0]
  ]);
  applyProtections(mockSpreadsheet);

  // Cell (2, 1) must be protected
  var cellRow2 = isCellProtected(tenuresSheet, 2, 1);
  assert.strictEqual(cellRow2.isProtected, true);
  assert.strictEqual(cellRow2.isWarningOnly, true);

  // Cell (3, 1) - empty row for new tenure - MUST NOT be protected!
  var cellRow3 = isCellProtected(tenuresSheet, 3, 1);
  assert.strictEqual(cellRow3.isProtected, false);
});

test('applyProtections: should strictly restrict editors to authorized admins and set warningOnly to false', () => {
  setAuthorizedAdmins(['lead@accenture.com']);
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var protections = membersSheet.getProtections();

  assert.ok(protections.length > 0);
  for (var i = 0; i < protections.length; i++) {
    var p = protections[i];
    // Strict block, not warning-only
    assert.strictEqual(p.isWarningOnly(), false);
    var editorEmails = p.getEditors().map(ed => ed.getEmail().toLowerCase());
    assert.ok(editorEmails.includes('owner@example.com'));
    assert.ok(editorEmails.includes('lead@accenture.com'));
  }
});

test('onEdit: should prompt user with blocking alert asking to snooze notifications for 5 minutes', () => {
  freshInit(mockSpreadsheet);
  clearProtectionWarningSnooze();
  var membersSheet = mockSpreadsheet.getSheetByName('Members');

  var ui = SpreadsheetApp.getUi();
  ui.alertResponse = 'YES';

  var mockEvent = {
    range: membersSheet.getRange(2, 4, 1, 1) // current_role (strictly protected)
  };

  onEdit(mockEvent);

  // Check that ui.alert was presented with YES_NO buttons
  assert.ok(ui.lastAlert);
  assert.strictEqual(ui.lastAlert.title, 'Strictly Protected Column Edited');
  assert.ok(ui.lastAlert.msg.includes('disable these warning notifications for the next 5 minutes'));
  assert.strictEqual(ui.lastAlert.buttons, 'YES_NO');

  // Check snooze was set
  assert.strictEqual(isProtectionWarningSnoozed(), true);
  assert.ok(mockSpreadsheet.lastToast);
  assert.strictEqual(mockSpreadsheet.lastToast.title, 'Notifications Snoozed');

  // Subsequent edit within snooze window should not trigger alert or warning toast
  ui.lastAlert = null;
  mockSpreadsheet.lastToast = null;
  onEdit(mockEvent);

  assert.strictEqual(ui.lastAlert, null);
  assert.strictEqual(mockSpreadsheet.lastToast, null);

  // Clearing snooze restores warning
  clearProtectionWarningSnooze();
  assert.strictEqual(isProtectionWarningSnoozed(), false);
  ui.alertResponse = 'NO';
  onEdit(mockEvent);
  assert.ok(ui.lastAlert);
  assert.strictEqual(mockSpreadsheet.lastToast.title, 'Strictly Protected Column Edited');
});


