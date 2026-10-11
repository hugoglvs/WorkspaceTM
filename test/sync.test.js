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

const { applyProtections } = require('../src/ProtectionManager');
global.applyProtections = applyProtections;

const { freshInit } = require('../src/Init');
const { extractCellValue, formatDateIso } = require('../src/Sidebar');
global.extractCellValue = extractCellValue;
global.formatDateIso = formatDateIso;
const { syncMembersAndTenures } = require('../src/SyncEngine');
const { onEdit } = require('../src/Code');

test('syncMembersAndTenures should populate tbl_members from tbl_tenures and create onboarding task', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  // Seed tenure directly for a new member (Columns: member, role, department, start_date, end_date, contract_type, fte_rate)
  tenuresSheet.getRange(2, 1, 1, 7).setValues([[
    'sarah.connor@accenture.com',
    'Tech Lead',
    'Cloud First',
    '2026-01-01',
    '',
    'Permanent',
    1.0
  ]]);

  var res = syncMembersAndTenures(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.addedToDirectory, 1);
  assert.strictEqual(res.tasksCreated, 1);

  // Check tbl_members
  // Columns: member(0), accenture_email(1), phone_number(2), current_role(3), current_department(4), tenure_status(5), today_presence(6)
  var memberRow = membersSheet.getRange(2, 1, 1, 7).getValues()[0];
  assert.strictEqual(memberRow[0], 'sarah.connor@accenture.com');
  assert.strictEqual(memberRow[1], 'sarah.connor@accenture.com'); // auto-filled accenture email
  assert.strictEqual(memberRow[3], 'Tech Lead');
  assert.strictEqual(memberRow[4], 'Cloud First');
  assert.strictEqual(memberRow[5], 'Active');

  // Check tbl_tasks: onboarding task created for sarah.connor@accenture.com
  // Columns: task_id(0), project_id(1), assigned_member(2), title(3), description(4), due_date(5), status(6), google_task_id(7)
  var taskRow = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  assert.strictEqual(taskRow[1], 'ONBOARDING');
  assert.strictEqual(taskRow[2], 'sarah.connor@accenture.com');
  assert.ok(taskRow[3].includes('enter Accenture email and phone number'));
  assert.ok(taskRow[4].includes('enter your Accenture email address and phone number'));
  assert.strictEqual(taskRow[6], 'Todo');
});

test('syncMembersAndTenures should update role and tenure_status for existing members', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');

  // Seed member in tbl_members
  membersSheet.getRange(2, 1, 1, 7).setValues([[
    'alice@example.com',
    'alice@accenture.com',
    '+33123456789',
    'Junior Analyst',
    'Advisory',
    'Active',
    'Remote'
  ]]);

  // Seed tenure in tbl_tenures with promotion to Manager
  tenuresSheet.getRange(2, 1, 1, 7).setValues([[
    'alice@example.com',
    'Manager',
    'Strategy',
    '2026-01-01',
    '',
    'Permanent',
    1.0
  ]]);

  var res = syncMembersAndTenures(mockSpreadsheet);
  assert.strictEqual(res.success, true);

  var updatedRow = membersSheet.getRange(2, 1, 1, 7).getValues()[0];
  assert.strictEqual(updatedRow[3], 'Manager');
  assert.strictEqual(updatedRow[4], 'Strategy');
  assert.strictEqual(updatedRow[5], 'Active');
  // Contact details preserved
  assert.strictEqual(updatedRow[1], 'alice@accenture.com');
  assert.strictEqual(updatedRow[2], '+33123456789');
});

test('applyProtections should lock read-only columns in tbl_members and unlock contact columns', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');

  var protections = membersSheet.getProtections();
  // Protected columns: member (col 1), current_role (col 4), current_department (col 5), tenure_status (col 6), today_presence (col 7)
  assert.strictEqual(protections.length, 5);

  var descriptions = protections.map(p => p.getDescription());
  assert.ok(descriptions.includes('Members_member_protected'));
  assert.ok(descriptions.includes('Members_current_role_protected'));
  assert.ok(descriptions.includes('Members_current_department_protected'));
  assert.ok(descriptions.includes('Members_tenure_status_protected'));
  assert.ok(descriptions.includes('Members_today_presence_protected'));
  // accenture_email and phone_number must NOT be in descriptions
  assert.ok(!descriptions.includes('Members_accenture_email_protected'));
  assert.ok(!descriptions.includes('Members_phone_number_protected'));

  // Ensure strict enforcement (warningOnly is false)
  for (var p = 0; p < protections.length; p++) {
    assert.strictEqual(protections[p].isWarningOnly(), false);
    var editors = protections[p].getEditors().map(e => e.getEmail().toLowerCase());
    assert.ok(editors.includes('owner@example.com'));
  }
});

test('syncMembersAndTenures should append multiple new members sequentially without overwriting row 2', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  // Add 3 different members into tbl_tenures
  tenuresSheet.getRange(2, 1, 3, 7).setValues([
    ['alice@accenture.com', 'Dev 1', 'Engineering', '2026-01-01', '', 'Permanent', 1.0],
    ['bob@accenture.com', 'Dev 2', 'Engineering', '2026-01-01', '', 'Contractor', 1.0],
    ['charlie@accenture.com', 'Dev 3', 'Engineering', '2026-01-01', '', 'Permanent', 0.8]
  ]);

  var res = syncMembersAndTenures(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.addedToDirectory, 3);
  assert.strictEqual(res.tasksCreated, 3);

  // Check tbl_members rows 2, 3, 4
  var mRows = membersSheet.getRange(2, 1, 3, 7).getValues();
  assert.strictEqual(mRows[0][0], 'alice@accenture.com');
  assert.strictEqual(mRows[0][3], 'Dev 1');
  assert.strictEqual(mRows[1][0], 'bob@accenture.com');
  assert.strictEqual(mRows[1][3], 'Dev 2');
  assert.strictEqual(mRows[2][0], 'charlie@accenture.com');
  assert.strictEqual(mRows[2][3], 'Dev 3');

  // Check tbl_tasks rows 2, 3, 4
  var tRows = tasksSheet.getRange(2, 1, 3, 7).getValues();
  assert.strictEqual(tRows[0][2], 'alice@accenture.com');
  assert.strictEqual(tRows[1][2], 'bob@accenture.com');
  assert.strictEqual(tRows[2][2], 'charlie@accenture.com');
});

test('onEdit in tbl_tenures automatically adds declared member to tbl_members', () => {
  freshInit(mockSpreadsheet);
  var tenuresSheet = mockSpreadsheet.getSheetByName('Contracts');
  var membersSheet = mockSpreadsheet.getSheetByName('Members');

  // Declare a new tenure on row 2
  tenuresSheet.getRange(2, 1, 1, 7).setValues([
    ['Hugo Gonçalves', 'Senior Lead', 'Strategy', '2026-01-01', '', 'Permanent', 1.0]
  ]);

  var mockEvent = {
    range: tenuresSheet.getRange(2, 1, 1, 1)
  };

  onEdit(mockEvent);

  // Assert member was automatically added to tbl_members
  var memberRow = membersSheet.getRange(2, 1, 1, 7).getValues()[0];
  assert.strictEqual(memberRow[0], 'Hugo Gonçalves');
  assert.strictEqual(memberRow[3], 'Senior Lead');
  assert.strictEqual(memberRow[4], 'Strategy');
  assert.strictEqual(memberRow[5], 'Active');
});

