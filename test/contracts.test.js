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
Object.assign(global, syncEngine);
const { freshInit } = require('../src/Init');
const sidebar = require('../src/Sidebar');
global.extractCellValue = sidebar.extractCellValue;
global.isUserMemberMatch = sidebar.isUserMemberMatch;
global.formatDateIso = sidebar.formatDateIso;

const contracts = require('../src/Contracts');
Object.assign(global, contracts);
const code = require('../src/Code');

function seed(tenureRows) {
  freshInit(mockSpreadsheet);
  if (tenureRows.length) {
    mockSpreadsheet.getSheetByName('Contracts').getRange(2, 1, tenureRows.length, 7).setValues(tenureRows);
  }
  syncEngine.syncMembersAndTenures(mockSpreadsheet);
}

function tenures() {
  return mockSpreadsheet.getSheetByName('Contracts').getRange(2, 1, 20, 7).getValues().filter(r => r[0]);
}


const HUGO = ['Hugo Gonçalves', 'Data Analyst', 'FAL ME', '2025-09-02', '', 'Permanent', 100];

test('Declare Contract: unhides and opens the Contracts tab on its first empty row', () => {
  seed([HUGO]);
  var sheet = mockSpreadsheet.getSheetByName('Contracts');
  sheet.hideSheet();

  var res = contracts.showContractsTab();

  assert.strictEqual(res.success, true);
  assert.strictEqual(sheet.isSheetHidden(), false);
  assert.strictEqual(sheet.activated, true);
  assert.strictEqual(res.row, 3);
  assert.ok(mockSpreadsheet.lastToast.msg.includes('one row per contract'));
});

test('onEdit: typing a new name in Members opens the Contracts tab for that person', () => {
  seed([HUGO]);
  var contractsSheet = mockSpreadsheet.getSheetByName('Contracts');
  contractsSheet.hideSheet();
  var cell = mockSpreadsheet.getSheetByName('Members').getRange(3, 1);
  cell.setValue('Ines Martin');

  code.onEdit({ range: cell, value: 'Ines Martin', source: mockSpreadsheet });

  assert.strictEqual(contractsSheet.isSheetHidden(), false);
  assert.ok(mockSpreadsheet.lastToast.msg.startsWith('Ines Martin has no contract yet.'));
  assert.strictEqual(SpreadsheetApp.getUi().shownSidebar, null, 'no sidebar any more');
});

test('onEdit: editing an existing Member with a contract does not open anything', () => {
  seed([HUGO]);
  var contractsSheet = mockSpreadsheet.getSheetByName('Contracts');
  contractsSheet.hideSheet();
  var cell = mockSpreadsheet.getSheetByName('Members').getRange(2, 3);
  cell.setValue('0102');

  code.onEdit({ range: cell, value: '0102', source: mockSpreadsheet });

  assert.strictEqual(contractsSheet.isSheetHidden(), true);
  assert.strictEqual(contracts.promptContractForNewMember({ range: mockSpreadsheet.getSheetByName('Members').getRange(2, 1), source: mockSpreadsheet }), null);
});

function d(iso) { var p = iso.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
const TODAY = d('2026-10-13');

test('contractProfile: active, last day, upcoming, departed, blank role, no contract', () => {
  var P = contracts.contractProfile;
  assert.deepStrictEqual(P([{ role: 'Dev', department: 'D1', start: d('2025-01-01'), end: null }], TODAY),
    { role: 'Dev', department: 'D1', status: 'Active' });
  assert.strictEqual(P([{ role: '', department: 'D1', start: d('2026-01-01'), end: d('2026-10-13') }], TODAY).status, 'Active', 'last day still active');
  assert.strictEqual(P([{ role: '', department: '', start: d('2026-01-01'), end: d('2026-10-13') }], TODAY).role, 'Team Member');
  assert.deepStrictEqual(P([
    { role: 'Old', department: 'D1', start: d('2024-01-01'), end: d('2025-12-31') },
    { role: 'Late', department: 'D4', start: d('2026-12-01'), end: null },
    { role: 'Early', department: 'D5', start: d('2026-11-15'), end: null }
  ], TODAY), { role: 'Early', department: 'D5', status: 'Upcoming' });
  assert.deepStrictEqual(P([
    { role: 'Y', department: 'D3', start: d('2024-02-01'), end: d('2025-06-30') },
    { role: 'X', department: 'D1', start: d('2023-01-01'), end: d('2024-01-01') }
  ], TODAY), { role: 'Y', department: 'D3', status: 'Departed' });
  assert.deepStrictEqual(P([], TODAY), { role: '', department: '', status: 'No Contract' });
});

test('Members formulas: written into every table row by the sync, reading the Contracts tab by name', () => {
  seed([HUGO]);
  var members = mockSpreadsheet.getSheetByName('Members');
  var f = members.getRange(2, 4, 1, 3).getFormulas()[0];
  assert.ok(f[0].startsWith('=IF($A2="",""') && f[0].includes("'Contracts'!$B$2:$B"), 'role reads Contracts role');
  assert.ok(f[1].includes("'Contracts'!$C$2:$C") && f[1].includes('"Delivery"'));
  assert.ok(f[2].includes('"No Contract"') && f[2].includes('TODAY()'));
  var lastRow = members.getRange(20, 4, 1, 3).getFormulas()[0];
  assert.ok(lastRow[2].includes('$A20'), 'empty table rows get the formula too');
  assert.deepStrictEqual(members.getRange(2, 4, 1, 3).getValues()[0], ['Data Analyst', 'FAL ME', 'Active']);
});

test('Members formulas: a new contract changes Role and Status without any write to Members', () => {
  seed([HUGO]);
  var contractsSheet = mockSpreadsheet.getSheetByName('Contracts');
  contractsSheet.getRange(2, 5).setValue('2026-10-01');
  contractsSheet.getRange(3, 1, 1, 7).setValues([['Hugo Gonçalves', 'Manager', 'FAL ME', '2027-01-01', '', 'Permanent', 100]]);

  assert.deepStrictEqual(mockSpreadsheet.getSheetByName('Members').getRange(2, 4, 1, 3).getValues()[0], ['Manager', 'FAL ME', 'Upcoming']);
});

test('Members formulas: typed-over cells are reported by Check Workspace and restored', () => {
  seed([HUGO]);
  var members = mockSpreadsheet.getSheetByName('Members');
  members.getRange(2, 6).setValue('Active');
  assert.deepStrictEqual(contracts.findTypedOverMemberCells(mockSpreadsheet), ['Hugo Gonçalves: Status']);

  var res = contracts.restoreMemberFormulas(mockSpreadsheet);
  assert.strictEqual(res.message, '1 cell(s) restored.');
  assert.deepStrictEqual(contracts.findTypedOverMemberCells(mockSpreadsheet), []);
  assert.strictEqual(contracts.restoreMemberFormulas(mockSpreadsheet).message, 'Nothing to restore.');
});

test('Members formulas: a renamed Contracts tab gets the formulas rewritten', () => {
  seed([HUGO]);
  mockSpreadsheet.getSheetByName('Contracts').setName("Team's contracts");
  global.getTable = function(ss, key) {
    return key === 'tenures' ? ss.getSheetByName("Team's contracts") : ss.getSheetByName(require('../src/Schema').SCHEMA_SPEC[key].sheetName);
  };
  try {
    contracts.applyMemberFormulas(mockSpreadsheet);
    var f = mockSpreadsheet.getSheetByName('Members').getRange(2, 6).getFormula();
    assert.ok(f.includes("'Team''s contracts'!$A$2:$A"), f);
    assert.strictEqual(mockSpreadsheet.getSheetByName('Members').getRange(2, 6).getValue(), 'Active');
  } finally {
    global.getTable = require('../src/Tables').getTable;
  }
});
