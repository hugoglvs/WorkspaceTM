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
