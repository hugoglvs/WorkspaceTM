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

const members = require('../src/MembersService');
Object.assign(global, members);
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

function memberRow(name) {
  return mockSpreadsheet.getSheetByName('Members').getRange(2, 1, 20, 7).getValues().find(r => r[0] === name);
}

const HUGO = ['Hugo Gonçalves', 'Data Analyst', 'FAL ME', '2025-09-02', '', 'Permanent', 100];

test('getMembersSidebarData: lists Members with their Tenures and the known roles and departments', () => {
  seed([HUGO]);

  var data = members.getMembersSidebarData(mockSpreadsheet);

  assert.strictEqual(data.members.length, 1);
  assert.strictEqual(data.members[0].name, 'Hugo Gonçalves');
  assert.strictEqual(data.members[0].status, 'Active');
  assert.deepStrictEqual(data.members[0].tenures, [{
    startDate: '2025-09-02', endDate: '', role: 'Data Analyst', department: 'FAL ME', contractType: 'Permanent', ftePercent: 100
  }]);
  assert.deepStrictEqual(data.departments, ['FAL ME']);
  assert.deepStrictEqual(data.roles, ['Data Analyst']);
  assert.ok(data.contractTypes.includes('Contractor'));
});

test('saveMemberTenure: a new Member gets a Tenure, a directory row and their contact details', () => {
  seed([]);

  var res = members.saveMemberTenure({
    member: 'Ines Martin', email: 'ines@accenture.com', phone: '+33 6 00 00 00 00',
    role: 'Consultant', department: 'FAL ME', startDate: '2026-11-02', endDate: '',
    contractType: 'Contractor', ftePercent: 80
  }, mockSpreadsheet);

  assert.strictEqual(res.success, true);
  assert.deepStrictEqual(tenures(), [['Ines Martin', 'Consultant', 'FAL ME', '2026-11-02', '', 'Contractor', 80]]);
  var row = memberRow('Ines Martin');
  assert.ok(row, 'directory row created');
  assert.strictEqual(row[1], 'ines@accenture.com');
  assert.strictEqual(row[2], '+33 6 00 00 00 00');
  assert.strictEqual(row[5], 'Upcoming');
});

test('saveMemberTenure: a new contract ends the open one the day before', () => {
  seed([HUGO]);

  var res = members.saveMemberTenure({
    member: 'Hugo Gonçalves', role: 'Senior Data Analyst', department: 'FAL ME',
    startDate: '2026-11-01', contractType: 'Permanent', ftePercent: 100
  }, mockSpreadsheet);

  assert.strictEqual(res.success, true);
  assert.strictEqual(res.closedPrevious, true);
  assert.deepStrictEqual(tenures(), [
    ['Hugo Gonçalves', 'Data Analyst', 'FAL ME', '2025-09-02', '2026-10-31', 'Permanent', 100],
    ['Hugo Gonçalves', 'Senior Data Analyst', 'FAL ME', '2026-11-01', '', 'Permanent', 100]
  ]);
});

test('saveMemberTenure: refuses a contract starting on or before an existing one, and missing fields', () => {
  seed([HUGO]);

  var clash = members.saveMemberTenure({
    member: 'Hugo Gonçalves', role: 'Analyst', department: 'FAL ME',
    startDate: '2025-09-02', contractType: 'Permanent', ftePercent: 100
  }, mockSpreadsheet);
  var missing = members.saveMemberTenure({ member: 'Ines Martin', role: '', department: 'FAL ME', startDate: '2026-11-02', contractType: 'Contractor', ftePercent: 80 }, mockSpreadsheet);
  var badFte = members.saveMemberTenure({ member: 'Ines Martin', role: 'Consultant', department: 'FAL ME', startDate: '2026-11-02', contractType: 'Contractor', ftePercent: 0 }, mockSpreadsheet);

  assert.strictEqual(clash.success, false);
  assert.ok(clash.message.includes('2025-09-02'));
  assert.strictEqual(missing.success, false);
  assert.strictEqual(badFte.success, false);
  assert.strictEqual(tenures().length, 1, 'nothing written');
});

test('endMemberTenure: sets the end date of a Tenure and updates the directory status', () => {
  seed([HUGO]);

  var res = members.endMemberTenure({ member: 'Hugo Gonçalves', startDate: '2025-09-02', endDate: '2026-01-15' }, mockSpreadsheet);

  assert.strictEqual(res.success, true);
  assert.strictEqual(tenures()[0][4], '2026-01-15');
  assert.strictEqual(memberRow('Hugo Gonçalves')[5], 'Departed');
  assert.strictEqual(members.endMemberTenure({ member: 'Hugo Gonçalves', startDate: '2025-09-02', endDate: '2025-01-01' }, mockSpreadsheet).success, false, 'end before start refused');
});

test('saveMemberProfile: updates email and phone', () => {
  seed([HUGO]);

  members.saveMemberProfile({ member: 'Hugo Gonçalves', email: 'hugo@accenture.com', phone: '0102' }, mockSpreadsheet);

  assert.strictEqual(memberRow('Hugo Gonçalves')[1], 'hugo@accenture.com');
  assert.strictEqual(memberRow('Hugo Gonçalves')[2], '0102');
});

test('onEdit: typing a new name in Members opens the sidebar pre-filled for that person', () => {
  seed([HUGO]);
  var cell = mockSpreadsheet.getSheetByName('Members').getRange(3, 1);
  cell.setValue('Ines Martin');

  code.onEdit({ range: cell, value: 'Ines Martin', source: mockSpreadsheet });

  var shown = SpreadsheetApp.getUi().shownSidebar;
  assert.ok(shown, 'sidebar opened');
  assert.strictEqual(shown.filename, 'MembersSidebar');
  assert.strictEqual(shown.template.prefillMember, 'Ines Martin');
});

test('onEdit: when the sidebar cannot open, a toast points to the menu instead', () => {
  seed([HUGO]);
  SpreadsheetApp.getUi().showSidebar = function() { throw new Error('You do not have permission to call Ui.showSidebar'); };
  var cell = mockSpreadsheet.getSheetByName('Members').getRange(3, 1);
  cell.setValue('Ines Martin');

  code.onEdit({ range: cell, value: 'Ines Martin', source: mockSpreadsheet });

  assert.ok(mockSpreadsheet.lastToast.msg.includes('Members & Contracts'));
});

test('onEdit: editing an existing Member with a Tenure does not open anything', () => {
  seed([HUGO]);
  var cell = mockSpreadsheet.getSheetByName('Members').getRange(2, 3);
  cell.setValue('0102');

  code.onEdit({ range: cell, value: '0102', source: mockSpreadsheet });

  assert.strictEqual(SpreadsheetApp.getUi().shownSidebar, null);
});

test('saveMemberTenure: ends the current contract on the date given in the form', () => {
  seed([HUGO]);

  var res = members.saveMemberTenure({
    member: 'Hugo Gonçalves', role: 'Senior Data Analyst', department: 'FAL ME',
    startDate: '2026-11-01', contractType: 'Permanent', ftePercent: 100, previousEndDate: '2026-10-31'
  }, mockSpreadsheet);

  assert.strictEqual(res.success, true);
  assert.strictEqual(res.gap, undefined);
  assert.strictEqual(tenures()[0][4], '2026-10-31');
});

test('saveMemberTenure: refuses a current-contract end date that overlaps the new contract or precedes its start', () => {
  seed([HUGO]);
  var base = { member: 'Hugo Gonçalves', role: 'Senior Data Analyst', department: 'FAL ME', startDate: '2026-11-01', contractType: 'Permanent', ftePercent: 100 };

  var overlap = members.saveMemberTenure(Object.assign({}, base, { previousEndDate: '2026-11-01' }), mockSpreadsheet);
  var beforeStart = members.saveMemberTenure(Object.assign({}, base, { previousEndDate: '2025-01-01' }), mockSpreadsheet);

  assert.strictEqual(overlap.success, false);
  assert.ok(overlap.message.includes('2026-11-01'));
  assert.strictEqual(beforeStart.success, false);
  assert.deepStrictEqual(tenures().map(t => t[4]), [''], 'nothing written');
});

test('saveMemberTenure: allows a gap between contracts and warns about it', () => {
  seed([HUGO]);

  var res = members.saveMemberTenure({
    member: 'Hugo Gonçalves', role: 'Senior Data Analyst', department: 'FAL ME',
    startDate: '2026-11-15', contractType: 'Permanent', ftePercent: 100, previousEndDate: '2026-10-31'
  }, mockSpreadsheet);

  assert.strictEqual(res.success, true);
  assert.deepStrictEqual(res.gap, { from: '2026-11-01', to: '2026-11-14' });
  assert.ok(res.message.includes('2026-11-01') && res.message.includes('2026-11-14'));
});

test('saveMemberTenure: a new Member declared by email gets a People chip in Contracts', () => {
  seed([]);
  global.Sheets.batchUpdateCalls.length = 0;

  var res = members.saveMemberTenure({
    member: 'Ines.Martin@gmail.com', role: 'Consultant', department: 'FAL ME', startDate: '2026-11-02',
    contractType: 'Contractor', ftePercent: 100
  }, mockSpreadsheet);

  assert.strictEqual(res.success, true, res.message);
  var chipCall = global.Sheets.batchUpdateCalls.find(c => c.resource.requests.some(r => r.updateCells && r.updateCells.rows[0].values[0].chipRuns));
  assert.ok(chipCall, 'a chipRuns write is sent');
  var cell = chipCall.resource.requests[0].updateCells.rows[0].values[0];
  assert.strictEqual(cell.chipRuns[0].chip.personProperties.email, 'ines.martin@gmail.com');
  assert.strictEqual(tenures()[0][0], 'ines.martin@gmail.com');
  assert.ok(memberRow('ines.martin@gmail.com'), 'the chip is copied into Members');
});

test('saveMemberTenure: the email of an existing Member adds a contract to that Member, not a new one', () => {
  seed([['Hugo Gonçalves', 'Data Analyst', 'FAL ME', '2025-09-02', '', 'Permanent', 100]]);
  members.saveMemberProfile({ member: 'Hugo Gonçalves', email: 'hugo@accenture.com' }, mockSpreadsheet);

  var res = members.saveMemberTenure({
    member: 'HUGO@accenture.com', role: 'Manager', department: 'FAL ME', startDate: '2026-11-02',
    contractType: 'Permanent', ftePercent: 100
  }, mockSpreadsheet);

  assert.strictEqual(res.success, true, res.message);
  assert.strictEqual(res.closedPrevious, true);
  assert.deepStrictEqual(tenures().map(t => t[0]), ['Hugo Gonçalves', 'Hugo Gonçalves']);
});

test('saveMemberTenure: a chip that resolves to an existing Member goes through the contract checks', () => {
  seed([['Hugo Gonçalves', 'Data Analyst', 'FAL ME', '2025-09-02', '', 'Permanent', 100]]);
  var api = global.Sheets.Spreadsheets.batchUpdate;
  global.Sheets.Spreadsheets.batchUpdate = function(resource, id) {
    var res = api.call(this, resource, id);
    var uc = resource.requests[0].updateCells;
    if (uc && uc.rows[0].values[0].chipRuns) {
      mockSpreadsheet.getSheetByName('Contracts').getRange(uc.range.startRowIndex + 1, 1).setValue('Hugo Gonçalves');
    }
    return res;
  };
  try {
    var res = members.saveMemberTenure({
      member: 'other.address@gmail.com', role: 'Manager', department: 'FAL ME', startDate: '2025-01-01',
      contractType: 'Permanent', ftePercent: 100
    }, mockSpreadsheet);
    assert.strictEqual(res.success, false);
    assert.match(res.message, /Hugo Gonçalves already has a contract/);
    assert.strictEqual(tenures().length, 1, 'the chip cell is cleared again');
  } finally {
    global.Sheets.Spreadsheets.batchUpdate = api;
  }
});
