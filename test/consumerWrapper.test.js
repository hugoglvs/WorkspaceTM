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
  delete global.WorkspacePM;
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
global.freshInit = freshInit;

const code = require('../src/Code');
const consumerWrapper = require('../consumer/Code');

test('ConsumerWrapper: onOpen creates minimal streamlined Project Management menu', () => {
  consumerWrapper.onOpen();
  var ui = SpreadsheetApp.getUi();
  assert.ok(ui.menu);
  assert.strictEqual(ui.menu.name, 'Project Management');

  var captions = ui.menu.items.map(function(item) { return item.caption; });
  assert.ok(captions.includes('Initialize / Upgrade Spreadsheet'));
  assert.ok(captions.includes('Declare Presence & Leaves'));
  assert.ok(captions.includes('Members & Contracts'));
  var actionItems = ui.menu.items.filter(function(item) { return !item.separator; });
  assert.strictEqual(actionItems.length, 3);
});

test('ConsumerWrapper: surfaces actionable UI alert when WorkspacePM library is missing', () => {
  // WorkspacePM is undefined
  var res = consumerWrapper.clientSetupOrUpgrade();
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.error, 'LIBRARY_NOT_LINKED');

  var ui = SpreadsheetApp.getUi();
  assert.ok(ui.lastAlert);
  assert.strictEqual(ui.lastAlert.title, 'WorkspacePM Library Not Found');
  assert.ok(ui.lastAlert.msg.includes("The 'WorkspacePM' library is not linked"));
});

test('ConsumerWrapper: surfaces toast when library method is unsupported', () => {
  global.WorkspacePM = {}; // Linked but empty object
  var res = consumerWrapper.clientSetupOrUpgrade();
  assert.strictEqual(res.success, false);
  assert.strictEqual(res.error, 'METHOD_NOT_SUPPORTED');

  assert.ok(mockSpreadsheet.lastToast);
  assert.strictEqual(mockSpreadsheet.lastToast.title, 'Unsupported Library Action');
});

test('ConsumerWrapper: invokes setupOrUpgrade and initializes tables when library is linked', () => {
  global.WorkspacePM = code; // Link library

  var res = consumerWrapper.clientSetupOrUpgrade();
  assert.strictEqual(res.success, true);
  assert.ok(mockSpreadsheet.getSheetByName('Members'));
  assert.ok(mockSpreadsheet.getSheetByName('Tasks'));
  assert.ok(mockSpreadsheet.getSheetByName('Events'));
  assert.ok(mockSpreadsheet.lastToast);
  assert.ok(mockSpreadsheet.lastToast.title.includes('Complete'));
});

test('ConsumerWrapper: client actions forward correctly to linked library', () => {
  var invokedMethods = [];
  global.WorkspacePM = {
    init: function() { invokedMethods.push('init'); return { success: true }; },
    setupOrUpgrade: function() { invokedMethods.push('setupOrUpgrade'); return { success: true }; },
    sync: function() { invokedMethods.push('sync'); return { success: true }; },
    validateAndReport: function() { invokedMethods.push('validateAndReport'); return { success: true }; },
    compilePresence: function() { invokedMethods.push('compilePresence'); return { success: true }; },
    compilePresenceMigration: function() { invokedMethods.push('compilePresenceMigration'); return { success: true }; },
    refreshWeeklyMatrix: function() { invokedMethods.push('refreshWeeklyMatrix'); return { success: true }; },
    refreshActiveResponsibles: function() { invokedMethods.push('refreshActiveResponsibles'); return { success: true }; },
    refreshKpis: function() { invokedMethods.push('refreshKpis'); return { success: true }; },
    checkBackupConflicts: function() { invokedMethods.push('checkBackupConflicts'); return { success: true }; },
    syncTasks: function() { invokedMethods.push('syncTasks'); return { success: true }; },
    syncCalendarEvents: function() { invokedMethods.push('syncCalendarEvents'); return { success: true }; },
    showBaselineSidebar: function() { invokedMethods.push('showBaselineSidebar'); return { success: true }; },
    installLiveTaskSync: function() { invokedMethods.push('installLiveTaskSync'); return { success: true }; },
    installAllTriggers: function() { invokedMethods.push('installAllTriggers'); return { success: true }; },
    nightlySyncJob: function() { invokedMethods.push('nightlySyncJob'); return { success: true }; },
    cleanupForeignTasks: function() { invokedMethods.push('cleanupForeignTasks'); return { success: true }; },
    debugUserIdentity: function() { invokedMethods.push('debugUserIdentity'); return { success: true }; },
    onEdit: function(e) { invokedMethods.push('onEdit'); },
    installedTaskEditHandler: function(e) { invokedMethods.push('installedTaskEditHandler'); return { success: true }; },
    getSidebarInitialData: function() { invokedMethods.push('getSidebarInitialData'); return { members: [] }; },
    getMemberBaselines: function(m) { invokedMethods.push('getMemberBaselines'); return { days: {} }; },
    saveMemberBaselines: function(p) { invokedMethods.push('saveMemberBaselines'); return { success: true }; },
    submitMemberLeave: function(p) { invokedMethods.push('submitMemberLeave'); return { success: true }; }
  };

  // Test canonical forwarders (direct library method names callable by menu items)
  consumerWrapper.showBaselineSidebar();
  assert.ok(invokedMethods.includes('showBaselineSidebar'));

  consumerWrapper.init();
  assert.ok(invokedMethods.includes('init'));

  consumerWrapper.setupOrUpgrade();
  assert.ok(invokedMethods.includes('setupOrUpgrade'));

  consumerWrapper.sync();
  assert.ok(invokedMethods.includes('sync'));

  consumerWrapper.validateAndReport();
  assert.ok(invokedMethods.includes('validateAndReport'));

  consumerWrapper.compilePresence();
  assert.ok(invokedMethods.includes('compilePresence'));

  consumerWrapper.compilePresenceMigration();
  assert.ok(invokedMethods.includes('compilePresenceMigration'));

  consumerWrapper.checkBackupConflicts();
  assert.ok(invokedMethods.includes('checkBackupConflicts'));

  consumerWrapper.installAllTriggers();
  assert.ok(invokedMethods.includes('installAllTriggers'));

  consumerWrapper.nightlySyncJob();
  assert.ok(invokedMethods.includes('nightlySyncJob'));

  consumerWrapper.cleanupForeignTasks();
  assert.ok(invokedMethods.includes('cleanupForeignTasks'));

  consumerWrapper.debugUserIdentity();
  assert.ok(invokedMethods.includes('debugUserIdentity'));

  // Test client aliases
  consumerWrapper.clientSyncMembersAndTenures();
  assert.ok(invokedMethods.includes('sync'));

  consumerWrapper.clientCompilePresence();
  assert.ok(invokedMethods.includes('compilePresence'));

  consumerWrapper.clientRefreshDashboards();
  assert.ok(invokedMethods.includes('refreshWeeklyMatrix'));
  assert.ok(invokedMethods.includes('checkBackupConflicts'));
  assert.ok(invokedMethods.includes('refreshKpis'));

  consumerWrapper.clientSyncTasks();
  assert.ok(invokedMethods.includes('syncTasks'));

  consumerWrapper.clientSyncCalendarEvents();
  assert.ok(invokedMethods.includes('syncCalendarEvents'));

  consumerWrapper.clientShowSidebar();
  assert.ok(invokedMethods.includes('showBaselineSidebar'));

  consumerWrapper.clientInstallTriggers();
  assert.ok(invokedMethods.includes('installAllTriggers'));

  consumerWrapper.onEdit({ range: {} });
  assert.ok(invokedMethods.includes('onEdit'));

  consumerWrapper.installedConsumerTriggerHandler({ range: {} });
  assert.ok(invokedMethods.includes('installedTaskEditHandler'));

  consumerWrapper.getSidebarInitialData();
  assert.ok(invokedMethods.includes('getSidebarInitialData'));

  consumerWrapper.getMemberBaselines('Alice');
  assert.ok(invokedMethods.includes('getMemberBaselines'));

  consumerWrapper.saveMemberBaselines({});
  assert.ok(invokedMethods.includes('saveMemberBaselines'));

  consumerWrapper.submitMemberLeave({});
  assert.ok(invokedMethods.includes('submitMemberLeave'));
});

test('installAllTriggers: installs real-time edit trigger and nightly 2:00 AM sync trigger', () => {
  var res = code.installAllTriggers(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.installed.length, 2);

  // Re-running should detect both triggers already active
  var res2 = code.installAllTriggers(mockSpreadsheet);
  assert.strictEqual(res2.success, true);
  assert.strictEqual(res2.alreadyActive.length, 2);
});

test('nightlySyncJob: runs automated maintenance across all engines', () => {
  var res = code.nightlySyncJob(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.ok(res.steps);
});

test('setupOrUpgrade: runs freshInit, installs triggers, and runs initial maintenance sync', () => {
  var res = code.setupOrUpgrade(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.ok(res.triggers);
  assert.ok(res.initialSync);
});


test('ConsumerWrapper: Members & Contracts sidebar calls reach the library with their payload', () => {
  var calls = [];
  global.WorkspacePM = {
    showMembersSidebar: function() { calls.push(['showMembersSidebar']); },
    getMembersSidebarData: function() { calls.push(['getMembersSidebarData']); return { members: [] }; },
    saveMemberTenure: function(p) { calls.push(['saveMemberTenure', p]); return { success: true }; },
    endMemberTenure: function(p) { calls.push(['endMemberTenure', p]); return { success: true }; },
    saveMemberProfile: function(p) { calls.push(['saveMemberProfile', p]); return { success: true }; }
  };

  consumerWrapper.showMembersSidebar();
  consumerWrapper.getMembersSidebarData();
  consumerWrapper.saveMemberTenure({ member: 'A' });
  consumerWrapper.endMemberTenure({ member: 'B' });
  consumerWrapper.saveMemberProfile({ member: 'C' });

  assert.deepStrictEqual(calls, [
    ['showMembersSidebar'], ['getMembersSidebarData'],
    ['saveMemberTenure', { member: 'A' }], ['endMemberTenure', { member: 'B' }], ['saveMemberProfile', { member: 'C' }]
  ]);
  delete global.WorkspacePM;
});
