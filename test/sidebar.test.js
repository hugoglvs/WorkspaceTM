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

const { freshInit } = require('../src/Init');
const syncEngine = require('../src/SyncEngine');
global.copyCellChip = syncEngine.copyCellChip;
global.syncBaselinesPeopleChips = syncEngine.syncBaselinesPeopleChips;
const sidebar = require('../src/Sidebar');
global.checkActiveUserStatus = sidebar.checkActiveUserStatus;
global.auditUserAndMembersIdentity = sidebar.auditUserAndMembersIdentity;
global.debugUserIdentity = sidebar.debugUserIdentity;
const { onOpen } = require('../src/Code');
const { compilePresenceLogs } = require('../src/PresenceEngine');
global.compilePresenceLogs = compilePresenceLogs;

test('Sidebar initial data should extract members from tbl_members', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  // Seed sample members in Column 1 (member / People chip)
  membersSheet.getRange(2, 1, 2, 1).setValues([
    ['alice@example.com'],
    ['bob@example.com']
  ]);

  var data = sidebar.getSidebarInitialData();
  assert.strictEqual(data.members.length, 2);
  assert.strictEqual(data.members[0].member, 'alice@example.com');
  assert.strictEqual(data.members[1].member, 'bob@example.com');
  assert.strictEqual(data.activeUserEmail, 'alice@example.com');
  assert.ok(data.locationOptions.includes('OneRoof'));
  assert.ok(data.locationOptions.includes('Remote'));
});

test('saveMemberBaselines and getMemberBaselines should write and retrieve baseline routine', () => {
  freshInit(mockSpreadsheet);

  var payload = {
    member: 'alice@example.com',
    startDate: '2026-01-05',
    days: {
      Monday: { location_am: 'Remote', location_pm: 'Remote' },
      Tuesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Wednesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Thursday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Friday: { location_am: 'Remote', location_pm: 'Remote' }
    }
  };

  var res = sidebar.saveMemberBaselines(payload);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.savedRows, 5);
  assert.strictEqual(res.member, 'alice@example.com');

  var retrieved = sidebar.getMemberBaselines('alice@example.com');
  assert.strictEqual(retrieved.startDate, '2026-01-05');
  assert.strictEqual(retrieved.days.Monday.location_am, 'Remote');
  assert.strictEqual(retrieved.days.Monday.location_pm, 'Remote');
  assert.strictEqual(retrieved.days.Tuesday.location_am, 'Main Office');
  assert.strictEqual(retrieved.days.Friday.location_am, 'Remote');
});

test('showBaselineSidebar and onOpen should create UI sidebar and menus without errors', () => {
  onOpen();
  var menu = SpreadsheetApp.getUi().menu;
  assert.strictEqual(menu.name, 'Project Management');
  assert.strictEqual(menu.items.length, 14); // init, sync, validate, compilePresence, compilePresenceMigration, refreshKpis, checkBackupConflicts, syncTasks, syncCalendarEvents, installLiveTaskSync, cleanupForeignTasks, separator, sidebar, debugUserIdentity

  sidebar.showBaselineSidebar();
  var shown = SpreadsheetApp.getUi().shownSidebar;
  assert.ok(shown);
  assert.strictEqual(shown.filename, 'BaselineSidebar');
  assert.strictEqual(shown.title, 'Presence & Leaves');
});

test('Sidebar should write baselines directly starting at row 2 with People Chip identity', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');

  // Seed member in column 1 (People chip identity)
  membersSheet.getRange(2, 1, 1, 1).setValues([
    ['charlie@peanuts.com']
  ]);

  var data = sidebar.getSidebarInitialData();
  assert.strictEqual(data.members.length, 1);
  assert.strictEqual(data.members[0].member, 'charlie@peanuts.com');

  // Save baseline using the member People Chip
  var saveRes = sidebar.saveMemberBaselines({
    member: 'charlie@peanuts.com',
    startDate: '2026-02-01',
    days: {
      Monday: { location_am: 'Remote', location_pm: 'Remote' },
      Tuesday: { location_am: 'Remote', location_pm: 'Remote' },
      Wednesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Thursday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Friday: { location_am: 'Remote', location_pm: 'Remote' }
    }
  });

  assert.strictEqual(saveRes.success, true);

  // Verify rows are written starting at row 2 in tbl_presence_baselines
  var baselinesSheet = mockSpreadsheet.getSheetByName('Routines');
  var writtenValues = baselinesSheet.getRange(2, 1, 5, 7).getValues();
  assert.strictEqual(writtenValues.length, 5);
  assert.strictEqual(writtenValues[0][1], 'charlie@peanuts.com'); // member column
  assert.strictEqual(writtenValues[0][4], 'Monday');
  assert.strictEqual(writtenValues[4][4], 'Friday');
});

test('checkBaselineDateOverlap: should detect overlaps between ranges including open-ended ranges', () => {
  // Overlapping open-ended
  assert.strictEqual(sidebar.checkBaselineDateOverlap('2026-01-01', '', '2026-06-01', ''), true);
  // Overlapping bounded
  assert.strictEqual(sidebar.checkBaselineDateOverlap('2026-01-01', '2026-08-01', '2026-06-01', '2026-12-31'), true);
  // Adjacent: s2 == e1 (non-overlapping because previous ended on the start date of next)
  assert.strictEqual(sidebar.checkBaselineDateOverlap('2026-01-01', '2026-06-01', '2026-06-01', ''), true, 'end dates are inclusive');
  // Completely separated: s2 > e1
  assert.strictEqual(sidebar.checkBaselineDateOverlap('2026-01-01', '2026-05-31', '2026-06-01', '2026-12-31'), false);
});

test('saveMemberBaselines: should not override previous baseline, should detect overlap and adjust end date upon confirmation', () => {
  freshInit(mockSpreadsheet);
  var baselinesSheet = mockSpreadsheet.getSheetByName('Routines');

  // 1. Save Initial Baseline for Alice (open-ended starting 2026-01-01)
  var res1 = sidebar.saveMemberBaselines({
    member: 'alice@accenture.com',
    startDate: '2026-01-01',
    days: {
      Monday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Tuesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Wednesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Thursday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Friday: { location_am: 'Main Office', location_pm: 'Main Office' }
    }
  });
  assert.strictEqual(res1.success, true);
  assert.strictEqual(res1.savedRows, 5);

  // 2. Submit a second baseline starting 2026-06-01 WITHOUT confirmation -> must detect overlap and prompt!
  var res2 = sidebar.saveMemberBaselines({
    member: 'alice@accenture.com',
    startDate: '2026-06-01',
    days: {
      Monday: { location_am: 'Remote', location_pm: 'Remote' },
      Tuesday: { location_am: 'Remote', location_pm: 'Remote' },
      Wednesday: { location_am: 'Remote', location_pm: 'Remote' },
      Thursday: { location_am: 'Remote', location_pm: 'Remote' },
      Friday: { location_am: 'Remote', location_pm: 'Remote' }
    }
  });
  assert.strictEqual(res2.success, false);
  assert.strictEqual(res2.overlap, true);
  assert.ok(res2.message.includes('overlaps 1 existing routine'));
  assert.ok(res2.message.includes('will end on 2026-05-31'), res2.message);

  // Verify sheet still only has the original 5 rows
  assert.strictEqual(baselinesSheet.getRange(2, 1, 5, 7).getValues().length, 5);

  // 3. Submit with adjustPreviousEnd: true -> must update previous baseline end date AND append 5 new rows!
  var res3 = sidebar.saveMemberBaselines({
    member: 'alice@accenture.com',
    startDate: '2026-06-01',
    adjustPreviousEnd: true,
    days: {
      Monday: { location_am: 'Remote', location_pm: 'Remote' },
      Tuesday: { location_am: 'Remote', location_pm: 'Remote' },
      Wednesday: { location_am: 'Remote', location_pm: 'Remote' },
      Thursday: { location_am: 'Remote', location_pm: 'Remote' },
      Friday: { location_am: 'Remote', location_pm: 'Remote' }
    }
  });
  assert.strictEqual(res3.success, true);
  assert.strictEqual(res3.overlapAdjusted, true);

  // 4. Verify tbl_presence_baselines now has 10 total rows (previous baseline NOT deleted!)
  var allRows = baselinesSheet.getRange(2, 1, 10, 7).getValues();
  assert.strictEqual(allRows.length, 10);

  // First 5 rows (original baseline) now end the day before the new one (inclusive end dates)
  for (var r = 0; r < 5; r++) {
    assert.strictEqual(allRows[r][2], '2026-01-01'); // start_date
    assert.strictEqual(allRows[r][3], '2026-05-31'); // end_date updated!
    assert.strictEqual(allRows[r][5], 'Main Office');
  }

  // Next 5 rows (new baseline) start on 2026-06-01!
  for (var nr = 5; nr < 10; nr++) {
    assert.strictEqual(allRows[nr][2], '2026-06-01'); // start_date
    assert.strictEqual(allRows[nr][3], ''); // end_date open
    assert.strictEqual(allRows[nr][5], 'Remote');
  }

  // 5. getMemberBaselines returns the latest active baseline (Remote)
  var latest = sidebar.getMemberBaselines('alice@accenture.com');
  assert.strictEqual(latest.startDate, '2026-06-01');
  assert.strictEqual(latest.days.Monday.location_am, 'Remote');
  assert.strictEqual(latest.totalBaselinesCount, 2);
});

test('submitMemberLeave: should validate inputs and reject invalid date ranges', () => {
  freshInit(mockSpreadsheet);

  var res1 = sidebar.submitMemberLeave({
    member: '',
    startDate: '2026-06-01',
    endDate: '2026-06-05',
    leaveType: 'Vacation'
  });
  assert.strictEqual(res1.success, false);
  assert.ok(res1.message.includes('Missing required'));

  var res2 = sidebar.submitMemberLeave({
    member: 'alice@accenture.com',
    startDate: '2026-06-10',
    endDate: '2026-06-05',
    leaveType: 'Vacation'
  });
  assert.strictEqual(res2.success, false);
  assert.ok(res2.message.includes('Start date cannot be after end date'));
});

test('submitMemberLeave: should append leave row to tbl_leaves and compile presence logs', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var leavesSheet = mockSpreadsheet.getSheetByName('Leaves');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);

  var payload = {
    member: 'alice@accenture.com',
    startDate: '2026-06-01',
    endDate: '2026-06-03',
    leaveType: 'Sick Leave',
    status: 'Approved'
  };

  var res = sidebar.submitMemberLeave(payload, mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.member, 'alice@accenture.com');
  assert.strictEqual(res.leaveType, 'Sick Leave');
  assert.strictEqual(res.status, 'Approved');
  assert.strictEqual(res.compiled, true);

  // Verify row in tbl_leaves (8 columns: leave_id, member, start_date, start_period, end_date, end_period, leave_type, status)
  var leaveRow = leavesSheet.getRange(2, 1, 1, 8).getValues()[0];
  assert.ok(String(leaveRow[0]).startsWith('leave_'));
  assert.strictEqual(leaveRow[1], 'alice@accenture.com');
  assert.strictEqual(leaveRow[2], '2026-06-01');
  assert.strictEqual(leaveRow[3], 'AM');
  assert.strictEqual(leaveRow[4], '2026-06-03');
  assert.strictEqual(leaveRow[5], 'PM');
  assert.strictEqual(leaveRow[6], 'Sick Leave');
  assert.strictEqual(leaveRow[7], 'Approved');
});

test('checkActiveUserStatus: should detect missing baseline and incomplete profile', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');

  // Alice is in tbl_members with missing role and department, and NO baseline routine
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@example.com', 'alice@example.com', '', '', '', 'Active', 'Not Logged']
  ]);

  var status = sidebar.checkActiveUserStatus(mockSpreadsheet, true);
  assert.strictEqual(status.email, 'alice@example.com');
  assert.strictEqual(status.memberFound, true);
  assert.strictEqual(status.hasBaseline, false);
  assert.ok(status.missingFields.includes('Role'));
  assert.ok(status.missingFields.includes('Department'));
  assert.ok(status.warnings.length > 0);

  // Check toast was fired
  assert.ok(mockSpreadsheet.lastToast);
  assert.ok(mockSpreadsheet.lastToast.title.includes('Action Needed'));
  assert.ok(mockSpreadsheet.lastToast.msg.includes('Weekly presence routine is not set up'));
});

test('checkActiveUserStatus: should report clean status when member is fully configured', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@example.com', 'alice@example.com', '+33123456789', 'Lead', 'Delivery', 'Active', 'Main Office']
  ]);

  // Set baseline for Alice
  sidebar.saveMemberBaselines({
    member: 'alice@example.com',
    startDate: '2026-01-01',
    days: {
      Monday: { location_am: 'Remote', location_pm: 'Remote' },
      Tuesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Wednesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Thursday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Friday: { location_am: 'Remote', location_pm: 'Remote' }
    }
  });

  mockSpreadsheet.lastToast = null;
  var status = sidebar.checkActiveUserStatus(mockSpreadsheet, true);
  assert.strictEqual(status.email, 'alice@example.com');
  assert.strictEqual(status.memberFound, true);
  assert.strictEqual(status.hasBaseline, true);
  assert.strictEqual(status.baselinesCount, 5);
  assert.strictEqual(status.missingFields.length, 0);
  assert.strictEqual(status.warnings.length, 0);
  assert.strictEqual(status.todayPresence, 'Main Office');
  assert.strictEqual(mockSpreadsheet.lastToast, null); // No toast when all clean!
});

test('getSidebarInitialData: should include userChecklist for active user', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@example.com', 'alice@example.com', '', 'Developer', 'Engineering', 'Active', 'Remote']
  ]);

  var data = sidebar.getSidebarInitialData();
  assert.ok(data.userChecklist);
  assert.strictEqual(data.userChecklist.email, 'alice@example.com');
  assert.strictEqual(data.userChecklist.hasBaseline, false); // No baseline yet
  assert.strictEqual(data.userChecklist.todayPresence, 'Remote');
});

test('isUserMemberMatch: should match People chip with mailto link, chip object, or display name against corporate email without accenture_email', () => {
  // 1. Chip object with getEmail()
  var chipObj = {
    getName: function() { return 'Hugo Gonçalves'; },
    getEmail: function() { return 'hugo.goncalves@accenture.com'; }
  };
  assert.strictEqual(sidebar.isUserMemberMatch(chipObj, null, 'hugo.goncalves@accenture.com'), true);

  // 2. Chip with RichText mailto: link
  var mockRichText = {
    getLinkUrl: function() { return 'mailto:hugo.goncalves@accenture.com'; },
    getRuns: function() { return []; }
  };
  assert.strictEqual(sidebar.isUserMemberMatch('Hugo Gonçalves', mockRichText, 'hugo.goncalves@accenture.com'), true);

  // 3. Display name with diacritics / accents matching email prefix
  assert.strictEqual(sidebar.isUserMemberMatch('Hugo Gonçalves', null, 'hugo.goncalves@accenture.com'), true);
  assert.strictEqual(sidebar.isUserMemberMatch('Hugo Goncalves', null, 'hugo.goncalves@accenture.com'), true);

  // 4. Non-matching member
  assert.strictEqual(sidebar.isUserMemberMatch('Bob Marley', null, 'hugo.goncalves@accenture.com'), false);
});

test('checkActiveUserStatus: should match member with display name People chip when accenture_email is blank', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');

  // Change Session user to hugo.goncalves@accenture.com
  var originalSession = global.Session;
  global.Session = {
    getActiveUser: function() {
      return { getEmail: function() { return 'hugo.goncalves@accenture.com'; } };
    }
  };

  try {
    // Column 1 is People Chip with display name 'Hugo Gonçalves', column 2 is blank!
    membersSheet.getRange(2, 1, 1, 7).setValues([
      ['Hugo Gonçalves', '', '+33123456789', 'Lead', 'Delivery', 'Active', 'Remote']
    ]);

    var status = sidebar.checkActiveUserStatus(mockSpreadsheet, false);
    assert.strictEqual(status.email, 'hugo.goncalves@accenture.com');
    assert.strictEqual(status.memberFound, true); // Must be found!
    assert.strictEqual(status.name, 'Hugo Gonçalves');
    assert.strictEqual(status.todayPresence, 'Remote');
  } finally {
    global.Session = originalSession;
  }
});

test('onOpen should log active user and members emails to Logger.log', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 2, 7).setValues([
    ['alice@example.com', '', '', 'Lead', 'Engineering', 'Active', 'Main Office'],
    ['bob@example.com', '', '', 'Engineer', 'Engineering', 'Active', 'Remote']
  ]);

  Logger.clear();
  onOpen();

  var logs = Logger.getLogs();
  assert.ok(logs.length > 0, 'Logger should have logged messages on onOpen');
  var logText = logs.join('\n');
  assert.ok(logText.includes('=== [User & Member Identity Audit] ==='));
  assert.ok(logText.includes("Active User: 'alice@example.com'"));
  assert.ok(logText.includes("Name: 'alice@example.com'"));
  assert.ok(logText.includes("Name: 'bob@example.com'"));
  assert.ok(logText.includes("Match: YES"));
});

test('debugUserIdentity should display UI alert dialog with audit logs', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@example.com', '', '', 'Lead', 'Engineering', 'Active', 'Main Office']
  ]);

  var audit = sidebar.debugUserIdentity(mockSpreadsheet);
  assert.ok(audit);
  assert.strictEqual(audit.activeEmail, 'alice@example.com');
  assert.strictEqual(audit.matchingMember.name, 'alice@example.com');

  var lastAlert = SpreadsheetApp.getUi().lastAlert;
  assert.ok(lastAlert);
  assert.strictEqual(lastAlert.title, 'User & Member Identity Diagnostics');
  assert.ok(lastAlert.msg.includes('Active User'));
  assert.ok(lastAlert.msg.includes('Members'));
});

test('auditUserAndMembersIdentity should handle empty Session user email gracefully', () => {
  freshInit(mockSpreadsheet);
  var originalSession = global.Session;
  global.Session = {
    getActiveUser: function() { return { getEmail: function() { return ''; } }; },
    getEffectiveUser: function() { return { getEmail: function() { return 'admin@example.com'; } }; }
  };

  try {
    var audit = sidebar.auditUserAndMembersIdentity(mockSpreadsheet);
    assert.strictEqual(audit.activeEmail, '');
    assert.strictEqual(audit.matchingMember, null);
    var text = audit.logLines.join('\n');
    assert.ok(text.includes('WARNING: Active user email is empty'));
  } finally {
    global.Session = originalSession;
  }
});

test('extractEmailFromChip should handle personProperties, complex mailto, and query strings', () => {
  // Person properties
  var personObj = { personProperties: { email: 'developer@google.com' } };
  assert.strictEqual(sidebar.extractEmailFromChip(personObj), 'developer@google.com');

  // Mailto with query params
  var richTextWithParams = {
    getLinkUrl: function() { return 'mailto:test.user@domain.com?subject=Inquiry'; },
    getRuns: function() { return []; }
  };
  assert.strictEqual(sidebar.extractEmailFromChip('Test User', richTextWithParams), 'test.user@domain.com');

  // Contact URL with query or path email
  var richTextContactUrl = {
    getLinkUrl: function() { return 'https://mail.google.com/mail/?to=colleague@company.com'; },
    getRuns: function() { return []; }
  };
  assert.strictEqual(sidebar.extractEmailFromChip('Colleague', richTextContactUrl), 'colleague@company.com');
});

test('explainUserMemberMatch should match hugoglvs42@gmail.com with Hugo Gonçalves via name contraction', () => {
  var res = sidebar.explainUserMemberMatch('Hugo Gonçalves', null, 'hugoglvs42@gmail.com');
  assert.strictEqual(res.match, true);
  assert.ok(res.reason.includes('Name contraction match') || res.reason.includes('hugo'));
});

test('fetchMemberChipEmails should extract email from chipRuns metadata', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  // Row 2 is chip object with personProperties
  membersSheet.grid[2] = {
    1: {
      name: 'Hugo Gonçalves',
      personProperties: { email: 'hugoglvs42@gmail.com', name: 'Hugo Gonçalves' }
    }
  };

  var emails = sidebar.fetchMemberChipEmails(mockSpreadsheet.getId(), 'Members', 2, 1);
  assert.strictEqual(emails[2], 'hugoglvs42@gmail.com');
});

test('checkActiveUserStatus should match active user hugoglvs42@gmail.com with Hugo Gonçalves in tbl_members', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Hugo Gonçalves', '', '', 'Lead', 'Delivery', 'Active', 'Remote']
  ]);

  var originalSession = global.Session;
  global.Session = {
    getActiveUser: function() { return { getEmail: function() { return 'hugoglvs42@gmail.com'; } }; },
    getEffectiveUser: function() { return { getEmail: function() { return 'hugoglvs42@gmail.com'; } }; }
  };

  try {
    var status = sidebar.checkActiveUserStatus(mockSpreadsheet, false);
    assert.strictEqual(status.email, 'hugoglvs42@gmail.com');
    assert.strictEqual(status.memberFound, true);
    assert.strictEqual(status.name, 'Hugo Gonçalves');
    assert.strictEqual(status.todayPresence, 'Remote');
  } finally {
    global.Session = originalSession;
  }
});

test('syncBaselinesPeopleChips: should copy People Chips from tbl_members to tbl_presence_baselines', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var baselinesSheet = mockSpreadsheet.getSheetByName('Routines');

  // Member in tbl_members with display name
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Hugo Gonçalves', 'hugoglvs42@gmail.com', '', 'Engineer', 'Core', 'Active', 'Not Logged']
  ]);

  // Baselines entered with email or plain text
  baselinesSheet.getRange(2, 1, 5, 7).setValues([
    ['b1', 'hugoglvs42@gmail.com', '2026-01-05', '', 'Monday', 'Remote', 'Remote'],
    ['b2', 'hugoglvs42@gmail.com', '2026-01-05', '', 'Tuesday', 'Main Office', 'Main Office'],
    ['b3', 'hugoglvs42@gmail.com', '2026-01-05', '', 'Wednesday', 'Main Office', 'Main Office'],
    ['b4', 'hugoglvs42@gmail.com', '2026-01-05', '', 'Thursday', 'Main Office', 'Main Office'],
    ['b5', 'hugoglvs42@gmail.com', '2026-01-05', '', 'Friday', 'Remote', 'Remote']
  ]);

  // Run syncBaselinesPeopleChips
  var count = syncEngine.syncBaselinesPeopleChips(mockSpreadsheet);
  assert.strictEqual(count, 5);

  // Column 2 should now have the chip / canonical identity from tbl_members
  var rowVals = baselinesSheet.getRange(2, 2, 5, 1).getValues();
  assert.strictEqual(rowVals[0][0], 'Hugo Gonçalves');
});


