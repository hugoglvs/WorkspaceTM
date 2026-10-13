const test = require('node:test');
const assert = require('node:assert');

const { createMockSpreadsheet, setupGasGlobals, MailApp, setLegacyTasks, getLegacyTasks, legacyTaskCol } = require('./mockGas');
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
const { extractCellValue, formatDateIso, saveMemberBaselines, submitMemberLeave, getMemberBaselines, getLeaveColumnMap } = require('../src/Sidebar');
global.extractCellValue = extractCellValue;
global.formatDateIso = formatDateIso;
global.saveMemberBaselines = saveMemberBaselines;
global.submitMemberLeave = submitMemberLeave;
global.getMemberBaselines = getMemberBaselines;
global.getLeaveColumnMap = getLeaveColumnMap;

const { syncMembersAndTenures, copyCellChip } = require('../src/SyncEngine');
global.syncMembersAndTenures = syncMembersAndTenures;
global.copyCellChip = copyCellChip;

const { compilePresenceLogs, compilePresenceMigration, findEarliestHistoricalDate, getWeekdaysBetween, getAllDaysBetween, getWeekdayName } = require('../src/PresenceEngine');
global.compilePresenceLogs = compilePresenceLogs;
global.compilePresenceMigration = compilePresenceMigration;
global.findEarliestHistoricalDate = findEarliestHistoricalDate;
global.getWeekdaysBetween = getWeekdaysBetween;
global.getAllDaysBetween = getAllDaysBetween;
global.getWeekdayName = getWeekdayName;

const { checkProjectBackupConflicts, assignProjectBackup } = require('../src/ProjectBackupEngine');
global.checkProjectBackupConflicts = checkProjectBackupConflicts;
global.assignProjectBackup = assignProjectBackup;

const { onOpen, checkBackupConflicts } = require('../src/Code');
global.onOpen = onOpen;
global.checkBackupConflicts = checkBackupConflicts;

test('Half-day leaves: single day morning (AM) and afternoon (PM) leave resolution', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var leavesSheet = mockSpreadsheet.getSheetByName('Leaves');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '', 'Dev', 'Delivery', 'Active', 'Not Logged']
  ]);

  // Set baseline routine: Mon-Fri Main Office
  saveMemberBaselines({
    member: 'alice@accenture.com',
    startDate: '2026-06-01',
    days: {
      Monday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Tuesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Wednesday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Thursday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Friday: { location_am: 'Main Office', location_pm: 'Main Office' }
    }
  });

  // 1. Submit Morning Only Leave on Tuesday June 2: start AM, end AM
  var resLeave1 = submitMemberLeave({
    member: 'alice@accenture.com',
    startDate: '2026-06-02',
    startPeriod: 'AM',
    endDate: '2026-06-02',
    endPeriod: 'AM',
    leaveType: 'Vacation',
    status: 'Approved',
    asOfDate: new Date(2026, 5, 2)
  }, mockSpreadsheet);
  assert.strictEqual(resLeave1.success, true);

  // 2. Submit Afternoon Only Leave on Wednesday June 3: start PM, end PM
  var resLeave2 = submitMemberLeave({
    member: 'alice@accenture.com',
    startDate: '2026-06-03',
    startPeriod: 'PM',
    endDate: '2026-06-03',
    endPeriod: 'PM',
    leaveType: 'Sick Leave',
    status: 'Approved',
    asOfDate: new Date(2026, 5, 3)
  }, mockSpreadsheet);
  assert.strictEqual(resLeave2.success, true);

  // Compile June 1 (Mon) to June 3 (Wed) with June 3 as asOfDate
  var compRes = compilePresenceLogs(new Date(2026, 5, 1), new Date(2026, 5, 3), mockSpreadsheet, new Date(2026, 5, 3));
  assert.strictEqual(compRes.success, true);

  var logRows = logSheet.getRange(2, 1, 3, 6).getValues();
  // Monday: Full Main Office
  assert.strictEqual(logRows[0][1], '2026-06-01');
  assert.strictEqual(logRows[0][3], 'Main Office');
  assert.strictEqual(logRows[0][4], 'Main Office');

  // Tuesday: AM Vacation, PM Main Office (morning half-day)
  assert.strictEqual(logRows[1][1], '2026-06-02');
  assert.strictEqual(logRows[1][3], 'Vacation');
  assert.strictEqual(logRows[1][4], 'Main Office');

  // Wednesday: AM Main Office, PM Sick Leave (afternoon half-day)
  assert.strictEqual(logRows[2][1], '2026-06-03');
  assert.strictEqual(logRows[2][3], 'Main Office');
  assert.strictEqual(logRows[2][4], 'Sick Leave');

  // Check today_presence in tbl_members on Wednesday June 3: should reflect AM / PM split
  var todayPres = membersSheet.getRange(2, 7).getValue();
  assert.strictEqual(todayPres, 'Main Office / Sick Leave');
});

test('Half-day leaves: multi-day leave with PM start and AM end', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['bob@accenture.com', 'bob@accenture.com', '', 'Dev', 'Delivery', 'Active', 'Not Logged']
  ]);

  saveMemberBaselines({
    member: 'bob@accenture.com',
    startDate: '2026-06-01',
    days: {
      Monday: { location_am: 'Remote', location_pm: 'Remote' },
      Tuesday: { location_am: 'Remote', location_pm: 'Remote' },
      Wednesday: { location_am: 'Remote', location_pm: 'Remote' }
    }
  });

  // Leave starting Monday June 1 PM and ending Wednesday June 3 AM
  submitMemberLeave({
    member: 'bob@accenture.com',
    startDate: '2026-06-01',
    startPeriod: 'PM',
    endDate: '2026-06-03',
    endPeriod: 'AM',
    leaveType: 'Vacation',
    status: 'Approved',
    asOfDate: new Date(2026, 5, 3)
  }, mockSpreadsheet);

  compilePresenceLogs(new Date(2026, 5, 1), new Date(2026, 5, 3), mockSpreadsheet, new Date(2026, 5, 3));

  var logRows = logSheet.getRange(2, 1, 3, 6).getValues();
  // Monday: AM Remote, PM Vacation
  assert.strictEqual(logRows[0][1], '2026-06-01');
  assert.strictEqual(logRows[0][3], 'Remote');
  assert.strictEqual(logRows[0][4], 'Vacation');

  // Tuesday: full day Vacation
  assert.strictEqual(logRows[1][1], '2026-06-02');
  assert.strictEqual(logRows[1][3], 'Vacation');
  assert.strictEqual(logRows[1][4], 'Vacation');

  // Wednesday: AM Vacation, PM Remote
  assert.strictEqual(logRows[2][1], '2026-06-03');
  assert.strictEqual(logRows[2][3], 'Vacation');
  assert.strictEqual(logRows[2][4], 'Remote');
});

test('Weekend working days: weekend baselines default to Off and compile when configured with working location', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var baselinesSheet = mockSpreadsheet.getSheetByName('Routines');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  membersSheet.getRange(2, 1, 2, 7).setValues([
    ['alice@accenture.com', 'alice@accenture.com', '', 'Lead', 'Delivery', 'Active', 'Not Logged'],
    ['carol@accenture.com', 'carol@accenture.com', '', 'OnCall', 'Support', 'Active', 'Not Logged']
  ]);

  // Alice has standard Mon-Fri (no weekend working days)
  saveMemberBaselines({
    member: 'alice@accenture.com',
    startDate: '2026-06-01',
    days: {
      Friday: { location_am: 'Main Office', location_pm: 'Main Office' }
    }
  });

  // Carol works Saturday: Remote / Remote
  saveMemberBaselines({
    member: 'carol@accenture.com',
    startDate: '2026-06-01',
    days: {
      Friday: { location_am: 'Main Office', location_pm: 'Main Office' },
      Saturday: { location_am: 'Remote', location_pm: 'Remote' }
    }
  });

  // Verify baseline saved for Carol includes Saturday
  var carolBaselines = getMemberBaselines('carol@accenture.com');
  assert.strictEqual(carolBaselines.days['Saturday'].location_am, 'Remote');

  // Compile from Friday June 5 to Saturday June 6 as of Saturday June 6
  var asOfSaturday = new Date(2026, 5, 6);
  var compRes = compilePresenceLogs(new Date(2026, 5, 5), new Date(2026, 5, 6), mockSpreadsheet, asOfSaturday);
  assert.strictEqual(compRes.success, true);

  // Friday: 2 logs (Alice, Carol)
  // Saturday: strictly 1 log (Carol only, since Alice is Off on Saturday)
  var allLogs = logSheet.getRange(2, 1, 10, 6).getValues();
  var saturdayLogs = allLogs.filter(r => r[1] === '2026-06-06');
  assert.strictEqual(saturdayLogs.length, 1);
  assert.strictEqual(saturdayLogs[0][2], 'carol@accenture.com');
  assert.strictEqual(saturdayLogs[0][3], 'Remote');

  // Check today_presence in tbl_members on Saturday June 6:
  // Alice: 'Weekend'
  // Carol: 'Remote'
  assert.strictEqual(membersSheet.getRange(2, 7).getValue(), 'Weekend');
  assert.strictEqual(membersSheet.getRange(3, 7).getValue(), 'Remote');
});

test('Presence logs: direct People Chip writing without text replacement flicker', () => {
  freshInit(mockSpreadsheet);
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var logSheet = mockSpreadsheet.getSheetByName('Presence Log');

  var chipMock = {
    getName: function() { return 'Hugo Gonçalves'; },
    getEmail: function() { return 'hugo.goncalves@accenture.com'; }
  };

  membersSheet.getRange(2, 1, 1, 7).setValues([
    [chipMock, 'hugo.goncalves@accenture.com', '', 'Lead', 'Delivery', 'Active', 'Not Logged']
  ]);

  saveMemberBaselines({
    member: 'Hugo Gonçalves',
    startDate: '2026-06-01',
    days: {
      Monday: { location_am: 'Main Office', location_pm: 'Main Office' }
    }
  });

  var res = compilePresenceLogs(new Date(2026, 5, 1), new Date(2026, 5, 1), mockSpreadsheet, new Date(2026, 5, 1));
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.createdCount, 1);

  // Column 3 in tbl_presence_log receives chipMock directly
  var memberCell = logSheet.getRange(2, 3).getValue();
  assert.strictEqual(memberCell, chipMock);
});

test('Project lead and backup simultaneous absence conflict detection: red alert, email, and task creation', () => {
  freshInit(mockSpreadsheet);
  global.MailApp.clearSentEmails();

  var projectsSheet = mockSpreadsheet.getSheetByName('Projects');
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var leavesSheet = mockSpreadsheet.getSheetByName('Leaves');
  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  // Setup members: Alice (Lead), Bob (Backup), Dave (Manager)
  membersSheet.getRange(2, 1, 3, 7).setValues([
    ['Alice', 'alice@accenture.com', '', 'Tech Lead', 'Engineering', 'Active', 'Not Logged'],
    ['Bob', 'bob@accenture.com', '', 'Senior Dev', 'Engineering', 'Active', 'Not Logged'],
    ['Dave', 'dave.manager@accenture.com', '', 'Engineering Manager', 'Engineering', 'Active', 'Not Logged']
  ]);

  // Setup active project with Alice as Lead and Bob as Backup
  projectsSheet.getRange(2, 1, 1, 7).setValues([
    ['PRJ-001', 'Phoenix Initiative', 'Active', '2026-06-01', '2026-12-31', 'Alice', 'Bob']
  ]);

  // Alice is on approved leave from June 10 to June 20
  submitMemberLeave({
    member: 'Alice',
    startDate: '2026-06-10',
    endDate: '2026-06-20',
    leaveType: 'Vacation',
    status: 'Approved',
    asOfDate: new Date(2026, 5, 1)
  }, mockSpreadsheet);

  // Bob is on approved leave from June 15 to June 25 (Overlaps June 15 to June 20!)
  submitMemberLeave({
    member: 'Bob',
    startDate: '2026-06-15',
    endDate: '2026-06-25',
    leaveType: 'Vacation',
    status: 'Approved',
    asOfDate: new Date(2026, 5, 1)
  }, mockSpreadsheet);

  // Run conflict check
  var auditRes = checkProjectBackupConflicts(mockSpreadsheet, {
    asOfDateParam: new Date(2026, 5, 1) // June 1, Alice not yet on leave
  });

  assert.strictEqual(auditRes.success, true);
  assert.strictEqual(auditRes.conflictsFound, 1);
  assert.strictEqual(auditRes.conflicts[0].projectId, 'PRJ-001');
  assert.strictEqual(auditRes.conflicts[0].conflictStart, '2026-06-15');
  assert.strictEqual(auditRes.conflicts[0].conflictEnd, '2026-06-19'); // Saturday 20 is not a working day

  // 1. Verify project row formatting is alert red (#F8D7DA, #721C24, bold)
  var projectRowRange = projectsSheet.getRange(2, 1, 1, 8);
  assert.strictEqual(projectRowRange.getBackground(), '#F8D7DA');
  assert.strictEqual(projectRowRange.getFontColor(), '#721C24');
  assert.strictEqual(projectRowRange.getFontWeight(), 'bold');

  // 2. Verify email sent via MailApp to lead and manager
  var sentEmails = global.MailApp.getSentEmails();
  assert.strictEqual(sentEmails.length, 1);
  assert.ok(sentEmails[0].to.includes('alice@accenture.com'));
  assert.ok(sentEmails[0].to.includes('dave.manager@accenture.com'));
  assert.ok(sentEmails[0].subject.includes('Phoenix Initiative'));

  // 3. Verify interim backup task created in tbl_tasks assigned to Alice (Lead, not yet absent)
  var taskRows = getLegacyTasks(tasksSheet, 2, 1, 8);
  assert.strictEqual(taskRows[0][1], 'PRJ-001');
  assert.strictEqual(taskRows[0][2], 'Alice');
  assert.strictEqual(taskRows[0][3], 'Organize interim backup for project Phoenix Initiative');
  assert.strictEqual(taskRows[0][5], '2026-06-15');
  assert.strictEqual(taskRows[0][6], 'Todo');
});

test('Project conflict mitigation via tbl_backups: clears alert formatting', () => {
  freshInit(mockSpreadsheet);
  var projectsSheet = mockSpreadsheet.getSheetByName('Projects');
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var backupsSheet = mockSpreadsheet.getSheetByName('Backups');

  membersSheet.getRange(2, 1, 4, 7).setValues([
    ['Alice', 'alice@accenture.com', '', 'Lead', 'Eng', 'Active', 'Not Logged'],
    ['Bob', 'bob@accenture.com', '', 'Senior Dev', 'Eng', 'Active', 'Not Logged'],
    ['Charlie', 'charlie@accenture.com', '', 'Staff Dev', 'Eng', 'Active', 'Not Logged'],
    ['Dave', 'dave.manager@accenture.com', '', 'Manager', 'Eng', 'Active', 'Not Logged']
  ]);

  projectsSheet.getRange(2, 1, 1, 7).setValues([
    ['PRJ-002', 'Titan Project', 'Active', '2026-06-01', '2026-12-31', 'Alice', 'Bob']
  ]);

  // Alice and Bob overlapping leaves June 15 to June 20
  submitMemberLeave({
    member: 'Alice',
    startDate: '2026-06-15',
    endDate: '2026-06-20',
    leaveType: 'Vacation',
    status: 'Approved'
  }, mockSpreadsheet);

  submitMemberLeave({
    member: 'Bob',
    startDate: '2026-06-15',
    endDate: '2026-06-20',
    leaveType: 'Vacation',
    status: 'Approved'
  }, mockSpreadsheet);

  // Initial check (as of June 1) triggers conflict
  var res1 = checkProjectBackupConflicts(mockSpreadsheet, { asOfDateParam: new Date(2026, 5, 1) });
  assert.strictEqual(res1.conflictsFound, 1);
  assert.strictEqual(projectsSheet.getRange(2, 1, 1, 8).getBackground(), '#F8D7DA');

  // Assign Charlie as backup in tbl_backups covering the conflict period
  var bkpRes = assignProjectBackup({
    projectId: 'PRJ-002',
    primaryMember: 'Alice',
    backupMember: 'Charlie',
    startDate: '2026-06-15',
    endDate: '2026-06-20',
    status: 'Active',
    notes: 'Covering simultaneous absence of Alice and Bob'
  }, mockSpreadsheet);

  assert.strictEqual(bkpRes.success, true);

  // Row 2 is Bob's permanent backup (mirrored from tbl_projects), row 3 Charlie's interim backup
  var permanentRow = backupsSheet.getRange(2, 1, 1, 8).getValues()[0];
  assert.strictEqual(permanentRow[3], 'Bob');
  assert.strictEqual(permanentRow[5], '', 'permanent backup has no end date');
  var bkpRow = backupsSheet.getRange(3, 1, 1, 8).getValues()[0];
  assert.strictEqual(bkpRow[1], 'PRJ-002');
  assert.strictEqual(bkpRow[2], 'Alice');
  assert.strictEqual(bkpRow[3], 'Charlie');
  assert.strictEqual(bkpRow[6], 'Active');

  // Verify project row formatting is restored to normal white (#FFFFFF)
  assert.strictEqual(projectsSheet.getRange(2, 1, 1, 8).getBackground(), '#FFFFFF');
  assert.strictEqual(projectsSheet.getRange(2, 1, 1, 8).getFontColor(), '#000000');
  assert.strictEqual(projectsSheet.getRange(2, 1, 1, 8).getFontWeight(), 'normal');
});

test('Project conflict task assignment: assigns backup member when lead is already absent', () => {
  freshInit(mockSpreadsheet);
  var projectsSheet = mockSpreadsheet.getSheetByName('Projects');
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  membersSheet.getRange(2, 1, 3, 7).setValues([
    ['Alice', 'alice@accenture.com', '', 'Lead', 'Eng', 'Active', 'Vacation'], // Alice is ALREADY on vacation today
    ['Bob', 'bob@accenture.com', '', 'Senior Dev', 'Eng', 'Active', 'Main Office'],
    ['Dave', 'dave@accenture.com', '', 'Manager', 'Eng', 'Active', 'Main Office']
  ]);

  projectsSheet.getRange(2, 1, 1, 7).setValues([
    ['PRJ-003', 'Delta System', 'Active', '2026-06-01', '2026-12-31', 'Alice', 'Bob']
  ]);

  // Alice is on leave June 1 to June 10
  submitMemberLeave({
    member: 'Alice',
    startDate: '2026-06-01',
    endDate: '2026-06-10',
    leaveType: 'Vacation',
    status: 'Approved'
  }, mockSpreadsheet);

  // Bob is on leave June 5 to June 8
  submitMemberLeave({
    member: 'Bob',
    startDate: '2026-06-05',
    endDate: '2026-06-08',
    leaveType: 'Vacation',
    status: 'Approved'
  }, mockSpreadsheet);

  // As of June 6: Alice is already absent! Task should be assigned to Bob!
  var auditRes = checkProjectBackupConflicts(mockSpreadsheet, {
    asOfDateParam: new Date(2026, 5, 6)
  });

  assert.strictEqual(auditRes.conflictsFound, 1);

  var taskRows = getLegacyTasks(tasksSheet, 2, 1, 8);
  assert.strictEqual(taskRows[0][1], 'PRJ-003');
  // Since Alice is already absent, Bob is assigned
  assert.strictEqual(taskRows[0][2], 'Bob');
});

test('onOpen registers Check Project Backup Conflicts in menu and wrapper works', () => {
  freshInit(mockSpreadsheet);
  onOpen();

  var res = checkBackupConflicts(mockSpreadsheet);
  assert.strictEqual(res.success, true);
});
