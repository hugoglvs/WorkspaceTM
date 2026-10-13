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

const kpiEngine = require('../src/KpiEngine');
global.KpiEngine = kpiEngine;
global.renderPresenceFacts = kpiEngine.renderPresenceFacts;
global.refreshKpis = kpiEngine.refreshKpis;

const taskSyncService = require('../src/TaskSyncService');
global.TaskSyncService = taskSyncService;
global.syncGoogleTasks = taskSyncService.syncGoogleTasks;
global.syncTasks = taskSyncService.syncTasks;

const code = require('../src/Code');

test('getTasksListId & setTasksListId: manages Google Tasks List ID via Document Properties', () => {
  // Initial state should discover or fallback
  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  assert.ok(listId);

  // Set explicit ID
  taskSyncService.setTasksListId('custom-team-list-999');
  var retrievedId = taskSyncService.getTasksListId(mockSpreadsheet);
  assert.strictEqual(retrievedId, 'custom-team-list-999');
});

test('formatTaskDueRfc3339: converts date strings and Date objects to RFC 3339', () => {
  var dStr = taskSyncService.formatTaskDueRfc3339('2026-06-15');
  assert.ok(dStr);
  assert.strictEqual(dStr, '2026-06-15T00:00:00.000Z');

  var dObj = new Date(Date.UTC(2026, 7, 20, 0, 0, 0));
  var dObjStr = taskSyncService.formatTaskDueRfc3339(dObj);
  assert.strictEqual(dObjStr, '2026-08-20T00:00:00.000Z');

  assert.strictEqual(taskSyncService.formatTaskDueRfc3339(''), null);
});

test('status mapping: translates bidirectionally between sheet and Google Tasks', () => {
  assert.strictEqual(taskSyncService.mapSheetStatusToGoogleTaskStatus('Done'), 'completed');
  assert.strictEqual(taskSyncService.mapSheetStatusToGoogleTaskStatus('Todo'), 'needsAction');
  assert.strictEqual(taskSyncService.mapSheetStatusToGoogleTaskStatus('In Progress'), 'needsAction');
  assert.strictEqual(taskSyncService.mapSheetStatusToGoogleTaskStatus('Blocked'), 'needsAction');

  assert.strictEqual(taskSyncService.mapGoogleTaskStatusToSheetStatus('completed', 'Todo'), 'Done');
  assert.strictEqual(taskSyncService.mapGoogleTaskStatusToSheetStatus('needsAction', 'Done'), 'Todo');
  assert.strictEqual(taskSyncService.mapGoogleTaskStatusToSheetStatus('needsAction', 'In Progress'), 'In Progress');
});

test('syncGoogleTasks: pushes unsynced tasks with allUsers option', () => {
  freshInit(mockSpreadsheet);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  // Seed 2 unsynced tasks (columns: task_id, project_id, assigned_member, title, description, due_date, status, google_task_id)
  tasksSheet.getRange(2, 1, 2, 8).setValues([
    ['TSK-001', 'PRJ-001', 'Alice Smith', 'Implement OAuth flow', 'OAuth 2.0 flow with Google', '2026-06-30', 'Todo', ''],
    ['TSK-002', 'PRJ-001', 'Bob Jones', 'Build mobile layout', 'Responsive CSS adjustments', '2026-07-15', 'In Progress', '']
  ]);

  var res = taskSyncService.syncGoogleTasks(mockSpreadsheet, { allUsers: true });
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.pushedCount, 2);

  // Check that google_task_id was written back to column 8
  var updatedRows = tasksSheet.getRange(2, 1, 2, 8).getValues();
  var gId1 = updatedRows[0][7];
  var gId2 = updatedRows[1][7];

  assert.ok(gId1, 'Row 2 should have google_task_id populated');
  assert.ok(gId2, 'Row 3 should have google_task_id populated');
  assert.notStrictEqual(gId1, gId2);

  // Verify tasks exist in mock Google Tasks
  var gTask1 = Tasks.Tasks.get(res.tasklistId, gId1);
  assert.strictEqual(gTask1.title, 'Implement OAuth flow');
  assert.ok(gTask1.notes.includes('OAuth 2.0 flow with Google'), 'Notes should contain description');
  assert.ok(gTask1.notes.includes('TSK-001'));
  assert.strictEqual(gTask1.status, 'needsAction');
});

test('syncGoogleTasks: strictly isolates push to active user tasks and ignores tasks assigned to other members', () => {
  freshInit(mockSpreadsheet);

  // Setup tbl_members so Alice Smith is recognized
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 2, 7).setValues([
    ['Alice Smith', '', '', 'Senior Eng', 'Core', 'Active', 'Main Office'],
    ['Bob Jones', '', '', 'Designer', 'UX', 'Active', 'Remote']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  // Task 1 assigned to Alice (active user), Task 2 assigned to Bob (other member)
  tasksSheet.getRange(2, 1, 2, 8).setValues([
    ['TSK-010', 'PRJ-001', 'Alice Smith', 'Alice specific task', 'Alice task notes', '2026-06-30', 'Todo', ''],
    ['TSK-020', 'PRJ-001', 'Bob Jones', 'Bob specific task', 'Bob task notes', '2026-07-15', 'In Progress', '']
  ]);

  // Sync without allUsers option (active user is alice@example.com)
  var res = taskSyncService.syncGoogleTasks(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.pushedCount, 1, 'Only Alice task should be pushed into Alice Google Tasks');

  var rows = tasksSheet.getRange(2, 1, 2, 8).getValues();
  assert.ok(rows[0][7], 'Alice task should have google_task_id');
  assert.strictEqual(rows[1][7], '', 'Bob task should remain unsynced and not pushed to Alice Google Tasks');
});

test('cleanupMisassignedGoogleTasks: unlinks and deletes foreign tasks from active user Google Tasks', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 2, 7).setValues([
    ['Alice Smith', '', '', 'Senior Eng', 'Core', 'Active', 'Main Office'],
    ['Bob Jones', '', '', 'Designer', 'UX', 'Active', 'Remote']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  // Alice task with ID, and Bob task erroneously with ID (synced during Alice session)
  tasksSheet.getRange(2, 1, 2, 8).setValues([
    ['TSK-010', 'PRJ-001', 'Alice Smith', 'Alice task', 'Alice notes', '2026-06-30', 'Todo', 'gtask_alice_123'],
    ['TSK-020', 'PRJ-001', 'Bob Jones', 'Bob task', 'Bob notes', '2026-07-15', 'In Progress', 'gtask_bob_456']
  ]);

  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  // Put both tasks in mock Google Tasks
  Tasks.Tasks.insert({ id: 'gtask_alice_123', title: 'Alice task' }, listId);
  Tasks.Tasks.insert({ id: 'gtask_bob_456', title: 'Bob task' }, listId);

  // Run cleanup for active user (Alice)
  var cleanupRes = taskSyncService.cleanupMisassignedGoogleTasks(mockSpreadsheet);
  assert.strictEqual(cleanupRes.success, true);
  assert.strictEqual(cleanupRes.cleanedCount, 1, 'Should clean 1 task assigned to Bob');

  // Verify Bob task ID in sheet was cleared
  var rows = tasksSheet.getRange(2, 1, 2, 8).getValues();
  assert.strictEqual(rows[0][7], 'gtask_alice_123', 'Alice task ID should be preserved');
  assert.strictEqual(rows[1][7], '', 'Bob task ID should be cleared in sheet');

  // Verify Bob task was removed from Google Tasks
  assert.throws(() => {
    Tasks.Tasks.get(listId, 'gtask_bob_456');
  });
});

test('syncGoogleTasks: pulls completed status from Google Tasks and marks sheet task Done', () => {
  freshInit(mockSpreadsheet);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  tasksSheet.getRange(2, 1, 1, 8).setValues([
    ['TSK-100', 'PRJ-002', 'Alice Smith', 'Design API schema', 'Schema details', '2026-06-01', 'Todo', '']
  ]);

  // Push to Google Tasks first
  var pushRes = taskSyncService.syncGoogleTasks(mockSpreadsheet);
  assert.strictEqual(pushRes.pushedCount, 1);

  var gId = tasksSheet.getRange(2, 8).getValue();
  assert.ok(gId);

  // Simulate user marking the task as completed in Google Tasks mobile app / web
  Tasks.Tasks.patch({ status: 'completed' }, pushRes.tasklistId, gId);

  // Sync again
  var pullRes = taskSyncService.syncGoogleTasks(mockSpreadsheet);
  assert.strictEqual(pullRes.success, true);
  assert.strictEqual(pullRes.pushedCount, 0, 'No new tasks to push');
  assert.strictEqual(pullRes.pulledUpdatesCount, 1, '1 task status updated from Google Tasks');

  // Verify sheet status changed from Todo to Done
  var updatedStatus = tasksSheet.getRange(2, 7).getValue();
  assert.strictEqual(updatedStatus, 'Done');
});

test('syncGoogleTasks: patches Google Tasks when sheet status is marked Done', () => {
  freshInit(mockSpreadsheet);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  tasksSheet.getRange(2, 1, 1, 8).setValues([
    ['TSK-200', 'PRJ-003', 'Alice Smith', 'Write integration tests', 'End to end tests', '2026-06-10', 'Todo', '']
  ]);

  var pushRes = taskSyncService.syncGoogleTasks(mockSpreadsheet);
  var gId = tasksSheet.getRange(2, 8).getValue();

  // User marks task Done in spreadsheet
  tasksSheet.getRange(2, 7).setValue('Done');

  // Sync
  taskSyncService.syncGoogleTasks(mockSpreadsheet);

  // Google Task should now be completed
  var gTask = Tasks.Tasks.get(pushRes.tasklistId, gId);
  assert.strictEqual(gTask.status, 'completed');
});

test('syncTasks: public entry point toasts sync summary and handles missing service gracefully', () => {
  freshInit(mockSpreadsheet);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  tasksSheet.getRange(2, 1, 1, 8).setValues([
    ['TSK-300', 'PRJ-004', 'Alice Smith', 'Setup CI pipeline', 'CI YAML pipeline', '2026-07-01', 'Todo', '']
  ]);

  var res = code.syncTasks(mockSpreadsheet);
  assert.strictEqual(res.success, true);
  assert.strictEqual(res.pushedCount, 1);
  assert.ok(mockSpreadsheet.lastToast);
  assert.ok(mockSpreadsheet.lastToast.title.includes('Tasks Synchronized'));
});

test('automatic task creation onEdit: creates Google Task ID once title is set', () => {
  freshInit(mockSpreadsheet);

  // Seed tbl_members with active user Alice
  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Senior Eng', 'Core', 'Active', 'Main Office']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  // Simulate user entering a title on row 2 in tbl_tasks
  var titleRange = tasksSheet.getRange(2, 4); // Column 4: title
  titleRange.setValue('Build authentication modal');

  // User triggers edit event
  var editEvent = {
    range: titleRange,
    value: 'Build authentication modal'
  };

  code.onEdit(editEvent);

  // Inspect row 2
  var row = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  var taskId = row[0];
  var title = row[3];
  var status = row[6];
  var googleTaskId = row[7];

  assert.ok(taskId, 'Task ID should be auto-generated');
  assert.ok(taskId.startsWith('TSK-'), 'Task ID should start with TSK-');
  assert.strictEqual(title, 'Build authentication modal');
  assert.strictEqual(status, 'Todo', 'Status should default to Todo');
  assert.ok(googleTaskId, 'Google Task ID should be automatically created once title is set');

  // Verify task exists in Google Tasks
  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  var gTask = Tasks.Tasks.get(listId, googleTaskId);
  assert.strictEqual(gTask.title, 'Build authentication modal');
  assert.ok(gTask.notes.includes(taskId));
});

test('automatic task creation with description: includes description in Google Task notes', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Senior Eng', 'Core', 'Active', 'Main Office']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  // User enters description first, then sets title
  tasksSheet.getRange(2, 5).setValue('Detailed specifications for SSO SAML login integration');
  var titleRange = tasksSheet.getRange(2, 4);
  titleRange.setValue('Integrate SAML SSO');

  var editEvent = {
    range: titleRange,
    value: 'Integrate SAML SSO'
  };

  code.onEdit(editEvent);

  var row = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  var gId = row[7];
  assert.ok(gId, 'Google Task ID created');

  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  var gTask = Tasks.Tasks.get(listId, gId);
  assert.strictEqual(gTask.title, 'Integrate SAML SSO');
  assert.ok(gTask.notes.includes('Detailed specifications for SSO SAML login integration'), 'Google Task notes should contain description');
});

test('editing title/description on existing synced task: patches Google Tasks', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', '', '', 'Senior Eng', 'Core', 'Active', 'Main Office']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  // Create task via onEdit first
  var titleRange = tasksSheet.getRange(2, 4);
  titleRange.setValue('Initial title');
  code.onEdit({ range: titleRange });

  var gId = tasksSheet.getRange(2, 8).getValue();
  assert.ok(gId);

  // Now user edits description on existing task
  var descRange = tasksSheet.getRange(2, 5);
  descRange.setValue('Updated description text');
  code.onEdit({ range: descRange });

  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  var gTask = Tasks.Tasks.get(listId, gId);
  assert.ok(gTask.notes.includes('Updated description text'), 'Google Task notes should be updated via patch');
});

test('automatic task creation onEdit: skips foreign members tasks (isolation)', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 2, 7).setValues([
    ['Alice Smith', '', '', 'Senior Eng', 'Core', 'Active', 'Main Office'],
    ['Bob Jones', '', '', 'Designer', 'UX', 'Active', 'Remote']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  // Assign task to Bob Jones before setting title
  tasksSheet.getRange(2, 3).setValue('Bob Jones');
  var titleRange = tasksSheet.getRange(2, 4);
  titleRange.setValue('Bob feature task');

  code.onEdit({ range: titleRange });

  var row = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  var gId = row[7];
  assert.strictEqual(gId, '', 'Google Task ID should remain empty for foreign member task during active user session');
});

test('onOpen: registers Sync Google Tasks and Clean Up Foreign Google Tasks in custom menu', () => {
  code.onOpen();
  var ui = SpreadsheetApp.getUi();
  assert.ok(ui.menu);
  var captions = ui.menu.items.map(function(item) { return item.caption; });
  assert.ok(captions.includes('Sync Google Tasks'));
  assert.ok(captions.includes('Clean Up Foreign Google Tasks'));
});

test('automatic task creation onEdit: generates task_id, defaults status, and creates Google Task with task_id title when editing project_id first', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Senior Eng', 'Core', 'Active', 'Main Office']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  var projRange = tasksSheet.getRange(2, 2); // Column 2: project_id
  projRange.setValue('PRJ-INFRA');

  code.onEdit({ range: projRange, value: 'PRJ-INFRA' });

  var row = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  var taskId = row[0];
  var projectId = row[1];
  var assignedMember = row[2];
  var status = row[6];
  var googleTaskId = row[7];

  assert.ok(taskId, 'Task ID should be auto-generated immediately when project_id is entered');
  assert.ok(taskId.startsWith('TSK-'), 'Task ID must start with TSK-');
  assert.strictEqual(projectId, 'PRJ-INFRA');
  assert.strictEqual(status, 'Todo', 'Status should default to Todo immediately');
  assert.strictEqual(assignedMember, '', 'Assigned member should NOT be filled automatically');
  assert.strictEqual(row[5], '', 'Due date should remain empty / nullable');
  assert.ok(googleTaskId, 'Google Task ID should be automatically created immediately once task_id is generated');

  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  var gTask = Tasks.Tasks.get(listId, googleTaskId);
  assert.strictEqual(gTask.title, taskId, 'Google Task initial title should default to task_id');
});

test('automatic task creation onEdit: allows nullable due date and creates Google Task without due date', () => {
  freshInit(mockSpreadsheet);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  var titleRange = tasksSheet.getRange(2, 4);
  titleRange.setValue('Task with no due date');

  code.onEdit({ range: titleRange, value: 'Task with no due date' });

  var row = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  assert.ok(row[0].startsWith('TSK-'), 'Task ID generated');
  assert.strictEqual(row[2], '', 'Assigned member remains blank');
  assert.strictEqual(row[3], 'Task with no due date');
  assert.strictEqual(row[5], '', 'Due date is nullable / blank');
  assert.strictEqual(row[6], 'Todo');
  assert.ok(row[7], 'Google Task ID created');

  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  var gTask = Tasks.Tasks.get(listId, row[7]);
  assert.strictEqual(gTask.title, 'Task with no due date');
  assert.strictEqual(gTask.due, undefined, 'Google Task has no due date');
});

test('automatic task creation onEdit: creates Google Task with task_id title on row init, and editing title replaces it with real title', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Senior Eng', 'Core', 'Active', 'Main Office']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  // First, user entered project_id and description
  var projRange = tasksSheet.getRange(2, 2);
  projRange.setValue('PRJ-009');
  code.onEdit({ range: projRange });

  var row1 = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  var existingTaskId = row1[0];
  var initialGoogleTaskId = row1[7];
  assert.ok(existingTaskId, 'task_id generated');
  assert.ok(initialGoogleTaskId, 'google_task_id created immediately upon task_id creation');

  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  var gTaskInitial = Tasks.Tasks.get(listId, initialGoogleTaskId);
  assert.strictEqual(gTaskInitial.title, existingTaskId, 'Initial Google Task title should default to task_id');

  // Now, user enters the real title
  var titleRange = tasksSheet.getRange(2, 4);
  titleRange.setValue('Deploy Kubernetes cluster');
  code.onEdit({ range: titleRange });

  var row2 = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  assert.strictEqual(row2[0], existingTaskId, 'Existing task_id should be preserved');
  assert.strictEqual(row2[3], 'Deploy Kubernetes cluster');
  assert.strictEqual(row2[7], initialGoogleTaskId, 'Google Task ID should remain identical');

  // Verify Google Task title was patched and replaced by the real title
  var gTaskUpdated = Tasks.Tasks.get(listId, initialGoogleTaskId);
  assert.strictEqual(gTaskUpdated.title, 'Deploy Kubernetes cluster', 'Google Task title should be patched with real title');
  assert.ok(gTaskUpdated.notes.includes(existingTaskId), 'Google Task notes should contain task_id');
});

test('automatic task creation onEdit: resilient to human-readable headers with spaces and casing', () => {
  freshInit(mockSpreadsheet);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  // Replace headers with human readable names
  tasksSheet.getRange(1, 1, 1, 8).setValues([[
    'Task ID', 'Project ID', 'Assigned Member', 'Title', 'Description', 'Due Date', 'Status', 'Google Task ID'
  ]]);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Senior Eng', 'Core', 'Active', 'Main Office']
  ]);

  // User enters title in Column 4
  var titleRange = tasksSheet.getRange(2, 4);
  titleRange.setValue('Test human headers');
  code.onEdit({ range: titleRange });

  var row = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  assert.ok(row[0].startsWith('TSK-'), 'Task ID should be created in column 1');
  assert.strictEqual(row[3], 'Test human headers', 'Title should be in column 4');
  assert.strictEqual(row[6], 'Todo', 'Status should be in column 7');
  assert.ok(row[7], 'Google Task ID should be placed in column 8');
});

test('automatic task creation onEdit: in-sheet task creation succeeds even if Tasks API throws', () => {
  freshInit(mockSpreadsheet);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  // Temporarily sabotage Tasks.Tasks.insert to simulate unauthorized simple trigger exception
  var originalInsert = Tasks.Tasks.insert;
  Tasks.Tasks.insert = function() {
    throw new Error('Exception: You do not have permission to call Tasks.Tasks.insert');
  };

  try {
    var titleRange = tasksSheet.getRange(2, 4);
    titleRange.setValue('Offline created task');
    code.onEdit({ range: titleRange });

    var row = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
    assert.ok(row[0].startsWith('TSK-'), 'Task ID should be generated despite Tasks API failure');
    assert.strictEqual(row[3], 'Offline created task');
    assert.strictEqual(row[6], 'Todo', 'Status should default to Todo despite Tasks API failure');
    assert.strictEqual(row[7], '', 'Google Task ID should stay empty when Tasks API fails');
  } finally {
    Tasks.Tasks.insert = originalInsert;
  }
});

test('automatic task creation onEdit: multiple title revisions update Google Task title from task_id to new titles', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Senior Eng', 'Core', 'Active', 'Main Office']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');

  // Step 1: User fills description first (title is empty)
  var descRange = tasksSheet.getRange(2, 5);
  descRange.setValue('Initial description without title');
  code.onEdit({ range: descRange });

  var row1 = tasksSheet.getRange(2, 1, 1, 8).getValues()[0];
  var taskId = row1[0];
  var gId = row1[7];
  assert.ok(taskId, 'Task ID created');
  assert.ok(gId, 'Google Task ID created immediately');

  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  var gTask1 = Tasks.Tasks.get(listId, gId);
  assert.strictEqual(gTask1.title, taskId, 'Initial Google Task title is task_id');

  // Step 2: User sets title for the first time
  var titleRange = tasksSheet.getRange(2, 4);
  titleRange.setValue('First Real Title');
  code.onEdit({ range: titleRange });

  var gTask2 = Tasks.Tasks.get(listId, gId);
  assert.strictEqual(gTask2.title, 'First Real Title', 'Google Task title replaced with real title');

  // Step 3: User edits title a second time
  titleRange.setValue('Updated Second Title');
  code.onEdit({ range: titleRange });

  var gTask3 = Tasks.Tasks.get(listId, gId);
  assert.strictEqual(gTask3.title, 'Updated Second Title', 'Google Task title updated to second title');

  // Step 4: User clears title (empty string)
  titleRange.setValue('');
  code.onEdit({ range: titleRange });

  var gTask4 = Tasks.Tasks.get(listId, gId);
  assert.strictEqual(gTask4.title, taskId, 'Google Task title safely falls back to taskId when cleared');
});

test('validation: project_id in tasks is not required and passes validation', () => {
  freshInit(mockSpreadsheet);
  const { validateRow } = require('../src/ValidationManager');
  // Columns: task_id, project_id, assigned_member, title, description, due_date, status, google_task_id
  var rowWithoutProject = ['TSK-001', '', '', 'Test Task Without Project', 'Some notes', '2026-06-30', 'Todo', ''];
  var res = validateRow('tbl_tasks', rowWithoutProject);
  assert.strictEqual(res.valid, true, 'Task row without project_id must be valid');
  assert.strictEqual(res.missingColumns.includes('project_id'), false, 'project_id should not be in missingColumns');
});

test('syncGoogleTasks: updates all fields (title, description/notes, due_date, status) in Google Tasks', () => {
  freshInit(mockSpreadsheet);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  // Seed initial task
  tasksSheet.getRange(2, 1, 1, 8).setValues([
    ['TSK-500', 'PRJ-CORE', 'Alice Smith', 'Initial Name', 'Initial Notes', '2026-06-01', 'Todo', '']
  ]);

  var pushRes = taskSyncService.syncGoogleTasks(mockSpreadsheet);
  assert.strictEqual(pushRes.pushedCount, 1);
  var gId = tasksSheet.getRange(2, 8).getValue();

  // User edits title, description, due_date, and status in sheet
  tasksSheet.getRange(2, 4).setValue('Renamed Feature');
  tasksSheet.getRange(2, 5).setValue('Updated Architectural Notes');
  tasksSheet.getRange(2, 6).setValue('2026-07-20');
  tasksSheet.getRange(2, 7).setValue('Done');

  // Trigger sync
  var syncRes = taskSyncService.syncGoogleTasks(mockSpreadsheet);
  assert.strictEqual(syncRes.success, true);
  assert.ok(syncRes.pushedFieldUpdatesCount >= 1, 'Field updates should be pushed to Google Tasks');

  // Verify Google Task fields updated
  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  var gTask = Tasks.Tasks.get(listId, gId);
  assert.strictEqual(gTask.title, 'Renamed Feature');
  assert.ok(gTask.notes.includes('Updated Architectural Notes'));
  assert.ok(gTask.notes.includes('TSK-500'));
  assert.strictEqual(gTask.due, '2026-07-20T00:00:00.000Z');
  assert.strictEqual(gTask.status, 'completed');
});

test('changing status from Done: user cancels -> reverts status to Done without creating new task', () => {
  freshInit(mockSpreadsheet);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  // Seed a Done task
  tasksSheet.getRange(2, 1, 1, 8).setValues([
    ['TSK-600', 'PRJ-001', 'Alice Smith', 'Complete module A', 'Notes A', '2026-06-01', 'Done', 'gtask_600']
  ]);
  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  Tasks.Tasks.insert({ id: 'gtask_600', title: 'Complete module A', status: 'completed' }, listId);

  // User tries to change status to "In Progress", but cancels dialog (proceed = false)
  var statusRange = tasksSheet.getRange(2, 7);
  statusRange.setValue('In Progress');

  var editEvent = {
    range: statusRange,
    value: 'In Progress',
    oldValue: 'Done',
    proceed: false
  };

  code.onEdit(editEvent);

  // Status should be reverted to Done
  var currentStatus = tasksSheet.getRange(2, 7).getValue();
  assert.strictEqual(currentStatus, 'Done', 'Status must revert to Done');
  assert.strictEqual(tasksSheet.getLastRow(), 2, 'No new row should be created');
  assert.strictEqual(tasksSheet.isRowHiddenByUser(2), false, 'Row must not be hidden');
});

test('changing status from Done: user proceeds -> hides original row, creates follow-up task with incremented task_id and new Google Task', () => {
  freshInit(mockSpreadsheet);

  var membersSheet = mockSpreadsheet.getSheetByName('Members');
  membersSheet.getRange(2, 1, 1, 7).setValues([
    ['Alice Smith', 'alice@accenture.com', '', 'Senior Eng', 'Core', 'Active', 'Main Office']
  ]);

  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  tasksSheet.getRange(2, 1, 1, 8).setValues([
    ['TSK-700', 'PRJ-ALPHA', 'Alice Smith', 'Deliver feature X', 'Detailed spec X', '2026-06-15', 'Done', 'gtask_700']
  ]);
  var listId = taskSyncService.getTasksListId(mockSpreadsheet);
  Tasks.Tasks.insert({ id: 'gtask_700', title: 'Deliver feature X', status: 'completed' }, listId);

  // User tries to change status to "In Progress" and proceeds
  var statusRange = tasksSheet.getRange(2, 7);
  statusRange.setValue('In Progress');

  var editEvent = {
    range: statusRange,
    value: 'In Progress',
    oldValue: 'Done',
    proceed: true
  };

  code.onEdit(editEvent);

  // Original row 2 must be reverted to Done and hidden
  assert.strictEqual(tasksSheet.getRange(2, 7).getValue(), 'Done', 'Original row stays Done');
  assert.strictEqual(tasksSheet.isRowHiddenByUser(2), true, 'Original row must be hidden');

  // New row 3 must be created with incremented task ID and same metadata
  assert.strictEqual(tasksSheet.getLastRow(), 3, 'New follow-up row must be created');
  var newRow = tasksSheet.getRange(3, 1, 1, 8).getValues()[0];
  var newTaskId = newRow[0];
  var newProjectId = newRow[1];
  var newAssignee = newRow[2];
  var newTitle = newRow[3];
  var newDesc = newRow[4];
  var newDueDate = newRow[5];
  var newStatus = newRow[6];
  var newGoogleTaskId = newRow[7];

  assert.strictEqual(newTaskId, 'TSK-700.1', 'Task ID must be incremented with .1');
  assert.strictEqual(newProjectId, 'PRJ-ALPHA', 'project_id must be copied');
  assert.strictEqual(newAssignee, 'Alice Smith', 'assigned_member must be copied');
  assert.strictEqual(newTitle, 'Deliver feature X', 'title must be copied');
  assert.strictEqual(newDesc, 'Detailed spec X', 'description must be copied');
  assert.strictEqual(newDueDate, '2026-06-15', 'due_date must be copied');
  assert.strictEqual(newStatus, 'In Progress', 'new status must be applied');

  // Verify new Google Task created
  assert.ok(newGoogleTaskId, 'New Google Task ID must be created');
  assert.notStrictEqual(newGoogleTaskId, 'gtask_700', 'New Google Task ID must be distinct from original');

  var newGTask = Tasks.Tasks.get(listId, newGoogleTaskId);
  assert.strictEqual(newGTask.title, 'Deliver feature X');
  assert.strictEqual(newGTask.status, 'needsAction');
  assert.ok(newGTask.notes.includes('TSK-700.1'));

  // Test secondary follow-up increment from TSK-700.1 -> TSK-700.2
  // Simulate user changing row 3 from Done to In Progress
  var statusRange3 = tasksSheet.getRange(3, 7);
  statusRange3.setValue('In Progress');
  var editEvent2 = {
    range: statusRange3,
    value: 'In Progress',
    oldValue: 'Done',
    proceed: true
  };
  code.onEdit(editEvent2);

  assert.strictEqual(tasksSheet.getLastRow(), 4);
  var row4 = tasksSheet.getRange(4, 1, 1, 8).getValues()[0];
  assert.strictEqual(row4[0], 'TSK-700.2', 'Task ID must be incremented from .1 to .2');
});

test('live task sync: installTaskSyncTrigger registers onEdit trigger and triggers handler', () => {
  freshInit(mockSpreadsheet);

  var installRes = code.installLiveTaskSync(mockSpreadsheet);
  assert.strictEqual(installRes.success, true);
  assert.strictEqual(installRes.installed, true);

  // Calling again should report already installed
  var repeatRes = code.installLiveTaskSync(mockSpreadsheet);
  assert.strictEqual(repeatRes.alreadyInstalled, true);
});



test('Google Tasks date: the Planned date when set, else the Deadline; the deadline is written in the notes', () => {
  freshInit(mockSpreadsheet);
  var tasksSheet = mockSpreadsheet.getSheetByName('Tasks');
  assert.strictEqual(tasksSheet.getRange(1, 6).getValue(), 'Deadline');
  assert.strictEqual(tasksSheet.getRange(1, 13).getValue(), 'Planned date');
  tasksSheet.getRange(2, 1, 2, 8).setValues([
    ['TSK-001', 'PRJ-001', 'Alice Smith', 'SoW', 'Sign the SoW', '2026-10-30', 'Todo', ''],
    ['TSK-002', 'PRJ-001', 'Alice Smith', 'PRM', '', '2026-10-31', 'Todo', '']
  ]);
  tasksSheet.getRange(2, 13).setValue('2026-10-26');

  var res = taskSyncService.syncGoogleTasks(mockSpreadsheet, { allUsers: true });

  var rows = tasksSheet.getRange(2, 1, 2, 8).getValues();
  var planned = Tasks.Tasks.get(res.tasklistId, rows[0][7]);
  var deadlineOnly = Tasks.Tasks.get(res.tasklistId, rows[1][7]);
  assert.strictEqual(planned.due.substring(0, 10), '2026-10-26', 'planned day sent as the Google Tasks date');
  assert.match(planned.notes, /Sign the SoW\n\nDeadline: 2026-10-30\n\nTask ID: TSK-001/);
  assert.strictEqual(deadlineOnly.due.substring(0, 10), '2026-10-31', 'no planned date: the deadline is the day');

  // Moving the planned date moves the Google task; the deadline stays in the notes
  tasksSheet.getRange(2, 13).setValue('2026-10-28');
  taskSyncService.syncGoogleTasks(mockSpreadsheet, { allUsers: true });
  var moved = Tasks.Tasks.get(res.tasklistId, rows[0][7]);
  assert.strictEqual(moved.due.substring(0, 10), '2026-10-28');
  assert.match(moved.notes, /Deadline: 2026-10-30/);
});
