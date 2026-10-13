/**
 * Project Management Workspace - Google Tasks API Synchronization
 * Ticket 07: Google Tasks API Synchronization
 *
 * Integrates tbl_tasks with Google Tasks:
 * - Pushes unsynced tasks (where google_task_id is empty) to Google Tasks via Tasks.Tasks.insert().
 * - Writes back returned task IDs into tbl_tasks.google_task_id in a batch update.
 * - Inspects modified tasks in Google Tasks and syncs status updates back to tbl_tasks.
 * - Handles batching and error handling gracefully.
 * - Stores active Google Tasks List ID in Document Properties.
 */

/**
 * Key used in Document Properties for the active Google Tasks List ID.
 */
var GOOGLE_TASKS_LIST_ID_KEY = "GOOGLE_TASKS_LIST_ID";

/**
 * Retrieves the active Google Tasks List ID from Document Properties,
 * creating or discovering a dedicated "Project Management Tasks" list if not yet configured.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {string} Tasklist ID
 */
function getTasksListId(targetSpreadsheet) {
  // Each person syncs into their own list (ADR 0045): the ID lives in user properties
  var userProps = null;
  try { userProps = PropertiesService.getUserProperties(); } catch (eUp) {}
  var saved = userProps ? userProps.getProperty(GOOGLE_TASKS_LIST_ID_KEY) : "";
  if (saved && saved.trim()) return saved.trim();

  // A list configured for the sheet (setTasksListId) is used when this person can open it
  var docProps = null;
  try { docProps = PropertiesService.getDocumentProperties(); } catch (eDp) {}
  var configured = docProps ? docProps.getProperty(GOOGLE_TASKS_LIST_ID_KEY) : "";
  if (configured && configured.trim() && typeof Tasks !== "undefined" && Tasks.Tasklists) {
    try {
      if (configured === "@default" || Tasks.Tasklists.get(configured.trim())) {
        if (userProps) userProps.setProperty(GOOGLE_TASKS_LIST_ID_KEY, configured.trim());
        return configured.trim();
      }
    } catch (eGet) {}
  }

  // Otherwise find or create this person's "Project Management Tasks" list
  if (typeof Tasks !== "undefined" && Tasks.Tasklists) {
    try {
      var lists = Tasks.Tasklists.list();
      var items = (lists && lists.items) ? lists.items : [];
      for (var i = 0; i < items.length; i++) {
        if (items[i].title === "Project Management Tasks") {
          if (userProps) userProps.setProperty(GOOGLE_TASKS_LIST_ID_KEY, items[i].id);
          return items[i].id;
        }
      }
      var created = Tasks.Tasklists.insert({ title: "Project Management Tasks" });
      if (created && created.id) {
        if (userProps) userProps.setProperty(GOOGLE_TASKS_LIST_ID_KEY, created.id);
        return created.id;
      }
    } catch (eList) {
      Logger.log("getTasksListId discovery error: " + eList);
    }
  }
  return "@default";
}

/** Every task of a Google Tasks list (all pages, completed and hidden included). */
function listAllGoogleTasks(tasklistId) {
  var all = [];
  var pageToken = null;
  var guard = 0;
  do {
    var args = { showCompleted: true, showHidden: true, maxResults: 100 };
    if (pageToken) args.pageToken = pageToken;
    var page = Tasks.Tasks.list(tasklistId, args);
    all = all.concat((page && page.items) ? page.items : []);
    pageToken = page ? page.nextPageToken : null;
    guard++;
  } while (pageToken && guard < 50);
  return all;
}

/**
 * Sets the active Google Tasks List ID in Document Properties.
 *
 * @param {string} listId
 */
function setTasksListId(listId) {
  var id = String(listId || "@default").trim();
  PropertiesService.getDocumentProperties().setProperty(GOOGLE_TASKS_LIST_ID_KEY, id);
  // The person configuring it uses it right away; others adopt it if they can open it
  try { PropertiesService.getUserProperties().setProperty(GOOGLE_TASKS_LIST_ID_KEY, id); } catch (e) {}
}

/**
 * Formats a Date object or string into RFC 3339 timestamp string for Google Tasks API.
 *
 * @param {Date|string} rawDate
 * @return {string|null}
 */
function formatTaskDueRfc3339(rawDate) {
  if (!rawDate) return null;
  var d = null;
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    d = rawDate;
  } else {
    var str = String(rawDate).trim();
    var match = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (match) {
      d = new Date(Date.UTC(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10), 0, 0, 0, 0));
    } else {
      d = new Date(str);
    }
  }
  if (!d || isNaN(d.getTime())) return null;
  return d.toISOString();
}

/**
 * Maps spreadsheet status ("Todo", "In Progress", "Done", "Blocked") to Google Tasks status.
 *
 * @param {string} sheetStatus
 * @return {string} "needsAction" or "completed"
 */
function mapSheetStatusToGoogleTaskStatus(sheetStatus) {
  var s = String(sheetStatus || "").trim().toLowerCase();
  if (s === "done" || s === "completed") {
    return "completed";
  }
  return "needsAction";
}

/**
 * Maps Google Tasks status ("needsAction", "completed") to spreadsheet status.
 *
 * @param {string} googleTaskStatus
 * @param {string} currentSheetStatus
 * @return {string}
 */
function mapGoogleTaskStatusToSheetStatus(googleTaskStatus, currentSheetStatus) {
  var g = String(googleTaskStatus || "").trim().toLowerCase();
  var curr = String(currentSheetStatus || "").trim();

  if (g === "completed") {
    return "Done";
  }
  // If needsAction: if it was Done, revert to Todo; otherwise preserve current ("In Progress", "Blocked", "Todo")
  if (curr.toLowerCase() === "done") {
    return "Todo";
  }
  return curr || "Todo";
}

/**
 * Generates a unique task ID in standard TSK-YYMMDD-XXXX format.
 * Uses Utilities.formatDate if available, or native Date as reliable fallback.
 *
 * @return {string} Generated task ID
 */
function generateTaskId() {
  var prefix = "TSK-";
  try {
    if (typeof Utilities !== "undefined" && Utilities.formatDate && typeof Session !== "undefined" && Session.getScriptTimeZone) {
      return prefix + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "yyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
    }
  } catch (eFmt) {}
  var d = new Date();
  var yy = String(d.getFullYear()).slice(-2);
  var mNum = d.getMonth() + 1;
  var mm = (mNum < 10 ? "0" : "") + mNum;
  var dNum = d.getDate();
  var dd = (dNum < 10 ? "0" : "") + dNum;
  var rand = Math.floor(1000 + Math.random() * 9000);
  return prefix + yy + mm + dd + "-" + rand;
}

/**
 * The day a task shows in Google Tasks: its Planned date, else its Deadline
 * (the API has a single date field, "due", documented as the scheduled day,
 * not a deadline; ADR 0050).
 */
function taskGoogleDay(row, colMap) {
  var planned = (colMap.plannedDate > 0) ? row[colMap.plannedDate - 1] : "";
  return planned || row[colMap.dueDate - 1];
}

/**
 * Google Task notes: description, then the deadline (the API has no deadline
 * field), then the sheet references.
 */
function taskGoogleNotes(description, deadline, taskId, projectId, assignedMember) {
  var notesParts = [];
  if (description) notesParts.push(description);
  var deadlineIso = isoDate(deadline);
  if (deadlineIso) notesParts.push("Deadline: " + deadlineIso);
  var metaParts = [];
  if (taskId) metaParts.push("Task ID: " + taskId);
  if (projectId) metaParts.push("Project: " + projectId);
  if (assignedMember) metaParts.push("Assigned: " + assignedMember);
  if (metaParts.length > 0) notesParts.push(metaParts.join(" | "));
  return notesParts.join("\n\n");
}

/**
 * Resolves column positions for tbl_tasks dynamically from headers,
 * handling both the new 8-column layout (with description) and legacy 7-column layout.
 * Supports case-insensitive matching, spaces, and underscores.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} tasksSheet
 * @return {Object} Map of column names to 1-indexed column numbers
 */
function getTaskColumnMap(tasksSheet) {
  var map = {
    taskId: 1,
    projectId: 2,
    assignedMember: 3,
    title: 4,
    description: 5,
    dueDate: 6,
    status: 7,
    googleTaskId: 8,
    plannedDate: 0,
    origin: 0
  };
  if (!tasksSheet) return map;
  var lastCol = Math.max(typeof tasksSheet.getLastColumn === "function" ? (tasksSheet.getLastColumn() || 0) : 0, 8);
  if (lastCol < 1) return map;
  var headerVals = tasksSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var hasDesc = false;
  for (var c = 0; c < headerVals.length; c++) {
    var raw = String(headerVals[c] || "").trim().toLowerCase();
    var norm = raw.replace(/[^a-z0-9]/g, "");

    if (norm === "taskid" || norm === "id") {
      map.taskId = c + 1;
    } else if (norm === "projectid" || norm === "project") {
      map.projectId = c + 1;
    } else if (norm === "assignedmember" || norm === "member" || norm === "assignee" || norm === "assigned" || norm === "assignedto") {
      map.assignedMember = c + 1;
    } else if (norm === "title" || norm === "tasktitle" || norm === "name" || norm === "taskname") {
      map.title = c + 1;
    } else if (norm === "description" || norm === "desc" || norm === "notes" || norm === "details") {
      map.description = c + 1;
      hasDesc = true;
    } else if (norm === "duedate" || norm === "due" || norm === "deadline") {
      map.dueDate = c + 1;
    } else if (norm === "planneddate" || norm === "planned") {
      map.plannedDate = c + 1;
    } else if (norm === "origin" || norm === "recurring" || norm === "recurrence") {
      map.origin = c + 1;
    } else if (norm === "status" || norm === "state") {
      map.status = c + 1;
    } else if (norm === "googletaskid" || norm === "gtaskid" || norm === "googletasks" || norm === "googletask") {
      map.googleTaskId = c + 1;
    }
  }

  if (!hasDesc) {
    // If no description column was found anywhere and sheet has 7 columns or fewer
    var matchedMax = Math.max(map.taskId, map.projectId, map.assignedMember, map.title, map.dueDate, map.status, map.googleTaskId);
    if (lastCol <= 7 || matchedMax <= 7) {
      map.description = 0;
      if (map.dueDate === 6) map.dueDate = 5;
      if (map.status === 7) map.status = 6;
      if (map.googleTaskId === 8) map.googleTaskId = 7;
    }
  }
  return map;
}

/**
 * Synchronizes tbl_tasks with Google Tasks bi-directionally:
 * 1. Pushes unsynced tasks (missing google_task_id) to Google Tasks.
 * 2. Writes back generated task IDs into tbl_tasks.google_task_id in batch.
 * 3. Pulls status updates from Google Tasks and reconciles with tbl_tasks.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Object} [options]
 * @return {Object} Synchronization summary report
 */
function syncGoogleTasks(targetSpreadsheet, options) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet available" };

  options = options || {};

  if (typeof Tasks === "undefined" || !Tasks.Tasks) {
    return {
      success: false,
      message: "Google Tasks Advanced Service is not enabled. Please enable Tasks API in Extensions > Apps Script > Services."
    };
  }

  var tasksSheet = getTable(ss, "tasks");
  if (!tasksSheet) {
    return { success: false, message: "Sheet 'Tasks' not found" };
  }

  var lastRow = tasksSheet.getLastRow();
  if (lastRow <= 1) {
    return {
      success: true,
      message: "Tasks is empty. No tasks to synchronize.",
      pushedCount: 0,
      pulledUpdatesCount: 0,
      totalTasks: 0
    };
  }

  // Determine active user identity to avoid pushing other members' tasks to current user's Google Tasks
  var activeUserEmail = options.userEmail || "";
  if (!activeUserEmail) {
    try {
      var aUser = Session.getActiveUser();
      activeUserEmail = aUser ? aUser.getEmail() : "";
    } catch (eUser) {}
  }

  var activeMemberNames = [];
  if (activeUserEmail && ss) {
    var membersSheet = getTable(ss, "members");
    if (membersSheet && membersSheet.getMaxRows() > 1) {
      var mRows = membersSheet.getRange(2, 1, Math.min(membersSheet.getMaxRows() - 1, 300), Math.max(membersSheet.getLastColumn(), 7));
      var mVals = mRows.getValues();
      var mRich = (typeof mRows.getRichTextValues === "function") ? mRows.getRichTextValues() : [];
      for (var mi = 0; mi < mVals.length; mi++) {
        var mRaw = mVals[mi][0];
        var mR = (mRich && mRich[mi]) ? mRich[mi][0] : null;
        var mExtEmail = (typeof extractEmailFromChip === "function") ? extractEmailFromChip(mRaw, mR) : "";
        var mAccEmail = String(mVals[mi][1] || "").trim().toLowerCase();
        var matches = false;
        if (mExtEmail && mExtEmail.toLowerCase() === activeUserEmail.toLowerCase()) matches = true;
        else if (mAccEmail && mAccEmail === activeUserEmail.toLowerCase()) matches = true;
        else if (typeof isUserMemberMatch === "function" && isUserMemberMatch(mRaw, mR, activeUserEmail, mExtEmail)) matches = true;

        if (matches) {
          var mName = (typeof extractCellValue === "function") ? extractCellValue(mRaw, mR) : String(mRaw || "").trim();
          if (mName) activeMemberNames.push(mName.toLowerCase());
        }
      }
    }
  }

  var colMap = getTaskColumnMap(tasksSheet);
  var tasklistId = getTasksListId(ss);
  var numCols = Math.max(tasksSheet.getLastColumn(), 8);
  var range = tasksSheet.getRange(2, 1, lastRow - 1, numCols);
  var values = range.getValues();
  var richValues = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];

  var pushedCount = 0;
  var pushedFieldUpdatesCount = 0;
  var pulledUpdatesCount = 0;
  var pushedUpdates = []; // { row, col, googleTaskId }
  var statusUpdates = []; // { row, col, status }
  var removedRows = [];     // rows whose Google Task was removed from this person's list
  var myReferencedIds = {}; // Google Task IDs still referenced by this person's rows

  // 1. First, retrieve current tasks from Google Tasks for reconciliation
  var existingGoogleTasksMap = {};
  var listedOk = false;
  try {
    var gItems = listAllGoogleTasks(tasklistId);
    for (var g = 0; g < gItems.length; g++) {
      existingGoogleTasksMap[gItems[g].id] = gItems[g];
    }
    listedOk = true;
  } catch (eGList) {
    Logger.log("Error listing Google Tasks: " + eGList);
  }

  // 2. Iterate through rows in tbl_tasks
  for (var r = 0; r < values.length; r++) {
    var sheetRowIndex = r + 2;
    var row = values[r];
    var richRow = richValues[r] || [];

    var taskId = String(row[colMap.taskId - 1] || "").trim();
    var projectId = String(row[colMap.projectId - 1] || "").trim();

    var rawMember = row[colMap.assignedMember - 1];
    var richMember = (richRow && richRow[colMap.assignedMember - 1]) ? richRow[colMap.assignedMember - 1] : null;
    var assignedMember = (typeof extractCellValue === "function")
      ? extractCellValue(rawMember, richMember)
      : String(rawMember || "").trim();

    var title = String(row[colMap.title - 1] || "").trim();
    var description = (colMap.description > 0) ? String(row[colMap.description - 1] || "").trim() : "";
    var dueDate = row[colMap.dueDate - 1];
    var googleDay = taskGoogleDay(row, colMap);
    var status = String(row[colMap.status - 1] || "Todo").trim();
    var googleTaskId = String(row[colMap.googleTaskId - 1] || "").trim();

    // Skip completely empty rows
    if (!taskId && !title && !projectId) continue;

    // Use default title if missing
    if (!title) {
      title = taskId ? taskId : "Untitled Project Task";
    }

    // Determine if task is assigned to active user
    var isAssignedToActiveUser = true;
    if (activeUserEmail && !options.allUsers) {
      var taskExtEmail = (typeof extractEmailFromChip === "function") ? extractEmailFromChip(rawMember, richMember) : "";
      if (taskExtEmail && taskExtEmail.toLowerCase() === activeUserEmail.toLowerCase()) {
        isAssignedToActiveUser = true;
      } else if (typeof isUserMemberMatch === "function" && isUserMemberMatch(rawMember, richMember, activeUserEmail, taskExtEmail)) {
        isAssignedToActiveUser = true;
      } else if (activeMemberNames.length > 0 && assignedMember && activeMemberNames.indexOf(assignedMember.toLowerCase()) !== -1) {
        isAssignedToActiveUser = true;
      } else {
        isAssignedToActiveUser = false;
      }
    }

    var isCancelled = status.toLowerCase() === "cancelled";
    var inMyList = !!(googleTaskId && existingGoogleTasksMap[googleTaskId]);

    // Reassigned away from me, or cancelled: remove it from my list (ADR 0045)
    if (inMyList && (!isAssignedToActiveUser || isCancelled)) {
      try { Tasks.Tasks.remove(tasklistId, googleTaskId); } catch (eRm) { Logger.log("Google Task removal note: " + eRm); }
      delete existingGoogleTasksMap[googleTaskId];
      removedRows.push({ row: sheetRowIndex, col: colMap.googleTaskId });
      continue;
    }
    if (isCancelled) continue;
    // Mine but its Google Task is in someone else's list (it was reassigned to me): push a new one
    if (isAssignedToActiveUser && googleTaskId && listedOk && !inMyList) googleTaskId = "";
    if (googleTaskId && isAssignedToActiveUser) myReferencedIds[googleTaskId] = true;

    // A. Unsynced task: Push to Google Tasks (only if assigned to current user)
    if (!googleTaskId) {
      if (status.toLowerCase() === "done") continue; // nothing to remind about
      if (!isAssignedToActiveUser) {
        // Skip tasks that belong to other team members
        continue;
      }
      try {
        var taskResource = {
          title: title,
          notes: taskGoogleNotes(description, dueDate, taskId, projectId, assignedMember),
          status: mapSheetStatusToGoogleTaskStatus(status)
        };

        var dueStr = formatTaskDueRfc3339(googleDay);
        if (dueStr) {
          taskResource.due = dueStr;
        }

        var createdTask = Tasks.Tasks.insert(taskResource, tasklistId);
        if (createdTask && createdTask.id) {
          pushedCount++;
          pushedUpdates.push({
            row: sheetRowIndex,
            col: colMap.googleTaskId,
            googleTaskId: createdTask.id
          });
          // Track in map
          existingGoogleTasksMap[createdTask.id] = createdTask;
          myReferencedIds[createdTask.id] = true;
        }
      } catch (eInsert) {
        Logger.log("Failed to insert task (Row " + sheetRowIndex + "): " + eInsert);
      }
    } else {
      // B. Existing synced task: Reconcile bi-directionally (only if assigned to current user)
      if (!isAssignedToActiveUser) {
        continue;
      }

      var matchedGTask = existingGoogleTasksMap[googleTaskId];
      if (matchedGTask) {
        var gStatus = matchedGTask.status; // "needsAction" or "completed"
        var targetGStatus = mapSheetStatusToGoogleTaskStatus(status);
        var patchResource = {};

        // 1. Status Reconciliation
        if (status === "Done" && gStatus !== "completed") {
          patchResource.status = "completed";
        } else if (gStatus === "completed" && status !== "Done") {
          statusUpdates.push({
            row: sheetRowIndex,
            col: colMap.status,
            status: "Done"
          });
          pulledUpdatesCount++;
        } else if (gStatus !== targetGStatus) {
          patchResource.status = targetGStatus;
        }

        // 2. Title Reconciliation: push sheet title to Google Tasks if different
        var effectiveTitle = title || taskId;
        if (effectiveTitle && matchedGTask.title !== effectiveTitle) {
          patchResource.title = effectiveTitle;
        }

        // 3. Date Reconciliation (the Planned date, else the Deadline; ADR 0050)
        var dueStr = formatTaskDueRfc3339(googleDay);
        var gDueStr = matchedGTask.due ? matchedGTask.due.substring(0, 10) : "";
        var sDueStr = dueStr ? dueStr.substring(0, 10) : "";
        if (sDueStr !== gDueStr) {
          patchResource.due = dueStr || null;
        }

        // 4. Description & Notes Reconciliation
        var expectedNotes = taskGoogleNotes(description, dueDate, taskId, projectId, assignedMember);

        if ((matchedGTask.notes || "").trim() !== expectedNotes.trim()) {
          patchResource.notes = expectedNotes;
        }

        // Patch Google Task if any fields or status changed
        if (Object.keys(patchResource).length > 0) {
          try {
            Tasks.Tasks.patch(patchResource, tasklistId, googleTaskId);
            pushedFieldUpdatesCount++;
          } catch (ePatch) {
            Logger.log("Error updating Google Task fields (Row " + sheetRowIndex + "): " + ePatch);
          }
        }
      }
    }
  }

  // 3. Batch write back pushed google_task_id values
  if (pushedUpdates.length > 0) {
    for (var p = 0; p < pushedUpdates.length; p++) {
      var pUpdate = pushedUpdates[p];
      tasksSheet.getRange(pUpdate.row, pUpdate.col).setValue(pUpdate.googleTaskId);
    }
  }

  // 4. Batch write back status updates (ticked in Google Tasks: Done on, ADR 0045 / 0051)
  if (statusUpdates.length > 0) {
    for (var s = 0; s < statusUpdates.length; s++) {
      var sUpdate = statusUpdates[s];
      tasksSheet.getRange(sUpdate.row, sUpdate.col).setValue(sUpdate.status);
      if (typeof stampTaskDone === "function" && typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC.tasks.columns.indexOf("done_on") !== -1) {
        try { stampTaskDone(tasksSheet, sUpdate.row, sUpdate.status); } catch (eStamp) {}
      }
    }
  }

  // 5. Clear the IDs of Google Tasks removed from my list
  for (var rr = 0; rr < removedRows.length; rr++) {
    tasksSheet.getRange(removedRows[rr].row, removedRows[rr].col).setValue("");
  }

  // 6. Tasks this app made in my list that no row of mine points to any more
  var orphansRemoved = 0;
  if (listedOk && !options.allUsers) {
    for (var gid in existingGoogleTasksMap) {
      if (myReferencedIds[gid]) continue;
      var notes = String(existingGoogleTasksMap[gid].notes || "");
      if (notes.indexOf("Task ID: ") === -1) continue;
      try { Tasks.Tasks.remove(tasklistId, gid); orphansRemoved++; } catch (eOr) {}
    }
  }

  return {
    success: true,
    tasklistId: tasklistId,
    pushedCount: pushedCount,
    pushedFieldUpdatesCount: pushedFieldUpdatesCount,
    pulledUpdatesCount: pulledUpdatesCount,
    removedCount: removedRows.length + orphansRemoved,
    totalTasks: values.length
  };
}

/**
 * Public execution entry point to synchronize Google Tasks.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function syncTasks(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var res = syncGoogleTasks(ss);
  if (ss && typeof ss.toast === "function") {
    if (res.success) {
      ss.toast(
        "Google Tasks Synced: " + res.pushedCount + " new pushed, " + res.pulledUpdatesCount + " status updates.",
        "Tasks Synchronized",
        5
      );
    } else {
      ss.toast(res.message || "Failed to synchronize tasks", "Task Sync Error", 5);
    }
  }
  return res;
}

/**
 * Unlinks and deletes tasks from the active user's Google Tasks that were erroneously
 * synced for other members, clearing the google_task_id in tbl_tasks so the rightful
 * assignee can sync it.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Report of cleaned tasks
 */
function cleanupMisassignedGoogleTasks(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet available" };

  if (typeof Tasks === "undefined" || !Tasks.Tasks) {
    return { success: false, message: "Google Tasks API not enabled" };
  }

  var tasksSheet = getTable(ss, "tasks");
  if (!tasksSheet) return { success: false, message: "Sheet 'Tasks' not found" };

  var activeUserEmail = "";
  try {
    var aUser = Session.getActiveUser();
    activeUserEmail = aUser ? aUser.getEmail() : "";
  } catch (eUser) {}

  if (!activeUserEmail) {
    return { success: false, message: "Cannot determine active user identity." };
  }

  var activeMemberNames = [];
  var membersSheet = getTable(ss, "members");
  if (membersSheet && membersSheet.getMaxRows() > 1) {
    var mRows = membersSheet.getRange(2, 1, Math.min(membersSheet.getMaxRows() - 1, 300), Math.max(membersSheet.getLastColumn(), 7));
    var mVals = mRows.getValues();
    var mRich = (typeof mRows.getRichTextValues === "function") ? mRows.getRichTextValues() : [];
    for (var mi = 0; mi < mVals.length; mi++) {
      var mRaw = mVals[mi][0];
      var mR = (mRich && mRich[mi]) ? mRich[mi][0] : null;
      var mExtEmail = (typeof extractEmailFromChip === "function") ? extractEmailFromChip(mRaw, mR) : "";
      var mAccEmail = String(mVals[mi][1] || "").trim().toLowerCase();
      var matches = false;
      if (mExtEmail && mExtEmail.toLowerCase() === activeUserEmail.toLowerCase()) matches = true;
      else if (mAccEmail && mAccEmail === activeUserEmail.toLowerCase()) matches = true;
      else if (typeof isUserMemberMatch === "function" && isUserMemberMatch(mRaw, mR, activeUserEmail, mExtEmail)) matches = true;

      if (matches) {
        var mName = (typeof extractCellValue === "function") ? extractCellValue(mRaw, mR) : String(mRaw || "").trim();
        if (mName) activeMemberNames.push(mName.toLowerCase());
      }
    }
  }

  var colMap = getTaskColumnMap(tasksSheet);
  var tasklistId = getTasksListId(ss);
  var lastRow = tasksSheet.getLastRow();
  if (lastRow <= 1) return { success: true, cleanedCount: 0 };

  var numCols = Math.max(tasksSheet.getLastColumn(), 8);
  var range = tasksSheet.getRange(2, 1, lastRow - 1, numCols);
  var values = range.getValues();
  var richValues = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];

  var cleanedCount = 0;

  for (var r = 0; r < values.length; r++) {
    var sheetRow = r + 2;
    var row = values[r];
    var richRow = richValues[r] || [];

    var rawMember = row[colMap.assignedMember - 1];
    var richMember = (richRow && richRow[colMap.assignedMember - 1]) ? richRow[colMap.assignedMember - 1] : null;
    var assignedMember = (typeof extractCellValue === "function")
      ? extractCellValue(rawMember, richMember)
      : String(rawMember || "").trim();
    var googleTaskId = String(row[colMap.googleTaskId - 1] || "").trim();

    if (!googleTaskId) continue;

    var taskExtEmail = (typeof extractEmailFromChip === "function") ? extractEmailFromChip(rawMember, richMember) : "";
    var isUserTask = false;
    if (taskExtEmail && taskExtEmail.toLowerCase() === activeUserEmail.toLowerCase()) {
      isUserTask = true;
    } else if (typeof isUserMemberMatch === "function" && isUserMemberMatch(rawMember, richMember, activeUserEmail, taskExtEmail)) {
      isUserTask = true;
    } else if (activeMemberNames.length > 0 && assignedMember && activeMemberNames.indexOf(assignedMember.toLowerCase()) !== -1) {
      isUserTask = true;
    } else {
      isUserTask = false;
    }

    // If task has Google Task ID but does NOT belong to active user: Clean it up!
    if (!isUserTask) {
      try {
        if (typeof Tasks.Tasks.remove === "function") {
          Tasks.Tasks.remove(tasklistId, googleTaskId);
        } else if (typeof Tasks.Tasks.delete === "function") {
          Tasks.Tasks.delete(tasklistId, googleTaskId);
        }
      } catch (eDel) {
        Logger.log("Note removing task from Google Tasks: " + eDel);
      }
      // Clear google_task_id in sheet
      tasksSheet.getRange(sheetRow, colMap.googleTaskId).setValue("");
      cleanedCount++;
    }
  }

  if (ss && typeof ss.toast === "function") {
    ss.toast("Cleaned up " + cleanedCount + " tasks belonging to other members from your Google Tasks.", "Cleanup Completed", 5);
  }

  return {
    success: true,
    cleanedCount: cleanedCount
  };
}

/**
 * Automatically creates a Google Task for a given task row in tbl_tasks once task_id is created.
 * By default, the task_id is used as the initial title in Google Tasks if title is not yet set.
 * Defaults status if missing, respects member isolation,
 * and writes back the generated Google Task ID to the sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} tasksSheet
 * @param {number} rowIndex 1-indexed row number in tbl_tasks
 * @return {Object} Status of task creation
 */
function createGoogleTaskForTaskRow(tasksSheet, rowIndex) {
  if (!tasksSheet || rowIndex <= 1) {
    return { success: false, message: "Invalid sheet or header row" };
  }

  var colMap = getTaskColumnMap(tasksSheet);
  var numCols = Math.max(tasksSheet.getLastColumn(), 8);
  var rowVals = tasksSheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];
  var richRow = (typeof tasksSheet.getRange(rowIndex, 1, 1, numCols).getRichTextValues === "function")
    ? tasksSheet.getRange(rowIndex, 1, 1, numCols).getRichTextValues()[0]
    : [];

  var ss = tasksSheet.getParent ? tasksSheet.getParent() : SpreadsheetApp.getActiveSpreadsheet();

  // 1. Auto-generate Task ID if missing
  var taskId = String(rowVals[colMap.taskId - 1] || "").trim();
  if (!taskId) {
    taskId = generateTaskId();
    tasksSheet.getRange(rowIndex, colMap.taskId).setValue(taskId);
    rowVals[colMap.taskId - 1] = taskId;
    // A row typed by hand (ADR 0051)
    if (colMap.origin > 0 && !String(rowVals[colMap.origin - 1] || "").trim()) {
      tasksSheet.getRange(rowIndex, colMap.origin).setValue("manual");
    }
  }

  // 2. Default Status to Todo if missing
  var status = String(rowVals[colMap.status - 1] || "").trim();
  if (!status) {
    status = "Todo";
    tasksSheet.getRange(rowIndex, colMap.status).setValue(status);
    rowVals[colMap.status - 1] = status;
  }

  var projectId = String(rowVals[colMap.projectId - 1] || "").trim();
  var rawMember = rowVals[colMap.assignedMember - 1];
  var richMember = (richRow && richRow[colMap.assignedMember - 1]) ? richRow[colMap.assignedMember - 1] : null;
  var assignedMember = (typeof extractCellValue === "function")
    ? extractCellValue(rawMember, richMember)
    : String(rawMember || "").trim();

  // Active user session resolution
  var activeUserEmail = "";
  try {
    var aUser = Session.getActiveUser();
    activeUserEmail = aUser ? aUser.getEmail() : "";
  } catch (eU) {}

  // Google Task title defaults to taskId if title is not yet set
  var title = String(rowVals[colMap.title - 1] || "").trim();
  var effectiveTitle = title || taskId;
  if (!effectiveTitle) {
    return {
      success: false,
      inSheetCreated: true,
      taskId: taskId,
      message: "No title or task ID available."
    };
  }

  var googleTaskId = String(rowVals[colMap.googleTaskId - 1] || "").trim();
  if (googleTaskId) {
    return { success: true, alreadyExists: true, googleTaskId: googleTaskId, taskId: taskId };
  }

  // 4. Isolation Check (ADR 0021): Don't push other members' tasks into active user's Google Tasks
  var isAssignedToActiveUser = true;
  if (activeUserEmail && assignedMember) {
    var taskExtEmail = (typeof extractEmailFromChip === "function") ? extractEmailFromChip(rawMember, richMember) : "";
    if (taskExtEmail && taskExtEmail.toLowerCase() === activeUserEmail.toLowerCase()) {
      isAssignedToActiveUser = true;
    } else if (typeof isUserMemberMatch === "function" && isUserMemberMatch(rawMember, richMember, activeUserEmail, taskExtEmail)) {
      isAssignedToActiveUser = true;
    } else {
      isAssignedToActiveUser = false;
      if (ss) {
        var mSheetCheck = getTable(ss, "members");
        if (mSheetCheck && mSheetCheck.getLastRow() > 1) {
          var mValsCheck = mSheetCheck.getRange(2, 1, mSheetCheck.getLastRow() - 1, 3).getValues();
          for (var mc = 0; mc < mValsCheck.length; mc++) {
            if (typeof isUserMemberMatch === "function" && isUserMemberMatch(mValsCheck[mc][0], null, activeUserEmail, "")) {
              var aName = (typeof extractCellValue === "function") ? extractCellValue(mValsCheck[mc][0], null) : String(mValsCheck[mc][0] || "");
              if (aName && aName.toLowerCase() === assignedMember.toLowerCase()) {
                isAssignedToActiveUser = true;
                break;
              }
            }
          }
        }
      }
    }
  }

  if (!isAssignedToActiveUser) {
    return {
      success: true,
      skipped: true,
      taskId: taskId,
      message: "Task is assigned to teammate ('" + assignedMember + "'). It will sync when they trigger sync."
    };
  }

  // 5. Verify Google Tasks API availability
  if (typeof Tasks === "undefined" || !Tasks.Tasks) {
    return { success: false, inSheetCreated: true, taskId: taskId, message: "Google Tasks API not available" };
  }

  var tasklistId = getTasksListId(ss);
  var description = (colMap.description > 0) ? String(rowVals[colMap.description - 1] || "").trim() : "";
  var dueDate = rowVals[colMap.dueDate - 1];

  var taskResource = {
    title: effectiveTitle,
    notes: taskGoogleNotes(description, dueDate, taskId, projectId, assignedMember),
    status: mapSheetStatusToGoogleTaskStatus(status)
  };

  var dueStr = formatTaskDueRfc3339(taskGoogleDay(rowVals, colMap));
  if (dueStr) taskResource.due = dueStr;

  try {
    var created = Tasks.Tasks.insert(taskResource, tasklistId);
    if (created && created.id) {
      tasksSheet.getRange(rowIndex, colMap.googleTaskId).setValue(created.id);
      if (ss && typeof ss.toast === "function") {
        ss.toast("Google Task created for '" + effectiveTitle + "'", "Task Synchronized", 4);
      }
      return {
        success: true,
        googleTaskId: created.id,
        taskId: taskId,
        title: effectiveTitle
      };
    }
  } catch (eIns) {
    Logger.log("createGoogleTaskForTaskRow error: " + eIns);
    return { success: false, inSheetCreated: true, taskId: taskId, error: String(eIns) };
  }

  return { success: false, inSheetCreated: true, taskId: taskId, message: "Could not create task in Google Tasks" };
}

/**
 * Generates an incremented task ID when reopening or archiving a Done task.
 * Formats: TSK-YYMMDD-XXXX -> TSK-YYMMDD-XXXX.1, TSK-YYMMDD-XXXX.1 -> TSK-YYMMDD-XXXX.2
 *
 * @param {string} baseTaskId
 * @return {string} Incremented task ID
 */
function getIncrementedTaskId(baseTaskId) {
  var str = String(baseTaskId || "").trim();
  if (!str) return generateTaskId() + ".1";
  var match = str.match(/\.(\d+)$/);
  if (match) {
    var nextNum = parseInt(match[1], 10) + 1;
    return str.replace(/\.(\d+)$/, "." + nextNum);
  }
  return str + ".1";
}

/**
 * Handles attempting to modify the status of a task that is marked "Done".
 * Shows a warning dialog. If the user cancels, reverts the status to "Done".
 * If the user proceeds, hides the completed row and creates a new follow-up task
 * with the same metadata, an incremented Task ID (.<number>), and a new Google Task ID.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} editedRow
 * @param {string} taskId
 * @param {Array} rowVals
 * @param {Object|null} richMember
 * @param {Object} colMap
 * @param {string} newStatus
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} [e]
 * @return {Object} Result report
 */
function handleDoneStatusChange(sheet, editedRow, taskId, rowVals, richMember, colMap, newStatus, ss, e) {
  var proceed = true;
  if (e && e.proceed !== undefined) {
    proceed = !!e.proceed;
  } else {
    var ui = null;
    try {
      ui = SpreadsheetApp.getUi();
    } catch (eUi) {}
    if (ui && typeof ui.alert === "function") {
      var response = ui.alert(
        "Cannot Modify 'Done' Status",
        "You cannot modify a task that is marked 'Done'.\n\nWould you like to proceed? The current task will be archived and a new follow-up task will be created with the same metadata and an incremented Task ID.",
        ui.ButtonSet.YES_NO
      );
      var yesBtn = (ui.Button && ui.Button.YES) ? ui.Button.YES : "YES";
      proceed = (response === yesBtn);
    }
  }

  if (!proceed) {
    // Revert status back to Done
    sheet.getRange(editedRow, colMap.status).setValue("Done");
    if (ss && typeof ss.toast === "function") {
      ss.toast("Status change cancelled. Task remains 'Done'.", "Action Cancelled", 3);
    }
    return { success: false, cancelled: true, message: "Status change from Done cancelled" };
  }

  // User confirmed:
  // 1. Keep original row as Done and hide it
  sheet.getRange(editedRow, colMap.status).setValue("Done");
  try {
    sheet.hideRows(editedRow);
  } catch (eHide) {
    Logger.log("Could not hide row: " + eHide);
  }

  // 2. Generate incremented Task ID
  var newTaskId = getIncrementedTaskId(taskId);

  // 3. Prepare new row metadata
  var lastRow = sheet.getLastRow();
  var newRowIndex = lastRow + 1;
  var numCols = Math.max(sheet.getLastColumn(), 8);

  var projectId = String(rowVals[colMap.projectId - 1] || "").trim();
  var rawMember = rowVals[colMap.assignedMember - 1];
  var title = String(rowVals[colMap.title - 1] || "").trim();
  var description = (colMap.description > 0) ? String(rowVals[colMap.description - 1] || "").trim() : "";
  var dueDate = rowVals[colMap.dueDate - 1];
  var effectiveNewStatus = newStatus || "Todo";

  var newRowData = [];
  for (var c = 0; c < numCols; c++) {
    newRowData.push("");
  }
  newRowData[colMap.taskId - 1] = newTaskId;
  if (colMap.projectId > 0) newRowData[colMap.projectId - 1] = projectId;
  if (colMap.assignedMember > 0) newRowData[colMap.assignedMember - 1] = rawMember;
  if (colMap.title > 0) newRowData[colMap.title - 1] = title;
  if (colMap.description > 0) newRowData[colMap.description - 1] = description;
  if (colMap.dueDate > 0) newRowData[colMap.dueDate - 1] = dueDate;
  if (colMap.plannedDate > 0) newRowData[colMap.plannedDate - 1] = rowVals[colMap.plannedDate - 1];
  if (colMap.status > 0) newRowData[colMap.status - 1] = effectiveNewStatus;
  newRowData[colMap.googleTaskId - 1] = "";

  sheet.getRange(newRowIndex, 1, 1, numCols).setValues([newRowData]);

  // Copy People Chip if present
  if (colMap.assignedMember > 0 && richMember) {
    try {
      sheet.getRange(newRowIndex, colMap.assignedMember).setRichTextValue(richMember);
    } catch (eRich) {}
  }

  // Apply Typography
  if (typeof applyTypography === "function") {
    try {
      applyTypography(sheet.getRange(newRowIndex, 1, 1, numCols));
    } catch (eTypo) {}
  }

  // 4. Create new Google Task for the new row
  var createRes = createGoogleTaskForTaskRow(sheet, newRowIndex);

  if (ss && typeof ss.toast === "function") {
    ss.toast(
      "Archived completed task. Created follow-up task " + newTaskId + " with status '" + effectiveNewStatus + "'.",
      "Follow-up Task Created",
      5
    );
  }

  return {
    success: true,
    followUpCreated: true,
    archivedRow: editedRow,
    newRow: newRowIndex,
    newTaskId: newTaskId,
    newGoogleTaskId: createRes ? createRes.googleTaskId : null
  };
}

/**
 * Event handler called from onEdit when tbl_tasks is modified.
 * Whenever any cell in a task row is populated, immediately generates task_id,
 * defaults status to "Todo", and immediately creates the Google Task (defaulting title to task_id).
 * When the user subsequently edits or sets the real task title, description, due date, or status,
 * propagates the change to Google Tasks, replacing the default title with the real title.
 * If trying to modify status from "Done", presents a warning and creates a follow-up task upon confirmation.
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {Object|null}
 */
function handleTaskEdit(e) {
  if (!e || !e.range) return null;
  var sheet = e.range.getSheet();
  if (!sheet || tableKeyOfSheet(sheet) !== "tasks") return null;

  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();
  if (editedRow <= 1) return null; // Header row

  var colMap = getTaskColumnMap(sheet);
  var numCols = Math.max(sheet.getLastColumn(), 8);
  var rowVals = sheet.getRange(editedRow, 1, 1, numCols).getValues()[0];

  // Check if the row has ANY non-empty value
  var hasAnyValue = false;
  for (var i = 0; i < rowVals.length; i++) {
    var v = rowVals[i];
    if (v !== "" && v !== null && v !== undefined && String(v).trim() !== "") {
      hasAnyValue = true;
      break;
    }
  }
  if (!hasAnyValue) return null;

  var ss = sheet.getParent ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet();

  // 1. Auto-generate Task ID immediately if missing
  var taskId = String(rowVals[colMap.taskId - 1] || "").trim();
  if (!taskId) {
    taskId = generateTaskId();
    sheet.getRange(editedRow, colMap.taskId).setValue(taskId);
    rowVals[colMap.taskId - 1] = taskId;
  }

  // 2. Default Status to Todo immediately if missing
  var status = String(rowVals[colMap.status - 1] || "").trim();
  if (!status) {
    status = "Todo";
    sheet.getRange(editedRow, colMap.status).setValue(status);
    rowVals[colMap.status - 1] = status;
  }

  // 3. Extract assigned member if present (assigned member is not filled automatically)
  var rawMember = rowVals[colMap.assignedMember - 1];
  var richMember = (typeof sheet.getRange(editedRow, colMap.assignedMember).getRichTextValue === "function")
    ? sheet.getRange(editedRow, colMap.assignedMember).getRichTextValue()
    : null;
  var assignedMember = (typeof extractCellValue === "function")
    ? extractCellValue(rawMember, richMember)
    : String(rawMember || "").trim();

  // 4. Apply Typography formatting to row
  if (typeof applyTypography === "function") {
    try {
      applyTypography(sheet.getRange(editedRow, 1, 1, numCols));
    } catch (eTypo) {}
  }

  var title = String(rowVals[colMap.title - 1] || "").trim();
  var googleTaskId = String(rowVals[colMap.googleTaskId - 1] || "").trim();

  // Check if user is attempting to modify the status of a task that was marked "Done"
  var oldStatus = (e && e.oldValue !== undefined) ? String(e.oldValue).trim() : "";
  var wasDone = (oldStatus === "Done");
  if (!wasDone && !oldStatus && googleTaskId && typeof Tasks !== "undefined" && Tasks.Tasks) {
    try {
      var tListId = getTasksListId(ss);
      var curGTask = Tasks.Tasks.get(tListId, googleTaskId);
      if (curGTask && curGTask.status === "completed") {
        wasDone = true;
      }
    } catch (eGGet) {}
  }

  if (editedCol === colMap.status && wasDone && status !== "Done") {
    return handleDoneStatusChange(sheet, editedRow, taskId, rowVals, richMember, colMap, status, ss, e);
  }

  // Case 1: Unsynced task: automatically create Google Task immediately once task_id exists!
  // By default, the task_id is used as the initial title in Google Tasks.
  if (!googleTaskId && taskId) {
    try {
      return createGoogleTaskForTaskRow(sheet, editedRow);
    } catch (eGTask) {
      Logger.log("createGoogleTaskForTaskRow error: " + eGTask);
      return { success: false, taskId: taskId, inSheetCreated: true, error: String(eGTask) };
    }
  }

  // Case 2: Synced task (has google_task_id) and user edited title, description, due date, status, etc.
  if (googleTaskId && (editedCol === colMap.title || (colMap.description > 0 && editedCol === colMap.description) || editedCol === colMap.dueDate || editedCol === colMap.status || editedCol === colMap.projectId || editedCol === colMap.assignedMember)) {
    if (typeof Tasks !== "undefined" && Tasks.Tasks) {
      try {
        var tasklistId = getTasksListId(ss);
        var patchResource = {};

        if (editedCol === colMap.title) {
          // Replace default task_id title with real title, or fallback to taskId if cleared
          patchResource.title = title || taskId;
        }
        if (editedCol === colMap.status) {
          patchResource.status = mapSheetStatusToGoogleTaskStatus(status);
        }
        if (editedCol === colMap.dueDate) {
          var dueDate = rowVals[colMap.dueDate - 1];
          var rfc = formatTaskDueRfc3339(dueDate);
          patchResource.due = rfc;
        }
        if (editedCol === colMap.description || editedCol === colMap.title || editedCol === colMap.projectId || editedCol === colMap.assignedMember) {
          var desc = (colMap.description > 0) ? String(rowVals[colMap.description - 1] || "").trim() : "";
          var tId = taskId;
          var pId = String(rowVals[colMap.projectId - 1] || "").trim();
          var rawM = rowVals[colMap.assignedMember - 1];
          var aM = (typeof extractCellValue === "function") ? extractCellValue(rawM, null) : String(rawM || "").trim();
          var nParts = [];
          if (desc) nParts.push(desc);
          var mParts = [];
          if (tId) mParts.push("Task ID: " + tId);
          if (pId) mParts.push("Project: " + pId);
          if (aM) mParts.push("Assigned: " + aM);
          if (mParts.length > 0) nParts.push(mParts.join(" | "));
          patchResource.notes = nParts.join("\n\n");
        }

        Tasks.Tasks.patch(patchResource, tasklistId, googleTaskId);
        return { success: true, patched: true, googleTaskId: googleTaskId, patchedFields: patchResource };
      } catch (ePatch) {
        Logger.log("handleTaskEdit patch error: " + ePatch);
      }
    }
  }

  return { success: true, inSheetCreated: true, taskId: taskId };
}

/**
 * Installs an installable onEdit trigger for live real-time task synchronization.
 * Installable triggers run with the full OAuth authority of the user, allowing Tasks API calls onEdit.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function installTaskSyncTrigger(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof ScriptApp === "undefined" || !ScriptApp.getProjectTriggers) {
    return { success: false, message: "ScriptApp not available" };
  }

  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "installedTaskEditHandler") {
      if (ss && typeof ss.toast === "function") {
        ss.toast("Live task sync trigger is already active.", "Live Sync Active", 4);
      }
      return { success: true, alreadyInstalled: true };
    }
  }

  ScriptApp.newTrigger("installedTaskEditHandler")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  if (ss && typeof ss.toast === "function") {
    ss.toast("Live task sync trigger installed successfully! Real-time sync is now enabled.", "Live Sync Enabled", 5);
  }

  return { success: true, installed: true };
}

// ---------------------------------------------------------------------------
// Each person's own Google Tasks (ADR 0045)
// A script writes only into the Google Tasks of the account it runs as, so each
// member connects once: an hourly job under their own account syncs the tasks
// assigned to them (both ways).
// ---------------------------------------------------------------------------

var TASKS_CONNECTED_PROP = "TASKS_CONNECTED_USERS";

function tsEffectiveEmail() {
  try { var u = Session.getEffectiveUser(); if (u && u.getEmail()) return u.getEmail(); } catch (e1) {}
  try { var a = Session.getActiveUser(); if (a && a.getEmail()) return a.getEmail(); } catch (e2) {}
  return "";
}

/** Emails of the people who connected their Google Tasks. */
function connectedTaskUsers() {
  try {
    return JSON.parse(PropertiesService.getDocumentProperties().getProperty(TASKS_CONNECTED_PROP) || "[]") || [];
  } catch (e) { return []; }
}

/** Menu: Connect my Google Tasks. Installs the hourly job for the person who clicks, then syncs once. */
function connectMyGoogleTasks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var email = tsEffectiveEmail();
  var has = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === "myTasksSyncJob"; });
  if (!has) ScriptApp.newTrigger("myTasksSyncJob").timeBased().everyHours(1).create();
  if (email) {
    var list = connectedTaskUsers();
    if (list.indexOf(email.toLowerCase()) === -1) {
      list.push(email.toLowerCase());
      PropertiesService.getDocumentProperties().setProperty(TASKS_CONNECTED_PROP, JSON.stringify(list));
    }
  }
  var res = myTasksSyncJob();
  try {
    ss.toast("Your tasks now sync with your Google Tasks every hour (list \"Project Management Tasks\")." +
      (res && res.pushedCount ? " " + res.pushedCount + " task(s) added." : ""), "Google Tasks connected", 8);
  } catch (eT) {}
  return { success: true, email: email, alreadyConnected: has, sync: res };
}

/** Hourly, under each connected person's account: sync their tasks with their Google Tasks. */
function myTasksSyncJob() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lock = (typeof acquireSyncLock === "function") ? acquireSyncLock(20000) : null;
  if (typeof acquireSyncLock === "function" && !lock) return { success: false, reason: "sync_in_progress" };
  try {
    var res = syncGoogleTasks(ss, { userEmail: tsEffectiveEmail() });
    if (res && res.pulledUpdatesCount && typeof renderTaskCalendar === "function") {
      try { renderTaskCalendar(ss); } catch (eCal) {}
    }
    return res;
  } finally {
    if (lock) lock.releaseLock();
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    GOOGLE_TASKS_LIST_ID_KEY: GOOGLE_TASKS_LIST_ID_KEY,
    listAllGoogleTasks: listAllGoogleTasks,
    connectedTaskUsers: connectedTaskUsers,
    connectMyGoogleTasks: connectMyGoogleTasks,
    myTasksSyncJob: myTasksSyncJob,
    generateTaskId: generateTaskId,
    getIncrementedTaskId: getIncrementedTaskId,
    getTasksListId: getTasksListId,
    setTasksListId: setTasksListId,
    formatTaskDueRfc3339: formatTaskDueRfc3339,
    mapSheetStatusToGoogleTaskStatus: mapSheetStatusToGoogleTaskStatus,
    mapGoogleTaskStatusToSheetStatus: mapGoogleTaskStatusToSheetStatus,
    getTaskColumnMap: getTaskColumnMap,
    syncGoogleTasks: syncGoogleTasks,
    syncTasks: syncTasks,
    cleanupMisassignedGoogleTasks: cleanupMisassignedGoogleTasks,
    createGoogleTaskForTaskRow: createGoogleTaskForTaskRow,
    handleDoneStatusChange: handleDoneStatusChange,
    handleTaskEdit: handleTaskEdit,
    installTaskSyncTrigger: installTaskSyncTrigger
  };
}
