/**
 * Project Management Workspace - Project Backup Conflict Engine
 *
 * Verifies that lead_member and backup_member on any active project in tbl_projects
 * are not absent simultaneously (overlapping approved leaves or current absence).
 *
 * If a conflict is detected:
 * 1. Highlights the project row in alert red (#F8D7DA background, #721C24 bold text).
 * 2. Sends an email notification via MailApp to the lead member and project manager.
 * 3. Creates an interim backup organization task in tbl_tasks assigning the lead member,
 *    or the backup member if the lead member is already absent.
 * 4. Resolves the conflict if an active designated backup exists in tbl_backups covering the range.
 */

/**
 * Normalizes Date or date string to YYYY-MM-DD.
 *
 * @param {Date|string} d
 * @return {string}
 */
function formatBackupDateIso(d) {
  if (!d) return "";
  if (typeof d === "string") {
    return d.indexOf("T") !== -1 ? d.split("T")[0] : d.trim();
  }
  try {
    var y = d.getFullYear();
    var m = ("0" + (d.getMonth() + 1)).slice(-2);
    var day = ("0" + d.getDate()).slice(-2);
    return y + "-" + m + "-" + day;
  } catch (e) {
    return String(d);
  }
}

/**
 * Parses date safely to Date object set at 00:00:00.
 *
 * @param {Date|string} d
 * @return {Date|null}
 */
function parseBackupDate(d) {
  if (!d) return null;
  if (d instanceof Date) {
    var dt = new Date(d.getTime());
    dt.setHours(0, 0, 0, 0);
    return dt;
  }
  var s = String(d).trim();
  var match = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10), 0, 0, 0, 0);
  }
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    parsed.setHours(0, 0, 0, 0);
    return parsed;
  }
  return null;
}

/**
 * Scans tbl_projects and detects simultaneous absence conflicts between lead_member and backup_member.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Object} [options]
 * @param {Date|string} [options.asOfDateParam] - Reference date for today
 * @param {string} [options.managerEmail] - Explicit manager email override
 * @param {boolean} [options.sendNotifications] - Default true
 * @param {boolean} [options.createTasks] - Default true
 * @return {Object} Audit report with conflicts found and actions taken
 */
function checkProjectBackupConflicts(targetSpreadsheet, options) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet" };

  var projectsSheet = ss.getSheetByName("tbl_projects");
  var leavesSheet = ss.getSheetByName("tbl_leaves");
  var backupsSheet = ss.getSheetByName("tbl_backups");
  var tasksSheet = ss.getSheetByName("tbl_tasks");
  var membersSheet = ss.getSheetByName("tbl_members");

  if (!projectsSheet) {
    return { success: false, message: "Sheet 'tbl_projects' not found." };
  }

  var opt = options || {};
  var sendNotifications = opt.sendNotifications !== false;
  var createTasks = opt.createTasks !== false;

  var now = opt.asOfDateParam ? new Date(opt.asOfDateParam) : new Date();
  var todayDate = new Date(now.getTime());
  todayDate.setHours(0, 0, 0, 0);
  var todayIso = formatBackupDateIso(todayDate);

  // 1. Build Members map (for emails, chips, roles, and today presence)
  var membersRoster = {};
  var managerCandidateEmail = opt.managerEmail || "";

  if (membersSheet) {
    var mMax = Math.min(membersSheet.getMaxRows(), 300);
    if (mMax > 1) {
      var mRange = membersSheet.getRange(2, 1, mMax - 1, Math.max(membersSheet.getLastColumn(), 7));
      var mVals = mRange.getValues();
      var mRich = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : [];

      for (var m = 0; m < mVals.length; m++) {
        var rawM = mVals[m][0];
        var rM = (mRich && mRich[m]) ? mRich[m][0] : null;
        var nameM = (typeof extractCellValue === "function") ? extractCellValue(rawM, rM) : String(rawM || "").trim();
        if (!nameM) continue;

        var emailM = String(mVals[m][1] || "").trim().toLowerCase();
        if (!emailM && typeof extractEmailFromChip === "function") {
          emailM = extractEmailFromChip(rawM, rM).toLowerCase();
        }
        var roleM = String(mVals[m][3] || "").trim();
        var deptM = String(mVals[m][4] || "").trim();
        var tenureM = String(mVals[m][5] || "").trim();
        var todayPresence = String(mVals[m][6] || "").trim();

        var mObj = {
          row: m + 2,
          raw: rawM,
          rich: rM,
          name: nameM,
          email: emailM,
          role: roleM,
          dept: deptM,
          tenure: tenureM,
          todayPresence: todayPresence
        };

        membersRoster[nameM.toLowerCase()] = mObj;
        if (emailM) membersRoster[emailM] = mObj;

        // Check if member has a Manager / Lead role
        if (!managerCandidateEmail && emailM && /manager|head|director/i.test(roleM)) {
          managerCandidateEmail = emailM;
        }
      }
    }
  }

  // Fallback manager email if none found in members roster
  if (!managerCandidateEmail) {
    try {
      managerCandidateEmail = Session.getActiveUser().getEmail();
    } catch (eSes) {}
  }
  if (!managerCandidateEmail) {
    try {
      if (typeof ss.getOwner === "function" && ss.getOwner()) {
        managerCandidateEmail = ss.getOwner().getEmail();
      }
    } catch (eOwn) {}
  }

  // 2. Read Approved Leaves from tbl_leaves
  var leavesByMember = {};
  if (leavesSheet) {
    var lColMap = (typeof getLeaveColumnMap === "function")
      ? getLeaveColumnMap(leavesSheet)
      : { leave_id: 1, member: 2, start_date: 3, start_period: -1, end_date: 4, end_period: -1, leave_type: 5, status: 6, numCols: 6 };
    var lMax = Math.min(leavesSheet.getMaxRows(), 500);
    if (lMax > 1) {
      var lRange = leavesSheet.getRange(2, 1, lMax - 1, Math.max(leavesSheet.getLastColumn(), 6));
      var lVals = lRange.getValues();
      var lRich = (typeof lRange.getRichTextValues === "function") ? lRange.getRichTextValues() : [];

      for (var l = 0; l < lVals.length; l++) {
        var rawLMem = lVals[l][lColMap.member - 1];
        var rLMem = (lRich && lRich[l]) ? lRich[l][lColMap.member - 1] : null;
        var lMem = (typeof extractCellValue === "function") ? extractCellValue(rawLMem, rLMem) : String(rawLMem || "").trim();
        if (!lMem) continue;

        var lStatus = String(lVals[l][lColMap.status - 1] || "").trim().toLowerCase();
        if (lStatus !== "approved") continue;

        var lKey = lMem.toLowerCase();
        if (!leavesByMember[lKey]) leavesByMember[lKey] = [];

        var sDate = parseBackupDate(lVals[l][lColMap.start_date - 1]);
        var eDate = parseBackupDate(lVals[l][lColMap.end_date - 1]);
        if (!sDate) continue;
        if (!eDate) eDate = new Date(sDate.getTime());

        leavesByMember[lKey].push({
          startDate: sDate,
          endDate: eDate,
          leaveType: String(lVals[l][lColMap.leave_type - 1] || "Vacation").trim()
        });
      }
    }
  }

  // 3. Read Active Designated Backups from tbl_backups
  var backupsByProject = {};
  if (backupsSheet) {
    var bMax = Math.min(backupsSheet.getMaxRows(), 500);
    if (bMax > 1) {
      // Columns: backup_id(1), project_id(2), primary_member(3), backup_member(4), start_date(5), end_date(6), status(7), notes(8)
      var bVals = backupsSheet.getRange(2, 1, bMax - 1, Math.max(backupsSheet.getLastColumn(), 8)).getValues();
      var bRich = (typeof backupsSheet.getRange(2, 1, bMax - 1, Math.max(backupsSheet.getLastColumn(), 8)).getRichTextValues === "function")
        ? backupsSheet.getRange(2, 1, bMax - 1, Math.max(backupsSheet.getLastColumn(), 8)).getRichTextValues()
        : [];

      for (var b = 0; b < bVals.length; b++) {
        var bProjId = String(bVals[b][1] || "").trim();
        if (!bProjId) continue;

        var bStatus = String(bVals[b][6] || "").trim().toLowerCase();
        if (bStatus !== "active" && bStatus !== "planned") continue;

        var bPrimaryRaw = bVals[b][2];
        var bPrimary = (typeof extractCellValue === "function") ? extractCellValue(bPrimaryRaw) : String(bPrimaryRaw || "").trim();
        var bBackupRaw = bVals[b][3];
        var bBackup = (typeof extractCellValue === "function") ? extractCellValue(bBackupRaw) : String(bBackupRaw || "").trim();

        var bStart = parseBackupDate(bVals[b][4]);
        var bEnd = parseBackupDate(bVals[b][5]);

        if (!backupsByProject[bProjId]) backupsByProject[bProjId] = [];
        backupsByProject[bProjId].push({
          primaryMember: bPrimary,
          backupMember: bBackup,
          startDate: bStart,
          endDate: bEnd,
          status: bVals[b][6]
        });
      }
    }
  }

  // 4. Read Existing Tasks from tbl_tasks to ensure idempotent task creation
  var existingBackupTasks = {}; // key: projectId -> true
  if (tasksSheet) {
    var tMax = Math.min(tasksSheet.getMaxRows(), 500);
    if (tMax > 1) {
      // Columns: task_id(1), project_id(2), assigned_member(3), title(4), description(5), due_date(6), status(7)
      var tVals = tasksSheet.getRange(2, 1, tMax - 1, Math.max(tasksSheet.getLastColumn(), 7)).getValues();
      for (var t = 0; t < tVals.length; t++) {
        var tProj = String(tVals[t][1] || "").trim();
        var tTitle = String(tVals[t][3] || "").trim().toLowerCase();
        var tStatus = String(tVals[t][6] || "").trim().toLowerCase();
        if (tProj && tTitle.indexOf("organize interim backup") !== -1 && tStatus !== "done") {
          existingBackupTasks[tProj] = true;
        }
      }
    }
  }

  // Helper to test if a member is currently absent
  function isMemberCurrentlyAbsent(mKey) {
    var mR = membersRoster[mKey];
    if (mR && mR.todayPresence) {
      var pres = mR.todayPresence.toLowerCase();
      if (pres.indexOf("vacation") !== -1 || pres.indexOf("sick") !== -1 || pres.indexOf("leave") !== -1 || pres.indexOf("absent") !== -1) {
        return true;
      }
    }
    var lList = leavesByMember[mKey] || [];
    for (var i = 0; i < lList.length; i++) {
      if (todayDate >= lList[i].startDate && todayDate <= lList[i].endDate) {
        return true;
      }
    }
    return false;
  }

  // Helper to check if designated backup resolves an overlap period on project
  function isConflictMitigated(projId, overlapStart, overlapEnd) {
    var bList = backupsByProject[projId] || [];
    for (var i = 0; i < bList.length; i++) {
      var b = bList[i];
      if (!b.backupMember) continue;
      // If designated backup has dates, ensure they cover or overlap
      if (b.startDate && b.startDate > overlapStart) continue;
      if (b.endDate && b.endDate < overlapEnd) continue;

      // Check if designated backup is also absent
      var bKey = b.backupMember.toLowerCase();
      var bLeaves = leavesByMember[bKey] || [];
      var bAbsent = false;
      for (var bl = 0; bl < bLeaves.length; bl++) {
        if (bLeaves[bl].startDate <= overlapEnd && bLeaves[bl].endDate >= overlapStart) {
          bAbsent = true;
          break;
        }
      }
      if (!bAbsent) {
        return true;
      }
    }
    return false;
  }

  // 5. Scan Projects in tbl_projects
  // Columns: project_id(1), name(2), status(3), start_date(4), target_date(5), lead_member(6), backup_member(7)
  var pMax = Math.min(projectsSheet.getMaxRows(), 500);
  if (pMax <= 1) {
    return { success: true, projectsAudited: 0, conflictsFound: 0, message: "No project rows in tbl_projects." };
  }

  var pRange = projectsSheet.getRange(2, 1, pMax - 1, Math.max(projectsSheet.getLastColumn(), 7));
  var pVals = pRange.getValues();
  var pRich = (typeof pRange.getRichTextValues === "function") ? pRange.getRichTextValues() : [];

  var conflictsList = [];
  var resolvedCount = 0;

  for (var p = 0; p < pVals.length; p++) {
    var rowNum = p + 2;
    var projId = String(pVals[p][0] || "").trim();
    var projName = String(pVals[p][1] || projId).trim();
    var projStatus = String(pVals[p][2] || "").trim();
    if (!projId && !projName) continue;

    // Completed projects do not need backup conflict alerting
    if (projStatus.toLowerCase() === "completed") {
      projectsSheet.getRange(rowNum, 1, 1, 7).setBackground("#FFFFFF").setFontColor("#000000").setFontWeight("normal");
      continue;
    }

    var leadRaw = pVals[p][5];
    var leadR = (pRich && pRich[p]) ? pRich[p][5] : null;
    var leadName = (typeof extractCellValue === "function") ? extractCellValue(leadRaw, leadR) : String(leadRaw || "").trim();

    var backupRaw = pVals[p][6];
    var backupR = (pRich && pRich[p]) ? pRich[p][6] : null;
    var backupName = (typeof extractCellValue === "function") ? extractCellValue(backupRaw, backupR) : String(backupRaw || "").trim();

    if (!leadName || !backupName) {
      projectsSheet.getRange(rowNum, 1, 1, 7).setBackground("#FFFFFF").setFontColor("#000000").setFontWeight("normal");
      continue;
    }

    var leadKey = leadName.toLowerCase();
    var backupKey = backupName.toLowerCase();

    var leadLeaves = leavesByMember[leadKey] || [];
    var backupLeaves = leavesByMember[backupKey] || [];

    var detectedConflict = null;

    // A. Check for overlapping approved leaves
    for (var ll = 0; ll < leadLeaves.length; ll++) {
      var lLeave = leadLeaves[ll];
      for (var bl = 0; bl < backupLeaves.length; bl++) {
        var bLeave = backupLeaves[bl];
        var ovStart = new Date(Math.max(lLeave.startDate.getTime(), bLeave.startDate.getTime()));
        var ovEnd = new Date(Math.min(lLeave.endDate.getTime(), bLeave.endDate.getTime()));

        if (ovStart <= ovEnd) {
          detectedConflict = {
            startDate: ovStart,
            endDate: ovEnd,
            reason: "Overlapping approved leaves (" + formatBackupDateIso(ovStart) + " to " + formatBackupDateIso(ovEnd) + ")"
          };
          break;
        }
      }
      if (detectedConflict) break;
    }

    // B. Check if both are absent today
    if (!detectedConflict && isMemberCurrentlyAbsent(leadKey) && isMemberCurrentlyAbsent(backupKey)) {
      detectedConflict = {
        startDate: todayDate,
        endDate: todayDate,
        reason: "Simultaneous absence today (" + todayIso + ")"
      };
    }

    // C. Check if conflict is mitigated by an active entry in tbl_backups
    var isMitigated = false;
    if (detectedConflict) {
      isMitigated = isConflictMitigated(projId, detectedConflict.startDate, detectedConflict.endDate);
    }

    if (detectedConflict && !isMitigated) {
      // === CONFLICT DETECTED ===
      var targetRowRange = projectsSheet.getRange(rowNum, 1, 1, 7);
      var currentBg = String(targetRowRange.getBackground() || "").toUpperCase();
      var isAlreadyAlerted = (currentBg === "#F8D7DA");

      // 1. Highlight project row in alert red (#F8D7DA, #721C24 bold)
      targetRowRange.setBackground("#F8D7DA");
      targetRowRange.setFontColor("#721C24");
      targetRowRange.setFontWeight("bold");

      var conflictStartStr = formatBackupDateIso(detectedConflict.startDate);
      var conflictEndStr = formatBackupDateIso(detectedConflict.endDate);

      // Resolve lead member email and manager email
      var leadEmail = "";
      if (membersRoster[leadKey] && membersRoster[leadKey].email) {
        leadEmail = membersRoster[leadKey].email;
      } else if (typeof extractEmailFromChip === "function") {
        leadEmail = extractEmailFromChip(leadRaw, leadR);
      }

      var recipientsList = [];
      if (leadEmail) recipientsList.push(leadEmail);
      if (managerCandidateEmail && recipientsList.indexOf(managerCandidateEmail.toLowerCase()) === -1) {
        recipientsList.push(managerCandidateEmail);
      }

      // 2. Send email notification (only if newly alerted to avoid duplicate spam)
      var emailSent = false;
      if (sendNotifications && !isAlreadyAlerted && recipientsList.length > 0 && typeof MailApp !== "undefined" && typeof MailApp.sendEmail === "function") {
        try {
          var subject = "[ACTION REQUIRED] Simultaneous Absence Conflict on Project: " + projName;
          var body = "Attention Required: Simultaneous Absence Conflict Detected\n\n" +
            "Project: " + projName + " (ID: " + projId + ")\n" +
            "Lead Member: " + leadName + "\n" +
            "Backup Member: " + backupName + "\n" +
            "Conflict Period: " + conflictStartStr + " to " + conflictEndStr + "\n" +
            "Details: " + detectedConflict.reason + "\n\n" +
            "Both the project lead and backup member are absent simultaneously without a designated replacement in tbl_backups.\n" +
            "Please assign an interim backup member in the 'tbl_backups' table immediately to ensure continuity.";
          MailApp.sendEmail({
            to: recipientsList.join(","),
            subject: subject,
            body: body
          });
          emailSent = true;
        } catch (eMail) {
          Logger.log("Failed to send conflict email: " + eMail);
        }
      }

      // 3. Create interim backup task in tbl_tasks
      var taskCreated = false;
      var assignedMemberForTask = leadRaw; // Default to lead
      var isLeadAbsentNow = isMemberCurrentlyAbsent(leadKey) || (todayDate >= detectedConflict.startDate);
      if (isLeadAbsentNow) {
        assignedMemberForTask = backupRaw; // Assign backup member if lead is already absent
      }

      if (createTasks && tasksSheet && !existingBackupTasks[projId]) {
        try {
          var tLast = (typeof tasksSheet.getLastRow === "function") ? (tasksSheet.getLastRow() || 1) : 1;
          var tTargetRow = tLast + 1;
          if (tTargetRow <= 2) tTargetRow = 2;

          var taskId = "task_" + Math.random().toString(36).substring(2, 9);
          var taskTitle = "Organize interim backup for project " + projName;
          var taskDesc = "Simultaneous absence conflict between " + leadName + " and " + backupName + " from " + conflictStartStr + " to " + conflictEndStr + ". Assign an interim backup in tbl_backups.";

          var newTaskRow = [
            taskId,
            projId,
            (typeof extractCellValue === "function") ? extractCellValue(assignedMemberForTask) : leadName,
            taskTitle,
            taskDesc,
            conflictStartStr,
            "Todo",
            ""
          ];

          tasksSheet.getRange(tTargetRow, 1, 1, 8).setValues([newTaskRow]);

          // Copy People Chip if available
          var assignedNameStr = (typeof extractCellValue === "function") ? extractCellValue(assignedMemberForTask) : leadName;
          var assignRoster = membersRoster[assignedNameStr.toLowerCase()];
          if (assignRoster && typeof copyCellChip === "function") {
            copyCellChip(
              membersSheet,
              assignRoster.row,
              1,
              tasksSheet,
              tTargetRow,
              3,
              assignedNameStr,
              ss.getId()
            );
          }

          if (typeof applyTypography === "function") {
            applyTypography(tasksSheet.getRange(tTargetRow, 1, 1, 8));
          }

          existingBackupTasks[projId] = true;
          taskCreated = true;
        } catch (eTask) {
          Logger.log("Failed to create backup task: " + eTask);
        }
      }

      conflictsList.push({
        projectId: projId,
        projectName: projName,
        row: rowNum,
        leadMember: leadName,
        backupMember: backupName,
        conflictStart: conflictStartStr,
        conflictEnd: conflictEndStr,
        reason: detectedConflict.reason,
        emailSent: emailSent,
        recipients: recipientsList,
        taskCreated: taskCreated
      });
    } else {
      // === NO CONFLICT / MITIGATED ===
      // Clear alert red formatting
      projectsSheet.getRange(rowNum, 1, 1, 7)
        .setBackground("#FFFFFF")
        .setFontColor("#000000")
        .setFontWeight("normal");
      resolvedCount++;
    }
  }

  return {
    success: true,
    projectsAudited: pVals.length,
    conflictsFound: conflictsList.length,
    conflicts: conflictsList,
    resolvedCount: resolvedCount
  };
}

/**
 * Assigns an interim backup in tbl_backups and automatically re-evaluates project conflict formatting.
 *
 * @param {Object} payload
 * @param {string} payload.projectId
 * @param {string} payload.primaryMember
 * @param {string} payload.backupMember
 * @param {string} payload.startDate
 * @param {string} [payload.endDate]
 * @param {string} [payload.status] - Default "Active"
 * @param {string} [payload.notes]
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function assignProjectBackup(payload, targetSpreadsheet) {
  if (!payload || !payload.projectId || !payload.primaryMember || !payload.backupMember || !payload.startDate) {
    return { success: false, message: "Missing required fields (projectId, primaryMember, backupMember, startDate)." };
  }

  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet" };

  var backupsSheet = ss.getSheetByName("tbl_backups");
  var membersSheet = ss.getSheetByName("tbl_members");

  if (!backupsSheet) {
    return { success: false, message: "Sheet 'tbl_backups' not found. Please run init() first." };
  }

  var backupId = "bkp_" + Math.random().toString(36).substring(2, 9);
  var startDateStr = formatBackupDateIso(payload.startDate);
  var endDateStr = payload.endDate ? formatBackupDateIso(payload.endDate) : "";
  var status = payload.status || "Active";
  var notes = payload.notes || "";

  var bLast = (typeof backupsSheet.getLastRow === "function") ? (backupsSheet.getLastRow() || 1) : 1;
  var targetRow = bLast + 1;
  if (targetRow <= 2) targetRow = 2;

  var newRow = [
    backupId,
    payload.projectId,
    payload.primaryMember,
    payload.backupMember,
    startDateStr,
    endDateStr,
    status,
    notes
  ];

  backupsSheet.getRange(targetRow, 1, 1, 8).setValues([newRow]);

  // Copy People Chips from tbl_members if available
  if (membersSheet && typeof copyCellChip === "function") {
    var mMax = Math.min(membersSheet.getMaxRows(), 300);
    var mVals = membersSheet.getRange(2, 1, mMax - 1, 1).getValues();
    var primaryRow = -1;
    var backupRow = -1;

    for (var m = 0; m < mVals.length; m++) {
      var name = (typeof extractCellValue === "function") ? extractCellValue(mVals[m][0]).toLowerCase() : String(mVals[m][0] || "").toLowerCase();
      if (name === String(payload.primaryMember).toLowerCase()) primaryRow = m + 2;
      if (name === String(payload.backupMember).toLowerCase()) backupRow = m + 2;
    }

    if (primaryRow > 0) {
      copyCellChip(membersSheet, primaryRow, 1, backupsSheet, targetRow, 3, payload.primaryMember, ss.getId());
    }
    if (backupRow > 0) {
      copyCellChip(membersSheet, backupRow, 1, backupsSheet, targetRow, 4, payload.backupMember, ss.getId());
    }
  }

  if (typeof applyTypography === "function") {
    applyTypography(backupsSheet.getRange(targetRow, 1, 1, 8));
  }

  // Re-evaluate conflicts across projects to clear alert red formatting if resolved
  var conflictReport = null;
  try {
    conflictReport = checkProjectBackupConflicts(ss);
  } catch (eConf) {
    Logger.log("Conflict re-check note: " + eConf);
  }

  return {
    success: true,
    backupId: backupId,
    projectId: payload.projectId,
    primaryMember: payload.primaryMember,
    backupMember: payload.backupMember,
    startDate: startDateStr,
    endDate: endDateStr,
    status: status,
    conflictReport: conflictReport
  };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    checkProjectBackupConflicts: checkProjectBackupConflicts,
    assignProjectBackup: assignProjectBackup,
    formatBackupDateIso: formatBackupDateIso,
    parseBackupDate: parseBackupDate
  };
}
