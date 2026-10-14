/**
 * Project Management Workspace - Project Backup Coverage Engine (ADR 0038)
 *
 * For every active project and every working day from today on, the responsible person is:
 *   the Project Lead if present -> else an interim backup active that day (dated tbl_backups row)
 *   -> else the permanent backup (tbl_backups row with no end_date, mirrored from
 *      tbl_projects.backup_member) -> else nobody (uncovered day).
 * Uncovered days caused only by approved leaves are conflicts (red row, email, task);
 * days involving a leave request are warnings (orange row, email). Today's responsible
 * person is written to tbl_projects.active_responsible.
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


// Projects columns (ADR 0053: no ID, the name is the reference)
var PROJECT_COLS = { name: 1, status: 2, start: 3, target: 4, lead: 5, backup: 6, active: 7 };
var CONFLICT_STYLE = { bg: "#F8D7DA", color: "#721C24" };
var WARNING_STYLE = { bg: "#FCE8B2", color: "#7A4F01" };
var COVERAGE_HORIZON_DAYS = 366;



function isWeekendDay(d) {
  return d.getDay() === 0 || d.getDay() === 6;
}

/**
 * Leaves by lower-cased Member: approved and requested only (cancelled ignored).
 * @return {Object<string, Array<{startDate: Date, endDate: Date, status: string, leaveType: string}>>}
 */
function readCoverageLeaves(ss) {
  var sheet = getTable(ss, "leaves");
  var byMember = {};
  if (!sheet || sheet.getMaxRows() < 2) return byMember;
  var colMap = (typeof getLeaveColumnMap === "function")
    ? getLeaveColumnMap(sheet)
    : { member: 2, start_date: 3, end_date: 5, leave_type: 7, status: 8 };
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 1000) - 1, Math.max(sheet.getLastColumn(), 6));
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  for (var r = 0; r < vals.length; r++) {
    var name = cellName(vals[r][colMap.member - 1], rich[r] ? rich[r][colMap.member - 1] : null);
    if (!name) continue;
    var status = String(vals[r][colMap.status - 1] || "").trim().toLowerCase();
    if (status !== "approved" && status !== "requested") continue;
    var start = toDate(vals[r][colMap.start_date - 1]);
    if (!start) continue;
    var end = toDate(vals[r][colMap.end_date - 1]) || new Date(start.getTime());
    var key = name.toLowerCase();
    (byMember[key] = byMember[key] || []).push({
      startDate: start, endDate: end, status: status,
      leaveType: String(vals[r][colMap.leave_type - 1] || "Leave").trim()
    });
  }
  return byMember;
}

/**
 * Contracts by Member: [{start, end}] (end null = open-ended).
 * @return {Object<string, Array<{start: (Date|null), end: (Date|null)}>>}
 */
function readCoverageContracts(ss) {
  var sheet = getTable(ss, "tenures");
  var byMember = {};
  if (!sheet || sheet.getMaxRows() < 2) return byMember;
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 500) - 1, 5);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  for (var r = 0; r < vals.length; r++) {
    var name = cellName(vals[r][0], rich[r] ? rich[r][0] : null);
    if (!name) continue;
    (byMember[name.toLowerCase()] = byMember[name.toLowerCase()] || []).push({
      start: toDate(vals[r][3]), end: toDate(vals[r][4])
    });
  }
  return byMember;
}

/**
 * True when the Member has a contract covering the date. A Member with no
 * contract at all is not judged (older data), so this returns true.
 */
function withinContract(contracts, key, date) {
  var list = contracts[key];
  if (!list || !list.length) return true;
  return list.some(function(c) { return (!c.start || c.start <= date) && (!c.end || c.end >= date); });
}

/**
 * Backups rows by project name, lower case (Active / Planned only).
 * @return {Object<string, Array<{row: number, primary: string, backup: string, startDate: (Date|null), endDate: (Date|null)}>>}
 */
function readBackupAssignments(ss) {
  var sheet = getTable(ss, "backups");
  var byProject = {};
  if (!sheet || sheet.getMaxRows() < 2) return byProject;
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 1000) - 1, 8);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  for (var r = 0; r < vals.length; r++) {
    var projKey = String(vals[r][1] || "").trim().toLowerCase();
    if (!projKey) continue;
    var status = String(vals[r][6] || "").trim().toLowerCase();
    if (status !== "active" && status !== "planned") continue;
    (byProject[projKey] = byProject[projKey] || []).push({
      row: r + 2,
      primary: cellName(vals[r][2], rich[r] ? rich[r][2] : null),
      backup: cellName(vals[r][3], rich[r] ? rich[r][3] : null),
      startDate: toDate(vals[r][4]),
      endDate: toDate(vals[r][5])
    });
  }
  return byProject;
}

/**
 * Keeps tbl_backups in step with tbl_projects (ADR 0038):
 * - backup_member is the permanent backup: one tbl_backups row with no end_date.
 *   A changed backup ends the old row yesterday and adds a new one from today.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Date} todayDate
 * @return {{rowsAdded: number, rowsEnded: number}}
 */
function syncProjectBackups(ss, todayDate) {
  var res = { rowsAdded: 0, rowsEnded: 0 };
  var projectsSheet = getTable(ss, "projects");
  var backupsSheet = getTable(ss, "backups");
  if (!projectsSheet || !backupsSheet || projectsSheet.getMaxRows() < 2) return res;

  var pRange = projectsSheet.getRange(2, 1, readRowLimit(projectsSheet, 500) - 1, 6);
  var pVals = pRange.getValues();
  var pRich = (typeof pRange.getRichTextValues === "function") ? pRange.getRichTextValues() : [];
  var assignments = readBackupAssignments(ss);
  var todayIso = formatBackupDateIso(todayDate);
  var yesterdayIso = formatBackupDateIso(addDays(todayDate, -1));

  for (var p = 0; p < pVals.length; p++) {
    var name = String(pVals[p][PROJECT_COLS.name - 1] || "").trim();
    if (!name) continue;
    var lead = cellName(pVals[p][PROJECT_COLS.lead - 1], pRich[p] ? pRich[p][PROJECT_COLS.lead - 1] : null);
    var backup = cellName(pVals[p][PROJECT_COLS.backup - 1], pRich[p] ? pRich[p][PROJECT_COLS.backup - 1] : null);
    if (!backup) continue;

    var permanent = (assignments[name.toLowerCase()] || []).filter(function(a) { return !a.endDate; });
    var current = permanent.filter(function(a) { return a.backup.toLowerCase() === backup.toLowerCase(); })[0];
    if (current) continue;

    for (var k = 0; k < permanent.length; k++) {
      backupsSheet.getRange(permanent[k].row, 6).setValue(yesterdayIso);
      res.rowsEnded++;
    }
    var startIso = permanent.length ? todayIso : (formatBackupDateIso(toDate(pVals[p][PROJECT_COLS.start - 1])) || todayIso);
    var row = (typeof findNextAvailableRow === "function") ? findNextAvailableRow(backupsSheet) : backupsSheet.getLastRow() + 1;
    backupsSheet.getRange(row, 1, 1, 8).setValues([[
      "bkp_" + Math.random().toString(36).substring(2, 9), name, lead, backup, startIso, "", "Active", "Permanent backup"
    ]]);
    if (typeof copyCellChip === "function") {
      try { copyCellChip(projectsSheet, p + 2, PROJECT_COLS.backup, backupsSheet, row, 4, backup, ss.getId ? ss.getId() : ""); } catch (eChip) {}
    }
    res.rowsAdded++;
  }
  return res;
}

/**
 * Scans active projects for days with nobody available, writes today's
 * active_responsible, highlights rows, emails and creates interim-backup tasks.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Object} [options]
 * @param {Date|string} [options.asOfDateParam] - Reference date for today
 * @param {string} [options.managerEmail] - Explicit manager email override
 * @param {boolean} [options.sendNotifications] - Default true
 * @param {boolean} [options.createTasks] - Default true
 * @return {Object} { success, projectsAudited, conflictsFound, conflicts[], warningsFound, warnings[], resolvedCount }
 */
function checkProjectBackupConflicts(targetSpreadsheet, options) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet" };
  var projectsSheet = getTable(ss, "projects");
  var tasksSheet = getTable(ss, "tasks");
  var membersSheet = getTable(ss, "members");
  if (!projectsSheet) return { success: false, message: "Sheet 'Projects' not found." };

  var opt = options || {};
  var sendNotifications = opt.sendNotifications !== false;
  var createTasks = opt.createTasks !== false;
  var todayDate = toDate(opt.asOfDateParam ? opt.asOfDateParam : new Date());
  var todayIso = formatBackupDateIso(todayDate);

  syncProjectBackups(ss, todayDate);

  // 1. Members: emails, chips, today's presence, a manager to notify
  var membersRoster = {};
  var managerEmail = opt.managerEmail || "";
  if (membersSheet && membersSheet.getMaxRows() > 1) {
    var mRange = membersSheet.getRange(2, 1, readRowLimit(membersSheet, 300) - 1, 7);
    var mVals = mRange.getValues();
    var mRich = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : [];
    for (var m = 0; m < mVals.length; m++) {
      var nameM = cellName(mVals[m][0], mRich[m] ? mRich[m][0] : null);
      if (!nameM) continue;
      var emailM = String(mVals[m][1] || "").trim().toLowerCase();
      if (!emailM && typeof extractEmailFromChip === "function") emailM = String(extractEmailFromChip(mVals[m][0], mRich[m] ? mRich[m][0] : null) || "").toLowerCase();
      var roleM = String(mVals[m][3] || "").trim();
      membersRoster[nameM.toLowerCase()] = { row: m + 2, name: nameM, email: emailM, todayPresence: String(mVals[m][6] || "").trim().toLowerCase() };
      if (!managerEmail && emailM && /manager|head|director/i.test(roleM)) managerEmail = emailM;
    }
  }
  if (!managerEmail) { try { managerEmail = Session.getActiveUser().getEmail(); } catch (eSes) {} }
  if (!managerEmail) { try { managerEmail = ss.getOwner().getEmail(); } catch (eOwn) {} }

  var leaves = readCoverageLeaves(ss);
  var assignments = readBackupAssignments(ss);
  var coverageHolidays = (typeof readHolidays === "function") ? readHolidays(ss) : {};
  var contracts = readCoverageContracts(ss);

  // Absence of a Member on a date: null, "approved" or "requested"
  function absence(name, date) {
    if (!name) return "approved";
    var key = name.toLowerCase();
    // Outside every contract (left, or not started yet): not available (ADR 0043)
    if (!withinContract(contracts, key, date)) return "approved";
    var found = null;
    var list = leaves[key] || [];
    for (var i = 0; i < list.length; i++) {
      if (date >= list[i].startDate && date <= list[i].endDate) {
        if (list[i].status === "approved") return "approved";
        found = "requested";
      }
    }
    if (!found && date.getTime() === todayDate.getTime() && membersRoster[key]) {
      var pres = membersRoster[key].todayPresence;
      if (/vacation|sick|leave|absent/.test(pres)) return "approved";
    }
    return found;
  }

  // Who is responsible for a project on a date
  function responsibleOn(project, date) {
    var pending = false;
    var a = absence(project.lead, date);
    if (!a) return { name: project.lead, as: "lead" };
    if (a === "requested") pending = true;
    var rows = assignments[project.key] || [];
    var interim = rows.filter(function(r) { return r.endDate && (!r.startDate || r.startDate <= date) && r.endDate >= date; });
    var permanent = rows.filter(function(r) { return !r.endDate && (!r.startDate || r.startDate <= date); });
    var candidates = interim.map(function(r) { return { name: r.backup, as: "interim" }; })
      .concat(permanent.map(function(r) { return { name: r.backup, as: "backup" }; }));
    if (!permanent.length && project.backup) candidates.push({ name: project.backup, as: "backup" });
    for (var c = 0; c < candidates.length; c++) {
      var ab = absence(candidates[c].name, date);
      if (!ab) return { name: candidates[c].name, as: candidates[c].as };
      if (ab === "requested") pending = true;
    }
    return { name: null, as: null, pending: pending };
  }

  // Last day worth checking for a project: the latest leave end among the people involved
  function horizonFor(project) {
    var people = [project.lead, project.backup].concat((assignments[project.key] || []).map(function(r) { return r.backup; }));
    var last = todayDate;
    people.forEach(function(n) {
      (leaves[String(n || "").toLowerCase()] || []).forEach(function(l) { if (l.endDate > last) last = l.endDate; });
    });
    var cap = addDays(todayDate, COVERAGE_HORIZON_DAYS);
    // Someone whose contracts all end within the horizon is missing from then on
    people.forEach(function(n) {
      var list = contracts[String(n || "").toLowerCase()];
      if (list && list.length && !list.some(function(c) { return !c.end; })) last = cap;
    });
    return last > cap ? cap : last;
  }

  // Existing open interim-backup tasks, so each project gets at most one
  var openBackupTasks = {};
  if (tasksSheet && tasksSheet.getMaxRows() > 1) {
    var tc = {};
    SCHEMA_SPEC.tasks.columns.forEach(function(k, i) { tc[k] = i; });
    var tVals = tasksSheet.getRange(2, 1, readRowLimit(tasksSheet, 1000) - 1, SCHEMA_SPEC.tasks.columns.length).getValues();
    for (var t = 0; t < tVals.length; t++) {
      var tProj = String(tVals[t][tc.project] || "").trim().toLowerCase();
      if (tProj && String(tVals[t][tc.title] || "").toLowerCase().indexOf("organize interim backup") !== -1 && String(tVals[t][tc.status] || "").toLowerCase() !== "done") {
        openBackupTasks[tProj] = true;
      }
    }
  }

  var pRange = projectsSheet.getRange(2, 1, readRowLimit(projectsSheet, 500) - 1, PROJECT_COLS.active);
  var pVals = pRange.getValues();
  var pRich = (typeof pRange.getRichTextValues === "function") ? pRange.getRichTextValues() : [];
  var conflicts = [], warnings = [], resolvedCount = 0, audited = 0;
  var activeColumn = [];

  for (var p = 0; p < pVals.length; p++) {
    var rowNum = p + 2;
    var pc = PROJECT_COLS;
    var project = {
      name: String(pVals[p][pc.name - 1] || "").trim(),
      status: String(pVals[p][pc.status - 1] || "").trim().toLowerCase(),
      lead: cellName(pVals[p][pc.lead - 1], pRich[p] ? pRich[p][pc.lead - 1] : null),
      backup: cellName(pVals[p][pc.backup - 1], pRich[p] ? pRich[p][pc.backup - 1] : null)
    };
    project.key = project.name.toLowerCase();
    var rowRange = projectsSheet.getRange(rowNum, 1, 1, pc.active);
    if (!project.name) { activeColumn.push([pVals[p][pc.active - 1] || ""]); continue; }
    if (project.status === "completed" || !project.lead) {
      rowRange.setBackground("#FFFFFF").setFontColor("#000000").setFontWeight("normal");
      activeColumn.push([""]);
      continue;
    }
    audited++;

    // Today's responsible person
    var now = responsibleOn(project, todayDate);
    activeColumn.push([now.name
      ? now.name + (now.as === "lead" ? "" : " (" + now.as + ")")
      : "⚠ Nobody available" + (now.pending ? " (pending request)" : "")]);

    // Uncovered working days from today on, grouped into periods (weekends do not break a period)
    var periods = [];
    var open = null;
    var horizon = horizonFor(project);
    for (var d = new Date(todayDate.getTime()); d <= horizon; d = addDays(d, 1)) {
      if (isWeekendDay(d) || coverageHolidays[formatBackupDateIso(d)]) continue;
      var who = responsibleOn(project, d);
      if (who.name) { open = null; continue; }
      var kind = who.pending ? "warning" : "conflict";
      if (open && open.kind === kind) { open.end = new Date(d.getTime()); continue; }
      open = { kind: kind, start: new Date(d.getTime()), end: new Date(d.getTime()) };
      periods.push(open);
    }
    var conflictPeriods = periods.filter(function(x) { return x.kind === "conflict"; });
    var warningPeriods = periods.filter(function(x) { return x.kind === "warning"; });

    if (!periods.length) {
      rowRange.setBackground("#FFFFFF").setFontColor("#000000").setFontWeight("normal");
      resolvedCount++;
      continue;
    }

    var style = conflictPeriods.length ? CONFLICT_STYLE : WARNING_STYLE;
    var previousBg = String(rowRange.getBackground() || "").toUpperCase();
    var alreadyAlerted = previousBg === style.bg;
    rowRange.setBackground(style.bg).setFontColor(style.color).setFontWeight("bold");

    var first = (conflictPeriods[0] || warningPeriods[0]);
    var startStr = formatBackupDateIso(first.start), endStr = formatBackupDateIso(first.end);
    var isConflict = !!conflictPeriods.length;
    var reason = (isConflict ? "Nobody available" : "Nobody available if pending leave requests are approved") +
      " (" + startStr + (endStr !== startStr ? " to " + endStr : "") + ")";
    if (!withinContract(contracts, project.lead.toLowerCase(), first.start)) {
      reason += ": " + project.lead + "'s contract ends, so the project needs a new lead";
    }

    var leadEmail = membersRoster[project.lead.toLowerCase()] ? membersRoster[project.lead.toLowerCase()].email : "";
    var recipients = [];
    if (leadEmail) recipients.push(leadEmail);
    if (managerEmail && recipients.indexOf(managerEmail.toLowerCase()) === -1) recipients.push(managerEmail);

    var emailSent = false;
    if (sendNotifications && !alreadyAlerted && recipients.length && typeof MailApp !== "undefined" && MailApp.sendEmail) {
      try {
        MailApp.sendEmail({
          to: recipients.join(","),
          subject: (isConflict ? "[ACTION REQUIRED] No backup available on project: " : "[WARNING] Leave request leaves project without backup: ") + project.name,
          body: (isConflict ? "Nobody can cover this project on some working days.\n\n" : "If the pending leave request(s) are approved, nobody can cover this project on some working days.\n\n") +
            "Project: " + project.name + "\n" +
            "Project Lead: " + project.lead + "\n" +
            "Permanent backup: " + (project.backup || "none") + "\n" +
            "Uncovered: " + startStr + " to " + endStr + "\n\n" +
            "Add an interim backup for these dates with Project Management → Interim Backups, or adjust the leaves."
        });
        emailSent = true;
      } catch (eMail) { Logger.log("Backup coverage email note: " + eMail); }
    }

    var taskCreated = false;
    if (isConflict && createTasks && tasksSheet && !openBackupTasks[project.key]) {
      try {
        var assignee = absence(project.lead, todayDate) ? (project.backup || project.lead) : project.lead;
        var assigneeRoster = membersRoster[String(assignee).toLowerCase()];
        var newTask = (typeof createTaskRow === "function" ? createTaskRow : require("./RecurringTasks").createTaskRow)(ss, {
          project: project.name,
          member: assigneeRoster ? { name: assigneeRoster.name, row: assigneeRoster.row } : (assignee ? { name: assignee } : null),
          title: "Organize interim backup for project " + project.name,
          description: "Nobody (lead or backup) is available from " + startStr + " to " + endStr + ". Add an interim backup for these dates with Project Management → Interim Backups.",
          due: startStr, status: "Todo", origin: "trigger:coverageGap"
        });
        openBackupTasks[project.key] = true;
        taskCreated = true;
      } catch (eTask) { Logger.log("Backup task note: " + eTask); }
    }

    var entry = {
      projectName: project.name, row: rowNum,
      leadMember: project.lead, backupMember: project.backup,
      conflictStart: startStr, conflictEnd: endStr, reason: reason,
      pending: !isConflict, emailSent: emailSent, recipients: recipients, taskCreated: taskCreated,
      periods: periods.map(function(x) { return { kind: x.kind, start: formatBackupDateIso(x.start), end: formatBackupDateIso(x.end) }; })
    };
    (isConflict ? conflicts : warnings).push(entry);
  }

  if (activeColumn.length) {
    projectsSheet.getRange(2, PROJECT_COLS.active, activeColumn.length, 1).setValues(activeColumn);
  }

  return {
    success: true,
    asOf: todayIso,
    projectsAudited: audited,
    conflictsFound: conflicts.length,
    conflicts: conflicts,
    warningsFound: warnings.length,
    warnings: warnings,
    resolvedCount: resolvedCount
  };
}

/**
 * Asks for a backup when a project row has a name but no backup_member (onEdit on tbl_projects).
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {boolean} true when the toast was shown
 */
function promptProjectBackup(e) {
  if (!e || !e.range || e.range.getRow() < 2) return false;
  var sheet = e.range.getSheet();
  var vals = sheet.getRange(e.range.getRow(), 1, 1, PROJECT_COLS.active).getValues()[0];
  var name = String(vals[PROJECT_COLS.name - 1] || "").trim();
  if (!name || String(vals[PROJECT_COLS.backup - 1] || "").trim()) return false;
  var ss = e.source || (sheet.getParent ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet());
  try {
    ss.toast("Project \"" + name + "\" has no backup yet. Fill in backup_member so absences stay covered.", "Backup needed", 8);
  } catch (eToast) {}
  return true;
}

/**
 * Assigns an interim backup in tbl_backups and automatically re-evaluates project conflict formatting.
 *
 * @param {Object} payload
 * @param {string} payload.project The project name
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
  if (!payload || !payload.project || !payload.primaryMember || !payload.backupMember || !payload.startDate) {
    return { success: false, message: "Missing required fields (project, primaryMember, backupMember, startDate)." };
  }

  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet" };

  var backupsSheet = getTable(ss, "backups");
  var membersSheet = getTable(ss, "members");

  if (!backupsSheet) {
    return { success: false, message: "Sheet 'Backups' not found. Run Initialize / Upgrade first." };
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
    payload.project,
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
    var mMax = readRowLimit(membersSheet, 300);
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
    // Formatting and "In charge today" only: the nightly check sends the emails
    conflictReport = checkProjectBackupConflicts(ss, { sendNotifications: false, createTasks: false });
  } catch (eConf) {
    Logger.log("Conflict re-check note: " + eConf);
  }

  return {
    success: true,
    backupId: backupId,
    project: payload.project,
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
    readCoverageContracts: readCoverageContracts,
    withinContract: withinContract,
    formatBackupDateIso: formatBackupDateIso,
    readCoverageLeaves: readCoverageLeaves,
    checkProjectBackupConflicts: checkProjectBackupConflicts,
    syncProjectBackups: syncProjectBackups,
    promptProjectBackup: promptProjectBackup,
    assignProjectBackup: assignProjectBackup
  };
}
