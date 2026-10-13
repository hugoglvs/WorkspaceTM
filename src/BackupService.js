/**
 * Interim Backups sidebar (ADR 0041): add a dated backup to a project without
 * opening the hidden Backups tab. The sidebar shows each open project's lead,
 * permanent backup, coverage gaps (click one to fill the dates) and interim
 * backups still running or planned, which can be cancelled.
 */

function bsIso(d) {
  return (typeof formatBackupDateIso === "function") ? formatBackupDateIso(d) : String(d || "");
}

/**
 * Everything the sidebar needs. Refreshes coverage (no emails, no tasks) so the
 * gaps are current.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Date|string} [asOf]
 */
function getBackupSidebarData(targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var today = toDate(asOf || new Date());
  var todayIso = bsIso(today);

  var gapsByProject = {};
  try {
    var rep = checkProjectBackupConflicts(ss, { sendNotifications: false, createTasks: false, asOfDateParam: today });
    (rep.conflicts || []).concat(rep.warnings || []).forEach(function(c) {
      gapsByProject[c.projectId] = (gapsByProject[c.projectId] || []).concat(c.periods || []);
    });
  } catch (eCov) { Logger.log("Backup sidebar coverage note: " + eCov); }

  var interims = {};
  var bSheet = getTable(ss, "backups");
  if (bSheet && bSheet.getMaxRows() > 1) {
    var bRange = bSheet.getRange(2, 1, readRowLimit(bSheet, 1000) - 1, 8);
    var bVals = bRange.getValues();
    var bRich = (typeof bRange.getRichTextValues === "function") ? bRange.getRichTextValues() : [];
    bVals.forEach(function(v, i) {
      var projId = String(v[1] || "").trim();
      var status = String(v[6] || "").trim();
      var end = toDate(v[5]);
      if (!projId || !end || bsIso(end) < todayIso) return;
      if (["active", "planned"].indexOf(status.toLowerCase()) === -1) return;
      (interims[projId] = interims[projId] || []).push({
        id: String(v[0] || ""), primary: cellName(v[2], bRich[i] ? bRich[i][2] : null),
        backup: cellName(v[3], bRich[i] ? bRich[i][3] : null),
        start: bsIso(toDate(v[4])), end: bsIso(end), status: status, notes: String(v[7] || "")
      });
    });
  }

  var projects = [];
  var pSheet = getTable(ss, "projects");
  if (pSheet && pSheet.getMaxRows() > 1) {
    var pRange = pSheet.getRange(2, 1, readRowLimit(pSheet, 500) - 1, 7);
    var pVals = pRange.getValues();
    var pRich = (typeof pRange.getRichTextValues === "function") ? pRange.getRichTextValues() : [];
    pVals.forEach(function(v, i) {
      var id = String(v[0] || "").trim();
      var name = String(v[1] || "").trim();
      if (!id || !name || String(v[2] || "").trim().toLowerCase() === "completed") return;
      projects.push({
        id: id, name: name, status: String(v[2] || ""),
        lead: cellName(v[5], pRich[i] ? pRich[i][5] : null),
        backup: cellName(v[6], pRich[i] ? pRich[i][6] : null),
        gaps: gapsByProject[id] || [], interims: interims[id] || []
      });
    });
  }
  projects.sort(function(a, b) { return (b.gaps.length ? 1 : 0) - (a.gaps.length ? 1 : 0) || a.name.localeCompare(b.name); });

  var members = [];
  var mSheet = getTable(ss, "members");
  if (mSheet && mSheet.getMaxRows() > 1) {
    var mRange = mSheet.getRange(2, 1, readRowLimit(mSheet, 300) - 1, 6);
    var mVals = mRange.getValues();
    var mRich = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : [];
    mVals.forEach(function(v, i) {
      var name = cellName(v[0], mRich[i] ? mRich[i][0] : null);
      var status = String(v[5] || "").trim().toLowerCase();
      if (name && status !== "departed") members.push(name);
    });
  }
  members.sort(function(a, b) { return a.localeCompare(b); });

  var leaves = {};
  var lv = readCoverageLeaves(ss);
  Object.keys(lv).forEach(function(k) {
    leaves[k] = lv[k].map(function(l) { return { start: bsIso(l.startDate), end: bsIso(l.endDate), status: l.status }; });
  });

  return { today: todayIso, projects: projects, members: members, leaves: leaves };
}

/**
 * Adds an interim backup (a Backups row with an end date).
 *
 * @param {{projectId: string, backupMember: string, startDate: string, endDate: string, notes: (string|undefined)}} payload
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Date|string} [asOf]
 */
function saveInterimBackup(payload, targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var p = payload || {};
  var start = toDate(p.startDate);
  var end = toDate(p.endDate);
  if (!p.projectId || !p.backupMember) return { success: false, message: "Choose a project and a backup." };
  if (!start || !end) return { success: false, message: "An interim backup needs a start and an end date." };
  if (end < start) return { success: false, message: "The end date is before the start date." };

  var project = getBackupSidebarProject(ss, p.projectId);
  if (!project) return { success: false, message: "Project not found." };
  if (String(p.backupMember).toLowerCase() === String(project.lead).toLowerCase()) {
    return { success: false, message: p.backupMember + " is the project lead. Pick someone else." };
  }

  var lock = (typeof acquireSyncLock === "function") ? acquireSyncLock(10000) : null;
  if (typeof acquireSyncLock === "function" && !lock) return { success: false, message: "A sync is running. Try again in a minute." };
  try {
    var today = toDate(asOf || new Date());
    var res = assignProjectBackup({
      projectId: p.projectId, primaryMember: project.lead || project.backup, backupMember: p.backupMember,
      startDate: bsIso(start), endDate: bsIso(end), status: start > today ? "Planned" : "Active", notes: p.notes || ""
    }, ss);
    if (res.success) res.message = p.backupMember + " covers " + project.name + " from " + bsIso(start) + " to " + bsIso(end) + ".";
    return res;
  } finally {
    if (lock) lock.releaseLock();
  }
}

function getBackupSidebarProject(ss, projectId) {
  var sheet = getTable(ss, "projects");
  if (!sheet || sheet.getMaxRows() < 2) return null;
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 500) - 1, 7);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || "").trim() === String(projectId)) {
      return { id: projectId, name: String(vals[i][1] || "").trim(),
        lead: cellName(vals[i][5], rich[i] ? rich[i][5] : null),
        backup: cellName(vals[i][6], rich[i] ? rich[i][6] : null) };
    }
  }
  return null;
}

/** Cancels an interim backup (status Cancelled; the row stays as history). */
function cancelInterimBackup(backupId, targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getTable(ss, "backups");
  if (!sheet || !backupId) return { success: false, message: "Backup not found." };
  var ids = sheet.getRange(2, 1, readRowLimit(sheet, 1000) - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(backupId)) {
      sheet.getRange(i + 2, 7).setValue("Cancelled");
      try { checkProjectBackupConflicts(ss, { sendNotifications: false, createTasks: false }); } catch (e) {}
      return { success: true, message: "Interim backup cancelled." };
    }
  }
  return { success: false, message: "Backup not found." };
}

/** Menu: Interim Backups. */
function showBackupSidebar() {
  var html = HtmlService.createTemplateFromFile("BackupSidebar").evaluate().setTitle("Interim Backups").setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getBackupSidebarData: getBackupSidebarData,
    saveInterimBackup: saveInterimBackup,
    cancelInterimBackup: cancelInterimBackup,
    showBackupSidebar: showBackupSidebar
  };
}
