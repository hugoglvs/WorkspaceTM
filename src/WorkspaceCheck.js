/**
 * Check Workspace and the Monday presence reminder (ADR 0041).
 *
 * Check Workspace (menu) lists what needs attention: members without a contract,
 * projects without a backup, coverage gaps, members not logged today, members who
 * have not declared every day of this week, and duplicate Presence Log rows
 * (which it can remove).
 *
 * The Monday reminder emails each active member who has a working day this week
 * with no Weekly Presence entry, no routine and no leave.
 */

var WC_WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];

function wcIso(d) {
  return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2);
}

/** Date from a cell (Date or "YYYY-MM-DD…"), at local midnight; null if empty. */
function wcDate(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var m = String(v || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  if (!v) return null;
  var d = new Date(String(v));
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function wcName(val, rich) {
  return (typeof extractCellValue === "function") ? extractCellValue(val, rich) : String(val || "").trim();
}

function wcRows(sheet, numCols, cap) {
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var range = sheet.getRange(2, 1, readRowLimit(sheet, cap) - 1, numCols);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  return vals.map(function(v, i) { return { row: i + 2, vals: v, rich: rich[i] || [] }; });
}

/** Members with their status, emails and today's presence. */
function wcMembers(ss) {
  return wcRows(getTable(ss, "members"), 7, 300).map(function(r) {
    var name = wcName(r.vals[0], r.rich[0]);
    if (!name) return null;
    var email = String(r.vals[1] || "").trim();
    if (!email && typeof extractEmailFromChip === "function") email = String(extractEmailFromChip(r.vals[0], r.rich[0]) || "").trim();
    return { row: r.row, name: name, key: name.toLowerCase(), email: email, status: String(r.vals[5] || "").trim(), today: String(r.vals[6] || "").trim() };
  }).filter(Boolean);
}

/** Monday to Friday of the week holding asOf (a weekend looks at the next week). */
function wcWorkWeek(asOf) {
  var d = wcDate(asOf || new Date());
  var dow = d.getDay();
  var monday = new Date(d.getFullYear(), d.getMonth(), d.getDate() + (dow === 0 ? 1 : (dow === 6 ? 2 : 1 - dow)));
  var days = [];
  for (var i = 0; i < 5; i++) days.push(new Date(monday.getFullYear(), monday.getMonth(), monday.getDate() + i));
  return days;
}

/**
 * Active members with working days this week that nothing declares: no Weekly
 * Presence entry (Presence Log row), no routine for that weekday, no leave.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Date|string} [asOf]
 * @return {Array<{name: string, email: string, days: string[]}>}
 */
function findUndeclaredDays(ss, asOf) {
  var holidays = (typeof readHolidays === "function") ? readHolidays(ss) : {};
  var week = wcWorkWeek(asOf).filter(function(d) { return !holidays[wcIso(d)]; });
  var logged = {};
  wcRows(getTable(ss, "presenceLog"), 5, 2000).forEach(function(r) {
    var d = wcDate(r.vals[1]);
    var who = wcName(r.vals[2], r.rich[2]).toLowerCase();
    if (d && who && (r.vals[3] || r.vals[4])) logged[wcIso(d) + "|" + who] = true;
  });
  var routines = {};
  wcRows(getTable(ss, "baselines"), 7, 500).forEach(function(r) {
    var who = wcName(r.vals[1], r.rich[1]).toLowerCase();
    if (!who) return;
    (routines[who] = routines[who] || []).push({
      start: wcDate(r.vals[2]), end: wcDate(r.vals[3]), day: String(r.vals[4] || "").trim().toLowerCase()
    });
  });
  var leaves = (typeof readCoverageLeaves === "function") ? readCoverageLeaves(ss) : {};

  var res = [];
  wcMembers(ss).forEach(function(m) {
    if (m.status.toLowerCase() !== "active") return;
    var missing = week.filter(function(day) {
      if (logged[wcIso(day) + "|" + m.key]) return false;
      var dayName = WC_WEEKDAYS[day.getDay()].toLowerCase();
      var hasRoutine = (routines[m.key] || []).some(function(rt) {
        if (rt.day && rt.day !== "all" && rt.day !== dayName) return false;
        return (!rt.start || rt.start <= day) && (!rt.end || rt.end >= day);
      });
      if (hasRoutine) return false;
      return !(leaves[m.key] || []).some(function(l) { return l.startDate <= day && l.endDate >= day; });
    });
    if (missing.length) res.push({ name: m.name, email: m.email, days: missing.map(wcIso) });
  });
  return res;
}

/**
 * Presence Log rows that repeat a date + member. The last row wins (it is the
 * one Weekly Presence shows), so the earlier ones are the duplicates.
 *
 * @return {{count: number, rows: number[], examples: string[]}}
 */
function findPresenceDuplicates(ss) {
  var seen = {};
  var dupRows = [];
  var examples = [];
  var rows = wcRows(getTable(ss, "presenceLog"), 3, 2000);
  for (var i = rows.length - 1; i >= 0; i--) {
    var d = wcDate(rows[i].vals[1]);
    var who = wcName(rows[i].vals[2], rows[i].rich[2]);
    if (!d || !who) continue;
    var key = wcIso(d) + "|" + who.toLowerCase();
    if (seen[key]) {
      dupRows.push(rows[i].row);
      if (examples.length < 5) examples.push(who + " on " + wcIso(d));
    } else {
      seen[key] = true;
    }
  }
  return { count: dupRows.length, rows: dupRows, examples: examples };
}

/** Deletes duplicate Presence Log rows (bottom-up, under the sync lock). */
function removePresenceDuplicates(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var lock = (typeof acquireSyncLock === "function") ? acquireSyncLock(10000) : null;
  if (typeof acquireSyncLock === "function" && !lock) {
    return { success: false, message: "A sync is running. Try again in a minute." };
  }
  try {
    var dup = findPresenceDuplicates(ss);
    var sheet = getTable(ss, "presenceLog");
    dup.rows.sort(function(a, b) { return b - a; }).forEach(function(r) { sheet.deleteRow(r); });
    if (dup.count && typeof refreshWeeklyMatrix === "function") {
      try { refreshWeeklyMatrix(ss); } catch (eM) {}
    }
    return { success: true, removed: dup.count, message: dup.count + " duplicate row(s) removed." };
  } finally {
    if (lock) lock.releaseLock();
  }
}

/**
 * Everything Check Workspace reports.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Date|string} [asOf]
 * @return {{asOf: string, issues: number, sections: Array<{key: string, title: string, ok: string, items: string[], action: (string|undefined)}>}}
 */
function runWorkspaceCheck(targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var today = wcDate(asOf || new Date());
  var members = wcMembers(ss);
  var sections = [];

  sections.push({
    key: "contracts", title: "Members without a contract", ok: "Every member has a contract.",
    action: "members",
    items: members.filter(function(m) { return m.status.toLowerCase() === "no contract"; }).map(function(m) { return m.name; })
  });

  var projects = wcRows(getTable(ss, "projects"), 7, 500).filter(function(r) {
    return String(r.vals[1] || "").trim() && String(r.vals[2] || "").trim().toLowerCase() !== "completed";
  });
  sections.push({
    key: "backups", title: "Projects without a backup", ok: "Every open project has a backup.",
    items: projects.filter(function(r) { return !wcName(r.vals[6], r.rich[6]); }).map(function(r) { return String(r.vals[1]).trim(); })
  });

  var coverage = [];
  if (typeof checkProjectBackupConflicts === "function") {
    try {
      var rep = checkProjectBackupConflicts(ss, { sendNotifications: false, createTasks: false, asOfDateParam: today });
      (rep.conflicts || []).concat(rep.warnings || []).forEach(function(c) {
        (c.periods || []).forEach(function(p) {
          coverage.push(c.projectName + ": nobody " + (p.start === p.end ? "on " + p.start : "from " + p.start + " to " + p.end) +
            (p.kind === "warning" ? " (if a leave request is approved)" : ""));
        });
      });
    } catch (eCov) { Logger.log("Check Workspace coverage note: " + eCov); }
  }
  sections.push({
    key: "leavers", title: "People leaving: projects and tasks to hand over", ok: "No project or open task depends on someone leaving.",
    action: "members", items: findLeaverHandovers(ss, today, members)
  });

  sections.push({ key: "coverage", title: "Coverage gaps", ok: "Every project is covered every working day.", action: "backups", items: coverage });

  var holidayToday = ((typeof readHolidays === "function") ? readHolidays(ss) : {})[wcIso(today)];
  var weekend = today.getDay() === 0 || today.getDay() === 6 || !!holidayToday;
  sections.push({
    key: "notLogged", title: "Not logged today", ok: holidayToday ? holidayToday + ": nothing to log." : (weekend ? "Weekend: nothing to log." : "Everyone is logged today."),
    items: weekend ? [] : members.filter(function(m) { return m.today === "Not Logged"; }).map(function(m) { return m.name; })
  });

  sections.push({
    key: "undeclared", title: "Days not declared this week", ok: "Everyone has declared this week.", action: "presence",
    items: findUndeclaredDays(ss, today).map(function(u) { return u.name + ": " + u.days.map(wcShortDay).join(", "); })
  });

  var dup = findPresenceDuplicates(ss);
  sections.push({
    key: "duplicates", title: "Duplicate Weekly Presence entries", ok: "No duplicates.", action: dup.count ? "dedupe" : undefined,
    items: dup.count ? [dup.count + " duplicate row(s) in Presence Log, e.g. " + dup.examples.join("; ")] : []
  });

  var issues = sections.reduce(function(n, s) { return n + s.items.length; }, 0);
  return { asOf: wcIso(today), issues: issues, sections: sections };
}

function wcShortDay(iso) {
  var d = wcDate(iso);
  return WC_WEEKDAYS[d.getDay()].substring(0, 3) + " " + d.getDate();
}

/**
 * What still depends on people who have left or leave within 30 days: open
 * projects they lead or back up, and open tasks assigned to them.
 *
 * @return {string[]}
 */
function findLeaverHandovers(ss, today, members) {
  var contracts = (typeof readCoverageContracts === "function") ? readCoverageContracts(ss) : {};
  var horizon = new Date(today.getFullYear(), today.getMonth(), today.getDate() + 30);
  var leaving = {}; // key -> "left on …" / "leaves on …"
  Object.keys(contracts).forEach(function(key) {
    var list = contracts[key];
    if (list.some(function(c) { return !c.end; })) return;
    var last = list.reduce(function(max, c) { return (!max || c.end > max) ? c.end : max; }, null);
    if (last && last < horizon) leaving[key] = (last < today ? "left on " : "leaves on ") + wcIso(last);
  });
  (members || wcMembers(ss)).forEach(function(m) {
    if (!leaving[m.key] && m.status.toLowerCase() === "departed") leaving[m.key] = "has left";
  });
  if (!Object.keys(leaving).length) return [];

  var items = [];
  wcRows(getTable(ss, "projects"), 7, 500).forEach(function(r) {
    var name = String(r.vals[1] || "").trim();
    if (!name || String(r.vals[2] || "").trim().toLowerCase() === "completed") return;
    [["lead", 5], ["backup", 6]].forEach(function(role) {
      var who = wcName(r.vals[role[1]], r.rich[role[1]]);
      if (who && leaving[who.toLowerCase()]) items.push(name + ": " + role[0] + " " + who + " " + leaving[who.toLowerCase()]);
    });
  });
  wcRows(getTable(ss, "tasks"), 7, 1000).forEach(function(r) {
    var title = String(r.vals[3] || "").trim();
    var status = String(r.vals[6] || "").trim().toLowerCase();
    var who = wcName(r.vals[2], r.rich[2]);
    if (!title || !who || status === "done" || status === "completed" || status === "cancelled") return;
    if (leaving[who.toLowerCase()]) items.push("Task \"" + title + "\": " + who + " " + leaving[who.toLowerCase()]);
  });
  return items;
}

/** Menu: Check Workspace. Shows the report in a dialog. */
function showWorkspaceCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = runWorkspaceCheck(ss);
  var template = HtmlService.createTemplateFromFile("WorkspaceCheckDialog");
  template.reportJson = JSON.stringify(report).replace(/</g, "\\u003c");
  var html = template.evaluate().setWidth(520).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, "Check Workspace");
  return report;
}

/**
 * Emails every active member who has undeclared working days this week.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Date|string} [asOf]
 * @return {{sent: string[], noEmail: string[]}}
 */
function sendPresenceReminders(targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var res = { sent: [], noEmail: [] };
  var weekly = getTable(ss, "weeklyMatrix");
  var link = ss.getUrl() + (weekly ? "#gid=" + weekly.getSheetId() : "");
  findUndeclaredDays(ss, asOf).forEach(function(u) {
    if (!u.email) { res.noEmail.push(u.name); return; }
    var days = u.days.map(wcShortDay);
    var first = u.name.split(/[\s@.]/)[0];
    MailApp.sendEmail({
      to: u.email,
      subject: "Please declare your presence for this week",
      body: "Hi " + first + ",\n\nThese days of this week have no presence yet: " + days.join(", ") +
        ".\n\nFill them in Weekly Presence, or set a routine with Project Management → Declare Presence & Leaves:\n" + link + "\n",
      htmlBody: "<p>Hi " + wcHtml(first) + ",</p><p>These days of this week have no presence yet: <b>" + wcHtml(days.join(", ")) +
        "</b>.</p><p>Fill them in <a href=\"" + link + "\">Weekly Presence</a>, or set a routine with " +
        "<i>Project Management → Declare Presence &amp; Leaves</i>.</p>"
    });
    res.sent.push(u.name);
  });
  Logger.log("[mondayPresenceReminder] sent: " + res.sent.join(", ") + (res.noEmail.length ? " | no email: " + res.noEmail.join(", ") : ""));
  return res;
}

function wcHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/** Time trigger entry point (Mondays, 8 am). */
function mondayPresenceReminder() {
  return sendPresenceReminders(SpreadsheetApp.getActiveSpreadsheet());
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    findUndeclaredDays: findUndeclaredDays,
    findPresenceDuplicates: findPresenceDuplicates,
    findLeaverHandovers: findLeaverHandovers,
    removePresenceDuplicates: removePresenceDuplicates,
    runWorkspaceCheck: runWorkspaceCheck,
    showWorkspaceCheck: showWorkspaceCheck,
    sendPresenceReminders: sendPresenceReminders,
    mondayPresenceReminder: mondayPresenceReminder
  };
}
