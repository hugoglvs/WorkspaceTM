/**
 * Recurring tasks (ADR 0045).
 *
 * A rule in the hidden "Recurring Tasks" tab (ID | Type | Project ID | Who |
 * Applies to | Start | End | Active) turns into dated rows in Tasks, the way a
 * Routine turns into Weekly Presence rows:
 * - Who = "Each person": one task per person in "Applies to" each period
 *   (e.g. Time & expenses every month for every Consultant);
 * - Who = "Rotation": one person per period, the one whose last turn is the
 *   oldest, skipping anyone on approved leave on the due date (a skipped person
 *   therefore goes first next time).
 * "Applies to" lists members and roles ("Role: Consultant; Alice"); blank means
 * every member. Weekly rules are planned 8 weeks ahead, monthly ones 3 months.
 *
 * Generated rows carry "<rule id>|<period>" in Tasks.Recurring; a row changed by
 * hand in the Calendar gets "|manual" and is never moved automatically.
 */

var RECURRING_WEEKS_AHEAD = 8;
var RECURRING_MONTHS_AHEAD = 3;
var OPEN_TASK_STATUSES = ["todo", "in progress", "blocked", ""];

function taskCol(name) {
  return SCHEMA_SPEC.tasks.columns.indexOf(name) + 1;
}


function rtToday(asOf) {
  return toDate(asOf || new Date());
}

/** Members with role, status and emails, in tab order. */
function rtMembers(ss) {
  var sheet = getTable(ss, "members");
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 300) - 1, 8);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = rich[i] || [];
    var name = cellName(vals[i][0], r[0]);
    if (!name) continue;
    var google = (typeof extractEmailFromChip === "function") ? String(extractEmailFromChip(vals[i][0], r[0]) || "").trim() : "";
    var work = String(vals[i][1] || "").trim();
    out.push({
      row: i + 2, name: name, key: name.toLowerCase(), role: String(vals[i][3] || "").trim(),
      status: String(vals[i][5] || "").trim().toLowerCase(),
      email: work || google, googleEmail: google || work
    });
  }
  return out;
}

/**
 * Members named by an "Applies to" text, in order: names as listed, roles in
 * Members order. Blank = every member. Departed members are left out.
 */
function resolveAppliesTo(appliesTo, members) {
  var current = members.filter(function(m) { return m.status !== "departed"; });
  var text = String(appliesTo || "").trim();
  if (!text) return current.slice();
  var out = [];
  var seen = {};
  var add = function(m) { if (!seen[m.key]) { seen[m.key] = true; out.push(m); } };
  text.split(/[,;\n]/).forEach(function(tok) {
    var t = tok.trim();
    if (!t) return;
    var role = t.match(/^role\s*:\s*(.+)$/i);
    if (role) {
      var r = role[1].trim().toLowerCase();
      current.filter(function(m) { return m.role.toLowerCase() === r; }).forEach(add);
    } else {
      current.filter(function(m) { return m.key === t.toLowerCase() || (m.email && m.email.toLowerCase() === t.toLowerCase()); }).forEach(add);
    }
  });
  return out;
}

/** "approved", "requested" or null for a member on a date. */
function leaveStatusOn(leaves, key, date) {
  var found = null;
  (leaves[key] || []).forEach(function(l) {
    if (date >= l.startDate && date <= l.endDate) {
      if (l.status === "approved") found = "approved";
      else if (!found) found = "requested";
    }
  });
  return found;
}

function rtContext(ss) {
  return {
    members: rtMembers(ss),
    leaves: (typeof readCoverageLeaves === "function") ? readCoverageLeaves(ss) : {},
    contracts: (typeof readCoverageContracts === "function") ? readCoverageContracts(ss) : {},
    holidays: (typeof readHolidays === "function") ? readHolidays(ss) : {},
    types: readTaskTypes(ss)
  };
}

function rtWorksOn(ctx, member, date) {
  return (typeof withinContract === "function") ? withinContract(ctx.contracts, member.key, date) : true;
}

// ---------------------------------------------------------------------------
// Tasks rows
// ---------------------------------------------------------------------------

function parseRecurrence(text) {
  var parts = String(text || "").split("|");
  return { ruleId: parts[0] || "", periodKey: parts[1] || "", manual: parts[2] === "manual" };
}

/** Every Tasks row, with the fields the recurring engine and the Calendar use. */
function readTaskRows(ss) {
  var sheet = getTable(ss, "tasks");
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var n = SCHEMA_SPEC.tasks.columns.length;
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 500) - 1, n);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (!v[0] && !v[3]) continue;
    var r = rich[i] || [];
    var member = cellName(v[2], r[2]);
    var rec = parseRecurrence(v[11]);
    out.push({
      row: i + 2, id: String(v[0] || ""), projectId: String(v[1] || "").trim(), member: member, memberKey: member.toLowerCase(),
      title: String(v[3] || ""), due: toDate(v[5]), status: String(v[6] || "").trim(), googleTaskId: String(v[7] || ""),
      type: String(v[8] || "").trim(), doneOn: v[9], doneBy: String(v[10] || ""), recurrence: String(v[11] || ""),
      ruleId: rec.ruleId, periodKey: rec.periodKey, manual: rec.manual
    });
  }
  return out;
}

function isOpenStatus(status) {
  return OPEN_TASK_STATUSES.indexOf(String(status || "").trim().toLowerCase()) !== -1;
}

function copyMemberChip(ss, member, targetSheet, row, col) {
  if (!member || typeof copyCellChip !== "function") return;
  var mSheet = getTable(ss, "members");
  try { copyCellChip(mSheet, member.row, 1, targetSheet, row, col, member.name, ss.getId ? ss.getId() : ""); } catch (e) {}
}

/**
 * Adds a task row. t: {projectId, member (rtMembers entry), title, description,
 * due (Date), type, recurrence, status}.
 * @return {{row: number, id: string}}
 */
function createTaskRow(ss, t) {
  var sheet = getTable(ss, "tasks");
  var row = (typeof findNextAvailableRow === "function") ? findNextAvailableRow(sheet) : sheet.getLastRow() + 1;
  var id = (typeof generateTaskId === "function") ? generateTaskId() : "TSK-" + Date.now();
  var values = [[
    id, t.projectId || "", t.member ? t.member.name : "", t.title || t.type || "Task", t.description || "",
    t.due ? isoDate(t.due) : "", t.status || "Todo", "", t.type || "", "", "", t.recurrence || ""
  ]];
  sheet.getRange(row, 1, 1, values[0].length).setValues(values);
  copyMemberChip(ss, t.member, sheet, row, taskCol("assigned_member"));
  if (typeof applyTypography === "function") {
    try { applyTypography(sheet.getRange(row, 1, 1, values[0].length)); } catch (eT) {}
  }
  return { row: row, id: id };
}

/** Writes a task's status and keeps Done on / Done by in step (audit). */
function setTaskStatus(ss, row, status, doneBy, asOf) {
  var sheet = getTable(ss, "tasks");
  sheet.getRange(row, taskCol("status")).setValue(status);
  stampTaskDone(sheet, row, status, doneBy, asOf);
}

/** Done on / Done by follow the status: set when it becomes Done, cleared otherwise. */
function stampTaskDone(sheet, row, status, doneBy, asOf) {
  var onCell = sheet.getRange(row, taskCol("done_on"));
  var byCell = sheet.getRange(row, taskCol("done_by"));
  if (String(status || "").trim().toLowerCase() === "done") {
    if (!onCell.getValue()) {
      onCell.setValue(isoDate(rtToday(asOf)));
      byCell.setValue(doneBy || "");
    }
  } else if (onCell.getValue() || byCell.getValue()) {
    onCell.setValue("");
    byCell.setValue("");
  }
}

function setTaskMember(ss, row, member) {
  var sheet = getTable(ss, "tasks");
  sheet.getRange(row, taskCol("assigned_member")).setValue(member ? member.name : "");
  copyMemberChip(ss, member, sheet, row, taskCol("assigned_member"));
}

// ---------------------------------------------------------------------------
// Rules
// ---------------------------------------------------------------------------

function readRecurringRules(ss) {
  var sheet = getTable(ss, "recurringTasks");
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var vals = sheet.getRange(2, 1, readRowLimit(sheet, 200) - 1, 8).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (!v[0] && !v[1]) continue;
    var activeCell = v[7];
    out.push({
      row: i + 2, id: String(v[0] || "").trim() || ("R" + (i + 2)), type: String(v[1] || "").trim(),
      projectId: String(v[2] || "").trim(), mode: /rotat/i.test(String(v[3] || "")) ? "Rotation" : "Each person",
      appliesTo: String(v[4] || ""), start: toDate(v[5]), end: toDate(v[6]),
      active: (activeCell === "" || activeCell === null || activeCell === undefined) ? true : ttTruthy(activeCell)
    });
  }
  return out;
}

/** Adds a rule; returns its ID. */
function addRecurringRule(ss, rule) {
  var sheet = getTable(ss, "recurringTasks");
  var id = "REC-" + Math.random().toString(36).substring(2, 8).toUpperCase();
  var row = (typeof findNextAvailableRow === "function") ? findNextAvailableRow(sheet) : sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 8).setValues([[
    id, rule.type, rule.projectId || "", rule.mode === "Rotation" ? "Rotation" : "Each person", rule.appliesTo || "",
    rule.start ? isoDate(toDate(rule.start)) : "", rule.end ? isoDate(toDate(rule.end)) : "", true
  ]]);
  return id;
}

/**
 * Next person for a rotation turn: among the rule's people working that day
 * and not on approved leave, the one whose last turn is the oldest (never
 * first), in list order on a tie.
 */
function pickRotationMember(ctx, people, ruleTasks, due, excludeKey, periodKey) {
  var lastTurn = {};
  ruleTasks.forEach(function(t) {
    if (String(t.status).toLowerCase() === "cancelled" || t.periodKey === periodKey) return;
    if (!lastTurn[t.memberKey] || t.periodKey > lastTurn[t.memberKey]) lastTurn[t.memberKey] = t.periodKey;
  });
  var best = null;
  people.forEach(function(m, idx) {
    if (m.key === excludeKey) return;
    if (!rtWorksOn(ctx, m, due) || leaveStatusOn(ctx.leaves, m.key, due) === "approved") return;
    var turn = lastTurn[m.key] || "";
    if (!best || turn < best.turn || (turn === best.turn && idx < best.idx)) best = { member: m, turn: turn, idx: idx };
  });
  return best ? best.member : null;
}

function notifyReassignment(ss, from, to, typeName, due) {
  if (typeof MailApp === "undefined") return;
  var when = isoDate(due);
  try {
    if (from && from.email) MailApp.sendEmail({ to: from.email, subject: typeName + " on " + when + ": handed over",
      body: "You are on leave on " + when + ", so " + (to ? to.name : "nobody yet") + " takes your " + typeName + " turn." });
    if (to && to.email) MailApp.sendEmail({ to: to.email, subject: typeName + " on " + when + ": your turn",
      body: (from ? from.name : "The planned person") + " is on leave on " + when + ", so the " + typeName + " turn is yours. See the Calendar tab." });
  } catch (e) { Logger.log("Reassignment email note: " + e); }
}

/**
 * Creates the coming tasks of every active rule and moves rotation turns
 * that fall on someone's approved leave. Runs nightly and after rule changes.
 *
 * @return {{created: number, reassigned: number, unassigned: string[]}}
 */
function generateRecurringTasks(targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var today = rtToday(asOf);
  var res = { created: 0, reassigned: 0, unassigned: [] };
  var rules = readRecurringRules(ss).filter(function(r) { return r.active; });
  if (!rules.length) return res;
  var ctx = rtContext(ss);
  var tasks = readTaskRows(ss);

  rules.forEach(function(rule) {
    var type = findTaskType(ss, rule.type, ctx.types);
    if (!type || !type.active) return;
    var freq = type.frequency === "Monthly" ? "Monthly" : "Weekly";
    var ahead = freq === "Monthly" ? RECURRING_MONTHS_AHEAD : RECURRING_WEEKS_AHEAD;
    var ruleTasks = tasks.filter(function(t) { return t.ruleId === rule.id; });
    var people = resolveAppliesTo(rule.appliesTo, ctx.members);

    for (var i = 0; i < ahead; i++) {
      var p = periodShift(freq, today, i);
      var due = dueDateFor(type, p, ctx.holidays);
      if (due < today) continue;
      if (rule.start && due < rule.start) continue;
      if (rule.end && due > rule.end) continue;
      var inPeriod = ruleTasks.filter(function(t) { return t.periodKey === p.key; });

      if (rule.mode === "Each person") {
        people.forEach(function(m) {
          if (!rtWorksOn(ctx, m, due)) return;
          if (inPeriod.some(function(t) { return t.memberKey === m.key; })) return;
          var made = createTaskRow(ss, { projectId: rule.projectId, member: m, title: type.name, due: due, type: type.name,
            recurrence: rule.id + "|" + p.key });
          var t = { row: made.row, memberKey: m.key, periodKey: p.key, ruleId: rule.id, status: "Todo" };
          ruleTasks.push(t);
          res.created++;
        });
      } else {
        if (inPeriod.length) continue;
        var pick = pickRotationMember(ctx, people, ruleTasks, due, null, p.key);
        if (!pick) { res.unassigned.push(type.name + " " + p.key); continue; }
        var made2 = createTaskRow(ss, { projectId: rule.projectId, member: pick, title: type.name, due: due, type: type.name,
          recurrence: rule.id + "|" + p.key });
        ruleTasks.push({ row: made2.row, memberKey: pick.key, periodKey: p.key, ruleId: rule.id, status: "Todo" });
        res.created++;
      }
    }

    // Rotation turns now falling on an approved leave move to the next person
    if (rule.mode === "Rotation") {
      ruleTasks.forEach(function(t) {
        if (!t.due || t.due < today || t.manual || !isOpenStatus(t.status)) return;
        if (leaveStatusOn(ctx.leaves, t.memberKey, t.due) !== "approved") return;
        var from = ctx.members.filter(function(m) { return m.key === t.memberKey; })[0] || null;
        var to = pickRotationMember(ctx, people, ruleTasks, t.due, t.memberKey, t.periodKey);
        if (!to) return;
        setTaskMember(ss, t.row, to);
        t.memberKey = to.key;
        t.member = to.name;
        notifyReassignment(ss, from, to, type.name, t.due);
        res.reassigned++;
      });
    }
  });
  return res;
}

// ---------------------------------------------------------------------------
// Reminders
// ---------------------------------------------------------------------------

function workingDaysAfter(d, n, holidays) {
  var x = d;
  var count = 0;
  while (count < n) {
    x = addDays(x, 1);
    if (ttIsWorking(x, holidays)) count++;
  }
  return x;
}

function previousWorkingDay(d, holidays) {
  var x = addDays(d, -1);
  while (!ttIsWorking(x, holidays)) x = addDays(x, -1);
  return x;
}

/**
 * Emails people whose typed tasks are due in 3 working days, or were due on
 * the previous working day and are still open. Once a day, on working days.
 *
 * @return {{sent: string[], skipped: (string|undefined)}}
 */
function sendTaskReminders(targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var today = rtToday(asOf);
  var res = { sent: [] };
  var holidays = (typeof readHolidays === "function") ? readHolidays(ss) : {};
  if (!ttIsWorking(today, holidays)) { res.skipped = "non_working_day"; return res; }
  var props = (typeof PropertiesService !== "undefined") ? PropertiesService.getDocumentProperties() : null;
  if (props && props.getProperty("TASK_REMINDERS_SENT_ON") === isoDate(today)) { res.skipped = "already_sent_today"; return res; }

  var soon = isoDate(workingDaysAfter(today, 3, holidays));
  var lastDue = isoDate(previousWorkingDay(today, holidays));
  var members = rtMembers(ss);
  var byMember = {};
  readTaskRows(ss).forEach(function(t) {
    if (!t.type || !t.due || !isOpenStatus(t.status) || !t.memberKey) return;
    var d = isoDate(t.due);
    var kind = d === soon ? "soon" : (d === lastDue ? "late" : null);
    if (!kind) return;
    (byMember[t.memberKey] = byMember[t.memberKey] || []).push({ title: t.title || t.type, due: d, kind: kind });
  });
  Object.keys(byMember).forEach(function(key) {
    var m = members.filter(function(x) { return x.key === key; })[0];
    if (!m || !m.email || m.status === "departed") return;
    var items = byMember[key];
    var lines = items.map(function(it) {
      return (it.kind === "late" ? "Overdue (was due " : "Due ") + it.due + (it.kind === "late" ? "): " : ": ") + it.title;
    });
    MailApp.sendEmail({
      to: m.email,
      subject: items.some(function(it) { return it.kind === "late"; }) ? "Overdue task reminder" : "Tasks due in 3 working days",
      body: "Hi " + m.name.split(/[\s@.]/)[0] + ",\n\n" + lines.join("\n") +
        "\n\nTick them done in Google Tasks or in the sheet (Calendar or Tasks tab):\n" + ss.getUrl() + "\n"
    });
    res.sent.push(m.name);
  });
  if (props) props.setProperty("TASK_REMINDERS_SENT_ON", isoDate(today));
  return res;
}

/** Open typed tasks past their due date, for Check Workspace. */
function findOverdueTasks(ss, asOf) {
  var today = rtToday(asOf);
  return readTaskRows(ss).filter(function(t) {
    return t.type && t.due && t.due < today && isOpenStatus(t.status);
  }).map(function(t) { return (t.title || t.type) + ": " + (t.member || "unassigned") + ", due " + isoDate(t.due); });
}

/** Open tasks assigned to someone on approved leave on the due date. */
function findTasksOnLeave(ss, asOf) {
  var today = rtToday(asOf);
  var leaves = (typeof readCoverageLeaves === "function") ? readCoverageLeaves(ss) : {};
  return readTaskRows(ss).filter(function(t) {
    return t.due && t.due >= today && t.memberKey && isOpenStatus(t.status) && leaveStatusOn(leaves, t.memberKey, t.due) === "approved";
  }).map(function(t) { return (t.title || t.type) + " on " + isoDate(t.due) + ": " + t.member + " is on leave"; });
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    taskCol: taskCol,
    rtMembers: rtMembers,
    resolveAppliesTo: resolveAppliesTo,
    leaveStatusOn: leaveStatusOn,
    rtContext: rtContext,
    rtWorksOn: rtWorksOn,
    readTaskRows: readTaskRows,
    isOpenStatus: isOpenStatus,
    createTaskRow: createTaskRow,
    setTaskStatus: setTaskStatus,
    stampTaskDone: stampTaskDone,
    setTaskMember: setTaskMember,
    readRecurringRules: readRecurringRules,
    addRecurringRule: addRecurringRule,
    pickRotationMember: pickRotationMember,
    generateRecurringTasks: generateRecurringTasks,
    sendTaskReminders: sendTaskReminders,
    findOverdueTasks: findOverdueTasks,
    findTasksOnLeave: findTasksOnLeave
  };
}
