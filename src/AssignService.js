/**
 * Assign Tasks sidebar (ADR 0045): give a task type to several people at once,
 * by role and/or by name, once (this week / month or a chosen one) or as a
 * recurring rule (each person, or a rotation). Also lists the recurring rules
 * and stops them.
 */

function getAssignSidebarData(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var members = rtMembers(ss).filter(function(m) { return m.status !== "departed"; });
  var roles = {};
  members.forEach(function(m) { if (m.role) roles[m.role] = true; });
  var types = readTaskTypes(ss).filter(function(t) { return t.active; });
  var projects = calProjects(ss);
  var projectName = {};
  projects.forEach(function(p) { projectName[p.id] = p.name; });
  return {
    types: types.map(function(t) { return { name: t.name, frequency: t.frequency, mandatory: t.mandatory }; }),
    projects: projects,
    roles: Object.keys(roles).sort(),
    members: members.map(function(m) { return { name: m.name, role: m.role }; }),
    rules: readRecurringRules(ss).map(function(r) {
      return { id: r.id, type: r.type, project: projectName[r.projectId] || r.projectId || "", mode: r.mode,
        appliesTo: r.appliesTo || "Everyone", start: r.start ? isoDate(r.start) : "", end: r.end ? isoDate(r.end) : "", active: r.active };
    }),
    today: isoDate(toDate(new Date()))
  };
}

/**
 * payload: {type, projectId, mode ("Each person" | "Rotation"), appliesTo,
 *           repeat (boolean), date (one-off: a day in the target week/month),
 *           start, end (repeat)}
 */
function assignTasks(payload, targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var p = payload || {};
  var today = toDate(asOf || new Date());
  var ctx = rtContext(ss);
  var type = findTaskType(ss, p.type, ctx.types);
  if (!type) return { success: false, message: "Choose a task type." };
  var people = resolveAppliesTo(p.appliesTo, ctx.members);
  if (!people.length) return { success: false, message: "Nobody matches these roles or names." };
  var mode = p.mode === "Rotation" ? "Rotation" : "Each person";

  var lock = (typeof acquireSyncLock === "function") ? acquireSyncLock(10000) : null;
  if (typeof acquireSyncLock === "function" && !lock) return { success: false, message: "A sync is running. Try again in a minute." };
  try {
    var message;
    if (p.repeat || mode === "Rotation") {
      var start = toDate(p.start) || today;
      var end = toDate(p.end);
      if (end && end < start) return { success: false, message: "The end date is before the start date." };
      addRecurringRule(ss, { type: type.name, projectId: p.projectId || "", mode: mode, appliesTo: p.appliesTo || "", start: start, end: end });
      var gen = generateRecurringTasks(ss, today);
      message = "Recurring " + type.name + " saved (" + (mode === "Rotation" ? "rotation among " : "each of ") + people.length +
        " people). " + gen.created + " task(s) planned ahead.";
    } else {
      var freq = type.frequency === "Monthly" ? "Monthly" : "Weekly";
      var period = periodOf(freq, toDate(p.date) || today);
      var due = dueDateFor(type, period, ctx.holidays);
      var existing = readTaskRows(ss).filter(function(t) {
        return t.due && String(t.status).toLowerCase() !== "cancelled" && t.type.toLowerCase() === type.name.toLowerCase() &&
          (t.projectId || "") === (p.projectId || "") && periodOf(freq, t.due).key === period.key;
      });
      var created = 0;
      var onLeave = [];
      people.forEach(function(m) {
        if (!rtWorksOn(ctx, m, due)) return;
        if (existing.some(function(t) { return t.memberKey === m.key; })) return;
        if (leaveStatusOn(ctx.leaves, m.key, due) === "approved") onLeave.push(m.name);
        createTaskRow(ss, { projectId: p.projectId || "", member: m, title: type.name, due: due, type: type.name });
        created++;
      });
      message = created + " " + type.name + " task(s) assigned, due " + isoDate(due) + "." +
        (onLeave.length ? " On leave that day: " + onLeave.join(", ") + "." : "");
    }
    try {
      if (calState().type && calState().type.toLowerCase() === type.name.toLowerCase()) renderTaskCalendar(ss, today);
    } catch (eCal) {}
    return { success: true, message: message };
  } finally {
    if (lock) lock.releaseLock();
  }
}

/** Stops a recurring rule and cancels its future open tasks. */
function stopRecurringRule(ruleId, targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var today = toDate(asOf || new Date());
  var rule = readRecurringRules(ss).filter(function(r) { return r.id === ruleId; })[0];
  if (!rule) return { success: false, message: "Rule not found." };
  getTable(ss, "recurringTasks").getRange(rule.row, 8).setValue(false);
  var cancelled = 0;
  readTaskRows(ss).forEach(function(t) {
    if (t.ruleId === ruleId && t.due && t.due > today && isOpenStatus(t.status) && !t.manual) {
      setTaskStatus(ss, t.row, "Cancelled", today);
      cancelled++;
    }
  });
  return { success: true, message: "Stopped. " + cancelled + " future task(s) cancelled." };
}

/** Menu: Assign Tasks. */
function showAssignSidebar() {
  var html = HtmlService.createTemplateFromFile("AssignSidebar").evaluate().setTitle("Assign Tasks").setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Menu: Calendar (shows and refreshes the tab). */
function showTaskCalendar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = renderTaskCalendar(ss);
  var sheet = getTable(ss, "taskCalendar");
  if (sheet) { if (sheet.isSheetHidden()) sheet.showSheet(); sheet.activate(); }
  return res;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    getAssignSidebarData: getAssignSidebarData,
    assignTasks: assignTasks,
    stopRecurringRule: stopRecurringRule,
    showAssignSidebar: showAssignSidebar,
    showTaskCalendar: showTaskCalendar
  };
}
