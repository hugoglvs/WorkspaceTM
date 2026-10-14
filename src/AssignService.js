/**
 * Assign Tasks sidebar (ADR 0045): give a task type to several people at once,
 * by role and/or by name, once (this week / month or a chosen one) or as a
 * recurring rule (each person, or a rotation). Also lists the recurring rules
 * and stops them. The type's template (Task Templates tab, ADR 0052) pre-fills
 * the people and the schedule; a schedule changed here is saved on the
 * template, so it applies to everyone.
 */

function scheduleText(t) {
  var f = normalizeFrequency(t.frequency);
  if (f === "One-off") return "One-off";
  return f + (t.due ? ", due " + t.due : "");
}

function getAssignSidebarData(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var members = rtMembers(ss).filter(function(m) { return m.status !== "departed"; });
  var roles = {};
  members.forEach(function(m) { if (m.role) roles[m.role] = true; });
  var types = readTaskTypes(ss).filter(function(t) { return t.active; });
  var projects = projectNames(ss);
  return {
    types: types.map(function(t) {
      return { name: t.name, frequency: t.frequency, due: t.due, remind: t.remindText, mandatory: t.mandatory,
        assignTo: t.assignTo, instructions: t.instructions, schedule: scheduleText(t) };
    }),
    frequencies: FREQUENCIES.filter(function(f) { return f !== "One-off"; }),
    projects: projects,
    roles: Object.keys(roles).sort(),
    members: members.map(function(m) { return { name: m.name, role: m.role }; }),
    rules: readRecurringRules(ss).map(function(r) {
      var rt = types.filter(function(t) { return t.name.toLowerCase() === r.type.toLowerCase(); })[0];
      return { id: r.id, type: r.type, project: r.project || "", mode: r.mode,
        schedule: rt ? scheduleText(rt) : "",
        appliesTo: r.appliesTo || "Everyone", start: r.start ? isoDate(r.start) : "", end: r.end ? isoDate(r.end) : "", active: r.active };
    }),
    today: isoDate(toDate(new Date()))
  };
}

/**
 * payload: {type, project (name), mode ("Each person" | "Rotation"), appliesTo,
 *           repeat (boolean), date (one-off: a day in the target period),
 *           start, end, frequency, due, remind (repeat; the schedule is saved on the template)}
 */
function assignTasks(payload, targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var p = payload || {};
  var today = toDate(asOf || new Date());
  var ctx = rtContext(ss);
  var type = findTaskType(ss, p.type, ctx.types);
  if (!type) return { success: false, message: "Choose a task type." };
  var repeat = !!(p.repeat || p.mode === "Rotation");
  if (repeat && p.frequency && normalizeFrequency(p.frequency) === "One-off") {
    return { success: false, message: "Pick how often it repeats." };
  }
  var people = resolveAppliesTo(p.appliesTo, ctx.members);
  if (!people.length) return { success: false, message: "Nobody matches these roles or names." };
  var mode = p.mode === "Rotation" ? "Rotation" : "Each person";

  var lock = (typeof acquireSyncLock === "function") ? acquireSyncLock(10000) : null;
  if (typeof acquireSyncLock === "function" && !lock) return { success: false, message: "A sync is running. Try again in a minute." };
  try {
    var message;
    if (repeat) {
      var start = toDate(p.start) || today;
      var end = toDate(p.end);
      if (end && end < start) return { success: false, message: "The end date is before the start date." };
      var changed = setTaskTypeSchedule(ss, type.name, p.frequency, p.due, p.remind);
      if (changed) type = findTaskType(ss, type.name);
      addRecurringRule(ss, { type: type.name, project: p.project || "", mode: mode, appliesTo: p.appliesTo || "", start: start, end: end });
      var gen = generateRecurringTasks(ss, today);
      message = "Recurring " + type.name + " saved (" + scheduleText(type) + "; " + (mode === "Rotation" ? "rotation among " : "each of ") +
        people.length + " people). " + gen.created + " task(s) planned ahead." +
        (changed ? " The " + type.name + " schedule changed for everyone." : "");
    } else {
      var freq = periodFrequency(type.frequency);
      var period = periodOf(freq, toDate(p.date) || today);
      var due = dueDateFor(type, period, ctx.holidays);
      var existing = readTaskRows(ss).filter(function(t) {
        return t.due && String(t.status).toLowerCase() !== "cancelled" && t.type.toLowerCase() === type.name.toLowerCase() &&
          (t.project || "").toLowerCase() === (p.project || "").toLowerCase() && periodOf(freq, t.due).key === period.key;
      });
      var created = 0;
      var onLeave = [];
      people.forEach(function(m) {
        if (!rtWorksOn(ctx, m, due)) return;
        if (existing.some(function(t) { return t.memberKey === m.key; })) return;
        if (leaveStatusOn(ctx.leaves, m.key, due) === "approved") onLeave.push(m.name);
        createTaskRow(ss, { project: p.project || "", member: m, title: type.name, due: due, type: type.name,
          description: type.instructions });
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
