/**
 * Calendar view (ADR 0045): the Tasks list as a planning grid, like Weekly
 * Presence. Pick a Type (B1) and a Project (D1); rows are members, columns are
 * weeks (or months for monthly types). A cell shows that member's task of
 * that type in that period:
 *   "To do"  assigned / in charge   (blue; red once overdue)
 *   "Done"   done                   (green)
 *   empty    nothing                (grey when on approved leave, orange when a leave is requested)
 * Editing a cell writes Tasks: "To do" assigns (refused on approved leave),
 * "Done" marks done (Done on / Done by stamped), clearing cancels the task.
 * Changing a period before the current month asks for confirmation (audit).
 */

var CALENDAR_STATE_PROP = "TASK_CALENDAR_STATE";
var CALENDAR_ALL_PROJECTS = "All projects";
var CALENDAR_FIRST_ROW = 3;
var CALENDAR_STYLES = {
  todo: "#E8F0FE", late: "#FCE8E6", done: "#E6F4EA", leave: "#E8EAED", requested: "#FEF7E0", none: "#FFFFFF"
};

function calProps() {
  try { return PropertiesService.getDocumentProperties(); } catch (e) { return null; }
}

function calState() {
  var p = calProps();
  try { return JSON.parse((p && p.getProperty(CALENDAR_STATE_PROP)) || "{}") || {}; } catch (e) { return {}; }
}

function calSaveState(state) {
  var p = calProps();
  if (p) p.setProperty(CALENDAR_STATE_PROP, JSON.stringify(state));
}

/** Projects as {id, name} (named projects with an ID). */
function calProjects(ss) {
  var sheet = getTable(ss, "projects");
  if (!sheet || sheet.getMaxRows() < 2) return [];
  return sheet.getRange(2, 1, readRowLimit(sheet, 500) - 1, 2).getValues()
    .filter(function(r) { return String(r[0] || "").trim() && String(r[1] || "").trim(); })
    .map(function(r) { return { id: String(r[0]).trim(), name: String(r[1]).trim() }; });
}

function calPeriods(type, today) {
  var monthly = type && type.frequency === "Monthly";
  var out = [];
  var from = monthly ? -3 : -4;
  var to = monthly ? 8 : 12;
  for (var i = from; i <= to; i++) out.push(periodShift(monthly ? "Monthly" : "Weekly", today, i));
  return out;
}

function calFrequency(type) {
  return type && type.frequency === "Monthly" ? "Monthly" : "Weekly";
}

/**
 * Builds or refreshes the Calendar tab for the chosen type and project.
 * @return {{success: boolean, type: string, project: string, members: number, periods: number}}
 */
function renderTaskCalendar(targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var today = toDate(asOf || new Date());
  var sheet = getTable(ss, "taskCalendar");
  if (!sheet) sheet = ss.insertSheet(tableSheetName("taskCalendar"));

  var ctx = rtContext(ss);
  var activeTypes = ctx.types.filter(function(t) { return t.active; });
  var state = calState();
  var type = findTaskType(ss, state.type, ctx.types) || activeTypes[0] || ctx.types[0];
  var projects = calProjects(ss);
  var project = projects.filter(function(p) { return p.id === state.projectId; })[0] || null;
  var periods = calPeriods(type, today);
  var freq = calFrequency(type);
  var people = ctx.members.filter(function(m) { return m.status !== "departed"; });

  var tasks = readTaskRows(ss).filter(function(t) {
    if (!t.due || String(t.status).toLowerCase() === "cancelled") return false;
    if (t.type.toLowerCase() !== type.name.toLowerCase()) return false;
    return !project || t.projectId === project.id;
  });
  var cellTask = {};
  tasks.forEach(function(t) {
    var k = t.memberKey + "|" + periodOf(freq, t.due).key;
    if (!cellTask[k] || String(t.status).toLowerCase() !== "done") cellTask[k] = t;
  });

  var width = periods.length + 1;
  var values = [];
  var bgs = [];
  var notes = [];
  people.forEach(function(m) {
    var row = [m.name];
    var bg = ["#FFFFFF"];
    var nt = [""];
    periods.forEach(function(p) {
      var t = cellTask[m.key + "|" + p.key];
      var due = dueDateFor(type, p, ctx.holidays);
      var leave = leaveStatusOn(ctx.leaves, m.key, due);
      if (t) {
        var done = String(t.status).toLowerCase() === "done";
        row.push(done ? "Done" : "To do");
        bg.push(done ? CALENDAR_STYLES.done : (t.due < today ? CALENDAR_STYLES.late : CALENDAR_STYLES.todo));
        nt.push(done ? "Done" + (t.doneOn ? " on " + isoDate(toDate(t.doneOn)) : "") + (t.doneBy ? " by " + t.doneBy : "") : "Due " + isoDate(t.due) + (leave === "approved" ? " · on leave!" : ""));
      } else {
        row.push("");
        bg.push(leave === "approved" ? CALENDAR_STYLES.leave : (leave === "requested" ? CALENDAR_STYLES.requested : CALENDAR_STYLES.none));
        nt.push(leave === "approved" ? "On leave" : (leave === "requested" ? "Leave requested" : ""));
      }
    });
    values.push(row);
    bgs.push(bg);
    notes.push(nt);
  });

  sheet.clear();
  try { sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 3), Math.max(sheet.getMaxColumns(), width)).clearDataValidations(); } catch (eV) {}
  sheet.getRange(1, 1, 1, 5).setValues([["Type", type.name, "Project", project ? project.name : CALENDAR_ALL_PROJECTS,
    "To do = assigned / in charge · Done = done · clear = remove"]]);
  sheet.getRange(1, 2).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(activeTypes.map(function(t) { return t.name; }), true).setAllowInvalid(false).build());
  sheet.getRange(1, 4).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([CALENDAR_ALL_PROJECTS].concat(projects.map(function(p) { return p.name; })), true).setAllowInvalid(false).build());
  try {
    sheet.getRange(1, 1).setFontWeight("bold");
    sheet.getRange(1, 3).setFontWeight("bold");
    sheet.getRange(1, 5).setFontColor("#5F6368");
  } catch (eF) {}

  var header = ["Member"].concat(periods.map(function(p) { return p.label; }));
  // Plain text, or Sheets turns "Oct 2026" into a date
  try { sheet.getRange(2, 1, 1, width).setNumberFormat("@"); } catch (eNf) {}
  sheet.getRange(2, 1, 1, width).setValues([header]);
  try {
    sheet.getRange(2, 1, 1, width).setFontWeight("bold").setBackground("#E8EAED");
    var current = periodOf(freq, today).key;
    var curIdx = periods.map(function(p) { return p.key; }).indexOf(current);
    if (curIdx >= 0) sheet.getRange(2, curIdx + 2).setBackground("#D2E3FC");
  } catch (eH) {}

  if (values.length) {
    var grid = sheet.getRange(CALENDAR_FIRST_ROW, 1, values.length, width);
    grid.setValues(values);
    grid.setBackgrounds(bgs);
    try { grid.setNotes(notes); } catch (eN) {}
    sheet.getRange(CALENDAR_FIRST_ROW, 2, values.length, periods.length).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(["To do", "Done"], true).setAllowInvalid(false).build());
  }
  try {
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(1);
  } catch (eFz) {}

  calSaveState({
    type: type.name, projectId: project ? project.id : "",
    periodKeys: periods.map(function(p) { return p.key; }),
    memberKeys: people.map(function(m) { return m.key; })
  });
  return { success: true, type: type.name, project: project ? project.name : CALENDAR_ALL_PROJECTS, members: people.length, periods: periods.length };
}

function calEditorName(e, members) {
  var email = "";
  try { if (e && e.user && typeof e.user.getEmail === "function") email = e.user.getEmail(); } catch (eU) {}
  if (!email) { try { email = Session.getActiveUser().getEmail(); } catch (eS) {} }
  var m = members.filter(function(x) {
    return email && ((x.email && x.email.toLowerCase() === email.toLowerCase()) || (x.googleEmail && x.googleEmail.toLowerCase() === email.toLowerCase()));
  })[0];
  return m ? m.name : email;
}

/** A period ending before the current month is closed (audit). */
function isClosedPeriod(period, today) {
  return period.end < new Date(today.getFullYear(), today.getMonth(), 1);
}

function confirmClosedChange(label) {
  try {
    var ui = SpreadsheetApp.getUi();
    var answer = ui.alert("Closed period", label + " is in a closed month. Change it anyway? (The change is recorded for the audit.)", ui.ButtonSet.YES_NO);
    return answer === ui.Button.YES;
  } catch (e) {
    return true; // no UI (tests, triggers without a user): allow
  }
}

/**
 * onEdit of the Calendar tab.
 * @return {string|null} what happened
 */
function handleTaskCalendarEdit(e, asOf) {
  if (!e || !e.range) return null;
  var sheet = e.range.getSheet();
  var ss = e.source || (sheet.getParent ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet());
  var r = e.range.getRow();
  var c = e.range.getColumn();
  var today = toDate(asOf || new Date());

  // Type or project picked
  if (r === 1 && (c === 2 || c === 4)) {
    var st = calState();
    if (c === 2) st.type = String(sheet.getRange(1, 2).getValue() || "");
    if (c === 4) {
      var name = String(sheet.getRange(1, 4).getValue() || "");
      var proj = calProjects(ss).filter(function(p) { return p.name === name; })[0];
      st.projectId = proj ? proj.id : "";
    }
    calSaveState(st);
    renderTaskCalendar(ss, today);
    return "rendered";
  }
  if (r < CALENDAR_FIRST_ROW || c < 2) return null;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
    renderTaskCalendar(ss, today);
    try { ss.toast("Change one cell at a time in the Calendar.", "Calendar", 5); } catch (eT) {}
    return "multi";
  }

  var state = calState();
  var memberKey = (state.memberKeys || [])[r - CALENDAR_FIRST_ROW];
  var periodKey = (state.periodKeys || [])[c - 2];
  if (!memberKey || !periodKey) { renderTaskCalendar(ss, today); return null; }

  var ctx = rtContext(ss);
  var type = findTaskType(ss, state.type, ctx.types);
  var member = ctx.members.filter(function(m) { return m.key === memberKey; })[0];
  if (!type || !member) { renderTaskCalendar(ss, today); return null; }
  var freq = calFrequency(type);
  var period = periodFromKey(periodKey);
  var newVal = String(e.range.getValue() || "").trim();
  var oldVal = e.oldValue === undefined ? "" : String(e.oldValue);
  var label = type.name + " · " + member.name + " · " + period.label;

  if (isClosedPeriod(period, today) && !confirmClosedChange(label)) {
    e.range.setValue(oldVal);
    return "kept";
  }

  var existing = readTaskRows(ss).filter(function(t) {
    return t.due && String(t.status).toLowerCase() !== "cancelled" && t.memberKey === memberKey &&
      t.type.toLowerCase() === type.name.toLowerCase() && (!state.projectId || t.projectId === state.projectId) &&
      periodOf(freq, t.due).key === periodKey;
  });
  var due = dueDateFor(type, period, ctx.holidays);
  var by = calEditorName(e, ctx.members);
  var outcome;

  if (newVal === "") {
    existing.forEach(function(t) {
      setTaskStatus(ss, t.row, "Cancelled", by, today);
      if (t.ruleId && !t.manual) getTable(ss, "tasks").getRange(t.row, taskCol("recurrence")).setValue(t.ruleId + "|" + t.periodKey + "|manual");
    });
    outcome = existing.length ? "cancelled" : "nothing";
  } else if (newVal === "Done") {
    if (existing.length) existing.forEach(function(t) { setTaskStatus(ss, t.row, "Done", by, today); });
    else {
      var made = createTaskRow(ss, { projectId: state.projectId, member: member, title: type.name, due: due, type: type.name });
      setTaskStatus(ss, made.row, "Done", by, today);
    }
    outcome = "done";
  } else {
    if (existing.length) {
      existing.forEach(function(t) { if (String(t.status).toLowerCase() === "done") setTaskStatus(ss, t.row, "Todo", by, today); });
      outcome = "reopened";
    } else {
      var leave = leaveStatusOn(ctx.leaves, memberKey, due);
      if (leave === "approved") {
        e.range.setValue(oldVal);
        try { ss.toast(member.name + " is on leave on " + isoDate(due) + ". Pick someone else.", "On leave", 6); } catch (eT2) {}
        return "refused_leave";
      }
      if (leave === "requested") {
        try { ss.toast(member.name + " has requested leave on " + isoDate(due) + ".", "Leave requested", 6); } catch (eT3) {}
      }
      createTaskRow(ss, { projectId: state.projectId, member: member, title: type.name, due: due, type: type.name });
      outcome = "assigned";
    }
  }
  renderTaskCalendar(ss, today);
  return outcome;
}

/**
 * onEdit of Tasks.Status: Done on / Done by follow the status, and a change in
 * a closed month asks for confirmation first.
 */
function handleTaskStatusEdit(e, asOf) {
  if (!e || !e.range || e.range.getColumn() !== taskCol("status") || e.range.getRow() < 2 || e.range.getNumRows() > 1) return null;
  var sheet = e.range.getSheet();
  var ss = e.source || (sheet.getParent ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet());
  var today = toDate(asOf || new Date());
  var row = e.range.getRow();
  var due = toDate(sheet.getRange(row, taskCol("due_date")).getValue());
  var status = String(e.range.getValue() || "");
  if (due && due < new Date(today.getFullYear(), today.getMonth(), 1) && e.oldValue !== undefined &&
      !confirmClosedChange(String(sheet.getRange(row, taskCol("title")).getValue() || "This task") + " (due " + isoDate(due) + ")")) {
    e.range.setValue(e.oldValue);
    return "kept";
  }
  stampTaskDone(sheet, row, status, calEditorName(e, rtMembers(ss)), today);
  return "stamped";
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    renderTaskCalendar: renderTaskCalendar,
    handleTaskCalendarEdit: handleTaskCalendarEdit,
    handleTaskStatusEdit: handleTaskStatusEdit,
    isClosedPeriod: isClosedPeriod,
    calProjects: calProjects
  };
}
