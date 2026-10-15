/**
 * Calendar view (ADR 0045, 0052): the Tasks list as a planning grid, like
 * Weekly Presence. Pick a Type (B1) and a Project (D1); rows are members,
 * columns are the type's periods (weeks, half months or months). Row 2 shows
 * the year above its periods, row 3 the periods, row 4 counts the tasks done
 * per period ("9/12"). A cell shows that member's task of that
 * type in that period:
 *   checkbox, unticked  assigned        (blue; red once overdue)
 *   checkbox, ticked    done            (green)
 *   empty               nothing         (grey when on approved leave, orange when a leave is requested)
 * Editing a cell writes Tasks: inserting a checkbox assigns (refused on
 * approved leave), ticking marks done (Done on stamped), unticking reopens,
 * deleting the checkbox cancels the task. "To do" / "Done" typed by hand still work.
 * Changing a period before the current month asks for confirmation (audit).
 */

var CALENDAR_STATE_PROP = "TASK_CALENDAR_STATE";
var CALENDAR_ALL_PROJECTS = "All projects";
var CALENDAR_FIRST_ROW = 5;
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

/** Project names (ADR 0053: the name is the reference). */
function calProjects(ss) {
  return projectNames(ss);
}

function calPeriods(type, today) {
  var freq = calFrequency(type);
  var span = { Monthly: [-3, 8], "Twice a month": [-4, 10], Weekly: [-4, 12] }[freq];
  var out = [];
  for (var i = span[0]; i <= span[1]; i++) out.push(periodShift(freq, today, i));
  return out;
}

function calFrequency(type) {
  return periodFrequency(type && type.frequency);
}

/** Year a period belongs to (its key's year: the ISO year for weeks). */
function calPeriodYear(p) {
  return parseInt(String(p.key).slice(0, 4), 10);
}

/** A period's column label without its year, which row 2 shows. */
function calPeriodLabel(p) {
  return String(p.label).replace(/\s+\d{4}$/, "");
}

/** A Calendar cell value as "done", "todo" or "" (checkbox or text). */
function calCellState(v) {
  if (v === true) return "done";
  if (v === false) return "todo";
  var s = String(v === undefined || v === null ? "" : v).trim().toLowerCase();
  if (!s) return "";
  if (s === "true" || s === "done" || s === "x" || s === "✓") return "done";
  return "todo";
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
  var project = projects.filter(function(n) { return n.toLowerCase() === String(state.project || "").toLowerCase(); })[0] || null;
  var periods = calPeriods(type, today);
  var freq = calFrequency(type);
  var people = ctx.members.filter(function(m) { return m.status !== "departed"; });

  var tasks = readTaskRows(ss).filter(function(t) {
    if (!t.due || String(t.status).toLowerCase() === "cancelled") return false;
    if (t.type.toLowerCase() !== type.name.toLowerCase()) return false;
    return !project || t.project.toLowerCase() === project.toLowerCase();
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
  var checks = [];
  var counts = periods.map(function() { return { done: 0, total: 0 }; });
  var checkbox = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  people.forEach(function(m) {
    var row = [m.name];
    var bg = ["#FFFFFF"];
    var nt = [""];
    var ck = [];
    periods.forEach(function(p, pi) {
      var t = cellTask[m.key + "|" + p.key];
      var due = dueDateFor(type, p, ctx.holidays);
      var leave = leaveStatusOn(ctx.leaves, m.key, due);
      ck.push(t ? checkbox : null);
      if (t) {
        var done = String(t.status).toLowerCase() === "done";
        counts[pi].total++;
        if (done) counts[pi].done++;
        row.push(done);
        bg.push(done ? CALENDAR_STYLES.done : (t.due < today ? CALENDAR_STYLES.late : CALENDAR_STYLES.todo));
        nt.push(done ? "Done" + (t.doneOn ? " on " + isoDate(toDate(t.doneOn)) : "") + "" : "Due " + isoDate(t.due) + (leave === "approved" ? " · on leave!" : ""));
      } else {
        row.push("");
        bg.push(leave === "approved" ? CALENDAR_STYLES.leave : (leave === "requested" ? CALENDAR_STYLES.requested : CALENDAR_STYLES.none));
        nt.push(leave === "approved" ? "On leave" : (leave === "requested" ? "Leave requested" : ""));
      }
    });
    values.push(row);
    bgs.push(bg);
    notes.push(nt);
    checks.push(ck);
  });

  try { sheet.getRange(2, 1, 1, Math.max(sheet.getMaxColumns(), width)).breakApart(); } catch (eBa) {}
  sheet.clear();
  try { sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 3), Math.max(sheet.getMaxColumns(), width)).clearDataValidations(); } catch (eV) {}
  sheet.getRange(1, 1, 1, 5).setValues([["Type", type.name, "Project", project || CALENDAR_ALL_PROJECTS,
    "Tick = done · Insert → Checkbox in an empty cell = assign · delete the checkbox = remove"]]);
  sheet.getRange(1, 2).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(activeTypes.map(function(t) { return t.name; }), true).setAllowInvalid(false).build());
  sheet.getRange(1, 4).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([CALENDAR_ALL_PROJECTS].concat(projects), true).setAllowInvalid(false).build());
  try {
    sheet.getRange(1, 1).setFontWeight("bold");
    sheet.getRange(1, 3).setFontWeight("bold");
    sheet.getRange(1, 5).setFontColor("#5F6368");
  } catch (eF) {}

  // Row 2: the year above its periods (merged per year); row 3: the periods
  var years = periods.map(calPeriodYear);
  try { sheet.getRange(2, 1, 2, width).setNumberFormat("@"); } catch (eNf) {}
  sheet.getRange(2, 1, 1, width).setValues([["Year"].concat(years.map(function(y, i) { return i === 0 || y !== years[i - 1] ? String(y) : ""; }))]);
  // Plain text, or Sheets turns "Oct" into a date
  sheet.getRange(3, 1, 1, width).setValues([["Member"].concat(periods.map(calPeriodLabel))]);
  try {
    sheet.getRange(2, 1, 2, width).setFontWeight("bold").setBackground("#E8EAED");
    sheet.getRange(2, 2, 1, periods.length).setHorizontalAlignment("center");
    var runStart = 0;
    for (var yi = 1; yi <= years.length; yi++) {
      if (yi === years.length || years[yi] !== years[runStart]) {
        if (yi - runStart > 1) sheet.getRange(2, runStart + 2, 1, yi - runStart).merge();
        runStart = yi;
      }
    }
    var current = periodOf(freq, today).key;
    var curIdx = periods.map(function(p) { return p.key; }).indexOf(current);
    if (curIdx >= 0) sheet.getRange(3, curIdx + 2).setBackground("#D2E3FC");
  } catch (eH) {}

  // Row 4: done per period, for the manager's check (ADR 0052)
  try { sheet.getRange(4, 1, 1, width).setNumberFormat("@"); } catch (eNf3) {}
  sheet.getRange(4, 1, 1, width).setValues([["Done"].concat(counts.map(function(c) { return c.total ? c.done + "/" + c.total : ""; }))]);
  try {
    sheet.getRange(4, 1, 1, width).setFontColor("#5F6368").setHorizontalAlignment("center");
    sheet.getRange(4, 1).setHorizontalAlignment("left");
    sheet.getRange(4, 2, 1, periods.length).setBackgrounds([counts.map(function(c, i) {
      if (!c.total) return "#FFFFFF";
      if (c.done === c.total) return CALENDAR_STYLES.done;
      return periods[i].end < today ? CALENDAR_STYLES.late : "#FFFFFF";
    })]);
  } catch (eC) {}

  if (values.length) {
    var grid = sheet.getRange(CALENDAR_FIRST_ROW, 1, values.length, width);
    grid.setValues(values);
    grid.setBackgrounds(bgs);
    try { grid.setNotes(notes); } catch (eN) {}
    sheet.getRange(CALENDAR_FIRST_ROW, 2, values.length, periods.length).setDataValidations(checks);
    try { sheet.getRange(CALENDAR_FIRST_ROW, 2, values.length, periods.length).setHorizontalAlignment("center"); } catch (eA) {}
  }
  try {
    sheet.setFrozenRows(4);
    sheet.setFrozenColumns(1);
  } catch (eFz) {}

  calSaveState({
    type: type.name, project: project || "",
    periodKeys: periods.map(function(p) { return p.key; }),
    memberKeys: people.map(function(m) { return m.key; })
  });
  return { success: true, type: type.name, project: project || CALENDAR_ALL_PROJECTS, members: people.length, periods: periods.length };
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
      st.project = name === CALENDAR_ALL_PROJECTS ? "" : name;
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
  var newVal = calCellState(e.range.getValue());
  var label = type.name + " · " + member.name + " · " + period.label;

  if (isClosedPeriod(period, today) && !confirmClosedChange(label)) {
    renderTaskCalendar(ss, today);
    return "kept";
  }

  var existing = readTaskRows(ss).filter(function(t) {
    return t.due && String(t.status).toLowerCase() !== "cancelled" && t.memberKey === memberKey &&
      t.type.toLowerCase() === type.name.toLowerCase() && (!state.project || t.project.toLowerCase() === String(state.project).toLowerCase()) &&
      periodOf(freq, t.due).key === periodKey;
  });
  var due = dueDateFor(type, period, ctx.holidays);
  var outcome;

  if (newVal === "") {
    existing.forEach(function(t) {
      setTaskStatus(ss, t.row, "Cancelled", today);
      if (t.ruleId && !t.manual) getTable(ss, "tasks").getRange(t.row, taskCol("origin")).setValue(t.ruleId + "|" + t.periodKey + "|manual");
    });
    outcome = existing.length ? "cancelled" : "nothing";
  } else if (newVal === "done") {
    if (existing.length) existing.forEach(function(t) { setTaskStatus(ss, t.row, "Done", today); });
    else {
      var made = createTaskRow(ss, { project: state.project || "", member: member, title: type.name, due: due, type: type.name,
        description: type.instructions });
      setTaskStatus(ss, made.row, "Done", today);
    }
    outcome = "done";
  } else {
    if (existing.length) {
      existing.forEach(function(t) { if (String(t.status).toLowerCase() === "done") setTaskStatus(ss, t.row, "Todo", today); });
      outcome = "reopened";
    } else {
      var leave = leaveStatusOn(ctx.leaves, memberKey, due);
      if (leave === "approved") {
        renderTaskCalendar(ss, today);
        try { ss.toast(member.name + " is on leave on " + isoDate(due) + ". Pick someone else.", "On leave", 6); } catch (eT2) {}
        return "refused_leave";
      }
      if (leave === "requested") {
        try { ss.toast(member.name + " has requested leave on " + isoDate(due) + ".", "Leave requested", 6); } catch (eT3) {}
      }
      createTaskRow(ss, { project: state.project || "", member: member, title: type.name, due: due, type: type.name,
        description: type.instructions });
      outcome = "assigned";
    }
  }
  renderTaskCalendar(ss, today);
  return outcome;
}

/**
 * onEdit of Tasks.Status: Done on follows the status, and a change in
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
  stampTaskDone(sheet, row, status, today);
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
