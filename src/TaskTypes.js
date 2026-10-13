/**
 * Task types and task periods (ADR 0045).
 *
 * A Task Type (Quiz, Onboarding, Time & expenses…) says how often a task comes
 * back and when it is due. The hidden "Task Types" tab lists them:
 * Type | Frequency (Weekly, Monthly, One-off) | Due | Mandatory | Active.
 * Due: a weekday for weekly types ("Friday"); for monthly types a day number
 * ("20"), "Last working day" or "First working day". A due date that falls on a
 * weekend or public holiday moves to the working day before (after, for "First
 * working day"). Mandatory types are the audited ones (reminders, overdue list).
 * Created when / Days after / Assign to / Instructions (ADR 0051): a type created
 * by an event (src/TaskTriggers.js), e.g. Onboarding when a new member arrives.
 */

var ONBOARDING_INSTRUCTIONS = "Welcome! Please enter your Accenture email address and phone number in the Members tab (columns Accenture email and Phone).";
var DEFAULT_TASK_TYPES = [
  ["Onboarding", "One-off", "", false, true, "New member", 7, "Person concerned", ONBOARDING_INSTRUCTIONS],
  ["Quiz", "Weekly", "Friday", false, true],
  ["Team presentation", "Weekly", "Thursday", false, true],
  ["Peer review", "Monthly", "Last working day", true, true],
  ["SoW", "Monthly", "Last working day", true, true],
  ["PRM", "Monthly", "Last working day", true, true],
  ["Forecast", "Monthly", "Last working day", true, true],
  ["Time & expenses", "Monthly", "Last working day", true, true]
];
var TT_WEEKDAYS = ["sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"];
var TT_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];




function ttTruthy(v) {
  if (v === true) return true;
  if (v === false) return false;
  var s = String(v === undefined || v === null ? "" : v).trim().toLowerCase();
  return s === "true" || s === "yes" || s === "y" || s === "1" || s === "x" || s === "✓";
}

function normalizeFrequency(f) {
  var s = String(f || "").trim().toLowerCase();
  if (s.indexOf("month") === 0) return "Monthly";
  if (s.indexOf("one") === 0 || s === "once") return "One-off";
  return "Weekly";
}

function ttFromRow(v) {
  var activeCell = v[4];
  var days = parseInt(v[6], 10);
  return {
    name: String(v[0] || "").trim(), frequency: normalizeFrequency(v[1]), due: String(v[2] || "").trim(),
    mandatory: ttTruthy(v[3]),
    active: (activeCell === "" || activeCell === null || activeCell === undefined) ? true : ttTruthy(activeCell),
    createdWhen: String(v[5] || "").trim(), daysAfter: isNaN(days) ? 0 : days,
    assignTo: String(v[7] || "").trim(), instructions: String(v[8] || "").trim()
  };
}

/**
 * Every task type, in tab order. Falls back to the defaults when the tab is
 * missing or empty.
 * @return {Array<{name: string, frequency: string, due: string, mandatory: boolean, active: boolean,
 *   createdWhen: string, daysAfter: number, assignTo: string, instructions: string}>}
 */
function readTaskTypes(ss) {
  var sheet = (ss && typeof getTable === "function") ? getTable(ss, "taskTypes") : null;
  var list = [];
  if (sheet && sheet.getMaxRows() > 1) {
    var vals = sheet.getRange(2, 1, readRowLimit(sheet, 100) - 1, 9).getValues();
    for (var r = 0; r < vals.length; r++) {
      if (!String(vals[r][0] || "").trim()) continue;
      list.push(ttFromRow(vals[r]));
    }
  }
  if (!list.length) list = DEFAULT_TASK_TYPES.map(ttFromRow);
  return list;
}

function taskTypeNames(ss) {
  return readTaskTypes(ss).map(function(t) { return t.name; });
}

function findTaskType(ss, name, list) {
  var key = String(name || "").trim().toLowerCase();
  var types = list || readTaskTypes(ss);
  for (var i = 0; i < types.length; i++) if (types[i].name.toLowerCase() === key) return types[i];
  return null;
}

/** Writes the default task types into an empty Task Types tab. */
function seedDefaultTaskTypes(ss) {
  var sheet = getTable(ss, "taskTypes");
  if (!sheet) return 0;
  var existing = sheet.getMaxRows() > 1 ? sheet.getRange(2, 1, Math.min(sheet.getMaxRows(), 50) - 1, 1).getValues() : [];
  if (existing.some(function(r) { return String(r[0] || "").trim(); })) {
    upgradeTaskTypeTriggers(ss, sheet);
    return 0;
  }
  sheet.getRange(2, 1, DEFAULT_TASK_TYPES.length, 9).setValues(DEFAULT_TASK_TYPES);
  return DEFAULT_TASK_TYPES.length;
}

/**
 * Once per sheet (ADR 0051): default types already in the tab get their
 * trigger columns (Onboarding: New member, 7 days, Person concerned) when
 * those are still empty. Never runs again, so a team can clear them.
 */
function upgradeTaskTypeTriggers(ss, sheet) {
  var props = null;
  try { props = (typeof PropertiesService !== "undefined") ? PropertiesService.getDocumentProperties() : null; } catch (eP) {}
  if (!props || props.getProperty("TASK_TRIGGERS_SEEDED")) return 0;
  var n = readRowLimit(sheet, 100) - 1;
  var vals = n > 0 ? sheet.getRange(2, 1, n, 9).getValues() : [];
  var updated = 0;
  vals.forEach(function(v, i) {
    var d = DEFAULT_TASK_TYPES.filter(function(x) { return x[5] && x[0].toLowerCase() === String(v[0] || "").trim().toLowerCase(); })[0];
    if (!d || String(v[5] || "").trim()) return;
    sheet.getRange(i + 2, 6, 1, 4).setValues([[d[5], d[6], d[7], d[8]]]);
    updated++;
  });
  props.setProperty("TASK_TRIGGERS_SEEDED", "1");
  return updated;
}

// ---------------------------------------------------------------------------
// Periods: a week ("2026-W42") or a month ("2026-10")
// ---------------------------------------------------------------------------

/** ISO week key of a date, e.g. "2026-W42". */
function isoWeekKey(d) {
  var t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  var day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  var yearStart = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  var week = Math.ceil(((t - yearStart) / 86400000 + 1) / 7);
  return t.getUTCFullYear() + "-W" + ("0" + week).slice(-2);
}

function ttMonday(d) {
  return addDays(d, -((d.getDay() + 6) % 7));
}

/**
 * The period holding a date.
 * @return {{key: string, start: Date, end: Date, label: string, frequency: string}}
 */
function periodOf(frequency, date) {
  var d = toDate(date);
  if (normalizeFrequency(frequency) === "Monthly") {
    var start = new Date(d.getFullYear(), d.getMonth(), 1);
    var end = new Date(d.getFullYear(), d.getMonth() + 1, 0);
    return { key: start.getFullYear() + "-" + ("0" + (start.getMonth() + 1)).slice(-2), start: start, end: end,
      label: TT_MONTHS[start.getMonth()] + " " + start.getFullYear(), frequency: "Monthly" };
  }
  var mon = ttMonday(d);
  var key = isoWeekKey(mon);
  return { key: key, start: mon, end: addDays(mon, 6),
    label: "W" + key.slice(-2) + " · " + mon.getDate() + " " + TT_MONTHS[mon.getMonth()], frequency: "Weekly" };
}

/** Period n periods after the one holding `date` (n may be negative). */
function periodShift(frequency, date, n) {
  var d = toDate(date);
  if (normalizeFrequency(frequency) === "Monthly") return periodOf("Monthly", new Date(d.getFullYear(), d.getMonth() + n, 1));
  return periodOf("Weekly", addDays(ttMonday(d), 7 * n));
}

/** Period from its key ("2026-W42" or "2026-10"). */
function periodFromKey(key) {
  var m = String(key || "").match(/^(\d{4})-(\d{2})$/);
  if (m) return periodOf("Monthly", new Date(+m[1], +m[2] - 1, 1));
  var w = String(key || "").match(/^(\d{4})-W(\d{2})$/);
  if (!w) return null;
  // Monday of ISO week 1 is the Monday of the week holding 4 January
  var week1 = ttMonday(new Date(+w[1], 0, 4));
  return periodOf("Weekly", addDays(week1, 7 * (+w[2] - 1)));
}

function ttIsWorking(d, holidays) {
  if (typeof isWorkingDay === "function") return isWorkingDay(d, holidays);
  return d.getDay() !== 0 && d.getDay() !== 6 && !(holidays && holidays[isoDate(d)]);
}

/**
 * Due date of a task type in a period. Weekends and public holidays move to
 * the working day before (after for "First working day"), staying in the period
 * when possible.
 *
 * @param {{frequency: string, due: string}} type
 * @param {{start: Date, end: Date}} period
 * @param {Object<string, string>} [holidays]
 * @return {Date}
 */
function dueDateFor(type, period, holidays) {
  var rule = String(type && type.due || "").trim().toLowerCase();
  var d;
  if (normalizeFrequency(type && type.frequency) === "Monthly") {
    if (rule.indexOf("first") === 0) {
      d = new Date(period.start.getTime());
      while (!ttIsWorking(d, holidays) && d < period.end) d = addDays(d, 1);
      return d;
    }
    var n = parseInt(rule, 10);
    if (!isNaN(n) && n > 0) d = new Date(period.start.getFullYear(), period.start.getMonth(), Math.min(n, period.end.getDate()));
    else d = new Date(period.end.getTime()); // "Last working day" and default
  } else {
    var idx = TT_WEEKDAYS.indexOf(rule);
    if (idx < 0) idx = 5; // Friday
    d = addDays(period.start, (idx + 6) % 7);
  }
  var due = d;
  while (!ttIsWorking(due, holidays) && due > period.start) due = addDays(due, -1);
  if (ttIsWorking(due, holidays)) return due;
  // Nothing earlier in the period: take the next working day
  due = d;
  while (!ttIsWorking(due, holidays) && due < period.end) due = addDays(due, 1);
  return due;
}

/** Menu: Edit Task Types (shows the hidden tab). */
function showTaskTypesTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getTable(ss, "taskTypes");
  if (!sheet) return { success: false, message: "The Task Types tab was not found. Run Initialize / Upgrade first." };
  if (sheet.isSheetHidden()) sheet.showSheet();
  sheet.activate();
  try {
    ss.toast("Due: a weekday for weekly types, a day number or \"Last working day\" for monthly ones. Mandatory = audited (reminders). Created when: an event that creates the task, Days after it, Assign to (Person concerned, Role: X or names), Instructions. Hide the tab again when done.", "Task Types", 12);
  } catch (eToast) {}
  return { success: true };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    DEFAULT_TASK_TYPES: DEFAULT_TASK_TYPES,
    readTaskTypes: readTaskTypes,
    taskTypeNames: taskTypeNames,
    findTaskType: findTaskType,
    seedDefaultTaskTypes: seedDefaultTaskTypes,
    upgradeTaskTypeTriggers: upgradeTaskTypeTriggers,
    normalizeFrequency: normalizeFrequency,
    isoWeekKey: isoWeekKey,
    periodOf: periodOf,
    periodShift: periodShift,
    periodFromKey: periodFromKey,
    dueDateFor: dueDateFor,
    ttTruthy: ttTruthy,
    ttIsWorking: ttIsWorking,
    showTaskTypesTab: showTaskTypesTab
  };
}
