/**
 * Task types and task periods (ADR 0045).
 *
 * A Task Template (Quiz, Onboarding, Time & expenses…; key "taskTypes", tab
 * "Task Templates", ADR 0052) says how often a task comes back and when it is
 * due. Type | Frequency (Weekly, Twice a month, Monthly, One-off) | Due |
 * Mandatory | Active | Created when | Days after | Assign to | Instructions | Remind.
 * Due: a weekday for weekly types ("Friday"); for monthly types a day number
 * ("20"), "Last working day" or "First working day"; for twice-a-month types one
 * value per half month ("15; Last day"). A due date that falls on a weekend or
 * public holiday moves to the working day before (after, for "First working
 * day"). Mandatory types are the audited ones (manager digest, overdue list).
 * Created when / Days after (ADR 0051): a type created by an event
 * (src/TaskTriggers.js), e.g. Onboarding when a new member arrives.
 * Assign to / Instructions: who gets it by default and its description, used
 * by triggers and pre-filled in the Assign Tasks sidebar.
 * Remind: working days before the due date when the person is emailed
 * ("3; 1"; 0 = the day itself, at 9 am); blank = 3.
 */

var ONBOARDING_INSTRUCTIONS = "Welcome! Please enter your Accenture email address and phone number in the Members tab (columns Accenture email and Phone).";
var TT_COLUMNS = 10;
var FREQUENCIES = ["Weekly", "Twice a month", "Monthly", "One-off"];
var DEFAULT_TASK_TYPES = [
  ["Onboarding", "One-off", "", false, true, "New member", 7, "Person concerned", ONBOARDING_INSTRUCTIONS, ""],
  ["Quiz", "Weekly", "Friday", false, true, "", "", "", "", ""],
  ["Team presentation", "Weekly", "Thursday", false, true, "", "", "", "", ""],
  ["Peer review", "Weekly", "Friday", true, true, "", "", "", "", "1"],
  ["SoW", "Monthly", "Last working day", true, true, "", "", "", "", ""],
  ["PRM", "Monthly", "Last working day", true, true, "", "", "", "", ""],
  ["Forecast", "Monthly", "Last working day", true, true, "", "", "", "", ""],
  ["Time & expenses", "Twice a month", "15; Last day", true, true, "", "", "", "", "3; 1"]
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
  if (/^(twice|semi|bi-?month|2x)/.test(s)) return "Twice a month";
  if (s.indexOf("month") === 0) return "Monthly";
  if (s.indexOf("one") === 0 || s === "once") return "One-off";
  return "Weekly";
}

/** Frequency of a type's periods: One-off types are planned by week. */
function periodFrequency(f) {
  var n = normalizeFrequency(f);
  return n === "One-off" ? "Weekly" : n;
}

/** Remind text ("3; 1; 0") as working-day counts; blank = [3]. */
function parseRemindDays(text) {
  var out = [];
  String(text === undefined || text === null ? "" : text).split(/[;,\s]+/).forEach(function(tok) {
    var n = parseInt(tok, 10);
    if (!isNaN(n) && n >= 0 && out.indexOf(n) === -1) out.push(n);
  });
  return out.length ? out : [3];
}

function ttFromRow(v) {
  var activeCell = v[4];
  var days = parseInt(v[6], 10);
  return {
    name: String(v[0] || "").trim(), frequency: normalizeFrequency(v[1]), due: String(v[2] || "").trim(),
    mandatory: ttTruthy(v[3]),
    active: (activeCell === "" || activeCell === null || activeCell === undefined) ? true : ttTruthy(activeCell),
    createdWhen: String(v[5] || "").trim(), daysAfter: isNaN(days) ? 0 : days,
    assignTo: String(v[7] || "").trim(), instructions: String(v[8] || "").trim(),
    remind: parseRemindDays(v[9]), remindText: String(v[9] === undefined || v[9] === null ? "" : v[9]).trim()
  };
}

/**
 * Every task type, in tab order. Falls back to the defaults when the tab is
 * missing or empty.
 * @return {Array<{name: string, frequency: string, due: string, mandatory: boolean, active: boolean,
 *   createdWhen: string, daysAfter: number, assignTo: string, instructions: string, remind: number[],
 *   row: (number|undefined)}>}
 */
function readTaskTypes(ss) {
  var sheet = (ss && typeof getTable === "function") ? getTable(ss, "taskTypes") : null;
  var list = [];
  if (sheet && sheet.getMaxRows() > 1) {
    var vals = sheet.getRange(2, 1, readRowLimit(sheet, 100) - 1, TT_COLUMNS).getValues();
    for (var r = 0; r < vals.length; r++) {
      if (!String(vals[r][0] || "").trim()) continue;
      var tt = ttFromRow(vals[r]);
      tt.row = r + 2;
      list.push(tt);
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

/** Writes the default task types into an empty Task Templates tab. */
function seedDefaultTaskTypes(ss) {
  var sheet = getTable(ss, "taskTypes");
  if (!sheet) return 0;
  var existing = sheet.getMaxRows() > 1 ? sheet.getRange(2, 1, Math.min(sheet.getMaxRows(), 50) - 1, 1).getValues() : [];
  if (existing.some(function(r) { return String(r[0] || "").trim(); })) {
    upgradeTaskTypeTriggers(ss, sheet);
    upgradeTaskTemplateSchedules(ss, sheet);
    return 0;
  }
  sheet.getRange(2, 1, DEFAULT_TASK_TYPES.length, TT_COLUMNS).setValues(DEFAULT_TASK_TYPES);
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

/**
 * Once per sheet (ADR 0052): Time & expenses and Peer review still on their
 * former default schedule (Monthly, Last working day) take the new defaults
 * (twice a month on the 15th and last day; weekly on Friday), with their
 * reminder days. Rows changed by the team are left alone.
 */
function upgradeTaskTemplateSchedules(ss, sheet) {
  var props = null;
  try { props = (typeof PropertiesService !== "undefined") ? PropertiesService.getDocumentProperties() : null; } catch (eP) {}
  if (!props || props.getProperty("TASK_TEMPLATES_0052")) return 0;
  var n = readRowLimit(sheet, 100) - 1;
  var vals = n > 0 ? sheet.getRange(2, 1, n, TT_COLUMNS).getValues() : [];
  var updated = 0;
  vals.forEach(function(v, i) {
    var name = String(v[0] || "").trim().toLowerCase();
    if (name !== "time & expenses" && name !== "peer review") return;
    if (normalizeFrequency(v[1]) !== "Monthly" || String(v[2] || "").trim().toLowerCase() !== "last working day") return;
    var d = DEFAULT_TASK_TYPES.filter(function(x) { return x[0].toLowerCase() === name; })[0];
    sheet.getRange(i + 2, 2, 1, 2).setValues([[d[1], d[2]]]);
    if (!String(v[9] === undefined || v[9] === null ? "" : v[9]).trim() && d[9]) sheet.getRange(i + 2, 10).setValue(d[9]);
    updated++;
  });
  props.setProperty("TASK_TEMPLATES_0052", "1");
  return updated;
}

/**
 * Changes a template's schedule (Assign Tasks sidebar, ADR 0052): it applies to
 * everyone; generateRecurringTasks then moves the coming tasks.
 *
 * @return {boolean} whether something changed
 */
function setTaskTypeSchedule(ss, name, frequency, due, remind) {
  var type = findTaskType(ss, name);
  var sheet = getTable(ss, "taskTypes");
  if (!type || !type.row || !sheet) return false;
  var freq = normalizeFrequency(frequency || type.frequency);
  var newDue = due === undefined || due === null ? type.due : String(due).trim();
  var newRemind = remind === undefined || remind === null ? type.remindText : String(remind).trim();
  if (freq === type.frequency && newDue === type.due && newRemind === type.remindText) return false;
  sheet.getRange(type.row, 2, 1, 2).setValues([[freq, newDue]]);
  sheet.getRange(type.row, 10).setValue(newRemind);
  return true;
}

// ---------------------------------------------------------------------------
// Periods: a week ("2026-W42") or a month ("2026-10")
// ---------------------------------------------------------------------------

function TT_MONTH_ENDS(y, m) { return new Date(y, m + 1, 0).getDate(); }

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
  var f = periodFrequency(frequency);
  if (f === "Twice a month") {
    var y = d.getFullYear(), mo = d.getMonth(), last = TT_MONTH_ENDS(y, mo);
    var first = d.getDate() <= 15;
    var hs = new Date(y, mo, first ? 1 : 16);
    var he = new Date(y, mo, first ? 15 : last);
    return { key: y + "-" + ("0" + (mo + 1)).slice(-2) + (first ? "-H1" : "-H2"), start: hs, end: he,
      label: (first ? "1–15 " : "16–" + last + " ") + TT_MONTHS[mo] + " " + y, frequency: "Twice a month", half: first ? 1 : 2 };
  }
  if (f === "Monthly") {
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
  var f = periodFrequency(frequency);
  if (f === "Twice a month") {
    var idx = d.getFullYear() * 24 + d.getMonth() * 2 + (d.getDate() <= 15 ? 0 : 1) + n;
    var y = Math.floor(idx / 24), rest = idx - y * 24;
    return periodOf(f, new Date(y, Math.floor(rest / 2), rest % 2 ? 16 : 1));
  }
  if (f === "Monthly") return periodOf("Monthly", new Date(d.getFullYear(), d.getMonth() + n, 1));
  return periodOf("Weekly", addDays(ttMonday(d), 7 * n));
}

/** Period from its key ("2026-W42" or "2026-10"). */
function periodFromKey(key) {
  var h = String(key || "").match(/^(\d{4})-(\d{2})-H([12])$/);
  if (h) return periodOf("Twice a month", new Date(+h[1], +h[2] - 1, h[3] === "1" ? 1 : 16));
  var m = String(key || "").match(/^(\d{4})-(\d{2})$/);
  if (m) return periodOf("Monthly", new Date(+m[1], +m[2] - 1, 1));
  var w = String(key || "").match(/^(\d{4})-W(\d{2})$/);
  if (!w) return null;
  // Monday of ISO week 1 is the Monday of the week holding 4 January
  var week1 = ttMonday(new Date(+w[1], 0, 4));
  return periodOf("Weekly", addDays(week1, 7 * (+w[2] - 1)));
}

/** Frequency a period key belongs to ("2026-W42" Weekly, "2026-10" Monthly, "2026-10-H1" Twice a month). */
function periodKeyFrequency(key) {
  var k = String(key || "");
  if (/^\d{4}-\d{2}-H[12]$/.test(k)) return "Twice a month";
  if (/^\d{4}-\d{2}$/.test(k)) return "Monthly";
  if (/^\d{4}-W\d{2}$/.test(k)) return "Weekly";
  return "";
}

/** Periods planned ahead by the recurring engine (ADR 0045, 0052). */
function periodsAhead(frequency) {
  var f = periodFrequency(frequency);
  return f === "Monthly" ? 3 : (f === "Twice a month" ? 6 : 8);
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
  var f = periodFrequency(type && type.frequency);
  if (f === "Twice a month") {
    var toks = rule.split(/[;,]/).map(function(x) { return x.trim(); });
    rule = toks[(period.half || (period.start.getDate() > 1 ? 2 : 1)) - 1] || ((period.start.getDate() > 1) ? "last" : "15");
    if (/^\d+$/.test(rule)) {
      var day = Math.min(Math.max(+rule, period.start.getDate()), period.end.getDate());
      rule = String(day);
    }
  }
  if (f === "Monthly" || f === "Twice a month") {
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

/** Menu: Edit Task Templates (shows the hidden tab). */
function showTaskTypesTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getTable(ss, "taskTypes");
  if (!sheet) return { success: false, message: "The Task Templates tab was not found. Run Initialize / Upgrade first." };
  if (sheet.isSheetHidden()) sheet.showSheet();
  sheet.activate();
  try {
    ss.toast("Due: a weekday (weekly), a day number or \"Last working day\" (monthly), \"15; Last day\" (twice a month). Mandatory = in the manager digest. Assign to: Person concerned, Role: X or names. Remind: working days before (\"3; 1\"). Created when: an event that creates the task. Hide the tab again when done.", "Task Templates", 12);
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
    upgradeTaskTemplateSchedules: upgradeTaskTemplateSchedules,
    setTaskTypeSchedule: setTaskTypeSchedule,
    FREQUENCIES: FREQUENCIES,
    normalizeFrequency: normalizeFrequency,
    periodFrequency: periodFrequency,
    parseRemindDays: parseRemindDays,
    periodKeyFrequency: periodKeyFrequency,
    periodsAhead: periodsAhead,
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
