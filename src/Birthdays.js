/**
 * Birthdays (ADR 0042).
 *
 * Members has a "Birthday" column holding day and month only ("14 Mar"): no
 * year, so no ages in a sheet everyone can read. On the working day before a
 * birthday (Friday for a Saturday, Sunday or Monday birthday), every other Active
 * Member gets an email with an "Add to my Google Calendar" link (a yearly
 * all-day event). The script also keeps a shared "Team Birthdays" calendar that
 * Active Members can subscribe to once.
 */

var BD_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
var BD_MONTHS_LONG = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
var BD_WEEKDAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
var BD_PROP_CALENDAR = "BIRTHDAY_CALENDAR_ID";
var BD_PROP_EVENTS = "BIRTHDAY_EVENTS";
var BD_PROP_SHARED = "BIRTHDAY_CALENDAR_SHARED";
var BD_PROP_SENT_ON = "BIRTHDAY_REMINDED_ON";

function bdProps() {
  try {
    return (typeof PropertiesService !== "undefined") ? PropertiesService.getDocumentProperties() : null;
  } catch (e) {
    return null;
  }
}

function bdJson(key, fallback) {
  var props = bdProps();
  try { return JSON.parse((props && props.getProperty(key)) || "null") || fallback; } catch (e) { return fallback; }
}

function bdSetJson(key, value) {
  var props = bdProps();
  if (props) props.setProperty(key, JSON.stringify(value));
}

function bdIso(d) {
  return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2);
}

function bdMidnight(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var m = String(v || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  var d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function bdMonthIndex(word) {
  var w = String(word || "").trim().toLowerCase().replace(/\.$/, "");
  if (w.length < 3) return -1;
  for (var i = 0; i < 12; i++) {
    if (BD_MONTHS_LONG[i].toLowerCase().indexOf(w) === 0) return i;
  }
  return -1;
}

function bdValid(month, day) {
  var max = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1];
  return month >= 1 && month <= 12 && day >= 1 && day <= max;
}

/**
 * Reads a birthday typed in many ways: "14 Mar", "14 March", "March 14",
 * "14/03", "14-03", "14.03" (day first), "2001-03-14", or a date cell. The year
 * is dropped.
 *
 * @param {*} value
 * @return {{month: number, day: number}|null}
 */
function parseBirthday(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date && !isNaN(value.getTime())) return { month: value.getMonth() + 1, day: value.getDate() };
  var s = String(value).trim();
  var m;
  var out = null;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) out = { month: +m[2], day: +m[3] };
  else if ((m = s.match(/^(\d{1,2})\s*[\/.\-]\s*(\d{1,2})(?:\s*[\/.\-]\s*\d{2,4})?$/))) out = { month: +m[2], day: +m[1] };
  else if ((m = s.match(/^(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-zéû.]+)(?:\s+\d{4})?$/))) out = { month: bdMonthIndex(m[2]) + 1, day: +m[1] };
  else if ((m = s.match(/^([A-Za-zéû.]+)\s+(\d{1,2})(?:st|nd|rd|th)?,?(?:\s+\d{4})?$/))) out = { month: bdMonthIndex(m[1]) + 1, day: +m[2] };
  return out && bdValid(out.month, out.day) ? out : null;
}

/** "14 Mar" */
function formatBirthday(b) {
  return b ? b.day + " " + BD_MONTHS[b.month - 1] : "";
}

function bdIsLeap(y) {
  return (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
}

/** The birthday in a given year (29 Feb falls on 28 Feb in other years). */
function bdInYear(b, year) {
  if (b.month === 2 && b.day === 29 && !bdIsLeap(year)) return new Date(year, 1, 28);
  return new Date(year, b.month - 1, b.day);
}

/** Next birthday on or after `from`. */
function nextBirthdayDate(b, from) {
  var f = bdMidnight(from);
  var d = bdInYear(b, f.getFullYear());
  return d < f ? bdInYear(b, f.getFullYear() + 1) : d;
}

function bdColumn() {
  var cols = (typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC.members) ? SCHEMA_SPEC.members.columns : [];
  return cols.indexOf("birthday") + 1;
}

/** Members with status, emails and parsed birthday. */
function readBirthdayMembers(ss) {
  var sheet = getTable(ss, "members");
  var col = bdColumn();
  if (!sheet || sheet.getMaxRows() < 2 || col < 1) return [];
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 300) - 1, col);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var r = rich[i] || [];
    var name = (typeof extractCellValue === "function") ? extractCellValue(vals[i][0], r[0]) : String(vals[i][0] || "").trim();
    if (!name) continue;
    var google = (typeof extractEmailFromChip === "function") ? String(extractEmailFromChip(vals[i][0], r[0]) || "").trim() : "";
    var work = String(vals[i][1] || "").trim();
    out.push({
      row: i + 2, name: name, key: name.toLowerCase(),
      status: String(vals[i][5] || "").trim().toLowerCase(),
      email: work || google, googleEmail: google || work,
      birthday: parseBirthday(vals[i][col - 1])
    });
  }
  return out;
}

/**
 * onEdit (Members, Birthday column): rewrites what was typed as "14 Mar", or
 * explains the expected format.
 *
 * @return {string|null} The stored text, "invalid", or null when not concerned
 */
function normalizeBirthdayEdit(e) {
  if (!e || !e.range || e.range.getRow() < 2 || e.range.getColumn() !== bdColumn()) return null;
  var cell = e.range;
  var raw = cell.getValue();
  if (raw === "" || raw === null) return null;
  var b = parseBirthday(raw);
  var ss = e.source || SpreadsheetApp.getActiveSpreadsheet();
  if (!b) {
    try { ss.toast("Write the birthday as day and month, e.g. 14 Mar or 14/03. No year needed.", "Birthday", 8); } catch (eT) {}
    return "invalid";
  }
  var text = formatBirthday(b);
  // The Birthday column is a typed TEXT column, so the text stays text (and
  // Sheets refuses number formats there)
  if (raw !== text) cell.setValue(text);
  return text;
}

/**
 * Days announced on `asOf`: from the next day up to the next working day. On a
 * Friday that is Saturday, Sunday and Monday; public holidays are skipped like
 * weekends (ADR 0043); on a non-working day, nothing.
 */
function birthdayAnnouncementDays(asOf, holidays) {
  var d = bdMidnight(asOf);
  var working = function(x) {
    return x.getDay() !== 0 && x.getDay() !== 6 && !(holidays && holidays[bdIso(x)]);
  };
  if (!working(d)) return [];
  var days = [];
  var next = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
  while (true) {
    days.push(next);
    if (working(next)) break;
    next = new Date(next.getFullYear(), next.getMonth(), next.getDate() + 1);
  }
  return days;
}

/** Google Calendar link that opens a yearly all-day event, ready to save. */
function birthdayCalendarLink(name, date) {
  var end = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);
  var compact = function(d) { return bdIso(d).replace(/-/g, ""); };
  return "https://calendar.google.com/calendar/render?action=TEMPLATE" +
    "&text=" + encodeURIComponent("🎂 " + name + "'s birthday") +
    "&dates=" + compact(date) + "/" + compact(end) +
    "&recur=" + encodeURIComponent("RRULE:FREQ=YEARLY") +
    "&crm=AVAILABLE" +
    "&details=" + encodeURIComponent("Added from the team's WorkspacePM reminder.");
}

function birthdaySubscribeLink() {
  var props = bdProps();
  var id = props ? props.getProperty(BD_PROP_CALENDAR) : "";
  return id ? "https://calendar.google.com/calendar/r?cid=" + encodeURIComponent(id) : "";
}

function bdHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * Emails every Active Member about the birthdays of the coming days (their own
 * excepted). Runs once per day at most.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Date|string} [asOf]
 * @return {{celebrants: string[], sent: string[], skipped: (string|undefined)}}
 */
function sendBirthdayReminders(targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var today = bdMidnight(asOf || new Date());
  var res = { celebrants: [], sent: [] };
  var props = bdProps();
  if (props && props.getProperty(BD_PROP_SENT_ON) === bdIso(today)) { res.skipped = "already_sent_today"; return res; }

  var days = birthdayAnnouncementDays(today, (typeof readHolidays === "function") ? readHolidays(ss) : {});
  if (!days.length) return res;
  var members = readBirthdayMembers(ss).filter(function(m) { return m.status === "active"; });

  var celebrants = [];
  members.forEach(function(m) {
    if (!m.birthday) return;
    days.forEach(function(d) {
      var on = bdInYear(m.birthday, d.getFullYear());
      if (bdIso(on) === bdIso(d)) celebrants.push({ key: m.key, name: m.name, date: d, birthday: m.birthday });
    });
  });
  res.celebrants = celebrants.map(function(c) { return c.name; });
  if (!celebrants.length) return res;

  var subscribe = birthdaySubscribeLink();
  members.forEach(function(r) {
    if (!r.email) return;
    var list = celebrants.filter(function(c) { return c.key !== r.key; });
    if (!list.length) return;
    var when = function(c) {
      var label = BD_WEEKDAYS[c.date.getDay()] + " " + c.date.getDate() + " " + BD_MONTHS_LONG[c.date.getMonth()];
      return c.date.getTime() === days[0].getTime() ? "tomorrow, " + label : "on " + label;
    };
    var subject = list.length === 1
      ? "🎂 " + list[0].name + "'s birthday is " + when(list[0]).replace(/, .*/, "")
      : "🎂 " + list.length + " birthdays coming up";
    var text = list.map(function(c) {
      return c.name + "'s birthday is " + when(c) + ".\nAdd it to your Google Calendar (every year): " + birthdayCalendarLink(c.name, c.date);
    }).join("\n\n");
    var html = list.map(function(c) {
      return "<p>🎂 <b>" + bdHtml(c.name) + "</b>'s birthday is " + bdHtml(when(c)) + ".<br>" +
        "<a href=\"" + birthdayCalendarLink(c.name, c.date) + "\">Add to my Google Calendar (every year)</a></p>";
    }).join("");
    if (subscribe) {
      text += "\n\nOr subscribe once to all team birthdays: " + subscribe;
      html += "<p style=\"color:#5f6368;font-size:12px\">Or <a href=\"" + subscribe + "\">subscribe once to all team birthdays</a>; it stays up to date.</p>";
    }
    MailApp.sendEmail({ to: r.email, subject: subject, body: text, htmlBody: html });
    res.sent.push(r.name);
  });

  if (props) props.setProperty(BD_PROP_SENT_ON, bdIso(today));
  Logger.log("[birthdayReminderJob] " + res.celebrants.join(", ") + " → " + res.sent.length + " email(s)");
  return res;
}

/** The shared Team Birthdays calendar, created on first use. */
function getBirthdayCalendar(ss) {
  var props = bdProps();
  if (!props || typeof CalendarApp === "undefined") return null;
  var id = props.getProperty(BD_PROP_CALENDAR);
  var cal = null;
  if (id) { try { cal = CalendarApp.getCalendarById(id); } catch (e) { cal = null; } }
  if (!cal) {
    cal = CalendarApp.createCalendar("Team Birthdays" + (ss && typeof ss.getName === "function" ? " – " + ss.getName() : ""));
    props.setProperty(BD_PROP_CALENDAR, cal.getId());
    bdSetJson(BD_PROP_EVENTS, {});
    bdSetJson(BD_PROP_SHARED, []);
  }
  return cal;
}

/**
 * Keeps the Team Birthdays calendar in step with Members: one yearly all-day
 * event per Active Member with a birthday, shared read-only with Active Members.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Date|string} [asOf]
 * @return {{created: number, removed: number, shared: number, unshared: number}}
 */
function syncBirthdayCalendar(targetSpreadsheet, asOf) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var res = { created: 0, removed: 0, shared: 0, unshared: 0 };
  var members = readBirthdayMembers(ss).filter(function(m) { return m.status === "active"; });
  var withBirthday = members.filter(function(m) { return m.birthday; });
  var known = bdJson(BD_PROP_EVENTS, {});
  if (!withBirthday.length && !Object.keys(known).length) return res;

  var cal = getBirthdayCalendar(ss);
  if (!cal) return res;
  known = bdJson(BD_PROP_EVENTS, {});
  var today = bdMidnight(asOf || new Date());

  var wanted = {};
  withBirthday.forEach(function(m) { wanted[m.key] = m.name + "|" + formatBirthday(m.birthday); });

  Object.keys(known).forEach(function(key) {
    if (wanted[key] === known[key].label) return;
    try {
      var series = cal.getEventSeriesById(known[key].id);
      if (series) series.deleteEventSeries();
    } catch (eDel) { Logger.log("Birthday event removal note: " + eDel); }
    delete known[key];
    res.removed++;
  });

  withBirthday.forEach(function(m) {
    if (known[m.key]) return;
    var b = m.birthday;
    var year = today.getFullYear();
    if (b.month === 2 && b.day === 29) { while (!bdIsLeap(year)) year++; }
    var series = cal.createAllDayEventSeries("🎂 " + m.name + "'s birthday", new Date(year, b.month - 1, b.day),
      CalendarApp.newRecurrence().addYearlyRule());
    known[m.key] = { id: series.getId(), label: wanted[m.key] };
    res.created++;
  });
  bdSetJson(BD_PROP_EVENTS, known);

  // Read access for Active Members (their Google account), without share emails
  if (typeof Calendar !== "undefined" && Calendar.Acl) {
    var shared = bdJson(BD_PROP_SHARED, []);
    var emails = members.map(function(m) { return m.googleEmail.toLowerCase(); }).filter(Boolean);
    var calId = cal.getId();
    emails.forEach(function(em) {
      if (shared.indexOf(em) !== -1) return;
      try {
        Calendar.Acl.insert({ role: "reader", scope: { type: "user", value: em } }, calId, { sendNotifications: false });
        shared.push(em);
        res.shared++;
      } catch (eAcl) { Logger.log("Birthday calendar share note (" + em + "): " + eAcl); }
    });
    shared = shared.filter(function(em) {
      if (emails.indexOf(em) !== -1) return true;
      try { Calendar.Acl.remove(calId, "user:" + em); res.unshared++; } catch (eRm) {}
      return false;
    });
    bdSetJson(BD_PROP_SHARED, shared);
  }
  return res;
}

/** Time trigger entry point (every day, 9 am; does nothing at weekends). */
function birthdayReminderJob() {
  return dailyRemindersJob();
}

/**
 * The 9 am job: birthday reminders and the Team Birthdays calendar (ADR 0042),
 * then task reminders (ADR 0045).
 */
function dailyRemindersJob() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = {};
  try { syncBirthdayCalendar(ss); } catch (eCal) { Logger.log("Birthday calendar sync note: " + eCal); }
  try { res.birthdays = sendBirthdayReminders(ss); } catch (eB) { Logger.log("Birthday reminders note: " + eB); }
  try { if (typeof sendTaskReminders === "function") res.tasks = sendTaskReminders(ss); } catch (eT) { Logger.log("Task reminders note: " + eT); }
  return res;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    parseBirthday: parseBirthday,
    formatBirthday: formatBirthday,
    nextBirthdayDate: nextBirthdayDate,
    readBirthdayMembers: readBirthdayMembers,
    normalizeBirthdayEdit: normalizeBirthdayEdit,
    birthdayAnnouncementDays: birthdayAnnouncementDays,
    birthdayCalendarLink: birthdayCalendarLink,
    sendBirthdayReminders: sendBirthdayReminders,
    syncBirthdayCalendar: syncBirthdayCalendar,
    birthdayReminderJob: birthdayReminderJob,
    dailyRemindersJob: dailyRemindersJob
  };
}
