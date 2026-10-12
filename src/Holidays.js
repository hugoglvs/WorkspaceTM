/**
 * Public holidays (ADR 0043): a hidden "Holidays" tab (Date | Holiday) that every
 * engine treats like a weekend day. It is pre-filled with the French public
 * holidays of this year and next (each year once, so a deleted line stays
 * deleted); the team adds its own (bridges, office closures).
 * Menu: Project Management → Edit Holidays.
 */

var HOLIDAYS_SEEDED_PROP = "HOLIDAYS_SEEDED_YEARS";

function hdIso(d) {
  return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2);
}

function hdToIso(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return hdIso(v);
  var m = String(v || "").match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  if (!v) return "";
  var d = new Date(String(v));
  return isNaN(d.getTime()) ? "" : hdIso(d);
}

/** Easter Sunday (Gregorian calendar, anonymous algorithm). */
function easterSunday(year) {
  var a = year % 19, b = Math.floor(year / 100), c = year % 100;
  var d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25);
  var g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
  var i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7;
  var m = Math.floor((a + 11 * h + 22 * l) / 451);
  var month = Math.floor((h + l - 7 * m + 114) / 31);
  var day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}

/** The 11 French public holidays of a year, as [iso, name]. */
function frenchPublicHolidays(year) {
  var easter = easterSunday(year);
  var plus = function(n) { return hdIso(new Date(easter.getFullYear(), easter.getMonth(), easter.getDate() + n)); };
  var fixed = function(m, d) { return hdIso(new Date(year, m - 1, d)); };
  return [
    [fixed(1, 1), "New Year's Day"],
    [plus(1), "Easter Monday"],
    [fixed(5, 1), "Labour Day"],
    [fixed(5, 8), "Victory in Europe Day"],
    [plus(39), "Ascension Day"],
    [plus(50), "Whit Monday"],
    [fixed(7, 14), "Bastille Day"],
    [fixed(8, 15), "Assumption Day"],
    [fixed(11, 1), "All Saints' Day"],
    [fixed(11, 11), "Armistice Day"],
    [fixed(12, 25), "Christmas Day"]
  ].sort(function(x, y) { return x[0] < y[0] ? -1 : 1; });
}

function hdProps() {
  try {
    return (typeof PropertiesService !== "undefined") ? PropertiesService.getDocumentProperties() : null;
  } catch (e) {
    return null;
  }
}

/**
 * Adds the French public holidays of this year and next, each year only once.
 *
 * @return {{added: number, years: number[]}}
 */
function seedPublicHolidays(ss, asOf) {
  var res = { added: 0, years: [] };
  var sheet = getTable(ss, "holidays");
  if (!sheet) return res;
  var props = hdProps();
  var seeded = [];
  try { seeded = JSON.parse((props && props.getProperty(HOLIDAYS_SEEDED_PROP)) || "[]"); } catch (e) { seeded = []; }
  var year = (asOf ? new Date(hdToIso(asOf) + "T00:00:00") : new Date()).getFullYear();
  var existing = readHolidays(ss);
  var rows = [];
  [year, year + 1].forEach(function(y) {
    if (seeded.indexOf(y) !== -1) return;
    frenchPublicHolidays(y).forEach(function(h) { if (!existing[h[0]]) rows.push([h[0], h[1]]); });
    seeded.push(y);
    res.years.push(y);
  });
  if (rows.length) {
    var start = Math.max((typeof sheet.getLastRow === "function" ? sheet.getLastRow() : 1) || 1, 1) + 1;
    sheet.getRange(start, 1, rows.length, 2).setValues(rows);
    res.added = rows.length;
  }
  if (props && res.years.length) props.setProperty(HOLIDAYS_SEEDED_PROP, JSON.stringify(seeded));
  return res;
}

/**
 * Public holidays by ISO date.
 * @return {Object<string, string>} "2026-12-25" -> "Christmas Day"
 */
function readHolidays(ss) {
  var out = {};
  var sheet = (ss && typeof getTable === "function") ? getTable(ss, "holidays") : null;
  if (!sheet || sheet.getMaxRows() < 2) return out;
  var vals = sheet.getRange(2, 1, readRowLimit(sheet, 200) - 1, 2).getValues();
  for (var i = 0; i < vals.length; i++) {
    var iso = hdToIso(vals[i][0]);
    if (iso) out[iso] = String(vals[i][1] || "").trim() || "Public holiday";
  }
  return out;
}

/** Monday to Friday and not a public holiday. */
function isWorkingDay(date, holidays) {
  var dow = date.getDay();
  if (dow === 0 || dow === 6) return false;
  return !(holidays && holidays[hdIso(date)]);
}

/** Menu: Edit Holidays. Shows the hidden tab. */
function showHolidaysTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getTable(ss, "holidays");
  if (!sheet) return { success: false, message: "The Holidays tab was not found. Run Initialize / Upgrade first." };
  if (sheet.isSheetHidden()) sheet.showSheet();
  sheet.activate();
  try {
    ss.toast("Add or remove days (bridges, office closures), then hide the tab again (right-click → Hide sheet).", "Holidays", 8);
  } catch (eToast) {}
  return { success: true };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    easterSunday: easterSunday,
    frenchPublicHolidays: frenchPublicHolidays,
    seedPublicHolidays: seedPublicHolidays,
    readHolidays: readHolidays,
    isWorkingDay: isWorkingDay,
    showHolidaysTab: showHolidaysTab
  };
}
