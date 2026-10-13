/**
 * Shared helpers (ADR 0049): dates, People chip names, table rows, HTML.
 * Every Apps Script file shares one global scope, so these names are used
 * directly everywhere; Node tests load them through test/mockGas.js.
 */

/**
 * "yyyy-mm-dd" for a Date, an ISO-like string ("2026-10-13…") or any text a
 * Date can parse; "" when empty or unreadable.
 *
 * @param {*} v
 * @return {string}
 */
function isoDate(v) {
  if (v === "" || v === null || v === undefined) return "";
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return "";
    return v.getFullYear() + "-" + ("0" + (v.getMonth() + 1)).slice(-2) + "-" + ("0" + v.getDate()).slice(-2);
  }
  var str = String(v).trim();
  var m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[0];
  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? "" : isoDate(parsed);
}

/**
 * The day of `v` at local midnight (a Date, an ISO-like string or parseable
 * text), or null when empty or unreadable.
 *
 * @param {*} v
 * @return {Date|null}
 */
function toDate(v) {
  if (v === "" || v === null || v === undefined) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
  var m = String(v).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  var d = new Date(String(v).trim());
  return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

/** A new Date `n` days after `d` (time of day kept). */
function addDays(d, n) {
  var x = new Date(d.getTime());
  x.setDate(x.getDate() + n);
  return x;
}

/** The ISO date `days` after the ISO date `iso`. */
function shiftIso(iso, days) {
  var d = toDate(iso);
  return d ? isoDate(addDays(d, days)) : "";
}

/** The text of a cell: a People chip's name, rich text, a date as ISO, or plain text. */
function cellName(val, rich) {
  return (typeof extractCellValue === "function") ? extractCellValue(val, rich) : String(val === null || val === undefined ? "" : val).trim();
}

/**
 * The rows of a table below its header: [{row, vals, rich}] for `numCols`
 * columns, up to the last row with content (at least `cap` rows looked at).
 */
function readTableRows(sheet, numCols, cap) {
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var range = sheet.getRange(2, 1, readRowLimit(sheet, cap) - 1, numCols);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  return vals.map(function(v, i) { return { row: i + 2, vals: v, rich: rich[i] || [] }; });
}

/** Text made safe inside HTML. */
function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    isoDate: isoDate,
    toDate: toDate,
    addDays: addDays,
    shiftIso: shiftIso,
    cellName: cellName,
    readTableRows: readTableRows,
    escapeHtml: escapeHtml
  };
}
