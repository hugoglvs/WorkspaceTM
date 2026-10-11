/**
 * Test fixtures: seed tbl_presence_ranges from the legacy row shapes used
 * throughout the tests, so each test keeps its original data.
 */
var COLUMNS = ['range_id', 'member', 'range_type', 'start_date', 'start_period', 'end_date', 'end_period', 'day_of_week', 'location_am', 'location_pm', 'status'];

function appendRows(ss, rows) {
  var sheet = ss.getSheetByName('tbl_presence_ranges');
  var vals = sheet.getRange(2, 1, 200, 2).getValues();
  var first = 2;
  for (var i = vals.length - 1; i >= 0; i--) { if (vals[i][0] || vals[i][1]) { first = i + 3; break; } }
  sheet.getRange(first, 1, rows.length, COLUMNS.length).setValues(rows);
  return sheet;
}

/** rows: [baseline_id, member, start_date, end_date, day_of_week, location_am, location_pm] */
function seedBaselines(ss, rows) {
  return appendRows(ss, rows.map(function(r) {
    return [r[0], r[1], 'Baseline', r[2], '', r[3], '', r[4], r[5], r[6], 'Approved'];
  }));
}

/**
 * rows: 6 columns [leave_id, member, start_date, end_date, leave_type, status]
 *    or 8 columns [leave_id, member, start_date, start_period, end_date, end_period, leave_type, status]
 */
function seedLeaves(ss, rows) {
  return appendRows(ss, rows.map(function(r) {
    var full = r.length >= 8;
    var type = full ? r[6] : r[4];
    return [r[0], r[1], 'Leave', r[2], full ? r[3] : '', full ? r[4] : r[3], full ? r[5] : '', 'All', type, type, full ? r[7] : r[5]];
  }));
}

module.exports = { seedBaselines: seedBaselines, seedLeaves: seedLeaves, PRESENCE_RANGE_COLUMNS: COLUMNS };
