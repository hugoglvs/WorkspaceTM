/**
 * Project Management Workspace - View Builder
 * Ticket 04: Weekly Matrix Planning View (view_weekly_matrix)
 *
 * Renders team presence deployment into the management weekly matrix format:
 * - Header Row 1: Merged Member People Chips across paired AM/PM columns.
 * - Header Row 2: "AM" and "PM" subheaders.
 * - Rows: All consecutive days (Monday through Sunday) with merged ISO Week Numbers in Col A.
 * - Weekend days: Dark gray background (#5F6368) row styling.
 * - Past weeks: Hidden by default (only last 4 weeks visible) via hideRows.
 * - Live two-way synchronization: In-sheet presence edits immediately sync to tbl_presence_log and tbl_members.today_presence.
 * - Color-coded conditional formatting (Office, Remote, Leaves) and 12px "Google Sans" typography.
 */

/**
 * Calculates the ISO 8601 week number for a given Date.
 *
 * @param {Date} date
 * @return {number} Week number (1-53)
 */
function getIsoWeekNumber(date) {
  var d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  var dayNum = d.getUTCDay() || 7; // Sunday is 7 in ISO
  d.setUTCDate(d.getUTCDate() + 4 - dayNum); // Nearest Thursday
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  var weekNo = Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
  return weekNo;
}

/**
 * Formats ISO week number as "W01", "W02", etc.
 *
 * @param {Date} date
 * @return {string}
 */
function formatIsoWeekString(date) {
  var w = getIsoWeekNumber(date);
  return "W" + (w < 10 ? "0" + w : w);
}

/**
 * Formats a Date object as ISO date string "YYYY-MM-DD".
 *
 * @param {Date} date
 * @return {string}
 */
function formatMatrixDateIso(date) {
  var y = date.getFullYear();
  var m = String(date.getMonth() + 1).padStart(2, "0");
  var d = String(date.getDate()).padStart(2, "0");
  return y + "-" + m + "-" + d;
}

/**
 * Formats a Date object into standard datetime string: YYYY-MM-DD HH:mm:ss
 *
 * @param {Date} [d]
 * @return {string}
 */
function formatMatrixDateTime(d) {
  var date = d || new Date();
  var yr = date.getFullYear();
  var mo = ("0" + (date.getMonth() + 1)).slice(-2);
  var da = ("0" + date.getDate()).slice(-2);
  var hr = ("0" + date.getHours()).slice(-2);
  var mi = ("0" + date.getMinutes()).slice(-2);
  var se = ("0" + date.getSeconds()).slice(-2);
  return yr + "-" + mo + "-" + da + " " + hr + ":" + mi + ":" + se;
}

/**
 * Parses a date value (string or Date) into a midnight UTC/local Date object.
 *
 * @param {string|Date} val
 * @return {Date|null}
 */
function parseMatrixDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    var d = new Date(val.getTime());
    d.setHours(0, 0, 0, 0);
    return d;
  }
  var s = String(val).trim();
  var match = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    return new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10), 0, 0, 0, 0);
  }
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    parsed.setHours(0, 0, 0, 0);
    return parsed;
  }
  return null;
}

/**
 * Returns array of all consecutive Date objects between start and end (inclusive),
 * including both weekdays and weekends.
 *
 * @param {Date} startDate
 * @param {Date} endDate
 * @return {Array<Date>}
 */
function getAllMatrixDays(startDate, endDate) {
  var days = [];
  var curr = new Date(startDate.getTime());
  curr.setHours(0, 0, 0, 0);

  var end = new Date(endDate.getTime());
  end.setHours(0, 0, 0, 0);

  while (curr <= end) {
    days.push(new Date(curr.getTime()));
    curr.setDate(curr.getDate() + 1);
  }
  return days;
}

/**
 * Colour-codes the matrix cells by Work Location type (Locations tab, ADR 0040):
 * Office purple, Client site blue, Remote green, leaves red, Off grey (PRESENCE_STYLES).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {GoogleAppsScript.Spreadsheet.Range} targetRange
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 */
function applyMatrixConditionalFormatting(sheet, targetRange, ss) {
  if (!sheet || !targetRange) return;
  var entries = presenceStyleEntries(ss || sheet.getParent());
  var rules = entries.map(function(en) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(en.text)
      .setBackground(en.style[0])
      .setFontColor(en.style[1])
      .setRanges([targetRange])
      .build();
  });
  sheet.setConditionalFormatRules(rules);
}

/**
 * Builds or refreshes the `view_weekly_matrix` sheet tab.
 * - Header Row 1: Merged Member People Chips above AM and PM columns.
 * - Header Row 2: "AM" and "PM" subheaders.
 * - Rows: All calendar days (Monday-Sunday), with merged Week Numbers in Col A.
 * - Weekends: Dark gray styling.
 * - Past weeks: Hidden by default (only last 4 weeks visible).
 * - Interactive validation: In-cell dropdowns for rapid presence declaration.
 *
 * @param {string} [startDateStr] Optional start date YYYY-MM-DD
 * @param {string} [endDateStr] Optional end date YYYY-MM-DD
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function renderWeeklyMatrixView(startDateStr, endDateStr, targetSpreadsheet, asOfDateParam) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet available" };

  var membersSheet = getTable(ss, "members");
  var logSheet = getTable(ss, "presenceLog");

  if (!membersSheet) {
    return { success: false, message: "Sheet 'Members' not found" };
  }

  // 1. Extract active members
  var activeMembers = [];
  var mMax = readRowLimit(membersSheet, 300);
  if (mMax > 1) {
    var mRange = membersSheet.getRange(2, 1, mMax - 1, Math.max(membersSheet.getLastColumn(), 6));
    var mVals = mRange.getValues();
    var mRich = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : [];

    for (var i = 0; i < mVals.length; i++) {
      var rawVal = mVals[i][0];
      var rawR = (mRich && mRich[i]) ? mRich[i][0] : null;
      var memberName = (typeof extractCellValue === "function") ? extractCellValue(rawVal, rawR) : String(rawVal).trim();
      if (!memberName) continue;

      var tenureStatus = String(mVals[i][5] || "").trim().toLowerCase();
      if (tenureStatus === "archived" || tenureStatus === "terminated") {
        continue;
      }

      var memberKey = memberName.toLowerCase();
      activeMembers.push({
        name: memberName,
        key: memberKey,
        sourceRow: i + 2,
        rawVal: rawVal,
        departed: tenureStatus === "departed" // columns hidden, not removed (ADR 0040)
      });
    }
  }

  if (activeMembers.length === 0) {
    return { success: false, message: "No active members found in Members" };
  }

  // 2. Read existing daily records from tbl_presence_log
  var presenceData = {};
  var minLoggedDate = null;
  var maxLoggedDate = null;

  if (logSheet) {
    var lMax = readRowLimit(logSheet, 2000);
    if (lMax > 1) {
      var lRange = logSheet.getRange(2, 1, lMax - 1, 5);
      var lVals = lRange.getValues();
      var lRich = (typeof lRange.getRichTextValues === "function") ? lRange.getRichTextValues() : [];

      for (var l = 0; l < lVals.length; l++) {
        var dVal = lVals[l][1];
        if (!dVal) continue;
        var dObj = parseMatrixDate(dVal);
        if (!dObj) continue;

        var dIso = formatMatrixDateIso(dObj);
        var mVal = lVals[l][2];
        var mR = (lRich && lRich[l]) ? lRich[l][2] : null;
        var mName = (typeof extractCellValue === "function") ? extractCellValue(mVal, mR) : String(mVal).trim();
        var mKey = mName.toLowerCase();

        var locAm = String(lVals[l][3] || "").trim();
        var locPm = String(lVals[l][4] || "").trim();

        var entryKey = dIso + "::" + mKey;
        presenceData[entryKey] = { am: locAm, pm: locPm };

        if (!minLoggedDate || dObj < minLoggedDate) minLoggedDate = new Date(dObj.getTime());
        if (!maxLoggedDate || dObj > maxLoggedDate) maxLoggedDate = new Date(dObj.getTime());
      }
    }
  }

  // 3. Determine planning window (aligned to full weeks: Monday to Sunday)
  var startD = parseMatrixDate(startDateStr);
  var endD = parseMatrixDate(endDateStr);

  var today = asOfDateParam ? (parseMatrixDate(asOfDateParam) || new Date()) : new Date();
  today.setHours(0, 0, 0, 0);

  if (!startD) {
    if (minLoggedDate) {
      startD = new Date(minLoggedDate.getTime());
      // Expand backward to Monday of that week
      var sDay = startD.getDay(); // 0 is Sun, 1 is Mon
      var diffBack = (sDay === 0 ? 6 : sDay - 1);
      startD.setDate(startD.getDate() - diffBack);
    } else {
      // Default: 4 weeks ago starting on Monday
      startD = new Date(today.getTime());
      var curDay = startD.getDay();
      var backToMon = (curDay === 0 ? 6 : curDay - 1);
      startD.setDate(startD.getDate() - backToMon - 21); // 3 prior weeks + current week
    }
  }

  if (!endD) {
    if (maxLoggedDate && maxLoggedDate > today) {
      endD = new Date(maxLoggedDate.getTime());
    } else {
      endD = new Date(today.getTime());
    }
    // Expand forward to Sunday of that week
    var eDay = endD.getDay();
    var forwardToSun = (eDay === 0 ? 0 : 7 - eDay);
    endD.setDate(endD.getDate() + forwardToSun);
  }

  var allDays = getAllMatrixDays(startD, endD);
  if (allDays.length === 0) {
    return { success: false, message: "No calendar days found in date range" };
  }

  // 4. Locate or create view_weekly_matrix tab
  var viewSheet = getTable(ss, "weeklyMatrix");
  if (!viewSheet) {
    viewSheet = ss.insertSheet(tableSheetName("weeklyMatrix"));
  }

  // Unhide all rows and columns first so stale hidden states are cleared
  try {
    viewSheet.showRows(1, Math.max(viewSheet.getMaxRows(), 100));
  } catch (eShow) {}
  try {
    viewSheet.showColumns(1, Math.max(viewSheet.getMaxColumns(), 2));
  } catch (eShowCols) {}

  // 5. Build Header Rows (Two-row structure)
  // Row 1: "Week Number", "Date", Member 1 (merged 2 cols), Member 2 (merged 2 cols)...
  // Row 2: "",            "",     "AM", "PM",               "AM", "PM"...
  var numMembers = activeMembers.length;
  var numCols = 2 + (numMembers * 2);

  var row1Headers = ["Week Number", "Date"];
  var row2Headers = ["", ""];

  for (var a = 0; a < numMembers; a++) {
    row1Headers.push(activeMembers[a].name);
    row1Headers.push(""); // second merged cell
    row2Headers.push("AM");
    row2Headers.push("PM");
  }

  // 6. Build Matrix Data Rows
  var rows = [];
  var weekGroups = []; // Array of { weekStr: string, startRowIndex: number, rowCount: number }
  var currentWeekStr = "";
  var currentWeekStart = 0;
  var currentWeekCount = 0;

  for (var w = 0; w < allDays.length; w++) {
    var dayDate = allDays[w];
    var weekStr = formatIsoWeekString(dayDate);
    var dateStr = formatMatrixDateIso(dayDate);

    // Track week groups for Column A merging
    if (weekStr !== currentWeekStr) {
      if (currentWeekStr) {
        weekGroups.push({ weekStr: currentWeekStr, startRowIndex: currentWeekStart, rowCount: currentWeekCount });
      }
      currentWeekStr = weekStr;
      currentWeekStart = w;
      currentWeekCount = 1;
    } else {
      currentWeekCount++;
    }

    var row = [weekStr, dateStr];

    var dayNum = dayDate.getDay();
    var isWeekend = (dayNum === 0 || dayNum === 6);

    for (var mIdx = 0; mIdx < numMembers; mIdx++) {
      var member = activeMembers[mIdx];
      var key = dateStr + "::" + member.key;
      var log = presenceData[key];

      if (isWeekend) {
        // Weekends have dark gray background, no default presence required
        row.push(log ? log.am : "");
        row.push(log ? log.pm : "");
      } else {
        row.push(log ? log.am : "");
        row.push(log ? log.pm : "");
      }
    }
    rows.push(row);
  }

  if (currentWeekStr) {
    weekGroups.push({ weekStr: currentWeekStr, startRowIndex: currentWeekStart, rowCount: currentWeekCount });
  }

  var numDataRows = rows.length;
  var totalSheetRows = numDataRows + 2; // Rows 1 & 2 are headers, Data starts at row 3

  // Clear existing content & format
  viewSheet.getRange(1, 1, Math.max(viewSheet.getMaxRows(), totalSheetRows), Math.max(viewSheet.getMaxColumns(), numCols)).clearContent();

  // Write headers (Row 1 & Row 2)
  viewSheet.getRange(1, 1, 1, numCols).setValues([row1Headers]);
  viewSheet.getRange(2, 1, 1, numCols).setValues([row2Headers]);

  // Merge Column A (A1:A2) and Column B (B1:B2)
  try {
    viewSheet.getRange(1, 1, 2, 1).merge().setValue("Week Number");
    viewSheet.getRange(1, 2, 2, 1).merge().setValue("Date");
  } catch (eMergeA) {}

  // Merge member header cells in Row 1 across AM and PM columns and copy People Chips
  for (var m = 0; m < numMembers; m++) {
    var colStart = 3 + (m * 2);
    try {
      // 1. Copy People Chip to (1, colStart) first
      if (typeof copyCellChip === "function" && membersSheet && activeMembers[m].sourceRow) {
        copyCellChip(membersSheet, activeMembers[m].sourceRow, 1, viewSheet, 1, colStart, activeMembers[m].name, ss.getId());
      } else {
        viewSheet.getRange(1, colStart).setValue(activeMembers[m].name);
      }
      // 2. Explicitly merge the member's header cell across AM and PM columns AFTER copying chip so merge is not lost
      var memberRange = viewSheet.getRange(1, colStart, 1, 2);
      memberRange.merge();
      memberRange.setHorizontalAlignment("center");
      memberRange.setVerticalAlignment("middle");
      memberRange.setFontWeight("bold");
    } catch (eMergeM) {}
  }

  // Batch merge via Sheets API if available to guarantee persistent UI merge in Google Sheets
  try {
    if (typeof Sheets !== "undefined" && Sheets.Spreadsheets && ss.getId()) {
      var mergeRequests = [];
      // A1:A2 (Week Number) and B1:B2 (Date)
      mergeRequests.push({
        mergeCells: {
          range: { sheetId: viewSheet.getSheetId(), startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 1 },
          mergeType: "MERGE_ALL"
        }
      });
      mergeRequests.push({
        mergeCells: {
          range: { sheetId: viewSheet.getSheetId(), startRowIndex: 0, endRowIndex: 2, startColumnIndex: 1, endColumnIndex: 2 },
          mergeType: "MERGE_ALL"
        }
      });
      // Merged member headers in Row 1
      for (var mm = 0; mm < numMembers; mm++) {
        var mColStart = 2 + (mm * 2); // 0-indexed column
        mergeRequests.push({
          mergeCells: {
            range: { sheetId: viewSheet.getSheetId(), startRowIndex: 0, endRowIndex: 1, startColumnIndex: mColStart, endColumnIndex: mColStart + 2 },
            mergeType: "MERGE_ALL"
          }
        });
      }
      Sheets.Spreadsheets.batchUpdate({ requests: mergeRequests }, ss.getId());
    }
  } catch (eBatchMerge) {}

  // Format Header Rows (Rows 1 & 2)
  var fullHeaderRange = viewSheet.getRange(1, 1, 2, numCols);
  fullHeaderRange.setFontWeight("bold");
  fullHeaderRange.setBackground("#E8EAED");
  fullHeaderRange.setFontColor("#202124");
  fullHeaderRange.setHorizontalAlignment("center");
  fullHeaderRange.setVerticalAlignment("middle");

  // Freeze top 2 header rows
  viewSheet.setFrozenRows(2);

  // Write data rows starting at Row 3
  if (numDataRows > 0) {
    var dataRange = viewSheet.getRange(3, 1, numDataRows, numCols);
    dataRange.setValues(rows);

    // Merge Column A by week
    for (var g = 0; g < weekGroups.length; g++) {
      var group = weekGroups[g];
      var gStartRow = 3 + group.startRowIndex;
      try {
        var weekCell = viewSheet.getRange(gStartRow, 1, group.rowCount, 1);
        weekCell.merge().setValue(group.weekStr);
        weekCell.setVerticalAlignment("middle");
        weekCell.setHorizontalAlignment("center");
        weekCell.setFontWeight("bold");
      } catch (eMergeW) {}
    }

    // Apply dark gray background to weekend rows (#5F6368)
    for (var dIdx = 0; dIdx < allDays.length; dIdx++) {
      var dCheck = allDays[dIdx];
      var dDayNum = dCheck.getDay();
      if (dDayNum === 0 || dDayNum === 6) {
        var weekendRowNumber = 3 + dIdx;
        try {
          var weekendRange = viewSheet.getRange(weekendRowNumber, 1, 1, numCols);
          weekendRange.setBackground("#5F6368");
          weekendRange.setFontColor("#FFFFFF");
        } catch (eWkBg) {}
      }
    }

    // Alignment: Column B (Date) center-aligned
    viewSheet.getRange(3, 2, numDataRows, 1).setHorizontalAlignment("center");

    // Dropdown of ACTIVE locations from the current week on; past weeks get none, so
    // retired locations in history are not flagged as invalid (ADR 0040)
    try {
      var locationChoices = ((typeof activeLocationNames === "function") ? activeLocationNames(ss) : ["Remote"])
        .concat(["Off", "Vacation", "Sick Leave", "Holiday"]);
      var mondayIso = formatMatrixDateIso(new Date(today.getFullYear(), today.getMonth(), today.getDate() - ((today.getDay() + 6) % 7)));
      var firstEditable = numDataRows;
      for (var fr = 0; fr < rows.length; fr++) {
        if (String(rows[fr][1]) >= mondayIso) { firstEditable = fr; break; }
      }
      if (firstEditable > 0) {
        viewSheet.getRange(3, 3, Math.min(firstEditable, numDataRows), numCols - 2).clearDataValidations();
      }
      if (firstEditable < numDataRows) {
        var validationRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(locationChoices, true)
          .setAllowInvalid(true)
          .build();
        viewSheet.getRange(3 + firstEditable, 3, numDataRows - firstEditable, numCols - 2).setDataValidation(validationRule);
      }
    } catch (eVal) {}

    // Apply Conditional Formatting for presence modes
    applyMatrixConditionalFormatting(viewSheet, viewSheet.getRange(3, 3, numDataRows, numCols - 2), ss);

    // Hide departed Members (their columns stay, just hidden)
    for (var dm = 0; dm < numMembers; dm++) {
      if (activeMembers[dm].departed) {
        try { viewSheet.hideColumns(3 + dm * 2, 2); } catch (eHideCol) {}
      }
    }

    // Hide prior weeks: only display the last 4 weeks by default
    var totalWeeks = weekGroups.length;
    if (totalWeeks > 4) {
      var weeksToHide = totalWeeks - 4;
      var rowsToHideCount = 0;
      for (var h = 0; h < weeksToHide; h++) {
        rowsToHideCount += weekGroups[h].rowCount;
      }
      try {
        viewSheet.hideRows(3, rowsToHideCount);
      } catch (eHide) {}
    }
  }

  // Apply Typography
  if (typeof applyTypography === "function") {
    applyTypography(viewSheet.getRange(1, 1, totalSheetRows, numCols));
  } else {
    viewSheet.getRange(1, 1, totalSheetRows, numCols).setFontFamily("Google Sans").setFontSize(12);
  }

  // Column Widths
  try {
    viewSheet.setColumnWidth(1, 110); // Week Number
    viewSheet.setColumnWidth(2, 120); // Date
    for (var c = 3; c <= numCols; c++) {
      viewSheet.setColumnWidth(c, 110);
    }
  } catch (eWidth) {}

  return {
    success: true,
    sheetName: viewSheet.getName(),
    rowsRendered: numDataRows,
    membersCount: numMembers,
    totalWeeks: weekGroups.length,
    startDate: formatMatrixDateIso(startD),
    endDate: formatMatrixDateIso(endD)
  };
}

/**
 * Handles in-sheet presence declaration when a member cell in `view_weekly_matrix` is edited.
 * Preserves data integrity by syncing edits immediately to `tbl_presence_log` and `tbl_members.today_presence`.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 * @return {Object|null}
 */
function handleWeeklyMatrixEdit(e) {
  if (!e || !e.range) return null;
  var sheet = e.range.getSheet();
  if (!sheet || tableKeyOfSheet(sheet) !== "weeklyMatrix") return null;

  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();

  // Data rows start at Row 3; member columns start at Column 3
  if (editedRow < 3 || editedCol < 3) return null;

  var ss = (typeof sheet.getParent === "function" ? sheet.getParent() : null) || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;

  // 1. Identify date from Column 2 on that row
  var dateVal = sheet.getRange(editedRow, 2).getValue();
  var dateObj = parseMatrixDate(dateVal);
  if (!dateObj) return null;
  var dateIso = formatMatrixDateIso(dateObj);

  // 2. Identify member and period (AM or PM) from column
  var memberIdx = Math.floor((editedCol - 3) / 2);
  var isPm = ((editedCol - 3) % 2) === 1;
  var period = isPm ? "PM" : "AM";

  var memberHeaderCol = 3 + (memberIdx * 2);
  var memberHeaderCell = sheet.getRange(1, memberHeaderCol);
  var memberHeaderVal = memberHeaderCell.getValue();
  var memberHeaderRich = (typeof memberHeaderCell.getRichTextValue === "function") ? memberHeaderCell.getRichTextValue() : null;
  var memberName = (typeof extractCellValue === "function") ? extractCellValue(memberHeaderVal, memberHeaderRich) : String(memberHeaderVal).trim();

  if (!memberName) return null;

  var newValue = String(e.range.getValue() || "").trim();

  // 3. Update or Insert in tbl_presence_log
  var logSheet = getTable(ss, "presenceLog");
  if (!logSheet) return null;

  var logLastRow = logSheet.getLastRow();
  var foundRow = -1;
  var existingAm = "";
  var existingPm = "";

  if (logLastRow > 1) {
    var lRange = logSheet.getRange(2, 1, logLastRow - 1, 5);
    var lVals = lRange.getValues();
    var lRich = (typeof lRange.getRichTextValues === "function") ? lRange.getRichTextValues() : [];

    for (var l = 0; l < lVals.length; l++) {
      var lDate = parseMatrixDate(lVals[l][1]);
      if (!lDate) continue;
      if (formatMatrixDateIso(lDate) === dateIso) {
        var mVal = lVals[l][2];
        var mR = (lRich && lRich[l]) ? lRich[l][2] : null;
        var mName = (typeof extractCellValue === "function") ? extractCellValue(mVal, mR) : String(mVal).trim();
        var isMatch = (typeof isUserMemberMatch === "function") ? isUserMemberMatch(mVal, mR, memberName) : (mName.toLowerCase() === memberName.toLowerCase());
        if (isMatch || mName.toLowerCase() === memberName.toLowerCase()) {
          foundRow = l + 2;
          existingAm = String(lVals[l][3] || "").trim();
          existingPm = String(lVals[l][4] || "").trim();
          break;
        }
      }
    }
  }

  var finalAm = isPm ? existingAm : newValue;
  var finalPm = isPm ? newValue : existingPm;

  var currentTimestampStr = formatMatrixDateTime(new Date());

  if (foundRow > 0) {
    if (isPm) {
      logSheet.getRange(foundRow, 5).setValue(newValue);
    } else {
      logSheet.getRange(foundRow, 4).setValue(newValue);
    }
    var tsCell = logSheet.getRange(foundRow, 6);
    tsCell.setValue(currentTimestampStr);
    try { tsCell.setNumberFormat("@"); } catch (eFmt) {}
  } else if (newValue) {
    var targetRow = logLastRow + 1;
    var logId = "log-" + (typeof Utilities !== "undefined" && Utilities.getUuid ? Utilities.getUuid() : Math.floor(Math.random() * 1000000));
    logSheet.getRange(targetRow, 1).setValue(logId);
    logSheet.getRange(targetRow, 2).setValue(dateIso);

    // Copy member People Chip from tbl_members
    var membersSheet = getTable(ss, "members");
    var chipCopied = false;
    if (membersSheet && typeof copyCellChip === "function") {
      var mMax = readRowLimit(membersSheet, 200);
      var mVals = membersSheet.getRange(2, 1, mMax - 1, 1).getValues();
      for (var mi = 0; mi < mVals.length; mi++) {
        var mRowName = (typeof extractCellValue === "function") ? extractCellValue(mVals[mi][0]) : String(mVals[mi][0]).trim();
        if (mRowName.toLowerCase() === memberName.toLowerCase()) {
          copyCellChip(membersSheet, mi + 2, 1, logSheet, targetRow, 3, memberName, ss.getId());
          chipCopied = true;
          break;
        }
      }
    }
    if (!chipCopied) {
      logSheet.getRange(targetRow, 3).setValue(memberName);
    }

    logSheet.getRange(targetRow, 4).setValue(finalAm);
    logSheet.getRange(targetRow, 5).setValue(finalPm);
    var tsCellNew = logSheet.getRange(targetRow, 6);
    tsCellNew.setValue(currentTimestampStr);
    try { tsCellNew.setNumberFormat("@"); } catch (eFmt2) {}

    if (typeof applyTypography === "function") {
      applyTypography(logSheet.getRange(targetRow, 1, 1, 6));
    }
  }

  // 4. Update today_presence in tbl_members if date is today
  var todayIso = formatMatrixDateIso(new Date());
  if (dateIso === todayIso) {
    var membersSheet2 = getTable(ss, "members");
    if (membersSheet2) {
      var mMax2 = readRowLimit(membersSheet2, 200);
      var mRange2 = membersSheet2.getRange(2, 1, mMax2 - 1, 7);
      var mVals2 = mRange2.getValues();
      for (var m2 = 0; m2 < mVals2.length; m2++) {
        var memName2 = (typeof extractCellValue === "function") ? extractCellValue(mVals2[m2][0]) : String(mVals2[m2][0]).trim();
        if (memName2.toLowerCase() === memberName.toLowerCase()) {
          var todayPresenceVal = (finalAm && finalPm && finalAm === finalPm) ? finalAm : ((finalAm || "") + (finalAm && finalPm ? " / " : "") + (finalPm || ""));
          membersSheet2.getRange(m2 + 2, 7).setValue(todayPresenceVal || "Not Logged");
          break;
        }
      }
    }
  }

  try {
    ss.toast(
      "Saved " + memberName + " " + dateIso + " " + period + ": " + (newValue || "cleared") + ". " +
        "For a recurring routine or a leave, use Project Management → Declare Presence & Leaves.",
      "Presence Logged",
      6
    );
  } catch (eToast) {}

  return {
    success: true,
    member: memberName,
    date: dateIso,
    period: period,
    value: newValue
  };
}

/**
 * Custom Menu action to refresh the Weekly Matrix view.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object}
 */
function refreshWeeklyMatrix(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var res = renderWeeklyMatrixView(null, null, ss);
  if (res.success && ss) {
    try {
      ss.toast(
        "Refreshed Weekly Matrix view with " + res.rowsRendered + " days across " + res.membersCount + " members.",
        "Weekly Matrix Updated",
        5
      );
    } catch (eToast) {}
  }
  return res;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    renderWeeklyMatrixView: renderWeeklyMatrixView,
    refreshWeeklyMatrix: refreshWeeklyMatrix,
    handleWeeklyMatrixEdit: handleWeeklyMatrixEdit,
    getIsoWeekNumber: getIsoWeekNumber,
    formatIsoWeekString: formatIsoWeekString,
    formatMatrixDateIso: formatMatrixDateIso,
    formatMatrixDateTime: formatMatrixDateTime,
    parseMatrixDate: parseMatrixDate,
    getAllMatrixDays: getAllMatrixDays,
    applyMatrixConditionalFormatting: applyMatrixConditionalFormatting
  };
}
