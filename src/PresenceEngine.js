/**
 * Project Management Workspace - Presence Engine
 * Ticket 03: Baseline & Leave Range Compiler to Daily Presence Log
 *
 * Expands declarative baselines (tbl_presence_baselines) and approved leaves (tbl_leaves)
 * into concrete, queryable daily records in tbl_presence_log for working days (Mon-Fri).
 * Preserves existing ad-hoc manual adjustments (idempotent compilation).
 */

/**
 * Returns array of Date objects for weekdays (Monday to Friday) between start and end (inclusive).
 *
 * @param {Date} startDate
 * @param {Date} endDate
 * @return {Array<Date>}
 */
function getWeekdaysBetween(startDate, endDate) {
  var days = [];
  var curr = new Date(startDate.getTime());
  curr.setHours(0, 0, 0, 0);

  var end = new Date(endDate.getTime());
  end.setHours(0, 0, 0, 0);

  while (curr <= end) {
    var dayOfWeek = curr.getDay(); // 0 = Sun, 1 = Mon, ..., 6 = Sat
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      days.push(new Date(curr.getTime()));
    }
    curr.setDate(curr.getDate() + 1);
  }
  return days;
}

/**
 * Helper to get English weekday name from Date object.
 *
 * @param {Date} date
 * @return {string} "Monday", "Tuesday", etc.
 */
function getWeekdayName(date) {
  var names = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  return names[date.getDay()];
}

/**
 * Returns array of all Date objects (including weekends) between start and end (inclusive).
 *
 * @param {Date} startDate
 * @param {Date} endDate
 * @return {Array<Date>}
 */
function getAllDaysBetween(startDate, endDate) {
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
 * Scans tbl_tenures, tbl_presence_baselines, tbl_leaves, and tbl_presence_log
 * to discover the earliest recorded date for data migration and historical compilation.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Date|null}
 */
function findEarliestHistoricalDate(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;

  var earliestTime = null;

  function checkDate(rawVal) {
    if (!rawVal) return;
    var d = null;
    if (typeof parseMatrixDate === "function") {
      d = parseMatrixDate(rawVal);
    } else if (rawVal instanceof Date) {
      d = new Date(rawVal.getTime());
      d.setHours(0, 0, 0, 0);
    } else {
      var s = String(rawVal).trim();
      var match = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
      if (match) {
        d = new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10), 0, 0, 0, 0);
      } else {
        var parsed = new Date(s);
        if (!isNaN(parsed.getTime())) {
          parsed.setHours(0, 0, 0, 0);
          d = parsed;
        }
      }
    }
    if (d && !isNaN(d.getTime())) {
      if (d.getFullYear() >= 2000 && d.getFullYear() <= 2100) {
        var t = d.getTime();
        if (earliestTime === null || t < earliestTime) {
          earliestTime = t;
        }
      }
    }
  }

  // 1. Check tbl_tenures (Col 4: start_date)
  var tenuresSheet = getTable(ss, "tenures");
  if (tenuresSheet) {
    var tMax = readRowLimit(tenuresSheet, 500);
    if (tMax > 1) {
      var tVals = tenuresSheet.getRange(2, 4, tMax - 1, 1).getValues();
      for (var t = 0; t < tVals.length; t++) {
        checkDate(tVals[t][0]);
      }
    }
  }

  // 2. Check tbl_presence_baselines (Col 3: start_date)
  var baselinesSheet = getTable(ss, "baselines");
  if (baselinesSheet) {
    var bMax = readRowLimit(baselinesSheet, 500);
    if (bMax > 1) {
      var bVals = baselinesSheet.getRange(2, 3, bMax - 1, 1).getValues();
      for (var b = 0; b < bVals.length; b++) {
        checkDate(bVals[b][0]);
      }
    }
  }

  // 3. Check tbl_leaves (Col 3: start_date)
  var leavesSheet = getTable(ss, "leaves");
  if (leavesSheet) {
    var lMax = readRowLimit(leavesSheet, 500);
    if (lMax > 1) {
      var lVals = leavesSheet.getRange(2, 3, lMax - 1, 1).getValues();
      for (var l = 0; l < lVals.length; l++) {
        checkDate(lVals[l][0]);
      }
    }
  }

  // 4. Check tbl_presence_log (Col 2: date)
  var logSheet = getTable(ss, "presenceLog");
  if (logSheet) {
    var lgMax = readRowLimit(logSheet, 2000);
    if (lgMax > 1) {
      var lgVals = logSheet.getRange(2, 2, lgMax - 1, 1).getValues();
      for (var lg = 0; lg < lgVals.length; lg++) {
        checkDate(lgVals[lg][0]);
      }
    }
  }

  return earliestTime !== null ? new Date(earliestTime) : null;
}

/**
 * Compiles presence logs for a given date range across all active members.
 * Reads tbl_presence_baselines and tbl_leaves, and writes into tbl_presence_log.
 * Preserves existing rows with manual edits. Updates tbl_members.today_presence.
 *
 * NOTE: tbl_presence_log serves as the ground-truth database for actual recorded presences.
 * Compilation is strictly bounded by today's date; future dates are never compiled into the log.
 *
 * @param {Date|string} [startDateParam] - Default: 1st of current month (or earliest if options.fromEarliest / 'EARLIEST')
 * @param {Date|string} [endDateParam] - Default: Today's date (strictly capped at today)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Date|string} [asOfDateParam] - Optional reference date representing 'today' (defaults to current date)
 * @param {Object} [options] - Optional configuration options, e.g. { fromEarliest: true }
 * @return {Object} Compilation summary report.
 */
function compilePresenceLogs(startDateParam, endDateParam, targetSpreadsheet, asOfDateParam, options) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet" };

  var membersSheet = getTable(ss, "members");
  var baselinesSheet = getTable(ss, "baselines");
  var leavesSheet = getTable(ss, "leaves");
  var logSheet = getTable(ss, "presenceLog");

  if (!membersSheet || !logSheet) {
    return { success: false, message: "Required tabs (Members, Presence Log) not found." };
  }

  // 1. Determine Date Range strictly capped at today's date (ground truth constraint)
  var now = asOfDateParam ? new Date(asOfDateParam) : new Date();
  var todayDate = new Date(now.getTime());
  todayDate.setHours(0, 0, 0, 0);

  var startDate = null;
  if (startDateParam && startDateParam !== "EARLIEST") {
    startDate = new Date(startDateParam);
  } else if (startDateParam === "EARLIEST" || (options && options.fromEarliest)) {
    var detected = findEarliestHistoricalDate(ss);
    startDate = detected || new Date(now.getFullYear(), 0, 1);
  } else {
    startDate = new Date(now.getFullYear(), now.getMonth(), 1);
  }

  var requestedEndDate = endDateParam ? new Date(endDateParam) : new Date(todayDate.getTime());

  startDate.setHours(0, 0, 0, 0);
  requestedEndDate.setHours(0, 0, 0, 0);

  // Ground truth constraint: compilation strictly cannot exceed today's date
  var endDate = requestedEndDate > todayDate ? todayDate : requestedEndDate;

  if (startDate > endDate) {
    return {
      success: true,
      createdCount: 0,
      preservedCount: 0,
      message: "Requested start date is after the ground-truth limit (today: " + (typeof formatDateIso === "function" ? formatDateIso(todayDate) : "") + ")."
    };
  }

  var workingDays = getWeekdaysBetween(startDate, endDate);
  var allCalendarDays = getAllDaysBetween(startDate, endDate);
  if (allCalendarDays.length === 0) {
    return { success: true, createdCount: 0, preservedCount: 0, workingDaysCount: 0, message: "No calendar days in range" };
  }

  // 2. Read Active Members from tbl_members
  // Columns: member(1), accenture_email(2), phone_number(3), current_role(4), current_department(5), tenure_status(6), today_presence(7)
  var membersMaxRows = readRowLimit(membersSheet, 300);
  var membersList = [];

  if (membersMaxRows > 1) {
    var mRange = membersSheet.getRange(2, 1, membersMaxRows - 1, Math.max(membersSheet.getLastColumn(), 7));
    var mValues = mRange.getValues();
    var mRich = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : [];
    for (var m = 0; m < mValues.length; m++) {
      var rawM = mValues[m][0];
      var richM = (mRich && mRich[m]) ? mRich[m][0] : null;
      var mVal = typeof extractCellValue === "function" ? extractCellValue(rawM, richM) : String(rawM || "").trim();
      if (!mVal) continue;

      var tStatus = typeof extractCellValue === "function" ? extractCellValue(mValues[m][5]) : String(mValues[m][5] || "").trim();
      // Skip departed members
      if (tStatus && tStatus.toLowerCase() === "departed") continue;

      membersList.push({
        row: m + 2,
        member: mVal,
        raw: rawM,
        richText: richM,
        key: mVal.toLowerCase()
      });
    }
  }

  if (membersList.length === 0) {
    return { success: true, createdCount: 0, preservedCount: 0, message: "No active members found." };
  }

  // Weekdays without a routine default to the first active office (Locations tab, ADR 0040)
  var defaultOffice = (typeof defaultOfficeLocation === "function") ? defaultOfficeLocation(ss) : "Office";

  // 3. Read Baselines from tbl_presence_baselines
  // Columns: baseline_id(1), member(2), start_date(3), end_date(4), day_of_week(5), location_am(6), location_pm(7)
  var baselinesByMember = {};
  if (baselinesSheet) {
    var bMaxRows = readRowLimit(baselinesSheet, 500);
    if (bMaxRows > 1) {
      var bValues = baselinesSheet.getRange(2, 1, bMaxRows - 1, 7).getValues();
      for (var b = 0; b < bValues.length; b++) {
        var bMember = typeof extractCellValue === "function" ? extractCellValue(bValues[b][1]) : String(bValues[b][1] || "").trim();
        if (!bMember) continue;

        var bKey = bMember.toLowerCase();
        if (!baselinesByMember[bKey]) baselinesByMember[bKey] = [];

        var sDate = bValues[b][2] ? new Date(bValues[b][2]) : null;
        var eDate = bValues[b][3] ? new Date(bValues[b][3]) : null;
        if (sDate) sDate.setHours(0, 0, 0, 0);
        if (eDate) eDate.setHours(0, 0, 0, 0);

        var dow = typeof extractCellValue === "function" ? extractCellValue(bValues[b][4]) : String(bValues[b][4] || "").trim();
        var locAm = typeof extractCellValue === "function" ? extractCellValue(bValues[b][5]) : String(bValues[b][5] || "").trim();
        var locPm = typeof extractCellValue === "function" ? extractCellValue(bValues[b][6]) : String(bValues[b][6] || "").trim();

        baselinesByMember[bKey].push({
          startDate: sDate,
          endDate: eDate,
          dayOfWeek: dow,
          locationAm: locAm || defaultOffice,
          locationPm: locPm || defaultOffice
        });
      }
    }
  }

  // 4. Read Leaves from tbl_leaves
  // Supports both modern 8-column (with start_period & end_period) and legacy 6-column schemas
  var leavesByMember = {};
  if (leavesSheet) {
    var lColMap = (typeof getLeaveColumnMap === "function")
      ? getLeaveColumnMap(leavesSheet)
      : { leave_id: 1, member: 2, start_date: 3, start_period: -1, end_date: 4, end_period: -1, leave_type: 5, status: 6, numCols: 6 };
    var lMaxRows = readRowLimit(leavesSheet, 500);
    if (lMaxRows > 1) {
      var lValues = leavesSheet.getRange(2, 1, lMaxRows - 1, Math.max(leavesSheet.getLastColumn(), 6)).getValues();
      for (var l = 0; l < lValues.length; l++) {
        var lMember = typeof extractCellValue === "function" ? extractCellValue(lValues[l][lColMap.member - 1]) : String(lValues[l][lColMap.member - 1] || "").trim();
        if (!lMember) continue;

        var lStatus = typeof extractCellValue === "function" ? extractCellValue(lValues[l][lColMap.status - 1]) : String(lValues[l][lColMap.status - 1] || "").trim();
        // Only factor Approved leaves
        if (lStatus && lStatus.toLowerCase() !== "approved") continue;

        var lKey = lMember.toLowerCase();
        if (!leavesByMember[lKey]) leavesByMember[lKey] = [];

        var lStart = lValues[l][lColMap.start_date - 1] ? new Date(lValues[l][lColMap.start_date - 1]) : null;
        var lEnd = lValues[l][lColMap.end_date - 1] ? new Date(lValues[l][lColMap.end_date - 1]) : null;
        if (lStart) lStart.setHours(0, 0, 0, 0);
        if (lEnd) lEnd.setHours(0, 0, 0, 0);

        var lStartPeriod = (lColMap.start_period > 0 && lValues[l][lColMap.start_period - 1])
          ? String(lValues[l][lColMap.start_period - 1]).toUpperCase().trim()
          : "AM";
        var lEndPeriod = (lColMap.end_period > 0 && lValues[l][lColMap.end_period - 1])
          ? String(lValues[l][lColMap.end_period - 1]).toUpperCase().trim()
          : "PM";

        var lType = typeof extractCellValue === "function" ? extractCellValue(lValues[l][lColMap.leave_type - 1]) : String(lValues[l][lColMap.leave_type - 1] || "").trim();

        leavesByMember[lKey].push({
          startDate: lStart,
          startPeriod: lStartPeriod,
          endDate: lEnd,
          endPeriod: lEndPeriod,
          leaveType: lType || "Vacation"
        });
      }
    }
  }

  // 5. Read Existing tbl_presence_log to Ensure Idempotence and Preserve Manual Edits
  // Columns: log_id(1), date(2), member(3), location_am(4), location_pm(5), timestamp(6)
  var logMaxRows = readRowLimit(logSheet, 2000);
  var existingLogMap = {}; // key: "YYYY-MM-DD|memberLower" -> { row, locationAm, locationPm }
  var preservedCount = 0;

  if (logMaxRows > 1) {
    var logValues = logSheet.getRange(2, 1, logMaxRows - 1, 6).getValues();
    for (var lg = 0; lg < logValues.length; lg++) {
      var lgId = logValues[lg][0];
      var lgDate = logValues[lg][1];
      var lgMember = typeof extractCellValue === "function" ? extractCellValue(logValues[lg][2]) : String(logValues[lg][2] || "").trim();
      if (!lgId && !lgDate && !lgMember) continue;

      var dateStr = typeof formatDateIso === "function" ? formatDateIso(lgDate) : String(lgDate).substring(0, 10);
      var key = dateStr + "|" + lgMember.toLowerCase();

      existingLogMap[key] = {
        row: lg + 2,
        locationAm: logValues[lg][3],
        locationPm: logValues[lg][4]
      };
    }
  }

  // Helper to resolve baseline location for a member on a specific date
  function resolveBaseline(mKey, dateObj, dayName) {
    var isWeekend = (dateObj.getDay() === 0 || dateObj.getDay() === 6);
    var rules = baselinesByMember[mKey] || [];
    for (var r = 0; r < rules.length; r++) {
      var rule = rules[r];
      if (rule.dayOfWeek && rule.dayOfWeek !== "All" && rule.dayOfWeek.toLowerCase() !== dayName.toLowerCase()) {
        continue;
      }
      if (rule.startDate && dateObj < rule.startDate) continue;
      if (rule.endDate && dateObj > rule.endDate) continue;
      return { am: rule.locationAm, pm: rule.locationPm };
    }
    var defaultLoc = isWeekend ? "Off" : defaultOffice;
    return { am: defaultLoc, pm: defaultLoc };
  }

  // Helper to check if member has an approved leave on date, returning { am, pm } honoring half-day boundaries
  function resolveLeave(mKey, dateObj) {
    var lList = leavesByMember[mKey] || [];
    var dTime = dateObj.getTime();

    for (var i = 0; i < lList.length; i++) {
      var lv = lList[i];
      if (!lv.startDate || !lv.endDate) continue;
      var sTime = lv.startDate.getTime();
      var eTime = lv.endDate.getTime();

      if (dTime < sTime || dTime > eTime) continue;

      var amLeave = null;
      var pmLeave = null;

      if (sTime === eTime) {
        // Single day leave
        if (lv.startPeriod === "AM" && lv.endPeriod === "PM") {
          amLeave = lv.leaveType;
          pmLeave = lv.leaveType;
        } else if (lv.startPeriod === "AM" && lv.endPeriod === "AM") {
          amLeave = lv.leaveType;
          pmLeave = null;
        } else if (lv.startPeriod === "PM" && lv.endPeriod === "PM") {
          amLeave = null;
          pmLeave = lv.leaveType;
        } else {
          amLeave = lv.leaveType;
          pmLeave = lv.leaveType;
        }
      } else {
        // Multi-day leave
        if (dTime === sTime) {
          amLeave = (lv.startPeriod === "PM") ? null : lv.leaveType;
          pmLeave = lv.leaveType;
        } else if (dTime === eTime) {
          amLeave = lv.leaveType;
          pmLeave = (lv.endPeriod === "AM") ? null : lv.leaveType;
        } else {
          amLeave = lv.leaveType;
          pmLeave = lv.leaveType;
        }
      }

      return {
        am: amLeave,
        pm: pmLeave,
        leaveType: lv.leaveType
      };
    }
    return null;
  }

  // 6. Generate New Daily Presence Records
  var newLogRows = [];
  var newLogRichTexts = [];
  var newLogChipsToCopy = [];
  var todayDateStr = typeof formatDateIso === "function" ? formatDateIso(now) : String(now).substring(0, 10);
  var todayPresenceMap = {}; // memberKey -> string representation

  var currentGenTimestamp = (typeof formatDateIso === "function" ? formatDateIso(now) : "") + " " +
    ("0" + now.getHours()).slice(-2) + ":" + ("0" + now.getMinutes()).slice(-2) + ":" + ("0" + now.getSeconds()).slice(-2);

  for (var w = 0; w < allCalendarDays.length; w++) {
    var wDate = allCalendarDays[w];
    var isWeekendDay = (wDate.getDay() === 0 || wDate.getDay() === 6);
    var wDateStr = typeof formatDateIso === "function" ? formatDateIso(wDate) : String(wDate).substring(0, 10);
    var dayName = getWeekdayName(wDate);

    for (var mi = 0; mi < membersList.length; mi++) {
      var mem = membersList[mi];
      var mKey = mem.key;
      var mapKey = wDateStr + "|" + mKey;

      if (existingLogMap[mapKey]) {
        // Already exists: Preserve manual entries!
        preservedCount++;
        if (wDateStr === todayDateStr) {
          var ex = existingLogMap[mapKey];
          var exLoc = ex.locationAm === ex.locationPm ? ex.locationAm : (ex.locationAm + " / " + ex.locationPm);
          todayPresenceMap[mKey] = exLoc;
        }
        continue;
      }

      // 1. Evaluate baseline default
      var baseLoc = resolveBaseline(mKey, wDate, dayName);
      var locAm = baseLoc.am;
      var locPm = baseLoc.pm;

      // 2. Override with approved leave if active on this day
      var leaveRes = resolveLeave(mKey, wDate);
      if (leaveRes) {
        if (leaveRes.am) locAm = leaveRes.am;
        if (leaveRes.pm) locPm = leaveRes.pm;
      }

      // 3. Weekend handling: Saturday/Sunday only logged if member actually works or has leave
      if (isWeekendDay) {
        var isWorkingWeekend = (locAm !== "Off" || locPm !== "Off");
        var hasLeaveOnWeekend = (leaveRes !== null && (leaveRes.am !== null || leaveRes.pm !== null));
        if (!isWorkingWeekend && !hasLeaveOnWeekend) {
          continue;
        }
      }

      var logUuid = "log_" + Math.random().toString(36).substring(2, 9);

      newLogRows.push([
        logUuid,
        wDateStr,
        mem.raw, // Directly write raw cell chip/object without flashing plain text
        locAm,
        locPm,
        currentGenTimestamp
      ]);

      newLogRichTexts.push([ mem.richText ]);

      newLogChipsToCopy.push({
        sourceRow: mem.row,
        sourceCol: 1,
        memberName: mem.member
      });

      if (wDateStr === todayDateStr) {
        var tLoc = (locAm === locPm) ? locAm : (locAm + " / " + locPm);
        todayPresenceMap[mKey] = tLoc;
      }
    }
  }

  // 7. Append New Rows to tbl_presence_log
  var createdCount = newLogRows.length;
  if (createdCount > 0) {
    var startAppendRow = 2;
    var lastR = (typeof logSheet.getLastRow === "function") ? (logSheet.getLastRow() || 1) : 1;
    if (lastR > 1) {
      // Find actual last populated row
      var checkVals = logSheet.getRange(2, 1, lastR - 1, 6).getValues();
      var actualEnd = 1;
      for (var cv = checkVals.length - 1; cv >= 0; cv--) {
        if (checkVals[cv][0] || checkVals[cv][1] || checkVals[cv][2]) {
          actualEnd = cv + 2;
          break;
        }
      }
      startAppendRow = actualEnd + 1;
    }

    logSheet.getRange(startAppendRow, 1, createdCount, 6).setValues(newLogRows);

    // If rich text values exist, set them directly in bulk to column 3
    try {
      var hasRich = false;
      for (var rk = 0; rk < newLogRichTexts.length; rk++) {
        if (newLogRichTexts[rk][0]) { hasRich = true; break; }
      }
      if (hasRich && typeof logSheet.getRange(startAppendRow, 3, createdCount, 1).setRichTextValues === "function") {
        logSheet.getRange(startAppendRow, 3, createdCount, 1).setRichTextValues(newLogRichTexts);
      }
    } catch (eRich) {}

    // Fallback chip copying if cell needs native Range.copyTo
    if (typeof copyCellChip === "function") {
      for (var c = 0; c < newLogChipsToCopy.length; c++) {
        var chipMeta = newLogChipsToCopy[c];
        var destRow = startAppendRow + c;
        copyCellChip(
          membersSheet,
          chipMeta.sourceRow,
          chipMeta.sourceCol,
          logSheet,
          destRow,
          3, // member is column 3 in tbl_presence_log
          chipMeta.memberName,
          ss.getId()
        );
      }
    }

    if (typeof applyTypography === "function") {
      applyTypography(logSheet.getRange(startAppendRow, 1, createdCount, 6));
    }
  }

  // 8. Update today_presence in tbl_members
  var isWeekendToday = (todayDate.getDay() === 0 || todayDate.getDay() === 6);
  var todayDayName = getWeekdayName(todayDate);
  for (var um = 0; um < membersList.length; um++) {
    var uMem = membersList[um];
    var todayVal = todayPresenceMap[uMem.key];
    if (!todayVal) {
      if (isWeekendToday) {
        var tLeave = resolveLeave(uMem.key, todayDate);
        var tBase = resolveBaseline(uMem.key, todayDate, todayDayName);
        if (tLeave && (tLeave.am || tLeave.pm)) {
          var lam = tLeave.am || tBase.am;
          var lpm = tLeave.pm || tBase.pm;
          todayVal = (lam === lpm) ? lam : (lam + " / " + lpm);
        } else if (tBase.am !== "Off" || tBase.pm !== "Off") {
          todayVal = (tBase.am === tBase.pm) ? tBase.am : (tBase.am + " / " + tBase.pm);
        } else {
          todayVal = "Weekend";
        }
      } else {
        todayVal = "Not Logged";
      }
    }
    membersSheet.getRange(uMem.row, 7).setValue(todayVal);
  }
  if (typeof applyTodayPresenceColours === "function") {
    try { applyTodayPresenceColours(ss); } catch (eTc) { Logger.log("Today colours note: " + eTc); }
  }

  // 9. Reapply protections so newly created log rows are protected while keeping empty rows open
  if (typeof applyProtections === "function") {
    try {
      applyProtections(ss);
    } catch (eProt) {
      Logger.log("Protection reapplication note: " + eProt);
    }
  }

  return {
    success: true,
    startDate: typeof formatDateIso === "function" ? formatDateIso(startDate) : String(startDate),
    endDate: typeof formatDateIso === "function" ? formatDateIso(endDate) : String(endDate),
    workingDaysCount: workingDays.length,
    activeMembersCount: membersList.length,
    createdCount: createdCount,
    preservedCount: preservedCount
  };
}

/**
 * Compiles presence logs from earliest/undefined historical date up to today
 * for data migration and historical population.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Compilation summary report
 */
function compilePresenceMigration(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet" };

  var detectedDate = findEarliestHistoricalDate(ss);
  var startDateToUse = detectedDate;

  var ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (eUi) {}

  if (ui && typeof ui.prompt === "function") {
    var detectedStr = detectedDate
      ? (typeof formatDateIso === "function" ? formatDateIso(detectedDate) : detectedDate.toISOString().slice(0, 10))
      : "None detected (defaults to Jan 1 of current year)";
    var promptRes = ui.prompt(
      "Compile Presence Logs (Migration / Undefined Start Date)",
      "Auto-detected earliest recorded date: " + detectedStr + "\n\nEnter start date (YYYY-MM-DD) or click OK to compile from earliest date:",
      ui.ButtonSet.OK_CANCEL
    );

    if (promptRes && typeof promptRes.getSelectedButton === "function") {
      var btn = promptRes.getSelectedButton();
      var cancelBtn = (ui.Button && ui.Button.CANCEL) ? ui.Button.CANCEL : "CANCEL";
      if (btn === cancelBtn) {
        if (typeof ss.toast === "function") {
          ss.toast("Presence migration cancelled.", "Migration Cancelled", 3);
        }
        return { success: false, cancelled: true, message: "Migration cancelled by user" };
      }
      var userText = String(promptRes.getResponseText ? promptRes.getResponseText() : "").trim();
      if (userText) {
        var match = userText.match(/^(\d{4})-(\d{2})-(\d{2})/);
        var parsedUserDate = match
          ? new Date(parseInt(match[1], 10), parseInt(match[2], 10) - 1, parseInt(match[3], 10), 0, 0, 0, 0)
          : new Date(userText);
        if (!isNaN(parsedUserDate.getTime())) {
          parsedUserDate.setHours(0, 0, 0, 0);
          startDateToUse = parsedUserDate;
        }
      }
    }
  }

  if (!startDateToUse) {
    var now = new Date();
    startDateToUse = new Date(now.getFullYear(), 0, 1); // Fallback: Jan 1 of current year
  }

  var res = compilePresenceLogs(startDateToUse, null, ss, null, { fromEarliest: true });

  if (res.success && ss) {
    if (typeof refreshWeeklyMatrix === "function") {
      try {
        refreshWeeklyMatrix(ss);
      } catch (eMatrix) {}
    }
    if (typeof ss.toast === "function") {
      var sStr = typeof formatDateIso === "function" ? formatDateIso(startDateToUse) : startDateToUse.toISOString().slice(0, 10);
      ss.toast(
        "Compiled " + res.createdCount + " migration daily presence logs from " + sStr + " (" + res.preservedCount + " existing preserved).",
        "Migration Completed",
        6
      );
    }
  }

  return res;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    compilePresenceLogs: compilePresenceLogs,
    compilePresenceMigration: compilePresenceMigration,
    findEarliestHistoricalDate: findEarliestHistoricalDate,
    getWeekdaysBetween: getWeekdaysBetween,
    getAllDaysBetween: getAllDaysBetween,
    getWeekdayName: getWeekdayName
  };
}
