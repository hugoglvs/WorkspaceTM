/**
 * Project Management Workspace - Bi-directional Synchronization Engine
 * Keeps tbl_members (Team Directory) synchronized from tbl_tenures.
 * Auto-creates profile completion tasks in tbl_tasks for new members.
 */

/**
 * Safe cell value extractor handling native People Chips, RichTextValues, Dates, and plain values.
 */
function safeExtractCell(val, richText) {
  if (typeof extractCellValue === "function") {
    return extractCellValue(val, richText);
  }
  if (richText && typeof richText.getText === "function") {
    var rtText = richText.getText();
    if (rtText) return String(rtText).trim();
  }
  if (val === null || val === undefined) return "";
  if (typeof val === "string") return val.trim();
  if (typeof val === "object") {
    if (val instanceof Date) return safeFormatDateIso(val);
    if (typeof val.getEmail === "function") return String(val.getEmail()).trim();
    if (typeof val.getName === "function") return String(val.getName()).trim();
    if (val.email) return String(val.email).trim();
    if (val.name) return String(val.name).trim();
  }
  return String(val).trim();
}

/**
 * Safe ISO date formatter (yyyy-mm-dd).
 */
function safeFormatDateIso(dateVal) {
  if (typeof formatDateIso === "function") {
    return formatDateIso(dateVal);
  }
  if (!dateVal) return "";
  if (dateVal instanceof Date) {
    var year = dateVal.getFullYear();
    var month = String(dateVal.getMonth() + 1).padStart(2, "0");
    var day = String(dateVal.getDate()).padStart(2, "0");
    return year + "-" + month + "-" + day;
  }
  return String(dateVal).trim();
}

/**
 * Synchronizes tbl_members from tbl_tenures:
 * 1. Reads tenures (columns: member, role, department, start_date, end_date, contract_type, fte_rate).
 * 2. Members added in tbl_tenures missing from tbl_members are copied as native People Chips.
 * 3. In tbl_members, current_role, current_department, and tenure_status are updated based on active tenure.
 * 4. When a new member is added, an onboarding task is assigned to their People Chip in tbl_tasks
 *    prompting them to fill in accenture_email and phone_number.
 * 5. Applies column protections to locked fields.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Summary report of additions and updates.
 */
function syncMembersAndTenures(targetSpreadsheet, options) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  // From an edit: skip the slow whole-workspace upkeep (chips in every Routine
  // row, all protections); the nightly sync does it (ADR 0044)
  var fromEdit = !!(options && options.fromEdit);
  if (!ss) return { success: false, message: "No active spreadsheet" };

  var membersSheet = getTable(ss, "members");
  var tenuresSheet = getTable(ss, "tenures");
  var tasksSheet = getTable(ss, "tasks");

  if (!membersSheet || !tenuresSheet) {
    return { success: false, message: "Required sheets not found" };
  }

  var todayDate = new Date();
  todayDate.setHours(0, 0, 0, 0);

  // -------------------------------------------------------------
  // Dynamic header resolution for tbl_tenures
  // -------------------------------------------------------------
  var tLastCol = Math.max(typeof tenuresSheet.getLastColumn === "function" ? (tenuresSheet.getLastColumn() || 0) : 0, 7);
  var tHeaders = tenuresSheet.getRange(1, 1, 1, tLastCol).getValues()[0];
  var tColMember = 1;
  var tColRole = 2;
  var tColDept = 3;
  var tColStart = 4;
  var tColEnd = 5;
  var tColType = 6;
  var tColFte = 7;

  for (var th = 0; th < tHeaders.length; th++) {
    var thName = String(tHeaders[th] || "").trim().toLowerCase();
    if (thName === "member") tColMember = th + 1;
    else if (thName === "role") tColRole = th + 1;
    else if (thName === "department") tColDept = th + 1;
    else if (thName === "start_date") tColStart = th + 1;
    else if (thName === "end_date") tColEnd = th + 1;
    else if (thName === "contract_type") tColType = th + 1;
    else if (thName === "fte_rate") tColFte = th + 1;
  }

  // -------------------------------------------------------------
  // Dynamic header resolution for tbl_members
  // -------------------------------------------------------------
  var mLastCol = Math.max(typeof membersSheet.getLastColumn === "function" ? (membersSheet.getLastColumn() || 0) : 0, 7);
  var mHeaders = membersSheet.getRange(1, 1, 1, mLastCol).getValues()[0];
  var mColMember = 1;
  var mColAccentureEmail = 2;
  var mColPhone = 3;
  var mColRole = 4;
  var mColDept = 5;
  var mColStatus = 6;
  var mColPresence = 7;

  for (var mh = 0; mh < mHeaders.length; mh++) {
    var mhName = String(mHeaders[mh] || "").trim().toLowerCase();
    if (mhName === "member") mColMember = mh + 1;
    else if (mhName === "accenture_email" || mhName === "accenture email") mColAccentureEmail = mh + 1;
    else if (mhName === "phone_number" || mhName === "phone") mColPhone = mh + 1;
    else if (mhName === "current_role" || mhName === "role") mColRole = mh + 1;
    else if (mhName === "current_department" || mhName === "department") mColDept = mh + 1;
    else if (mhName === "tenure_status" || mhName === "status") mColStatus = mh + 1;
    else if (mhName === "today_presence" || mhName === "today") mColPresence = mh + 1;
  }

  // -------------------------------------------------------------
  // Dynamic header resolution for tbl_tasks
  // -------------------------------------------------------------
  var tkColId = 1;
  var tkColProj = 2;
  var tkColMember = 3;
  var tkColTitle = 4;
  var tkColDesc = 0;
  var tkColDue = 5;
  var tkColStatus = 6;
  var tkColGTask = 7;

  if (tasksSheet) {
    var tkLastCol = Math.max(typeof tasksSheet.getLastColumn === "function" ? (tasksSheet.getLastColumn() || 0) : 0, 8);
    var tkHeaders = tasksSheet.getRange(1, 1, 1, tkLastCol).getValues()[0];
    for (var tkh = 0; tkh < tkHeaders.length; tkh++) {
      // Normalised so technical keys ("due_date") and readable labels ("Due date") both match
      var tkhName = String(tkHeaders[tkh] || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
      if (tkhName === "taskid" || tkhName === "id") tkColId = tkh + 1;
      else if (tkhName === "projectid") tkColProj = tkh + 1;
      else if (tkhName === "assignedmember" || tkhName === "assignedto") tkColMember = tkh + 1;
      else if (tkhName === "title") tkColTitle = tkh + 1;
      else if (tkhName === "description") tkColDesc = tkh + 1;
      else if (tkhName === "duedate") tkColDue = tkh + 1;
      else if (tkhName === "status") tkColStatus = tkh + 1;
      else if (tkhName === "googletaskid") tkColGTask = tkh + 1;
    }
  }

  // -------------------------------------------------------------
  // 1. Read tbl_tenures
  // -------------------------------------------------------------
  var tenuresMaxRows = readRowLimit(tenuresSheet, 300);
  var tenuresByMember = {}; // memberKey -> array of tenure records
  var uniqueTenureMembers = [];
  var validationWarnings = [];

  if (tenuresMaxRows > 1) {
    var tRange = tenuresSheet.getRange(2, 1, tenuresMaxRows - 1, tLastCol);
    var tValues = tRange.getValues();
    var tRich = (typeof tRange.getRichTextValues === "function") ? tRange.getRichTextValues() : null;

    for (var t = 0; t < tValues.length; t++) {
      var rawTVal = tValues[t][tColMember - 1];
      var rawTRich = (tRich && tRich[t]) ? tRich[t][tColMember - 1] : null;
      var tMember = safeExtractCell(rawTVal, rawTRich);
      if (!tMember) continue;

      // Check required fields for this tenure row (end_date is nullable)
      if (typeof validateRow === "function") {
        var vRowRes = validateRow("tenures", tValues[t], tHeaders);
        if (!vRowRes.valid) {
          validationWarnings.push({
            row: t + 2,
            member: tMember,
            missingColumns: vRowRes.missingColumns
          });
        }
      }

      var mKey = tMember.toLowerCase();
      if (!tenuresByMember[mKey]) {
        tenuresByMember[mKey] = [];
        uniqueTenureMembers.push(tMember);
      }

      var sDate = parseDateValue(tValues[t][tColStart - 1]);
      var eDate = parseDateValue(tValues[t][tColEnd - 1]);

      tenuresByMember[mKey].push({
        row: t + 2,
        member: tMember,
        role: safeExtractCell(tValues[t][tColRole - 1]),
        department: safeExtractCell(tValues[t][tColDept - 1]),
        startDate: sDate,
        endDate: eDate,
        contractType: safeExtractCell(tValues[t][tColType - 1]),
        fteRate: tValues[t][tColFte - 1]
      });
    }
  }

  // Helper to evaluate member tenure stats
  function evaluateTenureProfile(mKey) {
    var records = tenuresByMember[mKey];
    if (!records || records.length === 0) {
      return { role: "", department: "", status: "No Contract" };
    }

    // Active tenure: start <= today && (!end || end >= today)
    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      var startOk = !r.startDate || r.startDate <= todayDate;
      var endOk = !r.endDate || r.endDate >= todayDate;
      if (startOk && endOk) {
        return {
          role: r.role || "Team Member",
          department: r.department || "Delivery",
          status: "Active"
        };
      }
    }

    // Upcoming tenure: start > today
    for (var j = 0; j < records.length; j++) {
      if (records[j].startDate && records[j].startDate > todayDate) {
        return {
          role: records[j].role || "Team Member",
          department: records[j].department || "Delivery",
          status: "Upcoming"
        };
      }
    }

    // Past tenure: Departed
    var lastRecord = records[records.length - 1];
    return {
      role: lastRecord.role || "Team Member",
      department: lastRecord.department || "Delivery",
      status: "Departed"
    };
  }

  // -------------------------------------------------------------
  // 2. Read tbl_members
  // -------------------------------------------------------------
  var membersMaxRows = readRowLimit(membersSheet, 300);
  var existingMembers = [];
  var existingMemberMap = {};

  if (membersMaxRows > 1) {
    var mRange = membersSheet.getRange(2, 1, membersMaxRows - 1, mLastCol);
    var mValues = mRange.getValues();
    var mRich = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : null;

    for (var m = 0; m < mValues.length; m++) {
      var rawMVal = mValues[m][mColMember - 1];
      var rawMRich = (mRich && mRich[m]) ? mRich[m][mColMember - 1] : null;
      var mVal = safeExtractCell(rawMVal, rawMRich);
      if (!mVal) continue;

      var key = mVal.toLowerCase();
      var profile = evaluateTenureProfile(key);

      existingMembers.push({
        row: m + 2,
        member: mVal,
        accentureEmail: safeExtractCell(mValues[m][mColAccentureEmail - 1]),
        phoneNumber: safeExtractCell(mValues[m][mColPhone - 1]),
        currentRole: profile.role || safeExtractCell(mValues[m][mColRole - 1]),
        currentDepartment: profile.department || safeExtractCell(mValues[m][mColDept - 1]),
        tenureStatus: profile.status,
        todayPresence: safeExtractCell(mValues[m][mColPresence - 1]) || "Not Logged"
      });

      existingMemberMap[key] = true;
    }
  }

  // -------------------------------------------------------------
  // 3. Update existing rows in tbl_members with latest tenure lookups
  // -------------------------------------------------------------
  for (var u = 0; u < existingMembers.length; u++) {
    var em = existingMembers[u];
    membersSheet.getRange(em.row, mColRole).setValue(em.currentRole);
    membersSheet.getRange(em.row, mColDept).setValue(em.currentDepartment);
    membersSheet.getRange(em.row, mColStatus).setValue(em.tenureStatus);
  }

  // -------------------------------------------------------------
  // 4. Add new members from tbl_tenures to tbl_members
  //    Sequential row insertion prevents overwriting existing rows.
  // -------------------------------------------------------------
  var addedToDirectory = 0;
  var newlyAddedMembers = [];
  var nextMemberRow = findNextAvailableRow(membersSheet);

  for (var ut = 0; ut < uniqueTenureMembers.length; ut++) {
    var rawName = uniqueTenureMembers[ut];
    var uKey = rawName.toLowerCase();

    if (!existingMemberMap[uKey]) {
      var targetRow = nextMemberRow;
      nextMemberRow++; // Always advance for next member

      var prof = evaluateTenureProfile(uKey);
      var autoEmail = rawName.indexOf("@") !== -1 ? rawName : "";
      var srcTenureRow = tenuresByMember[uKey][0].row;

      // Copy exact People Chip from tbl_tenures to tbl_members
      copyCellChip(
        tenuresSheet,
        srcTenureRow,
        tColMember,
        membersSheet,
        targetRow,
        mColMember,
        autoEmail || rawName,
        ss.getId ? ss.getId() : ""
      );

      // Write remaining columns
      membersSheet.getRange(targetRow, mColAccentureEmail).setValue(autoEmail);
      membersSheet.getRange(targetRow, mColPhone).setValue("");
      membersSheet.getRange(targetRow, mColRole).setValue(prof.role);
      membersSheet.getRange(targetRow, mColDept).setValue(prof.department);
      membersSheet.getRange(targetRow, mColStatus).setValue(prof.status);
      membersSheet.getRange(targetRow, mColPresence).setValue("Not Logged");

      applyTypography(membersSheet.getRange(targetRow, 1, 1, mLastCol));

      existingMemberMap[uKey] = true;
      existingMembers.push({
        row: targetRow,
        member: rawName,
        currentRole: prof.role,
        currentDepartment: prof.department,
        tenureStatus: prof.status
      });

      newlyAddedMembers.push({
        member: rawName,
        memberRow: targetRow
      });

      addedToDirectory++;
    }
  }

  // -------------------------------------------------------------
  // 5. Assign Onboarding Task in tbl_tasks for new members
  // -------------------------------------------------------------
  var tasksCreated = 0;
  if (tasksSheet && newlyAddedMembers.length > 0) {
    try {
      var taskDueDate = new Date();
      taskDueDate.setDate(taskDueDate.getDate() + 7);
      var dueDateStr = safeFormatDateIso(taskDueDate);
      var nextTaskRow = findNextAvailableRow(tasksSheet);

      for (var nm = 0; nm < newlyAddedMembers.length; nm++) {
        var newM = newlyAddedMembers[nm];
        var targetTaskRow = nextTaskRow;
        nextTaskRow++;

        var taskId = "task_" + Math.random().toString(36).substring(2, 9);

        // Columns: task_id, project_id
        tasksSheet.getRange(targetTaskRow, tkColId).setValue(taskId);
        tasksSheet.getRange(targetTaskRow, tkColProj).setValue("ONBOARDING");

        // Copy People Chip to assigned_member
        copyCellChip(
          membersSheet,
          newM.memberRow,
          mColMember,
          tasksSheet,
          targetTaskRow,
          tkColMember,
          newM.member,
          ss.getId ? ss.getId() : ""
        );

        tasksSheet.getRange(targetTaskRow, tkColTitle).setValue(
          "Complete profile: enter Accenture email and phone number in Team Directory"
        );
        if (tkColDesc > 0 && tkColDesc <= tasksSheet.getLastColumn()) {
          tasksSheet.getRange(targetTaskRow, tkColDesc).setValue(
            "Please enter your Accenture email address and phone number in the Members tab (columns Accenture email and Phone)."
          );
        }
        tasksSheet.getRange(targetTaskRow, tkColDue).setValue(dueDateStr);
        tasksSheet.getRange(targetTaskRow, tkColStatus).setValue("Todo");
        // Task type (ADR 0045), so onboarding shows in the Calendar
        if (typeof SCHEMA_SPEC !== "undefined" && SCHEMA_SPEC.tasks && SCHEMA_SPEC.tasks.columns.indexOf("task_type") !== -1) {
          tasksSheet.getRange(targetTaskRow, SCHEMA_SPEC.tasks.columns.indexOf("task_type") + 1).setValue("Onboarding");
        }
        if (tkColGTask <= tasksSheet.getLastColumn()) {
          tasksSheet.getRange(targetTaskRow, tkColGTask).setValue("");
        }

        applyTypography(tasksSheet.getRange(targetTaskRow, 1, 1, Math.max(tasksSheet.getLastColumn(), 8)));
        tasksCreated++;
      }
    } catch (eTasks) {
      Logger.log("Onboarding task creation note: " + eTasks);
    }
  }

  // -------------------------------------------------------------
  // 5b. Synchronize People Chips into tbl_presence_baselines
  // -------------------------------------------------------------
  try {
    if (!fromEdit) syncBaselinesPeopleChips(ss);
  } catch (chipErr) {
    Logger.log("syncBaselinesPeopleChips note: " + chipErr);
  }

  // -------------------------------------------------------------
  // 6. Ensure column protections are applied
  // -------------------------------------------------------------
  if (typeof applyProtections === "function" && !fromEdit) {
    try {
      applyProtections(ss);
    } catch (protErr) {
      Logger.log("Protection apply note: " + protErr);
    }
  }

  return {
    success: true,
    addedToDirectory: addedToDirectory,
    tasksCreated: tasksCreated,
    updatedDirectoryRows: existingMembers.length,
    validationWarnings: validationWarnings
  };
}

/**
 * Copies a People Chip cell natively using SpreadsheetApp Range.copyTo
 * with Sheets API copyPaste and chipRuns as fallbacks.
 *
 * Supports both signatures:
 * - copyCellChip(srcSheet, srcRow, srcCol, dstSheet, dstRow, dstCol, fallbackEmail, spreadsheetId)
 * - copyCellChip(spreadsheetId, srcSheetId, srcRow, srcCol, dstSheetId, dstRow, dstCol, fallbackEmail)
 */
function copyCellChip(a1, a2, a3, a4, a5, a6, a7, a8) {
  var srcSheet, srcRow, srcCol, dstSheet, dstRow, dstCol, fallbackEmail, spreadsheetId;

  if (a1 && typeof a1 === "object" && typeof a1.getRange === "function") {
    // New signature: (srcSheet, srcRow, srcCol, dstSheet, dstRow, dstCol, fallbackEmail, spreadsheetId)
    srcSheet = a1;
    srcRow = a2;
    srcCol = a3;
    dstSheet = a4;
    dstRow = a5;
    dstCol = a6;
    fallbackEmail = a7;
    spreadsheetId = a8;
  } else {
    // Legacy signature: (spreadsheetId, srcSheetId, srcRow, srcCol, dstSheetId, dstRow, dstCol, fallbackEmail)
    spreadsheetId = a1;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) {
      var allSheets = ss.getSheets();
      for (var s = 0; s < allSheets.length; s++) {
        if (allSheets[s].getSheetId() === a2) srcSheet = allSheets[s];
        if (allSheets[s].getSheetId() === a5) dstSheet = allSheets[s];
      }
    }
    srcRow = a3;
    srcCol = a4;
    dstRow = a6;
    dstCol = a7;
    fallbackEmail = a8;
  }

  var copied = false;

  // 1. Native SpreadsheetApp Range.copyTo (Copies Smart Chips & Rich Text directly)
  if (srcSheet && dstSheet && srcRow && srcCol && dstRow && dstCol) {
    try {
      var srcRange = srcSheet.getRange(srcRow, srcCol);
      var dstRange = dstSheet.getRange(dstRow, dstCol);
      srcRange.copyTo(dstRange);
      var checkVal = safeExtractCell(dstRange.getValue(), typeof dstRange.getRichTextValue === "function" ? dstRange.getRichTextValue() : null);
      if (checkVal) {
        copied = true;
      }
    } catch (errNative) {
      Logger.log("Range.copyTo note: " + errNative);
    }
  }

  // 2. Sheets API batchUpdate copyPaste fallback
  if (!copied && typeof Sheets !== "undefined" && Sheets.Spreadsheets && spreadsheetId && srcSheet && dstSheet) {
    try {
      var copyReq = {
        copyPaste: {
          source: {
            sheetId: srcSheet.getSheetId(),
            startRowIndex: srcRow - 1,
            endRowIndex: srcRow,
            startColumnIndex: srcCol - 1,
            endColumnIndex: srcCol
          },
          destination: {
            sheetId: dstSheet.getSheetId(),
            startRowIndex: dstRow - 1,
            endRowIndex: dstRow,
            startColumnIndex: dstCol - 1,
            endColumnIndex: dstCol
          },
          pasteType: "PASTE_NORMAL"
        }
      };
      Sheets.Spreadsheets.batchUpdate({ requests: [copyReq] }, spreadsheetId);
      var checkVal2 = safeExtractCell(dstSheet.getRange(dstRow, dstCol).getValue(), typeof dstSheet.getRange(dstRow, dstCol).getRichTextValue === "function" ? dstSheet.getRange(dstRow, dstCol).getRichTextValue() : null);
      if (checkVal2) {
        copied = true;
      }
    } catch (errApi) {
      Logger.log("Sheets API copyPaste note: " + errApi);
    }
  }

  // 3. Sheets API updateCells with chipRuns fallback
  if (!copied && typeof Sheets !== "undefined" && Sheets.Spreadsheets && spreadsheetId && dstSheet && fallbackEmail && fallbackEmail.indexOf("@") !== -1) {
    try {
      var chipReq = {
        updateCells: {
          range: {
            sheetId: dstSheet.getSheetId(),
            startRowIndex: dstRow - 1,
            endRowIndex: dstRow,
            startColumnIndex: dstCol - 1,
            endColumnIndex: dstCol
          },
          rows: [
            {
              values: [
                {
                  userEnteredValue: { stringValue: "@" },
                  chipRuns: [
                    {
                      startIndex: 0,
                      chip: {
                        personProperties: {
                          email: fallbackEmail
                        }
                      }
                    }
                  ]
                }
              ]
            }
          ],
          fields: "userEnteredValue,chipRuns"
        }
      };
      Sheets.Spreadsheets.batchUpdate({ requests: [chipReq] }, spreadsheetId);
      copied = true;
    } catch (errChip) {
      Logger.log("chipRuns note: " + errChip);
    }
  }

  // 4. Value fallback if still empty
  if (!copied && dstSheet && dstRow && dstCol && fallbackEmail) {
    try {
      dstSheet.getRange(dstRow, dstCol).setValue(fallbackEmail);
      copied = true;
    } catch (errVal) {
      Logger.log("setValue note: " + errVal);
    }
  }
}

/**
 * Finds the next available empty row across all columns of a sheet.
 * Prevents treating a row as empty if any column has data.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {number} 1-indexed row number
 */
function findNextAvailableRow(sheet) {
  var maxRows = sheet.getMaxRows();
  var maxCols = Math.max(sheet.getLastColumn(), 1);
  if (maxRows <= 1) return 2;

  // Scan to just past the last row with data (never stop at a fixed cap: Tasks grows)
  var checkLimit = Math.min(maxRows, (typeof readRowLimit === "function" ? readRowLimit(sheet, 500) : Math.min(maxRows, 500)) + 1);
  var range = sheet.getRange(2, 1, checkLimit - 1, maxCols);
  var values = range.getValues();
  var richValues = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : null;

  for (var r = 0; r < values.length; r++) {
    var hasContent = false;
    for (var c = 0; c < values[r].length; c++) {
      var rawVal = values[r][c];
      var rawRich = (richValues && richValues[r]) ? richValues[r][c] : null;
      var textVal = safeExtractCell(rawVal, rawRich);
      if (textVal !== "") {
        hasContent = true;
        break;
      }
    }
    if (!hasContent) {
      return r + 2;
    }
  }
  if (checkLimit + 1 > maxRows && typeof sheet.insertRowsAfter === "function") {
    try { sheet.insertRowsAfter(maxRows, 50); } catch (eIns) {}
  }
  return checkLimit + 1;
}

/**
 * Backwards compatibility helper.
 */
function findFirstEmptyRow(sheet, checkCol) {
  return findNextAvailableRow(sheet);
}

/**
 * Safe date parsing for cell values (Date object or YYYY-MM-DD string).
 */
function parseDateValue(val) {
  if (!val) return null;
  if (val instanceof Date) {
    var d = new Date(val.getTime());
    d.setHours(0, 0, 0, 0);
    return d;
  }
  if (typeof val === "string" && val.trim().length >= 10) {
    var parts = val.trim().split("-");
    if (parts.length === 3) {
      var parsed = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      parsed.setHours(0, 0, 0, 0);
      return parsed;
    }
  }
  return null;
}

/**
 * Synchronizes and hydrates People Chips into tbl_presence_baselines
 * from tbl_members (or tbl_tenures), ensuring that baseline rows display native People Chips.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {number} Count of baseline rows hydrated with People Chips
 */
function syncBaselinesPeopleChips(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return 0;
  var membersSheet = getTable(ss, "members");
  var baselinesSheet = getTable(ss, "baselines");
  if (!membersSheet || !baselinesSheet || typeof copyCellChip !== "function") return 0;

  var mMax = readRowLimit(membersSheet, 300);
  if (mMax <= 1) return 0;
  var mRange = membersSheet.getRange(2, 1, mMax - 1, Math.max(membersSheet.getLastColumn(), 6));
  var mVals = mRange.getValues();
  var mRich = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : [];

  var membersRoster = [];
  for (var m = 0; m < mVals.length; m++) {
    var rawM = mVals[m][0];
    var rM = (mRich && mRich[m]) ? mRich[m][0] : null;
    var nameM = (typeof extractCellValue === "function") ? extractCellValue(rawM, rM) : String(rawM).trim();
    if (!nameM) continue;
    membersRoster.push({
      row: m + 2,
      raw: rawM,
      rich: rM,
      name: nameM,
      nameLower: nameM.toLowerCase(),
      email: String(mVals[m][1] || "").trim().toLowerCase()
    });
  }

  if (membersRoster.length === 0) return 0;

  var bMax = readRowLimit(baselinesSheet, 500);
  if (bMax <= 1) return 0;
  var bRange = baselinesSheet.getRange(2, 2, bMax - 1, 1);
  var bVals = bRange.getValues();
  var bRich = (typeof bRange.getRichTextValues === "function") ? bRange.getRichTextValues() : [];

  var updatedCount = 0;
  for (var b = 0; b < bVals.length; b++) {
    var rawB = bVals[b][0];
    var rB = (bRich && bRich[b]) ? bRich[b][0] : null;
    var nameB = (typeof extractCellValue === "function") ? extractCellValue(rawB, rB) : String(rawB).trim();
    if (!nameB) continue;

    var nameBLower = nameB.toLowerCase();
    var matchedMember = null;

    for (var k = 0; k < membersRoster.length; k++) {
      var mem = membersRoster[k];
      if (mem.nameLower === nameBLower || (mem.email && mem.email === nameBLower)) {
        matchedMember = mem;
        break;
      }
      if (typeof isUserMemberMatch === "function") {
        if (isUserMemberMatch(mem.raw, mem.rich, nameB) || isUserMemberMatch(rawB, rB, mem.name)) {
          matchedMember = mem;
          break;
        }
      }
    }

    if (matchedMember) {
      copyCellChip(membersSheet, matchedMember.row, 1, baselinesSheet, b + 2, 2, matchedMember.name, ss.getId());
      updatedCount++;
    }
  }
  return updatedCount;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    syncMembersAndTenures: syncMembersAndTenures,
    syncBaselinesPeopleChips: syncBaselinesPeopleChips,
    copyCellChip: copyCellChip,
    findNextAvailableRow: findNextAvailableRow,
    findFirstEmptyRow: findFirstEmptyRow,
    parseDateValue: parseDateValue,
    safeExtractCell: safeExtractCell,
    safeFormatDateIso: safeFormatDateIso
  };
}
