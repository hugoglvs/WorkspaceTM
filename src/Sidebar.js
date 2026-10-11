/**
 * Project Management Workspace - Presence Baseline Sidebar Backend
 * Uses Google Workspace People Chips natively as unique Member identity.
 */

/**
 * Opens the Presence Baseline Sidebar in Google Sheets.
 */
function showBaselineSidebar() {
  var ui = SpreadsheetApp.getUi();
  var html = HtmlService.createHtmlOutputFromFile('BaselineSidebar')
    .setTitle('Presence & Leaves')
    .setWidth(350);
  ui.showSidebar(html);
}

/**
 * Safely extracts string / display value from a cell that might contain
 * text, numbers, dates, or Google Sheets smart chips.
 */
function extractCellValue(val, richText) {
  if (richText && typeof richText.getText === "function") {
    var rtText = richText.getText();
    if (rtText) return String(rtText).trim();
  }
  if (val === null || val === undefined) return "";
  if (typeof val === "string") return val.trim();
  if (typeof val === "object") {
    if (val instanceof Date) return formatDateIso(val);
    if (typeof val.getEmail === "function") return String(val.getEmail()).trim();
    if (typeof val.getName === "function") return String(val.getName()).trim();
    if (val.email) return String(val.email).trim();
    if (val.name) return String(val.name).trim();
  }
  return String(val).trim();
}

/**
 * Normalizes a person's name or email prefix by lowercasing, stripping diacritics / accents,
 * and turning punctuation into spaces for robust matching.
 * e.g. "Hugo Gonçalves" -> "hugo goncalves", "hugo.goncalves" -> "hugo goncalves"
 *
 * @param {string} str
 * @return {string}
 */
function normalizePersonName(str) {
  if (!str) return "";
  var s = String(str).toLowerCase().trim();
  try {
    if (typeof s.normalize === "function") {
      s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    }
  } catch (eNorm) {}
  return s.replace(/[._\-+]/g, " ").replace(/\s+/g, " ").trim();
}

/**
 * Safely extracts an email address from a People Chip cell or object.
 * Checks:
 * 1. Smart Chip object methods (val.getEmail(), val.email, val.getPersonEmail(), personProperties)
 * 2. RichText mailto: link URL or run URLs (Standard Google Sheets People chip)
 * 3. String containing an email address
 *
 * @param {*} val
 * @param {GoogleAppsScript.Spreadsheet.RichTextValue} [richText]
 * @return {string}
 */
function extractEmailFromChip(val, richText) {
  // 1. Direct object inspection (Google People chip object / Person class)
  if (val && typeof val === "object") {
    if (typeof val.getEmail === "function") {
      var em1 = val.getEmail();
      if (em1) return String(em1).trim();
    }
    if (typeof val.getPersonEmail === "function") {
      var em2 = val.getPersonEmail();
      if (em2) return String(em2).trim();
    }
    if (val.email) return String(val.email).trim();
    if (val.mail) return String(val.mail).trim();
    if (val.personEmail) return String(val.personEmail).trim();
    if (val.personProperties && val.personProperties.email) {
      return String(val.personProperties.email).trim();
    }
  }

  // Helper to extract email from url or string
  function parseEmailFromUrl(url) {
    if (!url || typeof url !== "string") return "";
    var cleanUrl = url.trim();
    if (cleanUrl.toLowerCase().indexOf("mailto:") === 0) {
      return cleanUrl.substring(7).split("?")[0].trim();
    }
    var mailtoMatch = cleanUrl.match(/mailto:([^\s>?&]+)/i);
    if (mailtoMatch) return mailtoMatch[1];
    var emailRegexMatch = cleanUrl.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (emailRegexMatch) return emailRegexMatch[1];
    return "";
  }

  // 2. RichText mailto: link URL or runs (Standard Google Sheets People Chip representation)
  if (richText) {
    if (typeof richText.getLinkUrl === "function") {
      var emLink = parseEmailFromUrl(richText.getLinkUrl());
      if (emLink) return emLink;
    }
    if (typeof richText.getRuns === "function") {
      var runs = richText.getRuns() || [];
      for (var r = 0; r < runs.length; r++) {
        var rLink = runs[r].getLinkUrl ? runs[r].getLinkUrl() : "";
        var emRun = parseEmailFromUrl(rLink);
        if (emRun) return emRun;
      }
    }
  }

  // 3. String containing an email address
  if (typeof val === "string") {
    var str = val.trim();
    var emStr = parseEmailFromUrl(str);
    if (emStr) return emStr;
  }

  return "";
}

/**
 * Checks whether sub is a subsequence of str in order.
 * @param {string} sub
 * @param {string} str
 * @return {boolean}
 */
function isSubsequence(sub, str) {
  if (!sub || !str) return false;
  var sIdx = 0;
  for (var i = 0; i < str.length; i++) {
    if (str[i] === sub[sIdx]) {
      sIdx++;
      if (sIdx === sub.length) return true;
    }
  }
  return sIdx === sub.length;
}

/**
 * Reads People Smart Chip metadata (emails) directly from Google Sheets API v4.
 * Native Google People Chips store email in cell `chipRuns.chip.personProperties.email`.
 *
 * @param {string} spreadsheetId
 * @param {string} sheetName
 * @param {number} startRow
 * @param {number} numRows
 * @return {Object.<number, string>}
 */
function fetchMemberChipEmails(spreadsheetId, sheetName, startRow, numRows) {
  var chipEmails = {};
  if (typeof Sheets === "undefined" || !Sheets.Spreadsheets || !spreadsheetId) {
    return chipEmails;
  }
  try {
    var endRow = startRow + numRows - 1;
    var a1Range = (sheetName || "tbl_members") + "!A" + startRow + ":A" + endRow;
    var resp = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: [a1Range],
      fields: "sheets.data.rowData.values(formattedValue,chipRuns)"
    });
    if (resp && resp.sheets && resp.sheets[0] && resp.sheets[0].data && resp.sheets[0].data[0] && resp.sheets[0].data[0].rowData) {
      var rowData = resp.sheets[0].data[0].rowData;
      for (var r = 0; r < rowData.length; r++) {
        var rowIdx = startRow + r;
        var cell = (rowData[r] && rowData[r].values) ? rowData[r].values[0] : null;
        if (cell && cell.chipRuns && cell.chipRuns.length > 0) {
          for (var c = 0; c < cell.chipRuns.length; c++) {
            var chip = cell.chipRuns[c].chip;
            if (chip && chip.personProperties && chip.personProperties.email) {
              chipEmails[rowIdx] = String(chip.personProperties.email).trim().toLowerCase();
              break;
            }
          }
        }
      }
    }
  } catch (err) {
    try {
      Logger.log("[fetchMemberChipEmails] Note: Sheets API fetch skipped (" + (err ? err.message : err) + ")");
    } catch (eLog) {}
  }
  return chipEmails;
}

/**
 * Analyzes and explains whether a People Chip cell corresponds to the active user's identity.
 * Strictly compares the People Chip itself (never compares to accenture_email).
 *
 * @param {*} memberVal
 * @param {GoogleAppsScript.Spreadsheet.RichTextValue} [memberRichText]
 * @param {string} userEmail
 * @param {string} [overrideChipEmail]
 * @return {{match: boolean, reason: string, chipEmail: string, memberStr: string}}
 */
function explainUserMemberMatch(memberVal, memberRichText, userEmail, overrideChipEmail) {
  if (!userEmail) {
    return { match: false, reason: "No active user email provided", chipEmail: "", memberStr: "" };
  }
  var userEmailLower = String(userEmail).trim().toLowerCase();
  if (!userEmailLower) {
    return { match: false, reason: "Empty active user email", chipEmail: "", memberStr: "" };
  }

  // 1. People Chip extracted email (from Sheets API chip metadata, mailto link, or chip object)
  var chipEmail = overrideChipEmail || extractEmailFromChip(memberVal, memberRichText);
  if (chipEmail) {
    if (chipEmail.toLowerCase() === userEmailLower) {
      return { match: true, reason: "People Chip email exact match ('" + chipEmail + "')", chipEmail: chipEmail, memberStr: "" };
    }
  }

  // 2. Member string contains email
  var memberStr = extractCellValue(memberVal, memberRichText);
  var memberStrLower = memberStr.toLowerCase();
  if (memberStrLower && memberStrLower.indexOf(userEmailLower) !== -1) {
    return { match: true, reason: "Member text contains user email ('" + memberStr + "')", chipEmail: chipEmail, memberStr: memberStr };
  }

  // 3. Normalized Name <-> Email prefix comparison
  var emailPrefix = userEmailLower.split("@")[0];
  var normEmailPrefix = normalizePersonName(emailPrefix);
  var normMemberName = normalizePersonName(memberStr);

  if (normEmailPrefix && normMemberName) {
    if (normMemberName === normEmailPrefix) {
      return { match: true, reason: "Normalized display name matches email prefix ('" + normMemberName + "' == '" + normEmailPrefix + "')", chipEmail: chipEmail, memberStr: memberStr };
    }
    if (normMemberName.replace(/\s+/g, "") === normEmailPrefix.replace(/\s+/g, "")) {
      return { match: true, reason: "Collapsed display name matches collapsed email prefix ('" + normMemberName.replace(/\s+/g, "") + "')", chipEmail: chipEmail, memberStr: memberStr };
    }

    var cleanPrefix = normEmailPrefix.replace(/[0-9]/g, "").replace(/\s+/g, "");
    var cleanMember = normMemberName.replace(/\s+/g, "");
    if (cleanPrefix && cleanMember && cleanPrefix === cleanMember) {
      return { match: true, reason: "Prefix (ignoring trailing digits) matches member name ('" + cleanPrefix + "' == '" + cleanMember + "')", chipEmail: chipEmail, memberStr: memberStr };
    }

    var emailTokens = normEmailPrefix.split(" ").filter(Boolean);
    var nameTokens = normMemberName.split(" ").filter(Boolean);
    if (emailTokens.length >= 2 && nameTokens.length >= 2) {
      var allEmailTokensInName = emailTokens.every(function(t) { return normMemberName.indexOf(t) !== -1; });
      if (allEmailTokensInName) {
        return { match: true, reason: "All email tokens found in member name", chipEmail: chipEmail, memberStr: memberStr };
      }
      var allNameTokensInEmail = nameTokens.every(function(t) { return normEmailPrefix.indexOf(t) !== -1; });
      if (allNameTokensInEmail) {
        return { match: true, reason: "All member name tokens found in email prefix", chipEmail: chipEmail, memberStr: memberStr };
      }
    }

    // Name tokens contraction / abbreviation matching (e.g. "hugoglvs42" <-> "Hugo Gonçalves")
    if (nameTokens.length >= 2) {
      var firstName = nameTokens[0];
      var lastName = nameTokens[nameTokens.length - 1];
      // Case 0: cleanPrefix exactly matches first name (e.g. "alice" in "Alice Smith")
      if (cleanPrefix === firstName) {
        return { match: true, reason: "Email prefix matches first name ('" + cleanPrefix + "')", chipEmail: chipEmail, memberStr: memberStr };
      }
      // Case A: cleanPrefix starts with firstName and remainder is a subsequence of lastName (e.g. "hugo" + "glvs" in "goncalves")
      if (cleanPrefix.indexOf(firstName) === 0) {
        var remainder = cleanPrefix.substring(firstName.length);
        if (remainder && remainder.length >= 2 && isSubsequence(remainder, lastName)) {
          return { match: true, reason: "Name contraction match: first name '" + firstName + "' with last name contraction '" + remainder + "' in '" + lastName + "'", chipEmail: chipEmail, memberStr: memberStr };
        }
      }
      // Case B: cleanPrefix starts with first letter of firstName and rest is lastName (e.g. "hgoncalves" in "Hugo Gonçalves")
      if (cleanPrefix.indexOf(firstName[0]) === 0 && cleanPrefix.substring(1) === lastName) {
        return { match: true, reason: "Initial + last name match ('" + cleanPrefix + "')", chipEmail: chipEmail, memberStr: memberStr };
      }
      // Case C: cleanPrefix is lastName + initial (e.g. "goncalvesh")
      if (cleanPrefix.indexOf(lastName) === 0 && cleanPrefix.substring(lastName.length) === firstName[0]) {
        return { match: true, reason: "Last name + initial match ('" + cleanPrefix + "')", chipEmail: chipEmail, memberStr: memberStr };
      }
    }
  }

  return {
    match: false,
    reason: "No match (chipEmail='" + chipEmail + "', memberName='" + memberStr + "', userEmail='" + userEmailLower + "')",
    chipEmail: chipEmail,
    memberStr: memberStr
  };
}

/**
 * Checks whether a People Chip cell corresponds to the active user's identity.
 * Strictly compares the People Chip itself (never compares to accenture_email).
 *
 * @param {*} memberVal
 * @param {GoogleAppsScript.Spreadsheet.RichTextValue} [memberRichText]
 * @param {string} userEmail
 * @param {string} [overrideChipEmail]
 * @return {boolean}
 */
function isUserMemberMatch(memberVal, memberRichText, userEmail, overrideChipEmail) {
  return explainUserMemberMatch(memberVal, memberRichText, userEmail, overrideChipEmail).match;
}

/**
 * Diagnostic audit of active user session and tbl_members identities.
 * Logs all details via Logger.log() and returns structured findings.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object}
 */
function auditUserAndMembersIdentity(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var logs = [];

  function logMsg(msg) {
    logs.push(msg);
    try {
      Logger.log(msg);
    } catch (e) {}
  }

  logMsg("=== [User & Member Identity Audit] ===");

  var activeEmail = "";
  var effectiveEmail = "";
  try {
    var aUser = Session.getActiveUser();
    activeEmail = aUser ? aUser.getEmail() : "";
  } catch (e1) {
    logMsg("[Session] Error calling getActiveUser().getEmail(): " + e1.message);
  }
  try {
    var eUser = Session.getEffectiveUser();
    effectiveEmail = eUser ? eUser.getEmail() : "";
  } catch (e2) {
    logMsg("[Session] Error calling getEffectiveUser().getEmail(): " + e2.message);
  }

  logMsg("[Session] Active User: '" + activeEmail + "'");
  logMsg("[Session] Effective User: '" + effectiveEmail + "'");

  if (!activeEmail) {
    logMsg("[Session] WARNING: Active user email is empty! In Google Sheets simple onOpen triggers, Google restricts getActiveUser().getEmail() unless authorized. Run 'Project Management > Debug User & Member Identity' from the menu for authorized context.");
  }

  if (!ss) {
    logMsg("[Spreadsheet] No active spreadsheet available.");
    return { activeEmail: activeEmail, effectiveEmail: effectiveEmail, matchingMember: null, baselinesCount: 0, memberCount: 0, logLines: logs, reportText: logs.join("\n") };
  }

  var membersSheet = ss.getSheetByName("tbl_members");
  if (!membersSheet) {
    logMsg("[tbl_members] Sheet 'tbl_members' NOT FOUND in spreadsheet!");
    return { activeEmail: activeEmail, effectiveEmail: effectiveEmail, matchingMember: null, baselinesCount: 0, memberCount: 0, logLines: logs, reportText: logs.join("\n") };
  }

  var lastRow = membersSheet.getLastRow();
  var maxRows = Math.min(membersSheet.getMaxRows(), 300);
  logMsg("[tbl_members] Sheet found. Last row: " + lastRow + ", Max rows examined: " + maxRows);

  var ssId = "";
  try {
    ssId = (ss && typeof ss.getId === "function") ? ss.getId() : "";
  } catch (eId) {}

  var memberChipEmails = (ssId && maxRows > 1) ? fetchMemberChipEmails(ssId, "tbl_members", 2, maxRows - 1) : {};

  var matchingMemberRow = null;
  var memberCount = 0;

  if (maxRows > 1) {
    var range = membersSheet.getRange(2, 1, maxRows - 1, Math.max(membersSheet.getLastColumn(), 7));
    var values = range.getValues();
    var richTexts = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];

    for (var i = 0; i < values.length; i++) {
      var rowNum = i + 2;
      var rawVal = values[i][0];
      var rawRich = (richTexts && richTexts[i]) ? richTexts[i][0] : null;

      if (rawVal === "" && (!rawRich || !rawRich.getText || rawRich.getText() === "")) {
        continue;
      }
      memberCount++;

      var valType = typeof rawVal;
      var richLink = rawRich && typeof rawRich.getLinkUrl === "function" ? rawRich.getLinkUrl() : "";
      var richRuns = rawRich && typeof rawRich.getRuns === "function" ? rawRich.getRuns() : [];
      var runUrls = [];
      for (var r = 0; r < richRuns.length; r++) {
        var rLink = richRuns[r].getLinkUrl ? richRuns[r].getLinkUrl() : "";
        if (rLink) runUrls.push(rLink);
      }

      var extractedChipEmail = memberChipEmails[rowNum] || extractEmailFromChip(rawVal, rawRich);
      var displayName = extractCellValue(rawVal, rawRich);
      var matchInfo = explainUserMemberMatch(rawVal, rawRich, activeEmail, extractedChipEmail);

      var details = "[Row " + rowNum + "] " +
        "Name: '" + displayName + "' | " +
        "ValType: " + valType + " | " +
        "ExtractedEmail: '" + extractedChipEmail + "' | " +
        "LinkUrl: '" + (richLink || "") + "' | " +
        (runUrls.length > 0 ? "RunUrls: [" + runUrls.join(", ") + "] | " : "") +
        "Match: " + (matchInfo.match ? "YES (" + matchInfo.reason + ")" : "NO (" + matchInfo.reason + ")");

      logMsg(details);

      if (matchInfo.match && !matchingMemberRow) {
        matchingMemberRow = {
          row: rowNum,
          name: displayName,
          email: extractedChipEmail,
          role: values[i][3],
          department: values[i][4],
          todayPresence: values[i][6]
        };
      }
    }
  }

  logMsg("[tbl_members] Total members examined: " + memberCount + ". Active user match: " + (matchingMemberRow ? "FOUND on row " + matchingMemberRow.row : "NOT FOUND"));

  // Baselines check
  var baselinesSheet = ss.getSheetByName("tbl_presence_ranges");
  var baselinesCount = 0;
  if (baselinesSheet && matchingMemberRow) {
    var bMax = Math.min(baselinesSheet.getMaxRows(), 500);
    if (bMax > 1) {
      var bChipEmails = ssId ? fetchMemberChipEmails(ssId, "tbl_presence_ranges", 2, bMax - 1) : {};
      var bRange = baselinesSheet.getRange(2, 1, bMax - 1, 3);
      var bVals = bRange.getValues();
      var bRich = (typeof bRange.getRichTextValues === "function") ? bRange.getRichTextValues() : [];
      for (var b = 0; b < bVals.length; b++) {
        if (String(bVals[b][2] || "").trim() === "Leave") continue;
        var bVal = bVals[b][1];
        var bR = (bRich && bRich[b]) ? bRich[b][1] : null;
        var bExtEmail = bChipEmails[b + 2] || extractEmailFromChip(bVal, bR);
        var bMatch = (bExtEmail && bExtEmail.toLowerCase() === activeEmail.toLowerCase()) ||
                     explainUserMemberMatch(bVal, bR, activeEmail, bExtEmail).match ||
                     (matchingMemberRow.name && extractCellValue(bVal, bR).toLowerCase() === matchingMemberRow.name.toLowerCase());
        if (bMatch) {
          baselinesCount++;
        }
      }
    }
  }
  logMsg("[tbl_presence_ranges] Baselines configured for user: " + baselinesCount);

  // Tasks check in tbl_tasks
  var tasksSheet = ss.getSheetByName("tbl_tasks");
  var userTasksCount = 0;
  var otherTasksCount = 0;
  var misassignedGTasksCount = 0;
  var tasksExamined = 0;
  var taskAuditDetails = [];

  if (tasksSheet) {
    var tLast = tasksSheet.getLastRow();
    var tMax = Math.min(tasksSheet.getMaxRows(), 500);
    logMsg("[tbl_tasks] Sheet found. Last row: " + tLast + ", Max rows examined: " + tMax);

    if (tMax > 1) {
      var tColMap = (typeof getTaskColumnMap === "function")
        ? getTaskColumnMap(tasksSheet)
        : (function() {
            var m = { taskId: 1, projectId: 2, assignedMember: 3, title: 4, description: 0, dueDate: 5, status: 6, googleTaskId: 7 };
            if (tasksSheet) {
              var hdr = tasksSheet.getRange(1, 1, 1, Math.max(tasksSheet.getLastColumn(), 8)).getValues()[0];
              for (var ci = 0; ci < hdr.length; ci++) {
                var cName = String(hdr[ci] || "").trim().toLowerCase();
                if (cName === "task_id") m.taskId = ci + 1;
                else if (cName === "project_id") m.projectId = ci + 1;
                else if (cName === "assigned_member") m.assignedMember = ci + 1;
                else if (cName === "title") m.title = ci + 1;
                else if (cName === "description") m.description = ci + 1;
                else if (cName === "due_date") m.dueDate = ci + 1;
                else if (cName === "status") m.status = ci + 1;
                else if (cName === "google_task_id") m.googleTaskId = ci + 1;
              }
            }
            return m;
          })();

      var tChipEmails = ssId ? fetchMemberChipEmails(ssId, "tbl_tasks", 2, tMax - 1) : {};
      var tRange = tasksSheet.getRange(2, 1, tMax - 1, Math.max(tasksSheet.getLastColumn(), 8));
      var tVals = tRange.getValues();
      var tRich = (typeof tRange.getRichTextValues === "function") ? tRange.getRichTextValues() : [];

      for (var t = 0; t < tVals.length; t++) {
        var tRow = t + 2;
        var taskId = String(tVals[t][tColMap.taskId - 1] || "").trim();
        var projId = String(tVals[t][tColMap.projectId - 1] || "").trim();
        var rawAssigned = tVals[t][tColMap.assignedMember - 1];
        var richAssigned = (tRich && tRich[t]) ? tRich[t][tColMap.assignedMember - 1] : null;
        var taskTitle = String(tVals[t][tColMap.title - 1] || "").trim();
        var taskDesc = (tColMap.description > 0) ? String(tVals[t][tColMap.description - 1] || "").trim() : "";
        var taskStatus = String(tVals[t][tColMap.status - 1] || "").trim();
        var gTaskId = String(tVals[t][tColMap.googleTaskId - 1] || "").trim();

        if (!taskId && !projId && !taskTitle && !rawAssigned) continue;
        tasksExamined++;

        var assignedName = extractCellValue(rawAssigned, richAssigned);
        var tExtEmail = tChipEmails[tRow] || extractEmailFromChip(rawAssigned, richAssigned);
        var matchRes = explainUserMemberMatch(rawAssigned, richAssigned, activeEmail, tExtEmail);

        var isUserTask = matchRes.match;
        if (!isUserTask && matchingMemberRow && matchingMemberRow.name && assignedName) {
          if (assignedName.toLowerCase() === matchingMemberRow.name.toLowerCase()) {
            isUserTask = true;
          }
        }

        var tDetails = "[Task Row " + tRow + "] " +
          "ID: '" + taskId + "' | " +
          "Title: '" + taskTitle + "'" + (taskDesc ? " (Desc: '" + taskDesc + "')" : "") + " | " +
          "Project: '" + projId + "' | " +
          "Assigned: '" + assignedName + "' (Email: '" + tExtEmail + "') | " +
          "Match: " + (isUserTask ? "YES (" + matchRes.reason + ")" : "NO (" + matchRes.reason + ")") + " | " +
          "GoogleTaskId: '" + gTaskId + "' | " +
          "Status: '" + taskStatus + "'";

        logMsg(tDetails);

        if (isUserTask) {
          userTasksCount++;
        } else {
          otherTasksCount++;
          if (gTaskId) {
            logMsg("  -> ATTENTION: Task is assigned to '" + assignedName + "' (not active user), but has Google Task ID '" + gTaskId + "' (synced into your Google Tasks)!");
            misassignedGTasksCount++;
          }
        }

        taskAuditDetails.push({
          row: tRow,
          taskId: taskId,
          title: taskTitle,
          assigned: assignedName,
          assignedEmail: tExtEmail,
          isUserTask: isUserTask,
          googleTaskId: gTaskId,
          status: taskStatus
        });
      }
    }
    logMsg("[tbl_tasks] Total tasks examined: " + tasksExamined + ". Assigned to active user: " + userTasksCount + ". Assigned to other members: " + otherTasksCount + (misassignedGTasksCount > 0 ? " | WARNING: " + misassignedGTasksCount + " tasks for other members have Google Task IDs!" : ""));
  } else {
    logMsg("[tbl_tasks] Sheet 'tbl_tasks' NOT FOUND!");
  }

  logMsg("=== [End Audit] ===");

  return {
    activeEmail: activeEmail,
    effectiveEmail: effectiveEmail,
    matchingMember: matchingMemberRow,
    baselinesCount: baselinesCount,
    memberCount: memberCount,
    tasksExamined: tasksExamined,
    userTasksCount: userTasksCount,
    otherTasksCount: otherTasksCount,
    misassignedGTasksCount: misassignedGTasksCount,
    taskAuditDetails: taskAuditDetails,
    logLines: logs,
    reportText: logs.join("\n")
  };
}

/**
 * Interactive menu command: Audits user identity and displays diagnostic popup in Sheets.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object}
 */
function debugUserIdentity(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var audit = auditUserAndMembersIdentity(ss);
  try {
    var ui = SpreadsheetApp.getUi();
    if (ui) {
      var summary = audit.logLines.join("\n");
      if (audit.misassignedGTasksCount > 0) {
        summary += "\n\n⚠️ ATTENTION: " + audit.misassignedGTasksCount + " tasks assigned to other members currently have Google Task IDs and were synced to your personal Google Tasks. Run 'Project Management > Clean Up Foreign Google Tasks' from the menu to remove them from your personal list.";
      }
      if (summary.length > 1800) {
        summary = summary.substring(0, 1800) + "\n\n... (truncated, check Extensions > Apps Script > Executions for full logs)";
      }
      ui.alert("User & Member Identity Diagnostics", summary, ui.ButtonSet.OK);
    }
  } catch (eUi) {}
  return audit;
}

/**
 * Audits the active user's presence baseline and member profile status.
 * Logs diagnostics on execution and displays discrete toast if action is needed.
 * Strictly uses People Chip identity (never compares to accenture_email).
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {boolean} [showToast=true]
 * @return {Object|null}
 */
function checkActiveUserStatus(targetSpreadsheet, showToast) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;

  var audit = auditUserAndMembersIdentity(ss);
  var userEmail = audit.activeEmail || "";
  var memberFound = audit.matchingMember !== null;
  var memberName = (audit.matchingMember && audit.matchingMember.name) ? audit.matchingMember.name : (userEmail ? userEmail.split("@")[0] : "User");
  var missingFields = [];
  var todayPresence = "Not Logged";

  if (audit.matchingMember) {
    if (!extractCellValue(audit.matchingMember.role)) missingFields.push("Role");
    if (!extractCellValue(audit.matchingMember.department)) missingFields.push("Department");
    todayPresence = extractCellValue(audit.matchingMember.todayPresence) || "Not Logged";
  }

  var hasBaseline = audit.baselinesCount > 0;
  var warnings = [];

  if (!userEmail) {
    warnings.push("Active user email could not be read in onOpen trigger (privacy restriction). Run 'Debug User & Member Identity' from the menu.");
  } else if (!memberFound) {
    warnings.push("You are not registered in tbl_members yet.");
  } else {
    if (!hasBaseline) {
      warnings.push("Weekly presence routine is not set up.");
    }
    if (missingFields.length > 0) {
      warnings.push("Missing profile info: " + missingFields.join(", ") + ".");
    }
  }

  if (showToast !== false && warnings.length > 0) {
    try {
      if (!userEmail) {
        ss.toast(
          "Could not detect active user in onOpen trigger. Run 'Project Management > Debug User & Member Identity' to test.",
          "User Identity Notice",
          8
        );
      } else {
        ss.toast(
          warnings.join(" ") + " Open 'Presence & Leaves' from the menu to configure.",
          "Action Needed: " + memberName,
          10
        );
      }
    } catch (tErr) {}
  }

  return {
    email: userEmail,
    name: memberName,
    memberFound: memberFound,
    hasBaseline: hasBaseline,
    baselinesCount: audit.baselinesCount,
    missingFields: missingFields,
    todayPresence: todayPresence,
    warnings: warnings,
    auditLogs: audit.logLines
  };
}

/**
 * Returns initial setup data for the sidebar:
 * - list of members from tbl_members (Column 1: member [PEOPLE_CHIP], with extracted email)
 * - current user email
 * - standard location choices
 * - active user checklist
 *
 * @return {Object}
 */
function getSidebarInitialData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var membersSheet = ss ? ss.getSheetByName("tbl_members") : null;
  var members = [];

  if (membersSheet) {
    var maxRows = Math.min(membersSheet.getMaxRows(), 100);
    if (maxRows > 1) {
      var ssId = "";
      try {
        ssId = (ss && typeof ss.getId === "function") ? ss.getId() : "";
      } catch (eId) {}
      var chipEmails = ssId ? fetchMemberChipEmails(ssId, "tbl_members", 2, maxRows - 1) : {};
      var mRange = membersSheet.getRange(2, 1, maxRows - 1, 1);
      var data = mRange.getValues();
      var rData = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : [];

      for (var i = 0; i < data.length; i++) {
        var mRich = (rData && rData[i]) ? rData[i][0] : null;
        var mVal = extractCellValue(data[i][0], mRich);
        if (!mVal) continue;
        var chipEmail = chipEmails[i + 2] || extractEmailFromChip(data[i][0], mRich);

        members.push({
          member: mVal,
          displayName: mVal,
          email: chipEmail
        });
      }
    }
  }

  var activeUserEmail = "";
  try {
    activeUserEmail = Session.getActiveUser().getEmail();
  } catch (e) {}

  var locationOptions = [
    "Main Office",
    "Remote",
    "Downtown Office",
    "Branch Office",
    "Client Site",
    "Off"
  ];

  var userChecklist = null;
  if (activeUserEmail) {
    userChecklist = checkActiveUserStatus(ss, false);
  }

  return {
    members: members,
    activeUserEmail: activeUserEmail,
    locationOptions: locationOptions,
    leaveTypes: ["Vacation", "Sick Leave", "Maternity/Paternity", "Unpaid Leave"],
    leaveStatuses: ["Approved", "Requested"],
    userChecklist: userChecklist
  };
}

/**
 * Checks whether two date ranges overlap.
 * Open-ended ranges (null or empty string) extend infinitely.
 * Adjacent ranges where s2 == e1 do NOT overlap.
 *
 * @param {string} s1Str
 * @param {string} e1Str
 * @param {string} s2Str
 * @param {string} e2Str
 * @return {boolean}
 */
function checkBaselineDateOverlap(s1Str, e1Str, s2Str, e2Str) {
  if (!s1Str || !s2Str) return false;
  var s1 = new Date(s1Str);
  s1.setHours(0, 0, 0, 0);
  var s2 = new Date(s2Str);
  s2.setHours(0, 0, 0, 0);

  var e1 = e1Str ? new Date(e1Str) : null;
  if (e1) e1.setHours(0, 0, 0, 0);
  var e2 = e2Str ? new Date(e2Str) : null;
  if (e2) e2.setHours(0, 0, 0, 0);

  // Overlap if s1 <= (e2 || infinity) AND s2 < (e1 || infinity)
  var cond1 = !e2 || s1 <= e2;
  var cond2 = !e1 || s2 < e1;

  return cond1 && cond2;
}

/**
 * Retrieves the latest baseline records for a given member.
 * When multiple historical baselines exist, returns the most recent one (sorted by startDate desc).
 *
 * @param {string} memberIdentifier
 * @return {Object} Map of day of week to { location_am, location_pm }, startDate, endDate
 */
function getMemberBaselines(memberIdentifier) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {
    days: {},
    startDate: "",
    endDate: "",
    totalBaselinesCount: 0
  };
  if (!ss || !memberIdentifier) return result;

  var targetMember = String(memberIdentifier).trim().toLowerCase();
  var baselinesByStart = {};
  var ranges = readPresenceRanges(ss);

  for (var i = 0; i < ranges.length; i++) {
    var rg = ranges[i];
    if (rg.rangeType !== "Baseline" || rg.status.toLowerCase() === "cancelled") continue;
    if (rg.member.toLowerCase() !== targetMember) continue;
    var sDate = rg.startIso;
    if (!baselinesByStart[sDate]) {
      baselinesByStart[sDate] = { startDate: sDate, endDate: rg.endIso, days: {} };
    }
    if (rg.dayOfWeek) {
      baselinesByStart[sDate].days[rg.dayOfWeek] = { location_am: rg.locationAm, location_pm: rg.locationPm };
    }
  }

  var startKeys = Object.keys(baselinesByStart);
  result.totalBaselinesCount = startKeys.length;
  if (startKeys.length === 0) return result;

  // Sort by startDate descending to get the latest baseline
  startKeys.sort(function(a, b) {
    return new Date(b).getTime() - new Date(a).getTime();
  });

  var latest = baselinesByStart[startKeys[0]];
  result.startDate = latest.startDate;
  result.endDate = latest.endDate;
  result.days = latest.days;

  return result;
}

/**
 * Saves a new Baseline routine into tbl_presence_ranges (one row per weekday).
 * Does NOT overwrite older historical baselines; appends new rows.
 * Checks for date overlaps with existing baselines for the member.
 * If overlapping and adjustPreviousEnd is not true, prompts the user.
 * If adjustPreviousEnd is true, updates the end date of overlapping previous baselines to new startDate.
 *
 * @param {Object} payload { member, startDate, endDate, days, adjustPreviousEnd }
 * @return {Object} Status report
 */
function saveMemberBaselines(payload) {
  var member = payload ? (payload.member || payload.memberId) : null;
  if (!member) {
    throw new Error("Missing member");
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("tbl_presence_ranges");
  var membersSheet = ss.getSheetByName("tbl_members");
  if (!sheet) {
    throw new Error("Sheet tbl_presence_ranges not found. Please run Initialize / Upgrade first.");
  }

  var targetMember = String(member).trim();
  var targetMemberLower = targetMember.toLowerCase();
  var startDate = payload.startDate ? String(payload.startDate).trim() : formatDateIso(new Date());
  var endDate = payload.endDate ? String(payload.endDate).trim() : "";
  var days = payload.days || {};
  var adjustPreviousEnd = payload.adjustPreviousEnd === true;

  // 1. Scan this Member's existing Baselines in tbl_presence_ranges and detect overlaps
  var overlappingRows = [];
  var existingMemberRows = 0;
  var existingRanges = readPresenceRanges(ss);
  for (var i = 0; i < existingRanges.length; i++) {
    var rg = existingRanges[i];
    if (rg.rangeType !== "Baseline" || rg.status.toLowerCase() === "cancelled") continue;
    if (rg.member.toLowerCase() !== targetMemberLower) continue;
    existingMemberRows++;
    if (checkBaselineDateOverlap(rg.startIso, rg.endIso, startDate, endDate)) {
      overlappingRows.push({ rowNumber: rg.row, startDate: rg.startIso, endDate: rg.endIso });
    }
  }

  // 2. If overlap detected and confirmation not yet given, ask user
  if (overlappingRows.length > 0 && !adjustPreviousEnd) {
    var overlapStart = overlappingRows[0].startDate;
    return {
      success: false,
      overlap: true,
      message: "A previous baseline overlaps with the new date range (starts " + overlapStart + "). Do you want to set the end date of the previous baseline to the start date of the new baseline (" + startDate + ")?",
      overlappingStartDate: overlapStart,
      newStartDate: startDate
    };
  }

  // 3. If overlap confirmed, update end date of previous overlapping baselines to new startDate
  if (overlappingRows.length > 0 && adjustPreviousEnd) {
    for (var ov = 0; ov < overlappingRows.length; ov++) {
      sheet.getRange(overlappingRows[ov].rowNumber, presenceRangeColumn("end_date")).setValue(startDate);
    }
  }

  // 4-5. Build one Baseline range per weekday (weekend days only when provided)
  var allPossibleDays = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
  var newRows = [];

  for (var w = 0; w < allPossibleDays.length; w++) {
    var dayName = allPossibleDays[w];
    var isWeekend = (dayName === "Saturday" || dayName === "Sunday");
    if (isWeekend && (!days || !days[dayName])) {
      continue;
    }
    var fallbackLoc = isWeekend ? "Off" : "Main Office";
    var dayData = days[dayName] || { location_am: fallbackLoc, location_pm: fallbackLoc };
    newRows.push({
      member: targetMember,
      range_type: "Baseline",
      start_date: startDate,
      end_date: endDate,
      day_of_week: dayName,
      location_am: dayData.location_am || fallbackLoc,
      location_pm: dayData.location_pm || fallbackLoc,
      status: "Approved"
    });
  }

  // 6. Append without modifying non-overlapping historical data
  var targetStartRow = appendPresenceRanges(ss, newRows).firstRow;

  // Copy People Chip from tbl_members if available
  if (membersSheet && typeof copyCellChip === "function") {
    var mMax = Math.min(membersSheet.getMaxRows(), 300);
    var mRange = membersSheet.getRange(2, 1, mMax - 1, Math.max(membersSheet.getLastColumn(), 6));
    var mVals = mRange.getValues();
    var mRich = (typeof mRange.getRichTextValues === "function") ? mRange.getRichTextValues() : [];
    var sourceMRow = -1;
    var canonicalMemberName = targetMember;

    for (var m = 0; m < mVals.length; m++) {
      var rawM = mVals[m][0];
      var rM = (mRich && mRich[m]) ? mRich[m][0] : null;
      var nameM = (typeof extractCellValue === "function") ? extractCellValue(rawM, rM) : String(rawM).trim();
      var emailM = String(mVals[m][1] || "").trim().toLowerCase();

      var isMatch = false;
      if (nameM.toLowerCase() === targetMemberLower || (emailM && emailM === targetMemberLower)) {
        isMatch = true;
      } else if (typeof isUserMemberMatch === "function" && isUserMemberMatch(rawM, rM, targetMember)) {
        isMatch = true;
      }

      if (isMatch) {
        sourceMRow = m + 2;
        canonicalMemberName = nameM || targetMember;
        break;
      }
    }

    if (sourceMRow > 0) {
      for (var r = 0; r < newRows.length; r++) {
        copyCellChip(
          membersSheet,
          sourceMRow,
          1,
          sheet,
          targetStartRow + r,
          presenceRangeColumn("member"),
          canonicalMemberName,
          ss.getId()
        );
      }
    }
  }

  // Also synchronize and hydrate all People Chips in tbl_presence_ranges
  if (typeof syncBaselinesPeopleChips === "function") {
    try {
      syncBaselinesPeopleChips(ss);
    } catch (eSyncChips) {}
  }

  if (typeof applyProtections === "function") {
    try {
      applyProtections(ss);
    } catch (eProt) {}
  }

  return {
    success: true,
    savedRows: newRows.length,
    member: targetMember,
    overlapAdjusted: overlappingRows.length > 0
  };
}

/**
 * Helper to format date object to YYYY-MM-DD
 */
function formatDateIso(dateVal) {
  if (!dateVal) return "";
  if (typeof dateVal === "string") {
    return dateVal.indexOf("T") !== -1 ? dateVal.split("T")[0] : dateVal;
  }
  try {
    var y = dateVal.getFullYear();
    var m = ("0" + (dateVal.getMonth() + 1)).slice(-2);
    var d = ("0" + dateVal.getDate()).slice(-2);
    return y + "-" + m + "-" + d;
  } catch (e) {
    return String(dateVal);
  }
}

/**
 * Submits a Member Leave from the sidebar as a Leave range in tbl_presence_ranges
 * (the leave type is stored as the AM/PM Work Location).
 * Automatically copies the People Chip from tbl_members, applies typography and protections,
 * and triggers compilePresenceLogs and checkProjectBackupConflicts in real time.
 *
 * @param {Object} payload
 * @param {string} payload.member
 * @param {string} payload.startDate
 * @param {string} [payload.startPeriod] - Default "AM"
 * @param {string} payload.endDate
 * @param {string} [payload.endPeriod] - Default "PM"
 * @param {string} payload.leaveType
 * @param {string} [payload.status] - Default "Approved"
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object}
 */
function submitMemberLeave(payload, targetSpreadsheet) {
  if (!payload || !payload.member || !payload.startDate || !payload.endDate || !payload.leaveType) {
    return { success: false, message: "Missing required leave fields (member, startDate, endDate, leaveType)." };
  }

  var sDate = new Date(payload.startDate);
  var eDate = new Date(payload.endDate);
  sDate.setHours(0, 0, 0, 0);
  eDate.setHours(0, 0, 0, 0);

  if (sDate > eDate) {
    return { success: false, message: "Start date cannot be after end date." };
  }

  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet" };

  var leavesSheet = ss.getSheetByName("tbl_presence_ranges");
  var membersSheet = ss.getSheetByName("tbl_members");

  if (!leavesSheet) {
    return { success: false, message: "tbl_presence_ranges sheet not found." };
  }

  var targetMember = String(payload.member).trim();
  var targetMemberLower = targetMember.toLowerCase();
  var status = payload.status ? String(payload.status).trim() : "Approved";
  var leaveId = "leave_" + Math.random().toString(36).substring(2, 9);
  var startDateStr = formatDateIso(payload.startDate);
  var endDateStr = formatDateIso(payload.endDate);
  var startPeriod = (payload.startPeriod ? String(payload.startPeriod).toUpperCase() : "AM");
  var endPeriod = (payload.endPeriod ? String(payload.endPeriod).toUpperCase() : "PM");

  var targetRow = appendPresenceRanges(ss, [{
    range_id: leaveId,
    member: targetMember,
    range_type: "Leave",
    start_date: startDateStr,
    start_period: startPeriod,
    end_date: endDateStr,
    end_period: endPeriod,
    day_of_week: "All",
    location_am: payload.leaveType,
    location_pm: payload.leaveType,
    status: status
  }]).firstRow;

  // Copy People Chip from tbl_members if available
  if (membersSheet && typeof copyCellChip === "function") {
    var mMax = Math.min(membersSheet.getMaxRows(), 200);
    var mValues = membersSheet.getRange(2, 1, mMax - 1, 1).getValues();
    var sourceMRow = -1;
    for (var m = 0; m < mValues.length; m++) {
      if (extractCellValue(mValues[m][0]).toLowerCase() === targetMemberLower) {
        sourceMRow = m + 2;
        break;
      }
    }
    if (sourceMRow > 0) {
      copyCellChip(
        membersSheet,
        sourceMRow,
        1,
        leavesSheet,
        targetRow,
        presenceRangeColumn("member"),
        targetMember,
        ss.getId()
      );
    }
  }

  if (typeof applyProtections === "function") {
    try {
      applyProtections(ss);
    } catch (eProt) {}
  }

  // If approved, trigger real-time presence log compilation so today_presence and log update immediately
  var compileSummary = null;
  if (status.toLowerCase() === "approved" && typeof compilePresenceLogs === "function") {
    try {
      compileSummary = compilePresenceLogs(null, null, ss, payload.asOfDate);
    } catch (eComp) {
      Logger.log("Compilation note on leave submission: " + eComp);
    }
  }

  // Check project backup conflicts across active projects
  var backupReport = null;
  if (typeof checkProjectBackupConflicts === "function") {
    try {
      backupReport = checkProjectBackupConflicts(ss, { asOfDateParam: payload.asOfDate });
    } catch (eConf) {
      Logger.log("Project backup conflict note on leave submission: " + eConf);
    }
  }

  return {
    success: true,
    leaveId: leaveId,
    member: targetMember,
    startDate: startDateStr,
    startPeriod: startPeriod,
    endDate: endDateStr,
    endPeriod: endPeriod,
    leaveType: payload.leaveType,
    status: status,
    compiled: compileSummary !== null,
    backupReport: backupReport
  };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    showBaselineSidebar: showBaselineSidebar,
    getSidebarInitialData: getSidebarInitialData,
    checkActiveUserStatus: checkActiveUserStatus,
    auditUserAndMembersIdentity: auditUserAndMembersIdentity,
    debugUserIdentity: debugUserIdentity,
    explainUserMemberMatch: explainUserMemberMatch,
    getMemberBaselines: getMemberBaselines,
    saveMemberBaselines: saveMemberBaselines,
    submitMemberLeave: submitMemberLeave,
    checkBaselineDateOverlap: checkBaselineDateOverlap,
    extractCellValue: extractCellValue,
    extractEmailFromChip: extractEmailFromChip,
    isUserMemberMatch: isUserMemberMatch,
    fetchMemberChipEmails: fetchMemberChipEmails,
    normalizePersonName: normalizePersonName,
    formatDateIso: formatDateIso
  };
}

