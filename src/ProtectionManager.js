/**
 * Project Management Workspace - Protection Manager
 * Configures schema-driven protections:
 * - strictlyProtectedColumns: Strictly prevents non-admins from changing cells (warningOnly: false).
 * - warningProtectedColumns: Displays native confirmation warning dialog to users (warningOnly: true).
 * Displays warning notifications exclusively to authorized admins when they make manual edits to strictly protected columns.
 */

/**
 * Returns list of authorized admin emails allowed to edit strictly protected ranges.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {Array<string>}
 */
function getAuthorizedAdmins(ss) {
  var spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  var admins = {};

  // Script executor / effective user
  try {
    var me = Session.getEffectiveUser();
    if (me && me.getEmail()) admins[me.getEmail().toLowerCase()] = true;
  } catch (eMe) {}

  // Spreadsheet owner
  if (spreadsheet) {
    try {
      var owner = spreadsheet.getOwner();
      if (owner && owner.getEmail()) admins[owner.getEmail().toLowerCase()] = true;
    } catch (eOwner) {}
  }

  // Configured admins in Document Properties
  try {
    var props = PropertiesService.getDocumentProperties();
    var adminProp = props ? props.getProperty("ADMIN_EMAILS") : null;
    if (adminProp) {
      var list = adminProp.split(",");
      for (var a = 0; a < list.length; a++) {
        var em = list[a].trim().toLowerCase();
        if (em) admins[em] = true;
      }
    }
  } catch (eProps) {}

  return Object.keys(admins);
}

/**
 * Configures authorized admin emails in Document Properties.
 *
 * @param {Array<string>} emails
 */
function setAuthorizedAdmins(emails) {
  if (!emails || !Array.isArray(emails)) return;
  var props = PropertiesService.getDocumentProperties();
  if (props) {
    props.setProperty("ADMIN_EMAILS", emails.map(function(e) { return e.trim().toLowerCase(); }).join(","));
  }
}

/**
 * Returns the 1-based index of the last valid, fully created row in the sheet,
 * or 1 if only the header row exists.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Object} [config]
 * @return {number}
 */
function getLastCreatedRow(sheet, config) {
  var lastRow = (typeof sheet.getLastRow === "function") ? (sheet.getLastRow() || 1) : 1;
  if (lastRow <= 1) return 1;

  var numCols = (config && config.columns) ? config.columns.length : Math.max((sheet.getLastColumn && sheet.getLastColumn()) || 1, 1);
  try {
    var values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    var lastCreated = 1;

    for (var r = 0; r < values.length; r++) {
      var rowVals = values[r];
      if (typeof validateRow === "function" && config) {
        var vRes = validateRow(config.sheetName, rowVals, config.columns);
        if (vRes.populated && vRes.valid) {
          lastCreated = r + 2;
        }
      } else {
        var isPopulated = false;
        for (var c = 0; c < rowVals.length; c++) {
          var v = rowVals[c];
          if (v !== "" && v !== null && v !== undefined) {
            if (typeof v === "string" && v.trim() !== "") {
              isPopulated = true;
              break;
            } else if (typeof v !== "string") {
              isPopulated = true;
              break;
            }
          }
        }
        if (isPopulated) lastCreated = r + 2;
      }
    }
    return lastCreated;
  } catch (err) {
    return 1;
  }
}

/**
 * Checks if a specific cell (row, col) is covered by a protected range on the sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} row
 * @param {number} col
 * @return {Object} { isProtected: boolean, isWarningOnly: boolean, description: string }
 */
function isCellProtected(sheet, row, col) {
  if (!sheet || typeof sheet.getProtections !== "function") {
    return { isProtected: false, isWarningOnly: false, description: "" };
  }

  var protections = [];
  try {
    protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  } catch (e) {
    try {
      protections = sheet.getProtections();
    } catch (e2) {}
  }

  for (var i = 0; i < protections.length; i++) {
    var p = protections[i];
    var range = (typeof p.getRange === "function") ? p.getRange() : p.range;
    if (range) {
      var startR = (typeof range.getRow === "function") ? range.getRow() : range.startRow;
      var startC = (typeof range.getColumn === "function") ? range.getColumn() : range.startCol;
      var numR = (typeof range.getNumRows === "function") ? range.getNumRows() : range.numRows;
      var numC = (typeof range.getNumColumns === "function") ? range.getNumColumns() : range.numCols;

      if (row >= startR && row < (startR + numR) && col >= startC && col < (startC + numC)) {
        var isWarn = (typeof p.isWarningOnly === "function") ? p.isWarningOnly() : (p.warningOnly === true);
        var desc = (typeof p.getDescription === "function") ? p.getDescription() : (p.description || "");
        return {
          isProtected: true,
          isWarningOnly: isWarn,
          description: desc
        };
      }
    }
  }

  return { isProtected: false, isWarningOnly: false, description: "" };
}

/**
 * Applies schema-declared Google Sheets ProtectedRange locks:
 * - strictlyProtectedColumns: warningOnly=false, strips non-admin editors.
 * - warningProtectedColumns: warningOnly=true, presents warning confirmation dialog.
 * Crucial: Warning protections are applied ONLY to created (populated) rows (rows 2 to lastCreatedRow).
 * Empty rows remain completely unprotected so registering new rows (e.g. new tenures) never triggers a warning.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Report of protected ranges applied.
 */
function applyProtections(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet" };

  var strictCount = 0;
  var warningCount = 0;
  var adminList = getAuthorizedAdmins(ss);
  var adminMap = {};
  for (var a = 0; a < adminList.length; a++) {
    adminMap[adminList[a]] = true;
  }

  for (var tableKey in SCHEMA_SPEC) {
    if (!SCHEMA_SPEC.hasOwnProperty(tableKey)) continue;
    var config = SCHEMA_SPEC[tableKey];
    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) continue;

    var maxRows = Math.max(sheet.getMaxRows(), 50);
    var lastCreatedRow = getLastCreatedRow(sheet, config);
    var numCreatedRows = lastCreatedRow > 1 ? (lastCreatedRow - 1) : 0;

    // Clean up existing protections previously installed for this sheet
    try {
      var existingProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
      for (var i = 0; i < existingProtections.length; i++) {
        var desc = existingProtections[i].getDescription();
        if (desc && desc.indexOf("_protected") !== -1) {
          existingProtections[i].remove();
        }
      }
    } catch (err) {
      // Protection querying may vary by mock/environment
    }

    var strictCols = config.strictlyProtectedColumns || config.protectedColumns || [];
    var warnCols = config.warningProtectedColumns || [];

    // 1. Apply Strict Protections (setWarningOnly: false, restricted to admins)
    // For tbl_members: locks read-only columns across maxRows so users cannot manually add/edit member records.
    // For other tables: locks strictly protected system ID columns once rows are created.
    for (var s = 0; s < strictCols.length; s++) {
      var sColName = strictCols[s];
      if (warnCols.indexOf(sColName) !== -1) continue;

      var sColIdx = config.columns.indexOf(sColName) + 1;
      if (sColIdx > 0) {
        var sRows = (config.sheetName === "tbl_members") ? maxRows : (numCreatedRows > 0 ? numCreatedRows : 0);
        if (sRows > 0) {
          var sStartRow = (config.sheetName === "tbl_members") ? 1 : 2;
          var sRange = sheet.getRange(sStartRow, sColIdx, sRows, 1);
          var sProtection = sRange.protect().setDescription(config.sheetName + "_" + sColName + "_protected");

          sProtection.setWarningOnly(false);

          for (var sa = 0; sa < adminList.length; sa++) {
            try {
              sProtection.addEditor(adminList[sa]);
            } catch (addErr) {}
          }

          try {
            var currentEditors = sProtection.getEditors();
            var toRemove = [];
            for (var e = 0; e < currentEditors.length; e++) {
              var edEmail = currentEditors[e].getEmail().toLowerCase();
              if (!adminMap[edEmail]) {
                toRemove.push(currentEditors[e]);
              }
            }
            if (toRemove.length > 0) {
              sProtection.removeEditors(toRemove);
            }
          } catch (remErr) {}

          try {
            if (sProtection.canDomainEdit && sProtection.canDomainEdit()) {
              sProtection.setDomainEdit(false);
            }
          } catch (domErr) {}

          strictCount++;
        }
      }
    }

    // 2. Apply Warning Protections (setWarningOnly: true)
    // Applied ONLY once rows are created (rows 2 to lastCreatedRow).
    // Empty rows remain completely unprotected so registering new rows (e.g. new tenures) never triggers a warning.
    if (numCreatedRows > 0) {
      for (var w = 0; w < warnCols.length; w++) {
        var wColName = warnCols[w];
        var wColIdx = config.columns.indexOf(wColName) + 1;
        if (wColIdx > 0) {
          var wRange = sheet.getRange(2, wColIdx, numCreatedRows, 1);
          var wProtection = wRange.protect().setDescription(config.sheetName + "_" + wColName + "_protected");

          wProtection.setWarningOnly(true);
          warningCount++;
        }
      }
    }
  }

  return {
    success: true,
    strictlyProtectedCount: strictCount,
    warningProtectedCount: warningCount,
    protectedColumnsCount: strictCount + warningCount,
    authorizedAdmins: adminList
  };
}

/**
 * Checks whether protection warning notifications are currently snoozed/disabled
 * for the active user (either via CacheService or UserProperties).
 *
 * @return {boolean}
 */
function isProtectionWarningSnoozed() {
  try {
    var cache = (typeof CacheService !== "undefined" && CacheService.getUserCache) ? CacheService.getUserCache() : null;
    if (cache && cache.get("PROTECTION_WARNING_SNOOZE") === "true") {
      return true;
    }
  } catch (eC) {}

  try {
    var props = (typeof PropertiesService !== "undefined" && PropertiesService.getUserProperties) ? PropertiesService.getUserProperties() : null;
    if (props) {
      var until = props.getProperty("PROTECTION_WARNING_SNOOZE_UNTIL");
      if (until && Date.now() < parseInt(until, 10)) {
        return true;
      }
    }
  } catch (eP) {}

  return false;
}

/**
 * Snoozes protection warning notifications for the specified duration (default 5 minutes).
 * Persists in both UserCache (TTL seconds) and UserProperties (epoch timestamp).
 *
 * @param {number} [durationMinutes=5]
 */
function snoozeProtectionWarnings(durationMinutes) {
  var minutes = durationMinutes || 5;
  var snoozeMs = minutes * 60 * 1000;
  var snoozeUntil = Date.now() + snoozeMs;

  try {
    var cache = (typeof CacheService !== "undefined" && CacheService.getUserCache) ? CacheService.getUserCache() : null;
    if (cache) {
      cache.put("PROTECTION_WARNING_SNOOZE", "true", minutes * 60);
    }
  } catch (eC) {}

  try {
    var props = (typeof PropertiesService !== "undefined" && PropertiesService.getUserProperties) ? PropertiesService.getUserProperties() : null;
    if (props) {
      props.setProperty("PROTECTION_WARNING_SNOOZE_UNTIL", String(snoozeUntil));
    }
  } catch (eP) {}
}

/**
 * Clears any active protection warning notification snooze.
 */
function clearProtectionWarningSnooze() {
  try {
    var cache = (typeof CacheService !== "undefined" && CacheService.getUserCache) ? CacheService.getUserCache() : null;
    if (cache) cache.remove("PROTECTION_WARNING_SNOOZE");
  } catch (eC) {}

  try {
    var props = (typeof PropertiesService !== "undefined" && PropertiesService.getUserProperties) ? PropertiesService.getUserProperties() : null;
    if (props) props.deleteProperty("PROTECTION_WARNING_SNOOZE_UNTIL");
  } catch (eP) {}
}

/**
 * Handles an edit event on a protected cell.
 * Displays a blocking modal confirmation dialog (SpreadsheetApp.getUi().alert) asking
 * the user if they want to disable notifications for the next 5 minutes.
 * If snoozed, suppresses all alerts and toasts.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} editedRow
 * @param {number} editedCol
 * @param {string} colName
 * @param {Object} protInfo
 * @return {boolean} True if warning was handled, false if snoozed or skipped.
 */
function handleProtectedCellEdit(sheet, editedRow, editedCol, colName, protInfo) {
  if (!sheet || !protInfo || !protInfo.isProtected) return false;
  if (isProtectionWarningSnoozed()) return false;

  var sheetName = sheet.getName();
  var ss = (sheet.getParent && typeof sheet.getParent === "function") ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet();

  var isWarningOnly = protInfo.isWarningOnly === true;
  var alertTitle = isWarningOnly
    ? "Warning-Protected Column Edited"
    : "Strictly Protected Column Edited";

  var reason = isWarningOnly
    ? "Warning: You edited warning-protected column '" + colName + "' in " + sheetName + " (Row " + editedRow + "). Please confirm this modification was intentional."
    : "Admin Warning: You edited strictly protected column '" + colName + "' in " + sheetName + " (Row " + editedRow + "). Manual edits to calculated columns may be overwritten during synchronization.";

  var promptMsg = reason + "\n\nDo you want to disable these warning notifications for the next 5 minutes?";

  var userWantsSnooze = false;

  try {
    var ui = SpreadsheetApp.getUi();
    if (ui && typeof ui.alert === "function") {
      var btnSet = ui.ButtonSet ? ui.ButtonSet.YES_NO : "YES_NO";
      var res = ui.alert(alertTitle, promptMsg, btnSet);
      var yesBtn = (ui.Button && ui.Button.YES) ? ui.Button.YES : "YES";
      if (res === yesBtn || res === "YES") {
        userWantsSnooze = true;
      }
    }
  } catch (eUi) {
    Logger.log("UI alert note: " + eUi);
  }

  if (userWantsSnooze) {
    snoozeProtectionWarnings(5);
    if (ss && typeof ss.toast === "function") {
      ss.toast(
        "Protection warning notifications have been disabled for the next 5 minutes.",
        "Notifications Snoozed",
        4
      );
    }
    return true;
  }

  // If user declined snooze, or if dialog was displayed/closed, show toast notification as confirmation
  if (ss && typeof ss.toast === "function") {
    ss.toast(reason, alertTitle, 6);
  }
  return true;
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    applyProtections: applyProtections,
    getAuthorizedAdmins: getAuthorizedAdmins,
    setAuthorizedAdmins: setAuthorizedAdmins,
    getLastCreatedRow: getLastCreatedRow,
    isCellProtected: isCellProtected,
    isProtectionWarningSnoozed: isProtectionWarningSnoozed,
    snoozeProtectionWarnings: snoozeProtectionWarnings,
    clearProtectionWarningSnooze: clearProtectionWarningSnooze,
    handleProtectedCellEdit: handleProtectedCellEdit
  };
}

