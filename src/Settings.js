/**
 * Settings (ADR 0049): the organisation-specific values, kept in the hidden
 * "Settings" tab (Setting | Value | Description), edited with the menu
 * "Edit Settings". Missing rows are added with their default on upgrade;
 * existing values are never overwritten.
 */

var SETTINGS_DEFAULTS = [
  { key: "default_role", label: "Default role", value: "Team Member",
    description: "Role shown in Members when a contract has no role." },
  { key: "default_department", label: "Default department", value: "Delivery",
    description: "Department shown in Members when a contract has no department." },
  { key: "public_holidays", label: "Public holidays", value: "France",
    description: "Country whose public holidays are added to the Holidays tab each year: France, or None to add them yourself." },
  { key: "onboarding_project_id", label: "Onboarding project ID", value: "ONBOARDING",
    description: "Project ID of the onboarding task created for each new member." },
  { key: "onboarding_task_title", label: "Onboarding task title", value: "Complete profile: enter Accenture email and phone number in Team Directory",
    description: "Title of the onboarding task created for each new member." },
  { key: "onboarding_task_description", label: "Onboarding task description",
    value: "Please enter your Accenture email address and phone number in the Members tab (columns Accenture email and Phone).",
    description: "Description of the onboarding task." }
];

function settingDefault(key) {
  var d = SETTINGS_DEFAULTS.filter(function(s) { return s.key === key; })[0];
  return d ? d.value : "";
}

/** Every setting by key: the tab's value when filled in, else the default (read each time: a few cells). */
function readSettings(ss) {
  var out = {};
  SETTINGS_DEFAULTS.forEach(function(s) { out[s.key] = s.value; });
  var sheet = ss ? getTable(ss, "settings") : null;
  if (sheet) {
    var byLabel = {};
    SETTINGS_DEFAULTS.forEach(function(s) { byLabel[s.label.toLowerCase()] = s.key; byLabel[s.key] = s.key; });
    readTableRows(sheet, 2, 100).forEach(function(r) {
      var key = byLabel[String(r.vals[0] || "").trim().toLowerCase()];
      var value = String(r.vals[1] === null || r.vals[1] === undefined ? "" : r.vals[1]).trim();
      if (key && value) out[key] = value;
    });
  }
  return out;
}

/** One setting (tab value, else default). */
function getSetting(ss, key) {
  var all = readSettings(ss);
  return all.hasOwnProperty(key) ? all[key] : settingDefault(key);
}

/**
 * Adds the settings missing from the tab with their default value.
 *
 * @return {number} rows added
 */
function seedDefaultSettings(ss) {
  var sheet = getTable(ss, "settings");
  if (!sheet) return 0;
  var present = {};
  readTableRows(sheet, 1, 100).forEach(function(r) { present[String(r.vals[0] || "").trim().toLowerCase()] = true; });
  var rows = SETTINGS_DEFAULTS.filter(function(s) { return !present[s.label.toLowerCase()] && !present[s.key]; })
    .map(function(s) { return [s.label, s.value, s.description]; });
  if (!rows.length) return 0;
  var start = (typeof findNextAvailableRow === "function") ? findNextAvailableRow(sheet) : sheet.getLastRow() + 1;
  if (start + rows.length - 1 > sheet.getMaxRows() && typeof sheet.insertRowsAfter === "function") {
    sheet.insertRowsAfter(sheet.getMaxRows(), start + rows.length - 1 - sheet.getMaxRows());
  }
  sheet.getRange(start, 1, rows.length, 3).setValues(rows);
  return rows.length;
}

/** Menu: Edit Settings. */
function showSettingsTab() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getTable(ss, "settings");
  if (!sheet) return { success: false, message: "The Settings tab was not found. Run Initialize / Upgrade first." };
  if (sheet.isSheetHidden()) sheet.showSheet();
  sheet.activate();
  try {
    ss.toast("Change a value in the Value column (an empty value means the default), then hide the tab again (right-click → Hide sheet).", "Settings", 8);
  } catch (eToast) {}
  return { success: true };
}

/**
 * Edit on the Settings tab: re-apply what depends on it (the Members formulas
 * use the default role and department).
 */
function handleSettingsEdit(e) {
  var ss = (e && e.source) || SpreadsheetApp.getActiveSpreadsheet();
  if (typeof applyMemberFormulas === "function") applyMemberFormulas(ss);
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    SETTINGS_DEFAULTS: SETTINGS_DEFAULTS,
    readSettings: readSettings,
    getSetting: getSetting,
    seedDefaultSettings: seedDefaultSettings,
    showSettingsTab: showSettingsTab,
    handleSettingsEdit: handleSettingsEdit
  };
}
