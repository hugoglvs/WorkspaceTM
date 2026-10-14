/**
 * Projects (ADR 0053): the project name is the reference. Tasks, Backups and
 * Recurring Tasks hold the name in their "Project" column, a dropdown of the
 * Projects tab's names (schema `projectColumns`). Renaming a project renames it
 * in every one of those rows; a repeated name turns red in Projects.
 */

/** Column number of a Projects column key. */
function projectCol(key) {
  return SCHEMA_SPEC.projects.columns.indexOf(key) + 1;
}

/**
 * Every named project, in tab order.
 * @return {Array<{row: number, name: string, key: string, status: string, start: (Date|null),
 *   target: (Date|null), lead: string, backup: string}>}
 */
function readProjects(ss) {
  var sheet = ss ? getTable(ss, "projects") : null;
  if (!sheet || sheet.getMaxRows() < 2) return [];
  var cols = SCHEMA_SPEC.projects.columns;
  var c = {};
  cols.forEach(function(k, i) { c[k] = i; });
  var range = sheet.getRange(2, 1, readRowLimit(sheet, 500) - 1, cols.length);
  var vals = range.getValues();
  var rich = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    var name = String(v[c.name] || "").trim();
    if (!name) continue;
    var r = rich[i] || [];
    out.push({
      row: i + 2, name: name, key: name.toLowerCase(), status: String(v[c.status] || "").trim(),
      start: toDate(v[c.start_date]), target: toDate(v[c.target_date]),
      lead: cellName(v[c.lead_member], r[c.lead_member]), backup: cellName(v[c.backup_member], r[c.backup_member])
    });
  }
  return out;
}

/** Project names for the dropdowns, in tab order, without repeats. */
function projectNames(ss) {
  var seen = {};
  return readProjects(ss).filter(function(p) {
    if (seen[p.key]) return false;
    seen[p.key] = true;
    return true;
  }).map(function(p) { return p.name; });
}

/** A project by name (case-insensitive), or null. */
function findProject(ss, name) {
  var key = String(name || "").trim().toLowerCase();
  if (!key) return null;
  return readProjects(ss).filter(function(p) { return p.key === key; })[0] || null;
}

/**
 * Renames a project in every table column that refers to it.
 * @return {number} cells changed
 */
function renameProjectReferences(ss, oldName, newName) {
  var from = String(oldName || "").trim().toLowerCase();
  var to = String(newName || "").trim();
  if (!from || !to || String(oldName).trim() === to) return 0;
  var changed = 0;
  Object.keys(SCHEMA_SPEC).forEach(function(key) {
    var cfg = SCHEMA_SPEC[key];
    if (!cfg.projectColumns) return;
    var sheet = getTable(ss, key);
    if (!sheet || sheet.getMaxRows() < 2) return;
    var n = readRowLimit(sheet, 2000) - 1;
    if (n < 1) return;
    Object.keys(cfg.projectColumns).forEach(function(colKey) {
      var col = cfg.columns.indexOf(colKey) + 1;
      if (col < 1) return;
      var range = sheet.getRange(2, col, n, 1);
      var vals = range.getValues();
      var hit = false;
      for (var r = 0; r < vals.length; r++) {
        if (String(vals[r][0] || "").trim().toLowerCase() === from) { vals[r][0] = to; hit = true; changed++; }
      }
      if (hit) range.setValues(vals);
    });
  });
  // The Calendar's chosen project follows (ADR 0045)
  try {
    if (typeof calState === "function" && typeof calSaveState === "function") {
      var st = calState();
      if (String(st.project || "").toLowerCase() === from) { st.project = to; calSaveState(st); }
    }
  } catch (eCal) {}
  return changed;
}

/**
 * After the Projects tab lost its ID column (ADR 0053): rows that still hold
 * a former project ID ("PRJ-…") get the project name, read from the archive.
 * @return {number} cells changed
 */
function replaceProjectIdsFromArchive(ss, archive) {
  if (!archive || archive.getLastRow() < 2) return 0;
  var header = archive.getRange(1, 1, 1, archive.getLastColumn()).getValues()[0]
    .map(function(h) { return String(h || "").trim().toLowerCase().replace(/[^a-z0-9]/g, ""); });
  var idCol = header.indexOf("id") >= 0 ? header.indexOf("id") : header.indexOf("projectid");
  var nameCol = header.indexOf("project") >= 0 ? header.indexOf("project") : header.indexOf("name");
  if (idCol < 0 || nameCol < 0) return 0;
  var changed = 0;
  archive.getRange(2, 1, archive.getLastRow() - 1, archive.getLastColumn()).getValues().forEach(function(r) {
    var id = String(r[idCol] || "").trim();
    var name = String(r[nameCol] || "").trim();
    if (id && name && id !== name) changed += renameProjectReferences(ss, id, name);
  });
  return changed;
}

/**
 * Edit of a Projects name: a rename (old and new value both set) is carried to
 * Tasks, Backups and Recurring Tasks; any name change refreshes the dropdowns.
 */
function handleProjectNameEdit(e, ctx) {
  if (!e || !e.range || ctx.row < 2 || ctx.colName !== "name") return null;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
    if (typeof refreshLocationDropdowns === "function") refreshLocationDropdowns(ctx.ss);
    return "refreshed";
  }
  var oldName = e.oldValue === undefined ? "" : String(e.oldValue).trim();
  var newName = String(e.range.getValue() || "").trim();
  var changed = 0;
  if (oldName && newName && oldName !== newName) {
    var others = readProjects(ctx.ss).filter(function(p) { return p.row !== ctx.row && p.key === oldName.toLowerCase(); });
    // Another row still carries the old name: leave the references alone
    if (!others.length) changed = renameProjectReferences(ctx.ss, oldName, newName);
  }
  if (typeof refreshLocationDropdowns === "function") refreshLocationDropdowns(ctx.ss);
  if (changed) {
    try { ctx.ss.toast(oldName + " renamed " + newName + " in " + changed + " task, backup or recurring row(s).", "Project renamed", 6); } catch (eT) {}
  }
  return changed ? "renamed" : "refreshed";
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    projectCol: projectCol,
    readProjects: readProjects,
    projectNames: projectNames,
    findProject: findProject,
    renameProjectReferences: renameProjectReferences,
    replaceProjectIdsFromArchive: replaceProjectIdsFromArchive,
    handleProjectNameEdit: handleProjectNameEdit
  };
}
