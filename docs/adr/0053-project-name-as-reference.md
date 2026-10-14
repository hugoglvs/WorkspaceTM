# ADR 0053: The Project Name Is the Reference

## Status
Accepted (2026-10-15). Amends ADR 0038 (project IDs) and ADR 0051 (moved tables).

## Context
Projects had a generated ID ("PRJ-ABC12") that Tasks, Backups and Recurring Tasks stored in a "Project ID" column. Nobody reads or types these IDs; people know projects by name, as they know members by name.

## Decision
1. **No project ID.** The Projects tab starts with **Project** (the name). Tasks, Backups and Recurring Tasks have a **Project** column holding the name, a dropdown of the Projects tab's names (schema `projectColumns`, filled by `schemaValidationsFor` from `projectNames`, `src/Projects.js`). The dropdowns are refreshed on Initialize / Upgrade and whenever a project name is edited.
2. **Renaming follows.** Editing a project's name (`handleProjectNameEdit`, Projects entry of `EDIT_HANDLERS`) replaces the old name in every Project column (`renameProjectReferences`), and the Calendar's chosen project. If another Projects row still carries the old name, nothing is renamed.
3. **Names are unique.** A repeated name turns red in Projects (schema `uniqueColumns`, conditional formatting). Comparisons are case-insensitive.
4. **Projects are read by column key** (`readProjects` / `findProject`), never by position.
5. **Moved tables keep their rows.** When a release moves columns of a table holding data (ADR 0051), Initialize / Upgrade still archives the old tab, then copies its rows into the new table column by column, matched on header (`carryArchivedRows`), which keeps people chips. For Projects, former IDs still in Tasks, Backups and Recurring Tasks are replaced by the names read from the archive (`replaceProjectIdsFromArchive`).

## Consequences
- The upgrade after this release archives and rebuilds Projects (its first column is gone); the rows come back without retyping.
- A Project cell holding a name that is no longer in Projects shows as invalid in its dropdown.
