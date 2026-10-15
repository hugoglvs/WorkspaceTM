@AGENTS.md

## Project facts

- **WorkspaceTM**: a Google Apps Script **Library** (V8, `src/`) consumed by team **Consumer Sheets** through the thin wrapper `consumer/Code.js`. Deployment: `docs/distribution-guide.md`.
- No build, no `package.json`. Each `src/*.js` and `consumer/Code.js` is a global script in Apps Script and ends with a `if (typeof module !== 'undefined') module.exports = {...}` guard so Node tests can `require` it. Keep that guard when adding files or exports.
- Tests: `node --test test/*.js`, all GAS services mocked in `test/mockGas.js` (add new service mocks there). Baseline 2026-10-15: **262 pass, 0 fail**.
- Spec and tickets: `.scratch/workspace-pm-library/` (`spec.md`, `issues/NN-*.md`). Status words in use: `resolved`/`closed` = done, `open`, `ready-for-agent`.
- Every new decision lands as a `docs/adr/NNNN-*.md` (next number: 0056); new domain terms land in `GLOSSARY.md`.
- Product principles (ADR 0034): every grid stays editable, and there is no privacy model; data tables are hidden for declutter, never locked; sidebars are an extra entry path; KPIs go to Looker Studio through `fact_presence`.
- **Never hard-code a tab name** (ADR 0039): look tabs up with `getTable(ss, "<key>")` / `tableKeyOfSheet(sheet)` from `src/Tables.js`; keys: locations ("Locations"), members, tenures (tab "Contracts"), baselines ("Routines"), taskTypes ("Task Templates"), dropdowns ("Dropdowns"), leaves, presenceLog, projects, tasks, events, backups, weeklyMatrix ("Weekly Presence"), kpiData ("KPI Data"). `SCHEMA_SPEC` is keyed the same way; visible tables have readable `headers` labels, so match header text normalised (key or label).
- Shared helpers live in `src/Utils.js` (`isoDate`, `toDate`, `addDays`, `shiftIso`, `cellName`, `readTableRows`, `escapeHtml`; loaded as globals in tests by `test/mockGas.js`): use them instead of writing new date or name helpers (ADR 0049).
- Organisation-specific values come from the Settings tab via `getSetting(ss, key)` (`src/Settings.js`); add a row to `SETTINGS_DEFAULTS` rather than writing such a value in code (ADR 0049).
- Edits are dispatched per table by `EDIT_HANDLERS` in `src/EditHandlers.js` (`only` for views, `before` / `after` around the common checks): add a handler function to the table's entry, never a branch in `handleSheetEdit` (ADR 0049).
- Tasks rows are read and written by column key (`taskCol`, `readTaskRows`, `createTaskRow` in `src/RecurringTasks.js`), never by position; the column order changed in ADR 0051. Tasks created by an event go through `fireTaskTrigger` (`src/TaskTriggers.js`): add an event to `TASK_TRIGGERS`, the rest is configured in Task Templates. Periods come from `periodOf` / `periodShift` / `periodFrequency` (`src/TaskTypes.js`: Weekly, Twice a month, Monthly); never test `=== "Monthly"` by hand (ADR 0052).
- Members Role, Department and Status are formulas reading Contracts (ADR 0047, `src/Contracts.js`): never write values there; `applyMemberFormulas` (run by the members sync) writes them, and the mock evaluates them with `contractProfile`.
- Projects have no ID (ADR 0053): the name is the reference in every `project` column (dropdown via schema `projectColumns`); read Projects with `readProjects` / `findProject` (`src/Projects.js`), never by position.
- Chip style and option colours cannot be set by code (ADR 0054): a view (non-table) dropdown that needs them gets a model cell in the Dropdowns tab (`DROPDOWN_LISTS`, `src/Dropdowns.js`) applied with `applyDropdown`; tables ignore pasted dropdowns, so table colours are set in the table and `createGoogleSheetsTable` leaves unchanged columns alone (`tableColumnsSignature`).
- Work Locations come from the Locations tab via `src/Locations.js` (`activeLocationNames`, `allLocationNames`, `locationCategory`, `defaultOfficeLocation`); never hard-code a location (ADR 0040).
- Visible tabs are Members, Weekly Presence, Projects, Tasks and Events (ADR 0036, `src/Layout.js`); every other Library sheet is hidden on upgrade. KPIs exist only as `fact_presence` for Looker.
- Sidebars served by the Library call server functions through `google.script.run`, which reaches the **Consumer Sheet** script: every new sidebar server function needs a forwarder in `consumer/Code.js`, and both projects must be pushed. A `.js` and an `.html` file cannot share a base name in Apps Script (`MembersService.js` + `MembersSidebar.html`).
- All sync writes run under the **Sync Lock** (ADR 0032): new write paths triggered by edits or schedules go through `runWithEditLock` / `acquireSyncLock` in `src/Code.js`.
- Apps Script projects (owner `hugoglvs42@gmail.com`): Library `1vyJiEQTSluHaPbh8-B9gSOj2SeciqPEzHpJ5Xug0I_NVhDvtYHdZ1b6v` ← `src/`; Consumer Sheet "[NEW] ACN" (`1FOVefnly_zusD1TobiPAYrUEeAR7pLSAhIPxaXMFme0`) bound script `1IVORKENsN2r81QcA5hXa5qa-tgXub2RFFlhzD0-8fopojfAoQFixFujH` ← `consumer/`. Each folder has its own `.clasp.json`.
- **Deploy** (ADR 0033) from the shell on the user's Mac: `cd src && ~/.npm-global/bin/clasp push` (library), `cd consumer && ~/.npm-global/bin/clasp push` (wrapper). Run the tests, commit, tag the commit `deploy-YYYYMMDD-HHMM`, then push without asking: the user tests live and asks for a rollback when needed (redeploy an earlier `deploy-*` tag). The Consumer Sheet uses the library in development mode (HEAD), so a library push is live at once. `scripts/clasp-setup.sh` sets clasp up in the user's macOS Terminal and clones the live code into the git-ignored `.scratch/live/`.
## Working with the user

When the user suggests an idea, give an honest opinion first (merits, costs, alternatives) and wait for the decision before building it. Once a decision is made, implement, test and push without asking.

## Reporting to the user

The final message of every change lists each modified file as a direct link (`[path](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/<path>)`), with the changes applied to it as sub-bullets.
