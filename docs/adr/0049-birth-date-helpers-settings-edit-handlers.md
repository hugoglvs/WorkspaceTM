# ADR 0049: Full Birth Date, Shared Helpers, Settings Tab, Per-Table Edit Handlers

## Status
Accepted (2026-10-13). Amends ADR 0042 (birthday stored as day and month).

## Context
A code review found organisation-specific values written in the code, the same small helpers copied in many files, and one long function handling every edit of every tab. The team also wants the Birthday column to hold a full date.

## Decision
1. **Birthday is a date column** (`DATE`). Each person decides whether the year is real. A date is kept as typed; text with a year that Sheets did not read as a date ("1990-03-14", "March 14, 1990") is stored as a date; anything else gets a toast. Reminders and the Team Birthdays calendar use day and month only; older "14 Mar" text values are still read.
2. **Shared helpers** (`src/Utils.js`): `isoDate`, `toDate`, `addDays`, `shiftIso`, `cellName`, `readTableRows`, `escapeHtml`. 24 exact copies in 12 files were replaced. Look-alikes with different behaviour (`formatDateIso` keeps unparsed text, `parseEventDate` keeps the time, `presenceIso`, `bdMidnight`, `parseDateValue`) were left as they are.
3. **Settings tab** (hidden, menu *Edit Settings*, key `settings`): `Setting | Value | Description`. Rows: default role, default department, public holidays (France or None), onboarding project ID, onboarding task title and description. Missing rows are added with their default on Initialize / Upgrade; existing values are never overwritten; an empty value means the default. Editing the tab re-applies the Members formulas. The stale location lists in `Schema.js` (OneRoof, JLL, C04, C60) were removed: those dropdowns come from the Locations tab.
4. **Per-table edit handlers** (`src/EditHandlers.js`): `EDIT_HANDLERS` maps each table key to `only` (views: Weekly Presence, Calendar), `before` and `after` handlers around the common checks (missing required cells, protected columns). Each handler runs on its own: an error is logged and the next one still runs; returning `"stop"` ends the edit. `handleSheetEdit` in `Code.js` only dispatches.

## Consequences
- Adding behaviour for a tab means adding a function to its entry, not editing a 200-line function.
- Another team changes its defaults in Settings, not in the code. Only French holidays are built in.
- Known leftover: four functions exist in two files each (`syncTasks`, `refreshWeeklyMatrix`, `compilePresenceMigration`, `debugUserIdentity`, in `Code.js` and their own module); Apps Script keeps one, and both behave the same.
