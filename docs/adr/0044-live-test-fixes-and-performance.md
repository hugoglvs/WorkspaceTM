# ADR 0044: Fixes From the First Live Test, and Edit / Nightly Performance

## Status
Accepted (2026-10-12)

## Context
The first test in the real Sheet (Claude driving the browser) showed:
1. The nightly job of 2026-10-12 hit Apps Script's 6-minute limit ("Exceeded maximum execution time"), with no log of which step was slow.
2. One Members edit took 30–45 s. `applyProtections` removed and recreated every protection of every table on each Members edit, sidebar save and compile; the Members sync also re-copied the people chip into every Routine row.
3. A birthday typed in Members was not rewritten: `setNumberFormat` is refused in a typed (native table) column, and the failure surfaced later in the same execution.
4. Widening Members to add Birthday made Sheets stretch the "required cell" rule of *Today* onto *Birthday*, so an empty (optional) birthday showed red; each upgrade also stacked another copy of the required-cell rules.
5. Typing a new name in Members (the way to add someone) raised the blocking "Strictly Protected Column Edited" dialog first. Deleting a name left its Status in the row.
6. Check Workspace and Interim Backups take 15–25 s to compute; the dialog stayed blank until done.

## Decision
1. **Nightly steps are timed and budgeted**: each step logs its duration; once 5 minutes have passed, the remaining steps are skipped and listed. The coverage check now runs before Google Tasks / Calendar sync.
2. **Protections are incremental**: a protection already on the right range (same description, rows, column, warning mode) is kept; only missing or changed ones are recreated, stale ones removed. The compile only reapplies them when it created rows.
3. **Edits run a light sync** (`syncMembersAndTenures(ss, { fromEdit: true })`): no chip copy into Routines and no protections; the nightly sync does both. A Members edit now logs `[edit timing]` when it takes over 3 s.
4. Birthday rewrite: no number format; the typed TEXT column keeps text as typed.
5. Required-cell rules are replaced, not appended, on each upgrade (`isRequiredCellRule`).
6. A new name in an empty Members row skips the protected-column dialog; deleting a name clears the rest of its row.
7. Check Workspace opens at once with "Checking…" and loads the report in the dialog.

## Consequences
- Measured after the fixes: Members edit 45 s → ~15 s (first run; the rest is the Members sync and the Sheets round trips), nightly ~70 s on the test data.
- A contract row typed directly in the hidden Contracts tab gets its warning protection at the next nightly run.
- Confirmed live: the installable edit trigger handles edits (simple `onEdit` returns in under 1 s), toasts and dialogs from it reach the editor, and the Members & Contracts sidebar now opens by itself on a new name (issue 22 resolved).
