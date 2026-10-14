# ADR 0051: Task Model, Tasks Columns and Task Triggers

## Status
Accepted (2026-10-13); rows of a moved table are copied into the rebuilt one since ADR 0053. Amends ADR 0045 (Done by dropped, Recurring renamed Origin) and ADR 0049 (onboarding settings moved to Task Types).

## Context
The Tasks tab had grown to 13 columns, several technical, one redundant ("Done by" duplicated "Assigned to" and was unreliable because edits run under the trigger owner's account). Tasks also come from several sources: typed by hand, recurring rules, and events (only onboarding, written in the code with its texts in Settings). The team wants an easy way to add new automatic tasks later. Completion is visible to everyone in the sheet, so no digest is sent; Google Tasks stays optional per person.

## Decision
1. **One row = one task for one person with one deadline**, whatever created it.
2. **Tasks columns**, in this order: Type · Title · Assigned to · Project ID · Planned date · Deadline · Status · Description · Done on, then hidden: ID · Google Task ID · **Origin**. "Done by" is dropped. Origin says where the row comes from: `<rule>|<period>` (recurring, `|manual` once changed in the Calendar), `manual` (typed or assigned once), or `trigger:<event>:<subject>:<type>`. The schema gains `hiddenColumns`, applied on every Initialize / Upgrade.
3. **Moved columns**: when a release reorders or removes columns of a table that holds data (`columnsMoved` in `src/Init.js`, using `formerHeaders` / `removedHeaders`), Initialize / Upgrade renames the old tab to "Archive – <tab> <date>", hides it and builds the table again. Columns added at the end, or labels renamed in place, do not count as a move.
4. **Task triggers** (`src/TaskTriggers.js`): `TASK_TRIGGERS` lists the events the app notices; adding one is one entry plus one `fireTaskTrigger(ss, key, {subject, date, person, projectId})` call where the event is noticed. The Task Types tab says what each event creates, with four new columns: **Created when** (dropdown of trigger labels), **Days after** (relative to the event date, negative = before), **Assign to** ("Person concerned", "Role: X" or names), **Instructions** (the task description). Each task is created once per event (Origin checked).
5. **New member** is the first trigger: a person whose first contract starts today or later (or less than 30 days ago, to catch up), fired by the members sync and when the Contracts sidebar saves a first contract. Onboarding is a Task Types row (New member, 7 days after, Person concerned) instead of code plus three Settings rows, which upgrade clears. An existing Onboarding row gets these values once.
6. The interim-backup task created by the coverage check (ADR 0038) keeps its built-in rule and is marked `trigger:coverageGap`.

## Consequences
- 9 visible columns instead of 13; technical ones can still be shown (Sheets "Unhide columns").
- The live Tasks tab of an existing sheet is archived and rebuilt on the first upgrade after this release; its rows stay readable in the archive.
- Recurring tasks plan 3 months ahead (unchanged); completion shows in the Calendar and Tasks tabs for everyone.
