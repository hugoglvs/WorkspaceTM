# ADR 0052: Task Templates, Schedules, Reminders and the Manager Check

## Status
Accepted (2026-10-15). Amends ADR 0045 (Calendar cells, frequencies, reminders) and ADR 0051 ("no digest").

## Context
Time & expenses is due on the 15th and on the last day of each month at 10 am, peer reviews every week. People need reminding days before, and a manager wants to check that everyone did their mandatory tasks. The Assign Tasks sidebar hid the schedule ("Every period"), and the Calendar used "To do" / "Done" dropdowns where the team wants checkboxes.

## Decision
1. **Task Templates.** The "Task Types" tab is renamed **Task Templates** (key `taskTypes` unchanged; `formerNames` in `src/Tables.js` renames the tab on upgrade; menu *Edit Task Templates*). A template is the one place where a preconfigured task is defined: schedule, Mandatory, Assign to, Instructions, event (Created when), and a new **Remind** column. The Assign Tasks sidebar reads it.
2. **Twice a month.** A new frequency whose periods are half months (1st–15th, 16th–end; keys `2026-10-H1` / `-H2`). Due takes one value per half month ("15; Last day"); a weekend or public holiday moves to the working day before, as for other frequencies. Planned 6 half months ahead.
3. **The schedule belongs to the template.** The sidebar shows it (Repeats, Due, Remind) filled from the template; changing it there saves it on the template, for everyone. The recurring engine then makes the coming generated tasks follow: a new frequency cancels them and plans the new periods; a new due day moves their deadline. Tasks changed by hand (`|manual`) are never touched.
4. **Sidebar pre-fill.** Picking a template ticks the roles and names of its Assign to, selects Once (One-off) or Recurring, and shows its instructions, which become the task description (also for generated tasks).
5. **Remind** lists the working days before the due date when the person is emailed by the 9 am job ("3; 1"; 0 = the day itself; blank = 3). The late email stays. Since the daily trigger runs between 9 and 10 am, a deadline at 10 am should be reminded the day before (1), not on the day (0).
6. **Manager check.** A **Manager role** setting (default "Manager"): the working day after a mandatory task is due, every member with that role is emailed who has not done it (`sendManagerDigest`). The Calendar's done row counts the tasks done per period ("9/12", green when all done, red when a past period is incomplete).
7. **Calendar checkboxes.** A cell with a task holds a checkbox (ticked = done). Inserting a checkbox in an empty cell assigns the task, ticking marks it done, unticking reopens it, deleting the checkbox cancels it. Typed "To do" / "Done" still work. Row 2 shows the year above its periods (merged per year, labels without the year), row 3 the periods, row 4 the done count; members start at row 5.
8. **Defaults.** Time & expenses: twice a month, "15; Last day", Remind "3; 1". Peer review: weekly on Friday, Remind "1". An existing sheet whose two rows still hold the former default (Monthly, Last working day) gets these once.

## Consequences
- Existing recurring Time & expenses and Peer review rules switch to the new periods at the next generation; their coming monthly tasks are cancelled and replaced.
- One frequency per template keeps the Calendar one grid per template; a team needing two schedules for the same work makes two templates.
