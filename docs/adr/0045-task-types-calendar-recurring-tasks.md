# ADR 0045: Task Types, Calendar, Recurring Tasks, Audit Trail and Per-Person Google Tasks

## Status
Accepted (2026-10-12); Done by dropped and Recurring renamed Origin by ADR 0051.

## Context
The team keeps two hand-made sheets: a quiz calendar (one row per member, one column per week, a tick for who is in charge) and a "Team tasks" checklist of mandatory monthly tasks (peer reviews, SoWs, PRMs, forecasts, time & expenses) that are audited for the quality certification. Tasks only had a due date: no type, no recurrence, no completion record. Google Tasks sync only worked for the sheet owner, because a script can only write into the Google Tasks of the account it runs as.

## Decision
1. **Task Types** (hidden tab, menu *Edit Task Types*): `Type | Frequency (Weekly, Monthly, One-off) | Due | Mandatory | Active`. Due is a weekday for weekly types, a day number / "Last working day" / "First working day" for monthly ones; weekends and public holidays move to the working day before. Defaults: Onboarding, Quiz, Team presentation, Peer review, SoW, PRM, Forecast, Time & expenses (the last five mandatory). Tasks gets **Type**, **Done on**, **Done by**, **Recurring** columns and a **Cancelled** status.
2. **Calendar** (visible tab, like Weekly Presence): pick a Type (B1) and a Project (D1); rows are members, columns are weeks (months for monthly types), 4 weeks / 3 months back and 12 weeks / 8 months ahead. Cells: *To do* (assigned / in charge; red when overdue), *Done* (green), empty (grey on approved leave, orange on requested leave). Editing writes Tasks: *To do* assigns — **refused on approved leave** (on leave means not working; remote is for working elsewhere), a warning on a requested leave; *Done* marks done; clearing cancels. The selection is shared by everyone (accepted: small team).
3. **Recurring Tasks** (hidden tab, created from the *Assign Tasks* sidebar): `ID | Type | Project ID | Who | Applies to | Start | End | Active`. *Each person*: one task per person in "Applies to" every period. *Rotation*: one person per period, the one whose last turn is the oldest, skipping anyone on approved leave on the due date (so a skipped person goes first next time). "Applies to" mixes names and roles ("Role: Consultant; Alice"); blank = every member. The nightly job plans 8 weeks / 3 months ahead; a rotation turn that later falls on an approved leave moves to the next person and both are emailed. A turn changed by hand in the Calendar is marked `|manual` and never moved.
4. **Assign Tasks sidebar**: a type and project to several people at once, by role and/or name, once (a chosen week/month) or every period (a rule); lists rules and stops them (future open tasks cancelled).
5. **Audit trail**: Done on / Done by are stamped when a task becomes Done (Calendar, Tasks grid or Google Tasks) and cleared if reopened. Changing a task or a Calendar cell of a period before the current month asks for confirmation.
6. **Reminders** (daily 9 am job, renamed `dailyRemindersJob`, birthdays included): an email 3 working days before the due date and on the first working day after it if still open. Check Workspace adds *Overdue tasks*, *Tasks falling on a leave* and *Not connected to Google Tasks*.
7. **Per-person Google Tasks**: menu *Connect my Google Tasks* installs an hourly `myTasksSyncJob` under the clicking person's account. Each person syncs only the tasks assigned to them into their own "Project Management Tasks" list (list ID in user properties), both ways; a task reassigned away or cancelled is removed from the old list, tasks the app made that no row points to are cleaned up, and lists are read page by page.

## Consequences
- Tasks gains ~75 rows a month for a 15-person team with 5 mandatory monthly types; `findNextAvailableRow` no longer stops at row 500.
- Every member must click *Connect my Google Tasks* once and authorise; until then their tasks are only in the sheet and the reminder emails.
- Seven visible tabs: Members, Weekly Presence, Projects, Tasks, Calendar, Events.
