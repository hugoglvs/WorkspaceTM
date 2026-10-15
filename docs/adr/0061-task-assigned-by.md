# ADR 0061: Tasks Record Who Assigned Them

## Status
Accepted (2026-10-16). Extends ADR 0051 (task model).

## Context
A manager wants to see who created a task. ADR 0051 dropped "Done by" because the sheet's edit trigger runs as the script owner and cannot always tell who edited. Who creates a task is known in some cases and not in others.

## Decision
Tasks gain an **Assigned by** column (key `assigned_by`, people chip, editable), placed **last** so adding it is not a column move and the table is not rebuilt (ADR 0051, 0054). It is filled when the task is created:

| How the task is created | Assigned by |
|---|---|
| Assign sidebar | the person using the sidebar: their Members chip, else their email |
| Calendar checkbox | the person who ticked, when Google gives their identity (best effort) |
| Row typed in Tasks | the person typing, when Google gives their identity (best effort), when the row gets its ID |
| Recurring rules, task triggers, coverage gaps | "Automatic" |

`currentTaskAuthor(ss)` and `setTaskAuthor` (`src/RecurringTasks.js`) resolve and write it; `createTaskRow` takes `assignedBy`. When Google does not say who the user is (edit triggers for accounts outside the owner's Workspace domain), the cell stays empty and anyone can pick themselves. A reassigned rotation turn keeps "Automatic".

## Consequences
- Existing tasks have an empty Assigned by.
- The column appears after Done on (the hidden ID columns sit in between); run Initialize / Upgrade once to add it.
