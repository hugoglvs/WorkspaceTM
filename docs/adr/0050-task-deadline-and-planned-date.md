# ADR 0050: Task Deadline and Planned Date, Aligned with the Google Tasks API

## Status
Accepted (2026-10-13)

## Context
The Tasks tab had one "Due date", pushed to Google Tasks as the task's date and used for reminders, overdue and the audit. The Google Tasks app now separates a date (with optional time and repeat) from a deadline, but the Google Tasks API exposes only `title`, `notes`, `status`, `completed` and one date field, `due`, documented as the day the task is scheduled, "not the deadline"; only the date is kept. There is no deadline, time or repeat field in the API.

## Decision
1. The Tasks column `due_date` is now labelled **Deadline**: reminders, overdue, the Calendar view, Check Workspace and recurring tasks keep using it.
2. A new column **Planned date** (`planned_date`, date, optional, last column) is the day sent to Google Tasks (`due`). Empty, the Deadline is sent instead. The sheet stays the master: a date moved in Google Tasks is put back by the next sync.
3. The Google Task notes carry the deadline (`Deadline: 2026-10-30`) between the description and the sheet references.
4. Google's built-in repeat is not used, and its "a task with a deadline cannot repeat" rule is not copied: recurring tasks are one row per occurrence, each with its own deadline, which the audited monthly tasks need.
5. Times are not stored: the API drops them.

## Consequences
- In Google Tasks, the deadline shows in the notes, not as Google's own deadline badge (not settable through the API).
- Planned date sits at the end of the Tasks table, so existing columns keep their positions.
