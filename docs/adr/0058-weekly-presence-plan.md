# ADR 0058: Weekly Presence Shows the Plan up to the Last Declared Leave

## Status
Accepted (2026-10-15).

## Context
Weekly Presence ran from the first logged day to the end of the current week (or the last logged day). The Presence Log is compiled up to today only, so future leaves and routines were invisible in the view. Prior weeks were hidden by counting "the last 4 weeks" of the range, which breaks once the range runs into the future. The current week was not marked.

## Decision
1. **Range:** the view runs to the Sunday of the latest of today, the last logged day and the **end of the last declared leave** (approved or requested; rejected and cancelled ones don't count).
2. **Plan after today:** a day after today with no Presence Log entry shows what the compile would log: the routine (or the default office on weekdays), overridden by a leave, honouring half days; public holidays and plain weekends stay empty. Requested leaves are shown too, with a note "Leave requested, not approved yet". Nothing is written to the log until someone edits a cell. The logic is shared with the compile: `readPresenceRoutines`, `readPresenceLeaves`, `routineOn`, `leaveOn`, `plannedPresence` in `src/PresenceEngine.js`.
3. **Hidden weeks:** weeks before the 3 weeks preceding the current week are hidden (not removed); the current week and every later week stay visible.
4. **Current week:** its week-number cell is highlighted (light blue, `#D2E3FC`).

## Consequences
- The view grows with the furthest leave anyone declares.
- A planned value edited by hand becomes a log entry and wins over the plan from then on.
