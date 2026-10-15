# ADR 0057: Overlapping Routines Are Resolved by Date Order

## Status
Accepted (2026-10-15).

## Context
Saving a routine that overlapped an existing one asked whether to "set the end date of the previous baseline to the start date of the new baseline", and then did so for every overlapping routine. When the new routine started *before* an existing one (declaring an earlier period), the later routine was given an end date before its own start (e.g. start 2026-10-15, end 2026-09-02). The previous end was also the new start day itself, so both routines covered that day.

## Decision
Routine end dates are inclusive (`checkBaselineDateOverlap`, as in the presence compile). Once the person confirms, each overlapping routine is resolved by where it starts (`planBaselineOverlaps` in `src/Sidebar.js`):
- **starts earlier** than the new routine: it ends the day before the new one starts;
- **starts later**: it is kept as is, and the new routine ends the day before it (or earlier if the person gave an earlier end);
- **same start date**: it is replaced by the new routine.
The confirmation lists what will happen to each routine. No routine ever ends before it starts, and no day has two routines.

## Consequences
- Declaring a past routine "before" the current one fills the gap without touching the current routine.
- Rows already broken by the old behaviour are fixed by hand.
