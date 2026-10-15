# ADR 0059: No Presence Is Assumed Without a Routine

## Status
Accepted (2026-10-16). Amends ADR 0040 (item 5, default location).

## Context
A weekday with no routine for a member was logged, and shown in Weekly Presence, at the first active Office. A new member who had not declared a routine yet therefore appeared to be in the office every day.

## Decision
A weekday that no routine covers and no leave falls on is **not logged** and shows empty in Weekly Presence (past and planned days alike); the member's Today shows "Not Logged" until they declare a routine or type their presence. A half-day leave on such a day logs only its half. Weekends without a routine stay Off and unlogged, as before. `routineOn` (`src/PresenceEngine.js`) returns an empty location (`none: true`) for those days. The first active Office is still the default inside the routine sidebar (presets, a day left blank when saving a routine).

## Consequences
- Presence Log rows already written with the default office stay until removed by hand (Check Workspace or the Presence Log tab).
- Routines are now read with their chip text, like Members, so the member keys always match.
