# 0056. Routine sidebar: one location per day, AM / PM on request

## Status

Accepted (2026-10-15)

## Context

The Routine Baseline form of the Declare Presence & Leaves sidebar always showed two selects (AM and PM) for every weekday, although most people work from one place all day.

## Decision

- Each weekday shows a single location select ("whole day") by default.
- A "Split AM / PM" link on the day shows the AM and PM selects; "Whole day" switches back (the AM location then applies to both halves).
- A saved routine whose AM and PM differ opens split; presets and routines with equal halves open as whole days.
- Nothing changes in the data: the payload still carries `location_am` and `location_pm` per day, so the Routines table, the presence compiler and the server code are untouched.

## Consequences

- The form is shorter by default; half-day routines stay one click away.
- The sidebar HTML has no automated test (no browser in the test suite); the behaviour was checked with a DOM harness during development.
