# 0055. Weekly Presence is created and refreshed by the actions that change it

## Status

Accepted (2026-10-15)

## Context

The Weekly Presence tab was only built by the nightly sync (and by the "Refresh Weekly Matrix View" menu item of the Library menu, not shown in the Consumer Sheet menu). On a new workspace, Initialize / Upgrade runs that sync while Members is still empty, and the build stops with "No active members" before it creates the tab. The tab then stayed missing until the next nightly run, and a new member, routine or leave did not show in it before then.

## Decision

- Weekly Presence is created and refreshed right after the actions that change what it shows:
  - a Members or Contracts edit (after the directory sync, so departed members are hidden correctly): `editRefreshWeeklyPresence` in `src/EditHandlers.js`;
  - saving a routine in the Declare Presence & Leaves sidebar: the daily presence log is compiled, then the tab is refreshed;
  - submitting a leave (any status): the tab is refreshed (the log was already compiled for approved leaves);
  - the nightly sync still does it as a safety net.
- These refreshes are quiet (`refreshWeeklyMatrix(ss, true)`): no toast, the edit or sidebar shows its own result.
- The "Refresh Weekly Matrix View" item is removed from the Library menu. The `refreshWeeklyMatrix` forwarder stays in `consumer/Code.js` for the nightly job and Check Workspace.
- Presence declared by typing in the tab itself already updates the log, as before (ADR 0016).

## Consequences

- A routine save now also compiles the presence log; tests that compile the log themselves switch that off in their `saveMemberBaselines` wrapper.
- Members edits take a little longer (one more refresh after the ~15 s directory sync).
- Initialize / Upgrade on an empty workspace still creates no Weekly Presence tab: it appears with the first member.
