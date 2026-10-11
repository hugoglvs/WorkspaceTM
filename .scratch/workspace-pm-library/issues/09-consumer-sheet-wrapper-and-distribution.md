# 09: Consumer Sheet Wrapper & Distribution Package

**What to build:** The lightweight container-bound client script and setup distribution guide for other teams and departments. It provides the minimal code needed inside a team's spreadsheet to hook into the versioned Library, create the custom Google Sheets UI menu (`onOpen`), and invoke initialization, migration, sync, and dashboard refresh actions. It includes a clear README for team leads with step-by-step instructions on pinning library versions and granting one-time OAuth permissions.

**Blocked by:** 04: Weekly Matrix Planning View, 05: Project Gantt & Active Responsibles Dashboard, 06: Presence & Remote Work KPI Engine, 07: Google Tasks API Synchronization, 08: Google Calendar & Event Invitations Synchronization

**Status:** closed

- [x] A thin container-bound script template providing the `onOpen()` menu:
  - "Project Management" -> "Initialize / Upgrade Spreadsheet"
  - "Project Management" -> "Compile Daily Presence Logs"
  - "Project Management" -> "Refresh Dashboards (Weekly Matrix & Responsibles)"
  - "Project Management" -> "Sync Google Tasks"
  - "Project Management" -> "Sync Calendar Events"
- [x] Error handling surfaces actionable toast messages or modal alerts to the spreadsheet user.
- [x] Distribution guide (`docs/distribution-guide.md`) detailing:
  - How to link the Library Script ID in Apps Script.
  - How to pin the dependency to an immutable numbered version.
  - How to configure Document Properties (Calendar ID, Tasks List ID).
  - How to run the one-time OAuth authorization.
