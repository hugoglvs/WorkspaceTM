# 05: Project Gantt & Active Responsibles Dashboard

**What to build:** Two critical operational views:
1. `view_project_gantt`: A timeline sheet visualizing all active projects from `tbl_projects`, their start and target completion dates, progress status, and assigned members, allowing managers to see who is allocated where.
2. `view_active_responsibles`: A live status dashboard that lists every project, its designated Project Lead, and its Project Backup. It cross-references today's presence records in `tbl_presence_log` and `tbl_leaves`: if the Project Lead is active, they are marked "On Duty"; if the Project Lead is on leave, the Project Backup is highlighted as "Active Covering Lead"; if both are on leave or absent, an alerting status "CRITICAL: No Responsible Available" is displayed.

**Blocked by:** 03: Baseline & Leave Range Compiler to Daily Presence Log

**Status:** closed

> [!NOTE]
> Per user explicit directive, custom scripted `view_project_gantt` was omitted in favor of Google Workspace's native Timeline view (`Insert > Timeline` created manually by the user from `tbl_projects`).
> The Active Responsibles Dashboard (`view_active_responsibles`) has been fully implemented, integrated into UI menus and synchronization routines, and verified via automated tests.

- [x] (Skipped per user preference) Native Timeline view to be configured manually via `Insert > Timeline` on `tbl_projects`.
- [x] Sheet `view_active_responsibles` displays:
  - Project ID & Project Name
  - Project Lead (`full_name`) & Today's Presence Status
  - Project Backup (`full_name`) & Today's Presence Status
  - Effective Active Responsible on Duty
  - Coverage Status badge ("Lead Active", "Covered by Backup", "CRITICAL: Uncovered")
- [x] Conditional formatting highlights coverage gaps in red and backup escalations in amber.
- [x] All views apply 12px "Google Sans" formatting and native People Chips.
