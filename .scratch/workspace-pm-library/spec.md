# Specification: Google Workspace Project Management Library

Status: ready-for-agent

## Problem Statement

Teams across the organization currently track their rosters, member tenures, attendance, project tasks, and social events using fragmented, disparate tools or ad-hoc spreadsheets. This results in manual data duplication, inconsistent KPI calculations (such as remote work and in-office presence rates), lack of synchronization with individual Google Tasks and Google Calendars, and high operational overhead when maintainers attempt to deliver bug fixes or schema updates to multiple spreadsheets without corrupting team data.

## Solution

A centralized Google Apps Script Library that powers self-contained Consumer Sheets across different departments. The Library provides an initialization and automated schema migration engine, creates a normalized 3-tier structure (Pure Database tables, Analytical Aggregations, and Interactive Dashboards/Planning views), enforces role-based column protections, and offers scheduled or menu-driven synchronization with Google Tasks, Google Calendar, and Gmail notifications.

## User Stories

1. As a Team Lead, I want to initialize a blank Google Sheet using a single menu command, so that all required tables, headers, data validations, and protected ranges are generated automatically without manual configuration.
2. As a Team Lead, I want my spreadsheet to safely upgrade when a new version of the Library is deployed, so that my existing records and custom entries are never overwritten or lost.
3. As an Operations Manager, I want to maintain a centralized Member directory where current role and department are looked up from active Tenures, contractual status is evaluated from tenure dates, and real-time operational availability is evaluated from daily presence and leaves, so that employment data is never ambiguous.
4. As an HR Specialist, I want to record Member Tenures with role, department, contract types, and FTE rates, so that project capacity planning reflects accurate part-time and full-time allocations and role history.
5. As a Team Member, I want to declare my regular presence baselines using an integrated Google Sheets sidebar panel, so that I can set my weekly AM/PM routine without editing raw spreadsheet rows.
6. As a Team Member, I want to record my morning (AM) and afternoon (PM) Work Locations independently for each day, so that multi-site mobility and partial remote days are accurately captured.
7. As a Team Lead, I want presence and remote work KPIs calculated with half-day granularity (0.5 weighting per period), so that intra-day transitions reflect accurate attendance rates.
8. As a Team Lead, I want all data sheets formatted using Google Sheets Named Tables with consistent 12px "Google Sans" typography, so that tables are visually clean, modern, and provide native table filtering and typed headers.
9. As a Project Manager, I want to assign Tasks within Projects to specific Members with target due dates, so that delivery accountability is clear.
10. As a Team Member, I want my assigned Tasks in the sheet to automatically appear in my personal Google Tasks, so that I can track my daily to-dos in my native Google Workspace workflow.
11. As a Project Manager, I want changes made to Tasks in Google Tasks (such as task completion) to reflect back in the sheet, so that the status is synchronized across platforms.
12. As an Event Organizer, I want to log Social Events with dates and locations in the sheet, so that team gatherings are tracked alongside operational milestones.
13. As an Event Organizer, I want Social Events created in the sheet to be pushed directly to a shared Google Calendar with automated RSVP tracking, so that team members receive calendar invitations.
14. As a Team Member, I want computed and cross-referenced columns (such as member names and department lookups) to be read-only, so that I do not accidentally break lookup formulas or database keys.
15. As a Manager, I want a dedicated Weekly Matrix View where each row is a weekday (with week number and date) and each member has paired AM/PM location cells, so that weekly team deployments are visible at a glance in a table format.
16. As a Project Manager, I want a Project Gantt chart that visualizes project durations and member assignments over time, so that resource allocation across concurrent initiatives is clear.
17. As a Team Member, I want to declare scheduled leaves using start and end date ranges, so that I don't have to log every day individually from scratch.
18. As an Operations Manager, I want each project to have a designated Project Lead and Project Backup, so that accountability is clear during absences.
19. As a Manager, I want an Active Project Responsibles view that dynamically displays the current duty status (Lead on duty, Backup covering, or Alert: Neither available) based on today's presence and leaves, so that project continuity is guaranteed.
20. As an Organization Admin, I want all core logic to reside in an immutable versioned Library, so that multiple departments can adopt the system independently without diverging codebases.
21. As a newly onboarded Team Member, I want an onboarding Task automatically assigned to my People Chip in Google Tasks / tbl_tasks to enter my Accenture email and phone number in the Team Directory, so that team contact records are promptly completed.
22. As an Operations Manager, I want table schemas to strictly enforce required columns (such as member, role, department, and start date in tenures) while explicitly allowing nullable columns (such as end date for permanent tenures, or contact fields prior to onboarding), so that data integrity is maintained without restricting open-ended tenures or unassigned backlog tasks.
23. As a Spreadsheet Admin, I want table schemas to delineate strictly protected columns (preventing regular users from editing and alerting admins upon edit) from warning-protected columns (prompting users with confirmation warnings), so that computed and synchronized fields are safeguarded while authorized business key updates remain possible.
24. As an Operations Manager, I want protections on operational tables to apply only once a row is created, so that registering a new tenure on an empty row is completely unhindered by warnings or dialogs.
25. As a Team Member declaring presence baselines, I want new submissions to be logged as appended rows without overwriting previous routines, and if date ranges overlap, I want the system to prompt me to adjust the previous baseline's end date to the new start date, so that routine history is safely maintained without scheduling ambiguity.

## Implementation Decisions

### Modular Architecture
The Library is organized into distinct logical modules:
- **`SchemaEngine`**: Declares canonical table definitions, column orders, data validation rules, protected column lists, and table formatting styles (12px "Google Sans", alternating row shading, and Google Sheets Named Table conversions).
- **`MigrationManager`**: Reads the `SCHEMA_VERSION` stored in `PropertiesService.getDocumentProperties()` and executes ordered, non-destructive migrations to append missing columns, format new fields, and adjust protections.
- **`ProtectionManager`**: Applies Google Sheets `ProtectedRange` permissions, ensuring primary keys and formula-driven lookup columns are restricted to editors/admins while operational columns remain editable by team members.
- **`SidebarEngine`**: Provides the interactive Presence Baseline Sidebar via `HtmlService.createHtmlOutputFromFile`, allowing users to select recurring weekday AM/PM locations through a friendly panel that writes to `tbl_presence_baselines`.
- **`PresenceEngine`**: Compiles declarative `tbl_presence_baselines` and `tbl_leaves` ranges into concrete daily rows in `tbl_presence_log` while permitting ad-hoc member edits.
- **`ViewBuilder`**: Generates and refreshes presentation sheets:
  - **`view_weekly_matrix`**: Rows = weekdays (`week_number`, `date`); Columns = dynamically generated paired columns for each Member (`[Member] AM`, `[Member] PM`).
  - **`view_project_gantt`**: Formatted timeline view mapping `tbl_projects` start/target dates and assigned members.
  - **`view_active_responsibles`**: Dynamic status board evaluating whether the Project Lead is active or on Leave today, automatically elevating the Project Backup to active duty.
- **`TaskSyncService`**: Interfaces with the Google Tasks API (`Tasks.Tasks`), creating and updating tasks linked by external IDs (`google_task_id`).
- **`CalendarSyncService`**: Interfaces with Google Calendar (`CalendarApp`), generating events linked by external IDs (`google_calendar_event_id`) and managing invitation dispatches.
- **`KpiEngine`**: Populates or formats analytical views and dynamic aggregations (`calc_kpis`) factoring 0.5 FTE weighting per AM/PM period.

### Typography and Named Table Standard
- **Font Specification**: All data ranges, headers, and dashboard text initialized or migrated by the Library are programmatically styled to font family `"Google Sans"` and font size `12` (`setFontFamily("Google Sans")`, `setFontSize(12)`).
- **Named Tables Integration**: Each operational sheet tab (`tbl_*`) converts its active data range into a Google Sheets Named Table with typed columns (`PEOPLE_CHIP`, `DATE`, `DROPDOWN`, `DOUBLE`, `TEXT`).

### Schema Definition Format
The schema is strictly decoupled from spreadsheet manipulation:
```javascript
const CURRENT_SCHEMA_VERSION = 1;

const SCHEMA_SPEC = {
  tbl_members: {
    columns: ['member', 'accenture_email', 'phone_number', 'current_role', 'current_department', 'tenure_status', 'today_presence'],
    protected: ['member', 'current_role', 'current_department', 'tenure_status', 'today_presence']
  },
  tbl_tenures: {
    columns: ['member', 'role', 'department', 'start_date', 'end_date', 'contract_type', 'fte_rate'],
    protected: [],
    validations: { contract_type: ['Permanent', 'Contractor', 'Apprentice', 'Intern'] }
  },
  tbl_presence_baselines: {
    columns: ['baseline_id', 'member', 'start_date', 'end_date', 'day_of_week', 'location_am', 'location_pm'],
    protected: ['baseline_id'],
    validations: {
      day_of_week: ['All', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'],
      location_am: ['Remote', 'Main Office', 'Downtown Office', 'Branch Office', 'Client Site'],
      location_pm: ['Remote', 'Main Office', 'Downtown Office', 'Branch Office', 'Client Site']
    }
  },
  tbl_leaves: {
    columns: ['leave_id', 'member', 'start_date', 'end_date', 'leave_type', 'status'],
    protected: ['leave_id'],
    validations: {
      leave_type: ['Vacation', 'Sick Leave', 'Maternity/Paternity', 'Unpaid Leave'],
      status: ['Requested', 'Approved', 'Cancelled']
    }
  },
  tbl_presence_log: {
    columns: ['log_id', 'date', 'member', 'location_am', 'location_pm', 'timestamp'],
    protected: ['log_id', 'timestamp'],
    validations: { 
      location_am: ['Remote', 'Main Office', 'Downtown Office', 'Branch Office', 'Client Site', 'Vacation', 'Sick Leave', 'Holiday'],
      location_pm: ['Remote', 'Main Office', 'Downtown Office', 'Branch Office', 'Client Site', 'Vacation', 'Sick Leave', 'Holiday']
    }
  },
  tbl_projects: {
    columns: ['project_id', 'name', 'status', 'start_date', 'target_date', 'lead_member', 'backup_member'],
    protected: ['project_id'],
    validations: { status: ['Planning', 'Active', 'On Hold', 'Completed'] }
  },
  tbl_tasks: {
    columns: ['task_id', 'project_id', 'assigned_member', 'title', 'due_date', 'status', 'google_task_id'],
    protected: ['task_id', 'google_task_id'],
    validations: { status: ['Todo', 'In Progress', 'Done', 'Blocked'] }
  },
  tbl_events: {
    columns: ['event_id', 'title', 'date', 'location', 'google_calendar_event_id'],
    protected: ['event_id', 'google_calendar_event_id']
  }
};
```

### Migration Safety Rules
- Migrations must be purely additive: new columns are appended to the right of existing data to prevent shifting user-entered columns.
- Before executing schema modifications, existing table ranges are verified for column headers.
- If a table already contains data, `freshInit()` is skipped in favor of `runMigrations(currentVersion, targetVersion)`.

### External Sync Contracts
- External IDs (`google_task_id`, `google_calendar_event_id`) are immutable foreign keys written upon creation to prevent duplicate event or task generation.
- Batch operations must be used for all spreadsheet reads and writes to avoid hitting Apps Script service quotas.

## Testing Decisions

- **Testing Seam**: The public boundary of the Library (`setupOrUpgrade()`, `syncTasks()`, `syncEvents()`).
- **Migration Idempotency**: Tests must verify that running `setupOrUpgrade()` multiple times on a populated sheet produces identical headers, preserves all existing row values, and updates `SCHEMA_VERSION` accurately.
- **Protection Integrity**: Tests must verify that protected ranges accurately lock specified columns while preserving write permissions on user-fillable fields.
- **External API Decoupling**: Google Workspace API interactions (Tasks, Calendar, Gmail) must be abstractable or tested with mock adapters so unit tests can execute without live Workspace side effects.

## Out of Scope

- Building a standalone web portal outside of Google Sheets / Google Workspace.
- Bi-directional sync for task subtasks or complex recurring calendar event patterns in v1.
- Custom payroll or time-billing calculations (only FTE and presence status logging are supported).

## Further Notes

- Consumer spreadsheets will invoke this Library via a thin menu wrapper (`onOpen`).
- Team-level configurations (e.g. primary Google Calendar ID, default Task List ID) will be stored in `PropertiesService.getDocumentProperties()` so each team can point to their own calendar/task list.
