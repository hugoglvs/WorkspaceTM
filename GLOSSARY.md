# Project Management Workspace

A Google Workspace-based module for managing teams, presence tracking, project tasks, and social events using Google Sheets, Apps Script, and Workspace APIs.

## Language

**Member**:
An individual person belonging to a team or department, uniquely identified natively by their Google People Chip (`PEOPLE_CHIP`).
_Avoid_: User, Employee, Staff, Resource

**Team Directory**:
The central roster of Members (`tbl_members`) containing contact details (`accenture_email`, `phone_number`), current role, department, contractual tenure status, and today's presence.
_Avoid_: Contact List, Phonebook, Staff Directory

**Tenure**:
The active contract period, role, department, and FTE parameters of a Member's assignment, identified naturally by the Member's People Chip and date range.
_Avoid_: Contract, Duration, Employment Period, Tenure ID

**Tenure Status**:
The contractual employment state of a Member (`Active`, `Upcoming`, `Departed`, `No Contract`), evaluated from their Tenure dates by a formula in the Members Status column (ADR 0047).
_Avoid_: Employment Status, Employee State

**Today Presence**:
The real-time operational availability and Work Location of a Member for the current day, evaluated from Leaves, Presence Logs, and Events.
_Avoid_: Daily Status, Attendance State

**Presence Log**:
The ground-truth database (`tbl_presence_log`) capturing a Member's actual recorded presence and Work Location for morning (AM) and afternoon (PM) periods up to the current date. Future planning is dynamically projected on presentation views, keeping the presence log strictly historical and current.
_Avoid_: Attendance, Timesheet, Check-in, Speculative Schedule

**Work Location**:
A place where a Member works during a Period, configured in the hidden Locations tab with a type: Office (e.g. OneRoof), Client site (JLL, C04, C60) or Remote. Inactive locations stay valid in history but are not offered for new entries. `Off` and leave types are not Work Locations.
_Avoid_: Place, Desk, Site, Spot

**Period**:
A half-day working segment: Morning (AM) or Afternoon (PM).
_Avoid_: Half-day, Slot, Shift, Session

**Baseline**:
A recurring schedule or date-range rule defining default AM/PM Work Locations for a Member.
_Avoid_: Default, Preset, Template

**Presence Baseline Sidebar**:
An interactive HTML sidebar panel opened inside Google Sheets allowing Members to easily configure their default recurring or date-range Baselines.
_Avoid_: Baseline Form, Schedule Popup

**Declare Contract**:
The Contracts sidebar (menu *Project Management → Declare Contract*): pick a Member, see their Tenures, add a new one (the current one ends the day before) or edit an old one. It opens by itself when a new name without a contract is typed in Members (ADR 0048).
_Avoid_: Members & Contracts sidebar, Onboarding form, HR panel

**Leave**:
An approved non-working period (Vacation, Sick Leave) defined across a date range.
_Avoid_: Time-off, PTO, Absence

### Work & Scheduling

**Project**:
A high-level initiative or work stream containing Tasks. Identified by its name, unique in the Projects tab: Tasks, Backups and Recurring Tasks pick it from a dropdown, and a renamed project is renamed everywhere (ADR 0053).
_Avoid_: Epic, Program, Initiative

**Project Lead**:
The primary Member accountable for a Project.
_Avoid_: Owner, Manager, Responsible

**Project Backup**:
The permanent secondary Member who steps in for the Project Lead during absences or leave. Typed in `tbl_projects.backup_member` and mirrored into `tbl_backups` as the row with no `end_date`; a change ends the old row, which keeps the history.
_Avoid_: Substitute, Second, Deputy, Shadow

**Project Backup Conflict**:
A working day (from today on) when nobody can cover an active Project: the Project Lead, every Designated Backup active that day and the Project Backup are all on approved leave or absent. It turns the project row red, emails the lead and a manager, and creates an "Organize interim backup" task.
_Avoid_: Dual absence, Uncovered project, Coverage clash

**Coverage Warning**:
Like a Project Backup Conflict, but at least one of the absences is a leave still `Requested`: the row turns orange and an email warns, with no task, so the gap can be fixed before approval.
_Avoid_: Soft conflict, Pending conflict

**Active Responsible**:
The person in charge of a Project today, shown read-only in `tbl_projects.active_responsible`: the Project Lead, else "(interim)" Designated Backup, else "(backup)" Project Backup, else "⚠ Nobody available".
_Avoid_: Owner today, Duty person

**Designated Backup**:
A dated (interim) replacement in `tbl_backups` covering a Project for a specific date range. It takes precedence over the Project Backup on those days, and several consecutive ones can together cover a long absence.
_Avoid_: Interim cover, Temp replacement

**Task**:
An actionable work item within a Project assigned to a Member, synchronized with Google Tasks.
_Avoid_: Todo, Ticket, Issue, Action Item

**Social Event**:
A team event or gathering with date, location, and RSVPs, synchronized with Google Calendar.
_Avoid_: Meeting, Appointment, Calendar Event

### Dashboards & Views

**Weekly Matrix View**:
A structured planning sheet where rows are weekdays and columns represent Members with paired AM/PM Work Location cells.
_Avoid_: Weekly table, Schedule grid

**Project Gantt**:
A timeline visualization plotting active Projects, start/end dates, and assigned Members.
_Avoid_: Roadmap, Timeline chart

### Architecture & Distribution

**Library**:
The centralized Google Apps Script script project published as a versioned library and consumed by multiple spreadsheets.
_Avoid_: Script, Package, Add-on, Template

**Consumer Sheet**:
A team's operational Google Sheet containing protected operational tables, dashboards, and a thin script wrapper pointing to the Library.
_Avoid_: Client sheet, Local spreadsheet, Workbook

**Sync Lock**:
The per-Consumer Sheet lock (`LockService` document lock) that serialises the nightly sync and the script-side handling of live edits, so they never write from stale snapshots at the same time.
_Avoid_: Mutex, Sync flag, Script lock

**Presence Fact Table**:
The hidden `fact_presence` sheet rebuilt by `refreshKpis`: one row per Member × date × Period with `worked_days`, `leave_days` and `capacity_days`, aggregated by Looker Studio. It never holds totals.
_Avoid_: KPI export, Report data

**FTE Rate**:
A Tenure's share of full time. `fte_rate` accepts a fraction (`0.8`) or a percentage (`80`); both mean 80 %.
_Avoid_: Allocation, Workload

**Table Registry**:
The list in `src/Tables.js` mapping each table's stable key (`members`, `tenures`…) to its readable tab name ("Members", "Contracts"…) and its former technical name. Tabs are found by their stored sheet ID, so renaming a tab is safe.
_Avoid_: Sheet map, Tab config

**Required Column**:
A table column that must contain a valid, non-blank value for any active or populated row (e.g. `member`, `role`, `start_date` in `tbl_tenures`).
_Avoid_: Mandatory field, Not null column

**Nullable Column**:
A table column that may legitimately remain empty without compromising data validity (e.g. `end_date` in `tbl_tenures` for permanent contracts).
_Avoid_: Optional field, Blank column

**Strictly Protected Column**:
A table column containing computed, foreign, or read-only data that regular users are completely prevented from editing (`setWarningOnly(false)`), with authorized admins receiving an audit warning toast upon modification.
_Avoid_: Hard locked column, Read-only field

**Warning-Protected Column**:
A table column containing sensitive business keys or attributes where edits are allowed only after the user acknowledges a native Google Sheets confirmation dialog (`setWarningOnly(true)`).
_Avoid_: Soft protected column, Alert column


**Undeclared Day**:
A working day of the current week for which an Active Member has no Weekly Presence entry, no Routine for that weekday and no leave. Listed by Check Workspace and sent in the Monday reminder.
_Avoid_: Missing presence, Unlogged day (that is "Not Logged", today only)

**Check Workspace**:
The menu report of everything that needs attention (contracts, backups, Project Backup Conflicts, Not Logged, Undeclared Days, duplicate Presence Log rows), each with a link to its fix (ADR 0041).
_Avoid_: Health check, Audit

**Birthday**:
A Member's birth date, in the last Members column (a date; whether the year is real is up to each person, ADR 0049). Only day and month drive the birthday reminder and the Team Birthdays calendar (ADR 0042).
_Avoid_: Date of birth

**Settings**:
The hidden Settings tab (menu *Edit Settings*) holding the organisation-specific values: default role and department, public holidays country (ADR 0049).
_Avoid_: Config, Parameters

**Public Holiday**:
A day in the hidden Holidays tab (French public holidays pre-filled, team days added by hand). Treated like a weekend by presence, Weekly Presence, KPIs, coverage, reminders and birthdays (ADR 0043).
_Avoid_: Holiday (a personal leave type), Bank holiday

**Deadline**:
The date by which a task must be done (Tasks column *Deadline*). It drives reminders, overdue and the audit, and is written in the Google Task notes (ADR 0050).
_Avoid_: Due date (Google Tasks uses "due" for the planned day)

**Origin**:
Where a task row comes from, in the hidden Tasks column *Origin*: a recurring rule and period, "manual", or a task trigger (ADR 0051).
_Avoid_: Recurring (former column name), Source

**Task Trigger**:
An event the app notices (first: *New member*) that creates the tasks configured for it in Task Templates (*Created when*, *Days after*, *Assign to*, *Instructions*), once per occurrence (ADR 0051).
_Avoid_: Automation, Hook

**Planned Date**:
The day someone plans to do a task (Tasks column *Planned date*), sent to Google Tasks as the task's date; the Deadline when empty (ADR 0050).
_Avoid_: Scheduled date, Due date

**Task Template** (formerly Task Type):
A preconfigured kind of task (Quiz, Onboarding, Time & expenses…): its schedule (frequency: weekly, twice a month, monthly or one-off, and a due rule), a Mandatory flag for the tasks the manager checks, who gets it by default, its instructions, its reminder days and the event that creates it, if any. Listed in the hidden Task Templates tab; the Assign Tasks sidebar starts from it (ADR 0045, 0052).
_Avoid_: Task Type (former name)

**Dropdown Model**:
A cell in the hidden Dropdowns tab whose dropdown (chip style, colour per option) is set by hand and copied by the app onto every cell using that list, today Weekly Presence. Table dropdowns are coloured in the table itself (ADR 0054).

**Manager Digest**:
The email sent, the working day after a mandatory task is due, to every member with the Manager role (Settings), listing who has not done it (ADR 0052).

**Remind days**:
The working days before a task's due date when its person is emailed, set per Task Template ("3; 1"; blank = 3) (ADR 0052).
_Avoid_: Category, Task kind

**Recurring Task**:
A rule that creates dated tasks of one Task Type every period, either for each person it applies to or for one person in rotation (ADR 0045).
_Avoid_: Routine (that is presence), Template

**Rotation**:
A Recurring Task given to one person per period, the one whose last turn is the oldest, skipping people on approved leave.
_Avoid_: Round robin, Duty roster
