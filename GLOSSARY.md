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
The contractual employment state of a Member (`Active`, `Upcoming`, `Departed`), evaluated from their Tenure dates.
_Avoid_: Employment Status, Employee State

**Today Presence**:
The real-time operational availability and Work Location of a Member for the current day, evaluated from Leaves, Presence Logs, and Events.
_Avoid_: Daily Status, Attendance State

**Presence Log**:
The ground-truth database (`tbl_presence_log`) capturing a Member's actual recorded presence and Work Location for morning (AM) and afternoon (PM) periods up to the current date. Future planning is dynamically projected on presentation views, keeping the presence log strictly historical and current.
_Avoid_: Attendance, Timesheet, Check-in, Speculative Schedule

**Work Location**:
A recognized physical office, remote site, or absence status where a Member operates during a Period.
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

**Leave**:
An approved non-working period (Vacation, Sick Leave) defined across a date range.
_Avoid_: Time-off, PTO, Absence

### Work & Scheduling

**Project**:
A high-level initiative or work stream containing Tasks.
_Avoid_: Epic, Program, Initiative

**Project Lead**:
The primary Member accountable for a Project.
_Avoid_: Owner, Manager, Responsible

**Project Backup**:
The designated secondary Member who steps in for the Project Lead during absences or leave.
_Avoid_: Substitute, Second, Deputy, Shadow

**Project Backup Conflict**:
An operational risk scenario where both the Project Lead and Project Backup on an active Project are absent simultaneously (overlapping approved leaves or concurrent absences), requiring immediate interim backup designation in `tbl_backups`.
_Avoid_: Dual absence, Uncovered project, Coverage clash

**Designated Backup**:
An active, time-bounded replacement assignment recorded in `tbl_backups` defining who backs up who on which project for a specific date range, mitigating project backup conflicts.
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

