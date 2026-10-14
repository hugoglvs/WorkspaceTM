/**
 * Project Management Workspace - Schema Definition
 * Version: 1
 */

var CURRENT_SCHEMA_VERSION = 1;

var FONT_FAMILY = "Google Sans";
var FONT_SIZE = 12;
var DATE_FORMAT = "yyyy-mm-dd";

/**
 * Table schemas keyed by the registry key of src/Tables.js (ADR 0039).
 * sheetName = readable tab and native table name; legacyName = former technical name.
 * headers = readable column labels for the visible tables (column keys stay internal).
 */
var SCHEMA_SPEC = {
  locations: {
    sheetName: "Locations",
    legacyName: "tbl_locations",
    headers: { location: "Location", type: "Type", active: "Active" },
    columns: ["location", "type", "active"],
    requiredColumns: ["location", "type"],
    nullableColumns: ["active"],
    strictlyProtectedColumns: [],
    warningProtectedColumns: [],
    protectedColumns: [],
    columnTypes: { location: "TEXT", type: "DROPDOWN", active: "BOOLEAN" },
    validations: { type: ["Office", "Client site", "Remote"] }
  },
  members: {
    sheetName: "Members",
    legacyName: "tbl_members",
    headers: {
      member: "Member",
      accenture_email: "Accenture email",
      phone_number: "Phone",
      current_role: "Role",
      current_department: "Department",
      tenure_status: "Status",
      today_presence: "Today",
      birthday: "Birthday"
    },
    columns: ["member", "accenture_email", "phone_number", "current_role", "current_department", "tenure_status", "today_presence", "birthday"],
    requiredColumns: ["member", "current_role", "current_department", "tenure_status", "today_presence"],
    nullableColumns: ["accenture_email", "phone_number", "birthday"],
    strictlyProtectedColumns: ["member", "current_role", "current_department", "tenure_status", "today_presence"],
    warningProtectedColumns: [],
    protectedColumns: ["member", "current_role", "current_department", "tenure_status", "today_presence"],
    columnTypes: {
      member: "PEOPLE_CHIP",
      accenture_email: "TEXT",
      phone_number: "TEXT",
      current_role: "TEXT",
      current_department: "TEXT",
      tenure_status: "DROPDOWN",
      today_presence: "TEXT",
      birthday: "DATE" // full birth date; only day and month are used (ADR 0049)
    },
    validations: {
      // The four values syncMembersAndTenures writes (shown as coloured chips)
      tenure_status: ["Active", "Upcoming", "Departed", "No Contract"]
    }
  },
  tenures: {
    sheetName: "Contracts",
    legacyName: "tbl_tenures",
    columns: ["member", "role", "department", "start_date", "end_date", "contract_type", "fte_rate"],
    requiredColumns: ["member", "role", "department", "start_date", "contract_type", "fte_rate"],
    nullableColumns: ["end_date"],
    strictlyProtectedColumns: [],
    warningProtectedColumns: ["member", "start_date"],
    protectedColumns: ["member", "start_date"],
    columnTypes: {
      member: "PEOPLE_CHIP",
      role: "TEXT",
      department: "TEXT",
      start_date: "DATE",
      end_date: "DATE",
      contract_type: "DROPDOWN",
      fte_rate: "DOUBLE"
    },
    validations: {
      contract_type: ["Permanent", "Contractor", "Apprentice", "Intern"]
    }
  },
  baselines: {
    sheetName: "Routines",
    legacyName: "tbl_presence_baselines",
    columns: ["baseline_id", "member", "start_date", "end_date", "day_of_week", "location_am", "location_pm"],
    requiredColumns: ["baseline_id", "member", "start_date", "day_of_week", "location_am", "location_pm"],
    nullableColumns: ["end_date"],
    strictlyProtectedColumns: ["baseline_id"],
    warningProtectedColumns: ["member"],
    protectedColumns: ["baseline_id", "member"],
    columnTypes: {
      baseline_id: "TEXT",
      member: "PEOPLE_CHIP",
      start_date: "DATE",
      end_date: "DATE",
      day_of_week: "DROPDOWN",
      location_am: "DROPDOWN",
      location_pm: "DROPDOWN"
    },
    validations: {
      day_of_week: ["All", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
      // location_am / location_pm: the Locations tab plus "Off" (locationColumns)
    },
    // Filled at runtime from the Locations tab (src/Locations.js)
    locationColumns: { location_am: "plain", location_pm: "plain" }
  },
  leaves: {
    sheetName: "Leaves",
    legacyName: "tbl_leaves",
    columns: ["leave_id", "member", "start_date", "start_period", "end_date", "end_period", "leave_type", "status"],
    requiredColumns: ["leave_id", "member", "start_date", "end_date", "leave_type", "status"],
    nullableColumns: ["start_period", "end_period"],
    strictlyProtectedColumns: ["leave_id"],
    warningProtectedColumns: ["member", "leave_type"],
    protectedColumns: ["leave_id", "member", "leave_type"],
    columnTypes: {
      leave_id: "TEXT",
      member: "PEOPLE_CHIP",
      start_date: "DATE",
      start_period: "DROPDOWN",
      end_date: "DATE",
      end_period: "DROPDOWN",
      leave_type: "DROPDOWN",
      status: "DROPDOWN"
    },
    validations: {
      start_period: ["AM", "PM"],
      end_period: ["AM", "PM"],
      leave_type: ["Vacation", "Sick Leave", "Maternity/Paternity", "Unpaid Leave"],
      status: ["Requested", "Approved", "Cancelled"]
    }
  },
  presenceLog: {
    sheetName: "Presence Log",
    legacyName: "tbl_presence_log",
    columns: ["log_id", "date", "member", "location_am", "location_pm", "timestamp"],
    requiredColumns: ["log_id", "date", "member", "location_am", "location_pm", "timestamp"],
    nullableColumns: [],
    strictlyProtectedColumns: ["log_id", "timestamp"],
    warningProtectedColumns: ["date", "member"],
    protectedColumns: ["log_id", "timestamp", "date", "member"],
    columnTypes: {
      log_id: "TEXT",
      date: "DATE",
      member: "PEOPLE_CHIP",
      location_am: "DROPDOWN",
      location_pm: "DROPDOWN",
      timestamp: "TEXT"
    },
    // location_am / location_pm: the Locations tab, "Off" and leave types (locationColumns)
    validations: {},
    // Filled at runtime from the Locations tab (src/Locations.js)
    locationColumns: { location_am: "withLeaves", location_pm: "withLeaves" }
  },
  projects: {
    sheetName: "Projects",
    legacyName: "tbl_projects",
    // The project name is the reference used by Tasks, Backups and Recurring Tasks
    // (ADR 0053): no ID; a renamed project is renamed everywhere.
    headers: {
      name: "Project",
      status: "Status",
      start_date: "Start",
      target_date: "Target date",
      lead_member: "Lead",
      backup_member: "Backup",
      active_responsible: "In charge today"
    },
    columns: ["name", "status", "start_date", "target_date", "lead_member", "backup_member", "active_responsible"],
    removedHeaders: ["project_id", "ID"],
    uniqueColumns: ["name"],
    requiredColumns: ["name", "status", "start_date", "lead_member"],
    nullableColumns: ["target_date", "backup_member", "active_responsible"],
    strictlyProtectedColumns: ["active_responsible"],
    warningProtectedColumns: ["lead_member", "start_date"],
    protectedColumns: ["lead_member", "start_date", "active_responsible"],
    columnTypes: {
      name: "TEXT",
      status: "DROPDOWN",
      start_date: "DATE",
      target_date: "DATE",
      lead_member: "PEOPLE_CHIP",
      backup_member: "PEOPLE_CHIP",
      active_responsible: "TEXT"
    },
    validations: {
      status: ["Planning", "Active", "On Hold", "Completed"]
    }
  },
  tasks: {
    sheetName: "Tasks",
    legacyName: "tbl_tasks",
    headers: {
      task_type: "Type",
      title: "Title",
      assigned_member: "Assigned to",
      project: "Project",
      planned_date: "Planned date",
      due_date: "Deadline",
      status: "Status",
      description: "Description",
      done_on: "Done on",
      task_id: "ID",
      google_task_id: "Google Task ID",
      origin: "Origin"
    },
    // One row = one task for one person with one deadline (ADR 0051).
    // Deadline (key due_date) drives reminders, overdue and the Calendar; Planned date is the
    // day shown in Google Tasks, the deadline when empty (ADR 0050). Origin: "<rule>|<period>"
    // (recurring, "|manual" once changed by hand), "manual", or "trigger:<event>" (ADR 0051).
    columns: ["task_type", "title", "assigned_member", "project", "planned_date", "due_date", "status", "description", "done_on", "task_id", "google_task_id", "origin"],
    hiddenColumns: ["task_id", "google_task_id", "origin"],
    // Labels of earlier layouts, so Initialize / Upgrade recognises a moved table (ADR 0051)
    formerHeaders: { due_date: ["Due date"], origin: ["Recurring", "recurrence"], project: ["Project ID", "project_id"] },
    removedHeaders: ["Done by", "done_by"],
    requiredColumns: ["task_id", "title", "status"],
    nullableColumns: ["task_type", "project", "assigned_member", "planned_date", "due_date", "description", "done_on", "google_task_id", "origin"],
    strictlyProtectedColumns: ["task_id", "google_task_id", "origin"],
    warningProtectedColumns: [],
    protectedColumns: ["task_id", "google_task_id", "origin"],
    taskTypeColumns: { task_type: true },
    // Dropdown of the Projects tab's names (ADR 0053)
    projectColumns: { project: true },
    columnTypes: {
      task_type: "DROPDOWN",
      title: "TEXT",
      assigned_member: "PEOPLE_CHIP",
      project: "DROPDOWN",
      planned_date: "DATE",
      due_date: "DATE",
      status: "DROPDOWN",
      description: "TEXT",
      done_on: "DATE",
      task_id: "TEXT",
      google_task_id: "TEXT",
      origin: "TEXT"
    },
    validations: {
      status: ["Todo", "In Progress", "Done", "Blocked", "Cancelled"]
    }
  },
  events: {
    sheetName: "Events",
    legacyName: "tbl_events",
    headers: {
      event_id: "ID",
      title: "Title",
      date: "Date",
      duration: "Duration",
      location: "Location",
      organizer: "Organizer",
      google_calendar_event_id: "Google Calendar ID"
    },
    columns: ["event_id", "title", "date", "duration", "location", "organizer", "google_calendar_event_id"],
    requiredColumns: ["event_id", "title", "date"],
    nullableColumns: ["duration", "location", "organizer", "google_calendar_event_id"],
    strictlyProtectedColumns: ["event_id", "google_calendar_event_id"],
    warningProtectedColumns: ["date"],
    protectedColumns: ["event_id", "google_calendar_event_id", "date"],
    columnTypes: {
      event_id: "TEXT",
      title: "TEXT",
      date: "DATE",
      duration: "DOUBLE",
      location: "TEXT",
      organizer: "PEOPLE_CHIP",
      google_calendar_event_id: "TEXT"
    },
    validations: {}
  },
  backups: {
    sheetName: "Backups",
    legacyName: "tbl_backups",
    columns: ["backup_id", "project", "primary_member", "backup_member", "start_date", "end_date", "status", "notes"],
    formerHeaders: { project: ["Project ID", "project_id"] },
    requiredColumns: ["backup_id", "project", "primary_member", "backup_member", "start_date", "status"],
    nullableColumns: ["end_date", "notes"],
    strictlyProtectedColumns: ["backup_id"],
    warningProtectedColumns: ["project", "primary_member"],
    protectedColumns: ["backup_id", "project", "primary_member"],
    projectColumns: { project: true },
    columnTypes: {
      backup_id: "TEXT",
      project: "DROPDOWN",
      primary_member: "PEOPLE_CHIP",
      backup_member: "PEOPLE_CHIP",
      start_date: "DATE",
      end_date: "DATE",
      status: "DROPDOWN",
      notes: "TEXT"
    },
    validations: {
      status: ["Active", "Planned", "Completed", "Cancelled"]
    }
  },
  holidays: {
    sheetName: "Holidays",
    legacyName: "tbl_holidays",
    headers: { date: "Date", name: "Holiday" },
    columns: ["date", "name"],
    requiredColumns: ["date"],
    nullableColumns: ["name"],
    strictlyProtectedColumns: [],
    warningProtectedColumns: [],
    protectedColumns: [],
    columnTypes: { date: "DATE", name: "TEXT" },
    validations: {}
  },
  taskTypes: {
    sheetName: "Task Templates",
    legacyName: "tbl_task_types",
    headers: { type: "Type", frequency: "Frequency", due: "Due", mandatory: "Mandatory", active: "Active",
      created_when: "Created when", days_after: "Days after", assign_to: "Assign to", instructions: "Instructions", remind: "Remind" },
    // Created when / Days after: tasks created by an event (ADR 0051); Remind: working days before (ADR 0052)
    columns: ["type", "frequency", "due", "mandatory", "active", "created_when", "days_after", "assign_to", "instructions", "remind"],
    requiredColumns: ["type", "frequency"],
    nullableColumns: ["due", "mandatory", "active", "created_when", "days_after", "assign_to", "instructions", "remind"],
    strictlyProtectedColumns: [],
    warningProtectedColumns: [],
    protectedColumns: [],
    columnTypes: { type: "TEXT", frequency: "DROPDOWN", due: "TEXT", mandatory: "BOOLEAN", active: "BOOLEAN",
      created_when: "DROPDOWN", days_after: "DOUBLE", assign_to: "TEXT", instructions: "TEXT", remind: "TEXT" },
    // Filled at runtime from TASK_TRIGGERS (src/TaskTriggers.js)
    triggerColumns: { created_when: true },
    validations: { frequency: ["Weekly", "Twice a month", "Monthly", "One-off"] }
  },
  recurringTasks: {
    sheetName: "Recurring Tasks",
    legacyName: "tbl_recurring_tasks",
    headers: { rule_id: "ID", task_type: "Type", project: "Project", mode: "Who", applies_to: "Applies to", start_date: "Start", end_date: "End", active: "Active" },
    columns: ["rule_id", "task_type", "project", "mode", "applies_to", "start_date", "end_date", "active"],
    formerHeaders: { project: ["Project ID", "project_id"] },
    projectColumns: { project: true },
    requiredColumns: ["rule_id", "task_type", "mode"],
    nullableColumns: ["project", "applies_to", "start_date", "end_date", "active"],
    strictlyProtectedColumns: [],
    warningProtectedColumns: [],
    protectedColumns: [],
    columnTypes: { rule_id: "TEXT", task_type: "DROPDOWN", project: "DROPDOWN", mode: "DROPDOWN", applies_to: "TEXT", start_date: "DATE", end_date: "DATE", active: "BOOLEAN" },
    taskTypeColumns: { task_type: true },
    validations: { mode: ["Rotation", "Each person"] }
  },
  settings: {
    sheetName: "Settings",
    legacyName: "tbl_settings",
    headers: { setting: "Setting", value: "Value", description: "Description" },
    columns: ["setting", "value", "description"],
    requiredColumns: ["setting"],
    nullableColumns: ["value", "description"],
    strictlyProtectedColumns: [],
    warningProtectedColumns: ["setting"],
    protectedColumns: ["setting"],
    columnTypes: { setting: "TEXT", value: "TEXT", description: "TEXT" },
    validations: {}
  }
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    CURRENT_SCHEMA_VERSION: CURRENT_SCHEMA_VERSION,
    FONT_FAMILY: FONT_FAMILY,
    FONT_SIZE: FONT_SIZE,
    DATE_FORMAT: DATE_FORMAT,
    SCHEMA_SPEC: SCHEMA_SPEC,
    schemaHeaderRow: schemaHeaderRow,
    schemaColumnLabel: schemaColumnLabel
  };
}

/**
 * Header row for a table: readable labels when defined, otherwise the column keys.
 *
 * @param {Object} config A SCHEMA_SPEC entry
 * @return {string[]}
 */
function schemaHeaderRow(config) {
  return config.columns.map(function(col) {
    return (config.headers && config.headers[col]) ? config.headers[col] : col;
  });
}

/** Readable label of a column (falls back to the key). */
function schemaColumnLabel(config, col) {
  return (config && config.headers && config.headers[col]) ? config.headers[col] : col;
}
