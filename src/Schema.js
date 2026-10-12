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
      birthday: "TEXT" // day and month only, e.g. "14 Mar" (ADR 0042)
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
      day_of_week: ["All", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
      location_am: ["OneRoof", "Remote", "JLL", "C04", "C60", "Off"],
      location_pm: ["OneRoof", "Remote", "JLL", "C04", "C60", "Off"]
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
    validations: {
      location_am: ["OneRoof", "Remote", "JLL", "C04", "C60", "Off", "Vacation", "Sick Leave", "Maternity/Paternity", "Unpaid Leave", "Holiday"],
      location_pm: ["OneRoof", "Remote", "JLL", "C04", "C60", "Off", "Vacation", "Sick Leave", "Maternity/Paternity", "Unpaid Leave", "Holiday"]
    },
    // Filled at runtime from the Locations tab (src/Locations.js)
    locationColumns: { location_am: "withLeaves", location_pm: "withLeaves" }
  },
  projects: {
    sheetName: "Projects",
    legacyName: "tbl_projects",
    headers: {
      project_id: "ID",
      name: "Project",
      status: "Status",
      start_date: "Start",
      target_date: "Target date",
      lead_member: "Lead",
      backup_member: "Backup",
      active_responsible: "In charge today"
    },
    columns: ["project_id", "name", "status", "start_date", "target_date", "lead_member", "backup_member", "active_responsible"],
    requiredColumns: ["project_id", "name", "status", "start_date", "lead_member"],
    nullableColumns: ["target_date", "backup_member", "active_responsible"],
    strictlyProtectedColumns: ["project_id", "active_responsible"],
    warningProtectedColumns: ["lead_member", "start_date"],
    protectedColumns: ["project_id", "lead_member", "start_date", "active_responsible"],
    columnTypes: {
      project_id: "TEXT",
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
      task_id: "ID",
      project_id: "Project ID",
      assigned_member: "Assigned to",
      title: "Title",
      description: "Description",
      due_date: "Due date",
      status: "Status",
      google_task_id: "Google Task ID"
    },
    columns: ["task_id", "project_id", "assigned_member", "title", "description", "due_date", "status", "google_task_id"],
    requiredColumns: ["task_id", "title", "status"],
    nullableColumns: ["project_id", "assigned_member", "description", "due_date", "google_task_id"],
    strictlyProtectedColumns: ["task_id", "google_task_id"],
    warningProtectedColumns: [],
    protectedColumns: ["task_id", "google_task_id"],
    columnTypes: {
      task_id: "TEXT",
      project_id: "TEXT",
      assigned_member: "PEOPLE_CHIP",
      title: "TEXT",
      description: "TEXT",
      due_date: "DATE",
      status: "DROPDOWN",
      google_task_id: "TEXT"
    },
    validations: {
      status: ["Todo", "In Progress", "Done", "Blocked"]
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
    columns: ["backup_id", "project_id", "primary_member", "backup_member", "start_date", "end_date", "status", "notes"],
    requiredColumns: ["backup_id", "project_id", "primary_member", "backup_member", "start_date", "status"],
    nullableColumns: ["end_date", "notes"],
    strictlyProtectedColumns: ["backup_id"],
    warningProtectedColumns: ["project_id", "primary_member"],
    protectedColumns: ["backup_id", "project_id", "primary_member"],
    columnTypes: {
      backup_id: "TEXT",
      project_id: "TEXT",
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
