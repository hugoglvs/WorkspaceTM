/**
 * Project Management Workspace - Schema Definition
 * Version: 1
 */

var CURRENT_SCHEMA_VERSION = 1;

var FONT_FAMILY = "Google Sans";
var FONT_SIZE = 12;
var DATE_FORMAT = "yyyy-mm-dd";

var SCHEMA_SPEC = {
  tbl_members: {
    sheetName: "tbl_members",
    columns: ["member", "accenture_email", "phone_number", "current_role", "current_department", "tenure_status", "today_presence"],
    requiredColumns: ["member", "current_role", "current_department", "tenure_status", "today_presence"],
    nullableColumns: ["accenture_email", "phone_number"],
    strictlyProtectedColumns: ["member", "current_role", "current_department", "tenure_status", "today_presence"],
    warningProtectedColumns: [],
    protectedColumns: ["member", "current_role", "current_department", "tenure_status", "today_presence"],
    columnTypes: {
      member: "PEOPLE_CHIP",
      accenture_email: "TEXT",
      phone_number: "TEXT",
      current_role: "TEXT",
      current_department: "TEXT",
      tenure_status: "TEXT",
      today_presence: "TEXT"
    },
    validations: {}
  },
  tbl_tenures: {
    sheetName: "tbl_tenures",
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
  tbl_presence_ranges: {
    sheetName: "tbl_presence_ranges",
    hiddenByDefault: true,
    columns: ["range_id", "member", "range_type", "start_date", "start_period", "end_date", "end_period", "day_of_week", "location_am", "location_pm", "status"],
    requiredColumns: ["range_id", "member", "range_type", "start_date", "day_of_week", "location_am", "location_pm", "status"],
    nullableColumns: ["start_period", "end_date", "end_period"],
    strictlyProtectedColumns: ["range_id"],
    warningProtectedColumns: ["member"],
    protectedColumns: ["range_id", "member"],
    columnTypes: {
      range_id: "TEXT",
      member: "PEOPLE_CHIP",
      range_type: "DROPDOWN",
      start_date: "DATE",
      start_period: "DROPDOWN",
      end_date: "DATE",
      end_period: "DROPDOWN",
      day_of_week: "DROPDOWN",
      location_am: "DROPDOWN",
      location_pm: "DROPDOWN",
      status: "DROPDOWN"
    },
    validations: {
      range_type: ["Baseline", "Leave"],
      start_period: ["AM", "PM"],
      end_period: ["AM", "PM"],
      day_of_week: ["All", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
      location_am: ["Remote", "Main Office", "Downtown Office", "Branch Office", "Client Site", "Off", "Vacation", "Sick Leave", "Maternity/Paternity", "Unpaid Leave"],
      location_pm: ["Remote", "Main Office", "Downtown Office", "Branch Office", "Client Site", "Off", "Vacation", "Sick Leave", "Maternity/Paternity", "Unpaid Leave"],
      status: ["Requested", "Approved", "Cancelled"]
    }
  },
  tbl_presence_log: {
    sheetName: "tbl_presence_log",
    hiddenByDefault: true,
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
      location_am: ["Remote", "Main Office", "Downtown Office", "Branch Office", "Client Site", "Off", "Vacation", "Sick Leave", "Holiday"],
      location_pm: ["Remote", "Main Office", "Downtown Office", "Branch Office", "Client Site", "Off", "Vacation", "Sick Leave", "Holiday"]
    }
  },
  tbl_projects: {
    sheetName: "tbl_projects",
    columns: ["project_id", "name", "status", "start_date", "target_date", "lead_member", "backup_member"],
    requiredColumns: ["project_id", "name", "status", "start_date", "lead_member"],
    nullableColumns: ["target_date", "backup_member"],
    strictlyProtectedColumns: ["project_id"],
    warningProtectedColumns: ["lead_member", "start_date"],
    protectedColumns: ["project_id", "lead_member", "start_date"],
    columnTypes: {
      project_id: "TEXT",
      name: "TEXT",
      status: "DROPDOWN",
      start_date: "DATE",
      target_date: "DATE",
      lead_member: "PEOPLE_CHIP",
      backup_member: "PEOPLE_CHIP"
    },
    validations: {
      status: ["Planning", "Active", "On Hold", "Completed"]
    }
  },
  tbl_tasks: {
    sheetName: "tbl_tasks",
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
  tbl_events: {
    sheetName: "tbl_events",
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
  tbl_backups: {
    sheetName: "tbl_backups",
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
    SCHEMA_SPEC: SCHEMA_SPEC
  };
}
