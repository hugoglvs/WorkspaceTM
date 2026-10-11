/**
 * Project Management Workspace - Google Calendar Event Synchronization
 * Ticket 08: Google Calendar & Event Invitations Synchronization
 *
 * Integrates tbl_events with Google Calendar:
 * - Scans tbl_events for events missing google_calendar_event_id.
 * - Creates calendar events on the configured team calendar using CalendarApp.
 * - Invites active team members from tbl_members (using their Workspace emails).
 * - Writes back the generated event ID into tbl_events.google_calendar_event_id.
 * - Propagates title, date, and location edits on tbl_events to Google Calendar.
 * - Stores configured Google Calendar ID in Document Properties.
 */

var GOOGLE_CALENDAR_ID_KEY = "GOOGLE_CALENDAR_ID";

/**
 * Retrieves the configured Google Calendar ID from Document Properties,
 * defaulting to "primary" if not specified.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {string} Calendar ID
 */
function getCalendarId(targetSpreadsheet) {
  var props = PropertiesService.getDocumentProperties();
  var savedId = props.getProperty(GOOGLE_CALENDAR_ID_KEY);
  if (savedId && savedId.trim()) {
    return savedId.trim();
  }
  return "primary";
}

/**
 * Stores the target Google Calendar ID in Document Properties.
 *
 * @param {string} calendarId
 */
function setCalendarId(calendarId) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty(GOOGLE_CALENDAR_ID_KEY, String(calendarId || "primary").trim());
}

/**
 * Resolves column positions for tbl_events dynamically from header row.
 * Resilient against spaces, casing, underscores, and variable column orders.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} eventsSheet
 * @return {Object} Column map
 */
function getEventColumnMap(eventsSheet) {
  var defaultMap = {
    eventId: 1,
    title: 2,
    date: 3,
    duration: 4,
    location: 5,
    organizer: 6,
    googleCalendarEventId: 7
  };
  if (!eventsSheet) return defaultMap;
  var lastCol = Math.max(typeof eventsSheet.getLastColumn === "function" ? (eventsSheet.getLastColumn() || 0) : 0, 7);
  if (lastCol < 1) return defaultMap;

  var headerVals = eventsSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var found = {};
  var hasAnyHeader = false;

  for (var c = 0; c < headerVals.length; c++) {
    var raw = String(headerVals[c] || "").trim();
    if (!raw) continue;
    hasAnyHeader = true;
    var norm = raw.toLowerCase().replace(/[^a-z0-9]/g, "");
    if (norm === "eventid" || norm === "id") found.eventId = c + 1;
    else if (norm === "title" || norm === "eventtitle" || norm === "name" || norm === "eventname") found.title = c + 1;
    else if (norm === "date" || norm === "datetime" || norm === "eventdate" || norm === "startdate" || norm === "startdatetime") found.date = c + 1;
    else if (norm === "duration" || norm === "eventduration" || norm === "length" || norm === "minutes") found.duration = c + 1;
    else if (norm === "location" || norm === "place" || norm === "venue") found.location = c + 1;
    else if (norm === "organizer" || norm === "eventorganizer" || norm === "lead" || norm === "owner" || norm === "host") found.organizer = c + 1;
    else if (norm === "googlecalendareventid" || norm === "calendareventid" || norm === "geventid" || norm === "googleeventid") found.googleCalendarEventId = c + 1;
  }

  if (!hasAnyHeader) return defaultMap;

  return {
    eventId: found.eventId || 1,
    title: found.title || 2,
    date: found.date || 3,
    duration: found.duration || -1,
    location: found.location || (found.duration ? 5 : 4),
    organizer: found.organizer || -1,
    googleCalendarEventId: found.googleCalendarEventId || (found.organizer ? 7 : (found.duration ? 6 : 5))
  };
}

/**
 * Generates an event ID in standard EVT-YYMMDD-XXXX format.
 *
 * @return {string} Generated event ID
 */
function generateEventId() {
  var prefix = "EVT-";
  try {
    if (typeof Utilities !== "undefined" && Utilities.formatDate && typeof Session !== "undefined" && Session.getScriptTimeZone) {
      return prefix + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "yyMMdd") + "-" + Math.floor(1000 + Math.random() * 9000);
    }
  } catch (eFmt) {}
  var d = new Date();
  var yy = String(d.getFullYear()).slice(-2);
  var mNum = d.getMonth() + 1;
  var mm = (mNum < 10 ? "0" : "") + mNum;
  var dNum = d.getDate();
  var dd = (dNum < 10 ? "0" : "") + dNum;
  var rand = Math.floor(1000 + Math.random() * 9000);
  return prefix + yy + mm + dd + "-" + rand;
}

/**
 * Parses duration column value into minutes.
 * If null, empty, or undefined, returns null (indicating a whole-day event).
 * Supports numbers (minutes) and strings ("60", "60m", "1h", "1.5 hours", etc.).
 *
 * @param {number|string|null} rawDuration
 * @return {number|null} Duration in minutes, or null if whole day event
 */
function parseEventDuration(rawDuration) {
  if (rawDuration === null || rawDuration === undefined) return null;
  var str = String(rawDuration).trim();
  if (!str) return null;

  // Hours e.g. "1.5h", "2 hours", "1 hr"
  var hMatch = str.match(/^(\d+(?:\.\d+)?)\s*(?:h|hr|hrs|hour|hours)$/i);
  if (hMatch) {
    var hrs = parseFloat(hMatch[1]);
    return (!isNaN(hrs) && hrs > 0) ? Math.round(hrs * 60) : null;
  }

  // Minutes e.g. "60m", "45 mins", "90 min"
  var mMatch = str.match(/^(\d+(?:\.\d+)?)\s*(?:m|min|mins|minute|minutes)$/i);
  if (mMatch) {
    var mins = parseFloat(mMatch[1]);
    return (!isNaN(mins) && mins > 0) ? Math.round(mins) : null;
  }

  // Pure number
  var num = parseFloat(str);
  if (!isNaN(num) && num > 0) {
    return Math.round(num);
  }
  return null;
}

/**
 * Parses a Date object or date/datetime string into a Date object.
 * Preserves hours and minutes for datetime strings.
 *
 * @param {Date|string} rawDate
 * @return {Date|null}
 */
function parseEventDate(rawDate) {
  if (!rawDate) return null;
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
    return rawDate;
  }
  var str = String(rawDate).trim();
  // Match YYYY-MM-DD or YYYY-MM-DD HH:mm(:ss) or YYYY-MM-DDTHH:mm(:ss)
  var match = str.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (match) {
    var y = parseInt(match[1], 10);
    var m = parseInt(match[2], 10) - 1;
    var d = parseInt(match[3], 10);
    var hr = match[4] !== undefined ? parseInt(match[4], 10) : 0;
    var min = match[5] !== undefined ? parseInt(match[5], 10) : 0;
    var sec = match[6] !== undefined ? parseInt(match[6], 10) : 0;
    return new Date(y, m, d, hr, min, sec, 0);
  }
  var parsed = new Date(str);
  if (!isNaN(parsed.getTime())) return parsed;
  return null;
}

/**
 * Formats a Date object as YYYY-MM-DD string.
 *
 * @param {Date|string} d
 * @return {string}
 */
function formatEventDateIso(d) {
  if (!d) return "";
  var dateObj = parseEventDate(d);
  if (!dateObj) return "";
  var y = dateObj.getFullYear();
  var m = dateObj.getMonth() + 1;
  var day = dateObj.getDate();
  return y + "-" + (m < 10 ? "0" : "") + m + "-" + (day < 10 ? "0" : "") + day;
}

/**
 * Resolves the email address of the organizer from raw / rich cell values or tbl_members.
 * Defaults to Session.getActiveUser().getEmail() if available.
 *
 * @param {*} rawOrganizer
 * @param {*} richOrganizer
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {string}
 */
function getOrganizerEmail(rawOrganizer, richOrganizer, ss) {
  if (rawOrganizer) {
    if (typeof rawOrganizer === "object") {
      if (rawOrganizer.personProperties && rawOrganizer.personProperties.email) {
        return String(rawOrganizer.personProperties.email).trim().toLowerCase();
      }
      if (rawOrganizer.email) {
        return String(rawOrganizer.email).trim().toLowerCase();
      }
    }
    var str = String(rawOrganizer).trim();
    var matchMail = str.match(/([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/);
    if (matchMail) {
      return matchMail[1].toLowerCase();
    }
    if (richOrganizer && typeof richOrganizer.getLinkUrl === "function") {
      var link = richOrganizer.getLinkUrl();
      if (link) {
        var mLink = link.match(/mailto:([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})/i);
        if (mLink) return mLink[1].toLowerCase();
      }
    }
    if (typeof extractEmailFromChip === "function") {
      var extracted = extractEmailFromChip(rawOrganizer, richOrganizer);
      if (extracted) return extracted.toLowerCase();
    }
    if (ss) {
      var membersSheet = ss.getSheetByName("tbl_members");
      if (membersSheet && membersSheet.getLastRow() > 1) {
        var mRows = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, Math.max(membersSheet.getLastColumn(), 2)).getValues();
        for (var m = 0; m < mRows.length; m++) {
          var mName = String(mRows[m][0] || "").trim().toLowerCase();
          if (mName && mName === str.toLowerCase()) {
            var mEmail = String(mRows[m][1] || "").trim().toLowerCase();
            if (mEmail) return mEmail;
          }
        }
      }
    }
  }
  try {
    if (typeof Session !== "undefined" && Session.getActiveUser && Session.getActiveUser().getEmail) {
      var activeEmail = Session.getActiveUser().getEmail();
      if (activeEmail) return String(activeEmail).trim().toLowerCase();
    }
  } catch (eSes) {}
  return "";
}

/**
 * Retrieves the organizer's Google Calendar instance.
 * Falls back to default calendar if organizer calendar is not directly accessible.
 *
 * @param {string} organizerEmail
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
 * @return {GoogleAppsScript.Calendar.Calendar|null}
 */
function getOrganizerCalendar(organizerEmail, ss) {
  if (typeof CalendarApp === "undefined") return null;
  var activeUserEmail = "";
  try {
    if (typeof Session !== "undefined" && Session.getActiveUser && Session.getActiveUser().getEmail) {
      activeUserEmail = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
    }
  } catch (eSes) {}

  if (organizerEmail && (!activeUserEmail || organizerEmail.toLowerCase() !== activeUserEmail)) {
    try {
      var orgCal = CalendarApp.getCalendarById(organizerEmail);
      if (orgCal) return orgCal;
    } catch (e) {
      Logger.log("getCalendarById error for organizer " + organizerEmail + ": " + e);
    }
  }
  try {
    return CalendarApp.getDefaultCalendar();
  } catch (eDef) {
    Logger.log("getDefaultCalendar error: " + eDef);
    return null;
  }
}

/**
 * Gathers emails of active team members from tbl_members to invite to team events.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {string[]} Array of email addresses
 */
function getActiveTeamMemberEmails(ss) {
  var emails = [];
  if (!ss) return emails;
  var membersSheet = ss.getSheetByName("tbl_members");
  if (!membersSheet || membersSheet.getLastRow() <= 1) return emails;

  var lastRow = membersSheet.getLastRow();
  var lastCol = Math.max(membersSheet.getLastColumn(), 7);
  var rows = membersSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var richRows = (typeof membersSheet.getRange(2, 1, lastRow - 1, lastCol).getRichTextValues === "function")
    ? membersSheet.getRange(2, 1, lastRow - 1, lastCol).getRichTextValues()
    : [];

  for (var i = 0; i < rows.length; i++) {
    var rawMember = rows[i][0];
    var richMember = (richRows && richRows[i]) ? richRows[i][0] : null;
    var accEmail = String(rows[i][1] || "").trim().toLowerCase();
    var tenureStatus = String(rows[i][5] || "").trim().toLowerCase();

    // Skip departed members
    if (tenureStatus === "departed") continue;

    var extEmail = (typeof extractEmailFromChip === "function") ? extractEmailFromChip(rawMember, richMember) : "";
    var targetEmail = extEmail || accEmail;
    if (targetEmail && emails.indexOf(targetEmail.toLowerCase()) === -1) {
      emails.push(targetEmail.toLowerCase());
    }
  }
  return emails;
}

/**
 * Retrieves the target Google Calendar instance for event creation.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {GoogleAppsScript.Calendar.Calendar|null}
 */
function getTargetCalendar(ss) {
  if (typeof CalendarApp === "undefined") return null;
  var calId = getCalendarId(ss);
  if (calId && calId !== "primary") {
    try {
      var customCal = CalendarApp.getCalendarById(calId);
      if (customCal) return customCal;
    } catch (eCal) {
      Logger.log("getCalendarById error for " + calId + ": " + eCal);
    }
  }
  try {
    return CalendarApp.getDefaultCalendar();
  } catch (eDef) {
    Logger.log("getDefaultCalendar error: " + eDef);
    return null;
  }
}

/**
 * Creates a Google Calendar event for a given row in tbl_events.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} eventsSheet
 * @param {number} rowIndex
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Object} [options]
 * @return {Object} Status report
 */
/**
 * Creates a Google Calendar event for a given row in tbl_events.
 * Creates the event in the organizer's calendar (and team calendar if distinct).
 * Team members are NOT automatically invited (organizer invites from Google Calendar directly).
 * If duration is null/empty, creates an all-day event (doesn't take a slot in agenda).
 * If duration is provided, creates a timed event spanning start time to (start time + duration).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} eventsSheet
 * @param {number} rowIndex
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Object} [options]
 * @return {Object} Status report
 */
function createCalendarEventForRow(eventsSheet, rowIndex, targetSpreadsheet, options) {
  if (!eventsSheet || rowIndex <= 1) {
    return { success: false, message: "Invalid sheet or header row" };
  }
  options = options || {};
  var ss = targetSpreadsheet || (eventsSheet.getParent ? eventsSheet.getParent() : SpreadsheetApp.getActiveSpreadsheet());
  var colMap = getEventColumnMap(eventsSheet);
  var numCols = Math.max(eventsSheet.getLastColumn(), 7);
  var rowVals = eventsSheet.getRange(rowIndex, 1, 1, numCols).getValues()[0];

  // 1. Auto-generate event ID if missing
  var eventId = String(rowVals[colMap.eventId - 1] || "").trim();
  if (!eventId) {
    eventId = generateEventId();
    eventsSheet.getRange(rowIndex, colMap.eventId).setValue(eventId);
    rowVals[colMap.eventId - 1] = eventId;
  }

  var title = String(rowVals[colMap.title - 1] || "").trim();
  var rawDate = rowVals[colMap.date - 1];
  var dateObj = parseEventDate(rawDate);
  var rawDuration = colMap.duration > 0 ? rowVals[colMap.duration - 1] : null;
  var durationMinutes = parseEventDuration(rawDuration);
  var location = colMap.location > 0 ? String(rowVals[colMap.location - 1] || "").trim() : "";
  var rawOrganizer = colMap.organizer > 0 ? rowVals[colMap.organizer - 1] : null;
  var richOrganizer = (colMap.organizer > 0 && typeof eventsSheet.getRange(rowIndex, colMap.organizer).getRichTextValue === "function")
    ? eventsSheet.getRange(rowIndex, colMap.organizer).getRichTextValue()
    : null;
  var googleCalendarEventId = colMap.googleCalendarEventId > 0 ? String(rowVals[colMap.googleCalendarEventId - 1] || "").trim() : "";

  if (!title) {
    return { success: false, message: "Event title is required", eventId: eventId };
  }
  if (!dateObj) {
    return { success: false, message: "Valid event date is required", eventId: eventId };
  }
  if (googleCalendarEventId) {
    return { success: true, alreadyExists: true, eventId: eventId, googleCalendarEventId: googleCalendarEventId };
  }

  var organizerEmail = getOrganizerEmail(rawOrganizer, richOrganizer, ss);
  var organizerCal = getOrganizerCalendar(organizerEmail, ss);
  var teamCal = getTargetCalendar(ss);

  if (!organizerCal && !teamCal) {
    return { success: false, message: "Google Calendar service not available", eventId: eventId };
  }

  // By default, do NOT invite team members. It is the job of the organizer to invite people from Google Calendar directly.
  var guestEmails = [];
  if (options.inviteMembers === true) {
    guestEmails = getActiveTeamMemberEmails(ss);
  }

  var eventOptions = {
    location: location,
    description: "Event ID: " + eventId + (location ? (" | Location: " + location) : ""),
    sendInvites: false
  };
  if (guestEmails.length > 0) {
    eventOptions.guests = guestEmails.join(",");
  }

  var primaryCal = organizerCal || teamCal;
  var createdEvent = null;

  try {
    // If duration is null/empty -> "whole day" event that doesn't take a slot in the agenda
    if (durationMinutes === null) {
      createdEvent = primaryCal.createAllDayEvent(title, dateObj, eventOptions);
      if (teamCal && organizerCal && teamCal.getId() !== organizerCal.getId()) {
        try {
          teamCal.createAllDayEvent(title, dateObj, eventOptions);
        } catch (eTeam) {
          Logger.log("teamCal createAllDayEvent note: " + eTeam);
        }
      }
    } else {
      // Timed event -> takes a slot in the agenda with duration
      var startTime = dateObj;
      var endTime = new Date(startTime.getTime() + durationMinutes * 60 * 1000);
      createdEvent = primaryCal.createEvent(title, startTime, endTime, eventOptions);
      if (teamCal && organizerCal && teamCal.getId() !== organizerCal.getId()) {
        try {
          teamCal.createEvent(title, startTime, endTime, eventOptions);
        } catch (eTeam) {
          Logger.log("teamCal createEvent note: " + eTeam);
        }
      }
    }

    if (createdEvent && createdEvent.getId) {
      var gEvtId = createdEvent.getId();
      eventsSheet.getRange(rowIndex, colMap.googleCalendarEventId).setValue(gEvtId);
      if (ss && typeof ss.toast === "function") {
        var durDesc = durationMinutes === null ? "All Day" : (durationMinutes + " min");
        ss.toast("Calendar event created for '" + title + "' (" + durDesc + ").", "Event Synchronized", 4);
      }
      return {
        success: true,
        eventId: eventId,
        googleCalendarEventId: gEvtId,
        title: title,
        isAllDay: durationMinutes === null,
        durationMinutes: durationMinutes,
        organizerEmail: organizerEmail
      };
    }
  } catch (eCreate) {
    Logger.log("createCalendarEventForRow error: " + eCreate);
    return { success: false, error: String(eCreate), eventId: eventId };
  }

  return { success: false, message: "Could not create event in Google Calendar", eventId: eventId };
}

/**
 * Synchronizes tbl_events bi-directionally with Google Calendar:
 * 1. Pushes unsynced events (empty google_calendar_event_id) to the organizer's calendar.
 * 2. Does NOT auto-invite members (organizer manages invitations directly in Google Calendar).
 * 3. Handles duration (null for all-day agenda-free, numeric/string for timed events).
 * 4. Updates existing events in Google Calendar if title, date, duration, or location was modified in sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @param {Object} [options]
 * @return {Object} Synchronization summary report
 */
function syncGoogleCalendar(targetSpreadsheet, options) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return { success: false, message: "No active spreadsheet available" };
  options = options || {};

  var eventsSheet = ss.getSheetByName("tbl_events");
  if (!eventsSheet) {
    return { success: false, message: "Sheet 'tbl_events' not found" };
  }

  var lastRow = eventsSheet.getLastRow();
  if (lastRow <= 1) {
    return {
      success: true,
      message: "tbl_events is empty. No events to synchronize.",
      createdCount: 0,
      updatedCount: 0,
      totalEvents: 0
    };
  }

  var colMap = getEventColumnMap(eventsSheet);
  var numCols = Math.max(eventsSheet.getLastColumn(), 7);
  var range = eventsSheet.getRange(2, 1, lastRow - 1, numCols);
  var values = range.getValues();
  var richRows = (typeof range.getRichTextValues === "function") ? range.getRichTextValues() : [];

  var createdCount = 0;
  var updatedCount = 0;

  for (var r = 0; r < values.length; r++) {
    var sheetRowIndex = r + 2;
    var row = values[r];

    var eventId = String(row[colMap.eventId - 1] || "").trim();
    var title = String(row[colMap.title - 1] || "").trim();
    var rawDate = row[colMap.date - 1];
    var dateObj = parseEventDate(rawDate);
    var rawDuration = colMap.duration > 0 ? row[colMap.duration - 1] : null;
    var durationMinutes = parseEventDuration(rawDuration);
    var location = colMap.location > 0 ? String(row[colMap.location - 1] || "").trim() : "";
    var rawOrganizer = colMap.organizer > 0 ? row[colMap.organizer - 1] : null;
    var richOrganizer = (colMap.organizer > 0 && richRows && richRows[r]) ? richRows[r][colMap.organizer - 1] : null;
    var googleCalendarEventId = colMap.googleCalendarEventId > 0 ? String(row[colMap.googleCalendarEventId - 1] || "").trim() : "";

    // Skip empty rows
    if (!eventId && !title && !rawDate) continue;

    // Auto-generate eventId if missing
    if (!eventId) {
      eventId = generateEventId();
      eventsSheet.getRange(sheetRowIndex, colMap.eventId).setValue(eventId);
    }

    if (!title || !dateObj) continue;

    // A. Unsynced event: create in Google Calendar
    if (!googleCalendarEventId) {
      var resCreate = createCalendarEventForRow(eventsSheet, sheetRowIndex, ss, options);
      if (resCreate && resCreate.success) {
        createdCount++;
      }
    } else {
      // B. Existing synced event: check for title, date, duration, location updates
      try {
        var organizerEmail = getOrganizerEmail(rawOrganizer, richOrganizer, ss);
        var cal = getOrganizerCalendar(organizerEmail, ss) || getTargetCalendar(ss);
        if (!cal) continue;

        var existingEvent = cal.getEventById(googleCalendarEventId);
        if (existingEvent) {
          var modified = false;

          // Check title
          if (existingEvent.getTitle && existingEvent.getTitle() !== title) {
            existingEvent.setTitle(title);
            modified = true;
          }

          // Check location
          if (existingEvent.getLocation && (existingEvent.getLocation() || "") !== location) {
            existingEvent.setLocation(location);
            modified = true;
          }

          // Check date & duration
          var curIsAllDay = typeof existingEvent.isAllDayEvent === "function" ? existingEvent.isAllDayEvent() : (durationMinutes === null);

          if (durationMinutes === null) {
            // Should be all-day
            var curStart = existingEvent.getAllDayStartDate ? existingEvent.getAllDayStartDate() : (existingEvent.getStartTime ? existingEvent.getStartTime() : null);
            var curIso = curStart ? formatEventDateIso(curStart) : "";
            var sheetIso = formatEventDateIso(dateObj);
            if (!curIsAllDay || curIso !== sheetIso) {
              if (typeof existingEvent.setAllDayDate === "function") {
                existingEvent.setAllDayDate(dateObj);
                modified = true;
              }
            }
          } else {
            // Should be timed
            var start = dateObj;
            var end = new Date(start.getTime() + durationMinutes * 60 * 1000);
            var curStart = existingEvent.getStartTime ? existingEvent.getStartTime() : null;
            var curEnd = existingEvent.getEndTime ? existingEvent.getEndTime() : null;
            var timeDiff = curIsAllDay || !curStart || !curEnd || (curStart.getTime() !== start.getTime()) || (curEnd.getTime() !== end.getTime());
            if (timeDiff) {
              if (typeof existingEvent.setTime === "function") {
                existingEvent.setTime(start, end);
                modified = true;
              }
            }
          }

          if (modified) {
            updatedCount++;
          }
        }
      } catch (eUpd) {
        Logger.log("Error updating calendar event (Row " + sheetRowIndex + "): " + eUpd);
      }
    }
  }

  return {
    success: true,
    createdCount: createdCount,
    updatedCount: updatedCount,
    totalEvents: values.length
  };
}

/**
 * Public entry point to synchronize Google Calendar events with UI toast notification.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [targetSpreadsheet]
 * @return {Object} Status report
 */
function syncCalendar(targetSpreadsheet) {
  var ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var res = syncGoogleCalendar(ss);
  if (ss && typeof ss.toast === "function") {
    if (res.success) {
      ss.toast(
        "Google Calendar Synced: " + res.createdCount + " new event(s) pushed, " + res.updatedCount + " event(s) updated.",
        "Calendar Synchronized",
        5
      );
    } else {
      ss.toast(res.message || "Failed to synchronize calendar events", "Calendar Sync Error", 5);
    }
  }
  return res;
}

/**
 * Event handler called from onEdit when tbl_events is modified.
 * Auto-generates event_id and applies typography immediately.
 * Once title and date are populated, creates the Google Calendar event in real time.
 * If an existing synced event has its title, date, duration, or location edited, propagates the change.
 *
 * @param {Object} e Apps Script onEdit event object
 * @return {Object|null}
 */
function handleEventEdit(e) {
  if (!e || !e.range) return null;
  var sheet = e.range.getSheet();
  if (!sheet || sheet.getName() !== "tbl_events") return null;

  var editedRow = e.range.getRow();
  var editedCol = e.range.getColumn();
  if (editedRow <= 1) return null; // Header row

  var colMap = getEventColumnMap(sheet);
  var numCols = Math.max(sheet.getLastColumn(), 7);
  var rowVals = sheet.getRange(editedRow, 1, 1, numCols).getValues()[0];

  // Check if row has any content
  var hasAnyValue = false;
  for (var i = 0; i < rowVals.length; i++) {
    var v = rowVals[i];
    if (v !== "" && v !== null && v !== undefined && String(v).trim() !== "") {
      hasAnyValue = true;
      break;
    }
  }
  if (!hasAnyValue) return null;

  var ss = sheet.getParent ? sheet.getParent() : SpreadsheetApp.getActiveSpreadsheet();

  // 1. Auto-generate event ID immediately if missing
  var eventId = String(rowVals[colMap.eventId - 1] || "").trim();
  if (!eventId) {
    eventId = generateEventId();
    sheet.getRange(editedRow, colMap.eventId).setValue(eventId);
    rowVals[colMap.eventId - 1] = eventId;
  }

  // 2. Apply typography formatting
  if (typeof applyTypography === "function") {
    try {
      applyTypography(sheet.getRange(editedRow, 1, 1, numCols));
    } catch (eTypo) {}
  }

  var title = String(rowVals[colMap.title - 1] || "").trim();
  var rawDate = rowVals[colMap.date - 1];
  var dateObj = parseEventDate(rawDate);
  var rawDuration = colMap.duration > 0 ? rowVals[colMap.duration - 1] : null;
  var durationMinutes = parseEventDuration(rawDuration);
  var location = colMap.location > 0 ? String(rowVals[colMap.location - 1] || "").trim() : "";
  var rawOrganizer = colMap.organizer > 0 ? rowVals[colMap.organizer - 1] : null;
  var richOrganizer = (colMap.organizer > 0 && typeof sheet.getRange(editedRow, colMap.organizer).getRichTextValue === "function")
    ? sheet.getRange(editedRow, colMap.organizer).getRichTextValue()
    : null;
  var googleCalendarEventId = colMap.googleCalendarEventId > 0 ? String(rowVals[colMap.googleCalendarEventId - 1] || "").trim() : "";

  // Case 1: Unsynced event and both title & date are populated -> create event in Google Calendar
  if (!googleCalendarEventId && title && dateObj) {
    try {
      return createCalendarEventForRow(sheet, editedRow, ss);
    } catch (eCal) {
      Logger.log("handleEventEdit createCalendarEventForRow error: " + eCal);
      return { success: false, inSheetCreated: true, eventId: eventId, error: String(eCal) };
    }
  }

  // Case 2: Synced event -> update Google Calendar if title, date, duration, location, or organizer was edited
  if (googleCalendarEventId && (
    editedCol === colMap.title ||
    editedCol === colMap.date ||
    editedCol === colMap.duration ||
    editedCol === colMap.location ||
    editedCol === colMap.organizer
  )) {
    var organizerEmail = getOrganizerEmail(rawOrganizer, richOrganizer, ss);
    var calendar = getOrganizerCalendar(organizerEmail, ss) || getTargetCalendar(ss);
    if (calendar) {
      try {
        var existingEvent = calendar.getEventById(googleCalendarEventId);
        if (existingEvent) {
          if (editedCol === colMap.title && existingEvent.setTitle) {
            existingEvent.setTitle(title);
          }
          if (editedCol === colMap.location && existingEvent.setLocation) {
            existingEvent.setLocation(location);
          }
          if ((editedCol === colMap.date || editedCol === colMap.duration) && dateObj) {
            if (durationMinutes === null) {
              if (typeof existingEvent.setAllDayDate === "function") {
                existingEvent.setAllDayDate(dateObj);
              }
            } else {
              var start = dateObj;
              var end = new Date(start.getTime() + durationMinutes * 60 * 1000);
              if (typeof existingEvent.setTime === "function") {
                existingEvent.setTime(start, end);
              }
            }
          }
          return { success: true, updated: true, eventId: eventId, googleCalendarEventId: googleCalendarEventId };
        }
      } catch (ePatch) {
        Logger.log("handleEventEdit patch error: " + ePatch);
      }
    }
  }

  return { success: true, inSheetCreated: true, eventId: eventId };
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    GOOGLE_CALENDAR_ID_KEY: GOOGLE_CALENDAR_ID_KEY,
    getCalendarId: getCalendarId,
    setCalendarId: setCalendarId,
    getEventColumnMap: getEventColumnMap,
    generateEventId: generateEventId,
    parseEventDuration: parseEventDuration,
    parseEventDate: parseEventDate,
    formatEventDateIso: formatEventDateIso,
    getOrganizerEmail: getOrganizerEmail,
    getOrganizerCalendar: getOrganizerCalendar,
    getActiveTeamMemberEmails: getActiveTeamMemberEmails,
    getTargetCalendar: getTargetCalendar,
    createCalendarEventForRow: createCalendarEventForRow,
    syncGoogleCalendar: syncGoogleCalendar,
    syncCalendar: syncCalendar,
    handleEventEdit: handleEventEdit
  };
}
