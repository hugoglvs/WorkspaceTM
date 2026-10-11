# ADR 0027: Event Datetime, Duration Column, and Organizer Calendar Synchronization

## Status
Accepted

## Context
1. In ADR 0026, calendar events in `tbl_events` were synchronized as whole-day events on the configured team calendar, and active team members were automatically sent calendar invitations.
2. User operational feedback identified three crucial refinements:
   - **Organizer Ownership & Invitations:** The event should be created on the organizer's calendar, empowering the organizer to invite guests directly from Google Calendar, rather than dispatching automated broadcast invitations from Google Sheets.
   - **DateTime & Duration:** Event `date` values should support date-times (e.g. `2026-10-15 14:00`), and `tbl_events` requires a dedicated `duration` column (in minutes or formatted durations such as `60`, `1h`, `1.5h`).
   - **Agenda-Free Whole-Day Events:** When `duration` is `null` (or empty), the event must be created as an all-day event (`CalendarApp.createAllDayEvent`) placed in the top date banner without occupying an hourly time slot in the user's agenda. When `duration` is specified, it occupies a timed slot from the start time to `(start time + duration)`.

## Decision

1. **Schema Extension (`src/Schema.js`):**
   - Updated `SCHEMA_SPEC.tbl_events`:
     - Columns: `["event_id", "title", "date", "duration", "location", "organizer", "google_calendar_event_id"]`.
     - `requiredColumns`: `["event_id", "title", "date"]`.
     - `nullableColumns`: `["duration", "location", "organizer", "google_calendar_event_id"]`.
     - Column types (Google Sheets API `TableColumnProperties.ColumnType` enum):
       - `date`: `"DATE"` (accepts both dates and timestamps/datetimes).
       - `duration`: `"DOUBLE"` (numeric column type representing minutes).
       - `organizer`: `"PEOPLE_CHIP"` (organizer member chip/email).
       - `location`: `"TEXT"`.
       - `google_calendar_event_id`: `"TEXT"`.

2. **Native Table Type Support (`src/Typography.js`):**
   - In Google Sheets API v4 `TableColumnProperties.ColumnType`, valid enum values are `["TEXT", "DATE", "DOUBLE", "BOOLEAN", "DROPDOWN", "PEOPLE_CHIP", "FILE_CHIP", "CURRENCY", "PERCENT"]`. The API rejects `"DATETIME"` or `"INTEGER"`.
   - `createGoogleSheetsTable` strictly maps `DATE` (and any datetime representation) to `"DATE"` and numeric types (`INTEGER`, `NUMBER`, `DOUBLE`) to `"DOUBLE"`.

3. **Duration Parsing & Date-Time Resolution (`src/CalendarSyncService.js`):**
   - Added `parseEventDuration(rawDuration)`:
     - Returns `null` if empty, null, or invalid.
     - Parses integers (`60`), minute strings (`"45m"`, `"90 mins"`), and hour expressions (`"1h"`, `"1.5 hours"`).
   - Enhanced `parseEventDate(rawDate)`:
     - Preserves hours, minutes, and seconds for ISO and human date-time strings (`YYYY-MM-DD HH:mm:ss`), as well as existing Date objects.

4. **Organizer Calendar Resolution & Invitation Model (`src/CalendarSyncService.js`):**
   - Added `getOrganizerEmail(rawOrganizer, richOrganizer, ss)`:
     - Resolves email from People Chips, `mailto` links, plain email text, or by matching the member's name against `tbl_members`.
     - Falls back to `Session.getActiveUser().getEmail()`.
   - Added `getOrganizerCalendar(organizerEmail, ss)`:
     - Retrieves the organizer's calendar via `CalendarApp.getCalendarById(organizerEmail)`.
     - If the organizer is the active user or if custom access is restricted, uses `CalendarApp.getDefaultCalendar()`.
   - Removed automated guest broadcasts: `eventOptions.sendInvites = false`, `eventOptions.guests = ""` by default. The organizer owns the event and manages invitations directly in Google Calendar.

5. **Timed Slot vs. Whole-Day Creation:**
   - If `durationMinutes === null`: Calls `calendar.createAllDayEvent(title, dateObj, eventOptions)` (agenda-free banner).
   - If `durationMinutes !== null`: Calculates `endTime = new Date(startTime.getTime() + durationMinutes * 60 * 1000)` and calls `calendar.createEvent(title, startTime, endTime, eventOptions)` (takes exact agenda slot).
   - If a distinct team calendar is configured in `GOOGLE_CALENDAR_ID`, the event is also synchronized on the team calendar.

6. **Bi-Directional Updates & Live Edits:**
   - In `syncGoogleCalendar` and `handleEventEdit`:
     - Edits to `title`, `location`, `date`, `duration`, or `organizer` propagate to the calendar event.
     - Transitions between whole-day and timed events toggle between `setAllDayDate(date)` and `setTime(startTime, endTime)`.

## Consequences
- Organizers own events in their personal Google Calendars with full control over guest invitations.
- Supports both all-day milestone gatherings (agenda-free) and scheduled meetings/workshops (timed slots).
- 117/117 automated unit tests passing cleanly across the test suite (`test/*.js`).
