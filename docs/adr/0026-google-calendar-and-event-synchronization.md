# ADR 0026: Google Calendar and Event Synchronization

## Status
Accepted

## Context
1. In the project workspace, `tbl_events` tracks team gatherings, milestones, offsites, and social events.
2. Previously, these events remained static inside the spreadsheet without calendar presence, preventing team members from receiving calendar notifications or having events reflected in their personal schedules.
3. Teams require automated two-way synchronization between `tbl_events` and Google Calendar via Apps Script's `CalendarApp`.
4. Team invitations should automatically include all active members from `tbl_members` (using their Google Workspace / Accenture email accounts) while strictly omitting departed members.
5. In-sheet edits to event attributes (`title`, `date`, `location`) should propagate to existing Google Calendar events without duplicating events.
6. Target calendar management needs to be customizable through Document Properties (`GOOGLE_CALENDAR_ID`) with graceful fallback to the primary calendar (`"primary"`).

## Decision

1. **CalendarSyncService Module (`src/CalendarSyncService.js`):**
   - Implemented `getCalendarId(ss)` and `setCalendarId(id)` using `DocumentProperties` with key `"GOOGLE_CALENDAR_ID"` (defaulting to `"primary"`).
   - Implemented dynamic column detection via `getEventColumnMap(eventsSheet)`.
   - Implemented `generateEventId()` producing deterministic identifiers in the format `EVT-YYMMDD-XXXX`.
   - Implemented `getActiveTeamMemberEmails(ss)` scanning `tbl_members` for rows where `tenure_status` is not `"Departed"`, extracting emails from column 2 or People Chips via `extractEmailFromChip`.
   - Implemented `createCalendarEventForRow(eventsSheet, rowIndex, targetSpreadsheet, options)`:
     - Automatically generates `event_id` if blank.
     - Creates all-day events via `CalendarApp.createAllDayEvent(title, date, options)` on the target calendar.
     - Adds active team member emails to `options.guests` (comma-separated string).
     - Writes back `event.getId()` into `google_calendar_event_id` column.
   - Implemented `syncGoogleCalendar(ss, options)` and `syncCalendar(ss)`:
     - Iterates through all rows in `tbl_events`.
     - Pushes unsynced rows to Google Calendar.
     - Detects in-sheet edits on existing events (by comparing sheet values with existing calendar event `title`, `date`, and `location`), patching the Google Calendar event accordingly.
   - Implemented `handleEventEdit(e)`:
     - Captures edits on `tbl_events` in real time.
     - Automatically generates `event_id` and pushes newly titled/dated events to Google Calendar.
     - Updates existing events when title, date, or location is modified.

2. **UI and Trigger Wiring (`src/Code.js`):**
   - Added `syncCalendarEvents(targetSpreadsheet)` top-level function.
   - Added `"Sync Calendar Events"` item to `onOpen` menu under `Project Management`.
   - Wired `tbl_events` routing inside `onEdit(e)` and installable trigger `installedTaskEditHandler(e)` to `CalendarSyncService.handleEventEdit(e)`.
   - Exported `syncCalendarEvents` in `module.exports`.

3. **Comprehensive Test Suite & Mocking (`test/mockGas.js`, `test/calendarSync.test.js`):**
   - Extended `test/mockGas.js` with full mock implementation of `CalendarApp` (`createAllDayEvent`, `createEvent`, `getEventById`, `getDefaultCalendar`, `getCalendarById`, `getEvents`, and event mutation methods `setTitle`, `setLocation`, `setAllDayDate`, `addGuest`, `getGuestList`).
   - Created `test/calendarSync.test.js` covering calendar configuration, email gathering, event pushing, guest list verification, field updates, live edit routing, and menu registration.

## Consequences
- Team events in `tbl_events` synchronize directly with Google Calendar.
- Active team members automatically receive calendar invites while departed members are excluded.
- Real-time updates to title, date, and location propagate to Google Calendar seamlessly.
- 115/115 automated unit tests passing across the test suite (`test/*.js`).
