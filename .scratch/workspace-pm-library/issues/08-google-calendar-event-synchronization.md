# 08: Google Calendar & Event Invitations Synchronization

**What to build:** An integration service that connects `tbl_events` to Google Calendar. When new social events or milestone gatherings are added to `tbl_events`, this service creates calendar events using `CalendarApp`, populates title, date, and location, dispatches calendar invitations to team members, and writes back the generated `google_calendar_event_id` to the spreadsheet. If an event is updated or rescheduled in the sheet, the corresponding Google Calendar event is updated automatically.

**Blocked by:** 02: Migration Engine & Protection Manager

**Status:** closed

- [x] Scans `tbl_events` for rows where `google_calendar_event_id` is empty.
- [x] Creates the event on the configured team calendar using `CalendarApp.createEvent()`.
- [x] Invites active team members from `tbl_members` (using their Workspace emails) if invitations are enabled.
- [x] Writes back the unique `event.getId()` into `tbl_events.google_calendar_event_id`.
- [x] Updates existing calendar events if title, date, or location changes in `tbl_events`.
- [x] Document Properties stores the target Google Calendar ID per team.
