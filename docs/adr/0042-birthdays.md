# ADR 0042: Birthdays: Day and Month in Members, Reminder Email and Team Birthdays Calendar

## Status
Accepted (2026-10-12); the day-and-month storage is amended by ADR 0049 (full birth date).

## Context
The team wants each Member's birthday in Members, a reminder to the rest of the team, and an easy way to put it in their Google Calendar. The sheet is fully transparent (no privacy model), so a full birth date would expose everyone's age.

## Decision
1. **Members "Birthday" column** (`birthday`, last column, TEXT, optional) holds **day and month only**, stored as `14 Mar`. Typing `14/03`, `March 14`, `1990-03-14` or a date is rewritten to that form by the edit trigger; anything unreadable gets a toast. Blank = not shared. The Members & Contracts sidebar shows it in Contact details and the new-member form (`src/Birthdays.js`).
2. **Reminder on the working day before** (trigger `birthdayReminderJob`, daily 9 am, nothing at weekends): a Friday announces Saturday, Sunday and Monday birthdays. Every Active Member with an email gets one email listing the coming birthdays except their own; Departed members are neither celebrated nor emailed. 29 Feb is celebrated on 28 Feb in other years. A document property stops a second send on the same day.
3. **"Add to my Google Calendar"**: each birthday in the email links to Google Calendar's event template, a yearly all-day event marked *available* (`recur=RRULE:FREQ=YEARLY`), saved with one click in the reader's own calendar.
4. **Team Birthdays calendar**: the script (as the owner, from the daily job and the nightly sync, never from sidebars) keeps a calendar "Team Birthdays – <sheet>" with one yearly all-day event per Active Member with a birthday, recreated when a name or birthday changes and removed when someone leaves. It is shared read-only, without share emails, with each Active Member's Google account (people chip, else the email column) through the Calendar advanced service; access is removed when they leave. The reminder email offers a one-time subscribe link.

## Consequences
- The Calendar advanced service is enabled in the Library and the Consumer manifests, so Google asks to re-authorise once.
- Accenture addresses without a Google account can open the template link only if they are signed in to Google; subscribing needs the Google account named in the people chip.
- 29 Feb events in the shared calendar only show in leap years (the reminder still covers 28 Feb).
