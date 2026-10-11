# 14. Active User Status Audit and Personalized Sidebar Checklist

Date: 2026-10-10

## Status

Accepted

## Context

To enhance team operations and compliance without burdensome manual audits, users need automated feedback tailored specifically to them:
1. When opening the spreadsheet (`onOpen`), the system should identify the logged-in user via `Session.getActiveUser().getEmail()` and alert them if they have missing required profile information (e.g. role, department) or haven't configured their weekly presence baseline routine.
2. In the Presence & Leaves Sidebar (`BaselineSidebar.html`), the interface should present a personalized status card displaying the active user's identity, routine baseline configuration status, profile completeness, and current presence status for today.
3. In accordance with ADR 0005 ("Natural Identity via People Chips"), member identity is strictly anchored on native Google People Chips. `accenture_email` is merely a contact field and must never be used for identity resolution.

## Decisions

1. **Active User Audit Engine & People Chip Email Matching (`src/Sidebar.js` & `src/Code.js`)**:
   - Implemented `checkActiveUserStatus(targetSpreadsheet, showToast)` with robust People Chip email extraction:
     - Extracts the current authenticated user email via `Session.getActiveUser().getEmail()`.
     - In Google Sheets, a People Chip often displays a formatted name (e.g. `"Hugo Gonçalves"`) rather than raw text email.
     - Implemented `extractEmailFromChip(val, richText)` and `isUserMemberMatch(memberVal, memberRichText, userEmail)`:
       - Strictly compares the People Chip itself (never compares to `accenture_email`, preserving natural identity as per ADR 0005).
       - Extracts email from Smart Chip objects (`.getEmail()`, `.email`, `.getPersonEmail()`).
       - Extracts email from RichText `mailto:` link URL (`richText.getLinkUrl()`, text runs).
       - Matches normalized display names with diacritic stripping against corporate email username prefixes (e.g. `"Hugo Gonçalves"` <-> `"hugo.goncalves@..."`).
     - Verifies whether the user has at least one active routine in `tbl_presence_baselines` using the same People Chip matching.
     - Audits required profile fields (`Role`, `Department`).
     - Reads today's presence status from `tbl_members.today_presence`.
     - When `showToast` is true and warnings exist, displays a non-blocking toast warning (`SpreadsheetApp.getActiveSpreadsheet().toast(...)`) prompting them to open the sidebar.
2. **Spreadsheet Open Trigger (`src/Code.js`)**:
   - `onOpen()` executes `checkActiveUserStatus(ss, true)` upon load. If the user has everything set up, it remains silent; if action is needed, a targeted toast alerts them.
3. **Personalized Checklist Card in Sidebar (`src/BaselineSidebar.html` & `src/Sidebar.js`)**:
   - `getSidebarInitialData()` includes `userChecklist` metadata and attaches extracted chip emails to member choices for auto-selecting the user's People Chip.
   - `BaselineSidebar.html` renders a Google Workspace styled `#userStatusCard` with user avatar, name, and three live checklist indicators:
     - **Routine Baseline**: `[Configured (N rows)]` vs `[Missing - Set Below]`
     - **Profile Fields**: `[Complete]` vs `[Missing Role / Department]`
     - **Today Status**: `[Location / Weekend / Vacation / Not Logged]`

## Consequences

- Team members receive immediate, personalized guidance to complete their schedules and profiles without administrators having to chase them manually.
- Native Google People Chips are seamlessly compared to corporate email addresses across Google Sheets API representations without relying on contact columns.
- Clean separation between active user identification and the member dropdown selection (allowing leads to view other members while still seeing their own status badge).
- 100% test coverage with automated unit tests across 43 test cases.
