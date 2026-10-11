# ADR 0017: Historical Presence Compilation and Data Migration

## Status
Accepted

## Context
When bootstrapping a new Google Spreadsheet or importing historical records from legacy spreadsheets, project leads and administrators frequently load existing member tenures, presence baselines, and historical leave bookings.

The standard compilation routine (`compilePresence()`) is optimized for day-to-day operations and forward-looking routine compilations: it defaults to compiling from the 1st of the current month up to today (`todayDate`). However, during initial project setup or data migration, baselines and tenures often start months or years in the past (e.g. 2024 or 2025). Furthermore, imported baselines frequently omit an explicit `start_date` (leaving it undefined/blank), indicating that the baseline applies unconditionally from the beginning of the member's tenure.

Without a dedicated migration compilation option:
1. Historical dates prior to the current month were not compiled into `tbl_presence_log`.
2. Administrators had no simple UI mechanism to back-populate attendance across the entire project history.
3. The Weekly Matrix Planning View (`view_weekly_matrix`) lacked historical attendance context.

## Decision
1. **Earliest Date Auto-Discovery (`findEarliestHistoricalDate`):**
   - Implemented `findEarliestHistoricalDate(targetSpreadsheet)` in `src/PresenceEngine.js`.
   - Programmatically scans all relevant operational tables for the earliest recorded date:
     - `tbl_tenures` (Column 4: `start_date`)
     - `tbl_presence_baselines` (Column 3: `start_date`)
     - `tbl_leaves` (Column 3: `start_date`)
     - `tbl_presence_log` (Column 2: `date`)
   - Normalizes and filters out invalid or unreasonable dates (retaining dates between years 2000 and 2100).
   - If no dates are found in any table, cleanly falls back to January 1 of the current year.

2. **Full History / Migration Compilation (`compilePresenceMigration`):**
   - Implemented `compilePresenceMigration(targetSpreadsheet)` in `src/PresenceEngine.js` and exposed in `src/Code.js`.
   - When executed interactively from Google Sheets UI:
     - Prompts the administrator via `ui.prompt`, displaying the auto-detected earliest recorded date.
     - Allows the user to either accept the detected earliest date (clicking OK) or input an explicit custom start date (`YYYY-MM-DD`).
     - Supports graceful cancellation (`ui.Button.CANCEL`), leaving the spreadsheet untouched.
   - When executed programmatically or in test environments without active prompt UI, automatically runs using the earliest discovered date.

3. **Undefined Baseline Start Date Support:**
   - In `compilePresenceLogs`:
     - If a member's baseline in `tbl_presence_baselines` has an undefined/empty `start_date`, the rule applies backwards across all weekdays up to its `end_date` (or open-ended if `end_date` is also empty).
     - Idempotence is preserved: existing daily log rows with manual edits are never overwritten.
     - Approved leaves in `tbl_leaves` override baseline locations for their respective date spans.

4. **UI Menu Integration:**
   - Added menu item in `onOpen()`:
     **Project Management > Compile Presence Logs (Migration / Undefined Start Date)** (`compilePresenceMigration`).
   - Automatically triggers a refresh of the Weekly Matrix view (`refreshWeeklyMatrix(ss)`) upon successful completion.
   - Displays a detailed confirmation toast summarizing compiled and preserved records.

## Consequences
- Administrators can perform complete one-click historical data migrations from undefined or earliest start dates.
- Preserves the ground-truth principle: historical compilation is strictly capped at today (`todayDate`); future dates remain in presentation view.
- Comprehensive unit test coverage added in `test/presence.test.js` (62/62 tests passing across test suite).
