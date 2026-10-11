# ADR 0029: Protection Snooze Dialog, Sidebar Global Scope, and Tenure Sync Resilience

## Status
Accepted

## Context
Following deployment of the versioned library (`WorkspacePM`) and the consumer sheet wrapper (`ConsumerWrapper.gs`), three defects were encountered in live consumer spreadsheets:

1. **Protected Range Warnings Lacked Snooze Dialog:**
   When editing warning-protected cells (e.g. historical tenures), users received repetitive toast alerts with no option to temporarily dismiss them. The specification requires a modal confirmation dialog offering a 5-minute snooze.
2. **Presence & Leaves Sidebar Failed to Render:**
   Clicking **Project Management** > **Presence & Leaves Sidebar** produced no UI response.
   In Google Apps Script V8 runtime, all `.gs` / `.js` script files share a single flat global scope. In `src/Code.js`, top-level forwarder stubs (`showBaselineSidebar`, `getSidebarInitialData`, etc.) attempted conditional delegation using `typeof global !== "undefined"`. Because `global` is undefined in Apps Script (which uses `globalThis`), these stubs shadowed the actual implementations in `src/Sidebar.js`, returning `undefined` or `{}` and preventing the HTML sidebar dialog from rendering.
3. **Tenures Did Not Automatically Sync to Members:**
   Declaring a new tenure row in `tbl_tenures` failed to populate `tbl_members`. In Google Sheets, cells formatted as People Chips frequently return empty strings `""` when inspected via `range.getValues()`, storing the actual user name and email metadata inside `RichTextValue`. Consequently, `syncMembersAndTenures()` skipped valid tenure rows (`if (!tMember) continue;`). Furthermore, simple trigger execution context restrictions prevented advanced API calls from executing during simple `onEdit` events.

## Decision

1. **5-Minute Protection Warning Snooze Modal (`src/ProtectionManager.js`):**
   - Implemented dual-layer snooze caching:
     - Primary: `CacheService.getUserCache()` with 300-second TTL for zero-latency in-memory lookup.
     - Secondary / Persistent: `PropertiesService.getUserProperties()` (`PROTECTION_WARNING_SNOOZE_UNTIL`) storing epoch millisecond timestamp to persist across spreadsheet tab reloads.
   - In `handleProtectedCellEdit()`, if the user has not snoozed warnings:
     - Prompt with a blocking modal dialog: `SpreadsheetApp.getUi().alert("Protected Range Warning", "...", ui.ButtonSet.YES_NO)`.
     - **If YES:** Store snooze timestamp for 5 minutes (`snoozeProtectionWarnings(5)`) and display confirmation toast. All subsequent edits across protected cells suppress both alerts and toasts for the next 5 minutes.
     - **If NO / Close:** Display standard warning toast without activating the snooze.

2. **Sidebar Single Source of Truth & Clean Global Scope (`src/Sidebar.js`, `src/Code.js`):**
   - Eliminated shadowing stub declarations in `src/Code.js` top-level scope.
   - `src/Sidebar.js` remains the sole canonical definition of `showBaselineSidebar()`, `getSidebarInitialData()`, `getMemberBaselines()`, `saveMemberBaselines()`, and `submitMemberLeave()`.
   - In Node.js testing environments, `src/Code.js` exports delegators via `module.exports` to maintain backwards compatibility with existing test runners.
   - Added null-safety checks in `getSidebarInitialData()` and `getMemberBaselines()` for standalone and container-bound contexts.

3. **Tenures-to-Members People Chip Resilience (`src/SyncEngine.js`, `src/Code.js`, `src/ConsumerWrapper.js`):**
   - Updated `syncMembersAndTenures()` to retrieve both `range.getValues()` and `range.getRichTextValues()` simultaneously for both `tbl_tenures` and `tbl_members`.
   - Implemented `safeExtractCell(val, richText)` which inspects `richText.getText()` and chip metadata when `val` is empty or non-string.
   - Updated `findNextAvailableRow()` to verify both primitive values and rich text values before identifying empty rows.
   - In `copyCellChip()`, added fallback text writing (`autoEmail || rawName`) if native `copyTo()` or Sheets API does not populate the destination range.
   - In `src/Code.js`, wrapped `onEdit` trigger calls to `syncMembersAndTenures` in safe try/catch blocks that isolate non-critical exceptions and avoid user disruption.
   - Added `"Sync Directory & Tenures"` (`clientSyncMembersAndTenures()`) to `src/ConsumerWrapper.js` menu and exported it so administrators can trigger manual synchronization at any time.

## Consequences
- Protected range warnings can now be snoozed for 5 minutes via a standard modal dialog, eliminating notification fatigue while preserving audit integrity.
- The Presence & Leaves Sidebar opens reliably in both library-linked consumer spreadsheets and container-bound projects.
- Declaring tenures with People Chips immediately synchronizes to `tbl_members` without data loss or empty row skips.
- All 124 unit tests pass cleanly (`node --test test/*.js`).
