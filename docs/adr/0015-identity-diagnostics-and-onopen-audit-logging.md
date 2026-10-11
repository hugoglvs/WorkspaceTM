# ADR 0015: Identity Diagnostics and onOpen Audit Logging

## Status
Accepted

## Context
When opening the spreadsheet, users observed alerts stating that their profile was not registered in `tbl_members`, even when they were present in the directory. In Google Apps Script, matching the active user to Google People Chips can fail due to:
1. `Session.getActiveUser().getEmail()` returning an empty string in unauthenticated or domain-scoped simple triggers (`onOpen`).
2. Discrepancies between Google People Chip internal representations (display name, RichText mailto URL, Smart Chip person properties) and active user email strings.
3. Lack of runtime visibility into why `isUserMemberMatch` accepted or rejected each member row.

The user requested comprehensive diagnostic logging of active user information and members' email data upon `onOpen` so that failures can be debugged immediately.

## Decision
1. **Instrument `onOpen` with Structured Audit Logging:**
   - Call `auditUserAndMembersIdentity(spreadsheet)` on `onOpen`.
   - Log `Session.getActiveUser().getEmail()`, `Session.getEffectiveUser().getEmail()`, and execution context.
   - If active user email is blank, log a clear notice explaining simple trigger privacy restrictions in Google Apps Script.
2. **Detailed Per-Row Diagnostics for `tbl_members`:**
   - For every member row in `tbl_members`, log:
     - Row index, display name, cell value data type.
     - Extracted People Chip email (from Smart Chip object or RichText link URL).
     - Link URL and text run URLs.
     - Detailed match result with explicit reasoning (exact email match, text match, normalized name token match, or failure reason).
3. **Direct Extraction of People Smart Chip Metadata via Sheets API v4:**
   - Standard Google Apps Script `Range.getValues()` and `Range.getRichTextValues()` only return the display string (e.g. `'Hugo Gonçalves'`) with a null `getLinkUrl()` for native consumer Google People Chips.
   - Introduced `fetchMemberChipEmails(spreadsheetId, sheetName, startRow, numRows)` which uses `Sheets.Spreadsheets.get` to read native `cell.chipRuns[].chip.personProperties.email` metadata.
4. **Name Contraction and Abbreviation Matching:**
   - Handles email handles that combine first names and last name consonant contractions (e.g. `hugoglvs42@gmail.com` matching `'Hugo Gonçalves'` via `hugo` + `glvs` in `goncalves`).
5. **Dedicated Interactive Debug Menu Item:**
   - Add **Project Management > Debug User & Member Identity** (`debugUserIdentity`).
   - Running this command operates in an authorized user context (guaranteeing `Session.getActiveUser()` access) and displays the diagnostic report directly in a `SpreadsheetApp.getUi().alert()` dialog in addition to `Logger.log()`.
6. **Enhanced Extraction & Explainability:**
   - `extractCellValue` now checks `richText.getText()` before falling back to `String(val)` to prevent `"[object Object]"` strings for People Chips.
   - `extractEmailFromChip` parses query parameters (`mailto:...?...`), `personProperties`, and standard URLs.
   - `explainUserMemberMatch` returns `{ match, reason, chipEmail, memberStr }`.

## Consequences
- Full visibility into Apps Script executions via **Extensions > Apps Script > Executions**.
- Direct in-sheet debugging through the custom menu without needing to open the script editor.
- Preserves ADR 0005: strictly inspects Column 1 (`member` People Chip) and never relies on `accenture_email`.
