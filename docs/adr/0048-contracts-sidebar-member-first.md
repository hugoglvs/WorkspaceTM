# ADR 0048: Member First, Contracts in a Sidebar (Supersedes ADR 0046)

## Status
Accepted (2026-10-13). Supersedes ADR 0046 (contracts typed in the unhidden Contracts tab).

## Context
ADR 0046 dropped the Members & Contracts sidebar because writing a People chip from a typed email did not work in the enterprise domain, and made *Declare Contract* unhide the Contracts tab. Exposing a raw table for every contract change is clumsy, and the chip problem only came from creating the person in the sidebar.

## Decision
1. **A person starts in Members**: their name or email is typed in the Member column, where Sheets itself makes the People chip.
2. **Contracts sidebar** (`src/ContractsSidebar.html`, server functions in `src/Contracts.js`), opened by the menu *Declare Contract* (`showContractsSidebar`) or automatically when a new name without a contract is typed in Members (toast pointing to the menu if the sidebar cannot open):
   - pick a Member (only people already in Members; no email field, no chip writing);
   - see all their contracts, newest first, marked current / upcoming / ended;
   - **add a contract**: prefilled from the latest one; a contract still running on the new start date ends the day before by default (or on the date given); a start on or before an existing contract is refused; a gap between contracts is reported;
   - **edit any contract** in place (role, department, type, FTE, dates), refused if it would overlap another contract of the same person.
   `saveContract` writes the Contracts row and copies the name cell from Members (same chip); the Members formulas (ADR 0047) update at once.
3. The Contracts tab stays hidden; it remains editable for corrections, like every table (ADR 0034).
4. A Member without a contract is visible in the sheet: their Role and Department are empty, so those required cells are highlighted red; Status shows *No Contract*, and Check Workspace lists them with a *Declare Contract* button.

## Consequences
- One entry path with checks for contracts; no raw table exposed to the team.
- Automatic opening when a name is typed is verified for the trigger owner only; for other editors the toast fallback applies.
