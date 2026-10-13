# ADR 0046: Contracts Are Declared in the Contracts Tab (Members & Contracts Sidebar Removed)

## Status
Superseded by ADR 0048 (2026-10-13). Was: Accepted (2026-10-12), superseding ADR 0037.

## Context
The Members & Contracts sidebar (ADR 0037) asked for a name or a Google email when adding someone. Writing the email as a People chip from the script (Sheets API `chipRuns`) worked on a personal account but not in the Accenture Google Workspace, so new members ended up as plain text instead of chips. Typing an email in a People column of the sheet itself does make a chip, in every environment.

## Decision
1. The sidebar and its server functions (`getMembersSidebarData`, `saveMemberTenure`, `endMemberTenure`, `saveMemberProfile`) are removed, with `src/MembersService.js` and `src/MembersSidebar.html`.
2. Contracts are typed in the **Contracts** tab, like Locations and Holidays. The menu item **Declare Contract** (`showContractsTab`, `src/Contracts.js`) unhides the tab, selects its first empty row and shows a toast with what to fill in. Upgrade hides it again.
3. A new name typed in Members that has no contract opens the Contracts tab the same way (toast "X has no contract yet").
4. Check Workspace's *Members without a contract* action opens the Contracts tab.
5. Contact details and birthday are typed in the Members row directly (the birthday is still normalised on edit, ADR 0042).

## Consequences
- People chips come from Google Sheets itself, so they work in the enterprise domain.
- No more guided checks on contract dates (overlap refused, previous contract ended the day before): whoever types a new contract sets the end date of the current one. Check Workspace still reports members without a contract.
