# ADR 0047: Members Role, Department and Status Are Formulas Reading Contracts

## Status
Accepted (2026-10-13)

## Context
Role, Department and Status in Members were written by `syncMembersAndTenures`: on every Contracts or Members edit (edit trigger, Sync Lock, three writes per member) and by the 2 am job. A contract ending yesterday still showed "Active" until the night run, and the path depended on the installable edit trigger, which is the part most exposed to enterprise-domain restrictions.

## Decision
1. Each Members row holds three formulas (`memberFormulas`, `src/Contracts.js`) reading the Contracts tab by its current name:
   - **Status**: *Active* if a contract has started (or has no start) and has not ended (end blank or ≥ today); else *Upcoming* if one starts later; else *Departed*; *No Contract* if none.
   - **Role / Department**: from the earliest active contract, else the earliest upcoming one, else the contract that started last; "Team Member" / "Delivery" when blank.
   They use `TODAY()`, so they change on the day without any job. `contractProfile` states the same rule in JavaScript for tests.
2. The formulas are written per row by the script (`applyMemberFormulas`) into every row of the Members table, including empty ones, only where a cell does not already hold the right formula. It runs at the end of every members sync (edit, nightly, Initialize / Upgrade). This does not rely on Sheets copying formulas into new table rows, nor on an array formula filling a table column.
3. The sync no longer writes these three columns. Contracts edits still run the (now light) sync, to add a person typed only in Contracts to Members, with the onboarding task.
4. Check Workspace reports *Members cells typed over* (a Role, Department or Status cell without its formula) with a *Restore formulas* action (`restoreMemberFormulas`).
5. Members stays a typed list with People chips: a FILTER-generated list was rejected because typed details (email, phone, birthday) would drift to other rows when the list reorders, chips would become text, and departed members would disappear.

## Consequences
- Status is right every day, for everyone, without a trigger.
- The rule now exists twice (formula and `contractProfile`); tests evaluate the formulas with `contractProfile` in the mock, and the formulas were checked in a live sheet (active, last day, upcoming, departed, blank role, no contract).
- Nobody should type in these three columns; the strict protection and the Check Workspace check cover it.
- The required-cell highlight now counts a row as filled with `COUNTIF(row,"?*")+COUNT(row)` instead of `COUNTA`, which counted formulas returning "" and turned every empty Members row red.
