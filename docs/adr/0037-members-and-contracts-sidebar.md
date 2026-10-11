# ADR 0037: Members & Contracts Sidebar

## Status
Accepted (2026-10-12), resolves issue 22

## Context
1. Members is one of the five visible tabs (ADR 0036). New Members and Tenure changes (new contract, end date) should be declared in a sidebar, and starting a new row on Members should open it.
2. `syncMembersAndTenures` only flows from `tbl_tenures` to `tbl_members`. A name typed straight into Members has status "No Contract" until a Tenure exists.
3. Google's docs do not say whether a sidebar can be opened from an `onEdit` trigger for every editor.

## Decision
1. **Sidebar** `MembersSidebar.html` (menu *Project Management → Members & Contracts*), backed by `src/MembersSidebar.js`:
   - `getMembersSidebarData`: Members with their Tenures, plus known roles, departments and contract types.
   - `saveMemberTenure`: first contract for a new Member, or a new contract for an existing one. Any open or overlapping earlier Tenure ends **the day before** the new start. A contract starting on or before an existing one is refused. FTE is stored as a percentage (1–100).
   - `endMemberTenure`: sets the end date of one Tenure (Member + start date).
   - `saveMemberProfile`: Accenture email and phone.
   - Every write runs `syncMembersAndTenures`, so role, department and status in Members follow.
2. **New-row prompt:** `onEdit` on the Members `member` column, for a name with no Tenure, calls `showMembersSidebar(name)`, which opens pre-filled. If the trigger context refuses to open UI, a toast points to the menu instead.
3. Grids stay editable (ADR 0034): typing in `tbl_tenures` or Members still syncs as before.
4. Sidebar dates use the browser's **local** date (a UTC `toISOString` default was a day early between 00:00 and 02:00 in Paris; also fixed in the Presence & Leaves sidebar).

## Consequences
- The sidebar server functions need forwarders in the Consumer Sheet wrapper (`consumer/Code.js`), because `google.script.run` calls the bound script.
- Whether the automatic opening works for editors other than the script owner must be confirmed live; the toast covers the case where it does not.
