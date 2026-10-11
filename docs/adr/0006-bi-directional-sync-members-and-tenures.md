# 0006 Bi-directional Directory and Tenure Synchronization

Context:
The team directory (`tbl_members`) represents the central personnel roster, while `tbl_tenures` captures contract terms (roles, departments, dates, FTE). Users may onboard someone by adding a People Chip either into `tbl_members` or directly into `tbl_tenures`. Maintaining duplicate entries manually leads to inconsistencies where members exist in one table but not the other. Additionally, operational contact details (such as corporate "accenture email" and "phone number") are required in the team directory for communication and logistics.

Decision:
1. Extend `tbl_members` schema to include contact fields:
   - `accenture_email`: Corporate email address (`TEXT`).
   - `phone_number`: Contact telephone number (`TEXT`).
   Full column list: `['member', 'accenture_email', 'phone_number', 'current_role', 'current_department', 'tenure_status', 'today_presence']`.
2. Implement bi-directional synchronization (`syncMembersAndTenures()`):
   - **From `tbl_tenures` to `tbl_members`**: Any unique Member appearing in `tbl_tenures` that does not exist in `tbl_members` is automatically appended to `tbl_members`.
   - **From `tbl_members` to `tbl_tenures`**: Any Member in `tbl_members` that does not yet have a record in `tbl_tenures` receives an initial default Tenure record (`start_date = TODAY`, `contract_type = "Permanent"`, `fte_rate = 1.0`).
   - **Calculated Role, Department, and Status**: For every member in `tbl_members`, their `current_role`, `current_department`, and contractual `tenure_status` (`Active`, `Upcoming`, `Departed`) are updated from active or most recent Tenure dates.
3. Expose synchronization via `onEdit` trigger and explicit menu action:
   - `Project Management > Sync Directory & Tenures`.

Consequences:
- Adding a member to either table automatically populates the other.
- Contractual roles and departments stay in sync with the directory without manual copy-paste.
- Contact details (`accenture_email`, `phone_number`) are centrally visible in the team directory.
