# 0007 Tenure-Driven Directory, Profile Onboarding Tasks, and Column Protections

Context:
Employment contracts and assignments originate in `tbl_tenures`. Allowing manual row creation in `tbl_members` leads to orphaned directory entries without contracts. Furthermore, synthetic `tenure_id` identifiers are redundant when a tenure is naturally identified by its Member (`PEOPLE_CHIP`) and date range (`start_date` to `end_date`). Directory consistency requires locking all computed and tenure-derived columns (`member`, `current_role`, `current_department`, `tenure_status`, `today_presence`), leaving only contact fields (`accenture_email`, `phone_number`) editable. When new members are onboarded, automated accountability is needed to prompt them to fill in their contact information.

Decision:
1. **Tenure as Single Source of Truth**:
   - Rows in `tbl_members` are exclusively generated and maintained from `tbl_tenures`. Manual row creation in `tbl_members` is restricted.
2. **Elimination of `tenure_id`**:
   - `tenure_id` is removed from `tbl_tenures`.
   - `tbl_tenures` columns: `['member', 'role', 'department', 'start_date', 'end_date', 'contract_type', 'fte_rate']`.
   - Tenures are uniquely keyed by Member + date range.
3. **Strict Column Protection**:
   - In `tbl_members`, protected columns are: `member`, `current_role`, `current_department`, `tenure_status`, `today_presence`.
   - Only `accenture_email` and `phone_number` remain open for editor input.
4. **Automated Onboarding Profile Task**:
   - When a Member is newly synced to `tbl_members`, an onboarding task is automatically generated in `tbl_tasks`:
     - `title`: "Complete profile: enter Accenture email and phone number in Team Directory"
     - `assigned_member`: Member People Chip
     - `due_date`: 7 days from creation
     - `status`: "Todo"

Consequences:
- No orphaned members or duplicate contract entries.
- Clear separation between read-only directory data and editable contact fields.
- Automated onboarding flow directly integrates with task tracking.
