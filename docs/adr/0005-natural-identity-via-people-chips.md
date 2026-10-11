# 0005 Natural Identity via Google People Chips

Context:
Originally, tables used synthetic database keys (`member_id`) to identify individuals. However, in Google Workspace and native Google Sheets Tables, team members are naturally represented by People Chips (`PEOPLE_CHIP`). People Chips already bind globally unique Workspace user identities (email, full name, and avatar) directly within Google Sheets. Requiring human users to manage or look up synthetic `member_id` keys creates friction and duplicate identity fields.

Decision:
1. Eliminate synthetic `member_id` across all tables.
2. Use `member` (typed as `PEOPLE_CHIP`) as the natural identity column across:
   - `tbl_members`: Column 1 is `member`.
   - `tbl_tenures`: `member` references the person's People Chip.
   - `tbl_presence_baselines`: `member` references the person's People Chip.
   - `tbl_leaves`: `member` references the person's People Chip.
   - `tbl_presence_log`: `member` references the person's People Chip.
   - `tbl_projects`: `lead_member` and `backup_member`.
   - `tbl_tasks`: `assigned_member`.
3. In the Presence Baseline Sidebar and lookup formulas, identity matching operates directly on the `member` value.

Consequences:
- Entering a member in `tbl_members` simply requires typing `@name` to insert a People Chip in Column 1.
- No synthetic ID generation or manual ID typing is needed.
- Data across all operational tables seamlessly connects using the native Google Workspace identity.
