# 0003 Baselines, Leaves, and Project Backup Coverage Architecture

Context: Manually entering daily presence rows creates high friction. Team members need to define default working routines (Baselines) and date-range Leaves. Furthermore, projects need continuous oversight when a Project Lead is absent, requiring an immediate view of who is currently on duty.

Decision: 
1. We introduce two declarative range tables: `tbl_presence_baselines` (recurring patterns or date-range defaults) and `tbl_leaves` (date-range absences). An automated generation function compiles these ranges into concrete daily rows in `tbl_presence_log`, allowing individual manual overrides.
2. In `tbl_projects`, each project explicitly defines both a `lead_member_id` (Project Lead) and `backup_member_id` (Project Backup).
3. The dashboard provides an "Active Project Responsibles" live status card cross-referencing current presence: if the Project Lead is on Leave or absent, the Project Backup is automatically promoted to "Active Responsible On Duty" with clear visual alerting if both are unavailable.

Considered Options:
- Requiring team members to log every day manually row-by-row: Rejected due to high friction and incomplete attendance logs.
- Dynamic runtime formula evaluation for all historical dates without a daily log table: Rejected because formulas become sluggish over thousands of dates, and historical snapshots cannot be locked.

Consequences:
- `tbl_presence_log` remains the single source of truth for daily data, populated forward from Baselines and Leaves.
- Managers have continuous visibility into Project Backup coverage without manual check-ins.
