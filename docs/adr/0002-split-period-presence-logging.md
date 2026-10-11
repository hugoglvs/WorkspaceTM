# 0002 Split-Period (AM/PM) Presence Logging

Context: Team members frequently transition between physical office locations or between remote and office work within the same working day. We need a way to log and calculate presence metrics that accommodates intra-day movement without creating unnecessary data entry friction.

Decision: We will record presence as a single daily row per Member in `tbl_presence_log` featuring dedicated `location_am` and `location_pm` columns, rather than separate rows per half-day period. Each period represents 0.5 FTE units in presence calculations.

Considered Options:
- Two rows per member per day (with a `period` column): Rejected because it doubles the number of rows users must locate and edit, creating higher friction in Google Sheets interfaces.
- Single daily presence type with notes: Rejected because it does not support precise split-location metrics or office capacity tracking.

Consequences:
- Calculating full-day versus half-day presence or remote rates requires weighting AM and PM columns at 0.5 each.
- Visual planning dashboards can display AM/PM sub-columns or split badges per day.
