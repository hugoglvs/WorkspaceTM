# 06: Presence & Remote Work KPI Engine

**What to build:** An automated calculation and analytics sheet (`calc_kpis`) that evaluates attendance metrics across members, departments, and months. It computes remote work rates, in-office presence rates, and team capacity. The calculation engine explicitly accounts for intra-day mobility by weighting morning (AM) and afternoon (PM) periods at 0.5 working days each, producing accurate metrics even when members split their days across multiple offices or between home and office.

**Blocked by:** 03: Baseline & Leave Range Compiler to Daily Presence Log

**Status:** closed

- [x] Sheet `calc_kpis` generates aggregated metrics grouped by Month, Department, and Member.
- [x] Presence calculation correctly weights half-day periods:
  - Working periods = count of AM/PM slots excluding non-working leaves/holidays.
  - Remote rate = `(count of Remote periods * 0.5) / (total worked periods * 0.5)`.
  - In-Office rate = `(count of Physical Office periods * 0.5) / (total worked periods * 0.5)`.
- [x] Provides office-specific occupancy breakdowns (e.g. attendance distribution across Main Office, Downtown Office, Branch Office, Client Site).
- [x] Incorporates FTE rates from `tbl_tenures` to compute effective team capacity versus actual attendance.
- [x] Output table is formatted with 12px "Google Sans" and native percentage number formats.
