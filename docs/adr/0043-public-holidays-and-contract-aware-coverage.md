# ADR 0043: Public Holidays, and Coverage That Knows Contracts

## Status
Accepted (2026-10-12)

## Context
1. Every weekday counted as a working day: on a public holiday Members showed "Not Logged", the Monday reminder asked people to declare it, coverage could raise false conflicts, KPI capacity counted it, and the "working day before" birthday rule ignored bridges.
2. Coverage only looked at leaves. A lead or backup whose contract had ended, or ends soon, still counted as available.

## Decision
1. **Holidays tab** (registry key `holidays`, hidden, `Date | Holiday`; menu *Edit Holidays*; `src/Holidays.js`). The 11 French public holidays (Easter-based ones computed) of this year and next are added once per year (document property `HOLIDAYS_SEEDED_YEARS`), at Initialize / Upgrade and by the nightly job, so a deleted line stays deleted. The team adds its own days (bridges, office closures).
2. **A public holiday is a non-working day everywhere:**
   - Presence compile writes nothing for it (a hand-entered row is kept); Members *Today* shows "Public holiday" (grey).
   - Weekly Presence greys the row like a weekend, with the holiday name as a note on the date.
   - KPI Data: no row unless logged, no capacity; new column `is_public_holiday`.
   - Coverage skips it; the Monday reminder and Check Workspace do not count it as undeclared; birthdays are announced on the working day before, skipping holidays like weekends, and nothing is sent on a holiday.
3. **Coverage knows contracts.** A day outside every contract of a Member (left, or not started yet) counts as an approved absence. A Member with no contract at all is not judged (older data). When someone's contracts all end, the check runs to the 366-day horizon, and the reason says the lead's contract ends.
4. **Check Workspace "People leaving"**: open projects led or backed up by someone who has left or leaves within 30 days, and open tasks assigned to them, linking to Members & Contracts.

## Consequences
- Whit Monday is included (it is a public holiday even when companies work it); delete it in the Holidays tab if the team works.
- Issue 20 (IDs for leaves typed in the hidden grid) is closed as not needed.
