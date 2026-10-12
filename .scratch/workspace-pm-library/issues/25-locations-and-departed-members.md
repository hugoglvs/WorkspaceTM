# 25: Configurable Locations, and Departed Members Hidden in Weekly Presence

**Status:** resolved

**Blocked by:** _(none)_

**What to build:** A hidden Locations tab (Location | Type | Active) feeding every location dropdown and the KPI categories, plus departed Members hidden in Weekly Presence (ADR 0040).

- [x] `src/Locations.js`, Locations schema and registry entry, defaults OneRoof / Remote / JLL / C04 / C60.
- [x] Active locations only for new entries; history stays valid (no dropdown on past weeks, data tables accept all).
- [x] Edit hook and the *Edit Locations* menu item; KPI `client_site` category; default office instead of "Main Office".
- [x] Departed Members' columns hidden; Upcoming shown.

## Comments

_2026-10-12 — 9 new tests in `test/locations.test.js`; suite 174 pass._

## Follow-up (2026-10-12)
- [x] Toast when a location is unticked but ongoing Routines or future days still use it (`warnRetiredLocation`).
- [x] Members Status column shown as a dropdown.
- [ ] "Today" column: colours vs dropdown, awaiting decision.
