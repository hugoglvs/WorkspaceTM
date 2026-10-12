# ADR 0040: Configurable Work Locations, and Departed Members Hidden in Weekly Presence

## Status
Accepted (2026-10-12)

## Context
1. Work Locations were hard-coded in about 25 places (schema dropdowns, both sidebars, Weekly Presence, KPI keywords such as "main office"). The team needs its own list: OneRoof, Remote, JLL, C04, C60.
2. The KPI split (office / client site / remote) was guessed from words in the name, which fails for names like "JLL" or "C04".
3. Weekly Presence showed departed Members.

## Decision
1. **Locations tab** (registry key `locations`, hidden; menu *Project Management → Edit Locations* shows it): `Location | Type | Active`. Types are `Office`, `Client site` and `Remote`. Defaults are written when the tab is empty: OneRoof (Office), Remote (Remote), JLL, C04 and C60 (Client site). `src/Locations.js` owns reading it. A blank Active cell counts as active.
2. **Active vs history.** New entries offer **active** locations only (both sidebars, and Weekly Presence from the current week on). History stays valid: Weekly Presence past weeks carry no dropdown, and the hidden data tables (Routines, Presence Log) accept **every** listed location, active or not, plus `Off` (and leave types in Presence Log). So retiring a location never flags past cells as invalid.
3. Editing the Locations tab refreshes those dropdowns and Weekly Presence. Upgrades also refresh native table column properties; this is sent separately from renames so a refusal cannot block them.
4. **Type drives KPIs:** `KPI Data.category` is `office`, `client_site` or `remote` (plus leave / off / not_logged); client-site periods count as worked days. Unlisted (older) names are still guessed from their wording.
5. **Defaults:** a weekday without a routine defaults to the first active Office (OneRoof). The sidebar presets "All Office", "All Remote" and "Hybrid" use the first active Office and Remote locations.
6. **Weekly Presence colours** follow the type: Office green, Client site blue, Remote purple, leaves red, Off grey.
7. **Retired location still in use:** unticking Active shows a toast (no blocking) naming how many ongoing Routines (with Member names) and future Weekly Presence days still use it, so someone updates them via *Declare Presence & Leaves* or Weekly Presence. Past days are history and are not counted.
8. **Members Status** (`tenure_status`) shows a dropdown of `Active / Upcoming / Departed / No Contract`, for readability; the value is still computed from Contracts.
9. **Departed Members** are hidden in Weekly Presence (columns hidden, not removed; anyone can unhide them). Upcoming and No Contract Members stay visible. Columns are unhidden before each rebuild, so a returning Member reappears.

## Consequences
- `Off` and leave types stay built in; they are not places.
- Existing mock data with former names ("Main Office"…) is left as is (user decision).
