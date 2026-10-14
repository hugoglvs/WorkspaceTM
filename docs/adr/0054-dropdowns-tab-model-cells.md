# ADR 0054: The Dropdowns Tab and Model Cells

## Status
Accepted (2026-10-15).

## Context
The team wants dropdowns shown as chips with a colour per option (Weekly Presence, statuses). Neither Apps Script nor the Sheets API can set the chip style or the option colours (no field for them; open feature request), so code alone cannot produce them. Copying a cell's dropdown does carry them.

## Decision
1. **A hidden "Dropdowns" tab** (key `dropdowns`, `src/Dropdowns.js`, menu *Edit Dropdowns*) holds one **model cell** per list (`DROPDOWN_LISTS`), today the Presence list of Weekly Presence. Someone sets the chip style and colours on the model once, by hand (Data → Data validation).
2. **Options live beside the models.** Each model is a "Dropdown (from a range)" pointing at its list's options column in the same tab. Initialize / Upgrade and Locations edits rewrite the options (active locations, schema statuses); a model's dropdown is created only when missing, never replaced, so its style and colours survive. A new option starts without colour.
3. **The models are copied** (`applyDropdown`, paste of the data validation) onto the editable weeks of Weekly Presence when the view is rebuilt (`ViewBuilder`). Without a model, the view sets its own plain list.
4. **Tables are not fed from the tab.** Tested live: a Google Sheets table ignores a pasted dropdown (as a text column it shows none, as a dropdown column it keeps its own). Table dropdowns are already chips; their colours are set in the table itself (column menu → Edit column type, or the pencil of a cell's dropdown). `createGoogleSheetsTable` now sends a table's columns only when their names, types or lists changed (`tableColumnsSignature`), so those colours survive upgrades; a column whose list changes (a new location, task type or project) is sent again and may lose them.

## Consequences
- Verified live: Weekly Presence shows chips with the colour set on the Presence model; tables ignore pasted dropdowns.
- A list's options come from code; to add an option, change its source (Locations tab, schema), not the options column.
- Any future view (non-table) dropdown can be added to `DROPDOWN_LISTS`.
