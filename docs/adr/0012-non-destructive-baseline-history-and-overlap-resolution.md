# 12. Non-Destructive Baseline History and Overlap Resolution

Date: 2026-10-10

## Status

Accepted

## Context

Previously, `saveMemberBaselines()` purged all pre-existing baseline rows for a member when a new baseline was declared, completely wiping out historical schedule routines. Furthermore, there was no validation preventing overlapping effective date ranges between successive baseline routines.

Members need to maintain a historical record of their previous baselines over time (e.g. routine Q1 vs routine Q2). When submitting a new baseline, the system must not delete or overwrite the previous ones. It must detect date overlaps with existing baselines, prompt the user for confirmation to set the previous baseline's `end_date` to the new baseline's `start_date`, and log the new routine as appended rows.

## Decision

1. **Non-Destructive Appending**:
   - `saveMemberBaselines()` appends new 5-day weekday rows to `tbl_presence_baselines` without clearing or removing older routines.
   - The entire audit trail of past and upcoming baseline routines is preserved.

2. **Overlap Detection (`checkBaselineDateOverlap`)**:
   - Compares the submitted `[startDate, endDate]` range against existing baselines for the member.
   - Two ranges $[S_1, E_1]$ and $[S_2, E_2]$ overlap if $S_1 \le (E_2 || \infty)$ and $S_2 < (E_1 || \infty)$.
   - Adjacent boundaries where $S_2 = E_1$ are considered contiguous and non-overlapping.

3. **Interactive User Resolution Prompt**:
   - When an overlap is detected and the payload does not contain `adjustPreviousEnd: true`, the server returns `{ success: false, overlap: true, message: ... }`.
   - The Presence Baseline Sidebar UI renders an interactive modal dialog asking:
     *"A previous baseline overlaps with the new date range (starts [start]). Do you want to set the end date of the previous baseline to the start date of the new baseline ([newStart])?"*
   - If confirmed by the user, the request resubmits with `adjustPreviousEnd: true`.

4. **Automated End Date Adjustment**:
   - Upon confirmation, the `end_date` of overlapping prior baselines is updated to the `start_date` of the newly declared baseline.
   - The new baseline rows are appended starting at the next available row.

5. **Retrieval of Latest Baseline (`getMemberBaselines`)**:
   - `getMemberBaselines(memberIdentifier)` groups baselines by `startDate` and returns the most recent active baseline to populate the sidebar UI by default.

## Consequences

### Positive
- Complete schedule history is preserved for auditing and analytical reporting.
- Overlapping ambiguity is systematically prevented.
- Clear, interactive user experience with confirmation modal dialogs in the sidebar.
- All 35 unit tests pass cleanly.
