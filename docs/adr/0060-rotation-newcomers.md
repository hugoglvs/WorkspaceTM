# ADR 0060: Newcomers Join Rotations After a Delay, at the Back of the Queue

## Status
Accepted (2026-10-16). Extends ADR 0045 (Rotation).

## Context
A Rotation gives each turn to the person whose last turn is the oldest. A newcomer has never served, so they were picked at once, often in their first week, and they jumped ahead of everyone already waiting. Turns already created for the coming periods were also never re-dealt when someone joined, so the newcomer only appeared once the generator reached periods beyond them.

## Decision
1. **Joiner delay.** A member whose first contract (Contracts) started less than the setting **"Days before a newcomer joins a rotation"** (`rotation_joiner_delay`, default 14) before a turn's due date is skipped for that turn.
2. **Back of the queue.** A newcomer who never served, in a rotation that began before they joined, counts as having served the newest turn handed out before their delay ends: they wait behind everyone who served before them instead of going first. A rotation created after they joined treats them like everyone else.
3. **Re-plan on New member.** When a New member is detected (`fireNewMemberTriggers`), `replanRotationsForNewcomer` (`src/RecurringTasks.js`) re-deals, for each active Rotation rule that includes them, the coming open turns from the end of their delay (or today, if later). Turns that are done, cancelled, past or edited by hand (manual) are not touched; only the member of a turn whose pick changed is rewritten, so running it twice changes nothing.

Rules are generic: a rotation that applies to a role or department picks newcomers up as soon as Contracts gives them that role, with no edit to the rule.

## Consequences
- A member who returns after a gap keeps their old first contract start, so they are not delayed again.
- Re-dealt turns change hands without notice; the people concerned see it in Tasks.
