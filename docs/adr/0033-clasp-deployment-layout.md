# ADR 0033: One Folder per Apps Script Project, Deployed with clasp

## Status
Accepted

## Context
1. Until now, code was copied by hand into the Apps Script editor. The repo and the live projects can drift apart without anyone noticing. On 2026-10-11 the live Library had fallen one change set behind, and the live Consumer Sheet script was an older wrapper.
2. `src/` held both the Library and `ConsumerWrapper.js`. The wrapper defines the same globals as `Code.js` (`onOpen`, `onEdit`, `nightlySyncJob`, `installedTaskEditHandler`…), so pushing `src/` as-is would have put two competing definitions into the Library.
3. clasp pushes every script file under a project's root directory, and the Library and the Consumer Sheet script are separate Apps Script projects.

## Decision
1. `src/` = the Library only, with `src/.clasp.json` pointing at the Library script ID.
2. `consumer/` = the bound script of the reference Consumer Sheet: `consumer/Code.js` (formerly `src/ConsumerWrapper.js`), `consumer/appsscript.json` (library dependency + advanced services) and `consumer/.clasp.json`.
3. Deploy with `clasp push` from each folder after `clasp status` and a green test run. The live code can be pulled for comparison into the git-ignored `.scratch/live/` (`scripts/clasp-setup.sh`).
4. clasp credentials (`.clasprc.json`) never enter the repo (git-ignored).

## Consequences
- The repo is the single source of truth. A push replaces the live project's files with the folder's contents, so any edit made directly in the online editor is lost unless it is pulled first.
- The reference Consumer Sheet uses the Library in development mode (HEAD), so every Library push is live there immediately. Other teams should pin a numbered version (see `docs/distribution-guide.md`).
