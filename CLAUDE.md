@AGENTS.md

## Project facts

- **WorkspacePM**: a Google Apps Script **Library** (V8, `src/`) consumed by team **Consumer Sheets** through the thin `src/ConsumerWrapper.js`. Deployment: `docs/distribution-guide.md`.
- No build, no `package.json`. Each `src/*.js` is a global script in Apps Script and ends with a `if (typeof module !== 'undefined') module.exports = {...}` guard so Node tests can `require` it. Keep that guard when adding files or exports.
- Tests: `node --test test/*.js`, all GAS services mocked in `test/mockGas.js` (add new service mocks there). Baseline 2026-10-11: **127 pass, 0 fail**.
- Spec and tickets: `.scratch/workspace-pm-library/` (`spec.md`, `issues/NN-*.md`). Status words in use: `resolved`/`closed` = done, `open`, `ready-for-agent`.
- Every new decision lands as a `docs/adr/NNNN-*.md` (next number: 0033); new domain terms land in `GLOSSARY.md`.
- All sync writes run under the **Sync Lock** (ADR 0032): new write paths triggered by edits or schedules go through `runWithEditLock` / `acquireSyncLock` in `src/Code.js`.
- Apps Script projects (owner `hugoglvs42@gmail.com`): Library `1vyJiEQTSluHaPbh8-B9gSOj2SeciqPEzHpJ5Xug0I_NVhDvtYHdZ1b6v`, Consumer Sheet "[NEW] ACN" `1FOVefnly_zusD1TobiPAYrUEeAR7pLSAhIPxaXMFme0` with bound script `1IVORKENsN2r81QcA5hXa5qa-tgXub2RFFlhzD0-8fopojfAoQFixFujH`. Claude's sandboxes cannot reach Google APIs, so `clasp` runs only in the user's macOS Terminal (`scripts/clasp-setup.sh` clones the live code into the git-ignored `.scratch/live/`).

## Reporting to the user

The final message of every change lists each modified file as a direct link (`[path](file:///Users/hugogoncalves/Documents/antigravity/jolly-salk/<path>)`), with the changes applied to it as sub-bullets.
