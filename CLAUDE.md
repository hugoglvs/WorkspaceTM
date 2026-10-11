@AGENTS.md

## Project facts

- **WorkspacePM**: a Google Apps Script **Library** (V8, `src/`) consumed by team **Consumer Sheets** through the thin `src/ConsumerWrapper.js`. Deployment: `docs/distribution-guide.md`.
- No build, no `package.json`. Each `src/*.js` is a global script in Apps Script and ends with a `if (typeof module !== 'undefined') module.exports = {...}` guard so Node tests can `require` it. Keep that guard when adding files or exports.
- Tests: `node --test test/*.js`, all GAS services mocked in `test/mockGas.js` (add new service mocks there). Baseline 2026-10-11: **127 pass, 0 fail**.
- Spec and tickets: `.scratch/workspace-pm-library/` (`spec.md`, `issues/NN-*.md`). Status words in use: `resolved`/`closed` = done, `open`, `ready-for-agent`.
- Next open work: issue 13 (nightly sync concurrency; design only, user asked for no code yet when filed).
- Every new decision lands as a `docs/adr/NNNN-*.md` (next number: 0032); new domain terms land in `GLOSSARY.md`.
