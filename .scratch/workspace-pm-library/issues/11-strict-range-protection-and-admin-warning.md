# 11: Strict Range Protection and Admin Edit Warning

**What to build:** Enforce strict permission locks on protected columns across operational tables. Configure `setWarningOnly(false)` and restrict range editors exclusively to authorized admins (spreadsheet owner, script executor, and configured admin emails in Document Properties). Non-admin users are strictly blocked from changing protected columns. In `onEdit(e)`, detect manual changes made to protected columns and display warning notifications exclusively to authorized admins.

**Blocked by:** 02

**Status:** resolved

- [x] Configure `protection.setWarningOnly(false)` on all declared protected columns in `src/ProtectionManager.js`.
- [x] Implement `getAuthorizedAdmins(ss)` and `setAuthorizedAdmins(emails)` to manage authorized admin emails.
- [x] Strip non-admin editors via `protection.removeEditors()` and disable domain-level editing.
- [x] Add targeted admin warning toast in `onEdit(e)` in `src/Code.js` when an authorized admin edits a protected column.
- [x] Write automated tests in `test/validation.test.js` and `test/sync.test.js` validating strict protection and admin warning toasts.

## Answer

Implemented in `src/ProtectionManager.js` and `src/Code.js`.
Documented in ADR-0009 (`docs/adr/0009-strict-range-protection-and-admin-warning-audit.md`).
Protected ranges now strictly block unauthorized users (`isWarningOnly: false`), restrict editors to spreadsheet owner and authorized admins, and issue warning toasts exclusively when an authorized admin manually edits a protected column.
Verified with 23 passing tests.
