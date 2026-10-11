# 0001 Apps Script Library and Document Properties Migration Architecture

Context: The project management module needs to be deployed and maintained across multiple teams and departments without duplicating code or risk of breaking existing team spreadsheets upon library updates.

Decision: We will implement the core engine as a standalone Google Apps Script Library pinned by immutable semantic version numbers in Consumer Sheets. Each Consumer Sheet will run a thin wrapper script and track its active schema version inside `PropertiesService.getDocumentProperties()`. Schema updates will be managed via an incremental, non-destructive migration engine invoked during `init` / `upgrade`.

Considered Options:
- Standalone standalone copies of scripts in each sheet: Rejected because bug fixes and new features cannot be propagated without manual copy-pasting across all sheets.
- Google Workspace Add-on: Rejected due to heavier deployment review processes, domain-wide marketplace setup, and restricted access to direct sheet layout manipulation.

Consequences:
- Library modifications will never inadvertently break active Consumer Sheets pinned to older versions.
- Schema changes (adding columns or tables) must be backward-compatible and implemented as additive migrations that preserve existing row data.
