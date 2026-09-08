# Apps Script layer for "Banff Pride Partnerships"

`Code.gs` adds native Google Sheets design and automation to the live workbook. It writes only to the spreadsheet: no email, no Drive moves, no changes to signed agreements.

## Install (once, two minutes)

1. Open the Sheet. Extensions > Apps Script.
2. Delete the contents of the default `Code.gs`, paste this file, press Save.
3. In the function dropdown pick `setupWorkbook`, press Run, and approve the authorization prompt (it asks for spreadsheet access only).
4. Back in the Sheet, reload the browser tab. A **Banff Pride** menu appears.

`setupWorkbook` is idempotent: run it again after any structural change and it re-applies everything.

## What setup applies

- Fixes the v1.0 dashboard cell "Total support received" (it referenced itself).
- Tab colours, styled header rows, banded rows, filters, frozen headers and key columns, date and currency formats on every record tab.
- Dropdowns with help text on every status, owner, confidence and Partner ID column. Status columns reject values outside the Lists tab.
- Conditional formatting: overdue red, due-soon and missing-evidence amber, blocked purple, closed and inactive rows grey, confirmed and verified green. Applied on record tabs, view tabs and the dashboard.
- Helper formula columns grouped and collapsed on Opportunities and Deliverables; warning-only protection on view tabs, Lists, AuditLog, Contacts and all computed columns.
- Named ranges WarningDays, StaleDays, FestivalStart, FestivalEnd (Lists!B31:B34).
- ED Dashboard: each exception count links to its tab; quick links to the Partnerships folder, Brand Asset Library, archive, delivery folder and operating guide; a partner lookup panel (pick a partner in H11).
- BrandAssets column D filled with each partner's Brand Asset Library folder link.
- Header notes explaining the rules on the columns where people most often go wrong.

## Automation (runs on every edit, no install needed)

- New row on any record tab gets the next ID (ORG-, OPP-, DLV-, ...) as soon as you type in it. Decisions rows get today's date and "Awaiting approval"; Deliverables start "Not started"; Opportunities start "Prospect identified"; Organizations get Created.
- Typing a Partner ID fills the partner name from Organizations.
- Deliverables: "Delivered and verified" is refused when Evidence required is Yes and no Evidence ID is present (the status reverts and a message explains). Delivered date is stamped when the status is accepted.
- Decisions: choosing Approved, Rejected or Deferred stamps Decided date and Decided by.
- ThankYouEmail: Final approval cannot be set to Approved until Include? is Approved; the approver and date are stamped. Nothing is sent by the sheet.
- Organizations and Contacts get an Updated stamp.
- Every edit on a record tab is appended to AuditLog (actor, record ID, column, old and new value). Contact values are never written to the log.

## Menu

- **Go to ED Dashboard**.
- **Add a new partner**: prompts for the public name, creates the Organizations and Opportunities rows with IDs and a 7-day next action.
- **Verify workbook**: checks self-check counts, scans every view for formula errors, finds duplicate IDs, logs the result.
- **Log exception snapshot to AuditLog**: records the nine dashboard counts and the money panel as one audit row.
- **Schedule Monday 07:00 snapshot / Remove**: installs or removes a weekly time trigger for the snapshot (America/Edmonton). Nothing is scheduled unless you choose it.

## Not in scope

The script cannot send email, create Todoist tasks or move Drive files. Those remain approval-gated and are done by Claude on instruction.

## Update packs

When Claude reads new emails, meetings or tasks, the resulting changes arrive as a dated file `Updates_YYYY-MM-DD.gs`. Paste it as a second file in the same Apps Script project, save, pick its `applyUpdates_...` function and run it once. Each pack is guarded against a second run, cites the Gmail thread for every row, assigns IDs through the same helpers as the menu, copies formula columns for appended rows, and writes one AuditLog row. Review the pack before running it: it is the approval step for any financial values it adds.

| Pack | Covers | Adds or changes |
|---|---|---|
| `Updates_2026-09-08.gs` | 15 Gmail threads, 4 to 8 September | 5 organizations, 2 agreements, 3 contributions, 2 evidence rows, 18 deliverables (3 verified, 2 updated), 4 events, 1 ticket row, 13 opportunity updates, 3 brand asset rows, 3 reports, 3 thank-you rows, 12 communications, 15 source index rows |
