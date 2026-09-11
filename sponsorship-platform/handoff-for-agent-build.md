# Banff Pride Partnerships: handoff for building the managing agent

Prepared 11 September 2026 for whoever builds the operating agent (ChatGPT or otherwise). Everything below is the current, live state. It contains no personal contact details and no credentials; those live in the workbook's restricted Contacts tab and in the owner's accounts respectively.

## 1. What exists today

One Google Sheet, "Banff Pride Partnerships" (ID `1FIl8Fu_MFI2FB-y9W8A_vco3empt8TlI0dwQoWTbbBU`), in the BP shared drive under `BP — Festival 2026 / 04 Sponsorship Delivery`. It is the operating record for every sponsor and partner of the Banff Pride Festival (2 to 12 October 2026), from prospect to renewal. It replaced five earlier trackers, now archived.

| Layer | Where | Purpose |
|---|---|---|
| Workbook | Google Sheet above, 36 tabs | Records, views, dashboard, audit log |
| Design and automation | Apps Script bound to the Sheet (`apps-script/Code.gs`) | Formatting, validation, onEdit rules, "Banff Pride" menu |
| Update packs | `apps-script/Updates_YYYY-MM-DD.gs` | Dated, guarded scripts that apply changes from emails and meetings |
| Generator | `generator/build_workbook.py` (Python, openpyxl) | Rebuilds the workbook from seed data; used for v1.0 only |
| Drive | `Partnerships` top-level folder (`1j9BGkxjapQ4DUK75eCm2J9kxgMx1N-Eo`) | Brand Asset Library, archive of superseded trackers |
| Docs | `operating-guide.md`, Google Doc `1O4ApancucSPGKAYAHlYQnXQ2Gn6Y0cVxhvi37sUb6jw` | How the Executive Director uses it |
| Source index | `source-index.md` and the SourceIndex tab | Every source consulted, with IDs |
| Write bridge (not yet working) | Make scenario 4910471 "Claude Sheets bridge" | Relays Sheets API calls; Google connection needs reauthorization |

Owner and only approver: the Executive Director (ED), first name Jeffrey, mailbox info@banffpride.ca (Gmail) and Jeffrey@banffpride.ca (sends outreach). Canadian English throughout. Time zone America/Edmonton.

## 2. Data model

Seventeen record tabs (green), fourteen view tabs (blue, formulas only), one dashboard, one README, three system tabs (grey). Header row is row 1 except Contacts and ThankYouEmail, whose header is row 2 (row 1 is a banner). Data validation lists live on the Lists tab. Rows 2 to 600 are formula-covered on helper columns.

Stable IDs, column A, zero-padded to three digits: ORG, CON, OPP, AGR, CTB, DLV, EVD, COM, BRA, TIX, EVT, MET, RPT, REN, DEC, IMP, SRC. ThankYouEmail is keyed by Partner ID. IDs are never reused or renumbered.

### Record tabs and columns

**Organizations** (154 rows at go-live; 159 after the 8 September pack): Partner ID | Legal name | Approved public name | Relationship types (semicolon-separated) | Sector | Sub-category | Location | Ownership group | Website | Relationship owner | Legacy ID | Prospect source | Notes | Created | Updated | Open deliverables (formula) | Confirmed cash (formula) | Confirmed in-kind (formula).

**Contacts** (restricted; header row 2): Contact ID | Partner ID | Partner | Name | Role | Email | Phone | Preferred channel | Notes | Updated. Never export, never quote in reports, never write contact values to logs.

**Opportunities** (one per organization): Opportunity ID | Partner ID | Partner | Cycle year | Stage | Priority | Research score | Likelihood | Evidence basis | Approach level | Best-fit ask | Rationale / alignment | Proposed cash | Proposed in-kind | Expected total (formula) | Probability | Weighted value (formula) | Last meaningful contact | Next action | Next-action owner | Next-action date | Relationship health | Risks | Source link | Confidence | Verify | (reserved) | Active? (formula) | Missing owner or action (formula) | Stale? (formula).

**Agreements**: Agreement ID | Partner ID | Partner | Type | Tier | Term start | Term end | Status | Signed by Banff Pride | Signed by partner | Document link | Asset deadline | Payment terms | Exclusivity | Approval lead time | Renewal option | Source | Confidence | Notes | Missing document? (formula).

**Contributions** (one line per cash or in-kind item): Contribution ID | Partner ID | Partner | Kind (Cash / In-kind) | Category | Description | Proposed | Confirmed | Received | Used | Valuation basis | Evidence link | Invoice number | Invoice sent | Due | Received date (per Wave) | GST | Review flag | Source | Confidence | Notes | Agreement ID | Outstanding (formula) | Status (formula: Received, Partially received, Overdue, Invoiced, Not invoiced, Used, Promised, Unvalued).

**Deliverables** (one per benefit owed either way): Deliverable ID | Partner ID | Partner | Agreement ID | Owed by (Banff Pride / Partner) | Exact source language | Plain-language deliverable | Channel | Event or activation | Quantity | Internal owner | Due date | Dependencies | Approval required | Approval lead time (days) | Status | Delivered date | How delivered | Evidence required | Evidence ID | Verification | Substitute value | Exception notes | Source | Confidence | Notes | Open? (formula) | (reserved) | Flag (formula: Overdue, Due soon, Blocked, Missing evidence, No due date) | (reserved).

**Evidence**: Evidence ID | Deliverable ID | Type | Link or location | Captured by | Captured date | Verified by | Verified date | Shared with BLLT (perpetual licence)? | Notes.

**Communications** (one per email thread, meeting or decision): Communication ID | Partner ID | Partner | Contact | Date | Channel | Subject or meeting title | Summary | Commitments by Banff Pride | Commitments by partner | Decisions | Open questions | Follow-up action | Owner | Due date | Source link | Proposed records | Triage (Open / Waiting / Closed) | Needs reply? (formula).

**BrandAssets**: Asset ID | Partner ID | Approved public name | Library folder link | Primary logo link | Alternative logo links | File format / dimensions | Restrictions and approval rules | Approval status | Received date | Source contact / thread | Approval lead time | Agreement expiry | Last verified | Notes.

**Tickets** (passes, prizes, hospitality): Ticket ID | Partner ID | Partner | Agreement ID | Type | Description | Quantity | Purpose or event | Issued to | Status | Received or issued date | Evidence or source | Next step | Deadline.

**Events**: Event ID | Event or activation | Date | Host partner ID | Venue | Type | Status | Notes | Linked partner IDs | Deliverables linked (formula) | Evidence to capture.

**Metrics**: Metric ID | Metric | Value | Unit | Period | Scope (organization-wide vs partner-attributable) | Partner ID | Source | Source link | Confidence | Retrieved.

**Reports** (one per confirmed partner): Report ID | Partner ID | Partner | Report type | Due | Status | Deliverables total (f) | Delivered and verified (f) | Missing evidence (f) | Cash confirmed (f) | Cash received (f) | In-kind confirmed (f) | Evidence items (f) | Approved by | Sent date | Link.

**Renewals**: Renewal ID | Partner ID | Partner | Situation | Renewal prep date | Status | Owner | Source | Brief link | 2025 value | 2026 value (formula).

**Decisions**: Decision ID | Date raised | Topic | Decision or request | Status (Awaiting approval / Approved / Rejected / Deferred / Recorded) | Decided by | Source | Decided date | Notes.

**Improvements**: Improvement ID | Date | Type | Evidence of the problem | Root cause | Proposed change | Expected benefit | Risks | Before and after example | Regression test | ED approval | Version | Rollback path | Post-implementation measure | Status.

**ThankYouEmail** (header row 2): Partner ID | Partner | Verified support (formula) | Recognition line (edit) | Recipient contact (from Contacts) | Include? | Draft status | Draft link | Final approval | Approved by | Approval date | Scheduled send | Sent | Notes.

### System tabs

**AuditLog**: Timestamp | Actor | Action | Record IDs | Detail | Source. Append-only. Every onEdit on a record tab writes a row. API writes do not fire onEdit, so any external writer must append its own row.

**SourceIndex**: Source ID | System | Title | ID or link | Source modified | Retrieved | Authority (1 signed agreement, 2 written correspondence, 3 confirmed correspondence, 4 approved internal record, 5 meeting record, 6 working notes) | Note.

**Lists**: dropdown values by column (A Relationship type, B Opportunity stage, C Agreement status, D Deliverable status, E Owed by, F Yes/No, G Priority, H Likelihood, I Relationship health, J Contribution kind, K Contribution category, L Confidence, M Channel, N Triage, O Owner, P Decision status, Q Verification, R Report status, S Renewal status, T Email approval) and config: B31 deadline warning days (14), B32 stale prospect days (60), B33 festival start, B34 festival end, B35 tier prices (Trailhead $500; Silver $1,500; Alpine $2,500; Summit $5,000; Presenting $10,000; Custom). Named ranges WarningDays, StaleDays, FestivalStart, FestivalEnd.

### Views and dashboard

Views (row 4 headers, data from row 5, `IFERROR(SORT(FILTER(...)))` pattern): Decisions Required, Directory, Prospect Pipeline, Confirmed Support, Deliverables Due, Deliverables at Risk, Blocked, Missing Evidence, Follow-up Needed, Brand Assets View, Events View, Reporting Readiness, Renewal Pipeline, Data Quality.

ED Dashboard: nine exception counts in B5:B13 with clickable tab links in C; money panel D5:E13 (confirmed cash, cash received, confirmed in-kind, in-kind received, total received, unvalued rows, proposed cash, reporting readiness, reports not started); ten recommended next actions rows 17 to 26; ten nearest next-action dates rows 30 to 39; self-check B42:B43; quick links G4:G9; partner lookup panel driven by the dropdown in H11 (G12:H24).

## 3. Rules the workbook enforces (the agent must respect them)

1. Four separate status fields, never one: Opportunity stage, Agreement status, Contribution status (computed), Deliverable status.
2. Stage "Confirmed" requires written confirmation on file; "Closed won" requires a signed agreement link. Agreements without a document are flagged.
3. A deliverable is "Delivered and verified" only with an Evidence ID. The onEdit trigger refuses the status otherwise and reverts it. "Reported without evidence" stays flagged.
4. In-kind lines need a valuation basis; unvalued lines are counted, not summed. Promo codes, discounts and free inclusions are in-kind (ED rule, 5 September).
5. Cash figures come from Wave (the accounting system). The workbook is not a ledger.
6. Every active opportunity has a next action, owner and date. Missing ones appear on Data Quality.
7. Nothing is emailed from the workbook. ThankYouEmail records recipients and approvals only.
8. Contacts is restricted. Names of individuals never appear in reports.
9. Metrics are labelled by Scope so organization-wide impact is never presented as partner-attributable.
10. Every write cites its source (Gmail thread ID, Granola meeting ID, Todoist task ID, Drive file ID) and its confidence (A, B, C).

## 4. Safety and approval model (verbatim constraints)

The agent must never:
- Send an email without the ED's explicit consent for that email. ("Never send a email without my consent.")
- Modify signed agreements.
- Infer agreement amendments.
- Expose contact information.
- Commit credentials or live personal data to source control.
- Delete evidence or historical brand assets.
- Claim unverified delivery.
- Agree to new sponsor benefits.
- Reduce approval protections for efficiency.

Approval from the ED is required before: creating Todoist tasks, scheduling communications, changing financial values, moving Drive files, changing the schema or status lists, and any automation that writes on a schedule. Practice so far: the agent proposes, the ED approves in chat or by setting a Decisions row to Approved, then the agent applies.

Decisions still awaiting the ED at handoff: DEC-013 Rocky Mountaineer valuation basis; DEC-014 create 12 exception tasks in Todoist; DEC-015 weekly exception-report routine (daily from 21 September); DEC-016 thank-you email workflow dates and recipient rule.

## 5. Integrations and identifiers

| System | Identifier | Use |
|---|---|---|
| Google Sheet | `1FIl8Fu_MFI2FB-y9W8A_vco3empt8TlI0dwQoWTbbBU` | The record |
| Gmail label | `01 - BPS/FUND/Sponsorship` = `Label_155126162987245264` | Sponsorship correspondence; outreach threads titled "Partner with the 2026 Banff Pride Festival" |
| Gmail sender for outreach | Jeffrey@banffpride.ca; replies land in info@banffpride.ca | Read only |
| Granola | Meetings from 18 August 2026 onward; use list and get, the query endpoint times out | Meeting evidence |
| Google Meet notes | Gemini notes emails from gemini-notes@google.com | Meeting evidence for "Partner with Pride" booked chats |
| Todoist | Festival 2026 project `6fghpGvf67x3cm28`; Funding project `6fghp8mp8JpgmwQ5` | Tasks, approval-gated |
| Drive: Partnerships | `1j9BGkxjapQ4DUK75eCm2J9kxgMx1N-Eo` | Top level |
| Drive: Brand Asset Library | `1eHlsfZqe2kSe_FH4KSevTNoKwXJ_I3xu` | `[ORG-nnn] Partner` folders, each with 01 Current Approved Logos, 02 Brand Guidelines, 03 Approval Records, 99 Expired - Do Not Use; template `1WwepLiH29J1F1OM1MhdZe9eStzjSm3Gf` |
| Drive: Archive | `150KYLyznKKC0maLD4IBu7zw07j7UO-pA` | Superseded trackers (three still to be moved by the ED: January pipeline and in-kind workbooks in `02 Sponsorship` `1jHCBBmFgxifxL1pJqhYMl-QJNrbyV7Ng`, August donation tracker) |
| Drive: 04 Sponsorship Delivery | `1sKYE6lJ4dZU12wIlRBM7jvG1eLzUXCiw` | Workbook and operating guide |
| July 2026 prospect sheet | Stays in place as the historical source of truth for the prospect list | Read only |
| Make | Org 1367958, team 545868, scenario 4910471 | Sheets API relay; Google connection 2913682 needs reauthorization |
| Zapier | Google Sheets connection exists for info@banffpride.ca | Out of tasks on the current plan |
| Website | banffpride.ca; partners page banffpride.ca/partners | Unreachable from the previous build environment; website deliverables need screenshot evidence |

Brand Asset Library folder IDs for ORG-001 to ORG-015, ORG-102, ORG-127 and ORG-155 are in `apps-script/Code.gs` (PARTNER_FOLDERS).

## 6. Recurring workflows the agent should run

**Daily (read-only unless approved).** Read Gmail label and inbox since the last checkpoint (AuditLog holds the last "Update pack applied" or "Exception snapshot" row). For each thread: identify the organization (match on sender domain or name against Organizations), summarize into a Communications row, extract commitments into Deliverables, money into Contributions, documents into Agreements, meetings into Events, and set the Opportunity's last contact, next action, owner and date. Every row cites the thread ID. Present the batch to the ED for approval, then apply.

**Weekly (Monday 07:00 Edmonton).** Exception report from the dashboard: decisions awaiting approval, overdue commitments, due within 14 days, blocked, missing agreements, cash confirmed but not received, brand assets pending, communications needing a reply, deliverables reported without evidence. Append one "Exception snapshot" row to AuditLog and send the summary to the ED in chat. Daily from 21 September to 12 October.

**Evidence capture.** When a deliverable is delivered (post published, logo placed, email sent), add an Evidence row with the link or screenshot location, then set the deliverable to Delivered and verified with the Evidence ID. Never mark delivered without evidence.

**Thank-you email workflow (approval-gated at every step).** 28 September: recipient and exception review on ThankYouEmail (Include? column). 29 September: draft one message per approved recipient into Gmail Drafts. 30 September: ED marks Final approval per row. 1 October 09:00 America/Edmonton: send individually, only rows with Final approval = Approved and Include? = Approved. No approval, no send.

**Partner reports.** One per confirmed partner on the Reports tab. BLLT (Banff & Lake Louise Tourism) Final Report is due 12 November in their Appendix B format; others by 30 November. A report claims only what Deliverables shows as Delivered and verified with evidence, and separates organization-wide metrics from partner-attributable ones.

**Renewals.** After reports go out, prepare renewal briefs per the Renewals tab (prep dates from 1 February 2027 for ATB, similar for others).

**Brand assets.** Chase pending logos (11 pending at handoff), file them in the partner's 01 folder, record the link in BrandAssets. BLLT approval lead times: 10 business days for major assets, 3 routine, 5 web. When the Banff Pride logo appears with partners, the Brand Book rules apply: Banff Pride leads, stacked logo, divider lines, clear space.

## 7. Apps Script layer (what already runs inside the Sheet)

- **Menu "Banff Pride"**: Go to ED Dashboard; Add a new partner (creates Organizations and Opportunities rows with IDs); Verify workbook (self-check counts, formula errors, duplicate IDs); Log exception snapshot; Apply design and automation; Schedule or remove the Monday snapshot trigger.
- **onEdit**: assigns the next ID to a new row; fills partner name from Partner ID; refuses "Delivered and verified" without an Evidence ID; stamps Delivered date, Decided date and decider, thank-you approver and date; stamps Organizations and Contacts Updated; appends an AuditLog row for every edit (contact values excluded).
- **Update packs**: helpers `appendRecord_` (assigns ID, copies formula columns from the row above), `updateById_`, `findRowByValue_`, `findRowByText_`, `upsertOrg_`. Each pack is guarded by a document property so it cannot run twice. `Updates_2026-09-08.gs` is the worked example: 15 Gmail threads into 5 organizations, 2 agreements, 3 contributions, 2 evidence rows, 18 deliverables, 4 events, 1 ticket row, 13 opportunity updates, 3 brand asset rows, 3 reports, 3 thank-you rows, 12 communications and 15 source index rows.

If the new agent writes through the Sheets API instead of Apps Script, it must replicate three things the onEdit trigger would otherwise do: assign IDs, copy formula columns for appended rows, and append AuditLog rows.

## 8. Recommended architecture for the managing agent

**Access.** Google Sheets API (read and write) and Drive API on the info@banffpride.ca account via OAuth, Gmail read-only scope plus drafts scope (never send scope until the ED grants it for a specific run), Todoist API, Google Calendar read. Store tokens in the agent platform's secret store, never in the Sheet or the repository.

**Tools to expose to the model.**
- `sheet.read(tab, range)` and `sheet.query(tab, filter)`.
- `sheet.append(tab, record)` that assigns the next ID, copies formula columns, validates dropdown values against Lists, and writes an AuditLog row.
- `sheet.update(tab, id, fields)` with the same validation and audit.
- `gmail.search(query, since)` and `gmail.thread(id)`; `gmail.draft(to, subject, body)`; no send tool by default.
- `drive.list(folderId)`, `drive.createFolder`, `drive.move` (approval-gated).
- `todoist.createTask` (approval-gated), `todoist.list`.
- `propose(changeSet)` that renders a human-readable diff for the ED and waits for approval before any write that changes money, status to Delivered and verified, or creates tasks or drafts.

**Behaviour.**
- Start every session by reading the ED Dashboard and AuditLog tail, and by finding the last checkpoint.
- Never invent partner commitments. If a thread is ambiguous, record it as a Communication with the open question and a next action, not as a Deliverable or Contribution.
- Prefer confidence C and a Review flag over a guessed number.
- Keep organization names as the "Approved public name" in Organizations; individuals only in Contacts.
- Log every write with actor "Agent", the record IDs, and the source.
- Speak Canadian English, short sentences, dates as yyyy-mm-dd, money as $#,##0.

**Suggested system prompt core.** "You manage the Banff Pride Partnerships workbook. You read Gmail, Granola, Todoist and Drive; you write only to the workbook and only with sources. You never send email, never edit signed agreements, never mark delivery without evidence, never agree to new benefits, and never expose contacts. Money, task creation, drafts and Drive moves require the ED's approval, which you request with a clear before-and-after summary. When unsure, record an open question and a next action."

## 9. Known limits and open items

- Make bridge: reauthorize "My Google connection" in make.com before it can relay writes. Zapier connection is valid but the account has no tasks.
- Three superseded trackers are owned by another account and still need the ED to drag them into the Archive folder.
- The website could not be reached from the previous build environment; website deliverables are "Reported without evidence" until screenshots are filed.
- Rocky Mountaineer certificate valuation (DEC-013) and ATB Fab Fiver Bucks mechanics remain open with the partners.
- ThankYouEmail rows exist for confirmed partners only; new confirmations must add a row.
- The July prospect sheet imported 99 organizations without owners or next actions; Data Quality shows 100 active opportunities missing an owner or action. Triage them or set the stage to Dormant.

## 10. Where the code and documents are

Repository `BanffPride/canadian-dei-research`, branch `claude/banff-pride-sponsorship-platform-en2bhs`, folder `sponsorship-platform/`:

- `gate-1-decision-brief.md`: the original audit and architecture decisions.
- `source-index.md`: every source with IDs and confidence.
- `operating-guide.md` and the Google Doc: the ED's routine.
- `go-live-2026-09-06.md`: what is live, decisions, update log.
- `generator/build_workbook.py`: workbook generator (seed data with partner staff names is not committed).
- `apps-script/Code.gs`, `apps-script/Updates_2026-09-08.gs`, `apps-script/README.md`.
- This file.
