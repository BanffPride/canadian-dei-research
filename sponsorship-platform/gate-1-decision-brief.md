# Gate 1 Decision Brief: Sponsorship and Partnership Tracking Platform

**Prepared for:** Jeffrey Carpenter, Executive Director, Banff Pride Society
**Prepared by:** Claude Code (systems architect and implementation agent)
**Date:** 5 September 2026 (27 days before the Banff Pride Festival, 2 to 12 October 2026)
**Scope:** Read-only discovery and current-state audit. No live records, Drive folders, tasks, emails, calendar entries or website content were created or changed.
**Companion file:** `source-index.md` (every record reviewed, with Drive, Gmail, Granola and Todoist identifiers)

---

## 1. Executive summary

**Recommended decision:** Build one new, normalized Google Sheets workbook, "Banff Pride Partnerships", seeded from the July 2026 prospect sheet and the 2026 agreements, with Drive as the document store and Gmail, Granola, Todoist and Calendar linked by ID. Retire the four overlapping trackers after migration. Ship a deliberately small operating core before 2 October, and defer pipeline analytics, renewals and most automation to after the festival.

**What the audit found.**

- Banff Pride already has good design work but no operating system. A well-designed pipeline workbook and in-kind log were built in January 2026 and never used (4 rows, 0 rows). A July 2026 prospect research sheet is the most current record (106 ranked prospects, 112 venues) but has no owner or next action on any row. The live 2026 sponsorship story is held in agreements, Gmail, Granola notes and Todoist, not in any tracker.
- The 2026 written support base is real and larger than the trackers show. Signed or written-confirmed cash commitments total roughly $28,500, of which only $500 (Q-Tilities) is evidenced as received in Gmail. The BLLT agreement ($21,000, three invoices) has Invoice #2 on hold pending a reply that was due 4 September. ATB Financial approved $5,000 plus $5,000 in "Fab Fiver Bucks" on 28 August; Drive still records ATB as "application pending".
- In-kind value is systematically invisible. The 2025 Festival Report counts 78 in-kind sponsors; both 2026 budget files carry in-kind sponsorship at $0. Agreed in-kind values found in writing for 2026 total roughly $12,000 before any hotel, beverage or media discounts are valued.
- Fulfilment is not tracked at all. No record exists of which benefits have been delivered to whom, and no partner logos are stored in a findable location. The 2025 reports went to BLLT, BHC, Canadian Heritage and ATB; there is no evidence that other 2025 sponsors received a report.
- Eleven time-sensitive exceptions surfaced during the audit (section 4.4), several due this week.

**Operational consequence of approving:** one workbook to open, one clear next action per relationship, and a defensible fulfilment and evidence record in time for the BLLT Final Report (due 12 November) and for personalized partner reports after the festival.

**Approval requested (section 13):** five decisions, with the recommended option stated first in each case.

---

## 2. Sources reviewed

Connectors available and used: Google Drive, Gmail, Google Calendar, Granola, Todoist, Jotform (all connected, read access confirmed). Also connected but not needed for Gate 1: Slack, Make, Zapier, Canva, Square, Mailchimp. **banffpride.ca could not be reached** from this environment (the network egress proxy blocks the domain), so public recognition is unverified (section 15).

| Source | Coverage | Method |
|---|---|---|
| Google Drive | 5 title and full-text searches (sponsor, partner, prospect, agreement, MOU, proposal, invoice, logo, brand), 12 folder listings, 22 files read in full, 3 large files parsed locally | Read-only |
| Gmail | Label `01 - BPS/FUND/Sponsorship` plus targeted searches for 20 named partners, hotel, beverage and venue threads, the 3 September mass outreach (91 emails), and 2025 reporting | Read-only, delegated sweep |
| Granola | All 11 meetings available (18 August to 3 September 2026); nothing earlier exists in Granola | Read-only |
| Drive meeting notes | 14 "Partner with Pride" Gemini notes and transcripts, March to August 2026 | Read-only |
| Todoist | Projects Festival 2026 (48 tasks), Funding (17), Marketing & Comms (21), Inbox | Read-only |
| Google Calendar | 5 September to 15 October 2026, plus sponsor keyword search | Read-only |
| Jotform | Taste The Rainbow, Lavender Lounge and Pride Market vendor forms | Read-only (submission counts not returned by the connector) |
| Repository | `banffpride/canadian-dei-research` contains DEI market research only; no sponsorship code or data | Read |

Full inventory with identifiers: `source-index.md`.

---

## 3. Existing assets worth preserving

| Asset | Why it matters | Keep as |
|---|---|---|
| **July 2026 prospect sheet** (`1_54el2fUuyK3QmF5xNlY3rtJDfij2pUWdQyUy3Fm58s`) | Most current record. Fields already match the brief: priority, research score, sponsorship likelihood, sector, sub-category, location, ownership group, evidence basis, approach level, best-fit ask, status, owner, last contact, next action. Ownership-group and registry tables are genuinely useful. | Seed for Organizations and Opportunities. Keep the methodology tab as documentation. |
| **January 2026 Sponsorship Pipeline design** (`1TFqKOBBKFyuyXt1LmKAFl51YLcp9xis2xRQjlt3484E`) | Stable Sponsor ID scheme (`2026-BNF-SKIBIG3-836B96`), controlled lists, dashboard, task tab, calendar export. Unused since 26 January. | Reuse the ID scheme and list values. Do not migrate the 4 stale rows. |
| **January 2026 In-Kind log schema** (`1NP0FsWVMNx0G2Rj-i4N1r3ltz4TPqGRCm4Az2oRayNM`) | Entry ID, FMV, valuation basis, evidence type and link, booked-in-accounting flag, GL accounts, duplicate and missing-evidence flags. This is exactly the contribution record the brief asks for. | Adopt the schema for the Contributions tab. |
| **2026 agreement template** (`1GWbtbsazN_dfdHtv03qLE8Vce7iLM1-g_BrW8QAbE5g`) | Defines tiers (Trailhead $500, Silver $1,500, Alpine $2,500, Summit $5,000, Presenting $10,000), the benefits checklist, asset deadline, 30-day payment term and term end 31 December. The benefits checklist is the natural Deliverables template. | Source of the default deliverable set per tier. Fix the "October 02-122" typo. |
| **Signed 2026 agreements** (BLLT, SkiBig3 x2, Roam) | Authority level 1 records. BLLT contains the only formal logo-approval lead times (10 business days major, 3 routine, 5 web, deemed approval if silent) and the Final Report specification (Appendix B). | Link from Agreements tab; never edit. |
| **Festival 2026 Drive structure** (`BP — Festival 2026/04 Sponsorship Delivery/`) | Already has "01 Sponsor List & Benefits Tracker", "02 Assets by Sponsor", "03 Post-Festival Sponsor Reporting", "04 Donations". Empty or nearly empty, but the intent matches the platform. | Use as the document home. Add the Brand Asset Library alongside it (section 6). |
| **Outreach email templates** (`1fvXHkap_sVT8L23xO5V1dTpRmfRwS2ND1L6NKL4p6Uc`) | Four audience-specific templates plus Lavender Lounge copy. | Communication templates (changes require approval). |
| **Todoist conventions** | Tasks already carry a "Source: Drive/Gmail/Granola, date" line and hard deadlines. Sections NOW / NEXT / EVENT-FINAL and Funding / Sponsors & Partners exist. | Keep. The platform should write tasks in the same format. |
| **Gmail label** `01 - BPS/FUND/Sponsorship` | 88 threads. An inbound filter appears to apply it on replies. | Incremental index key for commitment detection. |
| **2026 Brand Book** (`1s8SDmlNV2JHESzo6qTPmj-WPGqMBXiqn`) | Co-branding rules: Banff Pride leads the group, stacked logo for co-branding, divider lines, clear space, minimum sizes, approved colour pairings only. | Binding for how partner logos appear next to Banff Pride's. |
| **2025 Festival Report** (`1wO_rVuuKiKR_1pGABz-w5L5H1rzj-WYe`) | Metrics set already accepted by BLLT and ATB (attendance 6,500, NPS 88, 79% visitors, reach figures, local spend). | Template for the Reporting Metrics tab. |
| **2025 Community Engagement Partnerships sheet** (`1UDGSTlnZ9BMwdIm4HbmHJrcxLdZGd8UZClfjQSm8Gcg`) | The richest 2025 operational record: 75 partners with status, deliverables from company, deliverables from Banff Pride, logo received (37 of 75), promo codes, and in-kind values (Wild Life Distillery, Moxy, Nourish, Arctos & Bird, Patagonia, Stoney Nakoda and 20 gift-card donors). | Relationship history for 2025. Strip personal contact columns before migration. |
| **2025 partner roster and values** (`1cIs0hz4rKt59LYaUegS2q_K5t2EUvLS-VABAjiTQ6LA`, `1iuP-9-CXFJEsbfVyuuS88yduNiUjD9a5YIUJ4nh6LA4`) | Definitive categorized 2025 roster (129 partners: 14 fiscal, 79 in-kind, 22 venues, 42 Taste The Rainbow, 10 Lavender Lounge, 31 non-profits) and the only file with 2025 cash and in-kind by partner (confirmed $35,000 cash, $9,862.27 in-kind). | Seed for renewal conversations and 2025 history fields. |
| **Tracker: Grants & Sponsors** (`1VbN4W9RsnWbEFZb43Uuh052rNzv1yikla2M-ahp6NrQ`) | Only source of the 2025 corporate ask-and-decline history with amounts (16 declines, 7 still "submitted"). | Import as closed opportunities so 2027 outreach does not repeat 2025 asks blind. |

---

## 4. Current-state problems

### 4.1 Duplication

- Six prospect or pipeline trackers overlap across two years: July 2026 prospect sheet (106 rows), business census tracker (149 rows, 28 July, superseded), January 2026 pipeline (4 rows), Sponsorship Intelligence (March 2026, 10 scored prospects, four already confirmed 2025 partners scored as cold), Sponsors - 2025 (427 rows, 395 still "Prospect", never updated after asks went out) and the BVPN readiness list (97 rows, different purpose but same organizations).
- Two in-kind trackers: the January log (empty, rich schema) and the August "Dontation - In-Kind Tracker" (3 rows, thin schema).
- Three sponsor-folder conventions: `2026-BNF-*` folders with 1,024-byte placeholder documents, the "01 Sponsor List & Benefits Tracker" folder where agreements actually live, and a separate SkiBig3 sub-folder.
- Two budget files disagree on 2026 sponsorship: the July 26 festival budget lists ATB, SkiBig3, Q-Tilities and Highline; the organizational 2026 Budget model carries only BLLT and BHC ($23,000) with in-kind at $0.

### 4.2 Conflicts found between sources

| Topic | Source A | Source B | Resolution needed |
|---|---|---|---|
| ATB Financial status | July prospect sheet: "Confirmed"; Todoist: "application pending" | Gmail 28 Aug: approved $5,000 sponsorship plus $5,000 Fab Fiver Bucks | Record as confirmed on the written approval. Decide whether the $5,000 Bucks is cash or activation in-kind. |
| Banff Lodging Company | Event Plan to BLLT (Aug): accommodation partner with promo code | July sheet: "Declined"; Gmail: declined sponsorship, offers 20% code and performer room block; BLLT notes: "not a lodging partner" | Two unanswered follow-ups from BLC about the room block. Classify as accommodation partner (promo code) with no sponsorship. |
| BHC cash and in-kind | January pipeline ask $2,000 + $2,000 | 2026 agreement doc $1,000 + $2,000; budget $1,000 | Agreement governs. Signature status unverified (section 15). |
| Q-Tilities $500 | Gmail 14 Aug: allocate full $1,000 to BVPN restricted fund | Agreement and invoice: $500 festival sponsorship | ED decision on fund allocation; affects the financial tab. |
| 4 October Drag Brunch | Moxy Banff finalizing 4 Oct brunch (Calendar 7 Sep) | BHC asked on 4 Aug to move the Dusty Boot drag brunch to 4 Oct, no reply seen | Confirm one venue; conflicting recognition risk. |
| SkiBig3 tier | Agreement ticks Silver ($1,500) | Cash $500 + in-kind $1,191 = $1,691 | Treat tier as Silver by total value; note cash is below tier price. |
| Roam contra value | Contract $1,339.50 incl. GST | Panels $990 + passes $300 ex-GST | Record $1,290 ex-GST as in-kind, $49.50 GST separately, flag for review. |

### 4.3 Missing and stale

- **Owner and next action are empty on all 106 ranked prospects and all 149 census rows.** Last contact is populated on 48 rows, all dated 4 or 6 August.
- **No fulfilment record.** No sheet tracks which benefits (website logo, social, signage, VIP tickets, activations) have been delivered.
- **No partner logo library.** "02 Assets by Sponsor" contains only Banff Pride's own logos. Highline, Moraine Lake Bus and SkiBig3 have sent logos by email; they are not filed.
- **No received-cash record** outside Wave. Gmail evidences one e-transfer ($500). BLLT Invoice #1 ($9,500) payment status is not visible.
- **No written 2026 record for Byhendo, Moxy Banff, Town of Banff or the Whyte Museum sponsorship element** (Whyte is a paid rental; BLC and Byhendo activations are agreed verbally or in meeting notes).
- **No complimentary ticket tracking.** SkiBig3 and BHC agreements include VIP tickets; ATB Bucks, influencer VIP passes and volunteer comp codes are handled ad hoc.
- **Reporting metrics tab is empty** (`01 Attendance & Metrics`, 1,024 bytes).
- **2025 partner reports** were sent to BLLT, BHC, Canadian Heritage and ATB only. The "03 Post-Festival Sponsor Reporting" folder is empty, and the 2026 marketing task list schedules a "sponsor cut" of the impact report only for 20 October to 7 November.
- Stale scaffolding: `2026-BNF-*` placeholder documents (empty since 26 January), "Logos - Local Partners" (2021, BLLT only).
- **Misfiled brand guidance:** both "Brand guidelines 2021.pdf" copies in the Banff Pride brand folders are the SMARTstart corporate guidelines, not Banff Pride's. The April 2026 brand critique confirms there is no guidance on partner logo lock-ups, size ratios or co-branded templates beyond the Brand Book's general co-branding page.
- **Three tier ladders, none reconciled:** the March 2025 package draft (Presenting $20,000+ down to Community $500), the March 2026 Sponsorship Intelligence proposal (Presenting $25,000+, Champion, Supporter, Community, with sector exclusivity) and the August 2026 agreement template (Trailhead $500 to Presenting $10,000). Only the agreement template has been used in signed documents. No 2025 partner was contracted at a named tier.
- **2025 contractual reporting not evidenced:** the Banff Lodging Co agreement required a final report by 31 December 2025 and Byhendo's required a thank-you in festival reporting; neither is found in Drive or Gmail. The 2025 Community Engagement sheet shows only 37 of 75 partners with logo received.
- **Known 2025 data-quality defects** the migration must not carry forward: ATB 2025 recorded as both "Declined $10,000" and "Confirmed $5,000"; Banff Lodging Co in-kind $11,500 planned vs $5,000 confirmed; "twelve (6) complimentary VIP tickets" in the BLC agreement; Fairmont's $2,000 note copied from the Encore row; date typos; Pursuit and Pauw Foundation $5,000 each in the plan tab but absent from the confirmed tab.

### 4.4 Exceptions that need action this week (found during audit, regardless of platform)

| Due | Item | Evidence |
|---|---|---|
| Overdue (4 Sep) | Reply to BLLT (Andrew Hercus) with finalized event and marketing plan so Invoice #2 ($9,500) is processed | Todoist 6hQc4pG7vpgF99Hg; Gmail 1a045111156753db |
| 8 Sep noon | Roam Transit bus panel artwork (30" x 11" PDF, quarter-inch bleed) | Gmail 19ff319e1d2378e9; Todoist 6hH2P8jrJJxW87mh (due 8 Sep, hard 11 Sep: the Gmail date is earlier) |
| Unanswered twice | Banff Lodging Co performer room block decision (chased 24 and 31 Aug) | Gmail 1a0348879df68f80 |
| Open | Confirm which venue holds the 4 October drag brunch (BHC Dusty Boot vs Moxy) | Gmail 19fd94025ee54917; Calendar 7 Sep |
| Open | Send Bell Media the free CTV "Our Community" assets (1920x1080 graphic, images, 10-second script) | Gmail 1a0349037701736b |
| Open | Bow Valley Insider needs ticket link and full Taste The Rainbow list, and approval of feature 2 | Gmail 19f861bcc8e81092 |
| Passed 4 Sep | Forward Biosphere Institute exhibit overview to Pursuit / Banff Gondola | Gmail 19e6f5aa8875094f |
| Open | ATB logo files not yet received; add ATB to website and Pride Market vendor list | Gmail 19ff81511fbf37d8 |
| ~15 Sep | Highline countersigned agreement and payment; post logo | Gmail 1a03e9d7f13f4a52 |
| Open | Verify SkiBig3 $500 payment received (Invoice #0813, 11 Aug) | Gmail 19fcddbc33e8d629 |
| Open | Obtain written 2026 terms from Byhendo (recap email promised 19 Aug) | Granola aa476d59 |
| Open | Pick up 50 Roam regional passes at Banff Visitor Centre | Gmail 19ff319e1d2378e9 |

These are surfaced only. No tasks were created.

---

## 5. Current information flow (as observed)

```
Prospect research (Drive sheets, July)  ──►  Mass outreach (Gmail, 3 Sep, 91 emails)
        │                                            │
        │ no owner / next action                     │ replies labelled 01-BPS/FUND/Sponsorship
        ▼                                            ▼
Partner chats (Calendar booking page ──► Gemini notes in Drive ──► Granola from 18 Aug)
        │
        │ commitments noted in meeting summaries only
        ▼
Agreement drafted from template (Drive 01 Sponsor List & Benefits Tracker) ──► PDF signed by email
        │
        ├──► Invoice created in Wave, sent by Gmail (BPS Invoice #08xx)      [payment status: Wave only]
        ├──► Logos arrive by email                                            [not filed]
        └──► Some actions become Todoist tasks with "Source:" line            [manual, partial]

Budget sheets (two, disagreeing)  ◄── manual entry, "Confirmed" typed as text
Website sponsor listing            ◄── manual, "confirmed 2026 sponsors only" rule (stated in Granola)
Post-festival report               ◄── assembled by hand in December from memory and Gmail
```

The gaps are between boxes: nothing carries a commitment from an email or meeting into an agreement, a deliverable, a task or a budget line, and nothing records delivery.

---

## 6. Recommended platform architecture

**Recommended:** one new workbook, "Banff Pride Partnerships", on the BP shared drive under `BP — Festival 2026/04 Sponsorship Delivery/`, year-independent (organizations persist; opportunities and agreements are per year). Deterministic IDs and validations by formula and Apps Script. Drive holds documents; the workbook holds links.

**Alternative considered:** upgrade the July 2026 prospect sheet in place. Rejected because it is a flat research table with 23 columns per prospect and no room for agreements, deliverables or evidence without becoming the "one enormous flat spreadsheet" the brief rules out. Its data migrates in full; its structure does not.

**Source of truth by data type**

| Data type | Authoritative source | Notes |
|---|---|---|
| Organization identity, public name, relationship type | Partnerships workbook: Organizations | Seeded from July sheet and agreements. Legal name from agreement where one exists. |
| Prospect research and scoring | Partnerships workbook: Opportunities (migrated from July sheet) | Method register preserved as a linked doc. |
| Commitments, tiers, benefits, term | Signed agreement PDF in Drive | Workbook holds a link and an extracted summary, never a substitute. |
| Cash invoiced and received | Wave | Workbook records invoice number, date sent, expected date and a "received (per Wave)" date entered by the ED. Not a ledger. |
| In-kind value | Partnerships workbook: Contributions | Every row needs valuation basis and evidence link; missing basis triggers a review flag. |
| Deliverables and fulfilment | Partnerships workbook: Deliverables and Evidence | Evidence is a Drive link (photo, screenshot, post URL). |
| Meetings and decisions | Granola (from 18 Aug) and Gemini notes in Drive (before) | Communications tab stores ID, summary and extracted commitments only. |
| Correspondence | Gmail | Communications tab stores thread ID. |
| Tasks | Todoist | Workbook stores task ID; Todoist stores the workbook row ID in the description. |
| Deadlines | Google Calendar for meetings and hard external dates; Todoist for work | Avoid double entry: calendar only for dates a partner also sees. |
| Brand assets | Drive Brand Asset Library | Workbook stores links and approval status. |
| Public recognition | banffpride.ca | Verified by the existing biweekly website audit routine once it can reach the site (section 15). |

**Brand Asset Library (proposed, no changes made):** `BP — Festival 2026/04 Sponsorship Delivery/02 Assets by Sponsor/` renamed to `Brand Asset Library` and moved one level up to `Partnerships/` so it outlives the 2026 festival, with the structure in the brief (`[PARTNER-ID] Partner Name/01 Current Approved Logos/{Primary, Reversed or White, Monochrome, Vector}/02 Brand Guidelines/03 Approval Records/99 Expired - Do Not Use`). Festival and campaign folders receive Drive shortcuts only. Banff Pride's own 16 logo files stay in `01 Brand & Assets`. Binding partner rules found so far: BLLT approval lead times (agreement section 2), Moraine Lake Bus request not to display the discount link publicly, Brand Book co-branding rules. No other partner supplied brand guidelines in writing.

---

## 7. Proposed workbook tabs

Record tabs (one row per record, stable ID, frozen header, dropdowns):

| # | Tab | ID prefix | Purpose |
|---|---|---|---|
| 1 | Organizations | ORG- | One row per organization, all relationship types as multi-select |
| 2 | Contacts | CON- | Restricted-access tab; name, role, org, preferred channel |
| 3 | Opportunities | OPP- | One row per organization per cycle year; pipeline stage and research fields |
| 4 | Agreements | AGR- | One row per signed or pending agreement; link to PDF |
| 5 | Contributions | CTB- | Cash and in-kind lines with proposed, confirmed, received, used |
| 6 | Deliverables | DLV- | Every benefit owed either way, one row each |
| 7 | Evidence | EVD- | Links proving delivery |
| 8 | Communications | COM- | Material emails and meetings only |
| 9 | Brand Assets | BRA- | Logo and guideline records with approval status |
| 10 | Tickets & Hospitality | TIX- | Comp tickets, VIP passes, promo codes, room blocks |
| 11 | Events & Activations | EVT- | Festival events and partner activations |
| 12 | Metrics | MET- | Reporting metrics with source and attribution scope |
| 13 | Partner Reports | RPT- | Generated report records and send status |
| 14 | Renewals | REN- | Renewal date, brief link, decision |
| 15 | Decisions | DEC- | Approvals requested and given, with date |
| 16 | Improvements | IMP- | Improvement Register |
| 17 | Audit Log | AUD- | Append-only script log of automated actions |
| 18 | Lists | | Controlled values, owners, tiers, config |
| 19 | Source Index | SRC- | Every indexed file, thread, meeting with modified date or hash |

View tabs (formula-driven, no manual entry): ED Dashboard, Decisions Required, Directory, Prospect Pipeline, Confirmed Support, Deliverables Due, Deliverables at Risk, Missing Evidence, Communications Requiring Follow-up, Brand Assets view, Event & Activation view, Reporting Readiness, Renewal Pipeline, Data Quality & Exceptions.

**MVP subset before 2 October:** tabs 1 to 8, 10, 11, 15, 17, 18, 19 plus the ED Dashboard, Decisions Required, Deliverables Due, Missing Evidence and Data Quality views. Tabs 9, 12, 13, 14, 16 and their views are built in the same workbook but populated after the festival (section 14).

---

## 8. Proposed field groups

- **Organizations:** Partner ID, legal name, approved public name, relationship types (multi), sector, sub-category, location, ownership group, approach level, website, social handles, relationship owner, prospect source, notes, created, updated.
- **Opportunities:** Opportunity ID, Partner ID, cycle year, stage, priority, research score, likelihood, evidence basis, alignment notes, best-fit ask, proposed cash, proposed in-kind, expected total, probability, weighted value, last meaningful contact, next action, next-action owner, next-action date, relationship health, risks, source links.
- **Agreements:** Agreement ID, Partner ID, type (sponsorship, partnership, venue, vendor contra, MOU), tier, term start, term end, status, signed by Banff Pride date, signed by partner date, document link, asset deadline, payment terms, exclusivity notes, approval lead time, renewal option.
- **Contributions:** Contribution ID, Partner ID, Agreement ID, kind (cash or in-kind), category, description, proposed, confirmed, received, used, outstanding (formula), valuation basis, evidence link, invoice number, invoice sent, due, received date (per Wave), GST treatment, review flag.
- **Deliverables:** Deliverable ID, Partner ID, Agreement ID, owed by (Banff Pride or partner), exact source language, plain-language description, channel, event or activation, quantity, internal owner, due date, dependency, approval required (yes or no, lead time), status, delivered date, method, evidence required, evidence ID, verification status, substitute value, exception notes.
- **Evidence:** Evidence ID, Deliverable ID, type (photo, screenshot, URL, document, email), link, captured by, captured date, verified by, verified date.
- **Communications:** Communication ID, Partner ID, contact, date, channel, subject, summary, commitments by Banff Pride, commitments by partner, decisions, open questions, follow-up, owner, due, source link (thread ID or meeting ID), proposed record links, triage status.
- **Brand Assets:** Asset ID, Partner ID, approved public name, library folder link, primary logo link, alternates, format, dimensions, received date, source contact, approval status, restrictions, approval lead time, agreement expiry, last verified.
- **Tickets & Hospitality:** Ticket ID, Partner ID, Agreement ID, type (VIP, comp, promo code, room block, pass), quantity, event, issued to, issued date, redemption evidence, deadline.
- **Metrics:** Metric ID, name, value, unit, period, scope (organization-wide or partner-attributable), Partner ID if attributable, source link, retrieved date, confidence.
- Every imported fact in every tab carries: source system, source link, source date, date retrieved, confidence (A, B, C), verify flag.

---

## 9. Proposed status taxonomy

**Recommendation: separate fields, not one status.** A single 17-step status would force the ED to pick between "active activation" and "invoice outstanding" for the same partner. Four fields cover the lifecycle without overlap:

1. **Opportunity stage** (Opportunities tab, one per organization per year): Prospect identified, Researched, Outreach prepared, Contacted, Conversation active, Proposal in development, Proposal sent, Negotiation, Verbal commitment, Agreement pending, Confirmed, Closed won (agreement signed), Closed declined, Closed lost, Dormant. "Confirmed" requires a written confirmation link; "Closed won" requires a signed agreement link. A validation rule blocks either without the link.
2. **Agreement status** (Agreements tab): Draft, Sent, Signed by Banff Pride, Signed by partner, Fully executed, Amended, Expired, Cancelled.
3. **Payment status** (Contributions tab, cash rows): Not invoiced, Invoiced, Partially received, Received, Overdue, Written off. In-kind rows use: Promised, Received, Partly used, Used, Unused, Expired.
4. **Fulfilment status** (Deliverables tab, per deliverable): Not started, In progress, Blocked by partner, Blocked internally, Delivered and verified, Partially delivered, Reported without evidence, Not delivered, Replaced with approved equivalent, No longer applicable.

Relationship-level roll-ups (Active activation, Fulfilment, Partner reporting, Renewal) are computed from these, not typed. Relationship health is a separate dropdown (Strong, Stable, At risk, Unknown) set by the ED with a reason.

---

## 10. Proposed integrations

| System | Read (automatic) | Write (approval required) | Mechanism |
|---|---|---|---|
| Google Sheets | Dashboard, validation, ID generation, roll-ups | Record edits | Formulas and Apps Script (deterministic) |
| Google Drive | Agreement and asset indexing by file ID and modified time | Folder creation, shortcuts, file moves | Apps Script or Claude with explicit approval |
| Gmail | Incremental scan of label `01 - BPS/FUND/Sponsorship` since last checkpoint (message ID) | Draft creation only; sending only via the governed workflow | Claude scheduled routine; Communications tab holds triage queue |
| Granola | New meeting summaries by meeting ID | None | Claude routine; commitments proposed, never auto-applied |
| Todoist | Task status by task ID | Create tasks in existing format ("Source:" line, section, deadline) | Claude routine after ED approves the queue |
| Google Calendar | Partner meetings and external deadlines | Create deadline events | Same as Todoist |
| Website | Sponsor listing check | None | Extend the existing "Banff Pride website audit (biweekly)" routine; requires an environment that can reach banffpride.ca |
| Wave | None available | None | ED enters received dates; reconciliation stays manual for 2026 |
| Make or Zapier | Not recommended for 2026 | | Adds a third system to open and a failure surface during the festival; revisit for 2027 |

All automations keyed by source ID (message ID, meeting ID, file ID, task ID) so a re-run cannot create duplicates.

---

## 11. High-impact UX optimizations

1. **ED Dashboard first, exception-only.** Counts and links for: decisions awaiting approval, overdue commitments, deliverables due in 14 days, blocked by partner, missing agreement or payment or asset, communications needing a reply, totals (confirmed cash, confirmed in-kind, received), reporting readiness, top three recommended actions.
2. **One row per relationship in the Directory with one "Next action" cell** and a link to the source record. No scrolling across 40 columns on a phone.
3. **Decisions Required as a queue** with Approve, Reject, Defer dropdowns; the audit log records the choice.
4. **Deliverables Due shows only the next 14 days, grouped by partner**, with the evidence-required flag visible so evidence is captured during the event.
5. **Controlled dropdowns everywhere a status is typed today** ("Confirned" appears twice in the budget).
6. **Conditional formatting limited to three colours:** overdue, due within 7 days, missing evidence.
7. **Mobile view tabs** kept under 8 visible columns; detail tabs hidden from the mobile tab strip.
8. **Links instead of copies:** agreement link, thread link, meeting link, folder link.

---

## 12. Automation opportunities (ranked by value against risk)

Before 2 October (approval-gated, low risk):
1. Exception report every Monday, daily from 21 September to 12 October, drawn from the Deliverables, Contributions and Communications tabs.
2. Deliverable due-date alerts at 14 and 3 days, respecting partner approval lead times (BLLT 10 business days).
3. "Delivered without evidence" detector.
4. Missing owner or missing next action detector on all active opportunities.
5. Gmail and Granola commitment queue: new items proposed into Communications with a suggested record link, awaiting ED approval.
6. Pre-festival thank-you workflow (section 14, fixed dates 28 September to 1 October).

After the festival:
7. Partner report generator (assembled from Agreements, Contributions, Deliverables, Evidence, Metrics).
8. Renewal briefs from December, using the SkiBig3 June 2027 payment and BLLT renewal window as first cases.
9. Website recognition verification once reachable.
10. Stale prospect detection (no contact in 60 days) for the 2027 pipeline.
11. Todoist and Calendar creation from approved queue items.

---

## 13. Critical decisions requiring ED approval

**Decision 1: Source of truth and retirement.** Recommended: adopt the July 2026 prospect sheet, the signed agreements and the Gmail evidence as the seed; retire the January pipeline, January in-kind log, August donation tracker and July census tracker after migration (preserved read-only in Archive). Alternative: keep the July sheet live in parallel. Consequence of the alternative: two places to update during the festival.

**Decision 2: Architecture.** Recommended: new normalized workbook as in sections 6 to 9, on the BP shared drive. Alternative: extend the July sheet. Consequence stated in section 6.

**Decision 3: "Confirmed" rule and hybrid values.** Recommended rules: Confirmed requires written confirmation; Closed won requires a signed agreement; ATB's $5,000 Fab Fiver Bucks recorded as activation in-kind, not cash; Rocky Mountaineer certificate recorded at $0 until a valuation basis other than "up to" is set; Roam recorded ex-GST; Q-Tilities $500 recorded as festival sponsorship until you decide the BVPN allocation. Each is flagged for review in the Contributions tab either way.

**Decision 4: Minimum viable scope before 2 October** as in section 14, with everything else deferred. Consequence: no pipeline analytics or renewal views until November, but every 2026 commitment, deliverable and piece of evidence is recorded before the festival opens.

**Decision 5: Pre-festival email and access.** Approve the recipient rule (partners with a signed agreement or written confirmation, plus in-kind contributors with evidence) and the four fixed dates; approve Claude reading the Sponsorship Gmail label and Granola on a schedule; confirm the Contacts tab is restricted to the ED and named staff.

No build begins until these are answered.

---

## 14. Recommended next build phase (Gate 2, then Gate 3)

**Gate 2 (Architecture and UX), target 8 to 10 September:** workbook schema, list values, dashboard wireframe, approval design and integration keys, presented as a document plus a synthetic prototype outline.

**Gate 3 (Synthetic prototype), target 12 to 15 September:** prototype workbook with synthetic data for 6 relationship types, validation rules, mobile and desktop screenshots, regression tests.

**Gate 4 (Live migration), target by 18 September:** migrate confirmed and in-discussion organizations (roughly 30 rows), 8 agreements, contributions, roughly 60 deliverables derived from agreement benefits, communications seeded from this audit, Brand Asset Library folders for confirmed partners. Human verification queue for the section 4.2 conflicts.

**Gate 5 (Automations), 19 to 27 September:** exception report, due-date alerts, evidence detector, commitment queue, thank-you workflow. Sponsor email: recipient review 28 September, drafts 29 September, approval 30 September, send 1 October 9:00 America/Edmonton, only if approval is recorded.

**Gate 6 (Acceptance), 28 September to 1 October:** data-quality results, permission tests, rollback check, one-page operating guide.

**Deferred to after 12 October:** partner report generator (built during the festival, used from 13 October; BLLT Final Report due 12 November), renewals, Improvement Register reviews, Todoist and Calendar automation, website verification, Wave reconciliation, 2027 pipeline import of the remaining 200 low-priority prospects.

---

## 15. Data and privacy risks

- **Contact PII** is scattered across the July sheet (venue emails), the BVPN readiness list (personal emails), agreement PDFs (personal phone numbers) and at least one Todoist description (a partner's phone number). Recommendation: Contacts tab with restricted access; no PII in the repository; no PII in generated reports.
- **Meeting transcripts** (Gemini notes) contain verbatim speech from partners. Store only summaries and commitment lines in the workbook.
- **BLLT content licence:** anything Banff Pride supplies to BLLT is licensed in perpetuity (agreement section 2). Evidence photos shared with BLLT should be tagged as such in the Evidence tab.
- **Financial data:** the workbook must not become a shadow ledger. Received amounts are entered from Wave with a date, and the tab says so.
- **Credentials:** none were seen or stored. Nothing in this repository contains live personal data; the source index uses organization names and IDs only.
- **Automation risk during the festival:** any write automation that fails between 2 and 12 October costs ED time. This is why writes stay approval-gated and Make or Zapier are deferred.

---

## 16. Not yet verified

1. banffpride.ca sponsor and partner listings (site unreachable from this environment).
2. Whether countersigned copies exist for BHC, Q-Tilities (agent saw "signed agreement returned Aug 28") and Highline outside the Drive copies, which show blank signature lines.
3. BLLT Invoice #1 ($9,500) payment status and SkiBig3 $500 receipt (Wave not connected).
4. Taste The Rainbow Jotform submission count (connector did not return data); Lavender Lounge form shows no submissions.
5. Whether the Town of Banff "Confirmed" status has a 2026 written basis.
6. Byhendo 2026 terms (no written record found).
7. Moxy Banff hosting terms for the 4 October brunch and the BHC Dusty Boot conflict.
8. Whether an inbound Gmail filter applies the Sponsorship label (filters cannot be listed with the available tools).
9. Whether the BP shared drive permissions allow a restricted Contacts tab (permissions not inspected).
10. 2025 sponsor values by partner rely on one sheet (Planned Contributions, confirmed tab: $35,000 cash, $9,862.27 in-kind) that conflicts with the grants tracker on ATB and omits Pursuit and Pauw; Wave is the only way to settle them.
11. Exact ATB invoice amount and whether Fab Fiver Bucks flow through Banff Pride's books.
12. Pursuit / Banff Gondola 2026 decision (no decision since June).

---

**Stop point.** This brief completes Gate 1. No further work will start until the five decisions in section 13 are answered.
