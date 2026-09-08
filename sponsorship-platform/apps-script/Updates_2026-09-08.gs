/**
 * Update pack: emails read 4 to 8 September 2026 (Gmail Sponsorship label and inbox).
 * Paste as a second file next to Code.gs, then run applyUpdates_20260908 once.
 * The run is guarded: a second run does nothing. Every row cites its Gmail thread.
 * Requires Code.gs (nextId_, appendAudit_, headerRow, actor_).
 */

const UPD_KEY = 'updates_20260908_applied';

function applyUpdates_20260908() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getDocumentProperties();
  if (props.getProperty(UPD_KEY)) { SpreadsheetApp.getUi().alert('Update pack 2026-09-08 has already been applied.'); return; }
  const D = s => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
  const log = [];

  // ---------------------------------------------------------- Organizations
  const org = {
    BIKE: upsertOrg_(ss, 'ORG-155', { 'Legal name': 'Bikescape', 'Approved public name': 'Bikescape', 'Relationship types': 'In-kind sponsor;Attraction or experience sponsor', 'Sector': 'Guided experiences', 'Sub-category': 'Bike tours and coaching', 'Location': 'Banff', 'Website': 'https://www.bikescape.ca', 'Relationship owner': 'Jeffrey', 'Prospect source': 'Sept 3 outreach', 'Notes': 'Offered a guided ride during the festival and a donated experience for prizing (Sept 5). Call booked Sept 10 09:30.' }),
    AIR: upsertOrg_(ss, 'ORG-156', { 'Legal name': 'Alpine Air Adventures', 'Approved public name': 'Alpine Air Adventures', 'Relationship types': 'Attraction or experience sponsor', 'Sector': 'Guided experiences', 'Sub-category': 'Helicopter and outdoor experiences', 'Location': 'Banff', 'Website': 'https://alpineairadventures.com', 'Relationship owner': 'Jeffrey', 'Prospect source': 'Sept 3 outreach', 'Notes': 'Asked what we had in mind (Sept 5). Trailhead $500 or a donated experience proposed Sept 8.' }),
    SUN: upsertOrg_(ss, 'ORG-157', { 'Legal name': 'Sunset Alpine Promotional Products', 'Approved public name': 'Sunset Alpine Promotional Products', 'Relationship types': 'Cash sponsor;Vendor or supplier', 'Sector': 'Promotional products', 'Sub-category': 'Branded merchandise', 'Location': 'Canmore', 'Website': 'https://www.sunsetalpine.com', 'Relationship owner': 'Jeffrey', 'Prospect source': 'BVPN introduction (James, Aug 26)', 'Notes': 'Owner committed $1,000 community sponsorship to Canmore Pride on Sept 8. Banff Pride ask not yet made.' }),
    JASE: upsertOrg_(ss, 'ORG-158', { 'Legal name': 'JASE Security Ltd.', 'Approved public name': 'JASE Security', 'Relationship types': 'Vendor or supplier', 'Sector': 'Security services', 'Sub-category': 'Event security', 'Location': 'Banff', 'Website': '', 'Relationship owner': 'Jeffrey', 'Prospect source': 'Vendor follow-up (2025 supplier)', 'Notes': 'Provided security in 2025. Quote requested Sept 8 for Mason Hall Oct 9, 10, 11 (20:00 to 01:00, up to 120 attendees). Grant-dependent budget.' }),
    FLOW: upsertOrg_(ss, 'ORG-159', { 'Legal name': 'Flow Skate & Friends', 'Approved public name': 'Flow Skate & Friends', 'Relationship types': 'Community partner', 'Sector': 'Arts and recreation', 'Sub-category': 'Skate and dance pop-up; visual art', 'Location': 'Calgary', 'Website': 'https://www.instagram.com/ziastolbie/', 'Relationship owner': 'Jeffrey', 'Prospect source': 'Partner with Pride booking (Sept 8)', 'Notes': 'Proposes an all-ages skate and dance pop-up at Banff Skatepark on Oct 3 and a Black queer joy art pop-up (Banff Centre or gondola) with Black Pride YYC artists.' })
  };
  org.COOP = 'ORG-127'; org.TEA = 'ORG-102'; org.BLLT = 'ORG-001'; org.PUR = 'ORG-021'; org.BIO = 'ORG-047'; org.MOXY = 'ORG-011'; org.BYH = 'ORG-010'; org.NIB = 'ORG-034'; org.BC = 'ORG-032';
  log.push('Organizations: ' + Object.values(org).slice(0, 5).join(', '));

  // ------------------------------------------------------------ Agreements
  const agrCoop = appendRecord_(ss, 'Agreements', { 'Partner ID': org.COOP, 'Partner': "Cooper's Dog Cabin", 'Type': 'Sponsorship agreement', 'Tier': 'Silver', 'Term start': D('2026-09-08'), 'Term end': D('2026-12-31'), 'Status': 'Sent (countersigned copy not filed)', 'Document link': 'Gmail 1a063b3e50267f65 (agreement and BPS Invoice #0908 attached, sent Sept 8)', 'Asset deadline': D('2026-09-14'), 'Payment terms': 'Payable within 30 days of signing (BPS Invoice #0908)', 'Source': 'Gmail 1a063b3e50267f65', 'Confidence': 'A', 'Notes': 'Silver accepted by email Sept 7. Benefits: website logo, social recognition, signage logo, featured social and email spotlight.' });
  const agrTea = appendRecord_(ss, 'Agreements', { 'Partner ID': org.TEA, 'Partner': 'Banff Tea Co.', 'Type': 'Written confirmation (no agreement document)', 'Tier': 'In-kind', 'Term start': D('2026-09-06'), 'Term end': D('2026-12-31'), 'Status': 'Confirmed in writing', 'Document link': 'Gmail 1a063b3d344298d2 (Sept 6 email)', 'Source': 'Gmail 1a063b3d344298d2', 'Confidence': 'A', 'Notes': 'Rainbow Dreams tea blend returns; full proceeds from sales donated to Banff Pride. Accepted by ED Sept 8.' });
  log.push('Agreements: ' + agrCoop + ', ' + agrTea);

  // ---------------------------------------------------------- Contributions
  const ctbCoop = appendRecord_(ss, 'Contributions', { 'Partner ID': org.COOP, 'Partner': "Cooper's Dog Cabin", 'Kind': 'Cash', 'Category': 'Sponsorship fee', 'Description': "Cooper's Dog Cabin Silver sponsorship", 'Proposed': 1500, 'Confirmed': 1500, 'Valuation basis': 'Silver tier ($1,500 per 2026 template); accepted by email Sept 7', 'Evidence link': 'Gmail 1a063b3e50267f65', 'Invoice number': 'BPS Invoice #0908', 'Invoice sent': D('2026-09-08'), 'Due': D('2026-10-08'), 'Review flag': 'Confirm invoice amount matches $1,500', 'Source': 'Gmail 1a063b3e50267f65', 'Confidence': 'B', 'Notes': 'Due date assumes signing on Sept 8; adjust to 30 days after actual signing.', 'Agreement ID': agrCoop });
  const ctbTea = appendRecord_(ss, 'Contributions', { 'Partner ID': org.TEA, 'Partner': 'Banff Tea Co.', 'Kind': 'In-kind', 'Category': 'Merchandise revenue share', 'Description': 'Rainbow Dreams tea blend: full proceeds donated', 'Valuation basis': 'Unvalued until sales are reported; ask for the 2025 proceeds figure as an estimate', 'Evidence link': 'Gmail 1a063b3d344298d2', 'Review flag': 'Request 2025 proceeds figure', 'Source': 'Gmail 1a063b3d344298d2', 'Confidence': 'A', 'Agreement ID': agrTea });
  const ctbBike = appendRecord_(ss, 'Contributions', { 'Partner ID': org.BIKE, 'Partner': 'Bikescape', 'Kind': 'In-kind', 'Category': 'Prizes / auction items', 'Description': 'Donated guided ride experience for prizing (offered Sept 5)', 'Valuation basis': 'Unvalued; value at retail once the experience is confirmed on the Sept 10 call', 'Evidence link': 'Gmail 1a063b74c624da44', 'Review flag': 'Confirm experience and retail value', 'Source': 'Gmail 1a063b74c624da44', 'Confidence': 'B' });
  updateById_(ss, 'Contributions', 'CTB-002', { 'Notes': 'BLLT submitted Invoice #2 for payment on Sept 8 after receiving the event and marketing plan addendums (Gmail 19fd3a82f9a41493).' });
  log.push('Contributions: ' + [ctbCoop, ctbTea, ctbBike].join(', ') + '; CTB-002 note');

  // ------------------------------------------------------------- Evidence
  const evdBllt = appendRecord_(ss, 'Evidence', { 'Deliverable ID': 'DLV-005', 'Type': 'Email', 'Link or location': 'Gmail thread 19fd3a82f9a41493: Sept 8 email with event plan and marketing plan addendums; Andrew Hercus acknowledged and submitted Invoice #2 for payment the same day', 'Captured by': 'Claude', 'Captured date': D('2026-09-08'), 'Verified by': 'Jeffrey', 'Verified date': D('2026-09-08'), 'Shared with BLLT (perpetual licence)?': 'No', 'Notes': 'Also evidences DLV-001 and DLV-002 (plans accepted by BLLT).' });
  const evdPur = appendRecord_(ss, 'Evidence', { 'Deliverable ID': '', 'Type': 'Email', 'Link or location': 'Gmail thread 19e6f5aa8875094f: Sept 7 email to Pursuit with the Queer(y)ing Ecologies exhibition overview, photos and video walkthrough', 'Captured by': 'Claude', 'Captured date': D('2026-09-08'), 'Verified by': 'Jeffrey', 'Verified date': D('2026-09-08'), 'Shared with BLLT (perpetual licence)?': 'No' });

  // ---------------------------------------------------------- Deliverables
  ['DLV-005', 'DLV-001', 'DLV-002'].forEach(id => updateById_(ss, 'Deliverables', id, { 'Status': 'Delivered and verified', 'Delivered date': D('2026-09-08'), 'How delivered': 'Email to Andrew Hercus with one-page addendums (event and marketing), Sept 8', 'Evidence ID': evdBllt, 'Verification': 'Verified', 'Exception notes': 'Accepted by BLLT: Invoice #2 submitted for payment Sept 8.' }));
  const dlvPurRow = findRowByText_(ss, 'Deliverables', 'Plain-language deliverable', 'Forward exhibit overview to Pursuit');
  if (dlvPurRow) {
    const dlvPurId = ss.getSheetByName('Deliverables').getRange(dlvPurRow, 1).getValue();
    updateById_(ss, 'Deliverables', dlvPurId, { 'Status': 'Delivered and verified', 'Delivered date': D('2026-09-07'), 'How delivered': 'Forwarded overview, photos and video to Pursuit (gondola) Sept 7', 'Evidence ID': evdPur, 'Verification': 'Verified' });
    updateById_(ss, 'Evidence', evdPur, { 'Deliverable ID': dlvPurId });
  }
  updateById_(ss, 'Deliverables', 'DLV-011', { 'Exact source language': 'Right to have the BLLT Visitor Experience Survey link (https://bll.fyi/VSEvents) distributed to Event attendees.', 'Dependencies': 'Post-festival attendee survey', 'Exception notes': 'BLLT sent the link, suggested copy and mandatory contest text with EN and FR rules on Sept 7 (Gmail 1a07d027283b0b98). ED committed Sept 8 to link it at the end of the post-festival survey. Mandatory text must appear wherever the link is shared.', 'Status': 'In progress' });
  const drRow = findRowByText_(ss, 'Deliverables', 'Plain-language deliverable', 'Drag Brunch details finalized');
  if (drRow) updateById_(ss, 'Deliverables', ss.getSheetByName('Deliverables').getRange(drRow, 1).getValue(), { 'Due date': D('2026-09-12'), 'Status': 'In progress', 'Exception notes': 'ByHendo (Hilary) resumes brunch planning with Moxy around Sept 11 (Gmail 1a06902394ace8fa). Outcome of the Sept 7 Moxy meeting not yet recorded.' });

  const dlv = (o) => appendRecord_(ss, 'Deliverables', Object.assign({ 'Quantity': 1, 'Internal owner': 'Jeffrey', 'Approval required': 'No', 'Evidence required': 'Yes', 'Status': 'Not started', 'Confidence': 'A' }, o));
  // BLLT
  dlv({ 'Partner ID': org.BLLT, 'Partner': 'Banff & Lake Louise Tourism', 'Agreement ID': 'AGR-001', 'Owed by': 'Partner', 'Exact source language': 'Request Sept 8: tagged link on the banfflakelouise.com Banff Pride listing', 'Plain-language deliverable': 'BLLT updates its Banff Pride listing with the tagged link', 'Channel': 'Website', 'Event or activation': 'Festival', 'Internal owner': 'Partner', 'Due date': D('2026-09-15'), 'Status': 'In progress', 'Exception notes': 'Sent to BLLT owned media team Sept 8 (Andrew).', 'Source': 'Gmail 19fd3a82f9a41493' });
  dlv({ 'Partner ID': org.BLLT, 'Partner': 'Banff & Lake Louise Tourism', 'Agreement ID': 'AGR-001', 'Owed by': 'Partner', 'Exact source language': 'Logo placement on the new partners page needs your approval', 'Plain-language deliverable': 'BLLT brand team approves BLLT logo placement on banffpride.ca/partners', 'Channel': 'Assets', 'Event or activation': 'Festival', 'Internal owner': 'Partner', 'Approval required': 'Yes', 'Approval lead time (days)': 10, 'Due date': D('2026-09-15'), 'Status': 'In progress', 'Exception notes': 'With BLLT brand team since Sept 8.', 'Source': 'Gmail 19fd3a82f9a41493' });
  dlv({ 'Partner ID': org.BLLT, 'Partner': 'Banff & Lake Louise Tourism', 'Agreement ID': 'AGR-001', 'Owed by': 'Banff Pride', 'Exact source language': 'Would you like to do a little spotlight on Banff Pride at the Event Advisory Group meeting on September 22?', 'Plain-language deliverable': 'Reply to Andrew on the Sept 22 Event Advisory Group spotlight', 'Channel': 'Email', 'Event or activation': 'Festival', 'Due date': D('2026-09-11'), 'Evidence required': 'No', 'Source': 'Gmail 19fd3a82f9a41493' });
  // Cooper's Dog Cabin (Silver)
  [['Banff Pride', 'Logo on the festival website', "Cooper's logo on banffpride.ca/partners", 'Website', D('2026-09-20'), 'Logo from partner'],
   ['Banff Pride', 'Recognition across our social channels', "Cooper's social recognition", 'Social', D('2026-10-12'), 'Logo from partner'],
   ['Banff Pride', 'Logo on festival and event signage', "Cooper's logo on signage", 'Signage / print', D('2026-09-25'), 'Logo by Sept 14 for print run'],
   ['Banff Pride', 'Featured spotlight in our social and email', "Cooper's featured spotlight (social and email)", 'Email / social', D('2026-10-01'), 'Description and logo from partner'],
   ['Banff Pride', 'Personalized partner report', "Cooper's partner report", 'Report', D('2026-11-30'), 'Evidence; metrics'],
   ['Partner', 'Sign and send it back whenever suits', "Cooper's signs the sponsorship agreement", 'Document', D('2026-09-15'), ''],
   ['Partner', 'Your logo and a short description of Cooper\'s, ideally by September 14', "Cooper's sends logo and description", 'Assets', D('2026-09-14'), ''],
   ['Partner', 'Payable within 30 days of signing', "Cooper's pays BPS Invoice #0908 ($1,500)", 'Payment', D('2026-10-08'), '']
  ].forEach(r => dlv({ 'Partner ID': org.COOP, 'Partner': "Cooper's Dog Cabin", 'Agreement ID': agrCoop, 'Owed by': r[0], 'Exact source language': r[1], 'Plain-language deliverable': r[2], 'Channel': r[3], 'Event or activation': 'Festival', 'Internal owner': r[0] === 'Partner' ? 'Partner' : 'Jeffrey', 'Due date': r[4], 'Dependencies': r[5], 'Source': 'Gmail 1a063b3e50267f65' }));
  // Banff Tea Co.
  dlv({ 'Partner ID': org.TEA, 'Partner': 'Banff Tea Co.', 'Agreement ID': agrTea, 'Owed by': 'Partner', 'Exact source language': 'Bring back the Rainbow Dreams tea blend; full proceeds from the sale of this tea', 'Plain-language deliverable': 'Rainbow Dreams blend on sale and proceeds remitted', 'Channel': 'Delivery', 'Event or activation': 'Festival', 'Internal owner': 'Partner', 'Due date': D('2026-10-31'), 'Source': 'Gmail 1a063b3d344298d2' });
  dlv({ 'Partner ID': org.TEA, 'Partner': 'Banff Tea Co.', 'Agreement ID': agrTea, 'Owed by': 'Banff Pride', 'Exact source language': 'Standard in-kind recognition (ED rule Sept 5)', 'Plain-language deliverable': 'Recognize Banff Tea Co. and promote the Rainbow Dreams blend on website and social', 'Channel': 'Website / social', 'Event or activation': 'Festival', 'Due date': D('2026-10-12'), 'Dependencies': 'Confirm launch date and where to buy', 'Source': 'Gmail 1a063b3d344298d2', 'Confidence': 'B' });
  // Bikescape (benefits offered by ED in the Sept 8 email; ride details pending the Sept 10 call)
  dlv({ 'Partner ID': org.BIKE, 'Partner': 'Bikescape', 'Owed by': 'Partner', 'Exact source language': 'We would love to participate with a guided ride', 'Plain-language deliverable': 'Guided Pride ride during the festival (Oct 6 to 9 or a Saturday)', 'Channel': 'Event', 'Event or activation': 'Guided ride', 'Internal owner': 'Partner', 'Due date': D('2026-10-09'), 'Dependencies': 'Sept 10 call: date, group size, ability level, bikes and helmets, registration and waivers, free or ticketed', 'Source': 'Gmail 1a063b74c624da44', 'Confidence': 'B' });
  dlv({ 'Partner ID': org.BIKE, 'Partner': 'Bikescape', 'Owed by': 'Partner', 'Exact source language': 'Keen to donate a guided experience as prizing', 'Plain-language deliverable': 'Donated guided experience for Pride Bingo (Oct 7) or Pride Market (Oct 10) prizing', 'Channel': 'Delivery', 'Event or activation': 'Drag Bingo', 'Internal owner': 'Partner', 'Due date': D('2026-10-07'), 'Dependencies': 'Confirmation on Sept 10 call', 'Source': 'Gmail 1a063b74c624da44', 'Confidence': 'B' });
  dlv({ 'Partner ID': org.BIKE, 'Partner': 'Bikescape', 'Owed by': 'Banff Pride', 'Exact source language': 'Logo placement on the festival website and in event listings, social posts across our channels, and named credit anywhere the ride or the prize appears', 'Plain-language deliverable': 'Bikescape website logo, event listing, social posts and named credit', 'Channel': 'Website / social', 'Event or activation': 'Guided ride', 'Due date': D('2026-10-12'), 'Dependencies': 'Bikescape confirmation; logo files', 'Source': 'Gmail 1a063b74c624da44 (ED email Sept 8)', 'Confidence': 'B' });
  log.push('Deliverables appended and updated');

  // ------------------------------------------------------------- Tickets
  appendRecord_(ss, 'Tickets', { 'Partner ID': org.BIKE, 'Partner': 'Bikescape', 'Type': 'Prize', 'Description': 'Guided ride experience (donated)', 'Purpose or event': 'Pride Bingo Oct 7 or Pride Market Oct 10', 'Issued to': 'Banff Pride', 'Status': 'Pending', 'Evidence or source': 'Gmail 1a063b74c624da44', 'Next step': 'Confirm on Sept 10 call', 'Deadline': D('2026-09-10') });

  // -------------------------------------------------------------- Events
  appendRecord_(ss, 'Events', { 'Event or activation': 'Lake Minnewanka cruise with Pursuit (ticketed)', 'Date': D('2026-10-03'), 'Host partner ID': org.PUR, 'Venue': 'Lake Minnewanka', 'Type': 'Signature', 'Status': 'Planned', 'Notes': 'Named as a new ticketed event in the Sept 8 plan addendum to BLLT (Gmail 19fd3a82f9a41493). Confirm terms with Pursuit and record them in Agreements.', 'Linked partner IDs': org.PUR, 'Evidence to capture': 'Ticket page; photos; attendance' });
  appendRecord_(ss, 'Events', { 'Event or activation': 'Pride guided ride with Bikescape (date TBD)', 'Date': '', 'Host partner ID': org.BIKE, 'Venue': 'Banff trails', 'Type': 'Outdoor', 'Status': 'Planning', 'Notes': 'Oct 6 to 9 daytime open; Sat Oct 3 or Oct 10 possible. Details on the Sept 10 call.', 'Linked partner IDs': org.BIKE, 'Evidence to capture': 'Event listing; photos; registrations' });
  appendRecord_(ss, 'Events', { 'Event or activation': 'Flow Skate & Friends pop-up (proposed)', 'Date': D('2026-10-03'), 'Host partner ID': org.FLOW, 'Venue': 'Banff Skatepark', 'Type': 'Community', 'Status': 'Planning', 'Notes': 'All-ages skate and dance session, 2 to 5 hours, Saturday preferred. Needs permit or rental check, budget for performers and a hydration sponsor (Gmail 1a082808ad1aa37f; Gemini notes 1a08212a26a853ca).', 'Linked partner IDs': org.FLOW, 'Evidence to capture': 'Permit; photos' });
  appendRecord_(ss, 'Events', { 'Event or activation': 'Black queer joy art pop-up (proposed)', 'Date': '', 'Host partner ID': org.FLOW, 'Venue': 'Banff Centre or Banff Gondola (TBD)', 'Type': 'Arts', 'Status': 'Planning', 'Notes': 'Curated by Zia Stolbie with Black Pride YYC artists. Banff Pride to introduce Banff Centre contacts.', 'Linked partner IDs': org.FLOW + ';' + org.BC + ';' + org.PUR, 'Evidence to capture': 'Venue confirmation' });

  // -------------------------------------------------------- Opportunities
  const oppUpd = (partnerId, fields) => { const row = findRowByValue_(ss, 'Opportunities', 'Partner ID', partnerId); if (row) setFields_(ss.getSheetByName('Opportunities'), row, fields); else appendRecord_(ss, 'Opportunities', Object.assign({ 'Partner ID': partnerId, 'Cycle year': 2026, 'Priority': 'Priority 3', 'Next-action owner': 'Jeffrey', 'Confidence': 'B', 'Verify': 'No' }, fields)); };
  oppUpd(org.BLLT, { 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Reply to Andrew on the Sept 22 Event Advisory Group spotlight; await tagged listing and logo approval', 'Next-action date': D('2026-09-11') });
  oppUpd(org.PUR, { 'Last meaningful contact': D('2026-09-07'), 'Next action': 'Chase Pursuit on the gondola exhibit and the sponsorship decision; Biosphere is holding the pieces', 'Next-action date': D('2026-09-11'), 'Risks': 'Biosphere took the exhibition down Sept 8 and is storing pieces pending the gondola answer.' });
  oppUpd(org.BIO, { 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Tell Tanya the gondola decision as soon as Pursuit replies', 'Next-action date': D('2026-09-11') });
  oppUpd(org.MOXY, { 'Last meaningful contact': D('2026-09-04'), 'Next action': 'Record the Sept 7 meeting outcome; confirm Drag Brunch details in writing; ByHendo resumes around Sept 11', 'Next-action date': D('2026-09-12') });
  oppUpd(org.NIB, { 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Await a booking via the calendar link; follow up if none', 'Next-action date': D('2026-09-15') });
  oppUpd(org.BC, { 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Introduce Zia Stolbie to Banff Centre contacts for the Black queer joy pop-up; confirm venue', 'Next-action date': D('2026-09-12') });
  oppUpd(org.COOP, { 'Partner': "Cooper's Dog Cabin", 'Stage': 'Agreement pending', 'Priority': 'Priority 2', 'Likelihood': 'High', 'Evidence basis': 'Email acceptance of Silver (Sept 7); agreement and invoice sent Sept 8', 'Best-fit ask': 'Silver $1,500', 'Proposed cash': 1500, 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Receive signed agreement, logo and description (by Sept 14); offer march, volunteer shift or prize donation', 'Next-action owner': 'Jeffrey', 'Next-action date': D('2026-09-14'), 'Relationship health': 'Strong', 'Source link': 'Gmail 1a063b3e50267f65', 'Confidence': 'A' });
  oppUpd(org.TEA, { 'Partner': 'Banff Tea Co.', 'Stage': 'Confirmed', 'Priority': 'Priority 3', 'Likelihood': 'High', 'Evidence basis': 'Email Sept 6: Rainbow Dreams blend returns, full proceeds donated', 'Best-fit ask': 'Rainbow Dreams tea blend, full proceeds', 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Agree launch timing, where to buy and how proceeds are remitted; ask for the 2025 proceeds figure', 'Next-action owner': 'Jeffrey', 'Next-action date': D('2026-09-15'), 'Relationship health': 'Strong', 'Source link': 'Gmail 1a063b3d344298d2', 'Confidence': 'A' });
  oppUpd(org.BIKE, { 'Partner': 'Bikescape', 'Stage': 'Verbal commitment', 'Priority': 'Priority 2', 'Likelihood': 'High', 'Evidence basis': 'Email Sept 5 offering a guided ride and a donated experience', 'Best-fit ask': 'Guided ride event plus donated experience for prizing', 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Call Thu Sept 10 09:30 (booked): settle ride date, group size, ability level, bikes and helmets, registration and waivers, free or ticketed; confirm the prize', 'Next-action date': D('2026-09-10'), 'Relationship health': 'Strong', 'Source link': 'Gmail 1a063b74c624da44; booking 1a083168420e01e1', 'Confidence': 'B' });
  oppUpd(org.AIR, { 'Partner': 'Alpine Air Adventures', 'Stage': 'Proposal sent', 'Priority': 'Priority 3', 'Likelihood': 'Medium', 'Evidence basis': 'Asked what we had in mind (Sept 5); Trailhead $500 or donated experience proposed Sept 8', 'Best-fit ask': 'Donated experience for prizing, or Trailhead $500', 'Proposed cash': 500, 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Await reply; follow up if nothing by Sept 15', 'Next-action date': D('2026-09-15'), 'Relationship health': 'Stable', 'Source link': 'Gmail 1a0731481fe5a55f' });
  oppUpd(org.SUN, { 'Partner': 'Sunset Alpine Promotional Products', 'Stage': 'Contacted', 'Priority': 'Priority 3', 'Likelihood': 'Medium-High', 'Evidence basis': 'BVPN intro Aug 26; owner committed $1,000 to Canmore Pride Sept 8 and said Sunset Alpine wants to support Banff Pride too', 'Best-fit ask': 'Trailhead or Silver cash, or branded merchandise in-kind', 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Reply on the intro thread with Banff Pride options (cash tier or promotional product in-kind)', 'Next-action date': D('2026-09-11'), 'Relationship health': 'Stable', 'Source link': 'Gmail 1a03eb85226cfbfe' });
  oppUpd(org.JASE, { 'Partner': 'JASE Security', 'Stage': 'Conversation active', 'Priority': 'Priority 5', 'Likelihood': 'Medium', 'Evidence basis': 'Vendor follow-up Sept 6; quote requested Sept 8', 'Best-fit ask': 'Security for Mason Hall Oct 9 to 11; ask for a community rate and record any discount as in-kind', 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Receive quote for Oct 9, 10, 11 (20:00 to 01:00, up to 120)', 'Next-action date': D('2026-09-12'), 'Relationship health': 'Stable', 'Source link': 'Gmail 1a079f58929fd33c' });
  oppUpd(org.FLOW, { 'Partner': 'Flow Skate & Friends', 'Stage': 'Conversation active', 'Priority': 'Priority 3', 'Likelihood': 'Medium', 'Evidence basis': 'Sept 8 call and written proposal', 'Best-fit ask': 'Community event partnership: skate pop-up Oct 3 and Black queer joy art pop-up', 'Last meaningful contact': D('2026-09-08'), 'Next action': 'Check skatepark permit or rental; introduce Banff Centre; confirm whether budget exists for performers and equipment', 'Next-action date': D('2026-09-12'), 'Relationship health': 'Stable', 'Source link': 'Gmail 1a082808ad1aa37f; 1a08212a26a853ca' });
  log.push('Opportunities updated');

  // --------------------------------------------------------- BrandAssets
  updateById_(ss, 'BrandAssets', 'BRA-001', { 'Approval status': 'Pending: BLLT brand team reviewing partners-page placement (requested Sept 8)', 'Last verified': D('2026-09-08') });
  appendRecord_(ss, 'BrandAssets', { 'Partner ID': org.COOP, 'Approved public name': "Cooper's Dog Cabin", 'Library folder link': 'https://drive.google.com/drive/folders/1gNzWFbVRe_DqrCuoHwQwvO0SPERtzlj-', 'Restrictions and approval rules': 'Use only files supplied by the partner; Silver placements (website, social, signage, spotlight)', 'Approval status': 'Pending: requested Sept 8, due Sept 14', 'Source contact / thread': 'Gmail 1a063b3e50267f65', 'Notes': 'Needed for the print and signage run.' });
  appendRecord_(ss, 'BrandAssets', { 'Partner ID': org.TEA, 'Approved public name': 'Banff Tea Co.', 'Library folder link': 'https://drive.google.com/drive/folders/15z9172oOSJiPS35g1mkY8J49AQ4stsr-', 'Approval status': 'Pending: not requested', 'Source contact / thread': 'Gmail 1a063b3d344298d2' });
  appendRecord_(ss, 'BrandAssets', { 'Partner ID': org.BIKE, 'Approved public name': 'Bikescape', 'Library folder link': 'https://drive.google.com/drive/folders/1RoTVe8CD2AfI-Z5nLV1ZpK7rkOVJ7p6C', 'Approval status': 'Pending: not requested', 'Source contact / thread': 'Gmail 1a063b74c624da44' });

  // ------------------------------------------------------------- Reports
  [[org.COOP, "Cooper's Dog Cabin"], [org.TEA, 'Banff Tea Co.'], [org.BIKE, 'Bikescape']].forEach(p => appendRecord_(ss, 'Reports', { 'Partner ID': p[0], 'Partner': p[1], 'Report type': 'Partner report', 'Due': D('2026-11-30'), 'Status': 'Not started' }));

  // ------------------------------------------------------- ThankYouEmail
  [[org.COOP, "Cooper's Dog Cabin"], [org.TEA, 'Banff Tea Co.'], [org.BIKE, 'Bikescape']].forEach(p => appendRecord_(ss, 'ThankYouEmail', { 'Partner ID': p[0], 'Partner': p[1], 'Include?': 'Not reviewed', 'Final approval': 'Not reviewed', 'Notes': 'Added Sept 8 from email confirmations.' }));

  // ------------------------------------------------------ Communications
  const com = (o) => appendRecord_(ss, 'Communications', Object.assign({ 'Owner': 'Jeffrey', 'Triage': 'Open' }, o));
  com({ 'Partner ID': org.BLLT, 'Partner': 'Banff & Lake Louise Tourism', 'Contact': 'Andrew Hercus', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'Schedule of Events & Deliverables for Invoice #2', 'Summary': 'Andrew asked for finalized plans; ED sent event and marketing addendums (headliners, Pride Market Oct 10, Pursuit cruise Oct 3, Bell Media and Bow Valley Insider in market) and asked for a tagged listing link and logo placement approval. Andrew sent both to his teams and submitted Invoice #2 for payment. He offered a Banff Pride spotlight at the Sept 22 Event Advisory Group.', 'Commitments by Banff Pride': 'Reply on the EAG spotlight', 'Commitments by partner': 'Tagged listing; logo placement decision; Invoice #2 paid', 'Decisions': 'Invoice #2 deliverables accepted', 'Open questions': 'EAG spotlight yes or no', 'Follow-up action': 'Reply to Andrew re Sept 22 spotlight', 'Due date': D('2026-09-11'), 'Source link': 'Gmail 19fd3a82f9a41493', 'Proposed records': 'DLV-001, DLV-002, DLV-005 verified; 3 new DLV; CTB-002 note' });
  com({ 'Partner ID': org.BLLT, 'Partner': 'Banff & Lake Louise Tourism', 'Contact': 'Madeleine Gribble', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'BLLT VES Survey information', 'Summary': 'BLLT sent the Visitor Experience Survey link, suggested copy and mandatory contest text with EN and FR rules, citing the contract clause. ED will link it at the end of the post-festival survey; BLLT confirmed the survey question updates automatically.', 'Commitments by Banff Pride': 'Include survey link and mandatory text in the post-festival survey', 'Decisions': 'Closed', 'Follow-up action': 'Build into post-festival survey', 'Due date': D('2026-10-13'), 'Source link': 'Gmail 1a07d027283b0b98', 'Proposed records': 'DLV-011 updated', 'Triage': 'Closed' });
  com({ 'Partner ID': org.BLLT, 'Partner': 'Banff & Lake Louise Tourism', 'Contact': 'Barb Bean', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'BLLT membership inquiry', 'Summary': 'ED asked BLLT which membership class applies to Banff Pride, the business category for the financial contribution, whether the $500 associate fee is annual, and for the bylaws and membership policies.', 'Open questions': 'Membership class; contribution amount; fee frequency; policies', 'Follow-up action': 'Await BLLT answers', 'Due date': D('2026-09-15'), 'Source link': 'Gmail 1a0681faa82afe03', 'Triage': 'Waiting' });
  com({ 'Partner ID': org.COOP, 'Partner': "Cooper's Dog Cabin", 'Contact': 'Florence', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'Re: Partner with the 2026 Banff Pride Festival', 'Summary': "Cooper's accepted Silver on Sept 7 and asked about other ways to help. ED sent the agreement and BPS Invoice #0908, asked for logo and description by Sept 14, and suggested the march, a volunteer shift or a prize donation.", 'Commitments by Banff Pride': 'Silver benefits (website, social, signage, spotlight)', 'Commitments by partner': 'Silver sponsorship; logo and description by Sept 14', 'Decisions': 'Silver accepted', 'Open questions': 'Prize donation; march participation', 'Follow-up action': 'Receive signed agreement and assets', 'Due date': D('2026-09-14'), 'Source link': 'Gmail 1a063b3e50267f65', 'Proposed records': agrCoop + '; ' + ctbCoop + '; 8 DLV' });
  com({ 'Partner ID': org.TEA, 'Partner': 'Banff Tea Co.', 'Contact': 'Siona', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'Re: Partner with the 2026 Banff Pride Festival', 'Summary': 'Banff Tea Co. will bring back the Rainbow Dreams blend and donate full proceeds. ED accepted Sept 8.', 'Commitments by partner': 'Rainbow Dreams blend; full proceeds', 'Decisions': 'Confirmed', 'Open questions': 'Launch date; where to buy; how proceeds are remitted', 'Follow-up action': 'Agree logistics and ask for 2025 proceeds figure', 'Due date': D('2026-09-15'), 'Source link': 'Gmail 1a063b3d344298d2', 'Proposed records': agrTea + '; ' + ctbTea });
  com({ 'Partner ID': org.BIKE, 'Partner': 'Bikescape', 'Contact': 'Clare McCann', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'Re: Partner with the 2026 Banff Pride Festival', 'Summary': 'Bikescape offered a guided ride and a donated experience. ED said yes to both, proposed Oct 6 to 9 or a Saturday, listed what he needs (group size, ability, bikes and helmets, registration and waivers, free or ticketed) and set out recognition. Call booked Thu Sept 10 09:30.', 'Commitments by Banff Pride': 'Website logo, event listing, social posts, named credit', 'Commitments by partner': 'Guided ride; donated experience', 'Open questions': 'Ride date and format; free or ticketed; prize placement (Bingo Oct 7 or Market Oct 10)', 'Follow-up action': 'Call Sept 10 09:30', 'Due date': D('2026-09-10'), 'Source link': 'Gmail 1a063b74c624da44; booking 1a083168420e01e1', 'Proposed records': ctbBike + '; 3 DLV; TIX; EVT' });
  com({ 'Partner ID': org.AIR, 'Partner': 'Alpine Air Adventures', 'Contact': 'Dominic Boon', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'Banff Pride 2026', 'Summary': 'Dominic asked what we had in mind. ED proposed promoting their October experiences to visitors, a donated experience for prizing with logo, social post and named credit, or Trailhead at $500.', 'Commitments by Banff Pride': 'Promote October experiences; prize recognition if donated', 'Open questions': 'Which option Alpine Air chooses', 'Follow-up action': 'Follow up if no reply', 'Due date': D('2026-09-15'), 'Source link': 'Gmail 1a0731481fe5a55f', 'Triage': 'Waiting' });
  com({ 'Partner ID': org.NIB, 'Partner': 'Nibble Tours', 'Contact': 'Nibble Tours', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'Re: Partner with the 2026 Banff Pride Festival', 'Summary': 'ED declined the tour offer for now, suggested Taste The Rainbow participation or in-kind tour passes as prizing, and sent the booking link.', 'Follow-up action': 'Await booking', 'Due date': D('2026-09-15'), 'Source link': 'Gmail 1a063b764c7a046e', 'Triage': 'Waiting' });
  com({ 'Partner ID': org.SUN, 'Partner': 'Sunset Alpine Promotional Products', 'Contact': 'Dov Simenauer', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'Introduction: Sunset Alpine Promotional Products', 'Summary': 'BVPN introduced Sunset Alpine (Canmore) to both Prides on Aug 26. On Sept 8 Dov asked Canmore Pride to record a $1,000 community sponsorship. Banff Pride has not yet replied on the thread.', 'Open questions': 'Banff Pride ask: cash tier or promotional product in-kind', 'Follow-up action': 'Reply to Dov with Banff Pride options', 'Due date': D('2026-09-11'), 'Source link': 'Gmail 1a03eb85226cfbfe' });
  com({ 'Partner ID': org.PUR, 'Partner': 'Pursuit', 'Contact': 'Pursuit (gondola)', 'Date': D('2026-09-07'), 'Channel': 'Email', 'Subject or meeting title': 'Banff Pride Discussion (exhibit overview forwarded)', 'Summary': 'ED forwarded the Queer(y)ing Ecologies overview, photos and video to Pursuit. Biosphere took the exhibition down Sept 8 and is holding pieces pending the gondola decision.', 'Commitments by Banff Pride': 'Keep Biosphere informed', 'Open questions': 'Gondola exhibit yes or no; sponsorship decision', 'Follow-up action': 'Chase Pursuit', 'Due date': D('2026-09-11'), 'Source link': 'Gmail 19e6f5aa8875094f', 'Proposed records': evdPur });
  com({ 'Partner ID': org.BYH, 'Partner': 'Byhendo', 'Contact': 'Hilary Urquhart', 'Date': D('2026-09-04'), 'Channel': 'Email', 'Subject or meeting title': 'Banff Pride Oct 04: Moxy + ByHendo', 'Summary': 'ED introduced ByHendo to Moxy for the Drag Brunch. Hilary will restart communication toward the end of the following week (around Sept 11).', 'Commitments by partner': 'Follow up with Moxy on brunch ideas', 'Follow-up action': 'Check in if nothing by Sept 12', 'Due date': D('2026-09-12'), 'Source link': 'Gmail 1a06902394ace8fa', 'Triage': 'Waiting' });
  com({ 'Partner ID': org.JASE, 'Partner': 'JASE Security', 'Contact': 'Wendell Minty', 'Date': D('2026-09-08'), 'Channel': 'Email', 'Subject or meeting title': 'JASE Security follow-up', 'Summary': 'JASE offered security again. ED explained the grant-dependent budget and requested a quote for Mason Hall on Oct 9, 10 and 11, 20:00 to 01:00, up to 120 attendees.', 'Commitments by partner': 'Quote', 'Follow-up action': 'Receive and review quote', 'Due date': D('2026-09-12'), 'Source link': 'Gmail 1a079f58929fd33c', 'Triage': 'Waiting' });
  com({ 'Partner ID': org.FLOW, 'Partner': 'Flow Skate & Friends', 'Contact': 'Zia Stolbie', 'Date': D('2026-09-08'), 'Channel': 'Meeting', 'Subject or meeting title': 'Partner with Pride chat and written proposal', 'Summary': 'Zia proposed an all-ages Flow Skate & Friends pop-up at Banff Skatepark on Oct 3 (2 to 5 hours, possible DJ or performers, hydration sponsor) and a curated Black queer joy art pop-up with Black Pride YYC artists at the Banff Centre or gondola.', 'Commitments by Banff Pride': 'Check skatepark permit or rental; introduce Banff Centre; review budget', 'Commitments by partner': 'Send proposal (received Sept 8)', 'Open questions': 'Budget for performers and equipment; venue for the art pop-up', 'Follow-up action': 'Permit check and Banff Centre intro', 'Due date': D('2026-09-12'), 'Source link': 'Gmail 1a082808ad1aa37f; Gemini notes 1a08212a26a853ca', 'Proposed records': '2 EVT' });
  log.push('Communications appended');

  // --------------------------------------------------------- SourceIndex
  [['Gmail', 'BLLT Invoice #2 deliverables thread', '19fd3a82f9a41493'], ['Gmail', 'BLLT VES survey', '1a07d027283b0b98'], ['Gmail', 'BLLT membership inquiry', '1a0681faa82afe03'],
   ['Gmail', "Cooper's Dog Cabin Silver", '1a063b3e50267f65'], ['Gmail', 'Banff Tea Co. Rainbow Dreams', '1a063b3d344298d2'], ['Gmail', 'Bikescape guided ride', '1a063b74c624da44'],
   ['Gmail', 'Bikescape call booking Sept 10', '1a083168420e01e1'], ['Gmail', 'Alpine Air Adventures', '1a0731481fe5a55f'], ['Gmail', 'Nibble Tours reply', '1a063b764c7a046e'],
   ['Gmail', 'Sunset Alpine introduction', '1a03eb85226cfbfe'], ['Gmail', 'Pursuit exhibit overview', '19e6f5aa8875094f'], ['Gmail', 'Moxy + ByHendo brunch', '1a06902394ace8fa'],
   ['Gmail', 'JASE Security quote request', '1a079f58929fd33c'], ['Gmail', 'Flow Skate & Friends proposal', '1a082808ad1aa37f'], ['Gemini notes', 'Zia Stolbie chat notes', '1a08212a26a853ca']]
    .forEach(r => appendRecord_(ss, 'SourceIndex', { 'System': r[0], 'Title': r[1], 'ID or link': r[2], 'Source modified': D('2026-09-08'), 'Retrieved': D('2026-09-08'), 'Authority': '3', 'Note': 'Update pack 2026-09-08' }));

  props.setProperty(UPD_KEY, new Date().toISOString());
  appendAudit_(ss, actor_(), 'Update pack applied', 'update-2026-09-08', log.join(' | ').slice(0, 1500), 'Gmail 4 to 8 Sept (15 threads)');
  SpreadsheetApp.getUi().alert('Update pack 2026-09-08 applied', log.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}

// ------------------------------------------------------------- helpers

function headers_(sh) {
  const hdr = headerRow(sh.getName());
  const vals = sh.getRange(hdr, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {}; vals.forEach((v, i) => { if (v !== '') map[String(v).trim()] = i + 1; });
  return map;
}

function setFields_(sh, row, fields) {
  const H = headers_(sh);
  Object.keys(fields).forEach(k => { if (!H[k]) throw new Error(sh.getName() + ': unknown column "' + k + '"'); sh.getRange(row, H[k]).setValue(fields[k]); });
}

function findRowByValue_(ss, sheetName, header, value) {
  const sh = ss.getSheetByName(sheetName); const H = headers_(sh); const hdr = headerRow(sheetName);
  const col = H[header]; if (!col) throw new Error(sheetName + ': unknown column ' + header);
  const vals = sh.getRange(hdr + 1, col, Math.max(sh.getLastRow() - hdr, 1), 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === value) return hdr + 1 + i;
  return 0;
}

function findRowByText_(ss, sheetName, header, text) {
  const sh = ss.getSheetByName(sheetName); const H = headers_(sh); const hdr = headerRow(sheetName);
  const vals = sh.getRange(hdr + 1, H[header], Math.max(sh.getLastRow() - hdr, 1), 1).getValues();
  for (let i = 0; i < vals.length; i++) if (String(vals[i][0]).indexOf(text) >= 0) return hdr + 1 + i;
  return 0;
}

function updateById_(ss, sheetName, id, fields) {
  const idHeader = headers_(ss.getSheetByName(sheetName));
  const first = Object.keys(idHeader).sort((a, b) => idHeader[a] - idHeader[b])[0];
  const row = findRowByValue_(ss, sheetName, first, id);
  if (!row) throw new Error(sheetName + ': ' + id + ' not found');
  setFields_(ss.getSheetByName(sheetName), row, fields);
}

/** Appends a record, assigns the next ID in column A (unless given), copies formula columns from the row above. */
function appendRecord_(ss, sheetName, fields) {
  const sh = ss.getSheetByName(sheetName); const H = headers_(sh); const hdr = headerRow(sheetName);
  const row = sh.getLastRow() + 1;
  const idCol = 1;
  const prefix = ID_PREFIX[sheetName] || (sheetName === 'SourceIndex' ? 'SRC' : null);
  const idHeaderName = Object.keys(H).find(k => H[k] === idCol);
  let id = fields[idHeaderName];
  if (!id && prefix) id = nextId_(sh, prefix, hdr);
  // copy formulas from the previous data row
  if (row - 1 > hdr) {
    const prev = sh.getRange(row - 1, 1, 1, sh.getLastColumn()).getFormulasR1C1()[0];
    prev.forEach((f, i) => { if (f) sh.getRange(row, i + 1).setFormulaR1C1(f); });
  }
  if (id && idHeaderName) sh.getRange(row, idCol).setValue(id);
  Object.keys(fields).forEach(k => { if (k === idHeaderName) return; if (!H[k]) throw new Error(sheetName + ': unknown column "' + k + '"'); sh.getRange(row, H[k]).setValue(fields[k]); });
  if (sheetName === 'Organizations') { sh.getRange(row, H['Created']).setValue(new Date()); sh.getRange(row, H['Updated']).setValue(new Date()); }
  return id;
}

/** Creates the organization with a fixed ID when free, otherwise the next free ID. Returns the ID. */
function upsertOrg_(ss, wantedId, fields) {
  const existing = findRowByValue_(ss, 'Organizations', 'Approved public name', fields['Approved public name']);
  if (existing) return ss.getSheetByName('Organizations').getRange(existing, 1).getValue();
  const taken = findRowByValue_(ss, 'Organizations', 'Partner ID', wantedId);
  const id = taken ? null : wantedId;
  const obj = Object.assign({}, fields); if (id) obj['Partner ID'] = id;
  return appendRecord_(ss, 'Organizations', obj);
}
