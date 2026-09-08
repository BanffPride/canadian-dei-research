/**
 * Banff Pride Partnerships: Google Sheets design and automation layer.
 * Version 1.1 (7 September 2026).
 *
 * Install once: Extensions > Apps Script, replace the default Code.gs with this
 * file, save, then run setupWorkbook from the editor (authorize when asked).
 * Reload the Sheet: a "Banff Pride" menu appears. Everything here is idempotent;
 * running setupWorkbook again is safe.
 *
 * Safety model (unchanged): this script never sends email, never edits signed
 * agreements, and never marks a contractual deliverable Delivered and verified
 * without an Evidence ID. It writes only to this spreadsheet.
 */

// ---------------------------------------------------------------- constants

const RECORD_TABS = ['Organizations', 'Contacts', 'Opportunities', 'Agreements', 'Contributions',
  'Deliverables', 'Evidence', 'Communications', 'BrandAssets', 'Tickets', 'Events', 'Metrics',
  'Reports', 'Renewals', 'Decisions', 'Improvements', 'ThankYouEmail'];
const VIEW_TABS = ['Decisions Required', 'Directory', 'Prospect Pipeline', 'Confirmed Support',
  'Deliverables Due', 'Deliverables at Risk', 'Blocked', 'Missing Evidence', 'Follow-up Needed',
  'Brand Assets View', 'Events View', 'Reporting Readiness', 'Renewal Pipeline', 'Data Quality'];
const SYSTEM_TABS = ['AuditLog', 'SourceIndex', 'Lists'];

// Header row per tab (data starts on the next row).
const HEADER_ROW = { Contacts: 2, ThankYouEmail: 2 };
function headerRow(name) { return HEADER_ROW[name] || 1; }

// ID prefixes per record tab (column A).
const ID_PREFIX = { Organizations: 'ORG', Contacts: 'CON', Opportunities: 'OPP', Agreements: 'AGR',
  Contributions: 'CTB', Deliverables: 'DLV', Evidence: 'EVD', Communications: 'COM', BrandAssets: 'BRA',
  Tickets: 'TIX', Events: 'EVT', Metrics: 'MET', Reports: 'RPT', Renewals: 'REN', Decisions: 'DEC',
  Improvements: 'IMP' };

// Tabs where column B is Partner ID and column C is the partner's public name.
const PARTNER_NAME_TABS = ['Contacts', 'Opportunities', 'Agreements', 'Contributions', 'Deliverables',
  'Communications', 'BrandAssets', 'Tickets', 'Reports', 'Renewals'];

const COLOUR = {
  view: '#1f4e79', record: '#1f6e43', system: '#7f7f7f', email: '#c05a2b', dashboard: '#6c0b06',
  headerBg: '#1f4e79', headerFg: '#ffffff',
  red: '#f4cccc', redFg: '#990000', amber: '#fff2cc', amberFg: '#7f6000', green: '#d9ead3', greenFg: '#274e13',
  purple: '#d9d2e9', purpleFg: '#20124d', grey: '#efefef', greyFg: '#666666', blue: '#cfe2f3', blueFg: '#0b5394'
};

// Drive locations (folder IDs are not secrets; they only resolve for people with access).
const LINKS = {
  'Partnerships (top level)': 'https://drive.google.com/drive/folders/1j9BGkxjapQ4DUK75eCm2J9kxgMx1N-Eo',
  'Brand Asset Library': 'https://drive.google.com/drive/folders/1eHlsfZqe2kSe_FH4KSevTNoKwXJ_I3xu',
  'Archive - superseded trackers': 'https://drive.google.com/drive/folders/150KYLyznKKC0maLD4IBu7zw07j7UO-pA',
  '04 Sponsorship Delivery': 'https://drive.google.com/drive/folders/1sKYE6lJ4dZU12wIlRBM7jvG1eLzUXCiw',
  'Operating Guide': 'https://docs.google.com/document/d/1O4ApancucSPGKAYAHlYQnXQ2Gn6Y0cVxhvi37sUb6jw/edit'
};
const PARTNER_FOLDERS = {
  'ORG-001': '1WlRbdXgJr1guhkVeC7lCTFbr6S68e3Yz', 'ORG-002': '1vWbIRZgwcUVA9QD_kHTK0WaZC7ixEYHe',
  'ORG-003': '1mO2I6wItzf9EtI_0P5oxt8Nmbo8Gn376', 'ORG-004': '1ShcCO_WM3gtHif9slcGbVRwUVfF72iXx',
  'ORG-005': '1eGcCB_sZvgOwTFwTTyDHSTN3YRT_xIvg', 'ORG-006': '15Khi97j-nVC1NUHOPLYyCSuQy1D96O9T',
  'ORG-007': '10Fw8GHFIcRJnBLrt9UDT-E5UjPt4tT_B', 'ORG-008': '106_qlbsjfFECNUMvWt3JDy6oGZn-NOdW',
  'ORG-009': '10wyWpw9DN4owUyMN7rce078R6ghvjnUQ', 'ORG-010': '1XIdHyXorKetIbK7enTkvprTJ5V7RZB3b',
  'ORG-011': '16z6Y47pZj3eSKfMJ4c_5yQNCx23PQofF', 'ORG-012': '19LXQgPYg8W1Rag83fGLClONr9gfhSM33',
  'ORG-013': '1Sqj_4uTr64UJI3oTXK1YifzcEeCRImxO', 'ORG-014': '1XIVf87eNHYn46n6vYBrLPwgmfvT9nkMC',
  'ORG-015': '1L-EBrtwd1xjlqBmObccl2tvo5JdHvXEt',
  'ORG-102': '15z9172oOSJiPS35g1mkY8J49AQ4stsr-', 'ORG-127': '1gNzWFbVRe_DqrCuoHwQwvO0SPERtzlj-', 'ORG-155': '1RoTVe8CD2AfI-Z5nLV1ZpK7rkOVJ7p6C'
};
const TEMPLATE_FOLDER = 'https://drive.google.com/drive/folders/1WwepLiH29J1F1OM1MhdZe9eStzjSm3Gf';

// -------------------------------------------------------------------- menu

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Banff Pride')
    .addItem('Go to ED Dashboard', 'goToDashboard')
    .addItem('Add a new partner', 'addNewPartner')
    .addItem('Verify workbook', 'verifyWorkbook')
    .addItem('Log exception snapshot to AuditLog', 'logExceptionSnapshot')
    .addSeparator()
    .addItem('Apply design and automation (setup)', 'setupWorkbook')
    .addItem('Schedule Monday 07:00 snapshot', 'scheduleWeeklySnapshot')
    .addItem('Remove scheduled snapshot', 'removeWeeklySnapshot')
    .addToUi();
}

function goToDashboard() {
  const ss = SpreadsheetApp.getActive();
  ss.setActiveSheet(ss.getSheetByName('ED Dashboard'));
}

// ------------------------------------------------------------------- setup

function setupWorkbook() {
  const ss = SpreadsheetApp.getActive();
  fixKnownIssues_(ss);
  setNamedRanges_(ss);
  styleTabs_(ss);
  setValidations_(ss);
  setConditionalFormats_(ss);
  setHeaderNotes_(ss);
  groupHelperColumns_(ss);
  protectTabs_(ss);
  buildDashboardExtras_(ss);
  linkBrandFolders_(ss);
  appendAudit_(ss, actor_(), 'Setup applied', 'all', 'Design and automation layer v1.1 applied (formats, validations, protections, dashboard links, onEdit rules).', 'Apps Script');
  SpreadsheetApp.getUi().alert('Banff Pride Partnerships', 'Design and automation applied. Reload the tab to see the menu.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function fixKnownIssues_(ss) {
  // Dashboard "Total support received" referenced itself in v1.0.
  ss.getRange('ED Dashboard!E9').setFormula('=E6+E8');
  // Literal "None" text left in the Communications follow-up column.
  const com = ss.getSheetByName('Communications');
  const rng = com.getRange(2, 13, Math.max(com.getLastRow() - 1, 1), 1);
  const vals = rng.getValues().map(r => [r[0] === 'None' ? '' : r[0]]);
  rng.setValues(vals);
}

function setNamedRanges_(ss) {
  const names = { WarningDays: 'Lists!B31', StaleDays: 'Lists!B32', FestivalStart: 'Lists!B33', FestivalEnd: 'Lists!B34' };
  Object.keys(names).forEach(n => ss.setNamedRange(n, ss.getRange(names[n])));
}

function styleTabs_(ss) {
  const style = (name, colour, hdr) => {
    const sh = ss.getSheetByName(name); if (!sh) return;
    sh.setTabColor(colour);
    const lastCol = Math.max(sh.getLastColumn(), 1);
    const header = sh.getRange(hdr, 1, 1, lastCol);
    header.setBackground(COLOUR.headerBg).setFontColor(COLOUR.headerFg).setFontWeight('bold')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('middle');
    sh.setRowHeight(hdr, 36);
    const lastRow = Math.max(sh.getLastRow(), hdr + 1);
    const body = sh.getRange(hdr + 1, 1, lastRow - hdr, lastCol);
    body.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment('top');
  };
  RECORD_TABS.forEach(n => {
    style(n, n === 'ThankYouEmail' ? COLOUR.email : COLOUR.record, headerRow(n));
    const sh = ss.getSheetByName(n);
    sh.getBandings().forEach(b => b.remove());
    const lastRow = Math.max(sh.getLastRow(), headerRow(n) + 1);
    sh.getRange(headerRow(n) + 1, 1, lastRow - headerRow(n), sh.getLastColumn())
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
    if (sh.getFilter()) sh.getFilter().remove();
    sh.getRange(headerRow(n), 1, Math.min(600, sh.getMaxRows()) - headerRow(n) + 1, sh.getLastColumn()).createFilter();
    sh.setFrozenRows(headerRow(n));
    sh.setFrozenColumns(n === 'Contacts' ? 3 : 2);
  });
  VIEW_TABS.forEach(n => {
    style(n, COLOUR.view, 4);
    const sh = ss.getSheetByName(n);
    sh.getRange('A1').setFontSize(14).setFontWeight('bold');
    sh.setFrozenRows(4);
  });
  SYSTEM_TABS.forEach(n => { style(n, COLOUR.system, 1); ss.getSheetByName(n).setFrozenRows(1); });
  const dash = ss.getSheetByName('ED Dashboard');
  dash.setTabColor(COLOUR.dashboard);
  dash.getRange('A1').setFontSize(16).setFontWeight('bold');
  ['A4:E4', 'A16:E16', 'A29:E29'].forEach(a => dash.getRange(a).setBackground(COLOUR.headerBg).setFontColor(COLOUR.headerFg).setFontWeight('bold'));
  dash.getRange('B5:B13').setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
  dash.getRange('E5:E13').setFontWeight('bold');
  dash.setFrozenRows(2);
  const readme = ss.getSheetByName('README');
  readme.setTabColor(COLOUR.dashboard);
  readme.getRange('A1').setFontSize(16).setFontWeight('bold');
  readme.getRange('A1:A40').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  readme.setColumnWidth(1, 900);
  // Number formats that must survive re-import.
  const fmt = (a1, f) => ss.getRange(a1).setNumberFormat(f);
  ['Organizations!N2:O600', 'Opportunities!R2:R600', 'Opportunities!U2:U600', 'Agreements!F2:G600', 'Agreements!I2:J600',
    'Agreements!L2:L600', 'Contributions!N2:P600', 'Deliverables!L2:L600', 'Deliverables!Q2:Q600', 'Evidence!F2:F600',
    'Evidence!H2:H600', 'Communications!E2:E600', 'Communications!O2:O600', 'BrandAssets!J2:J600', 'BrandAssets!M2:N600',
    'Tickets!K2:K600', 'Tickets!N2:N600', 'Events!C2:C600', 'Metrics!K2:K600', 'Reports!E2:E600', 'Reports!O2:O600',
    'Renewals!E2:E600', 'Decisions!B2:B600', 'Decisions!H2:H600', 'Improvements!B2:B600', 'ThankYouEmail!K3:M600',
    'Contacts!J3:J600'].forEach(a => fmt(a, 'yyyy-mm-dd'));
  ['Opportunities!M2:Q600', 'Contributions!G2:J600', 'Contributions!W2:W600', 'Organizations!Q2:R600',
    'Reports!J2:L600', 'Deliverables!V2:V600'].forEach(a => fmt(a, '$#,##0'));
  fmt('AuditLog!A2:A2000', 'yyyy-mm-dd hh:mm');
}

function setValidations_(ss) {
  const list = (a1, allowInvalid, help) => {
    const b = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getRange(a1), true).setAllowInvalid(!!allowInvalid);
    if (help) b.setHelpText(help);
    return b.build();
  };
  const set = (target, rule) => ss.getRange(target).setDataValidation(rule);
  set('Organizations!J2:J600', list('Lists!$O$2:$O$7', true, 'Relationship owner'));
  set('Opportunities!E2:E600', list('Lists!$B$2:$B$17', false, 'Confirmed needs written confirmation; Closed won needs a signed agreement.'));
  set('Opportunities!F2:F600', list('Lists!$G$2:$G$6', false));
  set('Opportunities!H2:H600', list('Lists!$H$2:$H$6', true));
  set('Opportunities!T2:T600', list('Lists!$O$2:$O$7', true));
  set('Opportunities!V2:V600', list('Lists!$I$2:$I$5', true));
  set('Opportunities!Y2:Y600', list('Lists!$L$2:$L$4', false));
  set('Opportunities!Z2:Z600', list('Lists!$F$2:$F$3', false));
  set('Agreements!H2:H600', list('Lists!$C$2:$C$10', false, 'Fully executed requires both signature dates and a document link.'));
  set('Agreements!R2:R600', list('Lists!$L$2:$L$4', false));
  set('Contributions!D2:D600', list('Lists!$J$2:$J$3', false));
  set('Contributions!E2:E600', list('Lists!$K$2:$K$16', false, 'Promo codes, discounts and free inclusions are in-kind (ED rule 2026-09-05).'));
  set('Contributions!T2:T600', list('Lists!$L$2:$L$4', false));
  set('Deliverables!E2:E600', list('Lists!$E$2:$E$3', false));
  set('Deliverables!K2:K600', list('Lists!$O$2:$O$7', true));
  set('Deliverables!N2:N600', list('Lists!$F$2:$F$3', false));
  set('Deliverables!P2:P600', list('Lists!$D$2:$D$11', false, 'Delivered and verified needs an Evidence ID in column T. The sheet will refuse it otherwise.'));
  set('Deliverables!S2:S600', list('Lists!$F$2:$F$3', false));
  set('Deliverables!U2:U600', list('Lists!$Q$2:$Q$4', false));
  set('Deliverables!Y2:Y600', list('Lists!$L$2:$L$4', false));
  set('Evidence!I2:I600', list('Lists!$F$2:$F$3', false));
  set('Communications!F2:F600', list('Lists!$M$2:$M$18', false));
  set('Communications!N2:N600', list('Lists!$O$2:$O$7', true));
  set('Communications!R2:R600', list('Lists!$N$2:$N$4', false));
  set('Reports!F2:F600', list('Lists!$R$2:$R$6', false));
  set('Renewals!F2:F600', list('Lists!$S$2:$S$6', false));
  set('Decisions!E2:E600', list('Lists!$P$2:$P$6', false, 'Change this to record your decision. Decided date fills in automatically.'));
  set('ThankYouEmail!F3:F600', list('Lists!$T$2:$T$5', false, 'Recipient review (28 September).'));
  set('ThankYouEmail!I3:I600', list('Lists!$T$2:$T$5', false, 'Final approval per recipient (30 September). Nothing sends without Approved here.'));
  // Partner ID pickers.
  ['Contacts!B3:B600', 'Opportunities!B2:B600', 'Agreements!B2:B600', 'Contributions!B2:B600', 'Deliverables!B2:B600',
    'Communications!B2:B600', 'BrandAssets!B2:B600', 'Tickets!B2:B600', 'Reports!B2:B600', 'Renewals!B2:B600',
    'ThankYouEmail!A3:A600', 'Events!D2:D600', 'Metrics!G2:G600']
    .forEach(a => set(a, list('Organizations!$A$2:$A$600', true, 'Pick a Partner ID; the name fills in automatically.')));
  ['Contributions!V2:V600', 'Deliverables!D2:D600', 'Tickets!D2:D600'].forEach(a => set(a, list('Agreements!$A$2:$A$600', true)));
  set('Deliverables!T2:T600', list('Evidence!$A$2:$A$600', true, 'Evidence ID from the Evidence tab.'));
  set('Evidence!B2:B600', list('Deliverables!$A$2:$A$600', true));
  // Dashboard partner lookup.
  set('ED Dashboard!H11', list('Organizations!$C$2:$C$600', false, 'Choose a partner'));
}

function setConditionalFormats_(ss) {
  const R = (sheet, a1) => sheet.getRange(a1);
  const textRule = (rng, text, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(text).setBackground(bg).setFontColor(fg).setRanges([rng]).build();
  const formulaRule = (rng, f, bg, fg) => { const b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f).setBackground(bg); if (fg) b.setFontColor(fg); return b.setRanges([rng]).build(); };
  const apply = (name, rules) => { const sh = ss.getSheetByName(name); if (sh) sh.setConditionalFormatRules(rules); };

  const dash = ss.getSheetByName('ED Dashboard');
  apply('ED Dashboard', [
    formulaRule(R(dash, 'A5:C13'), '=$B5>0', COLOUR.amber, COLOUR.amberFg),
    textRule(R(dash, 'E17:E26'), 'Overdue', COLOUR.red, COLOUR.redFg),
    textRule(R(dash, 'E17:E26'), 'Due soon', COLOUR.amber, COLOUR.amberFg),
    textRule(R(dash, 'E17:E26'), 'Blocked', COLOUR.purple, COLOUR.purpleFg),
    formulaRule(R(dash, 'A30:D39'), '=AND($B30<>"",$B30<TODAY())', COLOUR.red, COLOUR.redFg),
    formulaRule(R(dash, 'B42:B43'), '=NOT(ISNUMBER(B42))', COLOUR.red, COLOUR.redFg)
  ]);

  const dlv = ss.getSheetByName('Deliverables');
  apply('Deliverables', [
    formulaRule(R(dlv, 'A2:AD600'), '=$AC2="Overdue"', COLOUR.red, COLOUR.redFg),
    formulaRule(R(dlv, 'A2:AD600'), '=$AC2="Blocked"', COLOUR.purple, COLOUR.purpleFg),
    formulaRule(R(dlv, 'A2:AD600'), '=$AC2="Due soon"', COLOUR.amber, null),
    formulaRule(R(dlv, 'A2:AD600'), '=$AC2="Missing evidence"', COLOUR.amber, COLOUR.amberFg),
    textRule(R(dlv, 'P2:P600'), 'Delivered and verified', COLOUR.green, COLOUR.greenFg),
    formulaRule(R(dlv, 'A2:AD600'), '=OR($P2="No longer applicable",$P2="Replaced with approved equivalent")', COLOUR.grey, COLOUR.greyFg)
  ]);

  const opp = ss.getSheetByName('Opportunities');
  apply('Opportunities', [
    formulaRule(R(opp, 'A2:AD600'), '=$AB2=0', COLOUR.grey, COLOUR.greyFg),
    formulaRule(R(opp, 'S2:U600'), '=$AC2="Missing"', COLOUR.red, COLOUR.redFg),
    formulaRule(R(opp, 'R2:R600'), '=$AD2="Stale"', COLOUR.amber, COLOUR.amberFg),
    formulaRule(R(opp, 'U2:U600'), '=AND($U2<>"",$U2<TODAY(),$AB2=1)', COLOUR.red, COLOUR.redFg),
    textRule(R(opp, 'E2:E600'), 'Confirmed', COLOUR.green, COLOUR.greenFg),
    textRule(R(opp, 'E2:E600'), 'Closed won', COLOUR.green, COLOUR.greenFg),
    textRule(R(opp, 'V2:V600'), 'At risk', COLOUR.red, COLOUR.redFg),
    textRule(R(opp, 'V2:V600'), 'Strong', COLOUR.green, COLOUR.greenFg)
  ]);

  const ctb = ss.getSheetByName('Contributions');
  apply('Contributions', [
    textRule(R(ctb, 'X2:X600'), 'Overdue', COLOUR.red, COLOUR.redFg),
    textRule(R(ctb, 'X2:X600'), 'Received', COLOUR.green, COLOUR.greenFg),
    textRule(R(ctb, 'X2:X600'), 'Used', COLOUR.green, COLOUR.greenFg),
    textRule(R(ctb, 'X2:X600'), 'Unvalued', COLOUR.amber, COLOUR.amberFg),
    formulaRule(R(ctb, 'R2:R600'), '=$R2<>""', COLOUR.amber, COLOUR.amberFg)
  ]);

  const agr = ss.getSheetByName('Agreements');
  apply('Agreements', [
    textRule(R(agr, 'T2:T600'), 'Missing', COLOUR.red, COLOUR.redFg),
    textRule(R(agr, 'H2:H600'), 'Fully executed', COLOUR.green, COLOUR.greenFg),
    textRule(R(agr, 'H2:H600'), 'Confirmed in writing', COLOUR.green, COLOUR.greenFg),
    textRule(R(agr, 'H2:H600'), 'Sent (countersigned copy not filed)', COLOUR.amber, COLOUR.amberFg),
    formulaRule(R(agr, 'G2:G600'), '=AND($G2<>"",$G2<TODAY()+60)', COLOUR.amber, COLOUR.amberFg)
  ]);

  const com = ss.getSheetByName('Communications');
  apply('Communications', [
    formulaRule(R(com, 'A2:S600'), '=$S2="Yes"', COLOUR.amber, null),
    formulaRule(R(com, 'O2:O600'), '=AND($O2<>"",$O2<TODAY(),$R2<>"Closed")', COLOUR.red, COLOUR.redFg),
    textRule(R(com, 'R2:R600'), 'Closed', COLOUR.grey, COLOUR.greyFg)
  ]);

  const dec = ss.getSheetByName('Decisions');
  apply('Decisions', [
    textRule(R(dec, 'E2:E600'), 'Awaiting approval', COLOUR.amber, COLOUR.amberFg),
    textRule(R(dec, 'E2:E600'), 'Approved', COLOUR.green, COLOUR.greenFg),
    textRule(R(dec, 'E2:E600'), 'Rejected', COLOUR.red, COLOUR.redFg)
  ]);

  const tye = ss.getSheetByName('ThankYouEmail');
  apply('ThankYouEmail', [
    textRule(R(tye, 'F3:F600'), 'Approved', COLOUR.green, COLOUR.greenFg),
    textRule(R(tye, 'I3:I600'), 'Approved', COLOUR.green, COLOUR.greenFg),
    textRule(R(tye, 'F3:I600'), 'Hold', COLOUR.amber, COLOUR.amberFg),
    textRule(R(tye, 'F3:I600'), 'Exclude', COLOUR.grey, COLOUR.greyFg),
    textRule(R(tye, 'F3:I600'), 'Not reviewed', COLOUR.blue, COLOUR.blueFg)
  ]);

  const bra = ss.getSheetByName('BrandAssets');
  apply('BrandAssets', [
    formulaRule(R(bra, 'I2:I600'), '=LEFT($I2,7)="Pending"', COLOUR.amber, COLOUR.amberFg),
    textRule(R(bra, 'I2:I600'), 'Approved', COLOUR.green, COLOUR.greenFg),
    textRule(R(bra, 'I2:I600'), 'Expired', COLOUR.red, COLOUR.redFg)
  ]);

  const rpt = ss.getSheetByName('Reports');
  apply('Reports', [
    textRule(R(rpt, 'F2:F600'), 'Not started', COLOUR.grey, COLOUR.greyFg),
    textRule(R(rpt, 'F2:F600'), 'Sent', COLOUR.green, COLOUR.greenFg),
    formulaRule(R(rpt, 'I2:I600'), '=$I2>0', COLOUR.amber, COLOUR.amberFg)
  ]);

  const dq = ss.getSheetByName('Data Quality');
  apply('Data Quality', [formulaRule(R(dq, 'A5:C30'), '=$B5>0', COLOUR.amber, COLOUR.amberFg)]);

  // View tabs: colour flag words wherever they appear.
  VIEW_TABS.forEach(n => {
    if (n === 'Data Quality') return;
    const sh = ss.getSheetByName(n); if (!sh) return;
    const rng = sh.getRange(5, 1, 596, Math.max(sh.getLastColumn(), 1));
    sh.setConditionalFormatRules([
      textRule(rng, 'Overdue', COLOUR.red, COLOUR.redFg),
      textRule(rng, 'Due soon', COLOUR.amber, COLOUR.amberFg),
      textRule(rng, 'Blocked', COLOUR.purple, COLOUR.purpleFg),
      textRule(rng, 'Missing evidence', COLOUR.amber, COLOUR.amberFg),
      textRule(rng, 'Missing', COLOUR.red, COLOUR.redFg),
      textRule(rng, 'Stale', COLOUR.amber, COLOUR.amberFg),
      textRule(rng, 'At risk', COLOUR.red, COLOUR.redFg),
      textRule(rng, 'Awaiting approval', COLOUR.amber, COLOUR.amberFg),
      textRule(rng, 'Delivered and verified', COLOUR.green, COLOUR.greenFg)
    ]);
  });
}

function setHeaderNotes_(ss) {
  const note = (a1, text) => ss.getRange(a1).setNote(text);
  note('Deliverables!P1', 'Delivered and verified is refused unless column T holds an Evidence ID. Reported without evidence stays flagged.');
  note('Deliverables!T1', 'Evidence ID from the Evidence tab. Add the evidence row first, then link it here.');
  note('Deliverables!AC1', 'Computed flag: Overdue, Due soon (within WarningDays), Blocked, Missing evidence, No due date. Do not type here.');
  note('Opportunities!E1', 'Confirmed requires written confirmation on file. Closed won requires a signed agreement link in Agreements.');
  note('Opportunities!R1', 'Last meaningful contact. Prospects with no contact for StaleDays (Lists!B32) are flagged Stale.');
  note('Opportunities!S1', 'Every active relationship needs a next action, an owner (T) and a date (U). Missing ones show on Data Quality.');
  note('Contributions!H1', 'Confirmed value. Cash figures are entered from Wave; this workbook is not a ledger.');
  note('Contributions!I1', 'Received value per Wave (cash) or actual delivered value (in-kind).');
  note('Contributions!K1', 'Valuation basis is required for in-kind. Rows without one are counted as Unvalued, not summed.');
  note('Contributions!X1', 'Computed payment or in-kind status. Do not type here.');
  note('Agreements!K1', 'Drive link to the signed PDF or written confirmation. Never edit the signed document itself.');
  note('Decisions!E1', 'Set Approved, Rejected or Deferred here. Decided date and Decided by fill in automatically.');
  note('ThankYouEmail!I2', 'Final approval per recipient. Only rows marked Approved here, with Include? also Approved, are sent on 1 October at 09:00. Nothing is sent from this sheet.');
  note('Contacts!A1', 'RESTRICTED. Business contacts only. Do not share, export or paste into reports.');
  note('Organizations!A1', 'Partner ID is assigned automatically when you type in a new row.');
  note('ED Dashboard!H11', 'Pick a partner to see its status, money, deliverables and folder.');
  note('AuditLog!A1', 'Every edit on a record tab is logged here automatically (actor, record ID, old and new value).');
}

function groupHelperColumns_(ss) {
  [['Opportunities', 27, 4], ['Deliverables', 27, 4]].forEach(([name, start, count]) => {
    const sh = ss.getSheetByName(name);
    if (sh.getColumnGroupDepth(start) === 0) {
      sh.getRange(1, start, 1, count).shiftColumnGroupDepth(1);
      sh.collapseAllColumnGroups();
    }
  });
}

function protectTabs_(ss) {
  // Warning-only protection: editors see a confirmation before changing formula tabs.
  const protectSheet = (name, desc) => {
    const sh = ss.getSheetByName(name); if (!sh) return;
    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    sh.protect().setDescription(desc).setWarningOnly(true);
  };
  VIEW_TABS.forEach(n => protectSheet(n, 'View tab: formulas only. Edit the record tabs instead.'));
  protectSheet('ED Dashboard', 'Dashboard: formulas only (the partner lookup cell H11 is fine to change).');
  protectSheet('Lists', 'Dropdown values and config. Changing these changes the rules.');
  protectSheet('AuditLog', 'Append-only audit log.');
  protectSheet('Contacts', 'RESTRICTED: business contacts. Do not share or export.');
  // Helper formula columns on record tabs.
  const protectRange = (a1, desc) => {
    const r = ss.getRange(a1);
    r.getSheet().getProtections(SpreadsheetApp.ProtectionType.RANGE).filter(p => p.getDescription() === desc).forEach(p => p.remove());
    r.protect().setDescription(desc).setWarningOnly(true);
  };
  protectRange('Organizations!P2:R600', 'Computed columns');
  protectRange('Opportunities!AA2:AD600', 'Computed columns');
  protectRange('Opportunities!O2:O600', 'Computed columns');
  protectRange('Opportunities!Q2:Q600', 'Computed columns');
  protectRange('Agreements!T2:T600', 'Computed columns');
  protectRange('Contributions!W2:X600', 'Computed columns');
  protectRange('Deliverables!AA2:AD600', 'Computed columns');
  protectRange('Communications!S2:S600', 'Computed columns');
  protectRange('Events!J2:J600', 'Computed columns');
  protectRange('Reports!G2:M600', 'Computed columns');
  protectRange('Renewals!K2:K600', 'Computed columns');
  protectRange('ThankYouEmail!C3:C600', 'Computed columns');
}

function buildDashboardExtras_(ss) {
  const dash = ss.getSheetByName('ED Dashboard');
  dash.setColumnWidth(6, 12); dash.setColumnWidth(7, 210);
  // Row 5-13 column C: make the tab names clickable.
  for (let r = 5; r <= 13; r++) {
    const cell = dash.getRange(r, 3);
    const name = String(cell.getValue()).trim();
    const sh = ss.getSheetByName(name);
    if (sh) cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText('Open ' + name).setLinkUrl('#gid=' + sh.getSheetId()).build());
  }
  // Quick links.
  dash.getRange('G4:H4').setValues([['Quick links', '']]).setBackground(COLOUR.headerBg).setFontColor(COLOUR.headerFg).setFontWeight('bold');
  let r = 5;
  Object.keys(LINKS).forEach(k => {
    dash.getRange(r, 7).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(k).setLinkUrl(LINKS[k]).build());
    r++;
  });
  // Partner lookup.
  dash.getRange('G10:H10').setValues([['Partner lookup', '']]).setBackground(COLOUR.headerBg).setFontColor(COLOUR.headerFg).setFontWeight('bold');
  dash.getRange('G11').setValue('Partner');
  if (!dash.getRange('H11').getValue()) dash.getRange('H11').setValue('Banff & Lake Louise Tourism');
  dash.getRange('H11').setBackground(COLOUR.amber).setFontWeight('bold');
  const P = '$H$11';
  const rows = [
    ['Partner ID', `=IFERROR(INDEX(Organizations!$A$2:$A$600,MATCH(${P},Organizations!$C$2:$C$600,0)),"")`],
    ['Stage', `=IFERROR(INDEX(Opportunities!$E$2:$E$600,MATCH(${P},Opportunities!$C$2:$C$600,0)),"No opportunity row")`],
    ['Relationship health', `=IFERROR(INDEX(Opportunities!$V$2:$V$600,MATCH(${P},Opportunities!$C$2:$C$600,0)),"")`],
    ['Next action', `=IFERROR(INDEX(Opportunities!$S$2:$S$600,MATCH(${P},Opportunities!$C$2:$C$600,0)),"")`],
    ['Next-action owner and date', `=IFERROR(INDEX(Opportunities!$T$2:$T$600,MATCH(${P},Opportunities!$C$2:$C$600,0))&"  "&TEXT(INDEX(Opportunities!$U$2:$U$600,MATCH(${P},Opportunities!$C$2:$C$600,0)),"yyyy-mm-dd"),"")`],
    ['Agreement status', `=IFERROR(INDEX(Agreements!$H$2:$H$600,MATCH(${P},Agreements!$C$2:$C$600,0)),"No agreement")`],
    ['Cash confirmed / received', `=TEXT(SUMIFS(Contributions!$H$2:$H$600,Contributions!$C$2:$C$600,${P},Contributions!$D$2:$D$600,"Cash"),"$#,##0")&" / "&TEXT(SUMIFS(Contributions!$I$2:$I$600,Contributions!$C$2:$C$600,${P},Contributions!$D$2:$D$600,"Cash"),"$#,##0")`],
    ['In-kind confirmed / received', `=TEXT(SUMIFS(Contributions!$H$2:$H$600,Contributions!$C$2:$C$600,${P},Contributions!$D$2:$D$600,"In-kind"),"$#,##0")&" / "&TEXT(SUMIFS(Contributions!$I$2:$I$600,Contributions!$C$2:$C$600,${P},Contributions!$D$2:$D$600,"In-kind"),"$#,##0")`],
    ['Deliverables open / overdue', `=COUNTIFS(Deliverables!$C$2:$C$600,${P},Deliverables!$AA$2:$AA$600,1)&" / "&COUNTIFS(Deliverables!$C$2:$C$600,${P},Deliverables!$AC$2:$AC$600,"Overdue")`],
    ['Delivered and verified', `=COUNTIFS(Deliverables!$C$2:$C$600,${P},Deliverables!$P$2:$P$600,"Delivered and verified")&" of "&COUNTIF(Deliverables!$C$2:$C$600,${P})`],
    ['Communications needing reply', `=COUNTIFS(Communications!$C$2:$C$600,${P},Communications!$S$2:$S$600,"Yes")`],
    ['Brand asset status', `=IFERROR(INDEX(BrandAssets!$I$2:$I$600,MATCH(${P},BrandAssets!$C$2:$C$600,0)),"No row")`],
    ['Partner folder', `=IFERROR(HYPERLINK(INDEX(BrandAssets!$D$2:$D$600,MATCH(${P},BrandAssets!$C$2:$C$600,0)),"Open Brand Asset folder"),"No folder recorded")`]
  ];
  rows.forEach((row, i) => { dash.getRange(12 + i, 7).setValue(row[0]); dash.getRange(12 + i, 8).setFormula(row[1]); });
  dash.getRange('G12:G24').setFontColor(COLOUR.greyFg);
  dash.getRange('H12:H24').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setFontWeight('bold');
  dash.setColumnWidth(8, 260);
}

function linkBrandFolders_(ss) {
  const sh = ss.getSheetByName('BrandAssets');
  const last = sh.getLastRow(); if (last < 2) return;
  const ids = sh.getRange(2, 2, last - 1, 1).getValues();
  const links = sh.getRange(2, 4, last - 1, 1).getValues();
  const out = ids.map((r, i) => {
    const id = String(r[0]).trim();
    if (links[i][0]) return [links[i][0]];
    return [PARTNER_FOLDERS[id] ? 'https://drive.google.com/drive/folders/' + PARTNER_FOLDERS[id] : ''];
  });
  sh.getRange(2, 4, last - 1, 1).setValues(out);
  sh.getRange('D1').setNote('Brand Asset Library folder for this partner. New partners: copy _TEMPLATE, rename [ORG-nnn] Partner, paste the link here. Template: ' + TEMPLATE_FOLDER);
}

// ------------------------------------------------------------- automation

function onEdit(e) {
  if (!e || !e.range) return;
  const ss = e.source;
  const sh = e.range.getSheet();
  const name = sh.getName();
  if (RECORD_TABS.indexOf(name) < 0) return;
  const hdr = headerRow(name);
  const row = e.range.getRow(), col = e.range.getColumn();
  if (row <= hdr) return;
  const single = e.range.getNumRows() === 1 && e.range.getNumColumns() === 1;
  const today = new Date();

  // 1. Assign an ID to a new row (column A empty, something else typed).
  if (single && col !== 1 && ID_PREFIX[name]) {
    const idCell = sh.getRange(row, 1);
    if (!idCell.getValue() && e.value !== undefined && e.value !== '') {
      idCell.setValue(nextId_(sh, ID_PREFIX[name], hdr));
      if (name === 'Organizations') { sh.getRange(row, 14).setValue(today); }
      if (name === 'Decisions' && !sh.getRange(row, 2).getValue()) { sh.getRange(row, 2).setValue(today); sh.getRange(row, 5).setValue('Awaiting approval'); }
      if (name === 'Deliverables' && !sh.getRange(row, 16).getValue()) sh.getRange(row, 16).setValue('Not started');
      if (name === 'Opportunities' && !sh.getRange(row, 5).getValue()) sh.getRange(row, 5).setValue('Prospect identified');
    }
  }

  // 2. Partner ID typed: fill the partner name.
  const idColumn = name === 'ThankYouEmail' ? 1 : 2;
  const nameColumn = name === 'ThankYouEmail' ? 2 : 3;
  if (single && col === idColumn && (PARTNER_NAME_TABS.indexOf(name) >= 0 || name === 'ThankYouEmail')) {
    const nameCell = sh.getRange(row, nameColumn);
    const pname = partnerName_(ss, e.value);
    if (pname && !nameCell.getValue()) nameCell.setValue(pname);
  }

  // 3. Deliverables: evidence rule and delivered date.
  if (single && name === 'Deliverables' && col === 16) {
    if (e.value === 'Delivered and verified') {
      const evidence = sh.getRange(row, 20).getValue();
      const required = sh.getRange(row, 19).getValue();
      if (required === 'Yes' && !evidence) {
        e.range.setValue(e.oldValue || 'Reported without evidence');
        ss.toast('Add an Evidence row and put its ID in column T first. Status set to Reported without evidence.', 'Evidence required', 8);
        return;
      }
      if (!sh.getRange(row, 17).getValue()) sh.getRange(row, 17).setValue(today);
    }
  }

  // 4. Decisions: stamp decided date and decider.
  if (single && name === 'Decisions' && col === 5 && ['Approved', 'Rejected', 'Deferred'].indexOf(e.value) >= 0) {
    if (!sh.getRange(row, 8).getValue()) sh.getRange(row, 8).setValue(today);
    if (!sh.getRange(row, 6).getValue()) sh.getRange(row, 6).setValue(actor_());
  }

  // 5. ThankYouEmail: final approval needs Include? = Approved; stamp approver.
  if (single && name === 'ThankYouEmail' && col === 9 && e.value === 'Approved') {
    if (sh.getRange(row, 6).getValue() !== 'Approved') {
      e.range.setValue(e.oldValue || 'Not reviewed');
      ss.toast('Set Include? to Approved (recipient review) before Final approval.', 'Approval order', 8);
      return;
    }
    sh.getRange(row, 10).setValue(actor_());
    sh.getRange(row, 11).setValue(today);
  }

  // 6. Updated stamps.
  if (name === 'Organizations') sh.getRange(row, 15, e.range.getNumRows(), 1).setValue(today);
  if (name === 'Contacts') sh.getRange(row, 10, e.range.getNumRows(), 1).setValue(today);

  // 7. Audit trail.
  const recordId = sh.getRange(row, 1).getValue() || '(new row)';
  const header = sh.getRange(hdr, col).getValue();
  let detail;
  if (single) detail = `${name}!${e.range.getA1Notation()} [${header}]: "${trim_(e.oldValue)}" -> "${trim_(e.value)}"`;
  else detail = `${name}!${e.range.getA1Notation()}: ${e.range.getNumRows()} x ${e.range.getNumColumns()} cells changed`;
  if (name === 'Contacts') detail = `Contacts!${e.range.getA1Notation()} [${header}] changed`; // never log contact values
  appendAudit_(ss, actor_(), 'Edit', String(recordId), detail, 'onEdit');
}

function nextId_(sh, prefix, hdr) {
  const vals = sh.getRange(hdr + 1, 1, Math.max(sh.getLastRow() - hdr, 1), 1).getValues();
  let max = 0;
  vals.forEach(r => { const m = String(r[0]).match(new RegExp('^' + prefix + '-(\\d+)$')); if (m) max = Math.max(max, parseInt(m[1], 10)); });
  return prefix + '-' + String(max + 1).padStart(3, '0');
}

function partnerName_(ss, id) {
  if (!id) return '';
  const org = ss.getSheetByName('Organizations');
  const data = org.getRange(2, 1, Math.max(org.getLastRow() - 1, 1), 3).getValues();
  for (const r of data) if (r[0] === id) return r[2];
  return '';
}

function actor_() {
  try { const em = Session.getActiveUser().getEmail(); if (em) return em; } catch (err) { /* simple trigger without identity */ }
  return 'Sheet user';
}

function trim_(v) { const s = v === undefined || v === null ? '' : String(v); return s.length > 120 ? s.slice(0, 117) + '...' : s; }

function appendAudit_(ss, actor, action, ids, detail, source) {
  const log = ss.getSheetByName('AuditLog'); if (!log) return;
  log.appendRow([new Date(), actor, action, ids, detail, source]);
}

// ------------------------------------------------------------ menu actions

function addNewPartner() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const res = ui.prompt('Add a new partner', 'Approved public name of the organization:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK || !res.getResponseText().trim()) return;
  const pname = res.getResponseText().trim();
  const org = ss.getSheetByName('Organizations');
  const existing = org.getRange(2, 3, Math.max(org.getLastRow() - 1, 1), 1).getValues().map(r => String(r[0]).toLowerCase());
  if (existing.indexOf(pname.toLowerCase()) >= 0) { ui.alert('That organization already exists. Use the Directory to find it.'); return; }
  const orgId = nextId_(org, 'ORG', 1);
  const orgRow = org.getLastRow() + 1;
  org.getRange(orgRow, 1, 1, 3).setValues([[orgId, pname, pname]]);
  org.getRange(orgRow, 10).setValue('Jeffrey');
  org.getRange(orgRow, 12).setValue('Added via menu');
  org.getRange(orgRow, 14, 1, 2).setValues([[new Date(), new Date()]]);
  const opp = ss.getSheetByName('Opportunities');
  const oppId = nextId_(opp, 'OPP', 1);
  const oppRow = opp.getLastRow() + 1;
  const next = new Date(); next.setDate(next.getDate() + 7);
  opp.getRange(oppRow, 1, 1, 6).setValues([[oppId, orgId, pname, 2026, 'Prospect identified', 'Priority 3']]);
  opp.getRange(oppRow, 19, 1, 3).setValues([['Research and decide on approach', 'Jeffrey', next]]);
  opp.getRange(oppRow, 25).setValue('C');
  appendAudit_(ss, actor_(), 'Partner added', orgId + '; ' + oppId, pname + ' added with a Prospect identified opportunity and a 7-day next action.', 'Menu');
  ss.setActiveSheet(opp); opp.setActiveRange(opp.getRange(oppRow, 19));
  ui.alert('Added ' + pname + ' as ' + orgId + '. Fill in the next action, owner and date on the Opportunities row now selected.');
}

function verifyWorkbook() {
  const ss = SpreadsheetApp.getActive();
  const problems = [];
  const count = (name) => { const sh = ss.getSheetByName(name); return sh.getRange(headerRow(name) + 1, 1, 600, 1).getValues().filter(r => r[0] !== '').length; };
  const dash = ss.getSheetByName('ED Dashboard');
  if (dash.getRange('B42').getValue() !== count('Organizations')) problems.push('Dashboard organization self-check does not match the Organizations tab.');
  if (dash.getRange('B43').getValue() !== count('Deliverables')) problems.push('Dashboard deliverable self-check does not match the Deliverables tab.');
  VIEW_TABS.concat(['ED Dashboard']).forEach(n => {
    const sh = ss.getSheetByName(n); if (!sh) { problems.push('Missing tab ' + n); return; }
    const vals = sh.getDataRange().getDisplayValues();
    vals.forEach((row, i) => row.forEach((v, j) => {
      if (/^#(REF!|NAME\?|VALUE!|N\/A|ERROR!|DIV\/0!)$/.test(v)) problems.push(`${n}!${sh.getRange(i + 1, j + 1).getA1Notation()} shows ${v}`);
    }));
  });
  // Duplicate IDs.
  RECORD_TABS.forEach(n => {
    const sh = ss.getSheetByName(n);
    const ids = sh.getRange(headerRow(n) + 1, 1, 600, 1).getValues().map(r => r[0]).filter(v => v !== '');
    const seen = {}; ids.forEach(v => { if (seen[v]) problems.push(`${n}: duplicate ID ${v}`); seen[v] = true; });
  });
  const msg = problems.length ? problems.join('\n') : 'All checks passed: self-check counts match, no formula errors, no duplicate IDs.';
  appendAudit_(ss, actor_(), 'Verify', 'all', msg.slice(0, 500), 'Menu');
  SpreadsheetApp.getUi().alert('Verify workbook', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function logExceptionSnapshot() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName('ED Dashboard');
  const labels = dash.getRange('A5:A13').getValues().map(r => r[0]);
  const counts = dash.getRange('B5:B13').getDisplayValues().map(r => r[0]);
  const money = dash.getRange('D5:E13').getDisplayValues();
  const detail = labels.map((l, i) => `${l}: ${counts[i]}`).join('; ') + ' | ' + money.map(r => `${r[0]}: ${r[1]}`).join('; ');
  appendAudit_(ss, actor_(), 'Exception snapshot', 'dashboard', detail.slice(0, 1500), 'Snapshot');
  try { ss.toast('Snapshot written to AuditLog.', 'Banff Pride', 5); } catch (err) { /* no UI when run by trigger */ }
}

function scheduleWeeklySnapshot() {
  removeWeeklySnapshot();
  ScriptApp.newTrigger('logExceptionSnapshot').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).inTimezone('America/Edmonton').create();
  SpreadsheetApp.getUi().alert('A Monday 07:00 (Edmonton) snapshot of the dashboard exceptions will be appended to AuditLog. Remove it from the same menu.');
}

function removeWeeklySnapshot() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'logExceptionSnapshot').forEach(t => ScriptApp.deleteTrigger(t));
}
