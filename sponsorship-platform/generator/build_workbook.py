"""Build the Banff Pride Partnerships workbook (.xlsx, Google Sheets compatible).

Formulas use only functions shared by Excel 365 and Google Sheets: FILTER, SORT, INDEX,
COUNTIFS, SUMIFS, IFERROR, IF, TODAY, TEXT, HYPERLINK. Views read helper flag columns
so every FILTER has a single boolean condition.
"""
import json, re, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
import seed as S

MAXR = 600   # formula range limit for record tabs
EXT = 12     # blank rows pre-filled with row formulas below the seed data
D = dt.date
wb = Workbook(); wb.remove(wb.active)

HDR_FILL = PatternFill("solid", fgColor="333F27"); HDR_FONT = Font(bold=True, color="FEFAF4")
VIEW_FILL = PatternFill("solid", fgColor="88A3AA"); TITLE_FONT = Font(bold=True, size=14, color="333F27")
NOTE_FONT = Font(italic=True, color="666666"); BOLD = Font(bold=True)
RED = PatternFill("solid", fgColor="F8CBAD"); AMBER = PatternFill("solid", fgColor="FFF2CC"); GREY = PatternFill("solid", fgColor="E7E6E6")
thin = Side(style="thin", color="CCCCCC")

def sheet(name, headers, widths=None, color=None, freeze="C2", note=None):
    ws = wb.create_sheet(name)
    r = 1
    if note:
        ws.cell(1, 1, note).font = NOTE_FONT; r = 2
    for i, h in enumerate(headers, 1):
        c = ws.cell(r, i, h); c.fill = HDR_FILL; c.font = HDR_FONT; c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = ws.cell(r + 1, 3 if freeze == "C2" else 1).coordinate if freeze else None
    for i, w in enumerate(widths or [], 1):
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[r].height = 32
    if color: ws.sheet_properties.tabColor = color
    ws.auto_filter.ref = f"A{r}:{L(len(headers))}{r}"
    ws._hdr_row = r
    return ws

def put(ws, row, values, fmt=None):
    for i, v in enumerate(values, 1):
        c = ws.cell(row, i, v)
        if isinstance(v, (dt.date, dt.datetime)): c.number_format = "yyyy-mm-dd"
        if fmt and i in fmt: c.number_format = fmt[i]
        if isinstance(v, str) and len(v) > 60: c.alignment = Alignment(wrap_text=True, vertical="top")

def dv(ws, col_letter, list_ref, first=2, last=MAXR):
    d = DataValidation(type="list", formula1=list_ref, allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(d); d.add(f"{col_letter}{first}:{col_letter}{last}")

# ------------------------------------------------------------------ Lists
LISTS = {
 "Relationship type": ["Cash sponsor","In-kind sponsor","Venue sponsor or partner","Hotel and accommodation sponsor","Liquor, brewery, distillery or beverage sponsor","Community partner","Other non-profit","Vendor or supplier","Transportation sponsor or partner","Attraction or experience sponsor","Tourism or government partner","Media partner","Funder"],
 "Opportunity stage": ["Prospect identified","Researched","Outreach prepared","Contacted","Conversation active","Proposal in development","Proposal sent","Negotiation","Verbal commitment","Agreement pending","Confirmed","Closed won","Closed declined","Closed lost","Dormant","Do not contact"],
 "Agreement status": ["Draft","Sent (countersigned copy not filed)","Signed by Banff Pride","Signed by partner (Banff Pride copy to file)","Fully executed","Confirmed in writing","Amended","Expired","Cancelled"],
 "Deliverable status": ["Not started","In progress","Blocked by partner","Blocked internally","Delivered and verified","Partially delivered","Reported without evidence","Not delivered","Replaced with approved equivalent","No longer applicable"],
 "Owed by": ["Banff Pride","Partner"],
 "Yes/No": ["Yes","No"],
 "Priority": ["Priority 1","Priority 2","Priority 3","Priority 4","Priority 5"],
 "Likelihood": ["High","Medium-High","Medium","Low-Medium","Low"],
 "Relationship health": ["Strong","Stable","At risk","Unknown"],
 "Contribution kind": ["Cash","In-kind"],
 "Contribution category": ["Sponsorship fee","Venue / space","Accommodation","Food & beverage","Marketing / media","Production / AV","Transportation","Prizes / auction items","Professional services","Equipment / rentals","Printing","Fee waivers","Activation","Merchandise revenue share","Other"],
 "Confidence": ["A","B","C"],
 "Channel": ["Email","Meeting","Phone","Calendar","Document","Decision","Social","Website","Signage / print","Event","Report","Assets","Payment","Delivery","Volunteers","Compliance","Other"],
 "Triage": ["Open","Waiting","Closed"],
 "Owner": ["Jeffrey","James (BVPN)","Marketing Coordinator","Volunteer Coordinator","Board","Partner"],
 "Decision status": ["Awaiting approval","Approved","Rejected","Deferred","Recorded"],
 "Verification": ["Verified","Pending","Rejected"],
 "Report status": ["Not started","Assembling","Draft","Approved","Sent"],
 "Renewal status": ["Not started","In progress","Scheduled","Renewed","Declined"],
 "Email approval": ["Not reviewed","Approved","Hold","Exclude"],
}
lists_ws = sheet("Lists", list(LISTS.keys()), [26]*len(LISTS), "7F7F7F", freeze=None)
LISTREF = {}
for i, (k, vals) in enumerate(LISTS.items(), 1):
    for r, v in enumerate(vals, 2): lists_ws.cell(r, i, v)
    LISTREF[k] = f"Lists!${L(i)}$2:${L(i)}${len(vals)+1}"
lists_ws.cell(30, 1, "Config").font = BOLD
lists_ws.cell(31, 1, "Deadline warning days"); lists_ws.cell(31, 2, 14)
lists_ws.cell(32, 1, "Stale prospect days"); lists_ws.cell(32, 2, 60)
lists_ws.cell(33, 1, "Festival start"); lists_ws.cell(33, 2, D(2026,10,2)).number_format = "yyyy-mm-dd"
lists_ws.cell(34, 1, "Festival end"); lists_ws.cell(34, 2, D(2026,10,12)).number_format = "yyyy-mm-dd"
lists_ws.cell(35, 1, "Tiers (2026 template)"); lists_ws.cell(35, 2, "Trailhead $500; Silver $1,500; Alpine $2,500; Summit $5,000; Presenting $10,000; Custom")

# ------------------------------------------------------------------ Organizations
july = json.load(open("july.json"))
def norm(s): return re.sub(r"[^a-z0-9]", "", s.lower())
ALIAS = {"banffhospitalitycollective":"BHC","banffavebrewingco":"BABREW","banfflodgingcompany":"BLC","fairmontbanffsprings":"FAIRMONT","pursuitbanffjaspercollection":"PURSUIT","moxybanff":"MOXY","atbfinancial":"ATB","banffcentreforartsandcreativity":"BANFFCENTRE","banfflakelouisetourism":"BLLT","townofbanff":"TOB","skibig3":"SKIBIG3","flowstatebanff":"FLOW","beatniksalon":"BEATNIK","chillaudiovisualsolutions":"CHILLAV","valleysignanddesign":"VALLEYSIGN","proimagesignsprinting":"PROIMAGE","whytemuseumofthecanadianrockies":"WHYTE","banffpubliclibrary":"LIBRARY","melissasrestaurantbar":"MELS","roseandcrownpub":"ROSECROWN","pumptaptavern":"PUMPTAP","sheepdogbrewing":"SHEEPDOG","laterrazza":"LATERRAZZA","brazen":"BRAZEN","farmfire":"FARMFIRE","thebosskitchenbar":"BOSS","maclabbistroatbanffcentreforartsandcreativity":"MACLAB","banffairporter":"AIRPORTER","basecampsuitesbanff":"BASECAMP","luxcinema":"LUX","roampublictransit":"ROAM","morainelakebuscompany":"MLB"}
ASK2TYPE = [("room","Hotel and accommodation sponsor"),("bed","Hotel and accommodation sponsor"),("transport","Transportation sponsor or partner"),("shuttle","Transportation sponsor or partner"),("venue","Venue sponsor or partner"),("taste","Venue sponsor or partner"),("f&b","In-kind sponsor"),("beer","Liquor, brewery, distillery or beverage sponsor"),("media","Media partner"),("print","In-kind sponsor"),("in-kind","In-kind sponsor"),("cash","Cash sponsor")]
def types_from_ask(ask):
    t = []
    for k, v in ASK2TYPE:
        if k in ask.lower() and v not in t: t.append(v)
    return ";".join(t) or "Cash sponsor"

orgs = []   # dict rows
key2id = {}
for i, o in enumerate(S.ORGS, 1):
    pid = f"ORG-{i:03d}"; key2id[o[0]] = pid
    orgs.append(dict(id=pid, key=o[0], legal=o[1], public=o[2], types=o[3], sector=o[4], sub=o[5], loc=o[6], group=o[7], web=o[8], owner=o[9], notes=o[10], legacy=o[11], source="Gate 1 audit (agreements, Gmail, Granola)", july=None))
seen = {norm(o["legal"]) for o in orgs} | {norm(o["public"]) for o in orgs}
for row in july["main"]:
    n = norm(row["Business"])
    if n in ALIAS or n in seen:
        k = ALIAS.get(n)
        if k and not k.endswith("_G"):
            for o in orgs:
                if o["key"] == k: o["july"] = row
        elif n in seen:
            for o in orgs:
                if norm(o["legal"]) == n or norm(o["public"]) == n: o["july"] = row
        continue
    pid = f"ORG-{len(orgs)+1:03d}"; seen.add(n)
    orgs.append(dict(id=pid, key="J"+row["Rank"], legal=row["Business"], public=row["Business"], types=types_from_ask(row["Best-fit ask"]), sector=row["Sector"], sub=row["Sub-category"], loc=row["Location"], group=row["Ownership group"], web=row["Website / directory profile"], owner="Jeffrey", notes=row["Rationale / notes"], legacy="", source=f"July 2026 prospect sheet, rank {row['Rank']}", july=row))
venue_only = []
for v in july["venues"]:
    n = norm(v["Venue"])
    if n in seen or n in ALIAS:
        continue
    if not v["Status"]: continue
    pid = f"ORG-{len(orgs)+1:03d}"; seen.add(n)
    orgs.append(dict(id=pid, key="V", legal=v["Venue"], public=v["Venue"], types="Venue sponsor or partner", sector="Restaurants / bars", sub="", loc="Banff", group="", web="", owner="Jeffrey", notes=f"July venue list: status {v['Status']}; Taste The Rainbow {v['Taste The Rainbow']}; Lavender Lounge {v['Lavender Lounge']}", legacy="", source="July 2026 prospect sheet, venue table", july={"Status": v["Status"], "Recommended priority":"", "Research score":"", "Sponsorship likelihood":"", "Evidence basis":"", "Approach level":"Business", "Best-fit ask":"Taste The Rainbow / Lavender Lounge", "Rationale / notes":"", "Last contact":"", "Source URL":""}))
ORGN = {o["key"]: o for o in orgs}
def oid(key): return key2id[key] if key in key2id else ""
def oname(key): return ORGN[key]["public"] if key in ORGN else ""

ws = sheet("Organizations", ["Partner ID","Legal name","Approved public name","Relationship types","Sector","Sub-category","Location","Ownership group","Website","Relationship owner","Legacy ID","Prospect source","Notes","Created","Updated","Open deliverables","Confirmed cash","Confirmed in-kind"], [11,32,28,34,20,20,20,22,28,12,22,30,50,11,11,10,12,12], "1F6E43")
for r, o in enumerate(orgs, 2):
    put(ws, r, [o["id"], o["legal"], o["public"], o["types"], o["sector"], o["sub"], o["loc"], o["group"], o["web"], o["owner"], o["legacy"], o["source"], o["notes"], S.TODAY, S.TODAY,
        f'=COUNTIFS(Deliverables!$B$2:$B${MAXR},A{r},Deliverables!$AA$2:$AA${MAXR},1)',
        f'=SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$B$2:$B${MAXR},A{r},Contributions!$D$2:$D${MAXR},"Cash")',
        f'=SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$B$2:$B${MAXR},A{r},Contributions!$D$2:$D${MAXR},"In-kind")'], {17:"$#,##0", 18:"$#,##0"})
dv(ws, "J", LISTREF["Owner"])
ORG_LAST = len(orgs) + 1

# ------------------------------------------------------------------ Contacts (restricted)
ws = sheet("Contacts", ["Contact ID","Partner ID","Partner","Name","Role","Email","Phone","Preferred channel","Notes","Updated"], [10,11,28,22,30,30,14,12,50,11], "1F6E43", note="RESTRICTED: business contacts only. Do not share this tab or export it. Names of individuals never appear in reports.")
for i, c in enumerate(S.CONTACTS, 1):
    put(ws, i + 2, [f"CON-{i:03d}", oid(c[0]), oname(c[0]), c[1], c[2], c[3], c[4], c[5], c[6], S.TODAY])

# ------------------------------------------------------------------ Opportunities
STAGE_MAP = {"Confirmed":"Confirmed","In discussion":"Conversation active","Contacted":"Contacted","Declined":"Closed declined","Do not contact":"Do not contact","":"Prospect identified"}
CURATED = {  # key: (stage, priority, next action, next date, health, risks, last contact)
 "BLLT":("Closed won","Priority 1","Reply to the BLLT invoicing contact with finalized plans so Invoice #2 is processed",D(2026,9,8),"At risk","Invoice #2 ($9,500) on hold; Sept 8 Roam artwork and Sept 17 print need BLLT logo approval lead times",D(2026,9,2)),
 "SKIBIG3":("Closed won","Priority 1","Schedule the Adventure Hub Lavender Lounge session; confirm VIP ticket quantity",D(2026,9,20),"Strong","Sticker designs pending; exclusivity window Sept 21 to Oct 12",D(2026,8,28)),
 "BHC":("Confirmed","Priority 1","Send invoice and request countersigned agreement and logos",D(2026,9,12),"Stable","Countersigned copy and assets not filed",D(2026,8,14)),
 "QTIL":("Closed won","Priority 3","Send Pride Market booth details",D(2026,9,20),"Strong","",D(2026,8,31)),
 "HIGHLINE":("Confirmed","Priority 3","Post logo; watch for payment around Sept 15; request countersigned copy",D(2026,9,15),"Stable","",D(2026,8,28)),
 "ROAM":("Closed won","Priority 2","Send bus panel artwork by noon Sept 8",D(2026,9,8),"Stable","Hard artwork deadline",D(2026,8,17)),
 "ATB":("Confirmed","Priority 1","Chase logo files; add ATB to website and market vendor list",D(2026,9,10),"Strong","Fab Fiver Bucks mechanics unconfirmed",D(2026,8,28)),
 "MLB":("Confirmed","Priority 2","Verify website listing; send talking points; spotlight by Sept 11",D(2026,9,11),"Stable","Code must not be displayed publicly",D(2026,8,20)),
 "RMR":("Confirmed","Priority 3","Set raffle mechanics and recognition",D(2026,9,25),"Stable","Valuation basis is donor's 'up to'",D(2026,8,19)),
 "BYHENDO":("Verbal commitment","Priority 1","Confirm pricing and obtain written 2026 terms",D(2026,9,12),"Stable","No written record for 2026",D(2026,9,5)),
 "MOXY":("Confirmed","Priority 1","Finalize Drag Brunch details (Sept 7) and record written confirmation",D(2026,9,7),"Stable","",D(2026,9,4)),
 "WHYTE":("Closed won","Priority 2","Await floor plan and stage-piece confirmation",D(2026,9,10),"Strong","Carpet protection; no pipe and drape",D(2026,9,3)),
 "TOB":("Confirmed","Priority 1","File written basis for in-kind support",D(2026,9,20),"Stable","",D(2026,9,5)),
 "AIRPORTER":("Confirmed","Priority 2","Promote code; request redemption report after Oct 14",D(2026,10,15),"Stable","",D(2026,9,3)),
 "BASECAMP":("Confirmed","Priority 2","Agree redemption reporting",D(2026,9,11),"Stable","",D(2026,8,11)),
 "BLC":("Closed declined","Priority 2","Reply on performer room block (two unanswered follow-ups)",D(2026,9,8),"At risk","Silence risks the relationship; accommodation partner only",D(2026,8,31)),
 "MTNEVENTS":("Closed won","Priority 3","Send venue map; confirm pickup window",D(2026,9,10),"Strong","",D(2026,8,31)),
 "BELL":("Closed won","Priority 2","Send CTV Our Community assets",D(2026,9,9),"Stable","",D(2026,9,4)),
 "BVI":("Closed won","Priority 3","Send ticket link and Taste The Rainbow list; approve feature 2",D(2026,9,9),"Stable","Website inconsistencies flagged",D(2026,9,4)),
 "FAIRMONT":("Conversation active","Priority 1","Connect Fairmont with producers for the Rundle Lounge idea",D(2026,9,12),"Unknown","Tentative; no cash discussed",D(2026,7,31)),
 "PURSUIT":("Proposal sent","Priority 1","Forward Biosphere overview; chase the sponsorship decision",D(2026,9,8),"At risk","No decision since June; budget assumes $16,000",D(2026,9,3)),
 "FLOW":("Confirmed","Priority 3","Get registration link for Night Bloom",D(2026,9,15),"Stable","",D(2026,8,6)),
 "NIBBLE":("Conversation active","Priority 3","Schedule a call or food tour",D(2026,9,10),"Unknown","",D(2026,9,3)),
 "PROIMAGE":("Conversation active","Priority 3","Send print needs list",D(2026,9,9),"Unknown","",D(2026,9,4)),
 "WILDWATER":("Closed declined","Priority 3","None; thank and revisit 2027",None,"Unknown","",D(2026,9,4)),
 "ALPINEHELI":("Closed declined","Priority 3","None; thank and revisit 2027",None,"Unknown","",D(2026,9,4)),
 "BURNCO":("Contacted","Priority 3","Follow up late October",D(2026,10,26),"Unknown","",D(2026,9,4)),
 "REDBULL":("Outreach prepared","Priority 2","Festival sponsorship approach (Todoist)",D(2026,9,8),"Unknown","",None),
 "SUNCOR":("Researched","Priority 3","Research contact, alignment angle, target ask",D(2026,9,30),"Unknown","",None),
 "WALMART":("Researched","Priority 3","Research contact, alignment angle, target ask",D(2026,9,30),"Unknown","",None),
 "CHILLAV":("Confirmed","Priority 3","File written terms for in-kind A/V",D(2026,9,20),"Unknown","",None),
 "VALLEYSIGN":("Conversation active","Priority 3","Confirm 2026 in-kind print",D(2026,9,12),"Unknown","",None),
 "CANMOREPRIDE":("Closed won","Priority 3","Coordinate Roam and Q-Tilities joint arrangements",None,"Strong","",D(2026,8,17)),
 "SHEEPDOG":("Contacted","Priority 3","Follow up on Taste The Rainbow and Lavender Lounge",D(2026,9,10),"Unknown","",D(2026,8,7)),
 "LUX":("Confirmed","Priority 3","Finalize Rocky Horror date",D(2026,9,15),"Stable","",None),
 "BIOSPHERE":("Confirmed","Priority 3","Forward overview to Pursuit",D(2026,9,8),"Stable","",D(2026,9,3)),
 "WILDLIFE":("Prospect identified","Priority 2","Renewal outreach for 2026 in-kind product",D(2026,9,12),"Unknown","2025 partner not yet contacted",None),
 "BABREW":("Conversation active","Priority 1","Confirm in-kind venue and beer; Taste The Rainbow and Lavender Lounge",D(2026,9,10),"Unknown","",D(2026,8,7)),
 "BANFFCENTRE":("Conversation active","Priority 1","Follow the Sept 8 chat (Black Queer Joy pop-up); confirm venue role",D(2026,9,12),"Unknown","",D(2026,8,7)),
 "SIZZLERS":("Confirmed","Priority 2","Confirm Wild Things run-of-show and Byhendo pour",D(2026,9,20),"Stable","",D(2026,9,3)),
 "ROSECROWN":("Confirmed","Priority 2","Confirm Sunday Funday and Byhendo tasting",D(2026,9,20),"Stable","",D(2026,9,3)),
 "MELS":("Confirmed","Priority 2","Confirm Alpenglow set-up and Byhendo banner placement",D(2026,9,20),"Stable","",D(2026,9,3)),
 "PUMPTAP":("Confirmed","Priority 3","Confirm Punk Night details",D(2026,9,20),"Stable","",None),
 "LIBRARY":("Confirmed","Priority 3","Confirm Reading with Royalty logistics",D(2026,9,20),"Stable","",None),
 "BEATNIK":("Confirmed","Priority 3","Confirm Gender Free Haircut Club details",D(2026,9,20),"Stable","",None),
 "FRANKIED":("Confirmed","Priority 3","Confirm Donuts & Conversation",D(2026,9,20),"Stable","",None),
 "DETOURS":("Closed won","Priority 2","Confirm Alpenglow DJ lineup and fees",D(2026,9,15),"Stable","",D(2026,4,9)),
 "CSERIES":("Confirmed","Priority 3","Confirm Sunday Funday",D(2026,9,20),"Stable","",None),
 "LATERRAZZA":("Confirmed","Priority 3","Taste The Rainbow listing",D(2026,9,20),"Stable","",None),
 "BRAZEN":("Confirmed","Priority 3","Taste The Rainbow listing",D(2026,9,20),"Stable","",None),
 "FARMFIRE":("Confirmed","Priority 3","Taste The Rainbow listing",D(2026,9,20),"Stable","",None),
 "BOSS":("Confirmed","Priority 3","Taste The Rainbow listing",D(2026,9,20),"Stable","",None),
 "FATOX":("Confirmed","Priority 3","Taste The Rainbow listing",D(2026,9,20),"Stable","",None),
 "MACLAB":("Confirmed","Priority 3","Taste The Rainbow and Lavender Lounge listing",D(2026,9,20),"Stable","",None),
 "BVPN":("Closed won","Priority 5","n/a",None,"Strong","",None),
}
PROB = {"High":0.5,"Medium-High":0.35,"Medium":0.2,"Low-Medium":0.1,"Low":0.05}
OPP_HDR = ["Opportunity ID","Partner ID","Partner","Cycle year","Stage","Priority","Research score","Likelihood","Evidence basis","Approach level","Best-fit ask","Rationale / alignment","Proposed cash","Proposed in-kind","Expected total","Probability","Weighted value","Last meaningful contact","Next action","Next-action owner","Next-action date","Relationship health","Risks","Source link","Confidence","Verify","(reserved)","Active?","Missing owner or action","Stale?"]
ws = sheet("Opportunities", OPP_HDR, [11,11,28,8,20,10,8,12,12,12,26,40,11,11,11,9,11,12,44,12,12,12,36,30,9,8,9,7,10,7], "1F6E43")
r = 2
for o in orgs:
    j = o["july"] or {}
    c = CURATED.get(o["key"])
    stage = c[0] if c else STAGE_MAP.get(j.get("Status",""), "Prospect identified")
    prio = c[1] if c else j.get("Recommended priority","Priority 4")
    nxt = c[2] if c else ("Prepare tailored ask" if stage in ("Contacted","Conversation active") else "Review and decide whether to approach for 2027")
    nd = c[3] if c else None
    health = c[4] if c else "Unknown"
    risks = c[5] if c else ""
    lc = c[6] if c else (dt.datetime.strptime(j["Last contact"], "%m/%d/%Y").date() if j.get("Last contact") else None)
    if lc is None and j.get("Status") == "Contacted": lc = D(2026,9,3)
    lk = j.get("Sponsorship likelihood","")
    src = j.get("Source URL","") or o["source"]
    put(ws, r, [f"OPP-{r-1:03d}", o["id"], o["public"], 2026, stage, prio, int(j["Research score"]) if j.get("Research score") else None, lk, j.get("Evidence basis",""), j.get("Approach level",""), j.get("Best-fit ask",""), j.get("Rationale / notes","") or o["notes"], None, None,
        f'=IF(AND(M{r}="",N{r}=""),"",N(M{r})+N(N{r}))', PROB.get(lk), f'=IF(O{r}="","",O{r}*N(P{r}))', lc, nxt, "Jeffrey" if nxt != "n/a" else "", nd, health, risks, src, "A" if c else "B", "Yes" if (c and c[0] in ("Verbal commitment","Confirmed") and o["key"] in ("BYHENDO","TOB","CHILLAV","MOXY")) else "",
        "",
        f'=IF(OR(E{r}="Closed declined",E{r}="Closed lost",E{r}="Dormant",E{r}="Do not contact"),0,1)',
        f'=IF(AB{r}=0,"",IF(OR(S{r}="",T{r}="",AND(U{r}="",S{r}<>"n/a")),"Missing",""))',
        f'=IF(AB{r}=0,"",IF(AND(E{r}<>"Closed won",E{r}<>"Confirmed",OR(R{r}="",TODAY()-N(R{r})>Lists!$B$32)),"Stale",""))'],
        {13:"$#,##0",14:"$#,##0",15:"$#,##0",17:"$#,##0",16:"0%"})
    r += 1
OPP_LAST = r - 1
for col, key in [("E","Opportunity stage"),("F","Priority"),("H","Likelihood"),("T","Owner"),("V","Relationship health"),("Y","Confidence"),("Z","Yes/No")]: dv(ws, col, LISTREF[key])
ws.conditional_formatting.add(f"A2:AD{MAXR}", FormulaRule(formula=['$AC2="Missing"'], fill=AMBER))
ws.conditional_formatting.add(f"A2:AD{MAXR}", FormulaRule(formula=['$AD2="Stale"'], fill=GREY))

# ------------------------------------------------------------------ Agreements
ws = sheet("Agreements", ["Agreement ID","Partner ID","Partner","Type","Tier","Term start","Term end","Status","Signed by Banff Pride","Signed by partner","Document link","Asset deadline","Payment terms","Exclusivity","Approval lead time","Renewal option","Source","Confidence","Notes","Missing document?"], [11,11,28,26,20,11,11,26,12,12,40,11,26,40,36,24,30,9,44,10], "1F6E43")
for i, a in enumerate(S.AGRS, 2):
    put(ws, i, [a[0], oid(a[1]), oname(a[1])] + list(a[2:]) + [f'=IF(AND(H{i}<>"Draft",H{i}<>"Expired",K{i}=""),"Missing","")'])
dv(ws, "H", LISTREF["Agreement status"]); dv(ws, "R", LISTREF["Confidence"])
AGR_LAST = len(S.AGRS) + 1

# ------------------------------------------------------------------ Contributions
CTB_HDR = ["Contribution ID","Partner ID","Partner","Kind","Category","Description","Proposed","Confirmed","Received","Used","Valuation basis","Evidence link","Invoice number","Invoice sent","Due","Received date (per Wave)","GST","Review flag","Source","Confidence","Notes","Agreement ID","Outstanding","Status"]
ws = sheet("Contributions", CTB_HDR, [11,11,28,8,18,44,10,10,10,10,26,36,16,11,11,12,10,30,26,9,36,11,11,16], "1F6E43")
for i, c in enumerate(S.CTBS, 2):
    put(ws, i, [c[0], oid(c[1]), oname(c[1]), c[3], c[4], c[5], c[6], c[7], c[8], c[9], c[10], c[11], c[12], c[13], c[14], c[15], c[16], c[17], c[18], c[19], c[20], c[2],
        f'=IF(D{i}="Cash",N(H{i})-N(I{i}),N(I{i})-N(J{i}))',
        f'=IF(D{i}="Cash",IF(AND(N(H{i})>0,N(I{i})>=N(H{i})),"Received",IF(N(I{i})>0,"Partially received",IF(AND(N{i}<>"",O{i}<>"",O{i}<TODAY()),"Overdue",IF(N{i}<>"","Invoiced","Not invoiced")))),IF(N(I{i})>0,IF(N(J{i})>=N(I{i}),"Used","Received"),IF(N(H{i})>0,"Promised","Unvalued")))'],
        {7:"$#,##0",8:"$#,##0",9:"$#,##0",10:"$#,##0",23:"$#,##0"})
dv(ws, "D", LISTREF["Contribution kind"]); dv(ws, "E", LISTREF["Contribution category"]); dv(ws, "T", LISTREF["Confidence"])
ws.conditional_formatting.add(f"A2:X{MAXR}", FormulaRule(formula=['$X2="Overdue"'], fill=RED))
ws.conditional_formatting.add(f"A2:X{MAXR}", FormulaRule(formula=['AND($R2<>"",$A2<>"")'], fill=AMBER))
CTB_LAST = len(S.CTBS) + 1

# ------------------------------------------------------------------ Deliverables
DLV_HDR = ["Deliverable ID","Partner ID","Partner","Agreement ID","Owed by","Exact source language","Plain-language deliverable","Channel","Event or activation","Quantity","Internal owner","Due date","Dependencies","Approval required","Approval lead time (days)","Status","Delivered date","How delivered","Evidence required","Evidence ID","Verification","Substitute value","Exception notes","Source","Confidence","Notes","Open?","(reserved)","Flag","(reserved)"]
ws = sheet("Deliverables", DLV_HDR, [11,11,24,11,11,44,40,14,18,8,12,11,20,9,9,24,11,16,9,10,11,10,30,28,9,20,6,8,20,12], "1F6E43")
for i, d in enumerate(S.DLVS, 2):
    row = [d[0], oid(d[1]), oname(d[1]), d[2], d[3], d[4], d[5], d[6], d[7], d[8], d[9], d[10], d[11], d[12], d[13], d[14], d[15], d[16], d[17], d[18], d[19], d[20], d[21], d[22], d[23], ""]
    row += [f'=IF(OR(P{i}="Not started",P{i}="In progress",P{i}="Blocked by partner",P{i}="Blocked internally",P{i}="Partially delivered",P{i}="Reported without evidence"),1,0)',
            "",
            f'=IF(AA{i}=0,IF(AND(P{i}="Delivered and verified",S{i}="Yes",T{i}=""),"Missing evidence",""),IF(P{i}="Reported without evidence","Missing evidence",IF(OR(P{i}="Blocked by partner",P{i}="Blocked internally"),"Blocked",IF(L{i}="","No due date",IF(L{i}<TODAY(),"Overdue",IF(L{i}<=TODAY()+Lists!$B$31,"Due soon",""))))))',
            ""]
    put(ws, i, row, {30:"yyyy-mm-dd"})
for col, key in [("E","Owed by"),("K","Owner"),("N","Yes/No"),("P","Deliverable status"),("S","Yes/No"),("U","Verification"),("Y","Confidence")]: dv(ws, col, LISTREF[key])
ws.conditional_formatting.add(f"A2:AD{MAXR}", FormulaRule(formula=['$AC2="Overdue"'], fill=RED))
ws.conditional_formatting.add(f"A2:AD{MAXR}", FormulaRule(formula=['OR($AC2="Due soon",$AC2="Blocked")'], fill=AMBER))
ws.conditional_formatting.add(f"A2:AD{MAXR}", FormulaRule(formula=['$AC2="Missing evidence"'], fill=GREY))
DLV_LAST = len(S.DLVS) + 1

# ------------------------------------------------------------------ Evidence
ws = sheet("Evidence", ["Evidence ID","Deliverable ID","Type","Link or location","Captured by","Captured date","Verified by","Verified date","Shared with BLLT (perpetual licence)?","Notes"], [11,12,16,60,14,11,14,11,14,30], "1F6E43")
for i, e in enumerate(S.EVDS, 2): put(ws, i, list(e) + ["No", ""])
dv(ws, "I", LISTREF["Yes/No"])

# ------------------------------------------------------------------ Communications
COM_HDR = ["Communication ID","Partner ID","Partner","Contact","Date","Channel","Subject or meeting title","Summary","Commitments by Banff Pride","Commitments by partner","Decisions","Open questions","Follow-up action","Owner","Due date","Source link","Proposed records","Triage","Needs reply?"]
ws = sheet("Communications", COM_HDR, [11,11,24,18,11,10,34,60,34,34,20,24,30,10,11,36,26,9,10], "1F6E43")
for i, c in enumerate(S.COMS, 2):
    put(ws, i, [c[0], oid(c[1]), oname(c[1])] + list(c[2:]) + [f'=IF(AND(R{i}="Open",N{i}<>"Partner",O{i}<>"",O{i}<=TODAY()+7),"Yes","")'])
dv(ws, "F", LISTREF["Channel"]); dv(ws, "N", LISTREF["Owner"]); dv(ws, "R", LISTREF["Triage"])
ws.conditional_formatting.add(f"A2:S{MAXR}", FormulaRule(formula=['$S2="Yes"'], fill=AMBER))
COM_LAST = len(S.COMS) + 1

# ------------------------------------------------------------------ Brand assets
ws = sheet("BrandAssets", ["Asset ID","Partner ID","Approved public name","Library folder link","Primary logo link","Alternative logo links","File format / dimensions","Restrictions and approval rules","Approval status","Received date","Source contact / thread","Approval lead time","Agreement expiry","Last verified","Notes"], [10,11,28,36,30,30,18,44,30,11,36,14,11,11,30], "1F6E43")
for i, b in enumerate(S.BRAND, 2):
    put(ws, i, [b[0], oid(b[1]), b[2], b[3], b[4], b[5], b[6], b[7], b[8], None, b[10], "", None, None, b[9]])

# ------------------------------------------------------------------ Tickets
ws = sheet("Tickets", ["Ticket ID","Partner ID","Partner","Agreement ID","Type","Description","Quantity","Purpose or event","Issued to","Status","Received or issued date","Evidence or source","Next step","Deadline"], [10,11,24,11,18,44,8,26,14,14,11,30,34,11], "1F6E43")
for i, t in enumerate(S.TIX, 2):
    put(ws, i, [t[0], oid(t[1]), oname(t[1]), t[2], t[3], t[4], t[5], t[6], t[7], t[8], t[9], t[10], t[11], t[12]])

# ------------------------------------------------------------------ Events
ws = sheet("Events", ["Event ID","Event or activation","Date","Host partner ID","Venue","Type","Status","Notes","Linked partner IDs","Deliverables linked","Evidence to capture"], [10,34,11,12,22,12,10,60,30,11,40], "1F6E43")
for i, e in enumerate(S.EVENTS, 2):
    linked = ";".join(oid(k) for k in e[8].split(";") if k)
    put(ws, i, [e[0], e[1], e[2], oid(e[3]) if e[3] else "", e[4], e[5], e[6], e[7], linked, f'=COUNTIFS(Deliverables!$I$2:$I${MAXR},B{i})', "Photos of partner signage and activations; attendance count; screenshots of tagged posts"])

# ------------------------------------------------------------------ Metrics
ws = sheet("Metrics", ["Metric ID","Metric","Value","Unit","Period","Scope","Partner ID","Source","Source link","Confidence","Retrieved"], [10,40,14,10,14,20,11,26,36,9,11], "1F6E43")
for i, m in enumerate(S.METRICS, 2):
    put(ws, i, [m[0], m[1], m[2], m[3], m[4], m[5], oid(m[6]) if m[6] else "", m[7], m[8], m[9], S.TODAY if m[2] else None])

# ------------------------------------------------------------------ Reports
REPORT_KEYS = ["BLLT","SKIBIG3","BHC","QTIL","HIGHLINE","ATB","ROAM","MLB","RMR","BYHENDO","MOXY","TOB","AIRPORTER","BASECAMP","WHYTE","MTNEVENTS","BELL"]
ws = sheet("Reports", ["Report ID","Partner ID","Partner","Report type","Due","Status","Deliverables total","Delivered and verified","Missing evidence","Cash confirmed","Cash received","In-kind confirmed","Evidence items","Approved by","Sent date","Link"], [10,11,28,22,11,12,10,10,10,11,11,11,9,12,11,30], "1F6E43")
for i, k in enumerate(REPORT_KEYS, 2):
    due = D(2026,11,12) if k == "BLLT" else D(2026,11,30)
    put(ws, i, [f"RPT-{i-1:03d}", oid(k), oname(k), "BLLT Final Report (Appendix B)" if k=="BLLT" else "Personalized partner report", due, "Not started",
        f'=COUNTIFS(Deliverables!$B$2:$B${MAXR},B{i})', f'=COUNTIFS(Deliverables!$B$2:$B${MAXR},B{i},Deliverables!$P$2:$P${MAXR},"Delivered and verified")', f'=COUNTIFS(Deliverables!$B$2:$B${MAXR},B{i},Deliverables!$AC$2:$AC${MAXR},"Missing evidence")',
        f'=SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$B$2:$B${MAXR},B{i},Contributions!$D$2:$D${MAXR},"Cash")', f'=SUMIFS(Contributions!$I$2:$I${MAXR},Contributions!$B$2:$B${MAXR},B{i},Contributions!$D$2:$D${MAXR},"Cash")', f'=SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$B$2:$B${MAXR},B{i},Contributions!$D$2:$D${MAXR},"In-kind")',
        f'=SUMPRODUCT((Deliverables!$B$2:$B${MAXR}=B{i})*(Deliverables!$T$2:$T${MAXR}<>""))', "", None, ""], {10:"$#,##0",11:"$#,##0",12:"$#,##0"})
dv(ws, "F", LISTREF["Report status"])
RPT_LAST = len(REPORT_KEYS) + 1

# ------------------------------------------------------------------ Renewals
ws = sheet("Renewals", ["Renewal ID","Partner ID","Partner","Situation","Renewal prep date","Status","Owner","Source","Brief link","2025 value (cash / in-kind)","2026 value (cash / in-kind)"], [10,11,28,60,12,12,10,30,20,20,20], "1F6E43")
V25 = {"BLLT":"$20,000 / ~$7,500","BHC":"$1,000 / $3,000","ATB":"$5,000 / -","PURSUIT":"$5,000 / $2,196","FAIRMONT":"$2,000 / -","BLC":"- / $5,000","WILDLIFE":"- / $1,000 + code","SKIBIG3":"- / in-kind (2025 in-kind sponsor)"}
for i, rn in enumerate(S.RENEWALS, 2):
    put(ws, i, [rn[0], oid(rn[1]), oname(rn[1]), rn[2], rn[3], rn[4], rn[5], rn[6], "", V25.get(rn[1], ""), f'=TEXT(SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$B$2:$B${MAXR},B{i},Contributions!$D$2:$D${MAXR},"Cash"),"$#,##0")&" / "&TEXT(SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$B$2:$B${MAXR},B{i},Contributions!$D$2:$D${MAXR},"In-kind"),"$#,##0")'])
dv(ws, "F", LISTREF["Renewal status"])

# ------------------------------------------------------------------ Decisions
ws = sheet("Decisions", ["Decision ID","Date raised","Topic","Decision or request","Status","Decided by","Source","Decided date","Notes"], [10,11,22,70,16,12,26,11,30], "1F6E43")
for i, d in enumerate(S.DECISIONS, 2):
    put(ws, i, [d[0], d[1], d[2], d[3], d[4], d[5], d[6], d[1] if d[4] in ("Approved","Recorded") else None, ""])
dv(ws, "E", LISTREF["Decision status"])
DEC_LAST = len(S.DECISIONS) + 1

# ------------------------------------------------------------------ Improvements
ws = sheet("Improvements", ["Improvement ID","Date","Type","Evidence of the problem","Root cause","Proposed change","Expected benefit","Risks","Before and after example","Regression test","ED approval","Version","Rollback path","Post-implementation measure","Status"], [10,11,18,40,30,40,30,24,30,20,12,8,24,24,12], "1F6E43")
put(ws, 2, ["IMP-001", S.TODAY, "Tooling limitation", "The Drive connector cannot convert an uploaded .xlsx into a native Google Sheet, so the first version needs a one-click Save as Google Sheets by the ED.", "Connector conversion path rejects xlsx", "Once the native Sheet exists, future changes are made in place; keep the generator for schema changes only.", "No repeated conversions", "Formula differences between Excel and Sheets", "See operating guide", "Verified after ED conversion", "", "1.0", "Keep the .xlsx", "Formulas evaluate without #NAME", "Open"])
put(ws, 3, ["IMP-002", S.TODAY, "Data quality", "Owner and next action were empty on all 106 July prospects.", "Research sheet had no operating fields in use", "Opportunities tab flags Missing owner or action; Data Quality view counts them.", "One clear next action per relationship", "", "", "", "", "1.0", "", "Missing count trends to zero", "Open"])

# ------------------------------------------------------------------ Audit log
ws = sheet("AuditLog", ["Timestamp","Actor","Action","Record IDs","Detail","Source"], [18,12,26,30,70,30], "7F7F7F", freeze=None)
put(ws, 2, [dt.datetime(2026,9,6,0,0), "Claude", "Workbook built", "all", "Initial build from Gate 1 audit; seeded organizations, agreements, contributions, deliverables, communications, tickets, events, metrics, decisions.", "Gate 1 decision brief"])
put(ws, 3, [dt.datetime(2026,9,6,0,0), "Jeffrey", "Decisions recorded", "DEC-001 to DEC-011", "Gate 1 approvals and corrections (Moxy venue, payments received, Roam passes, Town of Banff, in-kind rule, Brand Asset Library location).", "Chat 2026-09-05"])

# ------------------------------------------------------------------ Source index
ws = sheet("SourceIndex", ["Source ID","System","Title","ID or link","Source modified","Retrieved","Authority","Note"], [10,10,50,40,12,12,26,44], "7F7F7F", freeze=None)
for i, s in enumerate(S.SOURCES, 2): put(ws, i, [f"SRC-{i-1:03d}"] + list(s))

# ------------------------------------------------------------------ Thank-you email workflow
ws = sheet("ThankYouEmail", ["Partner ID","Partner","Verified support (auto)","Recognition line (edit)","Recipient contact (from Contacts)","Include?","Draft status","Draft link","Final approval","Approved by","Approval date","Scheduled send","Sent","Notes"], [11,28,40,50,26,12,14,24,14,12,11,18,11,30], "C05A2B", note="Schedule: recipient review Sept 28; drafts Sept 29; approval deadline Sept 30; send Oct 1 09:00 America/Edmonton. Individual emails only. Nothing sends unless Final approval = Approved for that row.")
for i, k in enumerate(REPORT_KEYS, 3):
    put(ws, i, [oid(k), oname(k), f'=IFERROR("Cash confirmed "&TEXT(SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$B$2:$B${MAXR},A{i},Contributions!$D$2:$D${MAXR},"Cash"),"$#,##0")&"; in-kind confirmed "&TEXT(SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$B$2:$B${MAXR},A{i},Contributions!$D$2:$D${MAXR},"In-kind"),"$#,##0"),"")', "", f'=IFERROR(INDEX(Contacts!$D$3:$D${MAXR},MATCH(A{i},Contacts!$B$3:$B${MAXR},0)),"")', "Not reviewed", "Not started", "", "Not reviewed", "", None, "2026-10-01 09:00 America/Edmonton", "", ""])
dv(ws, "F", LISTREF["Email approval"], first=3); dv(ws, "I", LISTREF["Email approval"], first=3)
ws.cell(22, 1, "Message requirements").font = BOLD
for j, t in enumerate(["Thank the partner for their verified support (only what the Contributions tab confirms).","Recognize the organization by its approved public name.","Note that the Banff Pride Festival runs October 2 to 12.","Explain that email response times may be delayed during the festival.","Provide approved urgent-contact instructions where applicable.","Send individual messages, never a visible group email."], 23):
    ws.cell(j, 1, t)

# ================================================================== VIEWS
def view(name, title, headers, formula, widths=None, note=None, datecols=(), moneycols=()):
    ws = wb.create_sheet(name, 0)
    ws.sheet_properties.tabColor = "1F4E79"
    ws.cell(1, 1, title).font = TITLE_FONT
    ws.cell(2, 1, note or "Formula view. Do not type here; edit the record tabs.").font = NOTE_FONT
    for i, h in enumerate(headers, 1):
        c = ws.cell(4, i, h); c.fill = VIEW_FILL; c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True)
    ws.cell(5, 1, formula)
    ws.freeze_panes = "A5"
    for i, w in enumerate(widths or [18]*len(headers), 1): ws.column_dimensions[L(i)].width = w
    for rr in range(5, 175 if name=="Prospect Pipeline" else 60):
        for cc in datecols: ws.cell(rr, cc).number_format = "yyyy-mm-dd"
        for cc in moneycols: ws.cell(rr, cc).number_format = "$#,##0"
    return ws

# Deliverables projection helper: pick columns via CHOOSE is not portable; project by selecting ranges with FILTER on each block.
def fl(rng, cond, sort_col=None):
    core = f'FILTER({rng},{cond})'
    if sort_col: core = f'SORT({core},{sort_col},1)'
    return f'=IFERROR({core},"None")'

DL = f"Deliverables!$A$2:$Z${MAXR}"
view("Data Quality", "Data Quality and Exceptions", ["Check","Count","What to do"], "", [46,8,60], "Counts update automatically. Open the named tab and filter on the flag column to fix items.")
ws = wb["Data Quality"]
checks = [
 ("Deliverables overdue", f'=COUNTIF(Deliverables!$AC$2:$AC${MAXR},"Overdue")', "Deliverables tab, Flag = Overdue"),
 ("Deliverables blocked", f'=COUNTIF(Deliverables!$AC$2:$AC${MAXR},"Blocked")', "Deliverables tab, Flag = Blocked"),
 ("Deliverables reported without evidence", f'=COUNTIF(Deliverables!$AC$2:$AC${MAXR},"Missing evidence")', "Add an Evidence row and link it"),
 ("Open deliverables with no due date", f'=COUNTIF(Deliverables!$AC$2:$AC${MAXR},"No due date")', "Set a due date"),
 ("Active opportunities missing owner or next action", f'=COUNTIF(Opportunities!$AC$2:$AC${MAXR},"Missing")', "Opportunities tab, column Missing owner or action"),
 ("Stale prospects (no contact in 60 days)", f'=COUNTIF(Opportunities!$AD$2:$AD${MAXR},"Stale")', "Decide: contact, park for 2027, or close"),
 ("Agreements missing a document link", f'=COUNTIF(Agreements!$T$2:$T${MAXR},"Missing")', "Attach the signed PDF or written confirmation"),
 ("Contributions flagged for review", f'=COUNTIF(Contributions!$R$2:$R${MAXR},"?*")', "Contributions tab, Review flag column"),
 ("Cash confirmed but not received", f'=COUNTIFS(Contributions!$D$2:$D${MAXR},"Cash",Contributions!$X$2:$X${MAXR},"Invoiced")+COUNTIFS(Contributions!$D$2:$D${MAXR},"Cash",Contributions!$X$2:$X${MAXR},"Overdue")+COUNTIFS(Contributions!$D$2:$D${MAXR},"Cash",Contributions!$X$2:$X${MAXR},"Not invoiced")', "Check Wave, enter Received date"),
 ("In-kind rows with no valuation", f'=COUNTIFS(Contributions!$D$2:$D${MAXR},"In-kind",Contributions!$X$2:$X${MAXR},"Unvalued")', "Set a valuation basis or leave as unvalued support"),
 ("Communications needing a reply within 7 days", f'=COUNTIF(Communications!$S$2:$S${MAXR},"Yes")', "Communications tab, Needs reply = Yes"),
 ("Brand assets pending", f'=COUNTIF(BrandAssets!$I$2:$I${MAXR},"Pending*")', "Request logos; file in the Brand Asset Library"),
 ("Decisions awaiting approval", f'=COUNTIF(Decisions!$E$2:$E${MAXR},"Awaiting approval")', "Decisions Required tab"),
 ("Duplicate organization names", f'=SUMPRODUCT((COUNTIF(Organizations!$C$2:$C${MAXR},Organizations!$C$2:$C${MAXR})>1)*(Organizations!$C$2:$C${MAXR}<>""))', "Merge duplicates; keep one Partner ID"),
]
for i, (a, f, c) in enumerate(checks, 5):
    ws.cell(i, 1, a); ws.cell(i, 2, f); ws.cell(i, 3, c)
ws.conditional_formatting.add("B5:B30", FormulaRule(formula=['B5>0'], fill=AMBER))

view("Renewal Pipeline", "Renewal Pipeline", ["Renewal ID","Partner ID","Partner","Situation","Renewal prep date","Status","Owner","Source","Brief link","2025 value","2026 value"], fl(f"Renewals!$A$2:$K${MAXR}", f"Renewals!$A$2:$A${MAXR}<>\"\"", 5), [10,11,28,60,12,12,10,30,20,20,20], datecols=(5,))
view("Reporting Readiness", "Reporting Readiness", ["Report ID","Partner ID","Partner","Report type","Due","Status","Deliverables total","Delivered and verified","Missing evidence","Cash confirmed","Cash received","In-kind confirmed","Evidence items","Approved by","Sent date","Link"], fl(f"Reports!$A$2:$P${MAXR}", f"Reports!$A$2:$A${MAXR}<>\"\"", 5), [10,11,28,24,11,12,10,10,10,11,11,11,9,12,11,30], datecols=(5,15), moneycols=(10,11,12))
view("Events View", "Event and Activation View", ["Event ID","Event or activation","Date","Host partner ID","Venue","Type","Status","Notes","Linked partner IDs","Deliverables linked","Evidence to capture"], fl(f"Events!$A$2:$K${MAXR}", f"Events!$A$2:$A${MAXR}<>\"\"", 3), [10,34,11,12,22,12,10,60,30,11,40], datecols=(3,))
view("Brand Assets View", "Brand Assets", ["Asset ID","Partner ID","Approved public name","Library folder link","Primary logo link","Alternative logo links","Format","Restrictions and approval rules","Approval status","Received date","Source","Approval lead time","Agreement expiry","Last verified","Notes"], fl(f"BrandAssets!$A$2:$O${MAXR}", f"BrandAssets!$A$2:$A${MAXR}<>\"\"", 9), [10,11,28,36,30,30,18,44,30,11,36,14,11,11,30], datecols=(10,13,14))
view("Follow-up Needed", "Communications Requiring Follow-up", COM_HDR, fl(f"Communications!$A$2:$S${MAXR}", f"Communications!$S$2:$S${MAXR}=\"Yes\"", 15), [11,11,24,18,11,10,34,60,34,34,20,24,30,10,11,36,26,9,10], "Open items owned by Banff Pride with a due date within 7 days.", datecols=(5,15))
view("Missing Evidence", "Missing Evidence", DLV_HDR[:26], fl(DL, f"Deliverables!$AC$2:$AC${MAXR}=\"Missing evidence\"", 12), [11,11,24,11,11,44,40,14,18,8,12,11,20,9,9,24,11,16,9,10,11,10,30,28,9,20], "Deliverables reported done without an evidence link. Add a row in Evidence and link it.", datecols=(12,17))
view("Deliverables at Risk", "Deliverables at Risk", DLV_HDR[:26], fl(DL, f"Deliverables!$AC$2:$AC${MAXR}=\"Overdue\"", 12), [11,11,24,11,11,44,40,14,18,8,12,11,20,9,9,24,11,16,9,10,11,10,30,28,9,20], "Overdue items first. Blocked items are on the Blocked tab.", datecols=(12,17))
# second block placed after a fixed gap is not portable with spills; blocked items get their own tab instead
view("Blocked", "Deliverables Blocked", DLV_HDR[:26], fl(DL, f"Deliverables!$AC$2:$AC${MAXR}=\"Blocked\"", 12), [11,11,24,11,11,44,40,14,18,8,12,11,20,9,9,24,11,16,9,10,11,10,30,28,9,20], datecols=(12,17))
view("Deliverables Due", "Deliverables Due (next 14 days)", DLV_HDR[:26], fl(DL, f"Deliverables!$AC$2:$AC${MAXR}=\"Due soon\"", 12), [11,11,24,11,11,44,40,14,18,8,12,11,20,9,9,24,11,16,9,10,11,10,30,28,9,20], datecols=(12,17))
view("Confirmed Support", "Confirmed Support", CTB_HDR, fl(f"Contributions!$A$2:$X${MAXR}", f"Contributions!$H$2:$H${MAXR}>0", 3), [11,11,28,8,18,44,10,10,10,10,26,36,16,11,11,12,10,30,26,9,36,11,11,16], "Rows with a confirmed value. Unvalued in-kind support is listed in Contributions.", datecols=(14,15,16), moneycols=(7,8,9,10,23))
view("Prospect Pipeline", "Prospect Pipeline (active opportunities)", OPP_HDR[:26], fl(f"Opportunities!$A$2:$Z${MAXR}", f"Opportunities!$AB$2:$AB${MAXR}=1", 6), [11,11,28,8,20,10,8,12,12,12,26,40,11,11,11,9,11,12,44,12,12,12,36,30,9,8], datecols=(18,21), moneycols=(13,14,15,17))
view("Directory", "Sponsor and Partner Directory", ["Partner ID","Partner","Stage","Priority","Next action","Next-action owner","Next-action date","Relationship health","Last contact","Open deliverables"], "", [11,30,20,10,50,12,12,12,11,10], "One row per organization. Edit next actions in Opportunities.")
ws = wb["Directory"]
cond = f"Opportunities!$AB$2:$AB${MAXR}=1"
ws.cell(5, 1, f'=IFERROR(FILTER(Opportunities!$B$2:$C${MAXR},{cond}),"None")')
ws.cell(5, 3, f'=IFERROR(FILTER(Opportunities!$E$2:$F${MAXR},{cond}),"")')
ws.cell(5, 5, f'=IFERROR(FILTER(Opportunities!$S$2:$U${MAXR},{cond}),"")')
ws.cell(5, 8, f'=IFERROR(FILTER(Opportunities!$V$2:$V${MAXR},{cond}),"")')
ws.cell(5, 9, f'=IFERROR(FILTER(Opportunities!$R$2:$R${MAXR},{cond}),"")')
ws.cell(5, 10, f'=IFERROR(FILTER(Organizations!$P$2:$P${MAXR},Organizations!$A$2:$A${MAXR}<>""),"")')
for rr in range(5, 175):
    ws.cell(rr, 7).number_format = "yyyy-mm-dd"; ws.cell(rr, 9).number_format = "yyyy-mm-dd"
ws.cell(2, 1, "One row per active organization, in record order. Column J counts open deliverables for the same rows. Edit next actions in Opportunities.").font = NOTE_FONT
view("Decisions Required", "Decisions Required", ["Decision ID","Date raised","Topic","Decision or request","Status","Decided by","Source","Decided date","Notes"], fl(f"Decisions!$A$2:$I${MAXR}", f"Decisions!$E$2:$E${MAXR}=\"Awaiting approval\"", 2), [10,11,22,70,16,12,26,11,30], "Approve, reject or defer in the Decisions tab (Status column). Contribution review flags are listed in Data Quality.", datecols=(2,8))

# ------------------------------------------------------------------ ED Dashboard
ws = wb.create_sheet("ED Dashboard", 0); ws.sheet_properties.tabColor = "6C0B06"
ws.cell(1, 1, "Banff Pride Partnerships: ED Dashboard").font = TITLE_FONT
ws.cell(2, 1, '="Today: "&TEXT(TODAY(),"yyyy-mm-dd")&"   |   Days to festival: "&(Lists!$B$33-TODAY())').font = NOTE_FONT
for i, w in enumerate([44, 14, 22, 44, 14], 1): ws.column_dimensions[L(i)].width = w
kpis = [
 ("Decisions awaiting approval", f'=COUNTIF(Decisions!$E$2:$E${MAXR},"Awaiting approval")', "Decisions Required"),
 ("Overdue commitments", f'=COUNTIF(Deliverables!$AC$2:$AC${MAXR},"Overdue")', "Deliverables at Risk"),
 ("Deliverables due within 14 days", f'=COUNTIF(Deliverables!$AC$2:$AC${MAXR},"Due soon")', "Deliverables Due"),
 ("Work blocked by partners or internally", f'=COUNTIF(Deliverables!$AC$2:$AC${MAXR},"Blocked")', "Blocked"),
 ("Missing agreements (no document)", f'=COUNTIF(Agreements!$T$2:$T${MAXR},"Missing")', "Agreements"),
 ("Cash confirmed but not received", f'=SUMIFS(Contributions!$W$2:$W${MAXR},Contributions!$D$2:$D${MAXR},"Cash")', "Contributions"),
 ("Brand assets pending", f'=COUNTIF(BrandAssets!$I$2:$I${MAXR},"Pending*")', "Brand Assets View"),
 ("Communications needing a response", f'=COUNTIF(Communications!$S$2:$S${MAXR},"Yes")', "Follow-up Needed"),
 ("Deliverables reported without evidence", f'=COUNTIF(Deliverables!$AC$2:$AC${MAXR},"Missing evidence")', "Missing Evidence"),
]
ws.cell(4, 1, "Exceptions").font = BOLD; ws.cell(4, 2, "Count").font = BOLD; ws.cell(4, 3, "Open this tab").font = BOLD; ws.cell(4, 4, "Support").font = BOLD; ws.cell(4, 5, "Amount").font = BOLD
for i, (label, f, tab) in enumerate(kpis, 5):
    ws.cell(i, 1, label); ws.cell(i, 3, tab).font = NOTE_FONT; c = ws.cell(i, 2, f)
    if "Cash" in label: c.number_format = "$#,##0"
ws.conditional_formatting.add("B5:B13", FormulaRule(formula=['B5>0'], fill=AMBER))
money = [
 ("Total confirmed cash support", f'=SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$D$2:$D${MAXR},"Cash")'),
 ("Total cash received", f'=SUMIFS(Contributions!$I$2:$I${MAXR},Contributions!$D$2:$D${MAXR},"Cash")'),
 ("Total confirmed in-kind support (valued)", f'=SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$D$2:$D${MAXR},"In-kind")'),
 ("Total in-kind received (valued)", f'=SUMIFS(Contributions!$I$2:$I${MAXR},Contributions!$D$2:$D${MAXR},"In-kind")'),
 ("Total support received", f'=E7+E9'),
 ("In-kind rows still unvalued", f'=COUNTIFS(Contributions!$D$2:$D${MAXR},"In-kind",Contributions!$X$2:$X${MAXR},"Unvalued")'),
 ("Proposed (not confirmed) cash", f'=SUMIFS(Contributions!$G$2:$G${MAXR},Contributions!$D$2:$D${MAXR},"Cash")-SUMIFS(Contributions!$H$2:$H${MAXR},Contributions!$D$2:$D${MAXR},"Cash")'),
 ("Reporting readiness: deliverables delivered and verified", f'=COUNTIF(Deliverables!$P$2:$P${MAXR},"Delivered and verified")&" of "&COUNTA(Deliverables!$A$2:$A${MAXR})'),
 ("Partner reports not started", f'=COUNTIF(Reports!$F$2:$F${MAXR},"Not started")'),
]
for i, (label, f) in enumerate(money, 5):
    ws.cell(i, 4, label); c = ws.cell(i, 5, f)
    if "Total" in label or "Proposed" in label: c.number_format = "$#,##0"
ws.cell(15, 1, "Recommended next actions (overdue and due-soon deliverables owed by Banff Pride, earliest first)").font = BOLD
ws.cell(16, 1, "Deliverable").font = BOLD; ws.cell(16, 2, "Due").font = BOLD; ws.cell(16, 4, "Partner").font = BOLD; ws.cell(16, 5, "Flag").font = BOLD
for n in range(1, 11):
    rr = 16 + n
    base = f'SORT(FILTER(Deliverables!$C$2:$AC${MAXR},(Deliverables!$E$2:$E${MAXR}="Banff Pride")*((Deliverables!$AC$2:$AC${MAXR}="Overdue")+(Deliverables!$AC$2:$AC${MAXR}="Due soon")+(Deliverables!$AC$2:$AC${MAXR}="Blocked"))),10,1)'
    ws.cell(rr, 1, f'=IFERROR(INDEX({base},{n},5),"")')
    ws.cell(rr, 2, f'=IFERROR(INDEX({base},{n},10),"")').number_format = "yyyy-mm-dd"
    ws.cell(rr, 4, f'=IFERROR(INDEX({base},{n},1),"")')
    ws.cell(rr, 5, f'=IFERROR(INDEX({base},{n},27),"")')
ws.cell(28, 1, "Top open relationships by next-action date").font = BOLD
ws.cell(29, 1, "Partner").font = BOLD; ws.cell(29, 2, "Next-action date").font = BOLD; ws.cell(29, 4, "Next action").font = BOLD
for n in range(1, 11):
    rr = 29 + n
    base = f'SORT(FILTER(Opportunities!$C$2:$U${MAXR},(Opportunities!$AB$2:$AB${MAXR}=1)*(Opportunities!$U$2:$U${MAXR}<>"")),19,1)'
    ws.cell(rr, 1, f'=IFERROR(INDEX({base},{n},1),"")')
    ws.cell(rr, 2, f'=IFERROR(INDEX({base},{n},19),"")').number_format = "yyyy-mm-dd"
    ws.cell(rr, 4, f'=IFERROR(INDEX({base},{n},17),"")')
ws.cell(41, 1, "Self-check (should equal the row counts in the record tabs)").font = NOTE_FONT
ws.cell(42, 1, "Organizations"); ws.cell(42, 2, f'=COUNTA(Organizations!$A$2:$A${MAXR})')
ws.cell(43, 1, "Deliverables"); ws.cell(43, 2, f'=COUNTA(Deliverables!$A$2:$A${MAXR})')
ws.cell(44, 1, "If any cell above shows #NAME? open File > Save as Google Sheets").font = NOTE_FONT
ws.freeze_panes = "A4"

# helper flag columns used by the dashboard: Deliverables AD (BP-owed and actionable), Opportunities AE (active with a next-action date)
dws = wb["Deliverables"]
ows = wb["Opportunities"]
# extend row formulas on record tabs so new rows compute automatically
for i in range(DLV_LAST + 1, DLV_LAST + EXT + 1):
    for col, f in [(27, f'=IF(A{i}="","",IF(OR(P{i}="Not started",P{i}="In progress",P{i}="Blocked by partner",P{i}="Blocked internally",P{i}="Partially delivered",P{i}="Reported without evidence"),1,0))'),
                   (29, f'=IF(A{i}="","",IF(AA{i}=0,IF(AND(P{i}="Delivered and verified",S{i}="Yes",T{i}=""),"Missing evidence",""),IF(P{i}="Reported without evidence","Missing evidence",IF(OR(P{i}="Blocked by partner",P{i}="Blocked internally"),"Blocked",IF(L{i}="","No due date",IF(L{i}<TODAY(),"Overdue",IF(L{i}<=TODAY()+Lists!$B$31,"Due soon","")))))))')]:
        dws.cell(i, col, f)
for i in range(OPP_LAST + 1, OPP_LAST + EXT + 1):
    ows.cell(i, 28, f'=IF(A{i}="","",IF(OR(E{i}="Closed declined",E{i}="Closed lost",E{i}="Dormant",E{i}="Do not contact"),0,1))')
    ows.cell(i, 29, f'=IF(OR(A{i}="",AB{i}=0),"",IF(OR(S{i}="",T{i}="",AND(U{i}="",S{i}<>"n/a")),"Missing",""))')
    ows.cell(i, 30, f'=IF(OR(A{i}="",AB{i}=0),"",IF(AND(E{i}<>"Closed won",E{i}<>"Confirmed",OR(R{i}="",TODAY()-N(R{i})>Lists!$B$32)),"Stale",""))')
cws = wb["Contributions"]
for i in range(CTB_LAST + 1, CTB_LAST + EXT + 1):
    cws.cell(i, 23, f'=IF(A{i}="","",IF(D{i}="Cash",N(H{i})-N(I{i}),N(I{i})-N(J{i})))')
    cws.cell(i, 24, f'=IF(A{i}="","",IF(D{i}="Cash",IF(AND(N(H{i})>0,N(I{i})>=N(H{i})),"Received",IF(N(I{i})>0,"Partially received",IF(AND(N{i}<>"",O{i}<>"",O{i}<TODAY()),"Overdue",IF(N{i}<>"","Invoiced","Not invoiced")))),IF(N(I{i})>0,IF(N(J{i})>=N(I{i}),"Used","Received"),IF(N(H{i})>0,"Promised","Unvalued"))))')
mws = wb["Communications"]
for i in range(COM_LAST + 1, COM_LAST + EXT + 1):
    mws.cell(i, 19, f'=IF(A{i}="","",IF(AND(R{i}="Open",N{i}<>"Partner",O{i}<>"",O{i}<=TODAY()+7),"Yes",""))')
aws = wb["Agreements"]
for i in range(AGR_LAST + 1, AGR_LAST + EXT + 1):
    aws.cell(i, 20, f'=IF(A{i}="","",IF(AND(H{i}<>"Draft",H{i}<>"Expired",K{i}=""),"Missing",""))')

# README tab
ws = wb.create_sheet("README", 0); ws.sheet_properties.tabColor = "6C0B06"
ws.column_dimensions["A"].width = 120
lines = [
 ("Banff Pride Partnerships", TITLE_FONT),
 ("Version 1.0, built 2026-09-06 from the Gate 1 audit. Owner: Executive Director. Canadian English.", NOTE_FONT),
 ("", None),
 ("How to use", BOLD),
 ("1. Start on ED Dashboard. Every number links to the tab that explains it. Zero exceptions means nothing needs you today.", None),
 ("2. Blue tabs are views. They are formulas; never type in them. Green tabs are records; that is where you edit.", None),
 ("3. One row per thing: one organization, one agreement, one contribution line, one deliverable, one piece of evidence.", None),
 ("4. Every relationship has one next action in Opportunities (columns Next action, owner, date). Directory shows them all.", None),
 ("5. A deliverable is only Delivered and verified when an Evidence row is linked. Reported without evidence stays flagged.", None),
 ("6. Cash received is entered from Wave with the date. This workbook is not a ledger.", None),
 ("7. In-kind rows need a valuation basis. Promo codes, discounts and free inclusions are in-kind (ED rule 2026-09-05); unvalued rows are counted separately.", None),
 ("8. Contacts is restricted. Do not export or share it. Reports never name individuals.", None),
 ("9. Decisions Required lists what needs your approval. Change Status in the Decisions tab to record your call.", None),
 ("10. Nothing is emailed from this workbook. The ThankYouEmail tab only records recipients and approvals.", None),
 ("", None),
 ("Status fields (kept separate on purpose)", BOLD),
 ("Opportunity stage (Opportunities), Agreement status (Agreements), Payment or in-kind status (Contributions, computed), Fulfilment status (Deliverables), Relationship health (Opportunities).", None),
 ("", None),
 ("Authority order when sources conflict", BOLD),
 ("1 Signed agreement, 2 Approved written amendment, 3 Confirmed written correspondence, 4 Approved internal record, 5 Meeting record, 6 Working notes or inference. Confidence A/B/C on each row reflects this.", None),
 ("", None),
 ("Source of truth (ED decision 2026-09-05)", BOLD),
 ("July 2026 prospect sheet + signed agreements + Gmail evidence seeded this workbook; the August donation tracker was merged. Other trackers are archived. This workbook is now the operating record.", None),
 ("", None),
 ("Brand Asset Library", BOLD),
 ("Partnerships / Brand Asset Library / [Partner ID] Partner Name / 01 Current Approved Logos, 02 Brand Guidelines, 03 Approval Records, 99 Expired - Do Not Use. Use only files in 01. Never delete expired logos. Use shortcuts elsewhere.", None),
 ("", None),
 ("If a view shows #NAME?", BOLD),
 ("Use File > Save as Google Sheets once. Views use FILTER, SORT and INDEX, which Google Sheets supports natively.", None),
]
for i, (t, f) in enumerate(lines, 1):
    c = ws.cell(i, 1, t); c.alignment = Alignment(wrap_text=True, vertical="top")
    if f: c.font = f

# order: README, ED Dashboard, views..., records..., system
order = ["README","ED Dashboard","Decisions Required","Directory","Prospect Pipeline","Confirmed Support","Deliverables Due","Deliverables at Risk","Blocked","Missing Evidence","Follow-up Needed","Brand Assets View","Events View","Reporting Readiness","Renewal Pipeline","Data Quality",
         "Organizations","Contacts","Opportunities","Agreements","Contributions","Deliverables","Evidence","Communications","BrandAssets","Tickets","Events","Metrics","Reports","Renewals","Decisions","Improvements","ThankYouEmail","AuditLog","SourceIndex","Lists"]
wb._sheets = [wb[n] for n in order]
wb.save("Banff Pride Partnerships.xlsx")
print("orgs", len(orgs), "opps", OPP_LAST-1, "agr", AGR_LAST-1, "ctb", CTB_LAST-1, "dlv", DLV_LAST-1, "com", COM_LAST-1)
