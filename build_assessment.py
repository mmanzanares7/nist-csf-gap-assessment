"""Build csf_gap_assessment.xlsx: a NIST CSF 2.0 gap assessment of a fictional healthcare
services company. Gaps, priority scores, tiers, ranking, function summaries and the roadmap
are formulas. Blue cells are inputs (scores, criticality, evidence, actions).

Usage: python3 build_assessment.py ../csf_gap_assessment.xlsx
"""
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, RadarChart, Reference
from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = sys.argv[1] if len(sys.argv) > 1 else "csf_gap_assessment.xlsx"
NAVY, BLUE = "1F3864", "0000FF"
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill("solid", start_color=NAVY, end_color=NAVY)
RED = PatternFill("solid", start_color="FFC7CE", end_color="FFC7CE")
AMB = PatternFill("solid", start_color="FFEB9C", end_color="FFEB9C")
GRN = PatternFill("solid", start_color="C6EFCE", end_color="C6EFCE")


def f(bold=False, color="000000", size=10, italic=False):
    return Font(name="Arial", bold=bold, color=color, size=size, italic=italic)


def head(ws, row, labels):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill, c.border = f(True, "FFFFFF"), HEAD, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 36


def title(ws, text, sub):
    ws["A1"], ws["A2"] = text, sub
    ws["A1"].font, ws["A2"].font = f(True, NAVY, 14), f(italic=True, color="595959")


def put(ws, r, c, v, inp=False, center=False, bold=False, wrap=True):
    x = ws.cell(row=r, column=c, value=v)
    x.font = f(bold=bold, color=BLUE if inp else "000000")
    x.border = BOX
    x.alignment = Alignment(wrap_text=wrap, vertical="top", horizontal="center" if center else None)
    return x


# id, function, name, covers, current, target, criticality, evidence, action, owner, quarter, 800-53, soc2
cats = [
    ("GV.OC", "Govern", "Organizational Context", "Understanding the mission, stakeholders and legal duties that shape cybersecurity decisions.", 3, 4, 2,
     "ERP, payroll and customer database dependencies are documented in the PASTA threat model and vulnerability assessment scope. Regulatory obligations are not listed in one place.",
     "Document legal, regulatory and contractual obligations and map them to systems.", "Compliance Lead", "2027 Q1", "PM, PL", "CC1, CC2"),
    ("GV.RM", "Govern", "Risk Management Strategy", "Setting risk appetite, priorities and how risk is accepted.", 2, 4, 2,
     "A scored risk register with rating bands exists (risk-register-excel). No approved risk appetite statement or formal risk acceptance process.",
     "Approve a risk appetite statement and a risk acceptance procedure with sign-off levels.", "CISO", "2027 Q1", "PM, RA", "CC3"),
    ("GV.RR", "Govern", "Roles, Responsibilities, Authorities", "Defining who is accountable for cybersecurity and risk.", 2, 4, 2,
     "Risk owners are named in the register. No documented cybersecurity roles or responsibility matrix beyond risk ownership.",
     "Publish a RACI for security, risk and compliance activities.", "CISO", "2027 Q1", "PM, PS", "CC1"),
    ("GV.PO", "Govern", "Policy", "Written policy that is communicated, maintained and enforced.", 2, 4, 2,
     "The employee guide cites three policies (report phishing, no credentials by link, IT never asks for passwords). No full policy set with annual review.",
     "Create and approve access control, change management, incident response and acceptable use policies; review annually.", "Compliance Lead", "2027 Q2", "PL, PM", "CC5"),
    ("GV.OV", "Govern", "Oversight", "Leadership review of cybersecurity strategy and results.", 1, 3, 2,
     "Risk summaries can be produced, but there is no scheduled report of risk posture to leadership.",
     "Present the risk summary and top risks to leadership each quarter.", "CISO", "2027 Q1", "PM, CA", "CC4"),
    ("GV.SC", "Govern", "Cybersecurity Supply Chain Risk Management", "Managing cybersecurity risk from suppliers and service providers.", 2, 4, 3,
     "Several critical vendors were not reassessed in over 18 months (risk R-15). Vendor tiering and review are defined in third-party-risk-assessment.",
     "Adopt vendor tiering, a 12-month reassessment cycle and SOC 2 report review.", "Procurement", "2027 Q1", "SR, SA", "CC9"),
    ("ID.AM", "Identify", "Asset Management", "Knowing which systems, data and people the organization depends on.", 2, 4, 3,
     "Core systems are known. A legacy file server with a shared admin account (R-12) was found outside normal administration, so the inventory is incomplete.",
     "Build and reconcile an asset inventory with owner and data classification for every system.", "IT Infrastructure", "2027 Q1", "CM, PM", "CC6"),
    ("ID.RA", "Identify", "Risk Assessment", "Identifying and rating cybersecurity risk.", 3, 4, 3,
     "NIST SP 800-30 style assessment, PASTA threat model, ITGC testing and a scored register exist. Assessments are one-time, not on a set cycle.",
     "Run risk assessments on a fixed annual cycle and after major changes.", "CISO", "2027 Q2", "RA, CA", "CC3"),
    ("ID.IM", "Identify", "Improvement", "Learning from assessments and incidents to improve.", 2, 3, 2,
     "Incident journal reflections and ITGC recommendations capture lessons. There is no tracker that follows each improvement to closure.",
     "Track every audit and incident recommendation to closure in the risk register.", "Compliance Lead", "2027 Q2", "CA, PM, IR", "CC4"),
    ("PR.AA", "Protect", "Identity Management, Authentication, Access Control", "Managing identities, credentials and access rights.", 1, 4, 3,
     "ITGC tests found terminated users still active (2 of 4), overdue reviews (2 of 14) and a segregation of duties conflict. VPN has no MFA (R-10) and a legacy server uses shared admin credentials (R-12).",
     "Automate deprovisioning from HR, enforce MFA, remove shared admin accounts, and run semi-annual access reviews.", "IT Security", "2026 Q4", "AC, IA, PS", "CC6"),
    ("PR.AT", "Protect", "Awareness & Training", "Training people to recognize and report threats.", 3, 4, 2,
     "A plain-language phishing and malware guide exists. No phishing simulations or completion and click-rate metrics.",
     "Run quarterly phishing simulations and report click and report rates.", "IT Security", "2027 Q1", "AT", "CC1, CC2"),
    ("PR.DS", "Protect", "Data Security", "Protecting data at rest, in transit and in use.", 1, 4, 3,
     "Payroll data sat in unencrypted spreadsheets (R-14), database connections use deprecated SSL (R-08), and the customer database was reachable from the internet (R-05).",
     "Encrypt sensitive data, enforce TLS 1.2 or higher, restrict database access to an allow-list.", "IT Infrastructure", "2026 Q4", "SC, MP", "CC6"),
    ("PR.PS", "Protect", "Platform Security", "Secure configuration, patching and change control for systems.", 2, 4, 3,
     "A critical ERP vulnerability is unpatched (R-09) and 4 of 8 production changes lacked independent approval (ITGC test, D-04).",
     "Patch the ERP, block unapproved changes in the ticketing system, and set patch SLAs by severity.", "IT Infrastructure", "2026 Q4", "CM, SI, SA", "CC7, CC8"),
    ("PR.IR", "Protect", "Technology Infrastructure Resilience", "Resilient architecture, backups and capacity.", 3, 4, 3,
     "A full restoration test of the financial database was completed (R-13 closed). Public exposure of the customer database remains an architecture gap.",
     "Segment the database tier and repeat restoration tests every year.", "IT Infrastructure", "2027 Q1", "CP, SC, PE", "CC6, A1"),
    ("DE.CM", "Detect", "Continuous Monitoring", "Monitoring systems and networks for anomalies.", 2, 4, 3,
     "A Python script detects brute force, off-hours and geographic anomalies in exported authentication logs. It is not continuous or alerting.",
     "Centralize logs in a SIEM and schedule detections with alerts to an on-call owner.", "IT Security", "2027 Q2", "AU, SI, CA", "CC7"),
    ("DE.AE", "Detect", "Adverse Event Analysis", "Analyzing events to decide whether they are incidents.", 3, 4, 2,
     "Indicator analysis (Pyramid of Pain, VirusTotal) and packet analysis comparisons show analysis skills. No written triage criteria across event types.",
     "Define triage criteria and severity levels for alerts.", "IT Security", "2027 Q1", "AU, IR, SI", "CC7"),
    ("RS.MA", "Respond", "Incident Management", "Running the incident response process.", 2, 4, 3,
     "A phishing playbook exists. There is no tested ransomware runbook (R-11).",
     "Write and tabletop-test a ransomware runbook; extend the playbook set to other incident types.", "IT Security", "2027 Q1", "IR", "CC7"),
    ("RS.AN", "Respond", "Incident Analysis", "Investigating incidents and preserving evidence.", 3, 4, 2,
     "The incident handler's journal documents investigation steps and NIST lifecycle mapping. Evidence handling is not formalized.",
     "Add an evidence handling and chain of custody procedure.", "IT Security", "2027 Q2", "IR, AU", "CC7"),
    ("RS.CO", "Respond", "Incident Response Reporting & Communication", "Coordinating and reporting during incidents.", 2, 3, 2,
     "The phishing playbook includes escalation. There is no external communication or regulator notification plan.",
     "Add a communication plan covering leadership, customers, regulators and law enforcement.", "Compliance Lead", "2027 Q2", "IR", "CC2, CC7"),
    ("RS.MI", "Respond", "Incident Mitigation", "Containing and eradicating incidents.", 2, 3, 2,
     "Containment steps exist for phishing and malware (disconnect, do not self-remediate, call IT). Not tested for ransomware.",
     "Test containment steps during the ransomware tabletop exercise.", "IT Security", "2027 Q1", "IR", "CC7"),
    ("RC.RP", "Recover", "Incident Recovery Plan Execution", "Restoring systems and data after an incident.", 2, 4, 3,
     "Backup restoration has been tested once. No documented recovery plan with recovery time and recovery point objectives.",
     "Document recovery time and recovery point objectives and a prioritized recovery plan for ERP and payroll.", "IT Infrastructure", "2027 Q2", "CP, IR", "CC7, A1"),
    ("RC.CO", "Recover", "Incident Recovery Communication", "Communicating recovery progress and lessons.", 1, 3, 1,
     "No recovery communication procedure is documented.",
     "Add recovery status updates and a post-incident report template to the incident plan.", "Compliance Lead", "2027 Q3", "CP, IR", "CC2"),
]
FIRST, LAST = 5, 5 + len(cats) - 1   # 5..26

wb = Workbook()

# ---------------------------------------------------------------- Scope
sc = wb.active
sc.title = "Scope"
title(sc, "NIST CSF 2.0 Gap Assessment", "Fictional mid-sized healthcare services company. Scores are my judgments from the evidence in each row.")
info = [
    ("Purpose", "Compare current cybersecurity practice with a target for each of the 22 categories in NIST CSF 2.0, rank the gaps, and set a roadmap."),
    ("Organization", "A fictional regional healthcare services company with a financial ERP, a payroll system and a customer database. About 900 staff."),
    ("Method", "Each category gets a current and target maturity (1 to 5), a business criticality (1 to 3), and evidence. Gap = target minus current. "
               "Priority score = gap x criticality. Tiers come from the thresholds on the Scale sheet."),
    ("Evidence", "Current scores are based on findings from my other portfolio projects: ITGC testing, risk register, vulnerability assessment, PASTA threat model, "
                 "incident handler's journal, log anomaly detection and the security awareness guide."),
    ("Maturity scale", "The 1 to 5 scale is my own self-assessment scale. It is not the CSF Implementation Tiers, which describe an organization's overall risk governance."),
    ("Mappings", "The 800-53 family and SOC 2 criteria columns are indicative mappings from my own reading. Confirm them with NIST's CSF 2.0 Reference Tool before relying on them."),
    ("Sheets", "Assessment (inputs and formulas), Summary (function averages and charts), Roadmap (top 8 priorities), Scale (definitions and thresholds)."),
    ("Colour key", "Blue text is an input or judgment. Black text is a formula."),
    ("Data", "All data is synthetic. The organization and its findings are fictional."),
]
for i, (k, v) in enumerate(info):
    sc.cell(row=4 + i, column=1, value=k).font = f(True)
    c = sc.cell(row=4 + i, column=2, value=v)
    c.font, c.alignment = f(), Alignment(wrap_text=True, vertical="top")
    sc.cell(row=4 + i, column=1).alignment = Alignment(vertical="top")
sc.column_dimensions["A"].width = 18
sc.column_dimensions["B"].width = 110
sc.sheet_view.showGridLines = False

# ---------------------------------------------------------------- Scale
sl = wb.create_sheet("Scale")
title(sl, "Scales and Thresholds", "Edit the blue cells to change tiers; every sheet updates.")
head(sl, 4, ["Priority tier", "Minimum score", "Meaning"])
for i, (n, v, m) in enumerate([("Immediate", 6, "Start now; owner and date required"),
                               ("Next", 4, "Plan for the next two quarters"),
                               ("Planned", 0, "Schedule with normal improvement work")]):
    put(sl, 5 + i, 1, n, bold=True)
    put(sl, 5 + i, 2, v, inp=True, center=True)
    put(sl, 5 + i, 3, m)
head(sl, 10, ["Maturity", "Level", "Meaning"])
for i, (lv, n, m) in enumerate([(1, "Initial", "Ad hoc or undocumented; depends on individuals"),
                                (2, "Developing", "Some practice exists but is incomplete or inconsistent"),
                                (3, "Defined", "Documented and followed in most areas"),
                                (4, "Managed", "Measured, reviewed and improved on a schedule"),
                                (5, "Optimized", "Continuously improved using metrics and automation")]):
    put(sl, 11 + i, 1, lv, center=True)
    put(sl, 11 + i, 2, n, bold=True)
    put(sl, 11 + i, 3, m)
head(sl, 18, ["Criticality", "Level", "Meaning"])
for i, (lv, n, m) in enumerate([(1, "Supporting", "Failure has limited effect on operations or data"),
                                (2, "Important", "Failure disrupts a business function or exposes some sensitive data"),
                                (3, "Critical", "Failure stops financial or clinical operations or exposes regulated data")]):
    put(sl, 19 + i, 1, lv, center=True)
    put(sl, 19 + i, 2, n, bold=True)
    put(sl, 19 + i, 3, m)
sl.column_dimensions["A"].width = 16
sl.column_dimensions["B"].width = 16
sl.column_dimensions["C"].width = 70

# ---------------------------------------------------------------- Assessment
ws = wb.create_sheet("Assessment")
title(ws, "CSF 2.0 Assessment", "22 categories. Blue cells are inputs. Gap, score, tier and rank key are formulas.")
head(ws, 4, ["Function", "Category", "Name", "What it covers", "Current (1-5)", "Target (1-5)", "Gap", "Criticality (1-3)",
             "Priority score", "Priority tier", "Evidence and rationale", "Recommended action", "Owner", "Target quarter",
             "800-53 families (indicative)", "SOC 2 criteria (indicative)", "Rank key (helper)"])
for i, c in enumerate(cats):
    r = FIRST + i
    cid, fn, nm, cov, cur, tgt, crit, ev, act, own, qtr, f53, soc = c
    put(ws, r, 1, fn, inp=True, center=True)
    put(ws, r, 2, cid, inp=True, center=True, bold=True)
    put(ws, r, 3, nm, inp=True)
    put(ws, r, 4, cov, inp=True)
    put(ws, r, 5, cur, inp=True, center=True)
    put(ws, r, 6, tgt, inp=True, center=True)
    put(ws, r, 7, f"=MAX(F{r}-E{r},0)", center=True)
    put(ws, r, 8, crit, inp=True, center=True)
    put(ws, r, 9, f"=G{r}*H{r}", center=True, bold=True)
    put(ws, r, 10, f'=IF(I{r}>=Scale!$B$5,Scale!$A$5,IF(I{r}>=Scale!$B$6,Scale!$A$6,Scale!$A$7))', center=True, bold=True)
    put(ws, r, 11, ev, inp=True)
    put(ws, r, 12, act, inp=True)
    put(ws, r, 13, own, inp=True, center=True)
    put(ws, r, 14, qtr, inp=True, center=True)
    put(ws, r, 15, f53, inp=True, center=True)
    put(ws, r, 16, soc, inp=True, center=True)
    x = put(ws, r, 17, f"=I{r}+(5-E{r})/10+(100-ROW())/100000", center=True)
    x.number_format = "0.00000"
    x.font = f(color="7F7F7F")
    ws.row_dimensions[r].height = 78
for col, w in zip("ABCDEFGHIJKLMNOPQ", (10, 9, 26, 32, 9, 9, 7, 11, 9, 11, 58, 44, 16, 10, 13, 13, 10)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "D5"
ws.auto_filter.ref = f"A4:Q{LAST}"
dvs = DataValidation(type="whole", operator="between", formula1="1", formula2="5", showErrorMessage=True,
                     error="Enter a whole number from 1 to 5.")
dvc = DataValidation(type="whole", operator="between", formula1="1", formula2="3", showErrorMessage=True,
                     error="Enter 1, 2 or 3.")
ws.add_data_validation(dvs)
ws.add_data_validation(dvc)
dvs.add(f"E{FIRST}:F{LAST}")
dvc.add(f"H{FIRST}:H{LAST}")
for word, fill in (("Immediate", RED), ("Next", AMB), ("Planned", GRN)):
    ws.conditional_formatting.add(f"J{FIRST}:J{LAST}", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))
ws.conditional_formatting.add(f"G{FIRST}:G{LAST}", ColorScaleRule(start_type="num", start_value=0, start_color="FFFFFF",
                                                                 end_type="num", end_value=3, end_color="F4B183"))
RN = lambda col: f"Assessment!${col}${FIRST}:${col}${LAST}"

# ---------------------------------------------------------------- Summary
sm = wb.create_sheet("Summary")
title(sm, "Summary by Function", "Averages and counts are formulas over the Assessment sheet.")
head(sm, 4, ["Function", "Categories", "Avg current", "Avg target", "Avg gap", "Immediate", "Next", "Planned"])
funcs = ["Govern", "Identify", "Protect", "Detect", "Respond", "Recover"]
for i, fn in enumerate(funcs):
    r = 5 + i
    put(sm, r, 1, fn, bold=True)
    put(sm, r, 2, f"=COUNTIF({RN('A')},A{r})", center=True)
    for col, src in ((3, "E"), (4, "F"), (5, "G")):
        put(sm, r, col, f"=AVERAGEIF({RN('A')},$A{r},{RN(src)})", center=True).number_format = "0.0"
    for col, tier in ((6, "Immediate"), (7, "Next"), (8, "Planned")):
        put(sm, r, col, f'=COUNTIFS({RN("A")},$A{r},{RN("J")},"{tier}")', center=True)
put(sm, 11, 1, "All functions", bold=True)
put(sm, 11, 2, "=SUM(B5:B10)", center=True, bold=True)
for col, src in ((3, "E"), (4, "F"), (5, "G")):
    put(sm, 11, col, f"=AVERAGE({RN(src)})", center=True, bold=True).number_format = "0.0"
for col in (6, 7, 8):
    L = "FGH"[col - 6]
    put(sm, 11, col, f"=SUM({L}5:{L}10)", center=True, bold=True)
sm["A13"] = "Categories with a gap of 2 or more:"
sm["A13"].font = f(True)
put(sm, 13, 5, f'=COUNTIF({RN("G")},">=2")', center=True, bold=True)
for col, w in zip("ABCDEFGH", (34, 11, 12, 12, 10, 11, 8, 10)):
    sm.column_dimensions[col].width = w
sm.sheet_view.showGridLines = False

rc = RadarChart()
rc.type = "marker"
rc.title = "Average maturity by function: current vs target"
rc.add_data(Reference(sm, min_col=3, max_col=4, min_row=4, max_row=10), titles_from_data=True)
rc.set_categories(Reference(sm, min_col=1, min_row=5, max_row=10))
rc.y_axis.scaling.min, rc.y_axis.scaling.max = 0, 5
rc.series[0].graphicalProperties.line.solidFill = "C00000"
rc.series[1].graphicalProperties.line.solidFill = "2E75B6"
rc.height, rc.width = 9, 13
sm.add_chart(rc, "J4")
bc = BarChart()
bc.type = "bar"
bc.grouping = "stacked"
bc.overlap = 100
bc.title = "Categories by priority tier"
bc.add_data(Reference(sm, min_col=6, max_col=8, min_row=4, max_row=10), titles_from_data=True)
bc.set_categories(Reference(sm, min_col=1, min_row=5, max_row=10))
for s, colr in zip(bc.series, ("C00000", "FFC000", "70AD47")):
    s.graphicalProperties.solidFill = colr
bc.height, bc.width = 7.5, 13
bc.x_axis.delete = False
bc.y_axis.delete = False
sm.add_chart(bc, "J23")

# ---------------------------------------------------------------- Roadmap
rm = wb.create_sheet("Roadmap")
title(rm, "Roadmap: Top 8 Priorities", "Ranked by priority score. Ties go to the category with the lower current maturity, then Assessment sheet order.")
head(rm, 4, ["Rank", "Category", "Name", "Current", "Target", "Gap", "Criticality", "Score", "Tier", "Recommended action", "Owner", "Target quarter"])
for n in range(1, 9):
    r = 4 + n
    m = f"MATCH(LARGE({RN('Q')},{n}),{RN('Q')},0)"
    put(rm, r, 1, n, center=True, bold=True)
    for col, src in ((2, "B"), (3, "C"), (4, "E"), (5, "F"), (6, "G"), (7, "H"), (8, "I"), (9, "J"), (10, "L"), (11, "M"), (12, "N")):
        put(rm, r, col, f"=INDEX({RN(src)},{m})", center=col not in (3, 10))
    rm.row_dimensions[r].height = 44
for col, w in zip("ABCDEFGHIJKL", (7, 10, 34, 9, 9, 7, 11, 8, 11, 60, 17, 12)):
    rm.column_dimensions[col].width = w
for word, fill in (("Immediate", RED), ("Next", AMB), ("Planned", GRN)):
    rm.conditional_formatting.add("I5:I12", CellIsRule(operator="equal", formula=[f'"{word}"'], fill=fill))

wb._sheets = [wb[n] for n in ["Scope", "Assessment", "Summary", "Roadmap", "Scale"]]
wb.active = 1
wb.save(OUT)
print("saved", OUT)
