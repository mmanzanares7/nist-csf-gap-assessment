# NIST CSF 2.0 Gap Assessment

## Project Description

This project is a gap assessment of a fictional healthcare services company against all 22 categories of the NIST Cybersecurity Framework 2.0. Each category gets a current maturity, a target maturity, a business criticality and written evidence. Formulas turn those into a gap, a priority score, a priority tier and a ranked roadmap. The current scores come from findings in my other portfolio projects (ITGC testing, risk register, vulnerability assessment, incident response and monitoring work), so the assessment shows how technical findings roll up into a framework view that management can act on. This is the kind of baseline assessment a GRC analyst or consultant delivers at the start of a security program.

## Files

| File | Purpose |
|---|---|
| [`csf_gap_assessment.xlsx`](csf_gap_assessment.xlsx) | Workbook: Scope, Assessment, Summary (charts), Roadmap and Scale sheets (226 formulas, no errors) |
| [`csf_gap_assessment_export.csv`](csf_gap_assessment_export.csv) | Assessment with calculated values, readable on GitHub |
| [`images/maturity_by_function.png`](images/maturity_by_function.png) | Current against target by CSF function |
| [`scripts/build_assessment.py`](scripts/build_assessment.py) | Python (openpyxl) script that generates the workbook |

## Method

- **Categories:** the 22 CSF 2.0 categories across Govern, Identify, Protect, Detect, Respond and Recover.
- **Scores:** current and target maturity on a 1 to 5 scale, and business criticality from 1 to 3. The maturity scale is my own self-assessment scale, not the CSF Implementation Tiers.
- **Formulas:** gap = target minus current. Priority score = gap x criticality. Tier is Immediate at 6 or more, Next at 4 or more, otherwise Planned. The thresholds are inputs on the Scale sheet.
- **Ranking:** by priority score, with ties going to the category with the lower current maturity.
- **Mappings:** each category lists indicative 800-53 control families and SOC 2 criteria groups. These come from my own reading and should be checked against NIST's CSF 2.0 Reference Tool.

## Results

Average maturity is 2.1 against an average target of 3.8, and 13 of 22 categories have a gap of 2 or more. Recover and Protect have the largest average gaps (2.0 each), followed by Govern (1.8).

![Average maturity by function, current vs target](images/maturity_by_function.png)

| Rank | Category | Current | Target | Score | Why it ranks here |
|---|---|---|---|---|---|
| 1 | PR.AA Identity, Authentication, Access Control | 1 | 4 | 9 | Terminated users still active, overdue reviews, segregation of duties conflict, no VPN MFA, shared admin credentials |
| 2 | PR.DS Data Security | 1 | 4 | 9 | Unencrypted payroll spreadsheets, deprecated SSL, internet-facing customer database |
| 3 | GV.SC Supply Chain Risk Management | 2 | 4 | 6 | Critical vendors not reassessed in over 18 months |
| 4 | ID.AM Asset Management | 2 | 4 | 6 | Legacy server with shared admin account found outside normal administration |
| 5 | PR.PS Platform Security | 2 | 4 | 6 | Unpatched ERP vulnerability, changes without independent approval |
| 6 | DE.CM Continuous Monitoring | 2 | 4 | 6 | Log detections exist as a script, not continuous monitoring |
| 7 | RS.MA Incident Management | 2 | 4 | 6 | No tested ransomware runbook |
| 8 | RC.RP Recovery Plan Execution | 2 | 4 | 6 | No documented recovery objectives or plan |

Ranks 3 to 8 have the same score, so their order is a tie-break, not a judgment that one matters more than another. The roadmap gives each priority a recommended action, owner and target quarter.

## Connections to other projects

- [`itgc-controls-testing`](https://github.com/mmanzanares7/itgc-controls-testing) and [`sox-itgc-audit-workpapers`](https://github.com/mmanzanares7/sox-itgc-audit-workpapers) support the PR.AA and PR.PS scores.
- [`risk-register-excel`](https://github.com/mmanzanares7/risk-register-excel) supplies the risks cited for PR.DS, ID.AM, GV.SC and RS.MA.
- [`log-anomaly-detection`](https://github.com/mmanzanares7/log-anomaly-detection) supports the DE.CM score.
- [`third-party-risk-assessment`](https://github.com/mmanzanares7/third-party-risk-assessment) is the recommended action for GV.SC.

## Limitations

The organization and findings are fictional, and the scores are my judgments from the evidence in each row, not measurements. A real assessment would validate scores through interviews and evidence review, and would use the Implementation Tiers and target profiles agreed with leadership.

## Summary

I assessed a fictional company against every CSF 2.0 category, connected each score to specific evidence, and produced a ranked roadmap with owners and dates. Access control and data security come out as the first priorities, which matches the audit and risk findings in the rest of my portfolio.
