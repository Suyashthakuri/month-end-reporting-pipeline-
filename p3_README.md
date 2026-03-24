# Automated Month-End Reporting Pipeline
### Excel VBA automation replacing manual month-end consolidation — Financial Services Business Unit

---

## What this project does

This project automates the monthly expense reporting process for a financial services business unit. A VBA macro reads 500 raw transactions, validates data quality, categorises expenditure, calculates budget vs actual variances, flags items requiring review, and outputs a formatted P&L summary and executive dashboard — in under 60 seconds.

**The business problem it solves:** Finance teams spend 2–4 hours every month manually consolidating transaction exports, building variance tables, and formatting management reports. This pipeline reduces that to a single button click.

---

## Skills demonstrated

- **Excel VBA:** loops, conditionals, worksheet manipulation, data validation, automated formatting
- **Financial reporting:** P&L structure, budget vs actual variance, traffic-light KPI status
- **Process automation:** replacing manual month-end consolidation with a one-click pipeline
- **Financial services domain:** expense categorisation, cost centre reporting, approver workflow
- **Data quality:** automated validation flagging blank amounts and missing department codes

---

## How to use it

1. Download `raw_transactions.csv` and `month_end_report_template.xlsx`
2. Open the Excel file
3. Copy all data from `raw_transactions.csv`
4. Paste into the `Raw Data` sheet starting at row 3
5. Press `Alt + F11` to open the VBA Editor
6. Click `Insert → Module` and paste the VBA code from the `VBA Code` sheet
7. Close VBA Editor
8. Press `Alt + F8` → select `GenerateMonthEndReport` → click `Run`
9. Report generates automatically across all output sheets

---

## What the macro does — step by step

| Step | Action | Output |
|---|---|---|
| 1 | Validates all 500 transactions | Blank amounts and missing departments highlighted red |
| 2 | Aggregates by category | P&L Summary sheet populated with budget vs actual |
| 3 | Flags variance items | Variance Analysis sheet populated with review-required items |
| 4 | Updates dashboard | Executive Dashboard timestamp and KPIs refreshed |
| 5 | Completion message | Summary statistics displayed |

---

## Dataset

**500 transactions** across a 12-month financial year (FY2024) covering:

- 10 departments: Retail Banking, Business Banking, Wealth Management, Insurance, Markets, Operations, IT, HR, Finance, Risk
- 6 expense categories: Staff Costs, Technology, Occupancy, Professional Fees, Marketing, Operations
- Fields: transaction ID, month, department, category, subcategory, vendor, actual amount, budget amount, variance, status, approver

**Transaction status flags:**
- `On Track` — variance within 5% of budget
- `Monitor` — variance between 5–15%
- `Review Required` — variance exceeds 15%
- `Approved` — pre-approved line items (e.g. salaries)

---

## Output sheets

**Raw Data** — 500 transaction records with status colour coding (green/amber/red)

**P&L Summary** — Category-level budget vs actual with variance % and traffic-light status

**Variance Analysis** — All `Review Required` transactions sorted by impact size

**Executive Dashboard** — KPI cards (total expenditure, items on track, flagged items) + category breakdown table for CFO/board reporting

---

## Repository structure

```
month-end-reporting-pipeline/
├── README.md
├── data/
│   └── raw_transactions.csv
├── excel/
│   └── month_end_report_template.xlsx
└── screenshots/
    ├── 01_raw_data.png
    ├── 02_pl_summary.png
    ├── 03_variance_analysis.png
    └── 04_executive_dashboard.png
```

---

## Tools used

- Microsoft Excel — workbook design and output formatting
- Excel VBA — automation macro
- CSV — raw data input source

---

*Built as part of a reporting analytics portfolio targeting Reporting Analyst and BI Analyst roles in Sydney, Australia.*
