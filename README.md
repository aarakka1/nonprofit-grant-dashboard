# Nonprofit Grant Tracking Dashboard

An interactive Tableau dashboard built for a nonprofit organization to track grant utilization, program spending, and financial health across multiple funding sources and fiscal years.

---

## Problem Statement

The organization managed two federal grants (DFC and Alliance/ACHC) alongside seven additional programs, tracking thousands of transactions across multiple fiscal years. Financial data lived in two separate Excel files:

- A **bank transaction ledger** with 1,200+ rows of raw transactions allocated across 9 programs
- A **budget analysis template** with complex multi-header structure tracking monthly actuals vs. budgeted amounts by expense category

There was no unified visual tool to answer key questions like:
- How much of each grant has been spent vs. what remains?
- Which expense categories are consuming the most funding?
- How does spending compare month-over-month and year-over-year?
- Which programs are running a surplus or deficit?

---

## Solution

A multi-source Tableau dashboard combining both data files into a single interactive view, with a Python ETL script to transform the complex budget template into an analysis-ready format.

---

## Dashboard Features

### Running Balance Dashboard
- **KPI Cards** — Total Income, Total Expenses, Net Balance, Grant Amount, Remaining Balance
- **Grant Utilization Progress** — Side-by-side bar chart showing amount spent vs. grant ceiling for DFC ($125K) and Alliance ($38K), with overage labels
- **Monthly Activity** — Dual-axis bar chart showing income vs. expenses by month across the full timeline (Jul 2022 – Feb 2026)
- **Running Balance** — Cumulative net position line chart with zero reference baseline
- **Program Filter** — Single-select filter to switch between all 9 programs

### Budget Analysis Dashboard (Carol's Template)
- **Category Spending** — Horizontal bar chart comparing DFC vs. ACHC spending across 6 expense categories (Personnel, Fringe Benefits, Travel, Supplies, Contractor, Other)
- **Monthly Category Spending** — Multi-line chart showing spending trends by category over the fiscal year
- **Fiscal Year Comparison** — Side-by-side comparison of 2024-2025 vs. 2025-2026 spending by category

---

## Technical Highlights

### Data Sources
| File | Description | Rows |
|------|-------------|------|
| Running Balance (XLSX) | Bank transactions allocated across 9 programs | 1,237 |
| Budget Template (XLSX) | Monthly actuals vs. budget by expense category | 65 rows × 36 cols |

### ETL Script (`reshape_carol.py`)
The budget template used a complex multi-row header structure with merged cells that Tableau couldn't interpret directly. The Python script:

- Reads both `2024-2025` and `2025-2026` sheets
- Identifies monthly columns vs. summary columns by parsing row 4 (program) and row 5 (month/label)
- Handles Personnel and Fringe Benefits specially — these categories only have year totals in the source data, so the script spreads them evenly across 12 months
- Adds `Category`, `Program`, `Fiscal Year`, and `Month` columns
- Outputs a clean flat table (`Carol_Clean.xlsx`) ready for Tableau

```
Input:  Multi-header Excel (36 columns, complex structure)
Output: Flat table with 6 columns, 1,739 rows
        [Fiscal Year | Category | Line Item | Program | Month | Amount]
```

### Calculated Fields in Tableau
Key calculated fields built to power the visualizations:

```
// Grant Remaining Safe — prevents negative display
IF MIN([Program]) = "DFC" THEN
  MAX(0, 125000.0 - SUM([Program Expenses]))
ELSEIF MIN([Program]) = "Achc" THEN
  MAX(0, 38000.0 - SUM([Program Expenses]))
ELSE 0
END

// Overage Amount — shows how much over budget
IF MIN([Program]) = "DFC" THEN
  SUM([Program Expenses]) - 125000.0
ELSEIF MIN([Program]) = "Achc" THEN
  SUM([Program Expenses]) - 38000.0
ELSE 0
END

// Program Expenses — fiscal year scoped
IF [Date] >= #2024-10-01# AND [Date] <= #2025-09-30# THEN
  IF ZN([Program Amount]) < 0
  THEN ABS([Program Amount])
  ELSE 0
  END
ELSE 0
END
```

---

## Key Findings (2024-2025 Fiscal Year)

| Grant | Budget | Spent | Status |
|-------|--------|-------|--------|
| DFC | $125,000 | $128,246 | Over by $3,246 |
| ACHC (Alliance) | $38,000 | ~$40,946 | Over by ~$2,946 |

**Top spending categories (DFC):**
1. Personnel — $37,466
2. Fringe Benefits — $23,868
3. Other — $12,808
4. Travel — $8,740
5. Supplies — $5,894
6. Contractor — $3,183

---

## Tools & Technologies

| Tool | Usage |
|------|-------|
| Python 3 | ETL script, data reshaping |
| pandas | Excel reading, data transformation |
| openpyxl | Excel file handling |
| Tableau Desktop | Dashboard creation, calculated fields |
| Microsoft Excel | Source data files |

---

## Setup & Usage

### Running the ETL Script

**Requirements:**
```bash
pip install pandas openpyxl
```

**Run:**
```bash
# Place reshape_carol.py in the same folder as Template_from_Carol.xlsx
python3 reshape_carol.py
# Output: Carol_Clean.xlsx
```

**Load into Tableau:**
1. Open Tableau Desktop
2. Connect to `Carol_Clean.xlsx` as a new data source
3. Connect to the running balance file as a second data source
4. Build sheets using the field structure described above

---

## Project Structure

```
├── README.md
├── reshape_carol.py          # ETL script to reshape budget template
├── screenshots/
│   ├── dashboard_overview.png
│   ├── grant_progress.png
│   ├── category_spending.png
│   └── monthly_activity.png
└── tableau/
    └── Program_Balance_Dashboard.twbx   # Packaged Tableau workbook (data anonymized)
```

---

## Notes

- Source Excel files are not included in this repository as they contain real financial data
- Program and organization names have been kept generic in this README for privacy
- The dashboard is designed to update automatically when source files are refreshed with new data
- Grant amounts ($125K DFC, $38K ACHC) are hardcoded as calculated fields and should be updated annually

---

## Author

Built as a data analytics consulting project for a nonprofit organization.  
Tools: Python · pandas · Tableau · Excel
