
<img width="1525" height="677" alt="Raw_in_excel" src="https://github.com/user-attachments/assets/3b45908b-4aab-4ca1-bb9a-00612aaead99" /> <img width="1119" height="608" alt="02_Finance Dashboard" src="https://github.com/user-attachments/assets/373f2732-c1a0-4533-abda-3ef568e68531" />


# Finance Reporting & Analytics (Power BI)

Portfolio project simulating a real FP&A reporting workflow: transform messy P&L spreadsheets into a clean model and build a leadership-ready dashboard (Actual vs Budget vs Prior Year, trends, variances, drill-through).

## What’s included
- **Power BI semantic model (star schema)**: FactPnL + DimDate + DimAccount  
- **Core DAX measures**: Actual, Budget, Variance ($ / %), Revenue/COGS/OPEX/GP%/EBITDA, Prior Year (PY)
- **Report pages**
  - **Finance Dashboard**: KPI cards + Revenue/EBITDA trends + summary P&L matrix
  - **P&L Detail (Drill-through)**: line-level breakdown
  - **Data Quality / QA**: month/scenario coverage checks + refresh timestamp
- **Excel KPI model**: `excel/Excel_KPI_Model.xlsx` (dropdown MonthEndDate + SUMIFS-based KPIs + variances)
- **SQL examples**: KPI pack + QA queries (as if the data lived in a database)

## Data
Source workbook: `data/Financial_Raw_Data_2020_2024.xlsx`  
Contains raw P&L sheets (Actual/Budget 2020–2024) plus structured dimension/fact tables for reference.  
**All data is synthetic** (generated for learning/portfolio use).

## How to run
1. Open `powerbi/Finance_Dashboard.pbix`
2. If prompted, update the Excel path parameter to `data/Financial_Raw_Data_2020_2024.xlsx`
3. Refresh: **Home → Refresh**

## Folder structure
- `powerbi/` → PBIX + `screenshots/`
- `data/` → Excel source workbook
- `excel/` → `Excel_KPI_Model.xlsx` (SUMIFS/variance model)
- `scripts/` → python utilities (inspection/validation)
- `sql/` → example finance queries (KPI pack + QA checks)
- `docs/` → KPI glossary / refresh notes (optional)

## Skills demonstrated
Power Query (cleaning/unpivot), data modeling, DAX (variance & time intelligence), dashboard design, QA checks, Excel formulas (SUMIFS/XLOOKUP), SQL (CTEs/aggregations).
