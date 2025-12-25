# KPI / Measure Glossary

## Base measures
- Actual Amount: sum of FactPnL[Amount] where Scenario = Actual
- Budget Amount: sum of FactPnL[Amount] where Scenario = Budget
- PY: prior year version using SAMEPERIODLASTYEAR over DimCalendar[Date]

## KPIs
### Revenue
Accounts 4000–4999  
- Revenue (Actual), Revenue (Budget), Revenue (PY)
- Revenue Var $ = Actual - Budget
- Revenue Var % = Var $ / Budget
- Revenue YoY $ = Actual - PY
- Revenue YoY % = YoY $ / PY

### COGS
Accounts 5000–5999  
- COGS (Actual/Budget/PY)

### Gross Profit
Gross Profit = Revenue + COGS  
- GP (Actual/Budget/PY)

### GP% (Margin %)
GP% = Gross Profit / Revenue  
- GP% (Actual/Budget/PY)
- GP% Var (pp) = Actual - Budget
- GP% YoY (pp) = Actual - PY

### OPEX
Accounts 6000–6999  
- OPEX (Actual/Budget/PY)

### EBITDA
EBITDA = Gross Profit + OPEX  
- EBITDA (Actual/Budget/PY)
- EBITDA Var $ / Var %
- EBITDA YoY $ / YoY %

## Display notes
- Costs are negative by design; dashboard cards can use ABS() display measures.
- Percentage-point variance is used for margin KPIs (GP%).
