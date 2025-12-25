# Power BI build plan (high-level)

## Recommended tables to load
From the workbook:
- DimDate
- DimAccount
- DimEntity
- DimCostCenter
- DimProduct
- FactPnL_Monthly (clean fact)
- (Optional) FactSales_Operational, FactHeadcount (for driver visuals)

## Two ways to do the project
### A) Realistic test style (Power Query heavy)
Load **messy sheets**:
- Actual P&L 2024
- Budget P&L 2024
Clean + unpivot into FactPnL, then build the model.

### B) Production style (already normalized)
Load `FactPnL_Monthly` + dimensions directly to focus on modeling + DAX + visuals.

## Visual wireframe
Top row (cards):
- Revenue (Actual), Gross Profit, GP%, EBITDA, Variance, Variance %

Middle row:
- Line: Actual vs Budget by month
- Waterfall: variance bridge by PnLGroup or Level1

Bottom row:
- Matrix P&L (Account hierarchy) with Actual/Budget/Var/Var%
- Bar: Top 10 variance accounts
