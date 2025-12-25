# Data Model Overview

## Purpose
Finance reporting semantic model for Actual vs Budget vs Prior Year analysis (P&L).

## Fact Table
### FactPnL
**Grain:** Account × MonthEndDate × Scenario  
**Key fields:** Account, MonthEndDate, Scenario  
**Measure field:** Amount

## Dimensions
### DimAccount
- Account (key)
- Line Item
- SortOrder

### DimCalendar
- Date (MonthEndDate used in Fact)
- Year
- MonthNumber
- MonthName
- YearMonth
- Quarter

### PnL Layout (calculated table)
Defines the reporting structure for P&L lines (Revenue, COGS, OPEX, GP, EBITDA, Net Income), including sort order and line type (Account/Subtotal/Calc).

## Relationships
- FactPnL[Account] → DimAccount[Account] (Many-to-one, single direction)
- FactPnL[MonthEndDate] → DimCalendar[Date] (Many-to-one, single direction)

## Notes
- Costs are stored as negative values (COGS/OPEX), which is standard in finance models.
- Display measures may use ABS() for executive KPI cards.
