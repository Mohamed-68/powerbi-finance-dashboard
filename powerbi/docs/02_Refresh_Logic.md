# Refresh Logic (Power Query)

## Inputs
- Excel workbook with P&L sheets (Actual and Budget) across years.

## Transformation pattern
1. Load sheet
2. Remove non-data header rows and notes
3. Promote headers
4. Rename columns (Account, Line Item, months)
5. Unpivot month columns to create long format:
   - MonthName
   - Amount
6. Clean Amount:
   - remove commas / parentheses
   - convert to number
7. Add date attributes:
   - MonthNumber
   - MonthEndDate
   - Year
   - YearMonth
   - Quarter
8. Add Scenario column (Actual/Budget)
9. Append all years into FactPnL

## Output
- FactPnL table ready for the semantic model.
- DimAccount created from distinct accounts + sort order.
- DimCalendar created as a proper date dimension.

## Monthly update
- Add the new monthly file/sheet
- Refresh Power Query
- QA page checks confirm month/scenario completeness.
