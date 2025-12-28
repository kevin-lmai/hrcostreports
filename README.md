# HrCostCenter app

## this is app to be used by a hospital for internal HR record handling


### Input file: Excel file with 3-4 sheets. 
- Sheet 1 : base HR records for a month
- Sheet 2 : expanded HR records for corresponding month
- Sheet 3 : A table of Cost Centre Name with code
- Sheet 4 : A table of list order in report of Staff Category. If omitted, order will be in alphabetical order

### Output file:
- Department reports generated.

# Data constraint
- FTE in Override Sheet should be equal to 1.0 (100%). Report will be generated but issue number will be shown
- Staff Number found in Override Sheet but not found in Base Sheet, will also be shown
