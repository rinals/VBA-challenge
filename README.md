# VBA-challenge
1. quarterly_calculations.bas
The VBA script that I expoerted from Microsoft Excel Macros.
It works in the following manner:
 - Loops through all the worksheets in the Workbook (excel file)
 - For each work sheet
     - Loops through all the rows to find when a ticker symbol changes
     - Keeps track of all opening and closing prices
     - Keeps a separate tickerCounter and increments on each new Ticker open-to-close result
     - Calculates quarterly change for the opening price at the beginning and the closing price at the end when a ticker symbol changes.
     - percentage for each tickerCounter row is calculated by deviding the quarterly change by the opening price
     - Total volume is incremented on each row processed and updated when ticker symbol changes
     - Greatest quarterly increase, decrease and totalVolume is tracked and updated in a variable with if-else comparison

2. Screen shots
 - 4 Screen shots for each Quarter in the provided Excel file.
 - The screenshots were captured after running the macro