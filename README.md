# VBA_Challenge
VBA Challenge
A VBA script used to analyze real stock market data in a Microsoft Excel workbook.

# Background
For this project, I created a VBA (Visual Basic) script to analyze some stock market data. The data is inside a Microsoft Excel workbook and includes stock data for three years (2017 and 2018). Each year is a different tab/sheet inside the workbook.

# About the Script
You can find the script inside the VBAStocks folder of this repository. The script file is called AllStockAnalysisRefactored.bas and DQAnalysis.bas

-After you download and open up the All Stock Analysis Excel workbook, you can run the script by doing the following:
  - Click the Macro Button to "Run Analysis for All Stocks", which will run Sub(AllStockAnalysis.bas), when prompted just type either year "2017" or "2018"
  - Or, Click the Macro Button to run "Refactored - Run Analysis for All Stocks", which will run Sub(AllStocksAnalysisRefactore), when prompted just type either year "2017" or "2018"

- As the script runs, it is doing the following:
  It loops through all the stocks for one year for each run and takes the following information:
  - Ticker - The ticker symbol
  - Stock Volume - The total stock volume of the stock for that year
  - Return - Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

- It applies conditional formatting by highlighting positive yearly change values in green and negative yearly change values in red.
- Finally, it return the stock with the greatest percent increase, greatest percent decrease, and greatest total volume

# Sample Output
After the script has completed, go to the Excel workbook, and you should see the results of the script.

Screenshots are available in the VBA_Challenge/Resources folder of this repository.
