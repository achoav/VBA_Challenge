Attribute VB_Name = "Module1"
Sub DQAnalysis()
    'This Sub creates a summary for 2018 of the Total Daily Volume and Yearly Return for Stock "DQ"
    Worksheets("DQ Analysis").Activate
    'Changes the text on cell "A1" to "DAQO (Ticker: DQ"
    Range("A1").Value = "DAQO (Ticker: DQ"

    'Create a header row 3, columns A,B,C
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate
    
    'set initial volume to zero, for the stock analysis
    totalVolume = 0
    'setting data types to StartingPrice and EndingPrice
    
    Dim StartingPrice As Double
    Dim EndingPrice As Double
    
    'find the number if rows to loop over. RowCount will be the last row on the column
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows: starting the iterator at row 2 (row 1 is the label column)
    For i = 2 To RowCount
    
        If Cells(i, 1).Value = "DQ" Then
            'incease TotalVolume by the value in the current row and column 8 that contains the volume
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            StartingPrice = Cells(i, 6).Value
        End If
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
            EndingPrice = Cells(i, 6).Value
        End If
    Next i
    'Assign values to row 4
    
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (EndingPrice / StartingPrice) - 1
    
    Worksheets("All Stock Analysis").Activate
    'Changes the text on cell "A1" to "All Stocks (2018)"
    Range("A1").Value = "All Stocks (2018)"
    
End Sub

Sub AllStockAnalysis()
    'This Sub calculates the results for "All Stock Tickers" - Total Daily Volume and Yearly Returns
    'Further if formats negative returns in red and positive returns on green
'1. Format the output sheet on the "All Stocks Analysis" worksheet.
    '1.a. Activate the worksheet
    Dim startTime As Single
    Dim endTime As Single
    yearValue = InputBox("What year would you like to run the analysis on? (2017 or 2018)")
    startTimer = Timer
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    'Create a header row for A,B,C columns in row 3
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
'2. Initialize an array of all tickers.
    'Set tickers as a text type by saying String, and create an array of 12 tickers
    Dim tickers(12) As String
'3. Prepare for the analysis of tickers from 0 to 11.
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
'4. Initialize variables for the starting price and ending price.
    Dim StartingPrice As Single
    Dim EndingPrice As Single
'5. Activate the data worksheet.
    Worksheets(yearValue).Activate
'7. Find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
'8. Loop through tickers 0 thru 11
'Outer loop (via variable (i))
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
'Inner loop(via varibale (j))
    Worksheets("2018").Activate
        For j = 2 To RowCount
'9. Find the total volume for the current ticker..
    'This if will increase the totalVolume if the ticker matches the ticker on inner loop ticker=ticker(i)
        If Cells(j, 1).Value = ticker Then
        totalVolume = totalVolume + Cells(j, 8).Value
        End If
'10. Find the starting price for the current ticker.
    'Find if the row matches the selected ticker and with an If condition to determine it is the firstrow=startingPrice
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            StartingPrice = Cells(j, 6).Value
        End If
'11. Find the ending price for the current ticker.
    'Find if the row matches the selected ticker and with an If condition determine if it is the lastrow=endingPrice
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            EndingPrice = Cells(j, 6).Value
        End If
    'Close the outer loop that goes thru the rows
    Next j
'12. Output the data for the current ticker.    'Create a title for the worksheet "All Stocks 2018"
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    'Return on Investment
    Cells(4 + i, 3).Value = EndingPrice / StartingPrice - 1
    'Format the return on Investment, Green if Positive, Red if Negative and None if equal to zero
    If Cells(4 + i, 3) > 0 Then
        'Change cell color to green
        Cells(4 + i, 3).Interior.Color = vbGreen
        ElseIf Cells(4 + i, 3) < 0 Then
        'Change cell color to red
        Cells(4 + i, 3).Interior.Color = vbRed
        Else
        'Clear the cell color
        Cells(4 + i, 3).Interior.Color = xlNone
    End If
'Close the outer loop with the variable (i)
Next i
'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year" & (yearValue)
    If yearValue = "2017" Then
        Range("L3").Value = endTime
    Else
        Range("L4").Value = endTime
    End If
' Add section to display greatest percent increase, greatest percent decrease, and greatest total volume for each year.
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Get the last row
    lastRowState = Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Initialize variables and set values of variables initially to the first row in the list.
    greatest_percent_increase = Cells(4, 3).Value
    greatest_percent_increase_ticker = Cells(4, 1).Value
    greatest_percent_decrease = Cells(4, 3).Value
    greatest_percent_decrease_ticker = Cells(4, 1).Value
    greatest_stock_volume = Cells(4, 2).Value
    greatest_stock_volume_ticker = Cells(4, 1).Value
    
    
    ' skipping the header row, loop through the list of tickers.
    For i = 4 To 15
    
        ' Find the ticker with the greatest percent increase.
        If Cells(i, 3).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 3).Value
            greatest_percent_increase_ticker = Cells(i, 1).Value
        End If
        
        ' Find the ticker with the greatest percent decrease.
        If Cells(i, 3).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 3).Value
            greatest_percent_decrease_ticker = Cells(i, 1).Value
        End If
        
        ' Find the ticker with the greatest stock volume.
        If Cells(i, 2).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 2).Value
            greatest_stock_volume_ticker = Cells(i, 1).Value
        End If
        
    Next i
    
    ' Add the values for greatest percent increase, decrease, and stock volume to each worksheet.
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume


End Sub

Sub formatAllStocksAnalysiTable()
'This Sub formats with an underline and a Currency and % format
'Formatting - Bold
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
'Formatting - Border on the bottom
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
'Formatting - Column B in currency and column C in percentages
    Range("B4:B15").NumberFormat = "$ #,##0.00"
    Range("C4:C15").NumberFormat = "0.00%"
'Formating - Colum B to autofit
    Columns("B").AutoFit
'Attention: Format the return cell in the inner loop, green for positive return, and red for negative (this was done on the inner loop above)
End Sub

Sub ClearWorksheet()
    Sheets("All Stocks Analysis").Cells.Clear
    
End Sub

