Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

        startTime = Timer

    '1) format output sheet
    Worksheets("All Stocks Analysis").Activate
    
    'add a title
    Range("A1").Value = "All Stocks (" + yearValue + " )"
    
    'create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

'2) Initialize an array of all tickers.
Dim tickers(11) As String
    
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

'====================================

'3a) Initialize variables for the starting price and ending price
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
'3b) Activate the data worksheet

Worksheets(yearValue).Activate

'3c) Find the number of rows to loop over.

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'====================================

'4) Loop through the tickers. (outterloop)

For i = 0 To 11

    ticker = tickers(i)
    totalVolume = 0
    
        '====================================
    
        '5) Loop through rows in the date
        Worksheets(yearValue).Activate
        
        For J = 2 To RowCount

                '5a) Find the total volume for the current ticker.

                If Cells(J, 1).Value = ticker Then
                
                    totalVolume = totalVolume + Cells(J, 8).Value
                
                End If
    
                '5b) Find the starting price for the current ticker.

                If Cells(J - 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then

                    startingPrice = Cells(J, 6).Value
                
                End If

                '5c) Find the ending price for the current ticker.

                If Cells(J + 1, 1).Value <> ticker And Cells(J, 1).Value = ticker Then

                    endingPrice = Cells(J, 6).Value
                
                End If

        '====================================

        Next J
        '6) Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
      
Next i
    
endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


End Sub


Sub formatAllStocksAnalysisTable()

'Formatting
Worksheets("All Stocks Analysis").Activate
Range("A3:C3").Font.Bold = True
Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A3:C3").Font.Size = 12
Range("B4:B15").NumberFormat = "#,##0"
Range("C4:C15").NumberFormat = "0.00%"
Columns("B").AutoFit


dataRowStart = 4
dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then
                'Color the cell green
                Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then
                'Color the cell red
                Cells(i, 3).Interior.Color = vbRed
            
        Else
                'Clear the cell color
                Cells(i, 3).Interior.Color = xlNone

        End If


    Next i

    
End Sub


Sub ClearWorksheet()

Cells.Clear

End Sub


Sub yearValueAnalysis()

yearValue = InputBox("What year would you like to run the analysis on?")

End Sub
