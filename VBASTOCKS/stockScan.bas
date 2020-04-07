Attribute VB_Name = "Module1"
Sub stockScan()

    Dim tickerSymbol As String
    Dim nextTickerSymbol As String
    Dim dateTimeRange As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim SummaryRow As Integer
    
    For Each ws In Worksheets
        'setup new headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'hard_solution
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greated % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        SummaryRow = 2
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        openPrice = ws.Cells(2, 3)
    
        
        For i = 2 To lastRow
            Dim yearlyChange As Double
            Dim yearlyPercentChange As Double
            tickerSymbol = ws.Cells(i, 1).Value
            dateTimeRange = ws.Cells(i, 2).Value
            nextTickerSymbol = ws.Cells(i + 1, 1).Value
           
            
            If tickerSymbol <> nextTickerSymbol Then
                ws.Cells(SummaryRow, 9).Value = tickerSymbol
                closePrice = ws.Cells(i, 6).Value
                yearlyChange = closePrice - openPrice
                
                If openPrice = 0 Then
                    yearlyPercentageChange = 0
                Else
                    yearlyPercentageChange = (yearlyChange / openPrice)
                End If
                
                ws.Cells(SummaryRow, 9).Value = tickerSymbol
                If yearlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Value = yearlyChange
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(SummaryRow, 10).Value = yearlyChange
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(SummaryRow, 11).Value = FormatPercent(yearlyPercentageChange, 2)
                ws.Cells(SummaryRow, 12).Value = ws.Cells(i, 7).Value
                
                openPrice = ws.Cells(i + 1, 3)
                SummaryRow = SummaryRow + 1
            End If
        Next i
    
        'hard_solutions
        'Your solution will also be able to return the stock with the
        '"Greatest % increase"
        Dim maxRange As Double
        Dim maxRange2 As Double
        lastRow9 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        tickerSymbol = ws.Cells(i, 9).Value
        
        maxRange = ws.Cells(2, 11)
        minRange = ws.Cells(2, 11)
        For j = 2 To lastRow9
            maxRange2 = ws.Cells(j, 11).Value
            If maxRange2 > maxRange Then
                maxRange = ws.Cells(j, 11).Value
                maxTickerSymbol = ws.Cells(j, 9).Value
                
            End If
      
            minRange2 = ws.Cells(j, 11).Value
            If minRange2 < minRange Then
                minRange = ws.Cells(j, 11).Value
                minTickerSymbol = ws.Cells(j, 9).Value
                
            End If
    
        Next j
        
        valueRange = ws.Cells(2, 12)
        For k = 2 To lastRow9
            valueRange2 = ws.Cells(k, 12).Value
            If valueRange2 > valueRange Then
                valueRange = ws.Cells(k, 12).Value
                valueTickerSymbol = ws.Cells(k, 9).Value
                
            End If
    
        Next k
        
        ws.Cells(2, 16).Value = maxTickerSymbol
        ws.Cells(2, 17).Value = FormatPercent(maxRange, 2)
        ws.Cells(3, 16).Value = minTickerSymbol
        ws.Cells(3, 17).Value = FormatPercent(minRange, 2)
        ws.Cells(4, 16).Value = valueTickerSymbol
        ws.Cells(4, 17).Value = valueRange
        
    Next ws

End Sub

