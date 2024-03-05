Attribute VB_Name = "Module1"
Sub module2challenge()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, tickerStart As Long
    Dim Ticker As String, nextTicker As String
    Dim openingPrice As Double, closingPrice As Double
    Dim yearlyChange As Double, percentageChange As Double, totalVolume As Double
    Dim outputRow As Long
    

    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Initialize variables for new sheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        tickerStart = 2
        totalVolume = 0
        outputRow = 2 ' Start output on the second row
        
        ' Sets column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
    
        ' Loop through all rows in the current sheet
        For i = 2 To lastRow
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            Ticker = ws.Cells(i, 1).Value
            nextTicker = ws.Cells(i + 1, 1).Value
            
            If Ticker <> nextTicker Then
                openingPrice = ws.Cells(tickerStart, 3).Value
                closingPrice = ws.Cells(i, 6).Value
                yearlyChange = closingPrice - openingPrice
                percentageChange = ((closingPrice - openingPrice) / openingPrice)
    
                
                ' Output the data & conditional formatting
                With ws
                    
                    .Cells(outputRow, 9).Value = Ticker
                    .Cells(outputRow, 10).Value = yearlyChange
                    
                    ' Sets color of yearlyChange cells
                    If yearlyChange > 0 Then
                        .Cells(outputRow, 10).Interior.ColorIndex = 4
                    Else
                        .Cells(outputRow, 10).Interior.ColorIndex = 3
                    End If
                    
                    .Cells(outputRow, 11).Value = percentageChange
                    
                    ' Converts percentageChange to percentage format
                    .Cells(outputRow, 11).NumberFormat = "0.00%"
                    
                    .Cells(outputRow, 12).Value = totalVolume
                    
                End With
                
                ' Increment the output row for the next set of results
                outputRow = outputRow + 1
                
                ' Reset variables for the next ticker
                tickerStart = i + 1
                totalVolume = 0
            End If
        Next i
    
    Next ws
    
End Sub





