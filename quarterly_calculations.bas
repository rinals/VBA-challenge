Attribute VB_Name = "Module2"
Sub quarterly_change()

    ' Initialize variables
    Dim currentRow As Long
    Dim tickerSymbol As Variant
    Dim prevTickerSymbol As Variant
    Dim tickerCounter As Long
    Dim openPrice As Double
    Dim currentOpenPrice As Double
    Dim currentClosePrice As Double
    Dim prevClosePrice As Double
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim greatestPercentageIncrease As Double
    Dim greatestPercentageDecrease As Double
    Dim totalStockVolume As LongLong
    Dim greatestTotalVolume As LongLong
    Dim currentStockVolume As LongLong
    Dim ws As Worksheet
    
    
    For Each ws In ThisWorkbook.Worksheets
        ' Start at the second row
        currentRow = 2
        tickerCounter = currentRow
        prevTickerSymbol = ws.Cells(currentRow, 1).Value
        openPrice = ws.Cells(currentRow, 3).Value
        prevClosePrice = ws.Cells(currentRow, 6).Value
        currentStockVolume = ws.Cells(currentRow, 7).Value
        totalStockVolume = currentStockVolume
        greatestTotalVolume = 0
        greatestPercentageIncrease = 0
        greatestPercentageDecrease = 0
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    
        ' Loop through rows until an empty cell is found in column A
        Do While Not IsEmpty(ws.Cells(currentRow, 1).Value)
            ' Retrieve the value from column A of the current row
            tickerSymbol = ws.Cells(currentRow, 1).Value
            
            currentOpenPrice = ws.Cells(currentRow, 3).Value
            currentClosePrice = ws.Cells(currentRow, 6).Value
            currentStockVolume = ws.Cells(currentRow, 7).Value
                
            
            If tickerSymbol <> prevTickerSymbol Then
                ws.Cells(tickerCounter, 9).Value = prevTickerSymbol
                
                
                quarterlyChange = prevClosePrice - openPrice
                ws.Cells(tickerCounter, 10).Value = quarterlyChange
                If quarterlyChange < 0 Then
                    With ws.Cells(tickerCounter, 10).Interior
                        .Color = RGB(255, 0, 0)
                    End With
                End If
                If quarterlyChange > 0 Then
                    With ws.Cells(tickerCounter, 10).Interior
                        .Color = RGB(0, 255, 0)
                    End With
                End If
                
                percentageChange = quarterlyChange / openPrice
                ws.Cells(tickerCounter, 11).Value = percentageChange
                ws.Cells(tickerCounter, 11).NumberFormat = "0.00%"
                
                If percentageChange > greatestPercentageIncrease Then
                    greatestPercentageIncrease = percentageChange
                    ws.Cells(2, 17).Value = greatestPercentageIncrease
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                    ws.Cells(2, 16).Value = prevTickerSymbol
                    
                End If
                    
                If percentageChange < greatestPercentageDecrease Then
                    greatestPercentageDecrease = percentageChange
                    ws.Cells(3, 17).Value = greatestPercentageDecrease
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                    ws.Cells(3, 16).Value = prevTickerSymbol
                
                End If
            
                openPrice = currentOpenPrice
            
                ws.Cells(tickerCounter, 12).Value = totalStockVolume
                If totalStockVolume > greatestStockVolume Then
                    greatestStockVolume = totalStockVolume
                    ws.Cells(4, 17).Value = greatestStockVolume
                    ws.Cells(4, 16).Value = prevTickerSymbol
                    
                End If
            
                prevTickerSymbol = tickerSymbol
                
                totalStockVolume = 0
            
                tickerCounter = tickerCounter + 1
            End If
            
            totalStockVolume = totalStockVolume + currentStockVolume
            prevClosePrice = currentClosePrice

            ' Move to the next row
            currentRow = currentRow + 1
        Loop
    Next ws

    
End Sub
