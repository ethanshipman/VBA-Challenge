Attribute VB_Name = "Module1"
Sub StockData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim tickerSymbol As String
    Dim startPrice As Double
    Dim endPrice As Double
    Dim startDate As Variant
    Dim endDate As Variant
    Dim quarter As String
    Dim resultRow As Long
    Dim quarterlyChange As Double
    Dim percentageChange As Double
    Dim totalVolume As Double
    

    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim stockGreatestIncrease As String
    Dim stockGreatestDecrease As String
    Dim stockGreatestVolume As String
    
 
    greatestIncrease = -999999
    greatestDecrease = 999999
    greatestVolume = 0

  
    For Each ws In ThisWorkbook.Worksheets

        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row


        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Quarter"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Volume"

        resultRow = 2

       
        For i = 2 To lastRow
            tickerSymbol = ws.Cells(i, 1).Value
            startDate = ws.Cells(i, 2).Value
            quarter = GetQuarter(startDate)
            
            startPrice = ws.Cells(i, 3).Value
            endPrice = 0
            totalVolume = 0

           
            Do While ws.Cells(i, 1).Value = tickerSymbol And GetQuarter(ws.Cells(i, 2).Value) = quarter And i <= lastRow
                endPrice = ws.Cells(i, 4).Value
                totalVolume = totalVolume + ws.Cells(i, 5).Value
                i = i + 1
            Loop

            
            If startPrice <> 0 Then
                quarterlyChange = endPrice - startPrice
                percentageChange = Round((quarterlyChange / startPrice) * 100, 1)

                
                ws.Cells(resultRow, 8).Value = tickerSymbol
                ws.Cells(resultRow, 9).Value = quarter
                ws.Cells(resultRow, 10).Value = quarterlyChange
                ws.Cells(resultRow, 11).Value = percentageChange & "%"
                ws.Cells(resultRow, 12).Value = totalVolume

                
                If quarterlyChange > 0 Then
                    ws.Cells(resultRow, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(resultRow, 10).Interior.ColorIndex = 3
                End If

              
                If percentageChange > greatestIncrease Then
                    greatestIncrease = percentageChange
                    stockGreatestIncrease = tickerSymbol
                End If

           
                If percentageChange < greatestDecrease Then
                    greatestDecrease = percentageChange
                    stockGreatestDecrease = tickerSymbol
                End If

        
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    stockGreatestVolume = tickerSymbol
                End If

                resultRow = resultRow + 1
            End If

            i = i - 1
        Next i

    
        ws.Cells(2, 15).Value = "Greatest % Increase:"
        ws.Cells(2, 16).Value = stockGreatestIncrease
        ws.Cells(2, 17).Value = greatestIncrease & "%"

        ws.Cells(3, 15).Value = "Greatest % Decrease:"
        ws.Cells(3, 16).Value = stockGreatestDecrease
        ws.Cells(3, 17).Value = greatestDecrease & "%"

        ws.Cells(4, 15).Value = "Greatest Total Volume:"
        ws.Cells(4, 16).Value = stockGreatestVolume
        ws.Cells(4, 17).Value = greatestVolume

    Next ws

End Sub


Function GetQuarter(ByVal dt As Date) As String
    Dim monthNum As Integer
    monthNum = month(dt)

    Select Case monthNum
        Case 1 To 3
            GetQuarter = "Q1"
        Case 4 To 6
            GetQuarter = "Q2"
        Case 7 To 9
            GetQuarter = "Q3"
        Case 10 To 12
            GetQuarter = "Q4"
        Case Else
            GetQuarter = "Invalid Quarter"
    End Select
End Function
