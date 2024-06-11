Attribute VB_Name = "Module1"
Sub QuarterlyStockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim startRow As Long
    Dim endRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim i As Long
    Dim outputRow As Long
    Dim quarterlyChange As Double
    Dim percentChange As Double
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String

    greatestIncrease = 0 ' start with 0 for greatest increase
    greatestDecrease = 0 ' start with 0 for greatest decrease
    greatestVolume = 0   ' start with 0 for greatest volume

    For Each ws In ThisWorkbook.Worksheets
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        outputRow = 2
        startRow = 2

        For i = 2 To lastRow
            ' Doesf the ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                endRow = i

                openPrice = ws.Cells(startRow, 3).Value

    
                closePrice = ws.Cells(endRow, 6).Value

                quarterlyChange = closePrice - openPrice

                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If

                totalVolume = Application.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(endRow, 7)))

                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume

                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If

                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If

                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If

                outputRow = outputRow + 1

                startRow = i + 1
            End If
        Next i

        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("L").NumberFormat = "#,##0"


        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        ws.Cells(2, 16).Value = greatestIncreaseTicker
        ws.Cells(2, 17).Value = greatestIncrease

        ws.Cells(3, 16).Value = greatestDecreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease

        ws.Cells(4, 16).Value = greatestVolumeTicker
        ws.Cells(4, 17).Value = greatestVolume

    Next ws
End Sub

