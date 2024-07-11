Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock()

Dim i As Long
Dim LastRow As Long
Dim TotalStockVolume As Double
Dim QuarterlyChange As Double
Dim PercentChange As Double
Dim j As Integer
Dim ws As Worksheet
Dim OpeningValue As Double
Dim ClosingValue As Double
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim GreatestTotalVolume  As Double


For Each ws In Worksheets
        TotalStockVolume = 0
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
       'MsgBox (LastRow)
        j = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(j, 12).Value = TotalStockVolume
                j = j + 1
                TotalStockVolume = 0
            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
        Dim s As Integer
        s = 2
        For i = 2 To LastRow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ClosingValue = ws.Cells(i, 6).Value
                QuarterlyChange = ClosingValue - OpeningValue
                ws.Cells(s, 10).Value = QuarterlyChange
                If QuarterlyChange > 0 Then
                    ws.Cells(s, 10).Interior.ColorIndex = 4
                ElseIf QuarterlyChange < 0 Then
                    ws.Cells(s, 10).Interior.ColorIndex = 3
                End If
                PercentChange = ((ClosingValue - OpeningValue) / OpeningValue)
                ws.Cells(s, 11).Value = PercentChange
                ws.Cells(s, 11).NumberFormat = "0.00%"
                
                s = s + 1
                OpeningValue = 0
                ClosingValue = 0
            
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                OpeningValue = ws.Cells(i, 3).Value
                
            Else
                OpeningValue = OpeningValue
                ClosingValue = ClosingValue
            End If
        Next i
        
GreatestIncrease = 0
GreatestDecrease = 0
GreatestTotalVolume = 0

        For i = 2 To LastRow
            If ws.Cells(i, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 17).Value = GreatestIncrease
                ws.Cells(2, 17).NumberFormat = "0.00%"
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            Else
                GreatestIncrease = GreatestIncrease
            End If
            Next i
            
        For i = 2 To LastRow
                If ws.Cells(i, 11).Value < GreatestDecrease Then
                    GreatestDecrease = ws.Cells(i, 11).Value
                    ws.Cells(3, 17).Value = GreatestDecrease
                    ws.Cells(3, 17).NumberFormat = "0.00%"
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                Else
                    GreatestDecrease = GreatestDecrease
                End If
                Next i
                
        For i = 2 To LastRow
                If ws.Cells(i, 12).Value > GreatestTotalVolume Then
                    GreatestTotalVolume = ws.Cells(i, 12).Value
                    ws.Cells(4, 17).Value = GreatestTotalVolume
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                Else
                    GreatestTotalVolume = GreatestTotalVolume
                End If
                Next i
    Next ws
    

End Sub



