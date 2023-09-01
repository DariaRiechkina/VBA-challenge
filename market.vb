Sub market():
    Dim total_volume As Double
    Dim ticker As String
    Dim percentage As Double
    Dim price_change As Double
    Dim Row As Long
    Dim TargetRow As Long
    Dim FirstTickerRow As Long
    Dim Last_Row As Long
    Dim greatest_total As Double
    Dim greatest_increse As Double
    Dim greatest_decrease As Double
    
    For Each ws In Worksheets
        TargetRow = 2
        Row = 2
        FirstTickerRow = 2
        total_volume = 0
        
        Last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        Do While Row <= Last_Row
            
            ticker = ws.Cells(Row, 1).Value
            total_volume = total_volume + ws.Cells(Row, 7).Value
            
            If ws.Cells(Row + 1, 1).Value <> ticker Then
                price_change = ws.Cells(Row, 6).Value - ws.Cells(FirstTickerRow, 3).Value
                percentage = (price_change / ws.Cells(FirstTickerRow, 3).Value) * 100
                
                ws.Cells(TargetRow, 10).Value = ticker
                ws.Cells(TargetRow, 11).Value = price_change
                ws.Cells(TargetRow, 12).Value = percentage
                ws.Cells(TargetRow, 13).Value = total_volume
                If percentage >= 0 Then
                    ws.Cells(TargetRow, 12).Interior.ColorIndex = 4
                Else
                    ws.Cells(TargetRow, 12).Interior.ColorIndex = 3
                End If
                total_volume = 0
                Row = Row + 1
                TargetRow = TargetRow + 1
                FirstTickerRow = Row
            Else
                Row = Row + 1
            End If
        Loop
        
        Row = 2
        greatest_total = 0
        greatest_increse = 0
        greatest_decrease = 0
        greatest_total = 0
        
        Do While Row <= TargetRow
            If ws.Cells(Row, 12).Value > greatest_increse Then
                greatest_increse = ws.Cells(Row, 12).Value
                ws.Cells(2, 17).Value = ws.Cells(Row, 10).Value
                ws.Cells(2, 18).Value = ws.Cells(Row, 12).Value
            End If
            
            If ws.Cells(Row, 12).Value < greatest_decrease Then
                greatest_decrease = ws.Cells(Row, 12).Value
                ws.Cells(3, 17).Value = ws.Cells(Row, 10).Value
                ws.Cells(3, 18).Value = ws.Cells(Row, 12).Value
            End If
            
            If ws.Cells(Row, 13).Value > greatest_total Then
                greatest_total = ws.Cells(Row, 13).Value
                ws.Cells(4, 17).Value = ws.Cells(Row, 10).Value
                ws.Cells(4, 18).Value = ws.Cells(Row, 13).Value
            End If

            Row = Row + 1
        Loop
    Next ws
End Sub
