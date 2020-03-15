Sub TickerScript():
    For Each ws In Worksheets
        
        Dim TickerName As String
        Dim VolTotal As Double
        Dim SumTable As Integer
        Dim BeginYear As Double
        Dim EndYear As Double
        
        VolTotal = 0
        SumTable = 2
        BeginYear = Cells(2, 6).Value
        EndYear = 0
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Columns("I:L").AutoFit
        ws.Range("K1:K" & lastrow).NumberFormat = "0.00%"
        
        For i = 2 To lastrow
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerName = ws.Cells(i, 1).Value
                EndYear = ws.Cells(i, 6).Value
                VolTotal = VolTotal + ws.Cells(i, 7).Value
                ws.Range("I" & SumTable).Value = TickerName
                If (EndYear <> 0 And BeginYear <> 0 And VolTotal <> 0) Then
                    ws.Range("J" & SumTable).Value = (EndYear) - (BeginYear)
                    ws.Range("K" & SumTable).Value = ((EndYear) - (BeginYear)) / (BeginYear)
                    ws.Range("L" & SumTable).Value = VolTotal
                Else
                    ws.Range("J" & SumTable).Value = 0
                    ws.Range("K" & SumTable).Value = 0
                    ws.Range("L" & SumTable).Value = 0
                End If
                SumTable = SumTable + 1
                BeginYear = ws.Cells(i + 1, 6).Value
                EndYear = 0
                VolTotal = 0
            Else
                VolTotal = VolTotal + ws.Cells(i, 7).Value
            End If
        Next i
        
        For j = 2 To lastrow
            If ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            End If
        Next j
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        Dim GreatInc As Double
        Dim GreatDec As Double
        Dim GreatVol As Double
        
        GreatInc = 0
        GreatDec = 0
        GreatVol = 0
        
        For i = 2 To lastrow
            If ws.Cells(i, 11).Value > GreatInc Then
                GreatInc = ws.Cells(i, 11).Value
                ws.Range("P2").Value = ws.Cells(i, 9).Value
                ws.Range("Q2").Value = GreatInc
            End If
            
            If ws.Cells(i, 11).Value < GreatDec Then
                GreatDec = ws.Cells(i, 11).Value
                ws.Range("P3").Value = ws.Cells(i, 9).Value
                ws.Range("Q3").Value = GreatDec
            End If
            
            If ws.Cells(i, 12).Value > GreatVol Then
                GreatVol = ws.Cells(i, 12).Value
                ws.Range("P4").Value = ws.Cells(i, 9).Value
                ws.Range("Q4").Value = GreatVol
            End If
        Next i
        
        ws.Columns("O:Q").AutoFit
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
    Next ws
End Sub



