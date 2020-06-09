Sub TickerScript():
    'Create a loop to target every worksheet
    For Each ws In Worksheets
        
        'Declare variables
        Dim TickerName As String
        Dim VolTotal As Double
        Dim SumTable As Integer
        Dim BeginYear As Double
        Dim EndYear As Double
        
        'Create a counter for the total volume count; create a counter for the sum table starting at 2
        VolTotal = 0
        SumTable = 2
        
        'Targets the <vol> column for every worksheet
        BeginYear = ws.Cells(2, 6).Value
        EndYear = 0
        
        'Create a variable to get the lastrow
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Create headers for calculations soon to be collected
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Make the cells autofit to their respectable spaces
        ws.Columns("I:L").AutoFit
        
        'Format the numbers to display as percentages
        ws.Range("K1:K" & lastrow).NumberFormat = "0.00%"
        
        'Loop from the 2nd row to the last row
        For i = 2 To lastrow
            
            'Checks if the ticker name matches; if not, then perform the code
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Saves the ticker name
                TickerName = ws.Cells(i, 1).Value
                
                'Saves the end year value
                EndYear = ws.Cells(i, 6).Value
                
                'Updates the volume total
                VolTotal = VolTotal + ws.Cells(i, 7).Value
                
                'Prints the ticker name in the cell, using the SumTable counter
                ws.Range("I" & SumTable).Value = TickerName
                
                'If EndYear, BeginYear and VolTotal doesn't equal 0, then execute
                If (EndYear <> 0 And BeginYear <> 0 And VolTotal <> 0) Then
                    ws.Range("J" & SumTable).Value = (EndYear) - (BeginYear)
                    ws.Range("K" & SumTable).Value = ((EndYear) - (BeginYear)) / (BeginYear)
                    ws.Range("L" & SumTable).Value = VolTotal
                Else
                    ws.Range("J" & SumTable).Value = 0
                    ws.Range("K" & SumTable).Value = 0
                    ws.Range("L" & SumTable).Value = 0
                End If
                
                'Updates the SumTable to move down a row
                SumTable = SumTable + 1
                
                'Updates the BeginYear to move down a row
                BeginYear = ws.Cells(i + 1, 6).Value
                
                'Resets the EndYear and VolTotal to 0
                EndYear = 0
                VolTotal = 0
            Else
                'If no conditions are met, add the <vol> value to the VolTotal
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
        
        'Print the texts in the cells
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        Dim GreatInc As Double
        Dim GreatDec As Double
        Dim GreatVol As Double
        
        'Start the counter for each at 0 to add to it
        GreatInc = 0
        GreatDec = 0
        GreatVol = 0
        
        For i = 2 To lastrow
            'If the value in the cell is higher than variables, update it the highest number to occupy the variables
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
        
        'Auto fits the cells from O to Q; change the format to be percentages
        ws.Columns("O:Q").AutoFit
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
    Next ws
End Sub