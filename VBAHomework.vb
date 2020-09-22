Sub StockBuilder()

Dim YC As Double
Dim PC As Double
Dim TSV As Long
Dim VolumeTotal As Double
Dim currentticker As String
Dim nextticker As String
Dim currentrow As Long
Dim summary_row As Long
Dim openprice As Double
Dim closeprice As Double

Dim ws As Worksheet

    For Each ws In Worksheets
        ws.Activate
        totalrows = Cells(Rows.Count, "A").End(xlUp).Row
        summary_row = 2
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        For currentrow = 2 To totalrows
            currentticker = Cells(currentrow, 1).Value
            nextticker = Cells(currentrow + 1, 1).Value
            VolumeTotal = 0
            
            If currentticker = nextticker Then
                VolumeTotal = VolumeTotal + Cells(currentrow, 7).Value
                If Cells(currentrow, 3).Value <> 0 Then
                    openprice = Cells(currentrow, 3).Value
                    If openprice = 0 Then
                        openprice = Cells(currentrow + 1, 3).Value
                    End If
                End If
            Else
                VolumeTotal = VolumeTotal + Cells(currentrow, 7).Value
                closeprice = Cells(currentrow, 6).Value
                YC = (closeprice - openprice)
                PC = (PC1 / openprice) * 100
                
                
            'Summary Table
                Cells(summary_row, 9).Value = currentticker
                Cells(summary_row, 10).Value = YC
                Cells(summary_row, 11).NumberFormat = "###.##%"
                Cells(summary_row, 11).Value = PC
                Cells(summary_row, 12).Value = VolumeTotal
                
            'Summary Formatting
                If Cells(summary_row, 10).Value > 0 Then
                    Cells(summary_row, 10).Interior.ColorIndex = 4
                ElseIf Cells(summary_row, 10).Value < 0 Then
                    Cells(summary_row, 10).Interior.ColorIndex = 3
                End If
                
                summary_row = summary_row + 1
                VolumeTotal = 0
                PC = 0
                YC = 0
                TSV = 0
                closeprice = 0
            End If
        Next currentrow
        MsgBox ws.Name
    Next
End Sub

