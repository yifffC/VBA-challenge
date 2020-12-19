Sub alphabetical_testing()
Dim ticker As String
Dim yearly_change As Double
Dim percentage_change As Double
Dim volume As LongLong
Dim summary_table_row As Integer
Dim close_price As Double
Dim open_price As Double

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percentage Change"
Range("L1") = "Total Volume"

Range("N2") = "Greatest % Increase"
Range("N3") = "Greatest % Decrease"
Range("N4") = "Max Volume"
Range("O1") = "Ticker"
Range("P1") = "Value"

yearly_change = 0
percentage_change = 0
volume = 0
summary_table_row = 2
open_price = Cells(2, 3).Value
close_price = 0

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        volume = volume + Cells(i, 7).Value
        close_price = Cells(i, 6).Value
        yearly_change = close_price - open_price
        percentage_change = yearly_change / open_price
       
        If volume <> 0 Then
       
            Range("I" & summary_table_row).Value = ticker
            Range("J" & summary_table_row).Value = yearly_change
            Range("K" & summary_table_row).Value = percentage_change
            Range("K" & summary_table_row).NumberFormat = "0.00%"
            Range("L" & summary_table_row).Value = volume
            
            If yearly_change > 0 Then
                Range("J" & summary_table_row).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Range("J" & summary_table_row).Interior.ColorIndex = 3
            Else
                Range("J" & summary_table_row).Interior.ColorIndex = 0
            End If
            
            summary_table_row = summary_table_row + 1
            yearly_change = 0
            percentage_change = 0
            volume = 0
            If Cells(i + 1, 3).Value <> 0 Then
                open_price = Cells(i + 1, 3).Value
            Else
                For J = (i + 1) To lastrow
                    If Cells(J, 3).Value <> 0 Then
                        open_price = Cells(J, 3).Value
                    
                    End If
                Next J
            End If
        End If
    Else
        volume = volume + Cells(i, 7).Value
    End If
Next i

Dim increase As Double
Dim decrease As Double
Dim maxvol As LongLong
increase = Application.Max(Range("K:K"))
decrease = Application.Min(Range("K:K"))
maxvol = Application.Max(Range("L:L"))

For Row = 2 To lastrow
    If Cells(Row, 11).Value = increase Then
        Range("O2").Value = Cells(Row, 9)
        Range("P2").Value = increase
    ElseIf Cells(Row, 11).Value = decrease Then
        Range("O3").Value = Cells(Row, 9)
        Range("P3").Value = decrease
    ElseIf Cells(Row, 12).Value = maxvol Then
        Range("O4").Value = Cells(Row, 9)
        Range("P4").Value = maxvol
    End If
Next Row
Range("P2").NumberFormat = "0.00%"
Range("P3").NumberFormat = "0.00%"
End Sub