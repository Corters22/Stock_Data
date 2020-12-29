Attribute VB_Name = "Module1"
Sub stocks()
    Dim ticker As String
    Dim yearly_diff As Double
    Dim percent_change As Double
    Dim total As Long
    Dim current_ticker As String
    Dim row_number As Integer
        
    Range("I1") = "Ticker"
    Range("J1") = "Yearly change"
    Range("K1") = "Percent Change"
    Range("l1") = "Total Volume"
    last_row = Cells(Rows.Count, "A").End(xlUp).Row
    current_ticker = Cells(2, 1).Value
    row_number = 2
    
    For i = 2 To last_row
        If Cells(i, 1).Value <> current_ticker Then
        current_ticker = Cells(row_number, 9).Value
        row_number = row_number + 1
        current_ticker = Cells(i, 1).Value
    End If
    
    Next i
    
    Cells(row_number, 9).Value = current_ticker
    
End Sub
