Attribute VB_Name = "Module1"
Sub stocks()
    Dim Ticker As String
    Dim yearly_diff As Double
    Dim percent_change As Double
    Dim current_ticker As String
    Dim row_number As Integer
    Dim last_row As Long
    Dim open_amt As Double
    Dim close_amt As Double
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition
    Dim rng As Range
    
        
    Range("I1") = "Ticker"
    Range("J1") = "Yearly change"
    Range("K1") = "Percent Change"
    Range("l1") = "Total Volume"
    last_row = Cells(Rows.Count, "A").End(xlUp).Row
    current_ticker = Cells(2, 1).Value
    row_number = 2
    open_amt = Cells(2, 3).Value
    total = Cells(2, 7).Value
    
    
            For i = 2 To last_row
                If Cells(i, 1).Value <> current_ticker Then
                    Cells(row_number, 9).Value = current_ticker
                    Cells(row_number, 10).Value = Cells(i - 1, 6).Value - open_amt
                    Cells(row_number, 11).Value = Str(Cells(row_number, 10).Value / open_amt * 100) + "%"
                    Cells(row_number, 12).Value = total
                    row_number = row_number + 1
                    current_ticker = Cells(i, 1).Value
                    open_amt = Cells(i, 3).Value
                    total = Cells(i, 7).Value
             End If
                If Cells(i, 1).Value = current_ticker Then
                    total = total + Cells(i, 7).Value
                End If
            
        Next i
    
        Cells(row_number, 9).Value = current_ticker
        Cells(row_number, 10).Value = Cells(i - 1, 6).Value - open_amt
        Cells(row_number, 11).Value = Cells(row_number, 10).Value / open_amt * 100
        Cells(row_number, 12).Value = total
        
        Set rng = Range("j:j")
        rng.FormatConditions.Delete
        Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
        Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
        
        With condition1
        .Interior.Color = vbGreen
        End With
        
        With condition2
        .Interior.Color = vbRed
        End With
        
        Set rng = Range("j1")
        rng.FormatConditions.Delete
        
        
End Sub
