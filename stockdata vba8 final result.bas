Attribute VB_Name = "Stockdata"
Sub stocks()
    Dim current_ticker As String
    Dim row_number As Integer
    Dim last_row As Long
    Dim open_amt As Double
    Dim close_amt As Double
    Dim rng As Range
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim ws As Worksheet
    'Dim max_volume As Long
        
    For Each ws In Worksheets
        With ws
        ws.Activate
        

            Range("I1") = "Ticker"
            Range("J1") = "Yearly change"
            Range("K1") = "Percent Change"
            Range("l1") = "Total Volume"
            Range("n2") = "Greatest % Increase"
            Range("n3") = "Greatest % Decrease"
            Range("n4") = "Greatest Total Volume"
            Range("o1") = "Ticker"
            Range("p1") = "Value"
    
            last_row = Cells(Rows.Count, "A").End(xlUp).Row
            current_ticker = Cells(2, 1).Value
            row_number = 2
            open_amt = Cells(2, 3).Value
            total = Cells(2, 7).Value
    
    
                For i = 2 To last_row
                    If Cells(i, 1).Value <> current_ticker Then
                        Cells(row_number, 9).Value = current_ticker
                        Cells(row_number, 10).Value = Cells(i - 1, 6).Value - open_amt
                            If Cells(row_number, 10).Value >= 0 Then
                                Cells(row_number, 10).interior.color = vbGreen
                            Else: Cells(row_number, 10).interior.color = vbred
                            End If
                        If open_amt <> 0 Then
                            Cells(row_number, 11).Value = Cells(row_number, 10).Value / open_amt
                            Else: Cells(row_number, 11).Value = "NA"
                        End If
                                                
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
                If Cells(row_number, 10).Value >= 0 Then
                   Cells(row_number, 10).interior.color = vbGreen
                Else: Cells(row_number, 10).interior.color = vbred
                End If
            Cells(row_number, 11).Value = Cells(row_number, 10).Value / open_amt
            Cells(row_number, 12).Value = total
        
                          
            Set max_increase = ActiveSheet.Range("k:k").Find(what:=(Application.WorksheetFunction.Max(ActiveSheet.Range("k:k"))))
            Range("p2").Value = max_increase
            Range("o2").Value = Cells(max_increase.Row, 9).Value
        
            Set max_decrease = ActiveSheet.Range("k:k").Find(what:=(Application.WorksheetFunction.Min(ActiveSheet.Range("k:k"))))
            Range("p3").Value = max_decrease
            Range("o3").Value = Cells(max_decrease.Row, 9).Value
        
            Set max_volume = ActiveSheet.Range("l:l").Find(what:=(Application.WorksheetFunction.Max(ActiveSheet.Range("l:l"))))
            Range("p4").Value = max_volume
            Range("o4").Value = Cells(max_volume.Row, 9).Value
            
            Range("K:K").NumberFormat = "0.00%"
            Range("p2:p3").NumberFormat = "0.00%"
        
            ActiveSheet.Range("a:p").Columns.AutoFit
        End With
    Next ws
        
End Sub
