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
        
    'To run code in each worksheet
    For Each ws In Worksheets
        With ws
        ws.Activate
        
            'Headers for new data
            Range("I1") = "Ticker"
            Range("J1") = "Yearly change"
            Range("K1") = "Percent Change"
            Range("l1") = "Total Volume"
            Range("n2") = "Greatest % Increase"
            Range("n3") = "Greatest % Decrease"
            Range("n4") = "Greatest Total Volume"
            Range("o1") = "Ticker"
            Range("p1") = "Value"
    
            'defining variables for the loop
            last_row = Cells(Rows.Count, "A").End(xlUp).Row
            current_ticker = Cells(2, 1).Value
            row_number = 2
            open_amt = Cells(2, 3).Value
            total = Cells(2, 7).Value
    
                'Loop to pull tickers, difference between open amount and closed amount, and total volume
                For i = 2 To last_row
                    If Cells(i, 1).Value <> current_ticker Then
                        Cells(row_number, 9).Value = current_ticker
                        'Calculating the difference in open and closed amount
                        Cells(row_number, 10).Value = Cells(i - 1, 6).Value - open_amt
                            'Formatting cells based on calculated amount
                            If Cells(row_number, 10).Value >= 0 Then
                                Cells(row_number, 10).interior.color = vbGreen
                            Else: Cells(row_number, 10).interior.color = vbred
                            End If
                        'Calculating the percentage change
                        If open_amt <> 0 Then
                            Cells(row_number, 11).Value = Cells(row_number, 10).Value / open_amt
                            Else: Cells(row_number, 11).Value = "NA"
                        End If
                                                
                        'Printing total volume to new cell after change in tickers
                        Cells(row_number, 12).Value = total
                        row_number = row_number + 1
                        'Resetting values before next i
                        current_ticker = Cells(i, 1).Value
                        open_amt = Cells(i, 3).Value
                        total = Cells(i, 7).Value
                    End If

                    'Calculating total volume as the loop goes through                
                    If Cells(i, 1).Value = current_ticker Then
                        total = total + Cells(i, 7).Value
                    End If
            
                Next i
    
            'Pushing out final values of last row   
            Cells(row_number, 9).Value = current_ticker
            Cells(row_number, 10).Value = Cells(i - 1, 6).Value - open_amt
                If Cells(row_number, 10).Value >= 0 Then
                   Cells(row_number, 10).interior.color = vbGreen
                Else: Cells(row_number, 10).interior.color = vbred
                End If
            Cells(row_number, 11).Value = Cells(row_number, 10).Value / open_amt
            Cells(row_number, 12).Value = total
            
            'Calculating Max and Min
            'Defining variables based on row number of first value
            max_increase = 2
            max_decrease = 2
            max_volume = 2
    
    
                'Loop to calculate max and min
                For i = 3 To last_row
                    'Calculates Max of percentage change
                    If IsNumeric(Cells(i, 11).Value) = True Then
                        If Cells(i, 11).Value > Cells(max_increase, 11).Value Then max_increase = i
                    'Calculates Min of percentage change
                        If Cells(i, 11).Value < Cells(max_decrease, 11).Value Then max_decrease = i
                    End If
            
                    'Calculates Max total volume
                    If Cells(i, 12).Value > Cells(max_volume, 12).Value Then
                        max_volume = i
                    End If
                Next i
            
            'Pushes calculations and tickers to empty cells
            'Max %
            Cells(2, 16).Value = Cells(max_increase, 11).Value
            'ticker symbol
            Cells(2, 15).Value = Cells(max_increase, 9).Value
            'Min %
            Cells(3, 16).Value = Cells(max_decrease, 11).Value
            'ticker symbol
            Cells(3, 15).Value = Cells(max_decrease, 9).Value
            'Max volume
            Cells(4, 16).Value = Cells(max_volume, 12).Value
            'Ticker symbol
            Cells(4, 15).Value = Cells(max_volume, 9).Value
        
                          
            'Formatting spreadsheet           
            Range("K:K").NumberFormat = "0.00%"
            Range("p2:p3").NumberFormat = "0.00%"
        
            ActiveSheet.Range("a:p").Columns.AutoFit
        End With
    Next ws
        
End Sub
