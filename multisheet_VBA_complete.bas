Attribute VB_Name = "multisheet_VBA_complete"
Sub stocks_complete()

Dim ws As Worksheet

    For Each ws In Worksheets

        Dim counter As Integer
        Dim last_row As Long

        last_row = ws.Cells(Rows.count, "A").End(xlUp).row
        counter = 0
        volume = 0
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
    
    
            For i = 2 To last_row
        
                volume = volume + ws.Cells(i, 7).Value
                'tickers
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ws.Cells(2 + counter, 9).Value = ws.Cells(i, 1).Value
                    'totaling volume and reset
                    ws.Cells(2 + counter, 12).Value = volume
                    volume = 0
                    'next ticker
                    counter = counter + 1
                End If
        
                'yearly change and percent change
                If Right(ws.Cells(i, 2).Value, 4) = "0102" Then
                    opn = ws.Cells(i, 3).Value
                ElseIf Right(ws.Cells(i, 2).Value, 4) = "1231" Then
                    cls = ws.Cells(i, 6).Value
                    yr_change = (cls - opn)
                    ws.Cells(1 + counter, 10).Value = yr_change
                    prct_change = yr_change / opn
                    ws.Cells(1 + counter, 11).Value = prct_change
                    If ws.Cells(1 + counter, 10).Value > 0 Then
                        ws.Cells(1 + counter, 10).Interior.ColorIndex = 4
                    ElseIf ws.Cells(1 + counter, 10).Value < 0 Then
                        ws.Cells(1 + counter, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(1 + counter, 10).Interior.ColorIndex = 5
                    End If
                End If
            Next i
    
        max_change = 0
        min_change = 0
        max_vol = 0

            'Greatests
            For i = 2 To ws.Cells(Rows.count, "K").End(xlUp).row
        
                If ws.Cells(i, 11).Value > max_change Then
                    max_change = ws.Cells(i, 11).Value
                    max_change_tic = ws.Cells(i, 9).Value
                ElseIf ws.Cells(i, 11).Value < min_change Then
                    min_change = ws.Cells(i, 11).Value
                    min_change_tic = ws.Cells(i, 9).Value
                End If
        
                If ws.Cells(i, 12).Value > max_vol Then
                    max_vol = ws.Cells(i, 12).Value
                    max_vol_tic = ws.Cells(i, 9).Value
                End If
        
            Next i

        ws.Range("P2") = max_change_tic
        ws.Range("Q2") = max_change
        ws.Range("P3") = min_change_tic
        ws.Range("Q3") = min_change
        ws.Range("P4") = max_vol_tic
        ws.Range("Q4") = max_vol


        'formatting
        ws.Range("K2:K" & ws.Cells(Rows.count, "K").End(xlUp).row).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("Q4").NumberFormat = "0.00E+00"
        ws.Columns("A:Q").EntireColumn.AutoFit
    
    Next ws

End Sub
