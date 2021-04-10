Attribute VB_Name = "Module1"
Option Explicit

Sub stocks():

Dim i, percent_change, max_perc, min_perc, max_vol, last_row, row_lines, percent_data, open_price, close_price, total_vol, row_counter As Integer

For i = 1 To Worksheets.Count
Dim ws As Worksheet
Set ws = Worksheets(i)

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
Dim max_tick, max_perc_tick, min_perc_tick As String
    total_vol = 0
    max_vol = 0
    max_perc = 0
    row_counter = 2
    open_price = ws.Cells(2, 3).Value

    For row_lines = 2 To ws.Range("A1").End(xlDown).Row
        total_vol = total_vol + ws.Cells(row_lines, 7).Value
        If ws.Cells(row_lines + 1, 1) <> ws.Cells(row_lines, 1) Then
            ws.Cells(row_counter, 9).Value = ws.Cells(row_lines, 1).Value
            ws.Cells(row_counter, 12).Value = total_vol
            If total_vol > max_vol Then
                max_vol = total_vol
                max_tick = ws.Cells(row_lines, 1).Value
            End If
            total_vol = 0
            close_price = ws.Cells(row_lines, 6)
            ws.Cells(row_counter, 10).Value = close_price - open_price
            If open_price = 0 Then
                ws.Cells(row_counter, 11).Value = 0
            Else:
                percent_change = (close_price - open_price) / open_price
                ws.Cells(row_counter, 11).Value = percent_change
                If percent_change > max_perc Then
                    max_perc = percent_change
                    max_perc_tick = ws.Cells(row_lines, 1).Value
                End If
            End If
            open_price = ws.Cells(row_lines + 1, 3)
            row_counter = row_counter + 1
        End If
    Next row_lines
    min_perc = ws.Cells(2, 11).Value
    min_perc_tick = ws.Cells(2, 9).Value
    For row_lines = 3 To ws.Range("K2").End(xlDown).Row
        If ws.Cells(row_lines, 11) < min_perc Then
            min_perc = ws.Cells(row_lines, 11)
            min_perc_tick = ws.Cells(row_lines, 9)
        End If
    Next row_lines
    last_row = ws.Range("K1").End(xlDown).Row
    ws.Range("K2:K" & last_row).NumberFormat = "0.00%"
    ws.Range("P2:P3").NumberFormat = "0.00%"
    For percent_data = 2 To ws.Range("J1").End(xlDown).Row
        If ws.Cells(percent_data, 10).Value < 0 Then
            ws.Cells(percent_data, 10).Interior.Color = RGB(255, 0, 0)
        Else:
            ws.Cells(percent_data, 10).Interior.Color = RGB(0, 255, 0)
        End If
    Next percent_data
    ws.Cells(2, 15).Value = max_perc_tick
    ws.Cells(3, 15).Value = min_perc_tick
    ws.Cells(4, 15).Value = max_tick
    ws.Cells(2, 16).Value = max_perc
    ws.Cells(3, 16).Value = min_perc
    ws.Cells(4, 16).Value = max_vol
    
Next i

End Sub
