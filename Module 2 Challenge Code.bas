Attribute VB_Name = "Module1"
Sub Stock_Outputs()
    For Each ws In Worksheets
        ws.Range("I1:L1").EntireColumn.Insert
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
        ws.Columns("I:L").AutoFit
        end_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        out_range = 2
        tot_stock_vol = 0
    For r = 2 To end_row
        open_price = ws.Cells(r, "C").Value
        close_price = ws.Cells(r, "F").Value
        If (r = 2 Or ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1).Value) Then
            ticker_open = open_price
        End If
        If (ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value) Then
           yearly_delta = close_price - ticker_open
           percent_change = yearly_delta / ticker_open
           tot_stock_vol = tot_stock_vol + ws.Cells(r, "G").Value
           ws.Cells(out_range, "J") = yearly_delta
           ws.Cells(out_range, "J").NumberFormat = "0.00"
                If (yearly_delta < 0) Then
                    ws.Cells(out_range, "J").Interior.ColorIndex = 3
                Else
                    ws.Cells(out_range, "J").Interior.ColorIndex = 4
                End If
        ws.Cells(out_range, "K") = percent_change
        ws.Cells(out_range, "K").Style = "Percent"
        ws.Cells(out_range, "K").NumberFormat = "0.00%"
        ws.Cells(out_range, "I") = ws.Cells(r, 1).Value
        ws.Cells(out_range, "L") = tot_stock_vol
        out_range = out_range + 1
        tot_stock_vol = 0
        ticker_open = 0
        Else
           tot_stock_vol = tot_stock_vol + ws.Cells(r, "G").Value
        End If
Next r
Next ws
End Sub

Sub MaxMinVals()
For Each ws In Worksheets
    ws.Cells(1, "O").Value = "Ticker"
    ws.Cells(1, "P").Value = "Value"
    ws.Cells(2, "N").Value = "Greatest % Increase"
    ws.Cells(3, "N").Value = "Greatest % Decrease"
    ws.Cells(4, "N").Value = "Greatest Total Volume"
    ws.Columns("O:N").AutoFit
    max_p = WorksheetFunction.Max(ws.Range("K:K"))
    min_p = WorksheetFunction.Min(ws.Range("K:K"))
    max_vol = WorksheetFunction.Max(ws.Range("L:L"))
    ws.Cells(2, "p").Value = max_p
    ws.Cells(3, "p").Value = min_p
    ws.Cells(4, "p").Value = max_vol
    ws.Range("P2:P3").Style = "Percent"
    ws.Range("P2:P3").NumberFormat = "0.00%"
    ws.Cells(2, "O").Value = WorksheetFunction.XLookup(ws.Cells(2, "P"), ws.Range("K:K"), ws.Range("I:I"))
    ws.Cells(3, "O").Value = WorksheetFunction.XLookup(ws.Cells(3, "P"), ws.Range("K:K"), ws.Range("I:I"))
    ws.Cells(4, "O").Value = WorksheetFunction.XLookup(ws.Cells(4, "P"), ws.Range("L:L"), ws.Range("I:I"))
    Next ws
End Sub
 
