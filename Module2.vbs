Sub Stocks()
    For Each ws In Worksheets
        Dim ticker As String
        Dim yearly_change As Double
        opening_price = Cells(2, 3).Value
        Dim percent_change As Double
        Dim total_vol As Double
        Dim summary_table_row As Double
        summary_table_row = 2
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percentage Change"
        ws.Cells(1, 12) = "Total Volume"
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(4, 14) = "Greatest Total Volume"
        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 16) = "Value"
        '   get Last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        '   Loop through all ticker symbols
            For I = 2 To lastrow
                '   check if we are still on same ticker symbol
                If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                    '   Set Ticker Symbol
                    ticker = ws.Cells(I, 1).Value
                    '   Calc yearly change
                    closing_price = ws.Cells(I, 6).Value
                    yearly_change = closing_price - opening_price
                    '   Calc Percent Change
                    percentage_change = (yearly_change / opening_price)
                    '   calc total vol
                    total_vol = total_vol + ws.Cells(I, 7).Value
                    '   print ticker
                    ws.Range("I" & summary_table_row).Value = ticker
                    '   print yearly change
                    ws.Range("J" & summary_table_row).Value = yearly_change
                    '   print percentage change
                    ws.Range("K" & summary_table_row).Value = FormatPercent(percentage_change)
                    '   print total stock vol
                    ws.Range("L" & summary_table_row).Value = total_vol
                    '   add one to summary table row
                    summary_table_row = summary_table_row + 1
                    '   reset total vol
                    total_vol = 0
                    opening_price = ws.Cells(I + 1, 3).Value
                Else
                    total_vol = total_vol + ws.Cells(I, 7).Value
                
            End If
                If ws.Cells(I, 11).Value > 0 Then
                ws.Cells(I, 11).Interior.ColorIndex = 4
                ElseIf ws.Cells(I, 11).Value < 0 Then
                ws.Cells(I, 11).Interior.ColorIndex = 3
            End If
            Next I
            '   tickers for min and max -could not identify
            '   min and max of percentage
                Cells(2, 16).Value = FormatPercent(WorksheetFunction.Max(ws.Range("K:K")))
                Cells(3, 16).Value = WorksheetFunction.Min(ws.Range("K:K"))
                Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("L:L"))
    Next ws
End Sub




