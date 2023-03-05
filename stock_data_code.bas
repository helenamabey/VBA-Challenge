Attribute VB_Name = "Module1"
Sub stock_data()

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Columns("O:O").ColumnWidth = 17.17
    ws.Columns("L:L").ColumnWidth = 15
    ws.Columns("I:I").ColumnWidth = 12
    ws.Columns("J:J").ColumnWidth = 12
    
    Dim LastRow As Variant
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim i As Variant
    Dim open_price As Double
    Dim close_price As Double
    Dim ticker_row As Variant
    ticker_row = 1
    For i = 2 To LastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            open_price = ws.Cells(i, 3).Value
            ticker_row = ticker_row + 1
            ws.Cells(ticker_row, 9).Value = ws.Cells(i, 1).Value
        End If
        ws.Cells(ticker_row, 12).Value = ws.Cells(ticker_row, 12).Value + ws.Cells(i, 7).Value
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            close_price = ws.Cells(i, 6).Value
            ws.Cells(ticker_row, 10).Value = close_price - open_price
            ws.Cells(ticker_row, 11).Value = Format(ws.Cells(ticker_row, 10).Value / open_price, "0.00%")
            If ws.Cells(ticker_row, 10).Value >= 0 Then
                ws.Cells(ticker_row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(ticker_row, 10).Interior.ColorIndex = 3
            End If
        End If
    Next i
    
    Dim gpi_ticker As String
    Dim gpd_ticker As String
    Dim gtv_ticker As String
    Dim gpi As Double
    Dim gpd As Double
    Dim gtv As Variant
    
    gpi = 0
    gpd = 0
    gtv = 0
    LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To LastRow
        If ws.Cells(i, 11).Value > gpi Then
            gpi = ws.Cells(i, 11).Value
            gpi_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 11).Value < gpd Then
            gpd = ws.Cells(i, 11).Value
            gpd_ticker = ws.Cells(i, 9).Value
        End If
        If ws.Cells(i, 12).Value > gtv Then
            gtv = ws.Cells(i, 12).Value
            gtv_ticker = ws.Cells(i, 9).Value
        End If
    Next i
    
    ws.Cells(2, 16).Value = gpi_ticker
    ws.Cells(2, 17).Value = Format(gpi, "0.00%")
    ws.Cells(3, 16).Value = gpd_ticker
    ws.Cells(3, 17).Value = Format(gpd, "0.00%")
    ws.Cells(4, 16).Value = gtv_ticker
    ws.Cells(4, 17).Value = gtv
Next ws

End Sub
    
