Sub StockSummary()
 
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    'add column headers for summary table
    ws.cells(1,9).value = "Ticker"
    ws.cells(1,10).value = "Yearly Change"
    ws.cells(1,11).value = "Percent Change"
    ws.cells(1,12).value = "Total Stock Volume"
    'initialize variables
    rowEnd = ws.Cells(2, 1).End(xlDown).Row
    volume = 0 
    summary_row = 2
    'set initial opening price
    open_price = ws.cells(2,3).value 

    For row = 2 to rowEnd 
        current_stock = ws.Cells(row, 1).Value
        next_stock = ws.Cells(row + 1, 1).Value
        current_volume = ws.Cells(row, 7).Value

        If current_stock = next_stock Then
            volume = volume + current_volume       

        Else 
            volume = volume + current_volume
            closed_price = ws.cells(row, 6).value
            'deal with both open and closed price equalling zero
            if volume = 0 Then
                ws.Cells(summary_row, 10).Value = "0"
                ws.Cells(summary_row, 11).Value = "0"
            Else
                price_change = closed_price - open_price
                percent_change = price_change/open_price
                ws.Cells(summary_row, 10).Value = price_change
                ws.Cells(summary_row, 11).Value = percent_change
            End If
            'Display stock ticker and volume
            ws.Cells(summary_row, 9).Value = current_stock
            ws.Cells(summary_row, 12).Value = volume            
            'format cells 
            ws.Cells(summary_row, 11).NumberFormat = "0.00%"
            if percent_change >= 0 Then
                ws.Cells(summary_row, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(summary_row, 11).Interior.ColorIndex = 3
            End If
            summary_row = summary_row +1                
            volume = 0
            open_price = ws.cells(row + 1, 3).value
        End If
    Next row
Next ws
starting_ws.Activate
End Sub