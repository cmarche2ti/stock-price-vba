Sub StockSummary()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
        ws.cells(1,9).value = "Ticker"
        ws.cells(1,10).value = "Yearly Change"
        ws.cells(1,11).value = "Percent Change"
        ws.cells(1,12).value = "Total Stock Volume"
        rowEnd = ws.Cells(2, 1).End(xlDown).Row
        volume = 0 
        summary_row = 2

        For row = 2 to rowEnd 
            current_stock = ws.Cells(row, 1).Value
            next_stock = ws.Cells(row + 1, 1).Value
            current_volume = ws.Cells(row, 7).Value 

            If current_stock = next_stock Then 
                volume = volume + current_volume
            Else 
                volume = volume + current_volume
                ws.Cells(summary_row, 9).Value = current_stock
                ws.Cells(summary_row, 12).Value = volume
                summary_row = summary_row +1
                volume = 0
            End If

        Next row
Next ws
starting_ws.Activate
End Sub