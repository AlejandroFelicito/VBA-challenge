Sub stock():
    Dim ws As Worksheet
    
    For Each ws In Worksheets: ' For Each starts
        Dim total_rows&, input_row&, output_row&, output_row_count&, count_ticker& ' Long Integer
        Dim total_stock^ ' Long Long Int
        Dim opening_price!, closing_price!, yearly_change!, percent_change!  ' Single
        Dim current_ticker$ ' String
    
        total_rows = ws.UsedRange.Rows.Count ' Total rows
        output_row = 2 ' Initial output row
        current_ticker = ws.Cells(2, 1).Value ' Initial ticker
        count_ticker = 0 ' Initial tickers count
        output_row_count = 1 ' Initial output row count
        total_stock = 0 ' Initial total stock
        opening_price = ws.Cells(2, 3).Value ' Initial opening price
        closing_price = ws.Cells(2, 6).Value ' Initial closing price
        yearly_change = 0 ' Initial yearly change
        percent_change = 0 ' Initial percent change
    
'       MsgBox ws.Name
'       MsgBox ws.UsedRange.Rows.Count
    
        ' Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
'        ws.Cells(1, 13).Value = "Ticker Count"
'        ws.Cells(1, 14).Value = "Output row Count"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
    
        For input_row = 2 To total_rows
            count_ticker = count_ticker + 1 ' Add 1 ticker occurrence
            total_stock = total_stock + ws.Cells(input_row, 7).Value ' Add stocks
            If ws.Cells(input_row + 1, 1).Value <> current_ticker Then ' If next row ticker is different than current line ticker
                ws.Cells(output_row, 9).Value = current_ticker ' Output ticker
                closing_price = ws.Cells(input_row, 6) ' Final closing price
                yearly_change = closing_price - opening_price
                ws.Cells(output_row, 10).Value = yearly_change ' Output yearly change
                percent_change = yearly_change / opening_price
                ws.Cells(output_row, 11).Value = percent_change ' Output percent change
               
               ' Conditional formatting
               If yearly_change < 0 Then ' percent_change follows sign of yearly_change
                   ws.Cells(output_row, 10).Interior.ColorIndex = 3
                   ws.Cells(output_row, 11).Interior.ColorIndex = 3
               ElseIf yearly_change > 0 Then
                   ws.Cells(output_row, 10).Interior.ColorIndex = 4
                   ws.Cells(output_row, 11).Interior.ColorIndex = 4
               End If
                
                ws.Cells(output_row, 12).Value = total_stock ' Output total stock volume
'                ws.Cells(output_row, 13).Value = count_ticker ' Output ticker occurrences
                output_row_count = output_row_count + 1
'                ws.Cells(output_row, 14).Value = output_row_count ' Output row count
                output_row = output_row + 1
                count_ticker = 0
                total_stock = 0
                current_ticker = ws.Cells(input_row + 1, 1).Value ' Update ticker
                opening_price = ws.Cells(input_row + 1, 3).Value ' Update opening price
            End If
        Next input_row
    
        ' Start of added functionality
        For input_row = 2 To output_row_count ' row index in second table
            If ws.Cells(input_row, 12).Value > ws.Cells(4, 17).Value Then ' GOA Volume
                ws.Cells(4, 17).Value = ws.Cells(input_row, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(input_row, 9).Value
            End If
            
            If ws.Cells(input_row, 11).Value < ws.Cells(3, 17).Value Then ' GOA % Decrease
                ws.Cells(3, 17).Value = ws.Cells(input_row, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(input_row, 9).Value
            End If
            
            If ws.Cells(input_row, 11).Value > ws.Cells(2, 17).Value Then ' GOA % Increase
                ws.Cells(2, 17).Value = ws.Cells(input_row, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(input_row, 9).Value
            End If
        Next input_row
    Next ' For Each ends
End Sub

