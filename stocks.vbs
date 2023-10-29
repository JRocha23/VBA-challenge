Sub stocks():
For Each ws In Worksheets
    ws.Cells(1, "I") = "Ticker"
    ws.Cells(1, "J") = "Yearly Change"
    ws.Cells(1, "K") = "Percentage Change"
    ws.Cells(1, "L") = "Total Stock Vol"
    ws.Cells(1, "O") = "Ticker"
    ws.Cells(1, "P") = "Value"
    ws.Cells(2, "N") = "Greatest % Increase"
    ws.Cells(3, "N") = "Greatest % Decrease"
    ws.Cells(4, "N") = "Greatest Total Vol"

Dim total, table_row, open_price, close_price, increase, decrease, vol As Double
Dim ticker As String
total = 0
table_row = 2
increase = 0
decrease = 0
vol = 0

open_price = ws.Cells(2, "C")

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    total = total + ws.Cells(i, "G")

    ticker = ws.Cells(i, "A")

    If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
        ws.Cells(table_row, "L") = total

        ws.Cells(table_row, "I") = ticker
        
        close_price = ws.Cells(i, "F")
        
        ws.Cells(table_row, "J") = close_price - open_price
        
        If ws.Cells(table_row, "J") > 0 Then
            ws.Cells(table_row, "J").Interior.ColorIndex = 4
        Else
            ws.Cells(table_row, "J").Interior.ColorIndex = 3
            
        End If
        
        If open_price > 0 Then
            ws.Cells(table_row, "K") = FormatPercent((close_price - open_price) / open_price, 2)
            
        Else
            ws.Cells(table_row, "K") = 0
            
        End If

        If ws.Cells(table_row, "K") > increase Then
            increase = ws.Cells(table_row, "K")
            increase_ticker = ws.Cells(table_row, "I")
        End If
        
        If ws.Cells(table_row, "K") < decrease Then
            decrease = ws.Cells(table_row, "K")
            decrease_ticker = ws.Cells(table_row, "I")
        End If
        
        If ws.Cells(table_row, "L") > vol Then
            vol = ws.Cells(table_row, "L")
            vol_ticker = ws.Cells(table_row, "I")
        End If
                
        total = 0
        open_price = ws.Cells(i + 1, "C")
        table_row = table_row + 1

    End If

Next i

ws.Cells(, "O") = increase_ticker
ws.Cells(2, "P") = increase

ws.Cells(3, "O") = decrease_ticker
ws.Cells(3, "P") = decrease

ws.Cells(4, "O") = vol_ticker
ws.Cells(4, "P") = vol

Next ws

End Sub
