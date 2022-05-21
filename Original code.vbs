Sub tickersummary():

Dim current As Worksheet
For Each current In Worksheets

'set the summary table
    current.Cells(1, 9).Value = "Ticker"
    current.Cells(1, 10).Value = "Yearly Change"
    current.Cells(1, 11).Value = "Percentage Change"
    current.Cells(1, 12).Value = "Total Stock Volume"

    current.Cells(2, 15).Value = "Greatest % Increase"
    current.Cells(3, 15).Value = "Greatest % Decrease"
    current.Cells(4, 15).Value = "Greatest Total Volume"
    current.Cells(1, 16).Value = "Ticker"
    current.Cells(1, 17).Value = "Value"
    
' define intial variables
    last_row = current.Cells(Rows.Count, 1).End(xlUp).Row
    Summary_Table_Row = 2
    ticker_open = current.Cells(2, 3).Value
    ticker_sum = 0

    For i = 2 To last_row
    
        If current.Cells(i, 1) <> current.Cells(i + 1, 1) Then

' to determine ticker name & input to summary table
            Ticker_Name = current.Cells(i, 1).Value
            current.Cells(Summary_Table_Row, 9).Value = Ticker_Name

' to determine difference & % difference and to input to summary table
            ticker_close = current.Cells(i, 6).Value
            Ticker_diff = ticker_close - ticker_open
            
            'format the cell
            If Ticker_diff >= 0 Then
                current.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            Else
                current.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            End If
        
            ' account for open stock is o
            If ticker_open = 0 Then
                ticker_per = "NA"
            Else
                ticker_per = FormatPercent(Ticker_diff / ticker_open)
            End If
            
            current.Cells(Summary_Table_Row, 10).Value = Ticker_diff
            current.Cells(Summary_Table_Row, 11).Value = ticker_per

' to determine total volume stock & input to summary table
            ticker_sum = ticker_sum + current.Cells(i, 7).Value
            current.Cells(Summary_Table_Row, 12).Value = ticker_sum
    
' to assign new values
            Summary_Table_Row = Summary_Table_Row + 1
            ticker_open = current.Cells(i + 1, 3).Value
            ticker_sum = 0

        Else
            ticker_sum = ticker_sum + current.Cells(i, 7).Value

        End If
        
    Next i

Next

End Sub