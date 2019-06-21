Sub stock_data_easy()

'Loop through all sheets

    For Each ws In Worksheets
    
        'Set initial variable for holding the tick name
        Dim ticker As String

        'Set initial variable holding the volume per ticker
        Dim volume_total As Double
        volume_total = 0

        'Keep track of the location for each ticker
        Dim ticker_table_row As Integer
        ticker_table_row = 2
    
        'Create a variable to hold file name, and last row
        Dim WorksheetName As String
    
        'Last_Row of spreadsheets
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        WorksheetName = ws.Name
    
'Loop through all tickers
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Set ticker name
            ticker = ws.Cells(i, 1).Value
            
            'Add to the volume_total
            volume_total = volume_total + ws.Cells(i, 7).Value
            
            'Print tickers in summary table
            ws.Cells(1, 9).Value = "Tickers"
            ws.Range("I" & ticker_table_row).Value = ticker
            ws.Cells(1, 10).Value = "Total Stock Volume"
            ws.Range("J" & ticker_table_row).Value = volume_total
            
            'Add one to the ticker_table_row
            ticker_table_row = ticker_table_row + 1
            
            'Reset the volume_total
            volume_total = 0
        'If the cell immediately following the row is the same ticker
        Else
            'Add the volume total
            volume_total = volume_total + ws.Cells(i, 7).Value
        End If
        
    Next i
    
Next ws
        
End Sub

