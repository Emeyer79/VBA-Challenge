Sub StockMarketVBA():

'Start by creating a loop across all worksheets
Dim ws As Worksheet
For Each ws In Worksheets

    'Create summary and performance headers across all worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    'Create dimensions across all sheets
    Dim ticker As String
    Dim Vol_Total As Double
    total_vol = 0
    Dim rowcount As Long
    rowcount = 2
    'Use this to capture the first open
    Dim First_Open As Double
    First_Open = 0
    'Use this to capture the last close
    Dim Last_Close As Double
    Last_Close = 0
    'Use this to calculate the year change last_close - first_open
    Dim year_change As Double
    year_change = 0
    'Use this to calculate the yearly percent change
    Dim percent_change As Double
    percent_change = 0
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Use these dimensions for the bonus section performance
    Dim best_performer As String
    Dim best_value As Double
    Dim worst_performer As String
    Dim worst_value As Double
    Dim most_vol_performer As String
    Dim most_vol_value As Double
    
    
        'Set the inner loop for each worksheet to capture summary numbers using i
        For i = 2 To lastrow
        
            'Create Ticker change, Yearly difference and percent change formulas
            'Open Price formula
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                First_Open = ws.Cells(i, 3).Value
                
            End If
            
               'Total volume- Use Credit Card Lesson as reference
            Vol_Total = Vol_Total + ws.Cells(i, 7)

            'Ticker symbol change- Use Credit Card Lessons as example
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            'Move ticker symbol to summary table
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value

                'Move total stock volume to the summary table
                ws.Cells(rowcount, 12).Value = Vol_Total

                'Last_Close formula
                Last_Close = ws.Cells(i, 6).Value

                'Yearly price change formula
                year_change = Last_Close - First_Open
                ws.Cells(rowcount, 10).Value = year_change

                'Conditional to format to highlight positive or negative change.
                If year_change >= 0 Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If

                'Yearly Percent change formula
                If First_Open = 0 And Last_Close = 0 Then
                    percent_change = 0
                    ws.Cells(rowcount, 11).Value = percent_change
                     ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                    
                ElseIf First_Open = 0 Then
                
                    Dim percent_change_NA As String
                    percent_change_NA = "New Stock"
                    ws.Cells(rowcount, 11).Value = percent_change
                Else
                    percent_change = year_change / First_Open
                    ws.Cells(rowcount, 11).Value = percent_change
                     ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                    
                End If

                rowcount = rowcount + 1

                
                Vol_Total = 0
                First_Open = 0
                Last_Close = 0
                year_change = 0
                percent_change = 0
                
            End If
            
        Next i
        
        'Set last row for the new columns
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'Set best performer equal to the first stock
        best_value = ws.Cells(2, 11).Value

        'Set worst performer equal to the first stock
        worst_value = ws.Cells(2, 11).Value

        'Set most volume equal to the first stock
        most_vol_value = ws.Cells(2, 12).Value

        'Loop to search through summary table
        For j = 2 To lastrow

            'Conditional to determine best performer
            If ws.Cells(j, 11).Value > best_value Then
                best_value = ws.Cells(j, 11).Value
                best_performer = ws.Cells(j, 9).Value
            End If

            'Conditional to determine worst performer
            If ws.Cells(j, 11).Value < worst_value Then
                worst_value = ws.Cells(j, 11).Value
                worst_performer = ws.Cells(j, 9).Value
            End If

            'Conditional to determine stock with the greatest volume traded
            If ws.Cells(j, 12).Value > most_vol_value Then
                most_vol_value = ws.Cells(j, 12).Value
                most_vol_performer = ws.Cells(j, 9).Value
            End If

        Next j

        'Move best performer, worst performer, and stock with the most volume items to the performance table
        ws.Cells(2, 16).Value = best_performer
        ws.Cells(2, 17).Value = best_value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_performer
        ws.Cells(3, 17).Value = worst_value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_performer
        ws.Cells(4, 17).Value = most_vol_value
        
         'Autofit
         ws.Columns("I:L").EntireColumn.AutoFit
         ws.Columns("O:Q").EntireColumn.AutoFit
         ws.Range("A1:Q1").Font.Bold = True
         
    
            
Next ws

End Sub


