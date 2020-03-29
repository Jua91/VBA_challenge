Sub ticker():

    'Loop through each worksheets
    For Each ws In Worksheets
    
    'Set the headers of summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Set the lastRow
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set the variables for ticker, stock volume and summary table row
    Dim ticker_symbol As String
    Dim stock_volume As Double
    'Set the stock volume 0
    stock_volume = 0
    Dim summaryTable_row As Integer
    'Set the summary table row 2
    summaryTable_row = 2
    
    'Set the variables for open price, close price, yearly price, and percent change
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    open_price = ws.Cells(2, 3).Value
    
    'Loop through each row in the worksheet
        For i = 2 To lastRow
    
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                 'Assign the value of ticker
                 ticker_symbol = ws.Cells(i, 1).Value
                 'Add up the stock volume
                 stock_volume = stock_volume + ws.Cells(i, 7).Value
                 'Assign value of close price
                 close_price = ws.Cells(i, 6).Value
                 'Set the formula of yearly change
                 yearly_change = close_price - open_price
                  
                  'It's not divisible when the denominator (open_price) is 0, so set the percent change to  0 when the open price is 0.
                  If open_price = 0 Then
                    percent_change = 0
                  Else: percent_change = yearly_change / open_price
                  End If
                 
                 'Assign values to appropriate cells
                 ws.Range("I" & summaryTable_row).Value = ticker_symbol
                 ws.Range("J" & summaryTable_row).Value = yearly_change
                 ws.Range("K" & summaryTable_row).Value = percent_change
                 ws.Range("L" & summaryTable_row).Value = stock_volume
                 
                'Change the percent_change cells to percent format
                 ws.Range("K" & summaryTable_row).Style = "Percent"
                 
                'Move to next summaryTable_row by adding one
                 summaryTable_row = summaryTable_row + 1
                 
                 'Reset the stock_volume
                 stock_volume = 0
                 
                 'Reset the open price value
                 open_price = ws.Cells(i + 1, 3).Value
            Else
                'Add up the stock volume
                stock_volume = stock_volume + ws.Cells(i, 7).Value

            
            End If
        
        Next i
    
        'Set the last row of summary table
        lastRow_summaryTable = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
       'Loop through the summary table
        For i = 2 To lastRow_summaryTable
        
            If ws.Cells(i, 10).Value > 0 Then
                'If the value is positive, fill the cells with the green color
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                'Otherwise, fill the cells with the red color
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
            
        Next i
        
        
        'Set the header of the new table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Loop through the summary table
        For i = 2 To lastRow_summaryTable
            
            'Find the maximum, minimum percent changes, and the greatest total volume using conditionals
            If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRow_summaryTable)) Then
                'Assign the ticker symbol to the ticker cell
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                'Assign the maximum percent_change to the greatest % increase cell
                ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
                'Format the cell as percentage
                ws.Cells(2, 17).Style = "Percent"
            
            ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRow_summaryTable)) Then
                'Assign the ticker symbol to the ticker cell
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                'Assign the minimum percent_change to the greatest % decrease cell
                ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                'Format the cell as percentage
                ws.Cells(3, 17).Style = "Percent"
            
            ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRow_summaryTable)) Then
                'Assign the ticker symbol to the ticker cell
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                'Assign the maximum stock volume to the greatest total volume cell
                ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
            End If
        
        Next i
           
    Next ws
    
End Sub



