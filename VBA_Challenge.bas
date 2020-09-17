Attribute VB_Name = "Module1"
Sub VBA_Challenge()

'Create and Set initial variables
    Dim ws As Worksheet
    Dim ticker_symbol As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock_volume As Double
    Dim open_value As Double
    Dim close_value As Double
    Dim ws_last_row As Long
    Dim i As Long 'ticker symbol row data
    Dim j As Long 'ticker symbol column data

    
  'Loop through all of the worksheets in the active workbook
    For Each ws In ActiveWorkbook.Worksheets
    
    'Set values for each worksheet
        j = 0
        total_stock_volume = 0
        yearly_change = 0
        open_value = 2
 
        'Add headers for results
        ws.Cells(1, 9).Value = "Stock Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'AutoFit Columns
        ws.Range("I:L").Columns.AutoFit
         
        
        
        'get the row number of the last row with data
        ws_last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To ws_last_row

                'Check if we are still within the same stock ticker if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    'Add to the Volume Total
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                    
                    'Handle zero total volume
                    If total_stock_volume = 0 Then
                        
                        'print the results
                        ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                        ws.Range("J" & 2 + j).Value = 0
                        ws.Range("K" & 2 + j).Value = 0
                        ws.Range("L" & 2 + j).Value = 0
    
                'If the cell following a row is the same ticker symbol...
                Else
                    'Find First non zero starting value
                    If ws.Cells(open_value, 3) = 0 Then
                        For ticker_value = open_value To i
                            If ws.Cells(ticker_value, 3).Value <> 0 Then
                                    open_value = ticker_value
                                    Exit For
                            End If
                        Next ticker_value
                    End If
                    
                    'Calculate Change
                    yearly_change = (ws.Cells(i, 6) - ws.Cells(open_value, 3))
                    percent_change = Round((yearly_change / ws.Cells(open_value, 3) * 100), 2)
                    
                    'Start of the next ticker symbol
                    open_value = i + 1
                    
                    'print the results
                        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                        ws.Range("J" & 2 + j).Value = Round(yearly_change, 2)
                        ws.Range("K" & 2 + j).Value = Format(percent_change / 100, "#.##%")
                        ws.Range("L" & 2 + j).Value = total_stock_volume
                    
                    'colors cells - positives green and negatives red
                    If yearly_change > 0 Then
                                ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                            Else
                               ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                           End If
    
                End If
                   
                'reset variables for the new ticker symbol
                yearly_change = 0
                total_stock_volume = 0
                j = j + 1
            
            'If ticker symbol is the same - add results
            Else
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
           End If
                
        Next i
    Next ws
End Sub
