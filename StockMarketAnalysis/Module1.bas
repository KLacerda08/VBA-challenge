Attribute VB_Name = "Module1"
Sub Stock_market_analysis():

Dim ws As Variant

    For Each ws In Worksheets
        
        'Declare Variables
        
        Dim open_price As Double
        Dim close_price As Double
        Dim volume_total As Double
        Dim ticker As String
        Dim yearly_change As Double
        Dim dec_pct_change As Double
        Dim Summary_Table_Row As Double
        
        Dim maxRow As Double
        Dim i As Double
        Dim j As Double
        
        'Find max row of dataset on each worksheet
        maxRow = ws.UsedRange.Rows.Count
        
        'assign column fields for summary table on each sheet
        
        ws.Cells(1, "i").Value = "Ticker"
        ws.Cells(1, "j").Value = "Yearly Change"
        ws.Cells(1, "k").Value = "Percent Change"
        ws.Cells(1, "l").Value = "Total Stock Volume"
        
        'assign first row of summary table (header is row 1)
        Summary_Table_Row = 2
        
            'Run loop to find and calculate values within the dataset for each ticker
            For i = 2 To maxRow
            
                'Find open_price for each ticker, except when open_price = 0
                'when the ticker changes, and i is the first row with the next ticker (first day of the year), open_price is in row i
                              
                If ws.Cells(i - 1, "A").Value <> ws.Cells(i, "A").Value And open_price <> 0 Then
                    open_price = ws.Cells(i, "c").Value
                                           
                    'if the open_price = 0 (on the first day of the year), find the next open_price within the ticker dataset
                    'set open_price to the price on the day it opened
                    'if open price does not change, return the last occurrence of open_price prior to ticker change
                        
                ElseIf open_price = 0 Then
                        
                    If ws.Cells(i + 1, "C").Value <> ws.Cells(i, "C").Value Then
                    open_price = ws.Cells(i, "c").Value
                    
                        ElseIf ws.Cells(i + 1, "a").Value <> ws.Cells(i, "a").Value Then
                        open_price = ws.Cells(i, "c").Value
                                       
                    End If
                    
                End If
                
                'Find close_price - when the ticker changes, and i is the last row within the ticker, close_price is row i
                'identify the ticker in row i
                                
                If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
                    
                    close_price = ws.Cells(i, "F").Value
                    ticker = ws.Cells(i, "A").Value
                    
                    'add last volume in ticker dataset to volume_total
                    volume_total = volume_total + ws.Cells(i, "g").Value
                    
                   'Calculate yearly change
                    yearly_change = close_price - open_price
                        
                        'Calculate percent change as a decimal percent. Account for open_price = zero; zero cannot be in denominator for calc of percent change
                        If open_price = 0 Then
                            dec_pct_change = 0
                        Else:
                            dec_pct_change = (yearly_change / open_price)
                        End If
        
                    'assign data for the ticker to summary table row
                    ws.Cells(Summary_Table_Row, "i").Value = ticker
                    ws.Cells(Summary_Table_Row, "j").Value = yearly_change
                    ws.Cells(Summary_Table_Row, "k").Value = dec_pct_change
                    ws.Cells(Summary_Table_Row, "l").Value = volume_total
                    ws.Cells(Summary_Table_Row, "K").Value = FormatPercent(dec_pct_change, 2, False, False)
                    
                    'go to the next summary table row for the next ticker dataset
                    Summary_Table_Row = Summary_Table_Row + 1
        
                    'Reset volume_total for next ticker
                    volume_total = 0
               
                'add volume_total for ticker entries prior to change in ticker, to last volume in ticker dataset
                Else
        
                    volume_total = volume_total + ws.Cells(i, "g").Value
                    
                End If
            
            Next i
            
            'Establish separate loop for formatting. Code runs more efficiently this way.
            'j starts at 2, after the header row
            
            For j = 2 To maxRow
            
                If ws.Cells(j, "k").Value > 0 Then
                    ws.Cells(j, "k").Interior.ColorIndex = 4
                
                ElseIf ws.Cells(j, "k").Value < 0 Then
                    ws.Cells(j, "k").Interior.ColorIndex = 3
                    
                Else: ws.Cells(j, "k").Interior.ColorIndex = 0
                
                End If
            Next j

    'repeat loops for each subsequent worksheet
    Next ws
    
End Sub
