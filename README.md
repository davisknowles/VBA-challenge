# VBA-challenge
Summary_Table_Row = 2
      Last_Row = Cells(Rows.Count, "A").End(xlUp).Row
      open_price = Cells(2, 3).Value
        
            ' Loop through all rows for one year
            For i = 2 To Last_Row
                
                ' Check if we are still within the same Ticker Name
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                    ' Pull values
                    ticker_name = Cells(i, 1).Value
                     total_stock_volume = total_stock_volume + Cells(i, 7).Value
                     close_price = Cells(i, 6).Value
                    'calculate yearly change
                    yearly_change = close_price - open_price
                    'calculate percent change
                    percent_change = (yearly_change / open_price)
                    Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                 
                    ' Print in the Summary Table
                    Range("I" & Summary_Table_Row).Value = ticker_name
                    Range("L" & Summary_Table_Row).Value = total_stock_volume
                    Range("J" & Summary_Table_Row).Value = yearly_change
                    Range("K" & Summary_Table_Row).Value = percent_change
                    
                    'conditional formatting for highlight and color
                    If yearly_change >= 0 Then
                        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                        Else
                        Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    End If
        
                    ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
              
                    ' Reset the total stock volume
                    total_stock_volume = 0
                    

                ' If the cell immediately following a row is the same Ticker...
                Else
        
                    ' Add to the total stock volume
                    total_stock_volume = total_stock_volume + Cells(i, 7).Value
            End If
        
          Next i

End Sub

