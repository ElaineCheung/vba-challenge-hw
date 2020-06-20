Sub multiyearstock()

    For Each ws In Worksheets
    
    ' Add Summary Table Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
    ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
      ' Set an initial variable for holding the TickerSymbol and Count how Frequently it Occurs
         Dim Ticker_Symbol As String
    
      ' Set initial variables for holding the opening and closing prices, yearl change, percent change, and total stock volume
      
        Dim Open_Price As Double
        Open_Price = 0
        
        Dim Closing_Price As Double
        Closing_Price = 0
        
        Dim Yearly_Change As Double
        Yearly_Change = 0
        
        Dim Percent_Change As Variant
        Percent_Change = 0
        
        Dim Stock_Volume As Double
        Stock_Volume = 0
        
        'Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2

        ' Loop through all ticker
        For i = 2 To LastRow
        
                ' Check if we are still within the same ticker, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                 ' Set the Ticker Symbol
                Ticker_Symbol = ws.Cells(i, 1).Value
                    
                    ' Print the Ticker Symbol in the Summary Table
                    ws.Cells(Summary_Table_Row, 9).Value = Ticker_Symbol
                    
                   ' Set the Opening Price for the Ticker
                    Open_Price = Application.WorksheetFunction.VLookup(ws.Cells(i, 1).Value, ws.Range("A:G"), 3, False)
   
                   ' Set the Closing Price for the Ticker
                    Closing_Price = Application.WorksheetFunction.VLookup(ws.Cells(i, 1).Value, ws.Range("A:G"), 3, True)
        
                    'Determine the Yearly Change
                    Yearly_Change = Closing_Price - Open_Price
                    
                    'Determine the Percent Change
                        If Open_Price = 0 Then
                            Percent_Change = Null
                                    Else
                                        Percent_Change = Yearly_Change / Open_Price
                                    End If
                
                    'Add to the Total Volume
                      Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                    
                     ' Print the Yearly Change to the Summary Table
                     ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                    
                       'Conditional formatting for turning cells into green or red
                  
                        If ws.Cells(Summary_Table_Row, 10).Value < 0 Then
                            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                        Else: ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                        End If
                      
                        
                    ' Print the Percent Change to the Summary Table
                    ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
                    ws.Cells(Summary_Table_Row, 11).Style = "Percent"
                
                    'Print the Yearly Change to the Summary Table
                    ws.Cells(Summary_Table_Row, 12).Value = Stock_Volume
            
                    ' Add one to the summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
                        
                    ' Reset the opening price, closing price, yearly change, percent change, and stock volume
                    Open_Price = 0
                    Closing_Price = 0
                    Yearly_Change = 0
                    Percent_Change = 0
                    Stock_Volume = 0
                    Ticker_Count = 0
            
                Else
                    
                    'Add to the Total Volume
                    Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                                    
                End If
            
        Next i
        
     'Challenge
     
     'Create variables for challenge - determining value and ticker for greatest increase, greatest decrease, and greatest total volume
     
      Dim Great_Increase As Double
      Dim Increase_Position As Single
      Dim Ticker_Increase As String
      Dim Great_Decrease As Double
      Dim Decrease_Position As Single
      Dim Ticker_Decrease As String
      Dim Great_Volume As Double
      Dim Volume_Position As Single
      Dim Ticker_Volume As String
      
      'Find value of greatest increase
      
      Great_Increase = Application.WorksheetFunction.Max(ws.Range("J:J"))
      
      
       'Find ticker associated with greatest increase
       
       Increase_Position = Application.WorksheetFunction.Match(Great_Increase, ws.Range("J:J"), 0)
       Ticker_Increase = Application.WorksheetFunction.Index(Range("I:J"), Increase_Position, 1)
       
        'Print greatest increase ticker and value
        
      ws.Cells(2, 17).Value = Great_Increase
      ws.Cells(2, 16).Value = Ticker_Increase
      
      'Find  value of greatest decrease
      
      Great_Decrease = Application.WorksheetFunction.Min(ws.Range("J:J"))
      
        'Find ticker associated with greatest decrease
        Decrease_Position = Application.WorksheetFunction.Match(Great_Decrease, ws.Range("J:J"), 0)
       Ticker_Decrease = Application.WorksheetFunction.Index(Range("I:J"), Decrease_Position, 1)
    
    'Print greatest decrease
      ws.Cells(3, 17).Value = Great_Decrease
       ws.Cells(3, 16).Value = Ticker_Decrease
      
    'Find greatest total volume
      
      Great_Volume = Application.WorksheetFunction.Max(ws.Range("L:L"))
      
      'Find ticker associated with greatest total volume
       
        Volume_Position = Application.WorksheetFunction.Match(Great_Volume, ws.Range("L:L"), 0)
       Ticker_Volume = Application.WorksheetFunction.Index(Range("I:J"), Volume_Position, 1)
      
      'Print greatest total volume
      ws.Cells(4, 17).Value = Great_Volume
      ws.Cells(4, 16).Value = Ticker_Volume

    Next ws

End Sub

