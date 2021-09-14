Attribute VB_Name = "Module1"
Sub Stock_Data()

    
    
'loop through all sheets
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
    'output The ticker symbol.
    'assign a variable to ticker symbol
         Dim ticker As String
      
      'set variable for specifying the column of interest
         Dim column As Integer
        column = 1
      
    'set an inital variable for holding the total per ticker
         Dim total_stock_volume As Double
        total_stock_volume = 0
    
    'location for each ticker in the table
         Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
    'set variable for open price for Ticker
         Dim opening_price As Double
        
    'variable for closing price
         Dim closing_price As Double
        
    'variable for yearly change
          Dim yearly_change As Double
        
    'variable for percent change
          Dim percent_change As Double
        
    'loop through all stocks for 1 year
          For i = 2 To LastRow
        
        'get open price for ticker
                If i = 2 Then
                    opening_price = ws.Cells(i, 3).Value
              End If
        
        
        
        'check if the ticker is the same, if it is not
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'MsgBox (opening_price)
            'set ticker name
                    ticker = ws.Cells(i, 1).Value
                
            'set total volume name
                 total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            
            'set closing price for ticker
                  closing_price = ws.Cells(i, 6).Value
                  
            'Set opening price for ticker
                opening_price = ws.Cells(i + 1, 3).Value
            
                'MsgBox (opening_price)
            'print the ticker in the summary table
                 ws.Range("I" & Summary_Table_Row).Value = ticker
                
             'print the total stock volume in the summary table
                  ws.Range("L" & Summary_Table_Row).Value = total_stock_volume
                
            'calculate yearly change
            
                'MsgBox (opening_price)
                'MsgBox (closing_price)
                
                    yearly_change = closing_price - opening_price
                    ws.Range("J" & Summary_Table_Row).Value = yearly_change
            
                'MsgBox (yearly_change)
                
                'avoid 0s in percent change forumula
                If opening_price <> 0 Then
            'caluclate percent change
                    percent_change = yearly_change / opening_price
                    ws.Range("K" & Summary_Table_Row).Value = percent_change
                
                Else
                
                    percent_change = 0
                
                End If
             
                
            'Format percent change as percent
                    ws.Range("K" & Summary_Table_Row).Style = "Percent"
                
            'if yearly change is positive fill cell gren
                 If yearly_change >= 0 Then
                    
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    
                    ElseIf yearly_change < 0 Then
                
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    End If
                
            'add one to summary table row
                    Summary_Table_Row = Summary_Table_Row + 1
            
            'reset total stock volume
                    total_stock_volume = 0
                
                    yearly_change = 0
                
                    percent_change = 0
                
                Else
            
                'set total volume name
                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
         
            
        'MsgBox (yearly_change)
        
        'add to ticker column
            End If
        
        
            Next i
        Next ws
    
    'Output Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    
    
    'Output The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'conditionals to change color
    
    
    'Output The total stock volume of the stock.
    
    
End Sub
