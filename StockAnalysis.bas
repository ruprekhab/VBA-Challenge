Attribute VB_Name = "Module1"
Sub multiquarter_stock()
        
        'Declare ws as a worksheet object variable
        Dim ws As Worksheet
        'Loop through each worksheet in the worksbook
        For Each ws In Worksheets
        
       'create headers for summary table from column J to M
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Quarterly Change"
        ws.Cells(1, 12).Value = "Percentage Change"
        ws.Cells(1, 13).Value = "Total Stock volume"
       
        
        'create a variable i as a counter for the loop
        Dim i As Long
        
        'set variables to hold ticker related data
        Dim ticker_name As String
        Dim total_volume As Double
        Dim summary_table_row As Integer
        summary_table_row = 2
        
        
        'set variable for holding the opening price, closing price to calculate Quarterly Price Change, Percentage Change
        Dim opening_price As Double
        Dim closing_price As Double
        Dim price_change As Double
        Dim percentage_change As Double
                
        'set a variable to hold the last row.
        Dim lastrow As Long
                
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        opening_price = ws.Cells(2, 3).Value
                
        'Loop through tickers in column A
         For i = 2 To lastrow
                ticker_name = ws.Cells(i, 1).Value                                                                          'get the current ticker name
                   
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then                                            'check if the next row contains a different tickers
                            
                            ws.Range("J" & summary_table_row).Value = ticker_name                           'ticker names in column J
                    
                            total_volume = total_volume + ws.Cells(i, 7).Value                                      'calculating total volume
                            
                            ws.Range("M" & summary_table_row).Value = total_volume                         'Total volume for each ticker in Column M
                            
                            total_volume = 0                                                                                         'set initial total volume to 0
                            
                            closing_price = ws.Cells(i, 6).Value                                                              'storing the closing value at the end of the quarter
                            price_change = closing_price - opening_price                                             'calculate the quarterly price change for each ticker
                            
                            percentage_change = ((closing_price - opening_price) / opening_price)       'calculating percentage change for each ticker
                            ws.Range("L" & summary_table_row).Value = percentage_change                 'insert the percentage change in column L
                            ws.Range("L:L").NumberFormat = "0.00%"                                                     'format percentage change column
                            
                            
                            ws.Range("K" & summary_table_row).Value = price_change                          'insert the Quarterly change in column K
                            opening_price = ws.Cells(i + 1, 3).Value                                                       'set opening price
                            summary_table_row = summary_table_row + 1                                             'move to the next row in the summary table
                                
                        
                        Else                                                                                                                   'if the rows contain same ticker
                            total_volume = total_volume + ws.Cells(i, 7).Value                                        'calculate total volume
                            
                                
                            
                        End If
                        
                     
            Next i
                
       'conditional formatting for column K. Fill in green if the values are more than 0 and red if they are less than 0.
        For i = 2 To lastrow
                 If ws.Cells(i, 11) > 0 Then
                     ws.Cells(i, 11).Interior.ColorIndex = 4
                 ElseIf ws.Cells(i, 11) < 0 Then
                     ws.Cells(i, 11).Interior.ColorIndex = 3
                End If
         Next i
         
         'create greatest % increase, greatest % decrease, greatest total volume table
         ws.Range("O2").Value = "Greatest Increase"
         ws.Range("O3").Value = "Greatest Decrease"
         ws.Range("O4").Value = "Greatest Volume"
         ws.Range("P1").Value = "Ticker"
         ws.Range("Q1").Value = "Value"
         
        
         Dim maxvalue As Double               'set variable to hold the maximum increase in percentage change
         Dim minvalue As Double                'set the variable to hold the maximum decrease in percentage change
         Dim greatest_volume As Double     'set variable to hold maximum volume
  
         'find the maximum percentage increase, decrease and the volume
         maxvalue = Application.WorksheetFunction.Max(ws.Range("L1:L" & lastrow))               's
         minvalue = Application.WorksheetFunction.Min(ws.Range("L1:L" & lastrow))
         greatest_volume = Application.WorksheetFunction.Max(ws.Range("M1:M" & lastrow))
         
        ' Loop through the rows to find and write the corresponding tickers for max and min percentage changes and greatest volume
        For i = 2 To lastrow
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
         
                If ws.Cells(i, 12).Value = maxvalue Then
                ws.Range("Q2").Value = ws.Cells(i, 12).Value
                ws.Range("P2").Value = ws.Cells(i, 10).Value
                
                ElseIf ws.Cells(i, 12).Value = minvalue Then
                ws.Range("Q3").Value = ws.Cells(i, 12).Value
                ws.Range("P3").Value = ws.Cells(i, 10).Value
                
                
                ElseIf ws.Cells(i, 13).Value = greatest_volume Then
                ws.Range("Q4").Value = ws.Cells(i, 13).Value
                ws.Range("P4").Value = ws.Cells(i, 10).Value
            
             End If
            
            Next i
           
           
        Next ws
        MsgBox ("Summary table ready")  'After the script has finished looping through all the sheets, message box is displayed
        End Sub
      
        
        
        
        
      
        
 
     


