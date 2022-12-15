Attribute VB_Name = "Module1"
Sub finished_analysis()


For Each ws In Worksheets

Dim worksheetName As String

Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


  
   ' Set an initial variable for holding the ticker name
  Dim ticker As String

  ' Set an initial variable for holding the total per credit card brand
  Dim yearly_change As Double
  

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'an initial variable to hold opening price to use later
  Dim year_open_price As Double
  
  'an initial variable to hold closing price to be used later
  Dim year_close_price As Double
  
  Dim percent_change As Double
  
  Dim round_percent As String
  
  Dim total_stock_volume As LongLong
  
  
  

  ' Loop through rows in the column
  For i = 2 To Lastrow
  
   If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        year_open_price = ws.Cells(i, 3).Value
        
        
         
    End If
    

    ' Searches for when the value of the next cell is different than that of the current cell
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      
       ' Set the ticker name
      ticker = ws.Cells(i, 1).Value
      
      ' close price
      year_close_price = ws.Cells(i, 6).Value
      
      yearly_change = year_close_price - year_open_price

      ' Print the ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker

      ' Print the Brand Amount to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = yearly_change
      
      
      percent_change = (yearly_change / year_open_price) * 100
      
      round_percent = Round(percent_change, 2)
      
      'format string for %
      
      ws.Cells(Summary_Table_Row, 11).Value = round_percent + "%"
      
      ' Add to the stock volume Total
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value


      ' Print the stock volume total Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = total_stock_volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the stock volume total
      total_stock_volume = 0

    Else

      ' Add to the stock volume Total
      total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

    
  End If
        ' colouring for gain
    If ws.Cells(i, 10).Value > 0 Then
    
       ws.Cells(i, 10).Interior.ColorIndex = 4
       
       ws.Cells(i, 11).Interior.ColorIndex = 4
       
  End If
        'colouring for loss
     If ws.Cells(i, 10).Value < 0 Then
    
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
        ws.Cells(i, 11).Interior.ColorIndex = 3
        
  End If
      'colouring for no movement
    If ws.Cells(i, 10).Value = "0" Then
    
       ws.Cells(i, 10).Interior.ColorIndex = 15
       
       ws.Cells(i, 11).Interior.ColorIndex = 15
       
 End If
 
       
  Next i
  
  
  
  Dim max_increase_ticker As String
  
  Dim max_increase_value As Double
  
  Dim max_decrease_ticker As String
  
  Dim max_decrease_value As Double
  
  Dim greatest_stock_volume As LongLong
  
  Dim greatest_stock_ticker As String
  
  
  Lastrow2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
  
 
  
    For j = 2 To Lastrow2
    
    
    
        If ws.Cells(j, 11) > max_increase_value Then
    
        max_increase_value = ws.Cells(j, 11).Value
    
        max_increase_ticker = ws.Cells(j, 9).Value
    
        End If
        
        If ws.Cells(j, 11) < max_decrease_value Then
    
        max_decrease_value = ws.Cells(j, 11).Value
    
        max_decrease_ticker = ws.Cells(j, 9).Value
    
        End If
        
        If Cells(j, 12).Value > greatest_stock_volume Then
        
        greatest_stock_volume = Cells(j, 12).Value
        
        greatest_stock_ticker = Cells(j, 9).Value
        
        
        End If
        
        
    Next j
    
    ws.Cells(2, 16).Value = max_increase_ticker
    
    ws.Cells(2, 17).Value = max_increase_value
    
    
    ws.Cells(3, 16).Value = max_decrease_ticker
    
    ws.Cells(3, 17).Value = max_decrease_value
    
    ws.Cells(4, 16).Value = greatest_stock_ticker
    
    ws.Cells(4, 17).Value = greatest_stock_volume
        
  
  
  
  Next ws





End Sub
