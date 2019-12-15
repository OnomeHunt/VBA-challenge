Sub Stock()
    ' loop through sheets
    For Each wb In Worksheets
   
    Dim Worksheet As String
   
    Worksheet = wb.Name
    
    ' Create variable for holding the ticker symbol
    Dim ticker As String
    
    ' Create variable for holding the price change per year per ticker symbol
    Dim Yearly_Change As Double
    
    ' Create variable to hold open_price
    Dim open_price As Double
    open_price = wb.Cells(2, 3).Value
    
    ' Create variable to hold close_price
    Dim close_price As Double
    close_price = 0
    
    Dim lastrow As Long
    lastrow = wb.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' create variable for summary columns
    Dim new_column As Integer
    new_column = 1
    
    ' create postion to diplay summary results
    Dim new_row As Integer
    new_row = 2
    
    ' Get range to sue for min & max percentage_change for all combined tickcers
    Dim percentage_change_range As Range
    Set percentage_change_range = wb.Range("K2:K" & Rows.Count)
    
    ' Get range to use for max total volume for all combined tickers
    Dim greatest_total_volume_range As Range
    Set greatest_total_volume_range = wb.Range("L2:L" & Rows.Count)
    
    ' Convert percentage column to number to use to get max value
    percentage_change_range.NumberFormat = "0.00"
    
    greatest_increase = Application.WorksheetFunction.Max(percentage_change_range)
    greatest_decrease = Application.WorksheetFunction.Min(percentage_change_range)
    greatest_total_volume = Application.WorksheetFunction.Max(greatest_total_volume_range)
     
    
    
    ' Dim percentage_change As Variant
    
    Dim total_stock_volume As Double
    total_stock_volume = 0
    
    wb.Range("P" & 2) = "Ticker"
    wb.Range("Q" & 2) = "Value"
    
    wb.Range("P2").HorizontalAlignment = xlCenter
    wb.Range("Q2").HorizontalAlignment = xlCenter
    
    
    wb.Range("O" & 3) = "Greatest % Increase"
    wb.Range("O" & 4) = "Greatest % Decrease"
    wb.Range("Q" & 3) = greatest_increase
    wb.Range("Q" & 4) = greatest_decrease
    wb.Range("Q" & 3).NumberFormat = "0.00%"
    wb.Range("Q" & 4).NumberFormat = "0.00%"
    
    wb.Range("O" & 5) = "Greatest Total Volume"
    wb.Range("Q" & 5) = greatest_total_volume
    
   For j = 2 To lastrow
      If wb.Cells(j, 11).Value = wb.Cells(3, 17).Value Then
        Max_ticker_for_percentage_change = wb.Cells(j, 9).Value
        wb.Range("P3").Value = Max_ticker_for_percentage_change

    'ticker for the greatest % decrease
    ElseIf wb.Cells(j, 11).Value = wb.Cells(4, 17).Value Then
    Min_ticker_for_percentage_change = wb.Cells(j, 9).Value
    wb.Range("P4").Value = Min_ticker_for_percentage_change

    'ticker for the greatest volume total
    ElseIf wb.Cells(j, 12).Value = wb.Cells(5, 17).Value Then
    max_ticker_for_total_volume = wb.Cells(j, 9).Value
    wb.Range("P5").Value = max_ticker_for_total_volume
End If
Next j
  
     
    
    ' Loop through all stock
    For i = 2 To lastrow
         ticker = wb.Cells(i, 1).Value
         
        
         
                  
         If wb.Cells(i + 1, 1).Value <> wb.Cells(i, 1).Value Then
         
            close_price = wb.Cells(i, 6).Value
          
            Yearly_Change = close_price - open_price
            
            If Yearly_Change <> 0 And open_price <> 0 Then
                    
                percentage_change = (Yearly_Change / open_price)
            Else
                    percentage_change = 0
            End If
            
            total_stock_volume = total_stock_volume + wb.Cells(i, 7)
                
            wb.Range("I" & new_column) = "Ticker"
            wb.Range("I" & new_row) = ticker
            wb.Range("J" & new_column) = "Yearly Change"
            wb.Range("J" & new_row).Value = Yearly_Change
            wb.Range("K" & new_column) = "Percentage Change"
            wb.Range("K" & new_row).Value = percentage_change
            wb.Range("K" & new_row).NumberFormat = "0.00%"
            wb.Range("L" & new_column) = "Total Stock Volume"
            wb.Range("L" & new_row) = total_stock_volume
             
            If Yearly_Change < 0 Then
                wb.Range("K" & new_row).Interior.ColorIndex = 3
            Else
                wb.Range("K" & new_row).Interior.ColorIndex = 4
               End If
            
            
            
            
            
       ' Reset variables for next iteration
            open_price = wb.Cells(i + 1, 3)
            close_price = 0
            total_stock_volume = 0
            Yearly_Change = 0
            percentage_change = 0
            
            new_row = new_row + 1
            
            
        Else
            close_price = 0
            total_stock_volume = total_stock_volume + wb.Cells(i, 7)
    
        End If
            
        
        
    

    Next i
    
    
    
    
  Next wb





End Sub

