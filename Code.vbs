Sub Ticker()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Activate
  
      ' Set an initial variable for holding the ticker name
      Dim Ticker_Name As String
    
      ' Set an initial variable for holding the total per ticker
      Dim Vol_Total As Double
      Vol_Total = 0
    
      ' Set an initial variable for holding the open price
      Dim Opent As Double
      Opent = Cells(2, 3).Value
      
        ' Set an initial variable for holding the closing price
      Dim closet As Double
      closet = 0
    
      ' Keep track of the location for each ticker name in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
             
        ' Set an initial variable for gratest increase
      Dim increase As Double
     
        ' Set an initial variable for gratest decrease
      Dim decrease As Double

        ' Set an initial variable for gratest total volume
      Dim greatestvol As Double

        
        'Adding the headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
      
     ' Finding the last row
      lastrow = Cells(Rows.Count, 1).End(xlUp).Row
          
         ' Loop trough all tickers name
         For i = 2 To lastrow
         
             ' Check if we are still within the same ticker name, if it is not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
         
        
              ' Set the Ticker name
              Ticker_Name = Cells(i, 1).Value
        
              ' Add to the Ticker Total
              Vol_Total = Vol_Total + Cells(i, 7).Value
              
              'Set close price
              closet = Cells(i, 6).Value
        
              ' Print Ticker name in the Summary Table
              Range("I" & Summary_Table_Row).Value = Ticker_Name
        
              ' Print the Total volume to the Summary Table
              Range("L" & Summary_Table_Row).Value = Vol_Total
              
              'Print the yearly change
              Range("J" & Summary_Table_Row).Value = closet - Opent
              
              'Changing color of the cell
              If Range("J" & Summary_Table_Row).Value < 0 Then
                 Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
              ElseIf Range("J" & Summary_Table_Row).Value > 0 Then
                 Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
              End If
              
              'Print the percent change
              Range("K" & Summary_Table_Row).Value = closet / Opent - 1

              'Switch to % format
              Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
              ' Add one to the summary table row
              Summary_Table_Row = Summary_Table_Row + 1
              
              ' Reset the Vol Total
              Vol_Total = 0
        
              'New open price
              Opent = Cells(i + 1, 3).Value
        
            ' If the cell immediately following a row is the same ticker...
            Else
        
              ' Add to the ticker Total
              Vol_Total = Vol_Total + Cells(i, 7).Value
        
            End If
         
         Next i

 
      ' Set an initial variable for gratest increase
    increase = Cells(2, 11).Value
      
      ' Set an initial variable for gratest decrease
    decrease = Cells(2, 11).Value
      
      ' Set an initial variable for gratest total volume
    greatestvol = Cells(2, 12).Value
    
      ' Finding the last row for the Precent change
    lastrow2 = Cells(Rows.Count, 9).End(xlUp).Row
      
      ' Loop trough summary table to get the top numbers
      
        For j = 3 To lastrow2
        
            If increase < Cells(j, 11).Value Then
            increase = Cells(j, 11).Value
            End If
                    
            If decrease > Cells(j, 11).Value Then
            decrease = Cells(j, 11).Value
            End If
            
            If greatestvol < Cells(j, 12).Value Then
            greatestvol = Cells(j, 12).Value
            End If
        
        Next j
        
      ' Print the values and change format
      Cells(2, 17).Value = increase
      Cells(2, 17).NumberFormat = "0.00%"
      Cells(3, 17).Value = decrease
      Cells(3, 17).NumberFormat = "0.00%"
      Cells(4, 17).Value = greatestvol
 
 
     ' Loop trough summary table to get names for the top numbers
         
         For k = 2 To lastrow2
        
        If Cells(k, 11).Value = Cells(2, 17).Value Then
        Cells(2, 16).Value = Cells(k, 9).Value
        End If
        
        If Cells(k, 11).Value = Cells(3, 17).Value Then
        Cells(3, 16).Value = Cells(k, 9).Value
        End If
        
        If Cells(k, 12).Value = Cells(4, 17).Value Then
        Cells(4, 16).Value = Cells(k, 9).Value
        End If
        
        Next k



Next ws

End Sub