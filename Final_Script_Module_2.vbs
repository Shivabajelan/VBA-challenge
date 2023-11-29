Sub TotalStockVolum()

 'Loop through all the worksheets
 For Each ws In ThisWorkbook.Sheets

 ' Set an initial variable for holding the ticker name
 Dim Ticker As String
 
 ' Set an initial variable for holding the total volume per credit ticker name
 Dim Total_Stock_Volume As Double
 Total_Stock_Volume = 0
 
 ' Set an initial variable for holding the yearly change
 Dim Yearly_Change As Double
 
 ' Set an initial variable for holding the percentage change
 Dim Percentage_Change As Double
 
 
 ' Keep track of the location for each ticker name in the summary table
 Dim Summary_Table_Row As Integer
 Summary_Table_Row = 2
 
 'Giving headers to new created columns
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("I1").Value = "Ticker"
 ws.Range("L1").Value = "Total Stock Volume"
 
 'counts the number of the row/finding the last row
 LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 ' Keep track of the location for each opening price for the new ticker
 Dim j As Long
 j = 2
 
 'Loop through all the rows(tickers)
 For i = 2 To LastRow
 
 
     'check if we are still in same ticker or not
     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
 
     'set the ticker name
     Ticker_Name = ws.Cells(i, 1).Value
     
     'determinethe "Yearly Change"
     Yearly_Change = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
     
     'print the yearly change in the summary table
     ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
     
     'applying the conditiontive
     'Set the Cells Colours to Red and green
     If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
         ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
    
    'determinethe "Percentage Change"
     Percentage_Change = Yearly_Change / ws.Cells(j, 3).Value
     
     'print the percentage change in the summary table
     ws.Range("K" & Summary_Table_Row).Value = Percentage_Change
     
     'Format the values in the Percentage Change to %
     ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
    
 
     'add to the total valume
     Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
 
     'print the ticker name in the summary table
     ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
     
     'print the total volume in the summary table
     ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
      
     'Increment the summary table row
     Summary_Table_Row = Summary_Table_Row + 1
     
     'Increment the j
     j = i + 1
     
     'reset the total volume for the next ticker name
     Total_Stock_Volume = 0
     
  'If the following cell is the same ticker
   Else
   
       'add to the total volume
       Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
    
      
      End If
    Next i
  Next ws
End Sub