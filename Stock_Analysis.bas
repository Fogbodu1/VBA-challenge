Attribute VB_Name = "Module1"
Sub Stock_Analysis()

 ' Create a loop to cycle through the worksheets in the workbook
 ' Set a variable to cycle through the worksheets
    Dim ws As Worksheet

    'Start loop
    For Each ws In Worksheets
     
    
  ' Set an initial variable for holding the Ticker and summary table

    Dim Ticker As String
    Dim Summary_Table_Row As Integer
    Dim TotalStockVolume As Double
    Dim startPrice As Double
    Dim closePrice As Double
    Dim YearlyChange As Double
    Dim percentChange As Double
    percentChange = 0
    TotalStockVolume = 0
 
  ' Set the range for holding the for the Ticker and Total Stock Volume

   ws.Range("I1") = "Ticker"
   ws.Range("J1") = "Yearly Change"
   ws.Range("K1") = "Percent Change"
   ws.Range("L1") = "Total Stock Volume"

  ' Keep track of the location for each Ticker in the summary table
  Summary_Table_Row = 2
  j = 0

  ' Loop through all stock options
  For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Check if we are still within the same ticker, if it is not...
      If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
    
        startPrice = ws.Cells(i, 3).Value
          
          End If

 ' Set the close price
      If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
          
          closePrice = ws.Cells(i, 6).Value
      
      Ticker = ws.Cells(i, 1).Value

 ' Add to the Stock Volume Total
      TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

 ' Print Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker


 ' Print the Total Stock Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume
      
' Calculate the price change for the year and move it to the summary table.
      YearlyChange = closePrice - startPrice
      ws.Range("J" & Summary_Table_Row).Value = YearlyChange
      Select Case YearlyChange
            Case Is > 0
            ws.Range("J" & j + 2).Interior.ColorIndex = 4 'Green
            Case Is < 0
            ws.Range("J" & j + 2).Interior.ColorIndex = 3 'Red
      End Select
      j = j + 1
      
' Calculate the percent change for the year and move it to the summary table format as a percentage
' Conditional for calculating percent change
      If startPrice = 0 Then
      percentChange = YearlyChange / 1
      ElseIf startPrice <> closePrice Then
      percentChange = YearlyChange / startPrice
      Else
      percentChange = 0
      End If
        
      ws.Range("K" & Summary_Table_Row).Value = percentChange
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
 ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
 ' Reset the Stock Volume Total, year open price, year close price, year change, year percent change
      TotalStockVolume = 0
      startPrice = 0
      closePrice = 0
      YearlyChange = 0
      Percent_Change = 0

 ' If the cell immediately following a row is the same ticker...
      Else

 ' Add to the Stock Volume Total
      TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

    End If

    
  Next i
  
        'Create a best and worst performance table/titles
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        'Set variables to hold best performer, worst performer, and stock with the most volume
        Dim best_stock As String
        Dim best_value As Double
        Dim worst_stock As String
        Dim worst_value As Double
        Dim most_vol_stock As String
        Dim most_vol_value As Double
        
        'Set best performer equal to the first stock
         best_value = ws.Range("K2")
         'best_value = ws.Cells(2, 11).Value

        'Set worst performer equal to the first stock
        worst_value = ws.Range("K2")
        'worst_value = ws.Cells(2, 11).Value

        'Set most volume equal to the first stock
        most_vol_value = ws.Range("L2")
        'most_vol_value = ws.Cells(2, 12).Value
            
        
        
        'Loop to search through summary table
        For j = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
 
        'Conditional to determine best performer
            If ws.Cells(j, 11).Value > best_value Then
                best_value = ws.Cells(j, 11).Value
                best_stock = ws.Cells(j, 9).Value
            End If
            
        'Conditional to determine worst performer
            If ws.Cells(j, 11).Value < worst_value Then
                worst_value = ws.Cells(j, 11).Value
                worst_stock = ws.Cells(j, 9).Value
            End If
            
        'Conditional to determine stock with the greatest volume traded
            If ws.Cells(j, 12).Value > most_vol_value Then
                most_vol_value = ws.Cells(j, 12).Value
                most_vol_stock = ws.Cells(j, 9).Value
            End If
  Next j
  
        'Move best performer, worst performer, and stock with the most volume items to the performance table
         'ws.Range("P2") = best_stock
         'ws.Range("Q2") = best_value
         'ws.Range("Q2") = "0.00%"

        ws.Cells(2, 16).Value = best_stock
        ws.Cells(2, 17).Value = best_value
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 16).Value = worst_stock
        ws.Cells(3, 17).Value = worst_value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 16).Value = most_vol_stock
        ws.Cells(4, 17).Value = most_vol_value
        
        best_value = 0
        worst_value = 0
        most_vol_value = 0

  
  ' Analysis complete
  
  Next ws
  
  MsgBox ("Analysis Complete")

End Sub

