Sub YearlyTickerData()

For Each ws in Worksheets

  ' Set an initial variable for holding the Ticker Name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total Stock Volume
  Dim Stock_Volume 
  Stock_Volume = 0

  ' Set an initial variable for holding the Ticker Open
  Dim Ticker_Open
  Ticker_Open = 0

  ' Set an initial variable for holding the Ticker Close
  Dim Ticker_Close As Double
  Ticker_Close = 0

  ' Set an initial variable for holding the Yearly Change
  Dim Yearly_Change as Double
  Yearly_Change = 0

  ' Set an initial variable for holding the Percent Change
  Dim Percent_Change
  Percent_Change = 0

' Set an initial variable for holding the Greatest Increase
  Dim Greatest_Increase
  Greatest_Increase = 0

  ' Set an initial variable for holding the Greatest Decrease
  Dim Greatest_Decrease
  Greatest_Decrease = 0

  ' Set an initial variable for holding the Greatest Volume
  Dim Greatest_Volume
  Greatest_Volume = 0

  ' Set an initial variable for holding ticker name for Greatest Values
  Dim Greatest_Name as String
  
  ' Keep track of the location for all ticker data in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Determine the Last Row
  LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row 

  ' Determine the Last Column Number
  LastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

  'set column headers and formatting 
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("N2").Value = "Greatest % Increase"
  ws.Range("N3").Value = "Greatest % Decrease"
  ws.Range("N4").Value = "Greatest Total Volume"
  ws.Range("O1").Value = "Ticker"
  ws.Range("P1").Value = "Value"
  ws.Range("J1:J" & LastRow).ColumnWidth = 15
  ws.Range("K1:K" & LastRow).ColumnWidth = 15
  ws.Range("L1:L" & LastRow).ColumnWidth = 20
  ws.Range("N1:N" & LastRow).ColumnWidth = 20
  ws.Range("P1:P" & LastRow).ColumnWidth = 20
  ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
  ws.Range("P2").NumberFormat = "0.00%"
  ws.Range("P3").NumberFormat = "0.00%"

  ' Loop through all credit card purchases
  For i = 2 To LastRow

    ' Check if we are still within the same ticker name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker Name
      Ticker_Name = ws.Cells(i, 1).Value

      ' Set the Ticker Close
      Ticker_Close = ws.Cells(i, 6).value

      ' Add to the Stock Volume
      Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

      ' Calculate Yearly Change
      Yearly_Change = Ticker_Close - Ticker_Open

      ' Calculate Percent Change
      If Ticker_Open = 0 Then
        Ticker_Open = Null
      Else 
        Percent_Change = Yearly_Change / Ticker_Open
      End if

      ' Print the ticker data in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Yearly Change in the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

      ' Print the Percent Change in the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percent_Change

      ' Print the Brand Amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Stock_Volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Volume
      Stock_Volume = 0

    Elseif ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

    ' Set the Ticker Close
      Ticker_Open = ws.Cells(i, 3).value

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Stock Volume
      Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value

    End If

  Next i

  ' Loop through Yearly Change to format cell color
  For i = 2 To LastRow
    
    ' Check if it's a negative number
    If ws.Cells(i , 10).Value < 0 Then

        'Format Red
        ws.Cells(i, 10).interior.ColorIndex = 3

    ' Check if it's a positive number
    Elseif ws.Cells(i, 10).value > 0 Then
         
         'Format Green
         ws.Cells(i, 10).interior.ColorIndex = 4

    End If
  
  Next i

  ' Loop through columns for greatest values
  For i = 2 To LastRow

    ' Greatest Increase
    If ws.Cells(i , 11).Value > Greatest_Increase Then

        'Grab Greatest Increase and Ticker
        Greatest_Increase = ws.Cells(i, 11).value
        Greatest_Name = ws.Cells(i, 9).value

        ' Place Greatest Increase 
        ws.Range("P2") = Greatest_Increase
        ws.Range("O2") = Greatest_Name

    End If

  Next i

  ' Loop through columns for greatest values
  For i = 2 To LastRow  

    ' Greatest Decrease
    If ws.Cells(i , 11).Value < Greatest_Decrease Then

        'Grab Greatest Decrease and Ticker
        Greatest_Decrease = ws.Cells(i, 11).value
        Greatest_Name = ws.Cells(i, 9).value

        ' Place Greatest Decrease 
        ws.Range("P3") = Greatest_Decrease
        ws.Range("O3") = Greatest_Name

    End If

  Next i

  ' Loop through columns for greatest values
  For i = 2 To LastRow 

    ' Greatest Volume
    If ws.Cells(i , 12).Value > Greatest_Volume Then

        'Grab Greatest Volume and Ticker
        Greatest_Volume = ws.Cells(i, 12).value
        Greatest_Name = ws.Cells(i, 9).value

        ' Place Greatest Increase 
        ws.Range("P4") = Greatest_Volume
        ws.Range("O4") = Greatest_Name

    End If
    
  Next i  

Next ws

MsgBox ("Complete")

End Sub