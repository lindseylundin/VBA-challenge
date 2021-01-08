Attribute VB_Name = "Module1"
Sub vba_challenge()

' Loop through all worksheets
For Each ws In Worksheets

' Set an initial variable for holding the stock ticker symbol
Dim ticker_symbol As String

' Set an initial variable for holding the total per ticker symbol (stock)
Dim total_volume As Double
total_volume = 0

' Keep track of the location for each ticker symbol in the summary table
Dim summary_table_row As Integer
summary_table_row = 2

' To get last cell with data
Dim last_cell As Long
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' To add headers to the new columns
ws.Cells(1, 9).Value = "Ticker Symbol"
ws.Cells(1, 10).Value = "Total Stock Volume"


' Loop through all stock entries
    For i = 2 To LastRow

' Check if we are still within the same ticker symbol, if it is not...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

 ' Set the Ticker Symbol
      ticker_symbol = ws.Cells(i, 1).Value

      ' Add to the Total Volume
      total_volume = total_volume + ws.Cells(i, 7).Value

      ' Print the Ticker Symbol in the Summary Table
      ws.Range("I" & summary_table_row).Value = ticker_symbol

      ' Print the Total Volume to the Summary Table
      ws.Range("J" & summary_table_row).Value = total_volume

      ' Add one to the summary table row
      summary_table_row = summary_table_row + 1
      
      ' Reset the Total Volume
      total_volume = 0
    

    ' If the cell immediately following a row is the same Ticker Symbol...
    Else

      ' Add to the Total Stock Volume
      total_volume = total_volume + ws.Cells(i, 7).Value

         End If
         
    Next i
    
    
Next ws
    
End Sub

