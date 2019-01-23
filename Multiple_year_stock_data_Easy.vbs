Sub Multiple()

' Set an initial variable for holding ticker name
Dim ticker As String

'find last row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'set an initial variable for holding toal by ticker
Dim stock_volume As Double
stock_volume = 0

'track the location of each ticker in the summary table
Dim summary_table_row As Integer
summary_table_row = 2

'loop through all stock volume
For i = 2 To LastRow

' move through ticker ids for changes
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'set the ticker id
ticker = Cells(i, 1).Value

'add to stock volume
stock_volume = stock_volume + Cells(i, 7).Value

'update summary table with ticker id
Range("I" & summary_table_row).Value = ticker

'update summary table with stock volume
Range("J" & summary_table_row).Value = stock_volume

'Add one to the summary table row
summary_table_row = summary_table_row + 1

'reset the stock volume
stock_volume = 0

'if the cell immediately following is the same ticker id...
Else

'add to the stock volume
stock_volume = stock_volume + Cells(i, 7).Value

End If

Next i

End Sub

