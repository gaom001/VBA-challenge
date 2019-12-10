Sub stockloop()

' declare varibles'
Dim ticker_name As String
Dim open_price As Double
Dim high_price As Double
Dim low_price As Double
Dim close_price As Double
Dim stock_volume As Double


'establish a new sumamry table starting from I1 to L1
Dim summary_table_row As Integer
summary_table_row = 2
Cells(summary_table_row - 1, 9).Value = "ticker_symbol"
Cells(summary_table_row - 1, 10).Value = "yearly change"
Cells(summary_table_row - 1, 11).Value = "percent change"
Cells(summary_table_row - 1, 12).Value = "total_stock_volume"

'Set an initial variable for total stock volume and yearly change
Dim total_stock_volume As Double
total_stock_volume = 0

Dim yearly_change As Double


'Loop through all ticker names
  'identify the last row of the orignal table, then loop ticker namse from A2 to lastrow
  lastrow = Cells(Rows.Count,1).End(xlUp).Row

'The first open price in first ticker
open_price = Cells(2, 3).Value

For i = 2 To lastrow

    ' Check if we are still within the same ticker name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker_symbol = Cells(i, 1).Value
        
        close_price = Cells(i, 6).Value

        yearly_change = close_price - open_price

     'Check Division by 0 condition, and calculate percent change
        If open_price <> 0 Then
        percent_change = yearly_change / open_price * 100
        End If       
        
     'first open price in next ticker (replace the value of first ticker open price ex.(2,3) by next ticker open price ex.(264,3))
        open_price = Cells(i + 1, 3).Value

     'input the total stock volume
        total_stock_volume = total_stock_volume + Cells(i, 7).Value

        'Print the ticker symbol, yearly change, percent change and total stock volume into the new summary table
        Range("I" & summary_table_row).Value = ticker_symbol
        Range("J" & summary_table_row).Value = yearly_change
            If (yearly_change > 0) Then
                'Fill column with "GREEN" 
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            ElseIf (yearly_change<= 0) Then
                'Fill column with "RED"  
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
        Range("K" & summary_table_row).Value = percent_change & "%"
        Range("L" & summary_table_row).Value = total_stock_volume

        'Reset the total stock volume when executing next i
        total_stock_volume = 0
        close_price = 0
        
        'Print values into the next row when executing next i
        summary_table_row = summary_table_row + 1
        
    ' If the cell immediately following a row is the same ticker
    Else
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
          
 
    End If  
Next i




'ddeclare new varibles for the new summary table (column I : column L)
Dim max_percent As Double
Dim min_percent As Double
Dim max_volume As Double
Dim max_ticker as string
Dim max_volume_ticker as string

'set initial value for Greatest % increase, Greatest % Decrease and Greatest total volume
max_percent = 0
min_percent=0
max_volume=0

'identify the last row of the new summary table
Lrow=Cells(Rows.Count, 9).End(xlUp).Row
'Loop through column K values(percent change) and column L(total stock volume) values, 
For j = 2 To Lrow
    'loop to get the max value of percent change
    If Cells(j, 11).Value > max_percent Then
    max_percent = Cells(j, 11).Value
    max_ticker=cells(j,9).value
    End If

    'loop to get the min value of percent change
    If cells(j,11).value< min_percent Then
    min_percent=cells(j,11).value
    min_ticker=cells(j,9).value
    End if
   
   'loop to get the max value of total stock volume
    If Cells(j, 12).Value > max_volume Then
    max_volume = Cells(j, 12).Value
    max_volume_ticker=cells(j,9).value
    End If


'establish another new table starting from N1 to P1
    Range("O2").Value = "Greatest % increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest total volume"
    Range("P1").Value = "ticker name"
    Range("Q1").Value = "value"

'Print the Greatest % increase, Greatest % decrease and Greatest total volume with their tickers
    Cells(2, 17).Value = max_percent
    cells(2,16).value=max_ticker
    Cells(3, 17).Value = min_percent
    cells(3,16).value=min_ticker
    Cells(4, 17).Value = max_volume
    cells(4,16).value=max_volume_ticker

Next j



End Sub
