Sub ticker()
    ' Statement so code runs on all sheets within workbook
    Dim ws As Worksheet
    For Each ws In Worksheets
    ws.Activate
    
    ' Set initial variable for holding the ticker name
    Dim ticker_name As String

    ' Set an initial variable for holding the total stock volume for ticker
    Dim ticker_total As Double
    ticker_total = 0
    
    ' Setting variable for holding yearly stick price change
    Dim yearly_change As Double
    
    ' Set initial variable for holding initial stock price
    Dim begin_yearly_change As Double
    
    ' Set variable for holding ending stock price
    Dim end_yearly_change As Double
    
    ' Keep track of the location for each ticker name in the proper column
    Dim summary_ticker_row As Integer
    summary_ticker_row = 2

    ' Loop through all ticker names
    For i = 2 To 70925

        ' Check if we are still within the same ticker name, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the ticker name and titles
            ticker_name = Cells(i, 1).Value
            Cells(1, 8).Value = "<ticker symbol>"
            Cells(1, 9).Value = "<yearly change>"
            Cells(1, 10).Value = "<percent change>"
            Cells(1, 11).Value = "<total volume>"
            
            ' Set ticker value at open of first day
            begin_yearly_change = Cells(i - 260, 3).Value
            
            ' Set ticker value at close of first day
            end_yearly_change = Cells(i, 6).Value
            
            ' Add to the ticker total
            ticker_total = ticker_total + Cells(i, 7).Value
            
            ' Defining percent_change value
            yearly_change = Range("I" & summary_ticker_row).Value

            ' Print the ticker name in the proper column
            Range("H" & summary_ticker_row).Value = ticker_name

            ' Print the ticker volume amount to the proper column
            Range("K" & summary_ticker_row).Value = ticker_total
            
            ' Print the difference of end stock value and begining stock value
            Range("I" & summary_ticker_row).Value = Cells(i, 6).Value - Cells(i - 260, 3).Value
            
            ' Print percentage change in column J
            Range("J" & summary_ticker_row).Value = (yearly_change / begin_yearly_change) * 100

            ' Add one to the summary ticker column
            summary_ticker_row = summary_ticker_row + 1
      
            ' Reset totals
            ticker_total = 0
            end_yearly_change = 0
            begin_yearly_change = 0

        ' If the cell immediately following a row is the same ticker...
        Else

            ' Add to the ticker total
            ticker_total = ticker_total + Cells(i, 7).Value
            
        End If

    Next i
    
    ' Defining variables for side chart
    Dim percent_rng As Range
    Dim max_volume_rng As Range
    Dim max_function As Double
    Dim min_function As Double
    Dim max_volume As Double
    
    ' Creating labels and headers for side chart
    Cells(1, 14).Value = "Ticker"
    Cells(1, 15).Value = "Value"
    Cells(2, 13).Value = "Greatest % Increase"
    Cells(3, 13).Value = "Greatest % Decrease"
    Cells(4, 13).Value = "Greatest Total Volume"

    ' setting ranges to pull min and max values from
    Set percent_range = Range("J:J")
    Set max_volume_rng = Range("K:K")

    ' Functions that give min and max values
    max_function = Application.WorksheetFunction.Max(percent_range)
    min_function = Application.WorksheetFunction.Min(percent_range)
    max_volume = Application.WorksheetFunction.Max(max_volume_rng)

    ' Printing largest value in appropriate cell
    Cells(2, 15).Value = max_function
    Cells(3, 15).Value = min_function
    Cells(4, 15).Value = max_volume
    
  Next ws
  
End Sub


