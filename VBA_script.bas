Attribute VB_Name = "Module1"
Sub VBA_Challenge_Wall_Street():

'Set variable for holding the ticker name and count
Dim ticker_name As String
Dim tickercount As Long

'Set variables for holding the ticker opening price, closing price, total volume, yearly change, and percent change
Dim opening_price As Double
Dim closing_price As Double
Dim total_volume As Double
Dim yearly_change As Double
Dim percent_change As Double

'Set variable for last row
Dim lastrow As Long

'Set variable for summary table
Dim summary_table_row As Long

'Set variable for worksheet
Dim Ws As Integer

'Loop through each worksheet
For Ws = 1 To Sheets.Count
    
    'Activate worksheet
    Sheets(Ws).Activate
    
    'Determine the last row
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
 
    'Keep track of the row for each ticker in the summary table
    summary_table_row = 2

    'Print headings for the summary table
    Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")

    'Set the tickercount to be 0
    tickercount = 0

    'Set the total stock volume to be 0
    total_volume = 0

    'Loop through each row in the column
    For i = 2 To lastrow
    
        'Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the ticker name
            ticker_name = Cells(i, 1).Value
            
            'Print the ticker name in the summary table
            Range("I" & summary_table_row).Value = ticker_name
            
            'Set the closing price for the ticker
            closing_price = Cells(i, 6).Value
                                
            'Set the opening price for the ticker
            opening_price = Cells(i - tickercount, 3).Value
                   
            'Add to the stock total volume for the ticker
            total_volume = total_volume + Cells(i, 7).Value
                   
            'Print the stock total volume for the ticker in the summary table
            Range("L" & summary_table_row).Value = total_volume
            
            'Calculate yearly change
            yearly_change = closing_price - opening_price
    
            'Print yearly change
            Range("J" & summary_table_row).Value = yearly_change
       
            'Calculate percent change
            If opening_price = 0 Then
            percent_change = 0
            Else
            percent_change = yearly_change / opening_price
            End If
            
            'Print percent change
            Range("K" & summary_table_row).Value = percent_change
            Range("K" & summary_table_row).NumberFormat = "0.00%"
        
            'Conditional formatting that will highlight positive change in green and negative change in red
            If yearly_change >= 0 Then
            Range("J" & summary_table_row).Interior.ColorIndex = 4
            Else
            Range("J" & summary_table_row).Interior.ColorIndex = 3
            End If
    
            'Add one to the summary table row
            summary_table_row = summary_table_row + 1
            
            'Reset the ticker total stock volume
            total_volume = 0
                                          
            'Reset the ticker count
            tickercount = 0
            
        Else
        'Add 1 to tickercount
        tickercount = tickercount + 1
        
        'Add the volume to the total stock volume
        total_volume = total_volume + Cells(i, 7).Value
     
        End If
  
    Next i


    'Determine the last row of the summary table results
    summary_lastrow = Cells(Rows.Count, "I").End(xlUp).Row

    'Print headings for the max, min values table
    Range("O1:P1").Value = Array("Ticker", "Value")
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4") = "Greatest Total Volume"

    'Loop through rows to find values
    For i = 2 To summary_lastrow
  
        'Value in percent change column equals max value
        If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & summary_lastrow)) Then
                
        'Print the max value and ticker name
        Cells(2, 16).Value = Cells(i, 11).Value
        Cells(2, 16).NumberFormat = "0.00%"
        Cells(2, 15).Value = Cells(i, 9).Value
   
        'Value in percent change column equals min value
        ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & summary_lastrow)) Then
                
        'Print the min value and ticker name
        Cells(3, 16).Value = Cells(i, 11).Value
        Cells(3, 16).NumberFormat = "0.00%"
        Cells(3, 15).Value = Cells(i, 9).Value
                
        'Value in total stock volume equals max value
        ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & summary_lastrow)) Then
                
        'Print the max value and ticker name
        Cells(4, 16).Value = Cells(i, 12).Value
        Cells(4, 15).Value = Cells(i, 9).Value

        End If
            
    Next i
                     
    ' Autofit to display data
    Columns("I:P").AutoFit

Next Ws

End Sub



