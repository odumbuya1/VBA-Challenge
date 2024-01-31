Attribute VB_Name = "Module1"
Sub StockMarket1()

'Loop through all the sheets.
    For Each ws In Worksheets

' Declare variables
Dim ticker As String

'Set a variable to hold total count stock volume
Dim stock_volume As Double
stock_volume = 0

'Track summary table and 'Initialize summary table row
Dim summaryTableRow As Integer
summaryTableRow = 2

Dim opening_price As Double
' set initial open_price. Other opening prices will be determined in the conditional loop
opening_price = Cells(2, 3).Value

Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'Insert summary table headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


'Dim I As Long
'Set Worksheet
'Dim ws As Worksheet
'ws = ThisWorkbook.Sheets(2018)

Dim lastRow As Long
'Determine the last row with data in colume A
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'loop through rows
For I = 2 To lastRow

    ' Check if the next ticker symbol is different
   If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
   
        'Set ticker symbol
        ticker = ws.Cells(I, 1).Value
        
        'Add stock volume for current ticker
        stock_volume = stock_volume + ws.Cells(I, 7).Value
                 
        'Populate summary table
        ws.Range("I" & summaryTableRow).Value = ticker
        ws.Range("L" & summaryTableRow).Value = stock_volume
        
        'Get closing prices
         closing_price = ws.Cells(I, 6).Value
        
        'Calculate yearly change
        yearly_change = (closing_price - opening_price)
        
        'Print yearly change for each ticker in summary table
        ws.Range("J" & summaryTableRow).Value = yearly_change
        
                
        ' Calculate percent change
        If opening_price = 0 Then
            percent_change = 0
            
        Else
            percent_change = yearly_change / opening_price
       
        End If
                      
        'Print Percent change for each ticker in summary table
        ws.Range("K" & summaryTableRow).Value = percent_change
        ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
        
        'Reset summary table row counter
        summaryTableRow = summaryTableRow + 1
        
        'Reset stock volume to zero for the next ticker
        stock_volume = 0
        
        'Reset the opening price
         opening_price = Cells(I + 1, 3)
                    
    Else
    
      'Add stock volume for current ticker
      stock_volume = stock_volume + ws.Cells(I, 7).Value
      
      
    End If
    
Next I

'Use conditional formatting to show positive changes in green and neagtive changes in red.
'Determine last row of summary table

lastRow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Format yearly change column in summary table into green (if positive) or red (if negative)
    For I = 2 To lastRow_summary_table
    
        If ws.Cells(I, 10).Value > 0 Then
            ws.Cells(I, 10).Interior.ColorIndex = 4
            
        Else
             ws.Cells(I, 10).Interior.ColorIndex = 3
            
        End If
        
    Next I
     
  'Insert a sub table to hold the following  values
 ws.Cells(2, 15).Value = "Greatest % Increase"
 ws.Cells(3, 15).Value = "Greatest % Decrease"
 ws.Cells(4, 15).Value = "Greatest Total Volume"
 ws.Cells(1, 16).Value = "Ticker"
 ws.Cells(1, 17).Value = "Value"
 
 'Determine the maximum and minimum percent change and maximum total stock volume, ticker symbol
 
 For I = 2 To lastRow_summary_table
 
    'Get maximum percent change
    If ws.Cells(I, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastRow_summary_table)) Then
        ws.Cells(2, 16).Value = ws.Cells(I, 9).Value
        ws.Cells(2, 17).Value = ws.Cells(I, 11).Value
        ws.Cells(2, 17).NumberFormat = "0.00%"
 
 
       'Get minimum percent change
    ElseIf ws.Cells(I, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastRow_summary_table)) Then
        ws.Cells(3, 16).Value = ws.Cells(I, 9).Value
        ws.Cells(3, 17).Value = ws.Cells(I, 11).Value
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        
        'Get maximum stock volume
    ElseIf ws.Cells(I, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastRow_summary_table)) Then
        ws.Cells(4, 16).Value = ws.Cells(I, 9).Value
        ws.Cells(4, 17).Value = ws.Cells(I, 12).Value
       
       End If
       
     Next I
     
    Next ws
        
    
End Sub


