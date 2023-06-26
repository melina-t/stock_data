VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks():

'Looping through all the sheets
For Each ws In Worksheets
'------------------------------------------------------------------------
Dim ticker As String
Dim yearly_change As Double
Dim open_price As Double
Dim close_price As Double
Dim percent_change As Double
Dim total_volume As Double


'To determine the last row, taken from the Census Pt.1 script seen in class
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'This part was taken from the Credit Card VBA script seen in class
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'Adding a column with a name to all the WS's,taken from Census Pt.1
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

    open_price = ws.Cells(2, 3).Value

    For i = 2 To LastRow

    total_volume = total_volume + ws.Cells(i, 7).Value
    

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        ws.Cells(Summary_Table_Row, 9).Value = ticker
        
        yearly_change = (ws.Cells(i, 6).Value - open_price)
        ws.Cells(Summary_Table_Row, 10).Value = yearly_change
    
        percent_change = (yearly_change / open_price)
        ws.Cells(Summary_Table_Row, 11).Value = percent_change
        ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
        
        ws.Cells(Summary_Table_Row, 12).Value = total_volume
'Resetting the values
        total_volume = 0
        yearly_change = 0
        open_price = ws.Cells(i + 1, 3).Value
               
Summary_Table_Row = Summary_Table_Row + 1
   
    End If
    
    If ws.Cells(Summary_Table_Row, 10).Value >= 0 Then
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
     
    End If
        
    If ws.Cells(Summary_Table_Row, 10).Value < 0 Then
    ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
     
    End If

Next i


'Second part of activity, cells (2,3001)
Dim great_increase As Double
great_increase = 0
Dim great_decrease As Double
great_decrease = 0
Dim great_total As Double
great_total = 0

NextLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

For i = 2 To NextLastRow
    If great_increase < ws.Cells(i, 11).Value Then
    great_increase = ws.Cells(i, 11).Value
    ws.Cells(2, 16).Value = great_increase
    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"
    End If
    
    If great_decrease > ws.Cells(i, 11).Value Then
    great_decrease = ws.Cells(i, 11).Value
    ws.Cells(3, 16).Value = great_decrease
    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
    ws.Cells(3, 17).NumberFormat = "0.00%"
    End If
    
    If great_total < ws.Cells(i, 12) Then
    great_total = ws.Cells(i, 12).Value
    ws.Cells(4, 16).Value = great_total
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
    End If
    

Next i


Next ws
End Sub
