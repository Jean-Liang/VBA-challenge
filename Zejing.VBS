Attribute VB_Name = "Module1"
Sub StockData()

'Loop Through all sheets
For Each ws In Worksheets
Dim WorksheetName As String

'Determin the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
WorksheetName = ws.Name

'Insert column I for Ticker
ws.Range("I1").EntireColumn.Insert
'Name column I header as Ticker
ws.Cells(1, 9).Value = "Ticker"

'Insert column J for Yearly Change
ws.Range("J1").EntireColumn.Insert
'Name column I header as Yearly Change
ws.Cells(1, 10).Value = "Yearly Change"

'Insert column K for Percent Change
ws.Range("K1").EntireColumn.Insert
'Name column K header as Percent Change
ws.Cells(1, 11).Value = "Percent Change"


'Insert column L for Total Stock Volume
ws.Range("L1").EntireColumn.Insert
'Name column L header as Total Stock Volume
ws.Cells(1, 12).Value = "Total Stock Volume"


'Insert column O,P,Q for the bouns table
ws.Range("O1").EntireColumn.Insert
ws.Cells(2, 15).Value = "Greatest%Increase"
ws.Cells(3, 15).Value = "Greatest%Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

ws.Range("P1").EntireColumn.Insert
ws.Cells(1, 16).Value = "Ticker"
ws.Range("Q1").EntireColumn.Insert
ws.Cells(1, 17).Value = "Value"


'Below try to find the different ticker name in each worksheet
'first define the variable of ticker
Dim ticker As String

'Define variable of Total stock volume
Dim Stockvolume_Total As Double
Stockvolume_Total = 0

'Define variable of Yearly Change
Dim open_amount As Double
Dim close_amount As Double
Dim Yearly_Change As Double

'Define variable of percent change
Dim Percent_Change As Double


Dim Summary_table_row As Integer
Summary_table_row = 2

'Loop through all the rows
For i = 2 To LastRow

'Check if the first row of the ticker group name
If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
'To get the first row of the new ticker group, so we can have the open amount
open_amount = ws.Cells(i, 3).Value
End If

'Check if next ticker name is differnt than the current ticker name
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Set the ticker name
ticker = ws.Cells(i, 1).Value

'Add to the Stock Volume total
Stockvolume_Total = Stockvolume_Total + ws.Cells(i, 7).Value

'Get open and close amount value to calculate the yearly change
close_amount = ws.Cells(i, 6).Value
Yearly_Change = close_amount - open_amount

'To calculate the Percent change,we need consider if the open_amount is zero
'Use If, we set Percent_Change=0 when the open_amount is zero
If open_amount <> 0 Then
Percent_Change = Yearly_Change / open_amount
Else
Percent_Change = 0
End If


'Print the ticker name in the Summary table
ws.Range("I" & Summary_table_row).Value = ticker


'Print the yearly change in the summary table
ws.Range("J" & Summary_table_row).Value = Yearly_Change
ws.Range("J" & Summary_table_row).Style = "Currency"
If Yearly_Change > 0 Then
ws.Range("J" & Summary_table_row).Interior.ColorIndex = 4
ElseIf Yearly_Change < 0 Then
ws.Range("J" & Summary_table_row).Interior.ColorIndex = 3

End If


'Print the percent change in the summary table
'The reference format cell get from link https://learn.microsoft.com/en-us/office/vba/api/excel.cellformat.numberformat
ws.Range("K" & Summary_table_row).Value = Percent_Change
ws.Range("K" & Summary_table_row).NumberFormat = "0.00%"

If Percent_Change > 0 Then
ws.Range("K" & Summary_table_row).Interior.ColorIndex = 4
ElseIf Percent_Change < 0 Then
ws.Range("K" & Summary_table_row).Interior.ColorIndex = 3

End If

'Print total stock volume in the Summary table
ws.Range("L" & Summary_table_row).Value = Stockvolume_Total


'Add one to the summart table row
Summary_table_row = Summary_table_row + 1

Stockvolume_Total = 0

Else

'Add to the Total Stock Volume
Stockvolume_Total = Stockvolume_Total + ws.Cells(i, 7).Value


End If

Next i

Dim Greatincrease As Double
Dim Greatdecrease As Double
Dim GreatTotal As Double
'The worksheetfunction refernce link from https://www.wallstreetmojo.com/vba-max/
Greatincrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
Greatdecrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
GreatTotal = Application.WorksheetFunction.Max(ws.Range("L:L"))

ws.Cells(2, 17).Value = Greatincrease
ws.Cells(3, 17).Value = Greatdecrease
ws.Cells(4, 17).Value = GreatTotal

ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"


'Below is how to find the each ticker
For i = 2 To LastRow

If ws.Cells(i, 11).Value = Greatincrease Then
ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

ElseIf ws.Cells(i, 11).Value = Greatdecrease Then
ws.Cells(3, 16).Value = ws.Cells(i, 9).Value

ElseIf ws.Cells(i, 12).Value = GreatTotal Then
ws.Cells(4, 16).Value = ws.Cells(i, 9).Value



End If
Next i

Next ws


MsgBox ("Data update is completed")


End Sub

