Sub ticker()
'Declaring variables
Dim r_count, r_summary As Integer
Dim BoY, EoY, max_ltm, min_ltm, volume As Double
Dim ticker_id As String

'Headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percentage Charge"
Cells(1, 12).Value = "Stock Volume"
Cells(1, 13).Value = "Quote at BoY"
Cells(1, 14).Value = "Quote at EoY"
Cells(1, 15).Value = "LTM_Max"
Cells(1, 16).Value = "LTM_Min"

'Cells Formating
Range("I1:T1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True

'Percentage formating to Percent Change
   Range("K2:K" & Cells(Rows.Count, 1).End(xlUp).Row).Select
   Selection.Style = "Percent"
   Selection.NumberFormat = "0.00%"
   Range("T2:T3").Select
   Selection.Style = "Percent"
   Selection.NumberFormat = "0.00%"

'Formating volume with commas
    Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select ' a different way to scroll down selection
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    Columns("L:L").ColumnWidth = 15
    Columns("R:R").ColumnWidth = 19
    Columns("T:T").ColumnWidth = 15
    Range("T4").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
   
'Counting Rows and defining last_row value
Range("A1").Select
last_row = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox ("The table contains " & last_row & " rows")

'Loop to populate the total annual volume per ticker
r_summary = 2 'to set the initial row to populate the results
For r_count = 2 To last_row
  If Cells(r_count, 1).Value <> Cells(r_count + 1, 1).Value Then
    ticker_id = Cells(r_count, 1).Value
    volume = volume + Cells(r_count, 7).Value 'Last iteration before results are printing to capture the very last record
    Cells(r_summary, 9).Value = ticker_id
    Cells(r_summary, 12).Value = volume
    r_summary = r_summary + 1
    volume = 0 'to reset the volume for the next stock
  Else
    volume = volume + Cells(r_count, 7).Value
  End If
Next r_count

'Loop to populate the yearly change and annual yield
r_summary = 2 'to set the initial row to populate the results
For r_count = 2 To last_row
  If Cells(r_count, 1).Value <> Cells(r_count - 1, 1).Value Then
    BoY = Cells(r_count, 6).Value
  ElseIf Cells(r_count, 1).Value <> Cells(r_count + 1, 1).Value Then
    ticker_id = Cells(r_count, 1).Value
    EoY = Cells(r_count, 6).Value
    Cells(r_summary, 13).Value = BoY
    Cells(r_summary, 14).Value = EoY
    Cells(r_summary, 10).Value = EoY - BoY
    'The following lines of the script format the cells in case the stock went up or down
    If EoY > BoY Then  'Evaluates if the stock price increase
      Cells(r_summary, 10).Interior.ColorIndex = 4 'Green Color
    Else
      Cells(r_summary, 10).Interior.ColorIndex = 3 'Red Color
    End If
    'This conditional is to avoid the error of dividing a number by zero
    If BoY = 0 Then
      Cells(r_summary, 11).Value = 0
    Else
      Cells(r_summary, 11).Value = EoY / BoY - 1
    End If
    r_summary = r_summary + 1
    BoY = 0
    EoY = 0
  Else
    'none
  End If
Next r_count

'Loop to populate the maximum and minimum quotes
r_summary = 2 'to set the initial row to populate the results
max_ltm = 0
min_ltm = 10000

For r_count = 2 To last_row
  If Cells(r_count, 1).Value <> Cells(r_count + 1, 1).Value Then
    ticker_id = Cells(r_count, 1).Value
    Cells(r_summary, 9).Value = ticker_id
    
    If Cells(r_count, 6).Value > max_ltm Then 'Last iteration before results are printing to capture the very last record
      max_ltm = Cells(r_count, 6).Value
    Else
      'None
    End If
    
    If Cells(r_count, 6).Value < min_ltm Then 'Last iteration before results are printing to capture the very last record
      min_ltm = Cells(r_count, 6).Value
    Else
      'None
    End If
    
    Cells(r_summary, 15).Value = max_ltm
    Cells(r_summary, 16).Value = min_ltm
    r_summary = r_summary + 1
    max_ltm = 0
    min_ltm = 10000
  Else
    If Cells(r_count, 6).Value > max_ltm Then
      max_ltm = Cells(r_count, 6).Value
    Else
      'None
    End If
    
    If Cells(r_count, 6).Value < min_ltm Then
      min_ltm = Cells(r_count, 6).Value
    Else
      'None
    End If
   
  End If
Next r_count

'Challenges & Bonus
Cells(1, 19).Value = "Ticker"
Cells(1, 20).Value = "Value"
Cells(2, 18).Value = "Greatest % Increase"
Cells(3, 18).Value = "Greatest % Decrease"
Cells(4, 18).Value = "Greatest Total Volume"

Range("K1").Select 'Counting Rows and defining last_row value
last_row = Cells(Rows.Count, 11).End(xlUp).Row
'MsgBox ("There are " & last_row - 1 & " stocks in this table")
Cells(5, 18).Value = "Total number of stocks"
Cells(5, 19).Value = last_row - 1

max_ltm = 0 'We are recycling the use of this variable to avoid declaring more variables
min_ltm = 10000

'Loop for the Greatest % Increase
For r_count = 2 To last_row
  If Cells(r_count, 11).Value > max_ltm Then
    max_ltm = Cells(r_count, 11).Value
    ticker_id = Cells(r_count, 9).Value
  Else
    'None
  End If
Next r_count
Cells(2, 19).Value = ticker_id
Cells(2, 20).Value = max_ltm

'Loop for the Greatest % Decrease
For r_count = 2 To last_row
  If Cells(r_count, 11).Value < min_ltm Then
    min_ltm = Cells(r_count, 11).Value
    ticker_id = Cells(r_count, 9).Value
  Else
    'None
  End If
Next r_count
Cells(3, 19).Value = ticker_id
Cells(3, 20).Value = min_ltm

'Loop for the Greatest Total Volume
volume = 0 ' We are recycling this variable
For r_count = 2 To last_row
  If Cells(r_count, 12).Value > volume Then
    volume = Cells(r_count, 12).Value
    ticker_id = Cells(r_count, 9).Value
  Else
    'None
  End If
Next r_count
Cells(4, 19).Value = ticker_id
Cells(4, 20).Value = volume

'Missing Items:
' (i) Replicate buttom for all sheets

End Sub

Sub Clear()
' Macro to clear previous work

Worksheets(1).Activate 'Direct to the very first tab

'Selecting all sheets in workbook
Dim i As Long
  For i = ActiveSheet.Index To Sheets.Count
    Sheets(i).Select Replace:=False
  Next i

'Clering the space selection for all tabs
Columns("I:T").Select
Selection.ClearContents
Selection.ClearFormats
Range("A1").Select

ActiveSheet.Select Replace:=True

'Script for Clering the space selection for one tab
'    Range("I1").Select
'    Range(Selection, Selection.End(xlToRight)).Select
'    Range(Selection, Selection.End(xlDown)).Select
'    Selection.ClearContents
'    Selection.ClearFormats
'    Range("A1").Select
    
End Sub

Sub ticker_all()
' Subroutine to replicate the ticker subroute to all tabs

Worksheets(1).Activate 'Direct to the second tab

'Selecting all sheets in workbook
Dim i As Long
  For i = ActiveSheet.Index To Sheets.Count
    Worksheets(i).Activate
    ticker
  Next i

End Sub
