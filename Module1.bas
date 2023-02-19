Attribute VB_Name = "Module1"
Option Explicit
Sub openForm()
populatecomboboxes
showDate
MsgBox "Today's date is: " & converterForm.dateText.Text
converterForm.Show
End Sub
Sub converterQuery()
Dim url As String, i As Integer, indexToConv As Integer, indexConv As Integer
'date handling
'Dim dayDate As String, monthDate As String, yearDate As String, slashone As Integer, slashtwo As Integer
'slashone = InStr(converterForm.dateText.Text, "/")
'slashtwo = InStr(slashone + 1, converterForm.dateText.Text, "/")
'monthDate = Left(converterForm.dateText.Text, slashone - 1)
'dayDate = Mid(converterForm.dateText.Text, slashone + 1, 2)
'yearDate = Right(converterForm.dateText.Text, Len(converterForm.dateText.Text) - slashtwo)
'extracts month from the MM/DD/YYYY format
'If Len(monthDate) = 1 Then monthDate = "0" & monthDate
'If Len(dayDate) = 1 Then dayDate = "0" & dayDate
'bring up the chart in query import
Application.ScreenUpdating = False
Sheets("Sheet1").Visible = True
Sheets("Sheet1").Cells.Clear
'url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & yearDate & "-" & monthDate & "-" & dayDate   one way
    url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & Format(converterForm.dateText.Text, "yyyy-mm-dd")
    With Worksheets("Sheet1").QueryTables.Add(Connection:=url, Destination:=Worksheets("Sheet1").Range("A1"))
        .Name = "My Query"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
'look for  cell with value="USD"
For i = 1 To 10000
   If Worksheets("Sheet1").Range("A1:A" & 10000).Cells(i, 1).Text = "USD" Then
    Sheets("Sheet1").Select
    Range("A" & i).Select
    Exit For
    End If
    Next i
'new selection must be cell with USD , starting from this cell search selected index and convert
indexToConv = converterForm.ComboBox1.ListIndex
indexConv = converterForm.ComboBox2.ListIndex
If ActiveCell.Offset(indexToConv, 2).Value = 0 Or ActiveCell.Offset(indexConv, 2).Value = 0 Then
MsgBox "ONE OF THE CURRENCY RATES IS NOT AVAILABLE FOR THAT DATE, TRY A DIFFERENT YEAR "
Sheets("Sheet1").Visible = False
Exit Sub
End If
converterForm.convertedOutput = converterForm.TextBox1 * (ActiveCell.Offset(indexConv, 2).Value / ActiveCell.Offset(indexToConv, 2).Value)
Sheets("Sheet1").Visible = False
End Sub

Sub populatecomboboxes()
Dim i As Integer
Sheets("Sheet2").Visible = True
Sheets("Sheet2").Select
Range("A1").Select
Application.ScreenUpdating = False
For i = 1 To WorksheetFunction.CountA(Worksheets("Sheet2").Columns("A:A"))
    
    converterForm.ComboBox1.AddItem ActiveCell.Offset(i - 1, 0) & "-" & ActiveCell.Offset(i - 1, 1)
    converterForm.ComboBox2.AddItem ActiveCell.Offset(i - 1, 0) & "-" & ActiveCell.Offset(i - 1, 1)
Next i
converterForm.ComboBox1.Text = Worksheets("Sheet2").Range("A1") & "-" & Worksheets("Sheet2").Range("B1")
converterForm.ComboBox2.Text = Worksheets("Sheet2").Range("A2") & "-" & Worksheets("Sheet2").Range("B2")
Sheets("Sheet2").Visible = False
Sheets("ConverterSheet").Select
Range("A1").Select
End Sub

Sub showDate()
Dim datesVector
datesVector = Split(Now())
converterForm.dateText.Text = datesVector(0)

End Sub

Sub plotlastthirtydays()
Application.ScreenUpdating = False
Dim i As Integer, url As String, j As Integer, indexToConv As Integer, indexConv As Integer
Sheets("Sheet3").Visible = True
Sheets("Sheet3").Select

For i = 30 To 1 Step -1

    Range("A" & 30 - i + 1) = DateAdd("d", -i + 1, converterForm.dateText.Text)
    Next i
'populate cells B with currency
For i = 1 To 30

Sheets("Sheet1").Visible = True
Sheets("Sheet1").Cells.Clear
'url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & yearDate & "-" & monthDate & "-" & dayDate   one way  WHICH WONT WORK
    url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & Format(Worksheets("Sheet3").Range("A" & i).Text, "yyyy-mm-dd")
    With Worksheets("Sheet1").QueryTables.Add(Connection:=url, Destination:=Worksheets("Sheet1").Range("A1"))
        .Name = "My Query"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
'look for  cell with value="USD"
For j = 1 To 1000
   If Worksheets("Sheet1").Range("A1:A" & 1000).Cells(j, 1).Text = "USD" Then
    Sheets("Sheet1").Select
    Range("A" & j).Select
    Exit For
    End If
Next j
'new selection must be cell with USD , starting from this cell search selected index and convert
indexToConv = converterForm.ComboBox1.ListIndex
indexConv = converterForm.ComboBox2.ListIndex
Worksheets("Sheet3").Range("B" & i) = converterForm.TextBox1 * (ActiveCell.Offset(indexConv, 2).Value / ActiveCell.Offset(indexToConv, 2).Value)
Next i
Sheets("Sheet1").Visible = False
Sheets("Sheet3").Visible = False
MsgBox "Close and check the plot sheet"
End Sub

Sub refreshCombobox()
Dim i As Integer, k As Integer, url As String
Sheets("Sheet1").Visible = True
Sheets("Sheet1").Cells.Clear
For i = converterForm.ComboBox1.ListCount - 1 To 0 Step -1
        converterForm.ComboBox1.RemoveItem i
Next i
For i = converterForm.ComboBox2.ListCount - 1 To 0 Step -1
        converterForm.ComboBox2.RemoveItem i
Next i
'url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & yearDate & "-" & monthDate & "-" & dayDate   one way
    url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & Format(converterForm.dateText.Text, "yyyy-mm-dd")
    With Worksheets("Sheet1").QueryTables.Add(Connection:=url, Destination:=Worksheets("Sheet1").Range("A1"))
        .Name = "My Query"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
'look for  cell with value="USD"
For i = 1 To 10000
   If Worksheets("Sheet1").Range("A1:A" & 10000).Cells(i, 1).Text = "USD" Then
    Sheets("Sheet1").Select
    Range("A" & i).Select
    Exit For
    End If
    Next i
Do
    If Len(ActiveCell.Offset(k, 0).Text) <> 3 Then Exit Do
  converterForm.ComboBox1.AddItem ActiveCell.Offset(k, 0) & "-" & ActiveCell.Offset(k, 1)
    converterForm.ComboBox2.AddItem ActiveCell.Offset(k, 0) & "-" & ActiveCell.Offset(k, 1)
    k = k + 1
    
Loop
Sheets("Sheet1").Visible = False
Sheets("ConverterSheet").Select
End Sub


