Attribute VB_Name = "PlotCurrencyRateBehavior"
Option Explicit

'Set currency rate date
Sub showDate()


Dim datesVector

'Obtain the date format as a string array
datesVector = Split(Now())

'Save on string display field
converterForm.dateText.Text = datesVector(0)

End Sub

'Create the plot axis
Sub plotlastthirtydays()
Application.ScreenUpdating = False
Dim i As Integer, url As String, j As Integer, indexToConv As Integer, indexConv As Integer
Sheets("Sheet3").Visible = True
Sheets("Sheet3").Select

'Iterate over the last 30 days (before selected date)
For i = 30 To 1 Step -1

'Save date per iteration and create X axis
Range("A" & 30 - i + 1) = DateAdd("d", -i + 1, converterForm.dateText.Text)
Next i

'Iterate over the last 30 days (before selected date)
For i = 1 To 30

Sheets("Sheet1").Visible = True
Sheets("Sheet1").Cells.Clear



    'Create query connection recursively using the ith index which represents each day to plot
    url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & Format(Worksheets("Sheet3").Range("A" & i).Text, "yyyy-mm-dd")
    
    With Worksheets("Sheet1").QueryTables.Add(Connection:=url, Destination:=Worksheets("Sheet1").Range("A1"))
        
        'Set query parameters format that scraps daily conversion rates and populates Y axis
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




'Fetch the data set that corresponds to the daily conversion rate (i index iteration), this will be used as the Y axis
For j = 1 To 1000
   If Worksheets("Sheet1").Range("A1:A" & 1000).Cells(j, 1).Text = "USD" Then
    Sheets("Sheet1").Select
    Range("A" & j).Select
    Exit For
    End If
Next j




'Calculate converted amount and populate Y axis in the plot

'Use combo list index logic for currency conversion
indexToConv = converterForm.ComboBox1.ListIndex
indexConv = converterForm.ComboBox2.ListIndex

'Plot each conversion record
Worksheets("Sheet3").Range("B" & i) = converterForm.TextBox1 * (ActiveCell.Offset(indexConv, 2).Value / ActiveCell.Offset(indexToConv, 2).Value)
Next i
Sheets("Sheet1").Visible = False
Sheets("Sheet3").Visible = False
MsgBox "Close and check the plot sheet"
End Sub
