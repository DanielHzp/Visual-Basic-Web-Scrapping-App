Attribute VB_Name = "ConvertCurrency"
Option Explicit


'Executes the currency conversion using the rates extracted and the input fields on the form
Sub converterQuery()

'Declare auxiliary variables
Dim url As String, i As Integer, indexToConv As Integer, indexConv As Integer


'Handle spreadsheet visibility behavior when this method is executed
Application.ScreenUpdating = False
Sheets("Sheet1").Visible = True


'Clean data to avoid OVERWRITING
Sheets("Sheet1").Cells.Clear





'Create a URL that connects the query to the currency rates website www.XE.com (currencies API)

'url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & yearDate & "-" & monthDate & "-" & dayDate   one way
    url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & Format(converterForm.dateText.Text, "yyyy-mm-dd")
    
    
    
    
    
'Create query connection to extract conversion rates real time data
    With Worksheets("Sheet1").QueryTables.Add(Connection:=url, Destination:=Worksheets("Sheet1").Range("A1"))
        
        'Configure connection with the parameters used in the "Refresh Currency" method
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
    
    
    
    
    
    
    

'Find data extraction starting point using an index
For i = 1 To 10000

    'look for  cell with value="USD" iterating over the currencies extracted names
    If Worksheets("Sheet1").Range("A1:A" & 10000).Cells(i, 1).Text = "USD" Then
    Sheets("Sheet1").Select
    'New selection must be cell with USD currency rate, starting from this cell search selected index and convert the input amount
    
    'Find data extraction starting point using an index and save the cell position
    Range("A" & i).Select
    Exit For
    End If
    
Next i








'Create index that will iterate over the web service conversion rates (data extracted)

'"CONVERT FROM CURRENCY" name
indexToConv = converterForm.ComboBox1.ListIndex

'"CONVER TO CURRENCY" name
indexConv = converterForm.ComboBox2.ListIndex

'If the conversion rate is not available then display a user alert
If ActiveCell.Offset(indexToConv, 2).Value = 0 Or ActiveCell.Offset(indexConv, 2).Value = 0 Then

MsgBox "ONE OF THE CURRENCY RATES IS NOT AVAILABLE FOR THAT DATE, TRY A DIFFERENT YEAR "
Sheets("Sheet1").Visible = False
Exit Sub

End If







'Estimate the currency conversion and save the OUTPUT value that will be displayed in the form field
'Converted Value = Amount * ( Conversion currency rate ) / ( Initial currency rate)
converterForm.convertedOutput = converterForm.TextBox1 * (ActiveCell.Offset(indexConv, 2).Value / ActiveCell.Offset(indexToConv, 2).Value)

Sheets("Sheet1").Visible = False





End Sub
