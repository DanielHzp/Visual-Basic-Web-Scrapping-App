Attribute VB_Name = "UpdateCurrencyRates"
Option Explicit

'Updates the currency conversion rates and fetch data from web service www.XE.com (currencies API)
Sub refreshCombobox()

Dim i As Integer, k As Integer, url As String

'Handles spreadsheet visibility behavior when user clicks any button
Sheets("Sheet1").Visible = True

'Clear data to avoid OVERWRITING
Sheets("Sheet1").Cells.Clear

'Clean data on first dropdown lists in order to avoid OVERWRITING
For i = converterForm.ComboBox1.ListCount - 1 To 0 Step -1
        converterForm.ComboBox1.RemoveItem i
Next i



'Clean data on second dropdown lists in order to avoid OVERWRITING
For i = converterForm.ComboBox2.ListCount - 1 To 0 Step -1
        converterForm.ComboBox2.RemoveItem i
Next i

'url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & yearDate & "-" & monthDate & "-" & dayDate   one way

    'Set the query URL connection of the website www.XE.com
    url = "URL;https://www.xe.com/currencytables/?from=USD&date=" & Format(converterForm.dateText.Text, "yyyy-mm-dd")
    'Ideally this URL should be parameterized and administered in a spreadsheet cell
    
    
    
    'Create the connection to start data scraping from the website
    With Worksheets("Sheet1").QueryTables.Add(Connection:=url, Destination:=Worksheets("Sheet1").Range("A1"))
    
    'Configure the query parameters to save the extracted data
        .Name = "My Query"
        .RowNumbers = False
        .FillAdjacentFormulas = False
        
        'Use native data format
        .PreserveFormatting = True
        
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlOverwriteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        
        'Fetch the entire site metadata to avoid missing data
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        
        'Store data in column format for easier manipulation
        .WebPreFormattedTextToColumns = True
        
        'Avoid large data groups separated by columns
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        
        'Treat dates as strings
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        
        'Avoid query automatic updates
        .Refresh BackgroundQuery:=False
        
    End With
    



'Find data extraction starting point using an index
For i = 1 To 10000
    
    'look for  cell with value="USD" iterating over the currencies extracted names
   If Worksheets("Sheet1").Range("A1:A" & 10000).Cells(i, 1).Text = "USD" Then
    Sheets("Sheet1").Select
    
    'Find data extraction starting point using an index
    Range("A" & i).Select
    Exit For
    End If
Next i
    



'Force all dropdown lists to display updated currencies after data extraction
Do

    'If the currency symbol format is incorrect stop iteration
    If Len(ActiveCell.Offset(k, 0).Text) <> 3 Then Exit Do
    
    
    converterForm.ComboBox1.AddItem ActiveCell.Offset(k, 0) & "-" & ActiveCell.Offset(k, 1)
    converterForm.ComboBox2.AddItem ActiveCell.Offset(k, 0) & "-" & ActiveCell.Offset(k, 1)
    
    k = k + 1
    
Loop




'Handles spreadsheet visibility behavior after data scraping ends
Sheets("Sheet1").Visible = False
Sheets("ConverterSheet").Select
End Sub



