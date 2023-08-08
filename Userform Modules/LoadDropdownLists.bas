Attribute VB_Name = "LoadDropdownLists"
Option Explicit

'Load available currencies for conversions
Sub populatecomboboxes()

Dim i As Integer

'Handles spreadsheet visibility behavior
Sheets("Sheet2").Visible = True
Sheets("Sheet2").Select

'Set first record index
Range("A1").Select

'Avoid screen flickering when the user clicks and awaits command reponse
Application.ScreenUpdating = False

'Recursively iterate over currency names that must populate the combo boxes
'The query will save the currencies in column A of sheet 2
For i = 1 To WorksheetFunction.CountA(Worksheets("Sheet2").Columns("A:A"))
    
    'Add concatenated names ("Symbol - Display Name") to DROP DOWN LISTS
    converterForm.ComboBox1.AddItem ActiveCell.Offset(i - 1, 0) & "-" & ActiveCell.Offset(i - 1, 1)
    
    converterForm.ComboBox2.AddItem ActiveCell.Offset(i - 1, 0) & "-" & ActiveCell.Offset(i - 1, 1)
Next i


'Add titles of DROP DOWN LISTS
converterForm.ComboBox1.Text = Worksheets("Sheet2").Range("A1") & "-" & Worksheets("Sheet2").Range("B1")

converterForm.ComboBox2.Text = Worksheets("Sheet2").Range("A2") & "-" & Worksheets("Sheet2").Range("B2")

'Handles spreadsheet visibility behavior
Sheets("Sheet2").Visible = False
Sheets("ConverterSheet").Select
Range("A1").Select


End Sub
