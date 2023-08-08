Attribute VB_Name = "SaveQuit"
Option Explicit

Sub HIDEUNHIDE()
Attribute HIDEUNHIDE.VB_ProcData.VB_Invoke_Func = " \n14"
'
' HIDEUNHIDE web scraping data output
'

'
    Sheets("Sheet2").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Sheet1").Select
    Sheets("Sheet2").Visible = True
    
    
    
End Sub
