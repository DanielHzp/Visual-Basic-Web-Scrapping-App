VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} converterForm 
   Caption         =   "Currency Converter v1.0"
   ClientHeight    =   6264
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10560
   OleObjectBlob   =   "ConverterUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "converterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Button on-click actions ------------------------------------- Add here button commands to render data in the form -----------------------------

'Plot currency daily behavior
Private Sub CommandButton1_Click()

Call plotlastthirtydays

End Sub


'Convert input amount into selected currency

Private Sub CommandButton2_Click()

'Input validations before calculating amount

If TextBox1.Value < 0 Or TextBox1.Value = 0 Then
MsgBox " Enter a positive amount"
Exit Sub
ElseIf TextBox1.Value = "" Then
MsgBox "Enter a correct amount"
Exit Sub
ElseIf dateText.Text = "" Then
MsgBox "invalid date"
Exit Sub
ElseIf IsDate(dateText) = False Then
MsgBox " Input date isn't in the correct format"
Exit Sub
End If

If Not IsNumeric(TextBox1.Value) Then
    MsgBox "currency must be a number"
    Exit Sub
    End If


Call converterQuery


End Sub

'Close form and refresh drop down lists
Private Sub CommandButton3_Click()
Dim i As Integer
For i = ComboBox1.ListCount - 1 To 0 Step -1
        ComboBox1.RemoveItem i
Next i
For i = ComboBox2.ListCount - 1 To 0 Step -1
        ComboBox2.RemoveItem i
Next i


Unload converterForm

End Sub

'Refresh drop down lists and update rate display names
Private Sub CommandButton4_Click()

Call refreshCombobox


End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
'call UpdateCurrencyRates
End Sub
