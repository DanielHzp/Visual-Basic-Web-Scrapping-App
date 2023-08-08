Attribute VB_Name = "LoadForm"
Option Explicit



'Executed when an action button is clicked
Sub openForm()

'Call method that loads currency datasets to comboxes
populatecomboboxes

'Update date and trace it
showDate
MsgBox "Today's date is: " & converterForm.dateText.Text

'Load user form
converterForm.Show

End Sub
