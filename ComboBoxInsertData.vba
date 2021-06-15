'The sub will insert information in a ComboBox
Public Sub dataOnComboBox()
    lin = 1
    UserForm1.ComboBox.Clear
    Do Until Sheets("SheetName").Cells(lin, 1) = ""
      UserForm1.ComboBox.AddItem Sheets("aux").Cells(lin, 1)
      lin = lin + 1
    Loop
End Sub
