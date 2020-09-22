Attribute VB_Name = "Module1"
Public Sub CF(fName As Form)
'Centra la finestra sullo schermo
    fName.Left = (Screen.Width - fName.Width) / 2
    fName.Top = (Screen.Height - fName.Height) / 2
End Sub

