Attribute VB_Name = "modFormPos"
Public Sub FormPos(FormName As Form)
With FormName
.Top = (Screen.Height - .Height) / 2
.Left = (Screen.Width - .Width) / 2
.KeyPreview = True
End With
End Sub
