Attribute VB_Name = "modNumericOnly"
Public Function OnlyNumericKeys(KeyAscii As Integer, TextBox As Control) As Integer
  Select Case KeyAscii
    'allow Backspace(8),Numbers(48-57)
    Case 8, 46, 48 To 57 'refer the values here "http://msdn.microsoft.com/en-us/library/aa243025(v=vs.60).aspx" J4M
    
    Case Else
    
      KeyAscii = 0 'Reject everything else
      
  End Select
  
  OnlyNumericKeys = KeyAscii
  
End Function
