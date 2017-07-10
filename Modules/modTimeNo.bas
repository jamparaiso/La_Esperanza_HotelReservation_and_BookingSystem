Attribute VB_Name = "modTimeNo"
Public Function TimeNo(Time As String) As Long
TimeNo = CLng(Replace(Format(Time, "hhnnss"), ":", ""))
End Function
