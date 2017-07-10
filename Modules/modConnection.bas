Attribute VB_Name = "modConnection"
Global strSQL As String
Global strCommand As String
Global recSet As ADODB.Recordset
Global Conn As ADODB.Connection

Public Sub DBCON()
strCommand = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source =" & App.Path & "/La_Esperanza_Hotel_DataBase.mdb"

Set Conn = New ADODB.Connection
    With Conn
        .ConnectionString = strCommand
        .Open
    End With
    
End Sub
