Attribute VB_Name = "Module1"

Public ADOCN As ADODB.Connection
Public cadenacnx As String

Sub cargarconexion()
cadenacnx = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=" & App.Path & "\Base_SOLplus.mdb"
End Sub





