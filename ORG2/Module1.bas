Attribute VB_Name = "Module1"
Public Conn As ADODB.Connection
Public ID
Public ID2
Public Permiso
Public User
Public UserName
Public UserSector
Public Option1Selected As Integer
Public CantAntToDo
Dim PRN As Object

Sub HagoConexion()
Set Conn = New ADODB.Connection
Conn.Open "Provider=sqloledb;data Source=ar01mgr1sp;Initial Catalog=ITORGDESA;User Id=sa;Password=admALem855+"
'Conn.Open "Provider=sqloledb;data Source=ar01eft1vp;Initial Catalog=TEST;User Id=sa;Password=sa"
'Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\itorg.mdb"

End Sub


