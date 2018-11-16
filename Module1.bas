Attribute VB_Name = "Module1"
Public Conn As ADODB.Connection
Public ID
Public ID2
Public Permiso
Public User
Public UserName
Public UserSector
Public Option1Selected As Integer
Dim PRN As Object

Sub HagoConexion()
Set Conn = New ADODB.Connection
Conn.Open "Provider=sqloledb;data Source=ar01mgr1sp;Initial Catalog=ITORG;User Id=sa;Password=admALem855+"
End Sub


