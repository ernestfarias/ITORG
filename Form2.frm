VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Validacion"
   ClientHeight    =   1620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3420
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1620
   ScaleWidth      =   3420
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB1 As clsDB
Public Usr1 As clsUsuario
Private Sub Command1_Click()

user = Text1.Text
Dim Found As Boolean
Found = False
Set Usr1 = New clsUsuario
Set DB1 = New clsDB
'DB1.StringConnection = "Provider=sqloledb;data Source=ar01mgr1sp;Initial Catalog=ITORGDESA2;User Id=sa;Password=admALem855+"
DB1.DBConnect

Dim RSusr As ADODB.Recordset
Set RSusr = New ADODB.Recordset

DB1.CargarRecordset RSusr, "select * from ITusuarios"


While Not RSusr.EOF
If Text1.Text = RSusr("usuario") Then
If Text2.Text = RSusr("password") Then
Found = True
Usr1.Nombre = RSusr("nombre")
Usr1.UserID = RSusr("usuario")
Usr1.Permisos = RSusr("permisos")
Usr1.Sector = RSusr("sector")
Usr1.Password = RSusr("password")
permiso = RSusr("permisos")
End If
End If
RSusr.MoveNext
Wend

If Found = True Then

'Form1.Show
GoTo FINOK
End If

RSusr.Close

FIN:
MsgBox "Error de usuario o password"
End
FINOK:
RSusr.Close
frmLogin.Hide
Form1.Show


End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
