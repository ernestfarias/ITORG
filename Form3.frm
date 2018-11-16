VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de password"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1080
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmacion:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Anterior:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Text2.Text = Text3.Text Then
HagoConexion
SQL = "update itusuarios set password = '" & Text2.Text & "'" & " where usuario='" & User & "'"
Conn.Execute SQL
Conn.Close
MsgBox "Se ha cambiado el password", vbOKOnly
Form3.Hide
Else
MsgBox "No coinciden las contraseñas, pruebe nuevamente", vbOKOnly
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Form3.Hide
End Sub

Private Sub Form_Load()
Form3.Caption = "Organizar password " & "(" & User & ")"
End Sub
