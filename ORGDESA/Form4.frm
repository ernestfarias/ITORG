VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Exportar Archivo"
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Enviar a:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CommonDialog1.ShowOpen
Destino = CommonDialog1.FileName




Open Destino For Output As #2


HagoConexion
SQLuser = "where propietario ='" & Form1.Combo4.Text & "'"
SQLestado = "and estado='" & Form1.Option1(Option1Selected).Caption & "'"

If Form1.Combo4.Text = "Todos" Then
SQLuser = "where propietario IS NOT NULL "
End If

If Form1.Option1(Option1Selected).Caption = "Todas" Then
SQLestado = ""
End If




Dim RStodo As ADODB.Recordset
Set RStodo = New ADODB.Recordset

RStodo.Open "select * from ITtodo " & SQLuser & SQLestado & "order by prioridad", Conn

Write #2, "Prioridad", "Estado", "Porcentaje", "Descripcion", "Finicio", "Deadline"

While Not RStodo.EOF
Write #2, RStodo("prioridad"), RStodo("estado"), RStodo("porcentaje"), RStodo("descripcion"), RStodo("finicio"), RStodo("deadline")
RStodo.MoveNext

Wend

RStodo.Close
Close #2
End Sub

Private Sub File1_Click()

End Sub

Private Sub Command2_Click()
CommonDialog1.ShowOpen
caca = CommonDialog1.FileName
Text1.Text = caca
End Sub

Private Sub Form_Load()
'Combo1.Clear
HagoConexion
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset

RS.Open "select usuario from ITusuarios", Conn

'While Not RS.EOF
'Combo1.AddItem RS(0)
'RS.MoveNext
'Wend
'Combo1.Text = User
'RS.Close
'Conn.Close
'llenoGrid
End Sub
Private Sub llenoGrid()

End Sub
