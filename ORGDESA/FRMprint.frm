VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRMprint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   40005
   ClientLeft      =   3315
   ClientTop       =   -3675
   ClientWidth     =   10950
   HasDC           =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   705.644
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   193.146
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6360
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   10935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "llenar"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   2295
   End
End
Attribute VB_Name = "FRMprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.PrintForm

End Sub

Private Sub Command3_Click()
   CommonDialog1.CancelError = True
   On Error Resume Next
   ' display the dialog
   CommonDialog1.ShowPrinter
   If Err.Number = 32755 Then Exit Sub ' user cancelled
   If CommonDialog1.Orientation = cdlLandscape Then
       Printer.Orientation = cdlLandscape
   Else
       Printer.Orientation = cdlPortrait
   End If
   
      caca
End Sub

Private Sub caca()
Set PRN = Printer
Printer.Font.Size = 8
FRMprint.List1.Height = CInt(Form1.List1.ListCount * 240)
FRMprint.Height = CInt(Form1.List1.ListCount * 240) + 800

If Form1.TabStrip1.SelectedItem = "Tareas Programadas" Then

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
PRN.Print vbTab & vbTab & vbTab & "Usuario: " & User & vbTab & "Fecha de impresion: " & Now
PRN.Print
PRN.Print RStodo(0).Name & "|    " & RStodo("estado").Name & vbTab & RStodo("porcentaje").Name & "%" & vbTab & RStodo("descripcion").Name
PRN.Print

While Not RStodo.EOF
PRN.Print RStodo(0) & "|    " & RStodo("estado") & vbTab & RStodo("porcentaje") & "%" & vbTab & RStodo("descripcion")
RStodo.MoveNext
Wend
PRN.Print vbTab & vbTab & vbTab & "Usuario: " & User & vbTab & "Fecha de impresion: " & Now

RStodo.Close

End If

If Form1.TabStrip1.SelectedItem = "Tareas Diarias" Then
HagoConexion
SQLuser = "where propietario ='" & Form1.Combo9.Text & "'"
If Form1.Combo9.Text = "Todos" Then
SQLuser = "where propietario IS NOT NULL "
End If

Dim RSDiarias As ADODB.Recordset
Set RSDiarias = New ADODB.Recordset

RSDiarias.Open "select * from ITdiarias " & SQLuser & "order by id desc", Conn
PRN.Print vbTab & vbTab & vbTab & "Usuario: " & User & vbTab & "Fecha de impresion: " & Now
PRN.Print
PRN.Print RSDiarias(0).Name & "    " & RSDiarias(1).Name & "    " & RSDiarias("dia").Name & " " & RSDiarias("horas").Name & " " & vbTab & RSDiarias("descripcion").Name
PRN.Print

While Not RSDiarias.EOF
PRN.Print RSDiarias(0) & "|    " & RSDiarias(1) & "    " & RSDiarias("dia") & " " & RSDiarias("horas") & " " & vbTab & RSDiarias("descripcion")
RSDiarias.MoveNext
Wend
PRN.Print vbTab & vbTab & vbTab & "Usuario: " & User & vbTab & "Fecha de impresion: " & Now


RSDiarias.Close


End If


'Close #2


'For i = 0 To Form1.List1.ListCount
'i = i + 1
'Text1.Text = Form1.List1.List(Form1.List1.ListIndex)
'Next
Printer.EndDoc
End Sub

 
 
