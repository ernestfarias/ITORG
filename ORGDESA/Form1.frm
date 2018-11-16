VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " v"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8175
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   240
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   7680
      Begin VB.ListBox List3 
         Height          =   2010
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   7455
      End
      Begin VB.ComboBox Combo10 
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   "Combo10"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.ComboBox Combo9 
         Height          =   315
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox Combo8 
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   3360
         Width           =   735
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   1080
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   4920
         Width           =   6135
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   3600
         TabIndex        =   25
         Top             =   5520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Guardar"
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   5520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   6360
         TabIndex        =   23
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Modificar"
         Height          =   255
         Left            =   6360
         TabIndex        =   22
         Top             =   4560
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   6360
         TabIndex        =   21
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   4200
         Width           =   6135
      End
      Begin VB.Label Label22 
         Caption         =   "Categoria"
         Height          =   255
         Left            =   2520
         TabIndex        =   65
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Dueño:"
         Height          =   255
         Left            =   5640
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label14 
         Caption         =   "Comentarios"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label Label16 
         Caption         =   "Cantidad de Horas"
         Height          =   255
         Left            =   4800
         TabIndex        =   42
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Dia"
         Height          =   255
         Left            =   3000
         TabIndex        =   40
         Top             =   3360
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Propietario"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Prod 
         Caption         =   "ID"
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   2880
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton Command13 
         Caption         =   "Finalizar Tarea"
         Height          =   255
         Left            =   6120
         TabIndex        =   59
         Top             =   5520
         Width           =   1215
      End
      Begin VB.ComboBox Combo7 
         Height          =   315
         Left            =   3360
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   3720
         Width           =   615
      End
      Begin VB.TextBox Text16 
         Height          =   285
         Left            =   5880
         TabIndex        =   56
         Text            =   "Text16"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Height          =   495
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   7455
         Begin VB.ComboBox Combo4 
            Height          =   315
            ItemData        =   "Form1.frx":08CA
            Left            =   6000
            List            =   "Form1.frx":08CC
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   135
            Width           =   1410
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todas"
            Height          =   375
            Index           =   3
            Left            =   2895
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   105
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Finalizada"
            Height          =   375
            Index           =   2
            Left            =   1935
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   105
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "En Proceso"
            Height          =   375
            Index           =   1
            Left            =   975
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   105
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Pendiente"
            Height          =   375
            Index           =   0
            Left            =   15
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   105
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Dueño:"
            Height          =   255
            Left            =   5400
            TabIndex        =   53
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.TextBox Text15 
         Height          =   735
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   5040
         Width           =   5895
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   4320
         Width           =   5895
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Agregar"
         Height          =   255
         Left            =   6120
         TabIndex        =   9
         Top             =   4320
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Modificar"
         Height          =   255
         Left            =   6120
         TabIndex        =   10
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Borrar"
         Height          =   255
         Left            =   6120
         TabIndex        =   11
         Top             =   5040
         Width           =   1215
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3240
         Width           =   495
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3720
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Guardar"
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Top             =   5880
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   3240
         TabIndex        =   13
         Top             =   5880
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   2010
         ItemData        =   "Form1.frx":08CE
         Left            =   120
         List            =   "Form1.frx":08D0
         TabIndex        =   1
         Top             =   600
         Width           =   7455
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Label20"
         Height          =   195
         Left            =   5640
         TabIndex        =   60
         Top             =   5880
         Width           =   570
      End
      Begin VB.Label Label19 
         Caption         =   "Deadline"
         Height          =   255
         Left            =   5160
         TabIndex        =   57
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Prioridad"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         Height          =   255
         Left            =   2760
         TabIndex        =   32
         Top             =   3240
         Width           =   495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         Height          =   255
         Left            =   5400
         TabIndex        =   31
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin"
         Height          =   255
         Left            =   5520
         TabIndex        =   30
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Realizada (%)"
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Skill"
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   3720
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Comentarios"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Propietario"
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   2760
         Width           =   855
      End
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Refrescar"
      Height          =   255
      Left            =   2880
      TabIndex        =   46
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   6960
      Width           =   1095
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   11880
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tareas Programadas"
            Key             =   ""
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tareas Diarias"
            Key             =   ""
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Ocultar"
      Height          =   255
      Left            =   1440
      TabIndex        =   66
      Top             =   6960
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Exportar Archivo"
      Filter          =   "*.csv"
      Flags           =   1
      PrinterDefault  =   0   'False
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario logeado:"
      Height          =   255
      Left            =   5520
      TabIndex        =   45
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label11"
      Height          =   195
      Left            =   6840
      TabIndex        =   14
      Top             =   6960
      Width           =   570
   End
   Begin VB.Menu Opc 
      Caption         =   "Opciones"
      NegotiatePosition=   3  'Right
      Begin VB.Menu VistasUsurarios 
         Caption         =   "Vistas Usuarios"
         Enabled         =   0   'False
         WindowList      =   -1  'True
         Begin VB.Menu VistasU 
            Caption         =   "ver solo  mis tareas"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu VistasU 
            Caption         =   "ver todos los usuarios"
            Index           =   1
         End
      End
      Begin VB.Menu VistasEstado 
         Caption         =   "Vistas Estado"
         Enabled         =   0   'False
         Begin VB.Menu VistasE 
            Caption         =   "pendientes"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu VistasE 
            Caption         =   "en proceso"
            Index           =   1
         End
         Begin VB.Menu VistasE 
            Caption         =   "finalizadas"
            Index           =   2
         End
         Begin VB.Menu VistasE 
            Caption         =   "todas"
            Index           =   3
         End
      End
      Begin VB.Menu MSeguridad 
         Caption         =   "Seguridad"
         Begin VB.Menu MSUsuarios 
            Caption         =   "Usuarios..."
            Enabled         =   0   'False
         End
         Begin VB.Menu MSPassword 
            Caption         =   "Cambiar password"
         End
      End
   End
   Begin VB.Menu Consultas 
      Caption         =   "Consultas"
      Begin VB.Menu MCExportar 
         Caption         =   "Exportar"
      End
   End
   Begin VB.Menu Mimprimir 
      Caption         =   "Imprimir"
      Begin VB.Menu MprintActual 
         Caption         =   "Imprimir listado actual..."
      End
   End
   Begin VB.Menu about 
      Caption         =   "Info..."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Guardar
Dim GuardarDiarias
Dim Modificar
Dim ModificarDiarias
'---
Private Type NotifyIconData
  Size As Long
  Handle As Long
  ID As Long
  Flags As Long
  CallBackMessage As Long
  Icon As Long
  Tip As String * 64
End Type

' Constants for managing System Tray tasks, foudn in shellapi.h
Private Const AddIcon = &H0
Private Const ModifyIcon = &H1
Private Const DeleteIcon = &H2

Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202

Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Const MessageFlag = &H1
Private Const IconFlag = &H2
Private Const TipFlag = &H4
      
Private Declare Function Shell_NotifyIcon _
  Lib "shell32" Alias "Shell_NotifyIconA" ( _
  ByVal Message As Long, Data As NotifyIconData) As Boolean

Private Data As NotifyIconData

Private Sub Command14_Click()
Form1.Hide
End Sub

Private Sub Form_Resize()
'If Form1.Visible = True Then
'Form1.Hide
'End If
End Sub

Private Sub Form_Terminate()
  DeleteIconFromTray
End Sub

Private Sub AddIconToTray()

  Data.Size = Len(Data)
  Data.Handle = hWnd
  Data.ID = vbNull
  
  Data.Flags = IconFlag Or TipFlag Or MessageFlag
  Data.CallBackMessage = WM_MOUSEMOVE
  Data.Icon = Icon
  Data.Tip = App.Title & vbNullChar
  Call Shell_NotifyIcon(AddIcon, Data)

End Sub

Private Sub DeleteIconFromTray()
  Call Shell_NotifyIcon(DeleteIcon, Data)
End Sub

Private Sub Form_MouseMove(Button As Integer, _
  Shift As Integer, X As Single, Y As Single)
  
  Dim Message As Long
  Message = X / Screen.TwipsPerPixelX
  
  Select Case Message
    Case WM_LBUTTONDBLCLK
     'Form1.Show
      Form1.Visible = Not Form1.Visible
     Form1.WindowState = Abs(Not Form1.Visible)
  End Select
End Sub

Private Sub about_Click()
MsgBox "Desarrollado por Ernesto A Farias , int. 5065" & vbCrLf & "Ver. :" & App.Major & "." & App.Minor & "." & App.Revision, vbOKOnly
Info
End Sub

Private Sub Combo4_Click()
LeoTodo
End Sub

Private Sub Combo9_Click()
LeoDiarias
End Sub

Private Sub Command1_Click()
Unload Form1
Unload Form2
End
End Sub
Private Sub GuardoCat()
HagoConexion
Dim RSCat As ADODB.Recordset
Set RSCat = New ADODB.Recordset
With RSCat
    .ActiveConnection = Conn
   ' .CursorLocation = adUseClient
   ' .CursorType = adOpenKeyset
    .LockType = adLockPessimistic
    .Source = "SELECT * FROM ITcategorias"
    .Open
    
    .AddNew
    .Fields("categoria").Value = Combo10.Text
.Update
.Close
End With
End Sub

Private Sub Command10_Click()
'GUARDO CAT
encontro = 0
For i = 0 To Combo10.ListCount
If Combo10.Text = Combo10.List(i) Then
encontro = 1
End If
Next

If encontro = 0 Then
Res = MsgBox("La categoria no existe en la lista, desea agregarla?", vbYesNo)
If Res = vbNo Then
Combo10.SetFocus
GoTo NoGuardar
End If

If Res = vbYes Then
GuardoCat
End If

End If
'/GUARDO CAT


If ModificarDiarias = True Then
ModificarDiarias = False
Command10.Visible = False
Command11.Visible = False
HagoConexion
    Dim RSDiarias As ADODB.Recordset
Set RSDiarias = New ADODB.Recordset
With RSDiarias
    .ActiveConnection = Conn
   ' .CursorLocation = adUseClient
   ' .CursorType = adOpenKeyset
    .LockType = adLockPessimistic
    .Source = "SELECT * FROM ITdiarias where id='" & Text7.Text & "'"
    .Open
     
    .Fields("propietario").Value = Combo6.Text
'    .Fields("canthoras").Value = Text11.Text
    .Fields("categoria").Value = Combo10.Text
    .Fields("dia").Value = Text10.Text
    .Fields("descripcion").Value = Text5.Text
    .Fields("comentarios").Value = Text14.Text
    .Fields("horas").Value = Combo8.Text
    
    .Update
    .Close
    
    End With

Guardar = False

End If

If GuardarDiarias = True Then

HagoConexion
    'Dim RStodo As ADODB.Recordset
Set RSDiarias = New ADODB.Recordset
With RSDiarias
    .ActiveConnection = Conn
   ' .CursorLocation = adUseClient
   ' .CursorType = adOpenKeyset
    .LockType = adLockPessimistic
    .Source = "SELECT * FROM ITdiarias"
    .Open
    
    .AddNew
    .Fields("propietario").Value = Combo6.Text
  '  .Fields("canthoras").Value = Text11.Text
   ' .Fields("cantminutos").Value = Text13.Text
    .Fields("categoria").Value = Combo10.Text
    .Fields("dia").Value = Text10.Text
    .Fields("descripcion").Value = Text5.Text
    .Fields("comentarios").Value = Text14.Text
    .Fields("horas").Value = Combo8.Text
 
 
    .Update
    .Close
    
    End With

GuardarDiarias = False
End If
    
    
Command10.Visible = False
Command11.Visible = False
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Text7.Locked = True
Text10.Locked = True
'Text4.Locked = True
Combo6.Locked = True
'Text11.Locked = True
'Text13.Locked = True
Combo8.Locked = True
Combo10.Locked = True
Text5.Locked = True
Text14.Locked = True
    LeoDiarias
    LeoCat
NoGuardar:
End Sub

Private Sub Command11_Click()
LeoDiarias
Command7.Enabled = True
Command8.Enabled = True
Command9.Enabled = True
Command10.Visible = False
Command11.Visible = False
Text7.Locked = True
Text10.Locked = True
'Text4.Locked = True
Combo6.Locked = True
Combo8.Locked = True
Combo10.Locked = True
'Text11.Locked = True
'Text13.Locked = True
Text5.Locked = True
Text14.Locked = True
End Sub

Private Sub Command12_Click()
LeoDiarias
LeoTodo
End Sub

Private Sub Command13_Click()
Command3_Click
Combo2.Text = "Finalizada"
Combo7.Text = "100"
Text9.Text = Format(Date, "d/m/yyyy") 'pongo la fecha de hoy
Command5_Click
End Sub

Private Sub Command2_Click()
Text1.SetFocus
Guardar = True
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False

Command5.Visible = True
Command6.Visible = True
'hoy = CStr(Format(Date, "long date"))
Text1.Text = ""
'Text2.Text = User
Combo5.Text = User
'Text6.Text = 0
Text8.Text = Format(Date, "d/m/yyyy")
Text9.Text = ""
Text16.Text = ""

If Permiso = "admin" Then
Combo5.Locked = False
Combo5.Enabled = True
End If

Text3.Locked = False
'Text2.Locked = False
Combo3.Locked = False
Combo1.Locked = False
Combo2.Locked = False
'Text6.Locked = False
Combo7.Locked = False
Text16.Locked = False
Text8.Locked = False
Text9.Locked = False
Text1.Locked = False
Text15.Locked = False

End Sub

Private Sub Command3_Click()
If User <> Combo5.Text And Permiso <> "admin" Then
MsgBox "No tiene permiso para modificar tareas ajenas,gracias", vbOKOnly
GoTo NoModifica
End If

Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Visible = True
Command6.Visible = True
Modificar = True

If Permiso = "admin" Then
Combo5.Locked = False
Combo5.Enabled = True
End If

Text3.Locked = False
Combo5.Locked = False
'Text2.Locked = False
Combo3.Locked = False
Combo1.Locked = False
Combo2.Locked = False
'Text6.Locked = False
Combo7.Locked = False
Text16.Locked = False
Text8.Locked = False
Text9.Locked = False
Text1.Locked = False
Text15.Locked = False
NoModifica:
End Sub

Private Sub Command4_Click()
Res1 = MsgBox("Esta seguro que desea borrar: " & vbCrLf & Text1.Text, vbYesNo, "BORRAR")
If Res1 = vbNo Then
GoTo NoBorra
End If

If User <> Combo5.Text And Permiso <> "admin" Then
MsgBox "No tiene permiso para borrar tareas de otros usuarios", vbOKOnly
GoTo NoBorra
End If
HagoConexion
Conn.Execute "delete from ittodo where id='" & Text3.Text & "'"
    LeoTodo
    List1.ListIndex = (List1.ListCount - 1)

NoBorra:

End Sub

Private Sub Command5_Click()
If Modificar = True Then
Modificar = False
Command5.Visible = False
Command6.Visible = False
HagoConexion
    Dim RStodo As ADODB.Recordset
Set RStodo = New ADODB.Recordset
With RStodo
    .ActiveConnection = Conn
   ' .CursorLocation = adUseClient
   ' .CursorType = adOpenKeyset
    .LockType = adLockPessimistic
    .Source = "SELECT * FROM ITtodo where id='" & Text3.Text & "'"
    .Open
    
    .Fields("propietario").Value = Combo5.Text
    .Fields("prioridad").Value = Combo1.Text
    .Fields("estado").Value = Combo2.Text
    .Fields("porcentaje").Value = Combo7.Text
    .Fields("skill").Value = Combo3.Text
    .Fields("finicio").Value = Text8.Text
    .Fields("ffin").Value = Text9.Text
    .Fields("descripcion").Value = Text1.Text
    .Fields("comentarios").Value = Text15.Text
    .Fields("fmodificado").Value = Now
    .Fields("deadline").Value = Text16.Text
    .Update
    .Close
    
    End With

Guardar = False

End If

If Guardar = True Then

HagoConexion
    'Dim RStodo As ADODB.Recordset
Set RStodo = New ADODB.Recordset
With RStodo
    .ActiveConnection = Conn
   ' .CursorLocation = adUseClient
   ' .CursorType = adOpenKeyset
    .LockType = adLockPessimistic
    .Source = "SELECT * FROM ITtodo"
    .Open
    
    .AddNew
'    .Fields("propietario").Value = Text2.Text
     .Fields("propietario").Value = Combo5.Text
    .Fields("prioridad").Value = Combo1.Text
    .Fields("estado").Value = Combo2.Text
    .Fields("porcentaje").Value = Combo7.Text
    .Fields("skill").Value = Combo3.Text
    .Fields("finicio").Value = Text8.Text
    .Fields("ffin").Value = Text9.Text
    .Fields("descripcion").Value = Text1.Text
    .Fields("comentarios").Value = Text15.Text
    .Fields("fcreado").Value = Now
    .Fields("fmodificado").Value = Now
    .Fields("deadline").Value = Text16.Text
    
    .Update
    .Close
    
    End With

Guardar = False
End If
    
    
Command5.Visible = False
Command6.Visible = False
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Text3.Locked = True
'Text2.Locked = True
Combo5.Locked = True
Combo3.Locked = True
Combo1.Locked = True
Combo2.Locked = True
Combo7.Locked = True
'Text6.Locked = True
Text8.Locked = True
Text9.Locked = True
Text1.Locked = True
Text15.Locked = True
Text16.Locked = True
    LeoTodo




End Sub

Private Sub Command6_Click()
LeoTodo
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Visible = False
Command6.Visible = False
Text3.Locked = True
'Text2.Locked = True
Combo5.Locked = True
Combo3.Locked = True
Combo1.Locked = True
Combo2.Locked = True
Text16.Locked = True
'Text6.Locked = True
Combo7.Locked = True
Text8.Locked = True
Text9.Locked = True
Text1.Locked = True
Text15.Locked = True
End Sub

Private Sub Command7_Click()
Text5.SetFocus
GuardarDiarias = True
Combo10.Locked = False
Text7.Locked = False
Text10.Locked = False
'Text4.Locked = False
'Combo6.Locked = False
'Text11.Locked = False
'Text13.Locked = False
Combo8.Locked = False
Text5.Locked = False
Text14.Locked = False

Command10.Visible = True
Command11.Visible = True
Command8.Enabled = False
Command9.Enabled = False
Command7.Enabled = False
Text5.Text = ""
'Text4.Text = User
Combo6.Text = User
Text10.Text = Format(Date, "d/m/yyyy")
Text14.Text = ""
If Permiso = "admin" Then
Combo6.Enabled = True
Combo6.Locked = False
End If
End Sub

Private Sub Command8_Click()
If User <> Combo6.Text And Permiso <> "admin" Then
MsgBox "No tiene permiso para borrar tareas de otros usuarios", vbOKOnly
GoTo NoModificadiaria
End If

Combo10.Locked = False
ModificarDiarias = True
Command10.Visible = True
Command11.Visible = True
Command8.Enabled = False
Command9.Enabled = False
Command7.Enabled = False
Text7.Locked = False
Text10.Locked = False
'Text4.Locked = False
'Combo6.Locked = False
'Text11.Locked = False
'Text13.Locked = False
Combo8.Locked = False
Text5.Locked = False
Text14.Locked = False

If Permiso = "admin" Then
Combo6.Enabled = True
Combo6.Locked = False
End If

NoModificadiaria:
End Sub

Private Sub Command9_Click()
Res2 = MsgBox("Esta seguro que desea borrar: " & vbCrLf & Text5.Text, vbYesNo, "BORRAR")
If Res2 = vbNo Then
GoTo NoBorraDiaria
End If

If User <> Combo6.Text And Permiso <> "admin" Then
MsgBox "No tiene permiso para borrar tareas de otros usuarios", vbOKOnly
GoTo NoBorraDiaria
End If

HagoConexion
Conn.Execute "delete from itdiarias where id='" & Text7.Text & "'"
    LeoDiarias
List3.ListIndex = (List3.ListCount - 1)

NoBorraDiaria:
End Sub
Public Sub Form_Load()
Form1.Caption = App.ProductName & " " & App.Major & "." & App.Minor & "." & App.Revision
Frame1.Top = 480
Frame1.Left = 240
Frame2.Top = 480
Frame2.Left = 240
Frame1.Visible = True
Label11.Caption = User
'DoEvents
AddIconToTray

Combo1.Clear
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "5"
Combo1.Text = Combo1.List(1)

Combo2.Clear
Combo2.AddItem "En Proceso"
Combo2.AddItem "Pendiente"
Combo2.AddItem "Finalizada"
Combo2.Text = Combo2.List(0)

Combo3.Clear
Combo3.AddItem "1"
Combo3.AddItem "2"
Combo3.AddItem "3"
Combo3.AddItem "4"
Combo3.AddItem "5"
Combo3.Text = Combo3.List(2)

Combo7.Clear
Combo7.AddItem "0"
Combo7.AddItem "10"
Combo7.AddItem "20"
Combo7.AddItem "30"
Combo7.AddItem "40"
Combo7.AddItem "50"
Combo7.AddItem "60"
Combo7.AddItem "70"
Combo7.AddItem "80"
Combo7.AddItem "90"
Combo7.AddItem "100"
Combo7.Text = Combo7.List(0)

Combo8.Clear
Combo8.AddItem "00:30"
Combo8.AddItem "01:00"
Combo8.AddItem "01:30"
Combo8.AddItem "02:00"
Combo8.AddItem "02:30"
Combo8.AddItem "03:00"
Combo8.AddItem "03:30"
Combo8.AddItem "04:00"
Combo8.AddItem "04:30"
Combo8.AddItem "05:00"
Combo8.AddItem "05:30"
Combo8.AddItem "06:00"
Combo8.AddItem "06:30"
Combo8.AddItem "07:00"
Combo8.AddItem "07:30"
Combo8.AddItem "08:00"
Combo8.Text = Combo8.List(0)


Text16.Text = ""

Option1Selected = 0 ' por default pendientes siempre muestra primero
LeoUsuario
LeoTodo
LeoDiarias
LeoCat
On Error Resume Next
List1.ListIndex = (List1.ListCount - List1.ListCount + 1)
 
End Sub
Public Sub LeoCat()
Combo10.Clear
HagoConexion
Dim RSCat As ADODB.Recordset
Set RSCat = New ADODB.Recordset
RSCat.Open "select categoria from itcategorias", Conn
While Not RSCat.EOF
Combo10.AddItem RSCat(0)
RSCat.MoveNext
Wend
RSCat.Close

End Sub
Public Sub LeoDiarias()
List3.Clear
HagoConexion


SQLuser = "where propietario ='" & Combo9.Text & "'"
If Combo9.Text = "Todos" Then
SQLuser = "where propietario IS NOT NULL "
End If

Dim RSDiarias As ADODB.Recordset
Set RSDiarias = New ADODB.Recordset

RSDiarias.Open "select * from ITdiarias " & SQLuser & "order by id desc", Conn

While Not RSDiarias.EOF
List3.AddItem RSDiarias(0) & "|    " & RSDiarias(1) & "    " & RSDiarias("dia") & " " & RSDiarias("horas") & " " & vbTab & RSDiarias("descripcion")
RSDiarias.MoveNext
Wend

RSDiarias.Close

End Sub
Public Sub LeoTodo()
List1.Clear
HagoConexion
SQLuser = "where propietario ='" & Combo4.Text & "'"
SQLestado = "and estado='" & Option1(Option1Selected).Caption & "'"

If Combo4.Text = "Todos" Then
SQLuser = "where propietario IS NOT NULL "
End If

If Option1(Option1Selected).Caption = "Todas" Then
SQLestado = ""
End If




'If User = "admin" Then
'SQLuser = ""
'End If
Dim RStodo As ADODB.Recordset
Set RStodo = New ADODB.Recordset

'RStodo.Open "select * from ITtodo " & SQLuser & "order by id desc", Conn
RStodo.Open "select * from ITtodo " & SQLuser & SQLestado & "order by prioridad", Conn


While Not RStodo.EOF
List1.AddItem RStodo(0) & "|    " & RStodo("estado") & vbTab & RStodo("porcentaje") & "%" & vbTab & RStodo("descripcion")
RStodo.MoveNext

Wend

RStodo.Close

End Sub

Public Sub LeoUsuario()
Combo4.Clear
HagoConexion
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset

RS.Open "select usuario from ITusuarios", Conn

While Not RS.EOF
Combo4.AddItem RS(0)
Combo5.AddItem RS(0)
Combo6.AddItem RS(0)
Combo9.AddItem RS(0)
RS.MoveNext
Wend
Combo4.Text = User
Combo5.Text = User
Combo6.Text = User
Combo9.Text = User
Combo9.AddItem "Todos"
Combo4.AddItem "Todos"
RS.Close

End Sub


Private Sub Form_Unload(Cancel As Integer)
 DeleteIconFromTray
End
End Sub

Public Sub List1_Click()
On Error Resume Next
List1.ToolTipText = List1.List(List1.ListIndex)
Lt = InStr(List1.List(List1.ListIndex), "|") 'me da la pos donde esta el 1er |
Lt = CInt(Lt) - 1 'le resto 1 para que me saque el caracte "|"
ID = Left(List1.List(List1.ListIndex), Lt) ' me da el nro del campo id de la lista, todo para sacar el id para mostrar sus datos en el otro cuadro
CargoForm

End Sub
Private Sub CargoForm()
'On Error Resume Next
HagoConexion

Dim RStodo As ADODB.Recordset
Set RStodo = New ADODB.Recordset

RStodo.Open "select * from ITtodo where id='" & ID & "'", Conn

While Not RStodo.EOF
Text1.Text = RStodo("descripcion")
'Text2.Text = RStodo("propietario")
'Combo5.Text = RStodo("propietario")
Combo5.Text = RStodo("propietario")
Text3.Text = RStodo("id")
Combo1.Text = RStodo("prioridad")
Combo2.Text = RStodo("estado")
Combo7.Text = RStodo("porcentaje")
Combo3.Text = RStodo("skill")
Text8.Text = RStodo("finicio")
Text9.Text = RStodo("ffin")
Text15.Text = RStodo("comentarios")
Text16.Text = RStodo("deadline")
RStodo.MoveNext
Wend

If Combo2.Text <> "Finalizada" Then 'activo o no el boton de finalizar
Command13.Enabled = True
Else
Command13.Enabled = False
End If

Label20.Caption = "Deadline en: " & DateDiff("d", Format(Date, "d/m/yyyy"), Format(Text16.Text, "d/m/yyyy")) & " dias"

RStodo.Close

End Sub
Private Sub CargoForm2()
'On Error Resume Next
HagoConexion

Dim RSDiarias As ADODB.Recordset
Set RSDiarias = New ADODB.Recordset

RSDiarias.Open "select * from ITDiarias where id='" & ID2 & "'", Conn

While Not RSDiarias.EOF
Text7.Text = RSDiarias("id")
Text10.Text = RSDiarias("dia")
 
Combo6.Text = RSDiarias("propietario")
Combo10.Text = RSDiarias("categoria")
'Text13.Text = RSDiarias("cantminutos")
Text5.Text = RSDiarias("descripcion")
Text14.Text = RSDiarias("comentarios")
Combo8.Text = RSDiarias("horas")
RSDiarias.MoveNext
Wend

RSDiarias.Close
End Sub

Private Sub List3_Click()
List3.ToolTipText = List3.List(List3.ListIndex)
Lt2 = InStr(List3.List(List3.ListIndex), "|") 'me da la pos donde esta el 1er |
Lt2 = CInt(Lt2) - 1 'le resto 1 para que me saque el caracte "|"
ID2 = Left(List3.List(List3.ListIndex), Lt2) ' me da el nro del campo id de la lista, todo para sacar el id para mostrar sus datos en el otro cuadro
CargoForm2
End Sub

Private Sub MCExportar_Click()



If TabStrip1.SelectedItem = "Tareas Programadas" Then
CommonDialog1.FileName = User & Form1.Option1(Option1Selected).Caption & ".csv"
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
End If

If TabStrip1.SelectedItem = "Tareas Diarias" Then
CommonDialog1.FileName = User & "Diarias" & ".csv"
CommonDialog1.ShowOpen
Destino = CommonDialog1.FileName


Open Destino For Output As #2

HagoConexion


SQLuser = "where propietario ='" & Combo9.Text & "'"
If Combo9.Text = "Todos" Then
SQLuser = "where propietario IS NOT NULL "
End If

Dim RSDiarias As ADODB.Recordset
Set RSDiarias = New ADODB.Recordset

RSDiarias.Open "select * from ITdiarias " & SQLuser & "order by id desc", Conn


Write #2, "", "", "dia", "Descripcion", "horas", "descripcion"
While Not RSDiarias.EOF
Write #2, RSDiarias(0), RSDiarias(1), RSDiarias("dia"), RSDiarias("horas"), RSDiarias("descripcion")
RSDiarias.MoveNext

Wend



RSDiarias.Close


End If


Close #2

End Sub
Private Sub MprintActual_Click()
   CommonDialog1.CancelError = True
   On Error Resume Next
   CommonDialog1.ShowPrinter
   If Err.Number = 32755 Then Exit Sub
   If CommonDialog1.Orientation = cdlLandscape Then
       Printer.Orientation = cdlLandscape
   Else
       Printer.Orientation = cdlPortrait
   End If
   
      printCaca
      End Sub
      
    Private Sub printCaca()
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
PRN.Print
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
PRN.Print
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

Private Sub MSPassword_Click()
Form3.Show
End Sub

Private Sub Option1_Click(Index As Integer)

Option1Selected = Index
LeoTodo
End Sub

Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem = "Tareas Diarias" Then
Frame2.Visible = True
Frame1.Visible = False
End If
If TabStrip1.SelectedItem = "Tareas Programadas" Then
Frame1.Visible = True
Frame2.Visible = False
End If
End Sub

Private Sub Text11_gotfocus()
Text11.SelStart = 0
Text11.SelLength = 2

End Sub

Private Sub Text13_GotFocus()
Text13.SelStart = 0

Text13.SelLength = 2
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

End Sub

Private Sub VistasE_Click(Index As Integer)
For i = 0 To VistasE.uBound
VistasE(i).Checked = False
Next
VistasE(Index).Checked = True
End Sub

Private Sub VistasU_Click(Index As Integer)
For i = 0 To VistasU.uBound
VistasU(i).Checked = False
Next
VistasU(Index).Checked = True

End Sub
Private Sub Info()
'HagoConexion

'Dim RS As ADODB.Recordset
'Set RS = New ADODB.Recordset

'RS.Open "select ID from ITDiarias", Conn
'RS.Close

End Sub
