VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2235
   ClientLeft      =   3150
   ClientTop       =   1365
   ClientWidth     =   3900
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2235
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   1740
      Width           =   1395
   End
   Begin VB.CommandButton ok 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   900
      TabIndex        =   7
      Top             =   1740
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1572
      Left            =   96
      TabIndex        =   3
      Top             =   96
      Width           =   3720
      Begin VB.TextBox lPassword 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1620
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox lUsername 
         Height          =   345
         Left            =   1620
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lDatabaseName 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BPDC"
         Height          =   315
         Left            =   1620
         TabIndex        =   6
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuario :"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   720
         TabIndex        =   0
         Top             =   288
         Width           =   756
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Password :"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   564
         TabIndex        =   1
         Top             =   624
         Width           =   924
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&SID :"
         ForeColor       =   &H80000008&
         Height          =   192
         Left            =   1032
         TabIndex        =   2
         Top             =   1248
         Width           =   408
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancel_Click()
 Unload frmLogin
 If Not EstadoConexion Then End
End Sub

Private Sub Form_Load()

Call CenterForm(frmLogin)


End Sub

Private Sub lDatabaseName_GotFocus()

 
 

End Sub

Private Sub lPassword_GotFocus()
 
 lPassword.SelStart = 0
 lPassword.SelLength = Len(lPassword.Text)

End Sub

Private Sub lPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then OK_Click
End Sub

Private Sub lUsername_GotFocus()
 
 lUsername.SelStart = 0
 lUsername.SelLength = Len(lUsername.Text)

End Sub

Private Sub lUsername_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then OK_Click
End Sub

Private Sub OK_Click()


Set conbasa = New ADODB.Connection
    UserName = "BASA"
    Password = "1742"
    conbasa.Open "Provider=MSDAORA.1;Password=1742;User ID=basa;Data Source=bpdc;Persist Security Info=True"

'set the global var to false
    'to denote a failed login
    frmCargarRequerimientos.Show


' If lUsername.Text <> "" Then
'  Me.MousePointer = 11
'  UserName$ = lUsername.Text
'  Password$ = lPassword.Text
'  DatabaseName$ = lDatabaseName
'  Connect$ = UserName$ + "/" + Password$
'  If ISPRUEBA Then
'        ConexionODBC = "DSN=BASA_CONEXION_PRUEBA; UID=" + UserName$ + _
'        "; PWD=" + Password$ + "; DSQ=" + DatabaseName$
'  Else
'        ConexionODBC = "DSN=BASA_CONEXION; UID=" + UserName$ + _
'        "; PWD=" + Password$ + "; DSQ=" + DatabaseName$
'  End If
' On Error GoTo NoOraConnection
' Set conbasa = CreateObject("OracleInProcServer.Xconbasa")
' Set conbasa = conbasa.DbOpenDatabase(DatabaseName$, Connect$, 0&)
' Me.MousePointer = 0
' Unload frmLogin
' frmCargarRequerimientos.Show
' EstadoConexion = True
' Exit Sub
'
'NoOraConnection:
'  Me.MousePointer = 0
'  Unload frmLogin
'  lUsername = ""
'  lPassword = ""
'  EstadoConexion = False
'  frmLogOraError.Show MODAL
'
' End If
End Sub

