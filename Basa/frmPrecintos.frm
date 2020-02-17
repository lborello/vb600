VERSION 5.00
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmPrecintos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Precintos"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin Controles.cltGenerico ctlPersonal 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
   End
   Begin VB.TextBox txtPrecinto_2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2820
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtPrecinto_1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtCaja 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   540
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
   End
   Begin VB.Label Label4 
      Caption         =   "Precintos:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Caja :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Personal:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   540
      Width           =   735
   End
End
Attribute VB_Name = "frmPrecintos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub cmdAceptar_Click()
Dim Sql As String
If IsNumeric(txtCaja.Text) Then
        ExecutarSql "Update PRECINTOS Set ACTUAL = 0 Where COD_CLIENTE = " & ctlCliente.Valor & " And NRO_CAJA = " & txtCaja.Text
        Sql = " INSERT INTO PRECINTOS(COD_CLIENTE, NRO_CAJA, ACTUAL, PRECINTOS_1,"
        Sql = Sql & vbCrLf & " PRECINTOS_2, FECHA, ID_PERSONAL)"
        Sql = Sql & " VALUES (" & ctlCliente.Valor & "," & txtCaja.Text & ", '1','" & txtPrecinto_1.Text & " ', '" & txtPrecinto_2.Text & "' ," & SysDate & "," & ctlPersonal.Valor & ")"
        ExecutarSql Sql
        txtCaja.Text = ""
        txtPrecinto_1.Text = ""
        txtPrecinto_2.Text = ""
        Rem txtCaja.SetFocus
End If
End Sub

Private Sub Form_Load()
ctlPersonal.TipoControl = Personal
ctlCliente.TipoControl = Cliente
End Sub

Private Sub txtCaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub txtPrecinto_1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtPrecinto_2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

