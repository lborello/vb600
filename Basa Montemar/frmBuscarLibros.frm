VERSION 5.00
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmCargarLibros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Referncias Libros"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   6870
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2880
      TabIndex        =   7
      Top             =   2460
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5460
      TabIndex        =   6
      Top             =   2460
      Width           =   1200
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   315
      Left            =   780
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
   End
   Begin VB.TextBox txtLibroCliente 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1275
   End
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   6675
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4140
      TabIndex        =   0
      Top             =   2460
      Width           =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Libro:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   180
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmCargarLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Dim Sql As String
    Sql = " Update libros"
    Sql = Sql & " SET REFERENCIA = '" & UCase(txtDescripcion.Text) & "'"
    Sql = Sql & " Where NRO_LIBRO_INTERNO = " & txtLibroCliente.Text
    Sql = Sql & " AND COD_CLIENTE = " & ctlCliente.Valor
    ExecutarSql (Sql)
    txtDescripcion.Text = ""
    txtLibroCliente.Text = ""
End Sub
Private Sub cmdImprimir_Click()
    Dim Sql As String
        Sql = " SELECT * "
        Sql = Sql & " FROM  V_REFERENCIA_LIBROS "
        Sql = Sql & " Where COD_CLIENTE = " & ctlCliente.Valor
        frmReportes.ImprimirReporte PasoReportes & "rptReferenciasLibros.rpt", Sql, True
End Sub

Private Sub Form_Load()
    ctlCliente.TipoControl = Cliente
End Sub
