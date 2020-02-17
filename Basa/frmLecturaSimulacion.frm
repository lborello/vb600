VERSION 5.00
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmLecturaSimulacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulacion Lectura"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   9405
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
      Left            =   8040
      TabIndex        =   10
      Top             =   2100
      Width           =   915
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
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   7815
   End
   Begin VB.TextBox txtCajaHasta 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox txtCajaDesde 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   3195
   End
   Begin Controles.cltGenerico CtlPersonal 
      Height          =   375
      Left            =   5700
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
   End
   Begin VB.Label lblLectura 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      Top             =   1560
      Width           =   7815
   End
   Begin VB.Label Label6 
      Caption         =   "Lectura:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1620
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Hasta"
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
      Left            =   4680
      TabIndex        =   6
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Caja/s"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   675
   End
   Begin VB.Label Label2 
      Caption         =   "Personal:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   180
      Width           =   795
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmLecturaSimulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim i As Long
Dim C As Integer

C = 1
Sql = " SELECT MAX(NUMERO_LECTURA) AS MaxLectura From LECTURA_COLECTOR_CUERPO "
rs.Open Sql, ConActiva, 0, 1
Sql = " INSERT INTO LECTURA_COLECTOR_CUERPO     (NUMERO_LECTURA, USUARIO_CREACION,    FECHA_CREACION, DESCRIPCION)"
Sql = Sql & " VALUES (" & rs!MaxLectura + 1 & ",'" & ctlPersonal.Valor & "'," & SysDate & ",'Simulacion  " & txtDescripcion.Text & "')"
ExecutarSql Sql
For i = txtCajaDesde To txtCajaHasta.Text
    Sql = "INSERT INTO LECTURACOLECTOR  (NUMERO_LECTURA, CAJA, CLIENTE, ORDEN)"
    Sql = Sql & " VALUES (" & rs!MaxLectura + 1 & "," & i & "," & ctlCliente.Valor & "," & C & ")"
    C = C + 1
    ExecutarSql Sql
Next
lblLectura.Caption = rs!MaxLectura + 1
End Sub

Private Sub Form_Load()
ctlCliente.TipoControl = Cliente
ctlPersonal.TipoControl = PERSONAL
End Sub
