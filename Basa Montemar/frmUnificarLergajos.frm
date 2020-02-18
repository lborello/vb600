VERSION 5.00
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmUnificarLegajos 
   Caption         =   "Unificar Legajos"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8655
   MDIChild        =   -1  'True
   ScaleHeight     =   3930
   ScaleWidth      =   8655
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Left            =   7380
      TabIndex        =   16
      Top             =   2400
      Width           =   315
   End
   Begin VB.TextBox txtImagenes 
      Height          =   315
      Left            =   1380
      TabIndex        =   15
      Top             =   2400
      Width           =   5955
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   315
      Left            =   7380
      TabIndex        =   13
      Top             =   1920
      Width           =   315
   End
   Begin VB.TextBox txtOrdenHijos 
      Height          =   315
      Left            =   1380
      TabIndex        =   12
      Top             =   1920
      Width           =   5955
   End
   Begin Controles.cltGenerico ctlPersonal 
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   330
      Left            =   4860
      TabIndex        =   9
      Top             =   3300
      Width           =   1200
   End
   Begin VB.CommandButton CmdUnificar 
      Caption         =   "Unificar"
      Height          =   330
      Left            =   6240
      TabIndex        =   8
      Top             =   3300
      Width           =   1200
   End
   Begin VB.TextBox txtLegajosUnificarHijos 
      Height          =   315
      Left            =   1380
      TabIndex        =   5
      Top             =   1440
      Width           =   5955
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   315
      Left            =   7380
      TabIndex        =   4
      Top             =   1440
      Width           =   315
   End
   Begin VB.TextBox TxtLegajosUnificarPadre 
      Height          =   315
      Left            =   1380
      TabIndex        =   2
      Top             =   1020
      Width           =   1575
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
   End
   Begin VB.Label Label6 
      Caption         =   "Imagenes Hijos:"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   1275
   End
   Begin VB.Label Label5 
      Caption         =   "Orden Hijos:"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Responsable:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Legajos Hijos:"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   1020
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Legajos Padre:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "frmUnificarLegajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()
    Dim rs As ADODB.Recordset
    Dim Sql As String
    Sql = "  SELECT ID_CLIENTE_LEGAJO, CLIENTE_LEGAJO, DESCRIPCION,"
    Sql = Sql & vbCrLf & " Nombre , Cod_Estado"
    Sql = Sql & vbCrLf & " From LEGAJOS"
    Sql = Sql & vbCrLf & " WHERE COD_CLIENTE = " & ctlCliente.Valor
    Sql = Sql & vbCrLf & " AND ID_CLIENTE_LEGAJO IN (" & txtLegajosUnificarHijos.Text & ")"
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Sql, ConActiva, 0, 1
    Rem DATOSGRILLA grdLegajosUnificar, rs
    
End Sub

Private Sub CmdUnificar_Click()
Dim Sql As String


If txtLegajosUnificarHijos.Text <> "" Then
    Sql = " Update LEGAJOS"
    Sql = Sql & vbCrLf & " SET  DESCRIPCION ='UNIFICADO:" & TxtLegajosUnificarPadre & "'"
    Sql = Sql & vbCrLf & ", COD_ESTADO =9"
    Sql = Sql & vbCrLf & " Where ID_CLIENTE_LEGAJO in( " & txtLegajosUnificarHijos.Text & ")"
    Sql = Sql & vbCrLf & " And COD_CLIENTE = " & ctlCliente.Valor
    ExecutarSql Sql
End If


If txtImagenes.Text <> "" Then
    Sql = "  Update DOCUMENTOS_DIGITALES"
    Sql = Sql & vbCrLf & " SET   DESCRIPCION ='UNIFICADO LEGAJO:" & TxtLegajosUnificarPadre & "'"
    Sql = Sql & vbCrLf & " , COD_ESTADO =9"
    Sql = Sql & vbCrLf & " Where ID IN (" & txtImagenes.Text & ")"
    ExecutarSql Sql
End If


If txtOrdenHijos.Text <> "" Then
    Sql = " Update basasql.dbo.ORDENAR_DOCUMENTACION_DETALLE"
    Sql = Sql & vbCrLf & " SET  DESCRIPCION ='UNIFICADO LEGAJO:" & TxtLegajosUnificarPadre.Text & "'"
    Sql = Sql & vbCrLf & " Where ID IN (" & txtOrdenHijos.Text & ") "
    ExecutarSql Sql
End If

MsgBox "Legajos Unificado"

End Sub

Private Sub Form_Load()
ctlPersonal.TipoControl = Personal
ctlCliente.TipoControl = Cliente
End Sub
