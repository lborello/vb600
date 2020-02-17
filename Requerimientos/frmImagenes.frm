VERSION 5.00
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmImagenes 
   Caption         =   "Imagenes"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   8805
   Begin Controles.ctlVerImagenes ctlVerImagenes1 
      Height          =   4995
      Left            =   360
      TabIndex        =   9
      Top             =   1260
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   8811
   End
   Begin VB.CommandButton cmdActualizarCliente 
      Caption         =   "..."
      Height          =   315
      Left            =   6660
      TabIndex        =   7
      Top             =   660
      Width           =   375
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   315
      Left            =   2040
      TabIndex        =   6
      Top             =   660
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
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
      Left            =   7740
      TabIndex        =   5
      Top             =   180
      Width           =   1095
   End
   Begin VB.CommandButton cmdProx 
      Caption         =   "Prox."
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
      Left            =   3900
      TabIndex        =   4
      Top             =   180
      Width           =   735
   End
   Begin VB.CommandButton cmdCargarRegistros 
      Caption         =   "Carga"
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
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   675
   End
   Begin Controles.ctlClienteUsuario ctlClienteUsuario 
      Height          =   375
      Left            =   4740
      TabIndex        =   2
      Top             =   180
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
   End
   Begin VB.TextBox txtNumero 
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
      Left            =   2580
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label2 
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
      Left            =   780
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de remito:"
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
      Left            =   900
      TabIndex        =   0
      Top             =   180
      Width           =   1515
   End
End
Attribute VB_Name = "frmImagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
'Private Sub cboImagenes_Click()
'ViewImg1.MostrarImagen PasoImagenes & cboImagenes.ItemData(cboImagenes.ListIndex) & ".tif"
'End Sub

Private Sub cmdActualizar_Click()
Dim sql As String
    If IsNumeric(ctlClienteUsuario.Valor) And IsNumeric(txtNumero.Text) Then
        sql = " Update REMITOS_CUERPO Set COD_USUARIO_CLIENTE = " & ctlClienteUsuario.Valor & " Where NRO_REMITO = " & txtNumero.Text
        ExecutarSql sql
 
    End If
End Sub

Private Sub cmdCargarRegistros_Click()
Dim sql As String
Set rs = New ADODB.Recordset
sql = " SELECT REMITOS_CUERPO.NRO_REMITO,"
sql = sql & vbCrLf & "     REMITOS_CUERPO.ID_CLIENTE AS COD_CLIENTE"
sql = sql & vbCrLf & "  From REMITOS_CUERPO, CONTENEDOR"
sql = sql & vbCrLf & "  Where REMITOS_CUERPO.NRO_REMITO = CONTENEDOR.NRO_REMITO"
 sql = sql & vbCrLf & "     AND (NOT (REMITOS_CUERPO.ID_SQL IS NULL)) AND"
    sql = sql & vbCrLf & " (REMITOS_CUERPO.COD_USUARIO_CLIENTE IS NULL) AND"
    sql = sql & vbCrLf & " (REMITOS_CUERPO.TIPO = 2) AND"
 sql = sql & vbCrLf & "    (CONTENEDOR.ESTADO = 5)AND FECHA > '01/01/2005' "
sql = sql & vbCrLf & " GROUP BY REMITOS_CUERPO.NRO_REMITO,"
sql = sql & vbCrLf & "     REMITOS_CUERPO.ID_CLIENTE"
sql = sql & vbCrLf & "  HAVING (REMITOS_CUERPO.ID_CLIENTE < 200) AND"
sql = sql & vbCrLf & "     (REMITOS_CUERPO.ID_CLIENTE <> 39)"
sql = sql & vbCrLf & "  ORDER BY REMITOS_CUERPO.NRO_REMITO"

rs.Open sql, ConActiva, 0, 1

End Sub

Private Sub cmdProx_Click()
If Not rs.EOF Then
  txtNumero.Text = rs!NRO_REMITO
  ctlClienteUsuario.Clear
  ctlClienteUsuario.LlenarConCliente rs!Cod_cliente
  BuscarRemito txtNumero.Text
  rs.MoveNext
Else
MsgBox "FIN"
End If

End Sub


Private Sub Form_Load()
  ctlCliente.TipoControl = Cliente
End Sub

Private Sub Form_Resize()
If Me.Height > 950 Then
    ctlVerImagenes1.Height = Me.Height - 900
    ctlVerImagenes1.Width = Me.Width - 100
 End If
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If IsNumeric(txtNumero.Text) Then
            BuscarRemito txtNumero.Text
        End If
    End If
End Sub

Public Sub BuscarRemito(Remito As Long)
  Dim sql As String
  Dim i As Integer
  ReDim A(50) As String
  Dim rs As New ADODB.Recordset
        sql = "  SELECT ID_SQL  From IMAGENES "
        sql = sql & " Where TIPO_DOCUMENTO = 2"
        sql = sql & " And Elemento = " & Remito
        rs.Open sql, ConActiva, 0, 1
        i = 0
        Do While Not rs.EOF
            A(i) = PasoImagenes & rs!ID_SQL & ".tif"
            i = i + 1
            rs.MoveNext
        Loop
        If i = 0 Then
            MsgBox "No existen Imagenes", vbInformation
            ctlVerImagenes1.PonerImagen ""
        Else
           ReDim Preserve A(i)
            ctlVerImagenes1.PonerImagen A
        End If
End Sub
