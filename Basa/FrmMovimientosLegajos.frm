VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form FrmMovimientosLegajos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analisis de Movimientos legajos"
   ClientHeight    =   4455
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   11040
   Begin MSDataGridLib.DataGrid grdMovimientosCajas 
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLegajo 
      BackColor       =   &H00FFFFC0&
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
      Left            =   7140
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
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
      Left            =   8340
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   661
   End
   Begin VB.Label Label2 
      Caption         =   "ETIQUETA:"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   180
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "CLIENTE:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   795
   End
   Begin VB.Menu mnuRemito 
      Caption         =   "Remito"
      Visible         =   0   'False
      Begin VB.Menu mnuImagenRemito 
         Caption         =   "ImagenRemito"
      End
      Begin VB.Menu mnuCopiarDatos 
         Caption         =   "Copiar datos"
      End
   End
End
Attribute VB_Name = "FrmMovimientosLegajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Sub BuscarMovimientos(id_cliente As Integer, NRO_CAJA As Double, tipo_almacenamiento As Integer, ClienteRazon As String, estado As String)
    Dim OraMovimientos As New ADODB.Recordset
    Dim Sql As String
    Dim i As Integer
MousePointer = 11
   
   If tipo_almacenamiento = 0 Then
    lblCajaLibro.Caption = "Caja"
   Else
    lblCajaLibro.Caption = "Libro"
   End If
   LblCaja = NRO_CAJA
   lblCliente = ClienteRazon
   Select Case estado
   Case 2
        lblEstado.Caption = "En Planta"
   Case 3
        lblEstado.Caption = "Consulta"
   Case 4
        lblEstado.Caption = "Reservada"
   End Select
   
   GrdMovimientos.Clear
   For i = 2 To GrdMovimientos.Rows - 1
      GrdMovimientos.RemoveItem (2)
   Next
    GrdMovimientos.ColWidth(0) = 200
    GrdMovimientos.ColWidth(1) = 1000
    GrdMovimientos.ColWidth(2) = 1200
    GrdMovimientos.ColWidth(3) = 1800
    GrdMovimientos.ColWidth(4) = 1200
    GrdMovimientos.ColWidth(5) = 1200
    
    GrdMovimientos.TextMatrix(0, 1) = "Nº Remito"
    GrdMovimientos.TextMatrix(0, 2) = "Fecha"
    GrdMovimientos.TextMatrix(0, 3) = "Tipo"
    GrdMovimientos.TextMatrix(0, 4) = "Movimiento"
    GrdMovimientos.TextMatrix(0, 5) = "Nº R. Prov."



        
        Sql = "SELECT REMITOS_DETALLE.NRO_REMITO, REMITOS_DETALLE.NRO_CAJA, REMITOS_CUERPO.NRO_REM_PROV ,"
        Sql = Sql & vbCrLf & " REMITOS_DETALLE.DESDE, REMITOS_DETALLE.HASTA, REMITOS_DETALLE.TIPO_ALMACENADO,"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.ANULADO, REMITOS_CUERPO.FECHA, REMITO_TIPO.DESCRIPCION as TIPO_DESCRIPCION,"
        Sql = Sql & vbCrLf & " REMITO_OPERACION.Descripcion as OPERACION_Descripcion , REMITOS_CUERPO.Estado, REMITOS_CUERPO.id_cliente"
        Sql = Sql & vbCrLf & " From REMITOS_DETALLE, REMITOS_CUERPO, REMITO_TIPO, REMITO_OPERACION"
        Sql = Sql & vbCrLf & " WHERE ( (REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO) AND"
        Sql = Sql & vbCrLf & " (REMITOS_DETALLE.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO) AND"
        Sql = Sql & vbCrLf & " (REMITO_TIPO.ID = REMITOS_CUERPO.TIPO) AND"
        Sql = Sql & vbCrLf & " (REMITO_OPERACION.ID = REMITOS_CUERPO.OPERACION) AND"
        Sql = Sql & vbCrLf & " (REMITOS_CUERPO.ANULADO IS NULL) AND "
        Sql = Sql & vbCrLf & " (REMITOS_DETALLE.TIPO_ALMACENADO = " & tipo_almacenamiento & ") AND "
        Sql = Sql & vbCrLf & " (REMITOS_CUERPO.ID_CLIENTE =" & id_cliente & ") )"
        Sql = Sql & vbCrLf & " ORDER BY REMITOS_CUERPO.FECHA ASC,REMITOS_CUERPO.NRO_REMITO ASC"
        OraMovimientos.Open Sql, ConActiva, 0, 1
         i = 1
        Do While Not OraMovimientos.EOF
             If NRO_CAJA >= Val(OraMovimientos!Desde) And NRO_CAJA <= Val(OraMovimientos!Hasta) Then
                    If i <> 1 Then
                      GrdMovimientos.AddItem ""
                    End If
                    GrdMovimientos.TextMatrix(i, 1) = Trim(OraMovimientos!NRO_REMITO)
                    GrdMovimientos.TextMatrix(i, 2) = Trim(OraMovimientos!fecha)
                    GrdMovimientos.TextMatrix(i, 3) = Trim(OraMovimientos!TIPO_DESCRIPCION)
                    GrdMovimientos.TextMatrix(i, 4) = Trim(OraMovimientos!OPERACION_Descripcion)
                    GrdMovimientos.TextMatrix(i, 5) = IIf(IsNull(Trim(OraMovimientos!NRO_REM_PROV)), 0, Trim(OraMovimientos!NRO_REM_PROV))
                    i = i + 1
                Else
            End If
        OraMovimientos.MoveNext
        Loop
        
        MousePointer = 0
End Sub



Private Sub cmdBuscar_Click()
Dim rse As New ADODB.Recordset
rse.CursorLocation = adUseClient
Dim Sql As String
   
    Sql = " SELECT REMITOS_CUERPO.NRO_REMITO AS REMITO,"
    Sql = Sql & vbCrLf & " REMITOS_CUERPO.NRO_REM_PROV AS PROV,"
    Sql = Sql & vbCrLf & " REMITOS_CUERPO.FECHA, REMITOS_DETALLE.DESDE AS ETIQUETA,"
    Sql = Sql & vbCrLf & " REMITO_TIPO.DESCRIPCION AS TIPO,"
    Sql = Sql & vbCrLf & " REMITO_OPERACION.DESCRIPCION AS OPERACION"
    Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO, REMITOS_DETALLE, REMITO_TIPO,REMITO_OPERACION"
    Sql = Sql & vbCrLf & " Where REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO"
    Sql = Sql & vbCrLf & " AND REMITOS_CUERPO.TIPO = REMITO_TIPO.ID AND"
    Sql = Sql & vbCrLf & " REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID AND"
    Sql = Sql & vbCrLf & " (REMITOS_CUERPO.ID_CLIENTE = " & ctlCliente.Valor & ") AND"
    Sql = Sql & vbCrLf & " (REMITOS_DETALLE.TIPO_ALMACENADO in(3,2)) AND"
    Sql = Sql & vbCrLf & " (REMITOS_DETALLE.DESDE =" & txtLegajo.Text & " ) AND"
    Sql = Sql & vbCrLf & " (REMITOS_CUERPO.ANULADO IS NULL)"
    Sql = Sql & vbCrLf & " ORDER BY REMITOS_CUERPO.NRO_REMITO "
    
    rse.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
    DATOSGRILLA grdMovimientosCajas, rse

End Sub


Private Sub Form_Load()
ctlCliente.TipoControl = Cliente
End Sub

Private Sub grdMovimientosCajas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'If Button = 2 Then
'        If Dir("\\Server1basa\fax\remitos\" & "REMITO " & grdMovimientosCajas.Text & ".Tif") = "" Then
'            mnuImagenRemito.Enabled = False
'        Else
'            mnuImagenRemito.Enabled = True
'        End If
'        PopupMenu mnuRemito
'End If
End Sub

Private Sub mnuCopiarDatos_Click()
 CopiarDatosGrilla grdMovimientosCajas
End Sub

Private Sub mnuImagenRemito_Click()
    frmVerfax.Show
        frmVerfax.PonerImagen "\\Server1basa\fax\remitos\" & "REMITO " & grdMovimientosCajas.Text & ".Tif"
End Sub
