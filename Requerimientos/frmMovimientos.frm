VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmMovimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos"
   ClientHeight    =   6495
   ClientLeft      =   405
   ClientTop       =   2205
   ClientWidth     =   11745
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
   ScaleHeight     =   6495
   ScaleWidth      =   11745
   Begin MSDataGridLib.DataGrid grdMovimientos 
      Height          =   5175
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      Caption         =   "Movimientos de Elementos"
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
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   330
      Left            =   5280
      TabIndex        =   7
      Top             =   780
      Width           =   1080
   End
   Begin VB.TextBox txtElemento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   780
      Width           =   3615
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   60
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   661
   End
   Begin Controles.cltGenerico ctlElemento 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   420
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   556
   End
   Begin VB.Label Label3 
      Caption         =   "Elemento:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   780
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Elemento:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   420
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   675
   End
End
Attribute VB_Name = "frmMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BuscarElementos()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    rs.CursorLocation = adUseClient
        sql = " SELECT MOVIMIENTOS_ELEMENTOS.COD_REMITO,"
        sql = sql & vbCrLf & " TIPO_REMITO.DESCRIPCION AS TIPO,TIPO_OPERACION.DESCRIPCION AS OPERACION,"
        sql = sql & vbCrLf & " MOVIMIENTOS_ELEMENTOS.Fecha , MOVIMIENTOS_ELEMENTOS.Elemento"
        sql = sql & vbCrLf & " FROM MOVIMIENTOS_ELEMENTOS, TIPO_REMITO,TIPO_OPERACION"
        sql = sql & vbCrLf & " WHERE MOVIMIENTOS_ELEMENTOS.COD_TIPO = TIPO_REMITO.ID AND"
        sql = sql & vbCrLf & " MOVIMIENTOS_ELEMENTOS.COD_OPERACION = TIPO_OPERACION.ID"
        sql = sql & vbCrLf & " AND MOVIMIENTOS_ELEMENTOS.COD_CLIENTE =" & ctlCliente.Valor
        sql = sql & vbCrLf & " AND MOVIMIENTOS_ELEMENTOS.ELEMENTO = " & txtElemento.Text
        sql = sql & vbCrLf & " AND MOVIMIENTOS_ELEMENTOS.COD_TIPO_ALMACENAMIENTO =" & ctlElemento.Valor
        sql = sql & vbCrLf & " AND MOVIMIENTOS_ELEMENTOS.ANULADO IS NULL "
        sql = sql & vbCrLf & " ORDER BY MOVIMIENTOS_ELEMENTOS.FECHA"
        
    sql = "     SELECT     REMITOS_CUERPO.NRO_REMITO as Remtio, REMITOS_CUERPO.NRO_REM_PROV, TIPO_REMITO.DESCRIPCION AS TIPO,"
      sql = sql & vbCrLf & " TIPO_OPERACION.DESCRIPCION AS Doperacion, REMITOS_CUERPO.FECHA"
 sql = sql & vbCrLf & " FROM         REMITOS_CUERPO INNER JOIN"
                       sql = sql & vbCrLf & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO LEFT OUTER JOIN"
                       sql = sql & vbCrLf & " TIPO_OPERACION ON REMITOS_CUERPO.OPERACION = TIPO_OPERACION.ID LEFT OUTER JOIN"
                       sql = sql & vbCrLf & " TIPO_REMITO ON REMITOS_CUERPO.TIPO = TIPO_REMITO.ID"
 sql = sql & vbCrLf & " WHERE     REMITOS_CUERPO.ID_CLIENTE =  " & ctlCliente.Valor
  sql = sql & vbCrLf & " AND (REMITOS_CUERPO.ANULADO IS NULL) "
  sql = sql & vbCrLf & " AND  COD_TIPO_ALMACENAMIENTO = " & ctlElemento.Valor
   sql = sql & vbCrLf & " AND " & txtElemento.Text & "  BETWEEN REMITOS_DETALLE.DESDE AND REMITOS_DETALLE.HASTA"
 sql = sql & vbCrLf & " ORDER BY REMITOS_CUERPO.NRO_REMITO"
        
        rs.Open sql, ConActiva, 0, 1
        Set grdMovimientos.DataSource = rs.DataSource
        grdMovimientos.Refresh
End Sub

Private Sub cmdBuscar_Click()
     If IsNull(ctlCliente.Valor) Then
        MsgBox "Ingrese el Cliente", vbInformation
        Exit Sub
    End If
    If IsNull(ctlElemento.Valor) Then
        MsgBox "Ingrese el Tipo de elemento", vbInformation
        Exit Sub
    End If
    If Not IsNumeric(txtElemento.Text) Then
        MsgBox "El elemento no es correcto ", vbInformation
        Exit Sub
    End If
    BuscarElementos
End Sub

Private Sub Form_Load()
    ctlCliente.TipoControl = Cliente
    ctlElemento.TipoControl = Tipo_Remito_almacenamiento
End Sub

