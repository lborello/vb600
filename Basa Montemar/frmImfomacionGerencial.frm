VERSION 5.00
Object = "{A30B2DDF-C00F-469F-A23C-D6177608A128}#10.5#0"; "crviewer.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmImfomacionGerencial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informacion Gerencial"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   10890
   Begin VB.CommandButton cmdCopiarExcel 
      Caption         =   "Copiar Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6600
      TabIndex        =   14
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cmdRecalcularCaracteres 
      Caption         =   "Re Calcular Caracteres"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4320
      TabIndex        =   13
      Top             =   7920
      Width           =   2115
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "Informe"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   8760
      TabIndex        =   12
      Top             =   7920
      Width           =   1935
   End
   Begin VB.CommandButton cmdFactura 
      Caption         =   "facturas Sin Marcar"
      Height          =   390
      Left            =   120
      TabIndex        =   11
      Top             =   7920
      Width           =   1635
   End
   Begin VB.CommandButton cmdFacturasCustodia 
      Caption         =   "Facturas Custodia"
      Height          =   390
      Left            =   2280
      TabIndex        =   10
      Top             =   7920
      Width           =   1875
   End
   Begin CrystalActiveXReportViewerLib105Ctl.CrystalActiveXReportViewer CrystalActiveXReportViewer1 
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   8040
      Width           =   255
      _cx             =   450
      _cy             =   661
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.ComboBox cboTipoReferencia 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmImfomacionGerencial.frx":0000
      Left            =   1920
      List            =   "frmImfomacionGerencial.frx":0002
      TabIndex        =   7
      Top             =   1140
      Width           =   8775
   End
   Begin VB.ComboBox cboTipoInforme 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmImfomacionGerencial.frx":0004
      Left            =   1920
      List            =   "frmImfomacionGerencial.frx":00CE
      TabIndex        =   6
      Text            =   "cboTipoInforme"
      Top             =   120
      Width           =   8895
   End
   Begin VB.TextBox txtFechaHasta 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6000
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.TextBox txtFechaDesde 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid grdCargaLegajos 
      Height          =   6075
      Left            =   120
      TabIndex        =   0
      Top             =   1620
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   10716
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
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
         Size            =   9
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin VB.Label Label4 
      Caption         =   "Tipo de referencia:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1140
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha Hasta:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4380
      TabIndex        =   5
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Desde :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Informe:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmImfomacionGerencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim directorio As String
Private Sub cmdInformeGerenciaFacturacion_Click()
'Dim Ssql As String
'
'        Ssql = " SELECT * "
'
'       Ssql = "  SELECT CLIENTES.RAZON_SOCIAL, FACTURA.DESCRIPCION,"
'   Ssql = Ssql & vbCrLf & " FACTURA.CANTIDAD_ELEMENTO, FACTURA.MONTO,FACTURA.Lote"
'Ssql = Ssql & vbCrLf & " From FACTURA, CLIENTES"
'Ssql = Ssql & vbCrLf & " WHERE FACTURA.COD_CLIENTE = CLIENTES.ID_CLIENTE AND (FACTURA.LOTE = 88)"
'Ssql = Ssql & vbCrLf & " ORDER BY FACTURA.COD_CLIENTE"
'      Form4.SqlReporte = Ssql
'    Form4.Strre = "InformeGerenciaFacturacion"
'    Form4.Show
End Sub

Private Sub cmdCopiarExcel_Click()
CopiarDatosGrilla grdCargaLegajos
End Sub

Private Sub cmdImformeLegajosReducido_Click()

  Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
Dim PERSONALOLD As Integer



End Sub

Private Sub cmdFactura_Click()
        Dim Sql As String
        Dim conData As New ADODB.Connection
        Dim RsFactura As New ADODB.Recordset
        Dim rsTEM_IVA_DATA As New ADODB.Recordset
            conData.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=datas"
            Sql = "SELECT FacturaABC, NumeroFactura, MesFacturacion, AnoFacturacion  "
            Sql = Sql & " , Impresa ,FechaFacturacion  "
            Sql = Sql & " From factura  "
            Sql = Sql & " Where Impresa <> 'S'   "
            Sql = Sql & " and ( NumeroFactura > 10100000 and NumeroFactura < 20000000 ) "
            Sql = Sql & " ORDER BY NumeroFactura , FechaFacturacion "
            
         Rem   filtyro por fecha FechaFacturacion
            
            
            
            
            
'            Sql = "   SELECT CLIENTE.IDCLIENTE, CLIENTE.NOMBRE, FACTURA.FechaFacturacion, FACTURA.FacturaABC, FACTURA.NumeroFactura, FACDET.CANTIDAD, FACDET.PRECIOUNITARIO, FACDET.PRECIOTOTAL"
'            Sql = Sql & " FROM (CLIENTE INNER JOIN FACTURA ON CLIENTE.[IDCLIENTE] = FACTURA.[IDCliente]) INNER JOIN FACDET ON FACTURA.[NumeroFactura] = FACDET.[NUMEROFACTURA]"
'            Sql = Sql & " WHERE (((CLIENTE.IDCLIENTE) In (5003,5401,5403,5403,5001,123,156)) AND ((FACDET.DETALLE) Like '*ima*'))"
'            Sql = Sql & " ORDER BY CLIENTE.IDCLIENTE;"
            
            Set RsFactura = New ADODB.Recordset
            RsFactura.CursorLocation = adUseClient
            RsFactura.Open Sql, conData
            Set grdCargaLegajos.DataSource = RsFactura.DataSource

        If MsgBox(" Vos Jose quieres marcar las facturas por fecha", vbYesNo) = vbYes Then
            Sql = " UPDATE factura SET Impresa = 'S'"
            Sql = Sql & " Where Impresa <> 'S'"
            Sql = Sql & " and ( NumeroFactura > 10100000 and NumeroFactura < 20000000 ) "
            Sql = Sql & " and  FechaFacturacion = " & InputBox("Ingrese la FechaFacturacion como esta en la grilla")
            conData.Execute Sql
        End If
            
End Sub

Private Sub cmdFacturasCustodia_Click()
 
        Dim Sql As String
        Dim conData As New ADODB.Connection
        Dim RsFactura As New ADODB.Recordset
        Dim rsTEM_IVA_DATA As New ADODB.Recordset
            conData.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=datas"
            Sql = "SELECT FacturaABC, NumeroFactura, MesFacturacion, AnoFacturacion  "
            Sql = Sql & " , Impresa ,FechaFacturacion  "
            
            Sql = "SELECT *"
            Sql = Sql & " From factura  "
            Sql = Sql & " Where Impresa <> 'S'   "
            Sql = Sql & " and ( NumeroFactura > 10100000 and NumeroFactura < 20000000 ) "
            Sql = Sql & " ORDER BY NumeroFactura , FechaFacturacion "

            
            Set RsFactura = New ADODB.Recordset
            RsFactura.CursorLocation = adUseClient
            RsFactura.Open Sql, conData
            Set grdCargaLegajos.DataSource = RsFactura.DataSource

            conData.Execute Sql
            

End Sub

Private Sub cmdInforme_Click()
  Dim rs As New ADODB.Recordset
  Dim P As Integer
  
  Dim Usuario As Integer
  
On Error GoTo salir

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
Dim PERSONALOLD As Integer
Dim conRepo As New ADODB.Connection

        Dim Sql As String
        
        Rem ConBasa.CommandTimeout = 6000
        Set conRepo = New ADODB.Connection
       
        conRepo.Open strConBasa
        
        
        
'        sql de regional legajos
'
'        SELECT     LEGAJOS.NRO_CAJA, INDICES.DESCRIPCION, INDICES_1.DESCRIPCION AS Expr1, COUNT(*) AS Expr2
'FROM         LEGAJOS INNER JOIN
'                      INDICES ON LEGAJOS.COD_INDICE = INDICES.INDICE AND LEGAJOS.COD_CLIENTE = INDICES.COD_CLIENTE INNER JOIN
'                      INDICES INDICES_1 ON SUBSTRING(LEGAJOS.COD_INDICE, 1, 6) = INDICES_1.INDICE AND
'                      LEGAJOS.COD_CLIENTE = INDICES_1.COD_CLIENTE
'Where (LEGAJOS.COD_CLIENTE = 172)
'GROUPsql BY LEGAJOS.NRO_CAJA, INDICES.DESCRIPCION, INDICES_1.DESCRIPCION
'ORDER BY INDICES_1.DESCRIPCION, INDICES.DESCRIPCION
'
        

Select Case cboTipoInforme.ListIndex

Case 0 ' informe reducido carga
    Sql = " SELECT     CONVERT(char, LEGAJOS.FECHA_CREACION, 103) AS FECHACARGA, LEGAJOS.FK_PERSONAL_CREACION AS PERSONAL, COUNT(*) AS CANTIDADLEGAJOS,SUM(CANTIDAD_CARACTERES) AS CARACTERES,"
    Sql = Sql & "  Personal.Nombre , Personal.Apellido, PERSONAL.CARGA_HORARIA"
    Sql = Sql & " FROM         LEGAJOS INNER JOIN"
    Sql = Sql & " PERSONAL ON LEGAJOS.FK_PERSONAL_CREACION = PERSONAL.IDPERSONAL"
    Sql = Sql & " WHERE   (CONVERT(DATETIME, LEGAJOS.FECHA_CREACION, 103) BETWEEN  " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text) & ")"
    Sql = Sql & " GROUP BY CONVERT(char, LEGAJOS.FECHA_CREACION, 103), LEGAJOS.FK_PERSONAL_CREACION, PERSONAL.NOMBRE, PERSONAL.APELLIDO,PERSONAL.CARGA_HORARIA"
    Sql = Sql & " ORDER BY LEGAJOS.FK_PERSONAL_CREACION, CONVERT(char, LEGAJOS.FECHA_CREACION, 103)"

Case 1 'Informe  detallado carga
    Sql = " SELECT     CONVERT(char, LEGAJOS.FECHA_CREACION, 103) AS FECHACARGA, LEGAJOS.FK_PERSONAL_CREACION AS Personal, LEGAJOS.COD_CLIENTE,"
    Sql = Sql & " LEGAJOS.NRO_CAJA, COUNT(*) AS CANTIDADLEGAJOS, SUM(CANTIDAD_CARACTERES) AS CARACTERES , PERSONAL.NOMBRE, PERSONAL.APELLIDO"
    Sql = Sql & " FROM         LEGAJOS INNER JOIN"
    Sql = Sql & "  PERSONAL ON LEGAJOS.FK_PERSONAL_CREACION = PERSONAL.IDPERSONAL"
    Sql = Sql & " WHERE   (CONVERT(DATETIME, LEGAJOS.FECHA_CREACION, 103) BETWEEN  " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text) & ")"
    Sql = Sql & "  GROUP BY CONVERT(char, LEGAJOS.FECHA_CREACION, 103), LEGAJOS.FK_PERSONAL_CREACION, LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA,"
    Sql = Sql & "  Personal.Nombre , Personal.Apellido"
    Rem Sql = Sql & "  ORDER BY CONVERT(char, LEGAJOS.FECHA_CREACION, 103), LEGAJOS.FK_PERSONAL_CREACION, LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA"
    Sql = Sql & "  ORDER BY  LEGAJOS.FK_PERSONAL_CREACION,CONVERT(char, LEGAJOS.FECHA_CREACION, 103), LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA"
 Case 2
    Sql = "  SELECT     LEGAJOS.NRO_CAJA, INDICES.DESCRIPCION, COUNT(*) AS CANTIDADLEGAJOS, SUM(LEGAJOS.CANTIDAD_CARACTERES) AS CARACTERES"
    Sql = Sql & " FROM         LEGAJOS INNER JOIN"
    Sql = Sql & "  PERSONAL ON LEGAJOS.FK_PERSONAL_CREACION = PERSONAL.IDPERSONAL INNER JOIN"
    Sql = Sql & " INDICES ON LEGAJOS.FK_INDICES = INDICES.ID"
    Sql = Sql & " WHERE (CONVERT(DATETIME, LEGAJOS.FECHA_CREACION, 103) BETWEEN " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text) & ") "
    Sql = Sql & " AND  LEGAJOS.COD_CLIENTE = " & InputBox("Ingrese en numero de cliente")
    Sql = Sql & " GROUP BY  INDICES.DESCRIPCION, LEGAJOS.NRO_CAJA"
    Sql = Sql & "  ORDER BY INDICES.DESCRIPCION"
 Case 3  ' requerimientos
    Sql = "  SELECT     dbo.REQUERIMIENTO.IDREQUERIMIENTO, dbo.REQUERIMIENTO.IDREMITO, dbo.REQUERIMIENTO.ID_CLIENTE, CONVERT(datetime,"
    Sql = Sql & " dbo.REQUERIMIENTO.FECHARECEPCION, 102) AS fecha, dbo.TIPOREQUERIMIENTO.DESCRIPCION, dbo.REQUERIMIENTO.CANTIDAD,"
    Sql = Sql & " dbo.REQUERIMIENTO.CANTIDAD_IMAGENES, dbo.CLIENTEUSUARIO.APELLIDO_NOMBRE, dbo.INDICES.TituloHerencia "
    Sql = Sql & " FROM dbo.TIPOREQUERIMIENTO INNER JOIN"
    Sql = Sql & "  dbo.REQUERIMIENTO ON dbo.TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO = dbo.REQUERIMIENTO.IDTIPOREQUERIMIENTO LEFT OUTER JOIN"
    Sql = Sql & "  dbo.INDICES INNER JOIN"
    Sql = Sql & "  dbo.CLIENTEUSUARIO ON dbo.INDICES.INDICE = dbo.CLIENTEUSUARIO.COD_INDICE AND "
    Sql = Sql & "  dbo.INDICES.COD_CLIENTE = dbo.CLIENTEUSUARIO.COD_CLIENTE ON "
    Sql = Sql & "  dbo.REQUERIMIENTO.COD_USUARIO_CLIENTE = dbo.CLIENTEUSUARIO.ID_CLIENTEUSUARIO "
    Sql = Sql & "  WHERE  dbo.REQUERIMIENTO.ID_CLIENTE = " & InputBox("Ingrese el cliente", "Cliente", "231")
    Sql = Sql & "  AND (dbo.REQUERIMIENTO.ANULADO IS NULL) AND"
    Sql = Sql & " CONVERT(datetime, dbo.REQUERIMIENTO.FECHARECEPCION, 102)  BETWEEN " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text)
    Sql = Sql & "  ORDER BY INDICES.INDICE , dbo.INDICES.DESCRIPCION"
Case 4 ' cajas con legajos por cargar
    Sql = " SELECT CAJAS.FK_PERSONAL_ASIGNACION_TIPO AS PER_ASIG, CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, COUNT(*) AS CANT_LEGAJOS, CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL,"
    Sql = Sql & vbCrLf & " CONTENEDOR.VERTICAL, CONTENEDOR.ADELANTE_ATRAS, CONTENEDOR.UB_PROVISORIA, CLIENTES.RAZON_SOCIAL,"
    Sql = Sql & vbCrLf & " PARAMETROS.Descripcion"
    Sql = Sql & vbCrLf & " FROM         CAJAS INNER JOIN"
    Sql = Sql & vbCrLf & " CONTENEDOR ON CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA AND CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE INNER JOIN"
    Sql = Sql & vbCrLf & " CLIENTES ON CAJAS.FK_CLIENTE = CLIENTES.ID_CLIENTE INNER JOIN"
    Sql = Sql & vbCrLf & " PARAMETROS ON CAJAS.FK_TIPO_REFERENCIA = PARAMETROS.ID_PARAMETRO LEFT OUTER JOIN"
    Sql = Sql & vbCrLf & " LEGAJOS ON CAJAS.FK_CLIENTE = LEGAJOS.COD_CLIENTE AND CAJAS.NRO_CAJA = LEGAJOS.NRO_CAJA"
    Sql = Sql & vbCrLf & " WHERE     (CAJAS.FK_TIPO_REFERENCIA IN (1010, 1040, 1060))"
    Sql = Sql & vbCrLf & " GROUP BY CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL,"
    Sql = Sql & vbCrLf & " CONTENEDOR.Adelante_Atras , CONTENEDOR.UB_PROVISORIA, Clientes.RAZON_SOCIAL, PARAMETROS.Descripcion ,  CAJAS.FK_PERSONAL_ASIGNACION_TIPO  "
    Sql = Sql & vbCrLf & " ORDER BY CAJAS.FK_CLIENTE"
Case 5 ' control logico
    Sql = "  SELECT     LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_DESDE, LEGAJOS.LETRA_HASTA, PERSONAL.IDPERSONAL, PERSONAL.NOMBRE, PERSONAL.APELLIDO,"
    Sql = Sql & " LEGAJOS.NRO_CAJA , LEGAJOS.ID_LEGAJO, LEGAJOS.FECHA_CREACION,REGISTRO_VERIFICADO as VERIFICADO"
    Sql = Sql & "  FROM         LEGAJOS INNER JOIN"
    Sql = Sql & "  PERSONAL ON LEGAJOS.FK_PERSONAL_CREACION = PERSONAL.IDPERSONAL"
    Sql = Sql & "  WHERE     (LEGAJOS.COD_CLIENTE =" & InputBox("Ingrese el cliente", "Cliente", "231") & ") AND NOT (LEN(NRO_DESDE) IN (" & InputBox("Ingrese la longitud de números exceptuados separados por , ", "Longitud", "1, 8, 7") & "))"
    Sql = Sql & " AND  (CONVERT(DATETIME, LEGAJOS.FECHA_CREACION, 103) BETWEEN " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text) & ") "
    Sql = Sql & "  ORDER BY PERSONAL.IDPERSONAL, LEGAJOS.NRO_CAJA, LEGAJOS.ID_LEGAJO"
Case 6 ' remitos
    Sql = "   SELECT     dbo.REMITOS_CUERPO.ID_CLIENTE, dbo.REMITOS_CUERPO.TIPO, dbo.REMITOS_CUERPO.FECHA, dbo.REMITOS_CUERPO.CANTIDAD,"
    Sql = Sql & " dbo.INDICES.TituloHerencia , dbo.REMITOS_CUERPO.NRO_REMITO, dbo.TIPO_REMITO.ID, dbo.TIPO_REMITO.DESCRIPCION"
    Sql = Sql & " FROM dbo.TIPO_REMITO INNER JOIN"
    Sql = Sql & " dbo.REMITOS_CUERPO ON dbo.TIPO_REMITO.ID = dbo.REMITOS_CUERPO.TIPO LEFT OUTER JOIN"
    Sql = Sql & " dbo.INDICES INNER JOIN"
    Sql = Sql & " dbo.CLIENTEUSUARIO ON dbo.INDICES.INDICE = dbo.CLIENTEUSUARIO.COD_INDICE AND"
    Sql = Sql & " dbo.INDICES.COD_CLIENTE = dbo.CLIENTEUSUARIO.COD_CLIENTE ON"
    Sql = Sql & " dbo.REMITOS_CUERPO.COD_USUARIO_CLIENTE = dbo.CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
    Sql = Sql & " WHERE     dbo.REMITOS_CUERPO.ID_CLIENTE =" & InputBox("Ingrese el cliente", "Cliente", "231")
    Sql = Sql & " AND (TIPO IN (0, 2, 3, 4) OR (TIPO = 1) AND (OPERACION = 1))  "
    Sql = Sql & " AND (dbo.REMITOS_CUERPO.FECHA BETWEEN " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text) & ")"
    Sql = Sql & " ORDER BY dbo.INDICES.DESCRIPCION"
Case 7 ' Error de carga de legajos
    Sql = "   SELECT     CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, LEGAJOS.FK_PERSONAL_CREACION, CONVERT(CHAR, LEGAJOS.FECHA_CREACION, 103) AS FECHA, COUNT(*) AS CANTIDAD,"
    Sql = Sql & "                   Clientes.RAZON_SOCIAL "
    Sql = Sql & " FROM         CAJAS INNER JOIN"
    Sql = Sql & "                     CLIENTES ON CAJAS.FK_CLIENTE = CLIENTES.ID_CLIENTE LEFT OUTER JOIN"
    Sql = Sql & " LEGAJOS ON CAJAS.NRO_CAJA = LEGAJOS.NRO_CAJA AND CAJAS.FK_CLIENTE = LEGAJOS.COD_CLIENTE"
    Sql = Sql & " Where (Not (LEGAJOS.NRO_CAJA Is Null))"
    Sql = Sql & " and      LEGAJOS.FECHA_CREACION BETWEEN " & FechaFormato(txtFechaDesde.Text) & "  AND " & FechaFormato(txtFechaHasta.Text)
    Sql = Sql & " GROUP BY CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, LEGAJOS.FK_PERSONAL_CREACION, CONVERT(CHAR, LEGAJOS.FECHA_CREACION, 103), CLIENTES.RAZON_SOCIAL"
    Sql = Sql & " ORDER BY LEGAJOS.FK_PERSONAL_CREACION"
Case 8  ' CARGA DE REFERENCIAS
    Sql = "SELECT     dbo.REFERENCIAS.COD_CLIENTE, CONVERT(char(10), dbo.REFERENCIAS.FECHA_MODIFICACION, 103) AS FECHA, COUNT(*) AS Cantidad,"
    Sql = Sql & "                      dbo.Personal.Nombre , dbo.Personal.Apellido"
    Sql = Sql & "  FROM         dbo.REFERENCIAS INNER JOIN"
    Sql = Sql & "  dbo.PERSONAL ON dbo.REFERENCIAS.FK_PERSONAL_MODIFICACION = dbo.PERSONAL.IDPERSONAL"
    Sql = Sql & "  WHERE     dbo.REFERENCIAS.FECHA_MODIFICACION >  " & FechaFormato(txtFechaDesde.Text)
    Sql = Sql & "  GROUP BY CONVERT(char(10), dbo.REFERENCIAS.FECHA_MODIFICACION, 103), dbo.REFERENCIAS.COD_CLIENTE,"
    Sql = Sql & "  dbo.REFERENCIAS.FK_PERSONAL_MODIFICACION , dbo.Personal.Nombre, dbo.Personal.Apellido"
    Sql = Sql & "  Having (Not (dbo.REFERENCIAS.FK_PERSONAL_MODIFICACION Is Null))"
Case 9 ' detalle legajos
    Sql = " SELECT     ID_LEGAJO, COD_CLIENTE, NRO_CAJA, FK_PERSONAL_ACTUALIZACION, FECHA_ACTUALIZACION, CANTIDAD_CARACTERES"
    Sql = Sql & " From dbo.LEGAJOS"
   Sql = Sql & "  WHERE   CONVERT (date, FECHA_ACTUALIZACION) between  '" & txtFechaDesde.Text & "' and  '" & txtFechaHasta.Text & "'"
    Sql = Sql & " ORDER BY FK_PERSONAL_ACTUALIZACION, FECHA_ACTUALIZACION"
    
    
    
Case 10  'detalle referencias
    Sql = " SELECT     COD_CLIENTE, NRO_CAJA, FK_PERSONAL_MODIFICACION, FECHA_MODIFICACION, FECHA_DESDE, NRO_DESDE"
    Sql = Sql & "  From dbo.REFERENCIAS"
    Sql = Sql & "  WHERE    FECHA_MODIFICACION >  " & FechaFormato(txtFechaDesde.Text)
    Sql = Sql & "  ORDER BY FK_PERSONAL_MODIFICACION, FECHA_MODIFICACION"
Case 11
    ControlCajas
    Sql = " SELECT  *   "
    Sql = Sql & "   From TEM_CONTROL_CAJAS"
    Sql = Sql & "   ORDER BY FK_LECTURA, ORDEN"
    
Case 12
    InformeSupervielle
    
    Sql = "SELECT FORMA, REQUERIMIENTO, NRO_REMITO, NRO_REM_PROV, ELEMENTO, TIPO, SUBTIPO, FECHA, CANTIDADCONSULTAS, CANTIDADVACIAS,"
    Sql = Sql & vbCrLf & " CANTIDADGUARDAYCUSTODIA, CANTIDADLEGAJOS, CANTIDADFLETESNORMALES, RETIROS, CANTIDADFLETESURGENTES, CANTIDADIMAGENES,"
    Sql = Sql & vbCrLf & " RETIROSFUERADERADIO, ENVIOFUERADERADIO, CANTIDADHORASARCHIVISTA, PRECINTOS, APELLIDO_NOMBRE, PROVINCIA, SUCURSAL, COBRAR, FK_CLIENTE,"
    Sql = Sql & vbCrLf & " PASO_IMAGEN"
    Sql = Sql & vbCrLf & " From TEM_SUPERVIELLE"
    
    
    
    
    
    
    
    
    
    
    
    Sql = Sql & vbCrLf & " ORDER BY FORMA, TIPO, SUBTIPO"
    
    
    
    
    BUSCAR_REMITOS_NUEVO
    
    BUSCAR_Requerimientos
    frmReportes.ImprimirReporte PasoReportes + "FacturacionSupervielle.rpt", Sql, True
    
  Rem  frmReportes.Exportarpdf PasoReportes + "FacturacionSupervielle.rpt", sql, True
    
Case 13
    CajasSinReferencias CInt(InputBox("Ingrese el cliente", "Cliente", 0)), CInt(InputBox("Ingrese la sucursal o colocar 0 para todo el cliente", "Sucursal", 0))
    Sql = "SELECT     * "
    Sql = Sql & "   From CONTROL_REFERENCIAS"
    Sql = Sql & "   ORDER BY FK_CAJA"
Case 14
        Sql = " SELECT     CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.TIPO_REFERENCIA, CAJAS.FK_REMITO_CUSTODIA, CONTENEDOR.ESTANTERIA,"
        Sql = Sql & " CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, CONTENEDOR.ADELANTE_ATRAS, CONTENEDOR.ESTADO,"
        Sql = Sql & " CONTENEDOR.UB_PROVISORIA"
        Sql = Sql & " FROM CAJAS INNER JOIN"
        Sql = Sql & " CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA LEFT OUTER JOIN"
        Sql = Sql & " LEGAJOS ON CAJAS.NRO_CAJA = LEGAJOS.NRO_CAJA AND CAJAS.FK_CLIENTE = LEGAJOS.COD_CLIENTE LEFT OUTER JOIN"
        Sql = Sql & " REFERENCIAS ON CAJAS.FK_CLIENTE = REFERENCIAS.COD_CLIENTE AND CAJAS.NRO_CAJA = REFERENCIAS.NRO_CAJA"
        Sql = Sql & " WHERE     (CAJAS.TIPO_REFERENCIA = 'REFERENCIA EN PLANTA' OR"
        Sql = Sql & " CAJAS.TIPO_REFERENCIA = 'PARA CARGA DE LEGAJOS') "
        Sql = Sql & " AND (REFERENCIAS.NRO_CAJA IS NULL) AND (LEGAJOS.COD_CLIENTE IS NULL)"
        Sql = Sql & " ORDER BY CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"
        
        
        Sql = " SELECT     CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_TIPO_REFERENCIA, CONTENEDOR.ESTADO, CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL,"
        Sql = Sql & "                    CONTENEDOR.Vertical"
        Sql = Sql & " FROM         CAJAS INNER JOIN"
        Sql = Sql & "                     CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA"
        Sql = Sql & " Where (Cajas.FK_TIPO_REFERENCIA = 1010)"
        Sql = Sql & " ORDER BY CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"
        
        
        Sql = " SELECT     CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_PERSONAL_LEGAJO, CAJAS.FK_TIPO_REFERENCIA, CONTENEDOR.ESTADO,"
        Sql = Sql & "                      CONTENEDOR.Estanteria , CONTENEDOR.Horizontal, CONTENEDOR.Vertical"
        Sql = Sql & " FROM         CAJAS INNER JOIN"
        Sql = Sql & "                      CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA"
        Sql = Sql & " Where (Cajas.FK_TIPO_REFERENCIA = 1010)"
        Sql = Sql & " ORDER BY CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"
        
        
        
        Sql = "  SELECT     CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_PERSONAL_LEGAJO, CAJAS.FK_TIPO_REFERENCIA, CONTENEDOR.ESTADO,"
        Sql = Sql & "   CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, COUNT(*) AS CANTIDADLEGAJOS, LEGAJOS.FK_PERSONAL_CREACION"
        Sql = Sql & "   FROM         CAJAS INNER JOIN"
        Sql = Sql & "   CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA LEFT OUTER JOIN"
        Sql = Sql & "   LEGAJOS ON CAJAS.NRO_CAJA = LEGAJOS.NRO_CAJA AND CAJAS.FK_CLIENTE = LEGAJOS.COD_CLIENTE"
        Sql = Sql & "   GROUP BY CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_PERSONAL_LEGAJO, CAJAS.FK_TIPO_REFERENCIA, CONTENEDOR.ESTADO,"
        Sql = Sql & "   CONTENEDOR.Estanteria , CONTENEDOR.Horizontal, CONTENEDOR.Vertical, LEGAJOS.FK_PERSONAL_CREACION"
        Sql = Sql & "   Having (Cajas.FK_TIPO_REFERENCIA = 1010)"
        Sql = Sql & "   ORDER BY CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"
   
    
'Case 15
'    Sql = " SELECT     CAJAS.TIPO_REFERENCIA, CAJAS.FK_REMITO_CUSTODIA, REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.FECHA"
'    Sql = Sql & "  FROM         CAJAS INNER JOIN"
'    Sql = Sql & "                        REMITOS_CUERPO ON CAJAS.FK_REMITO_CUSTODIA = REMITOS_CUERPO.NRO_REMITO"
'    Sql = Sql & "  GROUP BY CAJAS.TIPO_REFERENCIA, CAJAS.FK_REMITO_CUSTODIA, REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.FECHA"
'    Sql = Sql & "  HAVING      REMITOS_CUERPO.FECHA > '" & txtFechaDesde.Text & "'"
'    Sql = Sql & "  ORDER BY CAJAS.FK_REMITO_CUSTODIA DESC"
Case 15
        Sql = " SELECT     COD_ID_REFERENCIA, COD_CLIENTE, NRO_CAJA,DESCRIPCION,  FK_PERSONAL_CREACION, FECHA_DESDE, NRO_DESDE, LETRA_DESDE,"
        Sql = Sql & " FECHA_MODIFICACION"
        Sql = Sql & " From REFERENCIAS"
        Sql = Sql & " WHERE     (FECHA_DESDE IS NULL) AND (NRO_DESDE IS NULL) AND (LETRA_DESDE IS NULL) AND (FECHA_MODIFICACION > CONVERT(DATETIME,"
        Sql = Sql & " '2011-01-01 00:00:00', 102))"
Case 16
        Sql = " SELECT LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN, REFERENCIAS.NRO_CAJA,"
        Sql = Sql & " LEGAJOS.NRO_CAJA AS Expr1, LEGAJOS.COD_CLIENTE"
        Sql = Sql & " FROM  LECTURACOLECTOR LEFT OUTER JOIN"
        Sql = Sql & " LEGAJOS ON LECTURACOLECTOR.CAJA = LEGAJOS.NRO_CAJA AND LECTURACOLECTOR.CLIENTE = LEGAJOS.COD_CLIENTE LEFT OUTER JOIN"
        Sql = Sql & " REFERENCIAS ON LECTURACOLECTOR.CLIENTE = REFERENCIAS.COD_CLIENTE AND LECTURACOLECTOR.CAJA = REFERENCIAS.NRO_CAJA"
        Sql = Sql & " GROUP BY LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN, REFERENCIAS.NRO_CAJA,"
        Sql = Sql & " LEGAJOS.NRO_CAJA , LEGAJOS.COD_CLIENTE"
       
        Sql = Sql & " HAVING LECTURACOLECTOR.NUMERO_LECTURA IN (" & InputBox("Ingrese los numeros de lectura Separador por , ") & ") And (REFERENCIAS.NRO_CAJA Is Null) And (LEGAJOS.NRO_CAJA Is Null)"
Case 17
        
        Sql = " SELECT CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, CONTENEDOR.ESTADO, CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL,"
        Sql = Sql & " CONTENEDOR.Vertical , CAJAS.FK_REMITO_CUSTODIA, REMITOS_CUERPO.fecha"
        Sql = Sql & " FROM CONTENEDOR LEFT OUTER JOIN CAJAS ON CONTENEDOR.COD_CLIENTE = CAJAS.FK_CLIENTE AND CONTENEDOR.NRO_CAJA = CAJAS.NRO_CAJA LEFT OUTER JOIN"
        Sql = Sql & " REFERENCIAS ON CONTENEDOR.COD_CLIENTE = REFERENCIAS.COD_CLIENTE AND CONTENEDOR.NRO_CAJA = REFERENCIAS.NRO_CAJA LEFT OUTER JOIN"
        Sql = Sql & " REMITOS_CUERPO ON CAJAS.FK_REMITO_CUSTODIA = REMITOS_CUERPO.NRO_REMITO"
        Sql = Sql & " Where (REFERENCIAS.COD_CLIENTE Is Null) "
        Sql = Sql & " And (CONTENEDOR.COD_CLIENTE = 1197) "
        Sql = Sql & " And (CONTENEDOR.estado = 2)"
        Sql = Sql & " ORDER BY CONTENEDOR.NRO_CAJA"
        

Case 18
        Sql = " SELECT     LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN, LEGAJOS.NRO_CAJA,"
        Sql = Sql & " CONTENEDOR.Estanteria"
        Sql = Sql & " FROM         LECTURACOLECTOR INNER JOIN"
        Sql = Sql & " CONTENEDOR ON LECTURACOLECTOR.CAJA = CONTENEDOR.NRO_CAJA AND LECTURACOLECTOR.CLIENTE = CONTENEDOR.COD_CLIENTE LEFT OUTER JOIN"
        Sql = Sql & " LEGAJOS ON LECTURACOLECTOR.CAJA = LEGAJOS.NRO_CAJA"
        Sql = Sql & " Where LECTURACOLECTOR.NUMERO_LECTURA =  " & InputBox("Ingrese la lectura")
        Sql = Sql & " And (LEGAJOS.NRO_CAJA Is Null)"
        Sql = Sql & " And LECTURACOLECTOR.Cliente = " & InputBox("Ingrese el cliente")
Case 19
        Sql = " SELECT     LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, INDICES.DESCRIPCION, SUBSTRING(REFERENCIAS.INDICE, 1, 9) AS SUCURSAL"
        Sql = Sql & " FROM         INDICES INNER JOIN"
        Sql = Sql & " REFERENCIAS ON INDICES.INDICE = SUBSTRING(REFERENCIAS.INDICE, 1, 9) AND INDICES.COD_CLIENTE = REFERENCIAS.COD_CLIENTE RIGHT OUTER JOIN"
        Sql = Sql & " LECTURACOLECTOR ON REFERENCIAS.COD_CLIENTE = LECTURACOLECTOR.CLIENTE AND REFERENCIAS.NRO_CAJA = LECTURACOLECTOR.CAJA"
        Sql = Sql & " GROUP BY LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, INDICES.DESCRIPCION, SUBSTRING(REFERENCIAS.INDICE, 1,9)"
        Sql = Sql & " HAVING LECTURACOLECTOR.NUMERO_LECTURA IN (" & InputBox("Ingrese La lectura ") & ")"
        Sql = Sql & " ORDER BY SUCURSAL, LECTURACOLECTOR.CAJA"
Case 20
        Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV, CONVERT(char, REMITOS_CUERPO.FECHA, 103) AS FECHA,"
        Sql = Sql & " REMITO_TIPO.DESCRIPCION AS TIPO , REMITOS_CUERPO.CANTIDAD, REMITOS_CUERPO.ID_CLIENTE, INDICES.DESCRIPCION"
        Sql = Sql & " FROM         CLIENTEUSUARIO INNER JOIN"
        Sql = Sql & " REMITOS_CUERPO ON CLIENTEUSUARIO.ID_CLIENTEUSUARIO = REMITOS_CUERPO.COD_USUARIO_CLIENTE INNER JOIN"
        Sql = Sql & "  INDICES ON CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE AND CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE INNER JOIN"
        Sql = Sql & "  REMITO_TIPO ON REMITOS_CUERPO.TIPO = REMITO_TIPO.ID"
        Sql = Sql & "  WHERE     (REMITOS_CUERPO.TIPO IN (0, 3)) "
        Sql = Sql & "  AND REMITOS_CUERPO.ID_CLIENTE =  " & InputBox("Ingrese el cliente")
        Sql = Sql & "  AND (REMITOS_CUERPO.ANULADO IS NULL)"
        Sql = Sql & "  ORDER BY INDICES.DESCRIPCION"
Case 21
        Sql = " SELECT REMITOS_CUERPO.NRO_REMITO, 'RM: ' + REMITOS_CUERPO.NRO_REM_PROV, CONVERT(char, REMITOS_CUERPO.FECHA, 103) AS FECHA, RTRIM(PERSONAL.NOMBRE)"
        Sql = Sql & " + ' ' + RTRIM(PERSONAL.APELLIDO) AS NOMBRE, REMITOS_CUERPO.AUDIT_FECHA, REMITOS_CUERPO.IMAGEN, REMITOS_CUERPO.ID_CLIENTE,"
        Sql = Sql & " REMITOS_CUERPO.FECHA_LECTURA_REMITO , Clientes.RAZON_SOCIAL "
        Sql = Sql & " FROM REMITOS_CUERPO INNER JOIN"
        Sql = Sql & " CLIENTES ON REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE INNER JOIN"
        Sql = Sql & " PERSONAL ON REMITOS_CUERPO.COD_PERSONAL_ENTREGA = PERSONAL.IDPERSONAL"
        Sql = Sql & " WHERE   (REMITOS_CUERPO.IMAGEN IS NULL) AND (REMITOS_CUERPO.ANULADO IS NULL) "
        Sql = Sql & " AND REMITOS_CUERPO.FECHA >= " & FechaFormato(txtFechaDesde.Text)
        Sql = Sql & " AND REMITOS_CUERPO.FECHA <= " & FechaFormato(txtFechaHasta.Text)
        Sql = Sql & " ORDER BY REMITOS_CUERPO.FECHA_LECTURA_REMITO , REMITOS_CUERPO.FECHA "
Case 22
        Sql = "  SELECT  REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV,CONVERT( char, REMITOS_CUERPO.FECHA ,103) as FECHA, REMITO_TIPO.DESCRIPCION,"
        Sql = Sql & " CLIENTEUSUARIO.APELLIDO_NOMBRE, REMITOS_DETALLE.DESDE AS CAJA, REMITOS_CUERPO.ID_CLIENTE"
        Sql = Sql & " FROM         REMITOS_CUERPO INNER JOIN"
        Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO INNER JOIN"
        Sql = Sql & " REMITO_TIPO ON REMITOS_CUERPO.TIPO = REMITO_TIPO.ID LEFT OUTER JOIN"
        Sql = Sql & " CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
        Sql = Sql & " Where (REMITOS_CUERPO.TIPO = 0) AND (REMITOS_CUERPO.ANULADO IS NULL) "
        Rem Sql = Sql & "  AND REMITOS_CUERPO.COD_USUARIO_CLIENTE in(3248, 4634, 2898, 3849)"
        Sql = Sql & " AND REMITOS_CUERPO.id_cliente =" & InputBox("INGRESE EL NUMERO DE CLIENTE", 0)
        Sql = Sql & " ORDER BY REMITOS_CUERPO.FECHA"
Case 23
        Sql = " SELECT ROLLO, ID_LEGAJO, FK_PERSONAL_CREACION, FECHA_CREACION"
        Sql = Sql & " From LEGAJOS"
        Sql = Sql & " Where (NRO_CAJA Is Null) And ROLLO = " & InputBox("Ingrese el rollo")
        Sql = Sql & " ORDER BY ID_LEGAJO"
Case 24
        Sql = " SELECT DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.LETRA_HASTA, DOCUMENTOS_DIGITALES.NRO_DESDE,"
        Sql = Sql & "  DOCUMENTOS_DIGITALES.NRO_HASTA, DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.FECHA_HASTA, DOCUMENTOS_DIGITALES.Nombre,"
        Sql = Sql & "  DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.NRO_CAJA, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
        Sql = Sql & "  DOCUMENTOS_DIGITALES.ID"
        Sql = Sql & "  FROM DOCUMENTOS_DIGITALES INNER JOIN"
        Sql = Sql & "  DOCUMENTOS_DIGITALES_LOTE ON"
        Sql = Sql & "  DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & "  Where DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES =  " & InputBox("INGRESE EL CLIENTE")
        Sql = Sql & "  And (Not (DOCUMENTOS_DIGITALES.NRO_DESDE Is Null))"
        Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE"
Case 25
        Sql = " SELECT     REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.ID_CLIENTE, CONVERT(char, REQUERIMIENTO.FECHARECEPCION, 103) AS FECHARECEPCION,"
        Sql = Sql & " REQUERIMIENTO.CANTIDAD, CONVERT(char(10), REQUERIMIENTO.FECHA_SISTEMA, 103) + ' ' + CONVERT(char(9), REQUERIMIENTO.FECHA_SISTEMA, 108)"
        Sql = Sql & " AS FECHA_CARGA_SISTEMA,  CONVERT(char(2), REQUERIMIENTO.FECHA_SISTEMA, 108) as HORA , CONVERT(char(10), REQUERIMIENTO.FECHAENTREGA, 103) AS COMPROMISO_ENTREGA,"
        Sql = Sql & " REQUERIMIENTO.COMPROMISO_ENTREGA AS Expr1, REQUERIMIENTO.IDPERSONAL, TIPOREQUERIMIENTO.DESCRIPCION, REQUERIMIENTO.TOMO,"
        Sql = Sql & " PERSONAL.Nombre , PERSONAL.Apellido"
        Sql = Sql & " FROM         REQUERIMIENTO INNER JOIN"
        Sql = Sql & " TIPOREQUERIMIENTO ON REQUERIMIENTO.IDTIPOREQUERIMIENTO = TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO INNER JOIN"
        Sql = Sql & " PERSONAL ON REQUERIMIENTO.TOMO = PERSONAL.IDPERSONAL"
        Sql = Sql & " WHERE     (REQUERIMIENTO.FECHAENTREGA > CONVERT(DATETIME, '2012-08-01 00:00:00', 102))"
        Sql = Sql & " ORDER BY REQUERIMIENTO.FECHA_SISTEMA"
Case 26
        Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV, CONVERT(char, REMITOS_CUERPO.FECHA, 103) AS FECHA,"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.CANTIDAD, REMITOS_DETALLE.DESDE AS CAJA, CLIENTEUSUARIO.APELLIDO_NOMBRE,"
        Sql = Sql & vbCrLf & " INDICES.Descripcion"
        Sql = Sql & vbCrLf & " FROM         INDICES INNER JOIN"
        Sql = Sql & vbCrLf & " CLIENTEUSUARIO ON INDICES.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND INDICES.INDICE = CLIENTEUSUARIO.COD_INDICE RIGHT OUTER JOIN"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO INNER JOIN"
        Sql = Sql & vbCrLf & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO ON"
        Sql = Sql & vbCrLf & " CLIENTEUSUARIO.ID_CLIENTEUSUARIO = REMITOS_CUERPO.COD_USUARIO_CLIENTE"
        Sql = Sql & vbCrLf & " WHERE     (REMITOS_CUERPO.TIPO = 0) AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) AND (REMITOS_CUERPO.ANULADO IS NULL) "
        Sql = Sql & vbCrLf & " AND      REMITOS_CUERPO.FECHA BETWEEN " & FechaFormato(txtFechaDesde.Text) & "  AND " & FechaFormato(txtFechaHasta.Text)
        Sql = Sql & vbCrLf & " AND    REMITOS_CUERPO.ID_CLIENTE = " & InputBox("INGRESE EL CLIENTE")
        Sql = Sql & vbCrLf & " ORDER BY REMITOS_CUERPO.FECHA,  REMITOS_DETALLE.DESDE  "
Case 27
        Sql = "  SELECT     REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV, CONVERT (CHAR , REMITOS_CUERPO.FECHA, 103) AS FECHA,"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.CANTIDAD, REMITOS_DETALLE.DESDE AS CAJAS"
        Sql = Sql & vbCrLf & " FROM         REMITOS_DETALLE INNER JOIN"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO ON REMITOS_DETALLE.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " REFERENCIAS ON REMITOS_CUERPO.ID_CLIENTE = REFERENCIAS.COD_CLIENTE AND REMITOS_DETALLE.DESDE = REFERENCIAS.NRO_CAJA"
        Sql = Sql & vbCrLf & " Where (REFERENCIAS.NRO_CAJA Is Null) "
        Sql = Sql & vbCrLf & " And REMITOS_CUERPO.NRO_REMITO = " & InputBox("Ingrese el Numero remito Sistema")
Case 28
        Sql = "  SELECT     REMITOS_CUERPO.ID_CLIENTE AS FK_CLIENTE, REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV, TIPO_REMITO.DESCRIPCION AS TIPO,"
        Sql = Sql & vbCrLf & "  REMITOS_CUERPO.CANTIDAD, REMITOS_CUERPO.TIPO AS TIPOVALOR, CONVERT(char, REMITOS_CUERPO.FECHA, 103) AS Fecha, REMITOS_CUERPO.ANULADO,"
        Sql = Sql & vbCrLf & "  CLIENTEUSUARIO.APELLIDO_NOMBRE, INDICES.DESCRIPCION AS SUCURSAL"
        Sql = Sql & vbCrLf & " FROM         TIPO_REMITO INNER JOIN"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO ON TIPO_REMITO.ID = REMITOS_CUERPO.TIPO LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " INDICES INNER JOIN"
        Sql = Sql & vbCrLf & " CLIENTEUSUARIO ON INDICES.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND INDICES.INDICE = CLIENTEUSUARIO.COD_INDICE ON"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
        Sql = Sql & vbCrLf & " WHERE     (REMITOS_CUERPO.ANULADO IS NULL) "
        Sql = Sql & vbCrLf & "  AND (REMITOS_CUERPO.TIPO IN (0, 3)) "
        Sql = Sql & vbCrLf & " AND  REMITOS_CUERPO.FECHA BETWEEN " & FechaFormato(txtFechaDesde.Text) & "  AND " & FechaFormato(txtFechaHasta.Text)
        Sql = Sql & vbCrLf & " AND REMITOS_CUERPO.ID_CLIENTE = " & InputBox("Ingrese el cliente")
        ExecutarSql " DELETE FROM basasql.dbo.TEM_DISCO"
        InsertarDisco Sql
        Sql = " SELECT * From basasql.dbo.TEM_DISCO "
        frmReportes.ImprimirReporte PasoReportes + "rptFacuracionDisco.rpt", Sql, True
Case 29
        Sql = " SELECT     REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.FK_SUCURSAL, TIPOREQUERIMIENTO.DESCRIPCION, REQUERIMIENTO.FECHARECEPCION,"
        Sql = Sql & vbCrLf & " Clientes.RAZON_SOCIAL , REQUERIMIENTO.Imagen, REQUERIMIENTO.ANULADO"
        Sql = Sql & vbCrLf & " FROM         REQUERIMIENTO INNER JOIN"
        Sql = Sql & vbCrLf & " TIPOREQUERIMIENTO ON REQUERIMIENTO.IDTIPOREQUERIMIENTO = TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO INNER JOIN"
        Sql = Sql & vbCrLf & " CLIENTES ON REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE"
        Sql = Sql & vbCrLf & " WHERE     (REQUERIMIENTO.IDTIPOREQUERIMIENTO IN (20, 13, 5, 8, 9, 14, 18, 19, 23, 22)) AND (REQUERIMIENTO.IMAGEN IS NULL) AND"
        Sql = Sql & vbCrLf & " NOT (REQUERIMIENTO.ANULADO IS NULL)"
        Sql = Sql & " AND REQUERIMIENTO. FECHAENTREGA BETWEEN " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text)
Case 30
        Sql = " SELECT     CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE,"
        Sql = Sql & vbCrLf & " CONTENEDOR.NRO_CAJA, CONTENEDOR.NRO_REMITO, REFERENCIAS.NRO_CAJA AS CAJAREFERENCIA, LEGAJOS.NRO_CAJA AS CAJALEGAJOS"
        Sql = Sql & vbCrLf & " FROM         CONTENEDOR LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " LEGAJOS ON CONTENEDOR.NRO_CAJA = LEGAJOS.NRO_CAJA AND CONTENEDOR.COD_CLIENTE = LEGAJOS.COD_CLIENTE LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " REFERENCIAS ON CONTENEDOR.NRO_CAJA = REFERENCIAS.NRO_CAJA AND CONTENEDOR.COD_CLIENTE = REFERENCIAS.COD_CLIENTE"
        Sql = Sql & vbCrLf & " GROUP BY CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE,"
        Sql = Sql & vbCrLf & " CONTENEDOR.NRO_CAJA , CONTENEDOR.NRO_REMITO, REFERENCIAS.NRO_CAJA, LEGAJOS.NRO_CAJA"
        Sql = Sql & vbCrLf & " HAVING      (CONTENEDOR.ESTANTERIA BETWEEN " & InputBox("Estanteria desde ") & " AND  " & InputBox("Estanteria Hasta") & ")"
        Sql = Sql & vbCrLf & " AND (CONTENEDOR.HORIZONTAL BETWEEN " & InputBox("Horizontal Desde") & " AND " & InputBox("Horizontal Hasta") & ")"
        Sql = Sql & vbCrLf & " AND (CONTENEDOR.VERTICAL BETWEEN " & InputBox("Vertical Desde") & " AND  " & InputBox("Vertical Hasta") & ") "
        Sql = Sql & vbCrLf & " AND (REFERENCIAS.NRO_CAJA IS NULL)"
        Sql = Sql & vbCrLf & " AND (NOT (CONTENEDOR.COD_CLIENTE IS NULL))"
        Sql = Sql & vbCrLf & " ORDER BY CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL"
Case 31
        Sql = " SELECT     CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE,"
        Sql = Sql & vbCrLf & "                     CONTENEDOR.NRO_CAJA, REFERENCIAS.NRO_CAJA AS CAJAREFERENCIA, LEGAJOS.NRO_CAJA AS CAJALEGAJOS, INDICES.DESCRIPCION AS DESCINDICE,"
        Sql = Sql & vbCrLf & "                     REFERENCIAS.DESCRIPCION, REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA, REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA,"
        Sql = Sql & vbCrLf & "                    REFERENCIAS.LETRA_DESDE , REFERENCIAS.LETRA_HASTA"
        Sql = Sql & vbCrLf & " FROM         INDICES INNER JOIN"
        Sql = Sql & vbCrLf & "                     REFERENCIAS ON INDICES.COD_CLIENTE = REFERENCIAS.COD_CLIENTE AND INDICES.INDICE = REFERENCIAS.INDICE RIGHT OUTER JOIN"
        Sql = Sql & vbCrLf & "                    CONTENEDOR LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & "                   LEGAJOS ON CONTENEDOR.NRO_CAJA = LEGAJOS.NRO_CAJA AND CONTENEDOR.COD_CLIENTE = LEGAJOS.COD_CLIENTE ON"
        Sql = Sql & vbCrLf & "                    REFERENCIAS.NRO_CAJA = CONTENEDOR.NRO_CAJA And REFERENCIAS.COD_CLIENTE = CONTENEDOR.COD_CLIENTE"
        Sql = Sql & vbCrLf & " WHERE   (NOT (CONTENEDOR.ESTADO IN (0, 1)) AND CONTENEDOR.ESTANTERIA BETWEEN " & InputBox("Estanteria desde ") & " AND  " & InputBox("Estanteria Hasta") & ")"
        Sql = Sql & vbCrLf & " AND (CONTENEDOR.HORIZONTAL BETWEEN " & InputBox("Horizontal Desde") & " AND " & InputBox("Horizontal Hasta") & ")"
        Sql = Sql & vbCrLf & " AND (CONTENEDOR.VERTICAL BETWEEN " & InputBox("Vertical Desde") & " AND  " & InputBox("Vertical Hasta") & ") "
        Sql = Sql & vbCrLf & " ORDER BY CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL"


Case 32
        Sql = "  SELECT     ID_CONTENEDOR, FK_CAJAS, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA, NRO_REMITO,"
        Sql = Sql & vbCrLf & " F_MODIFICACION, IDREQUERIMIENTO, NUEVA, BAJA, UB_PROVISORIA, COD_CAJA, JERAQUIA, COD_INDICE, COD_CLIENTE_USUARIO,"
        Sql = Sql & vbCrLf & " COD_RESPONSABLE_POSICION, FECHA_CREACION, MODULO_V, MODULO_H, CONTROL, MODULO, COD_REMITO_GUARDA, COD_USUARIO_CLIENTE_GUARDA,"
        Sql = Sql & vbCrLf & " COD_INDICE_SECTOR , Orden, FECHAPOSICION"
        Sql = Sql & vbCrLf & "  From basasql.dbo.CONTENEDOR_23012012_1942"
        Sql = Sql & vbCrLf & " Where NRO_CAJA = " & InputBox("Ingrese la caja")

Case 33
Case 34
        IVADATA
            Sql = " SELECT    CONVERT(char, FechaFacturacion , 103 ) as FECHA , NombreCliente, Letra,  NumeroFactura,  factura,  CUIT, Subtotal, IVA, TotalFacturado"
            Sql = Sql & vbCrLf & " From basasql.dbo.TEM_IVA_DATA"
        Select Case InputBox("1-Basa , 2 - Custodia , 3 - Electronica ")
        Case 1
            Sql = Sql & vbCrLf & "   WHERE     (FacturaABC IN (N'G', N'F')) AND (NumeroFactura > 1000)"
        Case 2
            Sql = Sql & vbCrLf & "   WHERE     (FacturaABC IN (N'A', N'B')) AND (NumeroFactura > 0 )"
        Case 3
            Sql = Sql & vbCrLf & "   WHERE     (FacturaABC IN (N'G', N'F' ,  N'E')) AND (NumeroFactura < 1000)"
        End Select
            Sql = Sql & vbCrLf & " ORDER BY FacturaABC, NumeroFactura, FechaFacturacion "

Case 35
        Sql = " SELECT     CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_TIPO_REFERENCIA, CONTENEDOR.ESTADO, CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL,"
        Sql = Sql & "                    CONTENEDOR.Vertical"
        Sql = Sql & " FROM         CAJAS INNER JOIN"
        Sql = Sql & "                     CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA"
        Sql = Sql & " Where (Cajas.FK_TIPO_REFERENCIA = 1090)"
        Sql = Sql & " ORDER BY CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"
Case 36
        Sql = "  SELECT     ESTANTERIA, ESTADO, COD_CLIENTE, NRO_CAJA"
        Sql = Sql & "  From CONTENEDOR"
        Sql = Sql & "  Where (Estanteria > 1000) And (estado = 5)"
        Sql = Sql & "  ORDER BY COD_CLIENTE, NRO_CAJA"
Case 37
        Sql = " SELECT     CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_TIPO_REFERENCIA, CAJAS.FK_TIPO_REFERENCIA_PERSONAL, CONVERT(char,"
        Sql = Sql & vbCrLf & " CAJAS.TIPO_REFERENCIA_FECHA, 103) AS fecha, PARAMETROS.DESCRIPCION, CONTENEDOR.ESTADO"
        Sql = Sql & vbCrLf & " FROM CAJAS INNER JOIN "
        Sql = Sql & vbCrLf & " PARAMETROS ON CAJAS.FK_TIPO_REFERENCIA = PARAMETROS.ID_PARAMETRO INNER JOIN"
        Sql = Sql & vbCrLf & " CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA"
        Sql = Sql & vbCrLf & " Where (Not (CAJAS.TIPO_REFERENCIA_FECHA Is Null)) And (CAJAS.FK_TIPO_REFERENCIA <> 1020)"
        Sql = Sql & vbCrLf & " ORDER BY CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"
 Case 38
        Sql = " SELECT     CLIENTES.ID_CLIENTE, CLIENTES.RAZON_SOCIAL, REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.TIPO, CONVERT(char, REMITOS_CUERPO.FECHA, 103)"
        Sql = Sql & vbCrLf & " AS FECHA, REMITOS_CUERPO.CANTIDAD"
        Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO INNER JOIN"
        Sql = Sql & vbCrLf & " CLIENTES ON REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE"
        Sql = Sql & vbCrLf & " Where (REMITOS_CUERPO.ANULADO Is Null) And (REMITOS_CUERPO.TIPO = 2)"
        Sql = Sql & vbCrLf & " AND REMITOS_CUERPO.FECHA BETWEEN " & FechaFormato(txtFechaDesde.Text) & "  AND " & FechaFormato(txtFechaHasta.Text)
        Sql = Sql & vbCrLf & "  ORDER BY FECHA"
Case 39
        Sql = " SELECT     CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, ORDEN_LEGAJOS.REARCHIVO_CAJA, CONTENEDOR.ESTANTERIA,"
        Sql = Sql & vbCrLf & " CONTENEDOR.Horizontal , CONTENEDOR.Vertical "
        Sql = Sql & vbCrLf & " FROM         CONTENEDOR LEFT OUTER JOIN "
        Sql = Sql & vbCrLf & " ORDEN_LEGAJOS ON CONTENEDOR.NRO_CAJA = ORDEN_LEGAJOS.REARCHIVO_CAJA "
        Sql = Sql & vbCrLf & " GROUP BY CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA"
        Sql = Sql & vbCrLf & " ,ORDEN_LEGAJOS.REARCHIVO_CAJA, CONTENEDOR.ESTANTERIA,"
        Sql = Sql & vbCrLf & " CONTENEDOR.Horizontal , CONTENEDOR.Vertical "
        Sql = Sql & vbCrLf & " Having (CONTENEDOR.COD_CLIENTE = 291) "
        Sql = Sql & vbCrLf & " And (CONTENEDOR.estado = 2) "
        Sql = Sql & vbCrLf & " And (ORDEN_LEGAJOS.REARCHIVO_CAJA Is Null) "
Case 40
        Sql = " SELECT     NRO_CAJA, COD_CLIENTE, COUNT(*) AS Cantidad"
        Sql = Sql & vbCrLf & " From basasql.dbo.CAMBIOPOSICION"
        Sql = Sql & vbCrLf & " GROUP BY NRO_CAJA, COD_CLIENTE"
        Sql = Sql & vbCrLf & " HAVING  (COD_CLIENTE = " & InputBox("Ingrese el cliente") & ")"
        Sql = Sql & vbCrLf & " AND (NRO_CAJA BETWEEN  " & InputBox("Ingrese la caja inicio") & "  AND " & InputBox("Ingrese la caja fin ")
        Sql = Sql & vbCrLf & " ) AND (COUNT(*) > 1)"
        Sql = Sql & vbCrLf & " ORDER BY Cantidad DESC"

Case 41
        Sql = " SELECT     NRO_CAJA, COD_CLIENTE, COUNT(*) AS Cantidad"
        Sql = Sql & vbCrLf & " From basasql.dbo.CAMBIOPOSICION"
        Sql = Sql & vbCrLf & " GROUP BY NRO_CAJA, COD_CLIENTE"
        Sql = Sql & vbCrLf & " HAVING  "
        Sql = Sql & vbCrLf & "  (NRO_CAJA BETWEEN  " & InputBox("Ingrese la caja inicio") & "  AND " & InputBox("Ingrese la caja fin ")
        Sql = Sql & vbCrLf & " ) AND (COUNT(*) > 1)"
        Sql = Sql & vbCrLf & " ORDER BY Cantidad DESC"

Case 42
        Sql = "  SELECT     DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES.LETRA_HASTA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE , DOCUMENTOS_DIGITALES_LOTE.Descripcion, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"
        Sql = Sql & vbCrLf & " FROM         DOCUMENTOS_DIGITALES INNER JOIN"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE ON"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & vbCrLf & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401)"
        Sql = Sql & vbCrLf & " AND DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS in(" & InputBox("Ingrese la caja ") & ") "
        Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"

Case 43
        ControlCajas5000SInMovimiento InputBox("Ingrese el Cliente")
        Sql = "SELECT     CAMBIOPOSICION.ESTANTERIA, CAMBIOPOSICION.HORIZONTAL, CAMBIOPOSICION.VERTICAL, CAMBIOPOSICION.NRO_ESTANTE, CAMBIOPOSICION.ESTADO,"
        Sql = Sql & vbCrLf & "                      CAMBIOPOSICION.COD_CLIENTE, CAMBIOPOSICION.NRO_CAJA, CAMBIOPOSICION.FECHA, CAMBIOPOSICION.ID_PERSONAL,"
        Sql = Sql & vbCrLf & "                     V_TEM_CONTROL_CAJAS_5000.CANTDUPLI"
        Sql = Sql & vbCrLf & " FROM         V_TEM_CONTROL_CAJAS_5000 INNER JOIN"
        Sql = Sql & vbCrLf & "                       CAMBIOPOSICION ON V_TEM_CONTROL_CAJAS_5000.COD_CLIENTE = CAMBIOPOSICION.COD_CLIENTE AND"
        Sql = Sql & vbCrLf & "                       V_TEM_CONTROL_CAJAS_5000.NRO_CAJA = CAMBIOPOSICION.NRO_CAJA"
        Sql = Sql & vbCrLf & " Where (CAMBIOPOSICION.Estanteria > 5000)"
        Sql = Sql & vbCrLf & " ORDER BY V_TEM_CONTROL_CAJAS_5000.CANTDUPLI DESC, CAMBIOPOSICION.FECHA"
Case 44
      
        Sql = " SELECT     ID_LEGAJO, NRO_CAJA AS CAJA , COD_CLIENTE AS CLIENTE , LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION, "
        Sql = Sql & " FECHA_ACTUALIZACION AS FECHA , FK_PERSONAL_ACTUALIZACION   as PERSONAL"
        Sql = Sql & " From LEGAJOS"
        Sql = Sql & " WHERE  FK_PERSONAL_ACTUALIZACION   = " & InputBox("INGRESE EL PERSONAL QUE CARGO EL LEGAJO")
        Sql = Sql & " AND convert(char,FECHA_ACTUALIZACION ,103) ='" & InputBox("INGRESE LA FECHA ") & "'"
        Sql = Sql & " ORDER BY  FECHA_ACTUALIZACION "
Case 45
        EXPURGO_DISCO InputBox("Ingrese el N° de Cliente", "Cliente", 0), InputBox("Ingrse el numero documento o indice", "Indice", 1), InputBox("Ingrese la fecha del filtro", "Fecha", "31/12/2004")
Case 46
        Sql = " SELECT     CONVERT(char, FECHA_CREACION, 103) AS FECHA_CREACION, NRO_CAJA, COD_CLIENTE, FK_PERSONAL_CREACION"
        Sql = Sql & "  From LEGAJOS"
        Sql = Sql & "  GROUP BY CONVERT(char, FECHA_CREACION, 103), NRO_CAJA, COD_CLIENTE, FK_PERSONAL_CREACION"
        Sql = Sql & "  HAVING      (CONVERT(char, FECHA_CREACION, 103) = '" & InputBox("Ingrese la fecha", "Fecha", Format(Now, "DD/MM/YYYY")) & "')"
        Sql = Sql & "  ORDER BY FECHA_CREACION"
Case 47
        Sql = " SELECT     REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.ID_CLIENTE, REQUELIBOSCAJAS.PERSONAL, CONVERT(char, REQUERIMIENTO.FECHARECEPCION , 103 ) as FECHA, PERSONAL.NOMBRE, PERSONAL.APELLIDO, TIPOREQUERIMIENTO.DESCRIPCION"
        Sql = Sql & " FROM         REQUELIBOSCAJAS INNER JOIN"
        Sql = Sql & " REQUERIMIENTO ON REQUELIBOSCAJAS.IDREQUERIMIENTOS = REQUERIMIENTO.IDREQUERIMIENTO INNER JOIN"
        Sql = Sql & " PERSONAL ON REQUELIBOSCAJAS.PERSONAL = PERSONAL.IDPERSONAL INNER JOIN"
        Sql = Sql & " TIPOREQUERIMIENTO ON REQUERIMIENTO.IDTIPOREQUERIMIENTO = TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO"
        Sql = Sql & " WHERE     REQUERIMIENTO.FECHARECEPCION > '" & InputBox("Ingrese la fecha", "Fecha", Format(Now, "DD/MM/YYYY")) & "' AND (NOT (REQUELIBOSCAJAS.PERSONAL IS NULL))"
        Sql = Sql & " ORDER BY REQUERIMIENTO.IDREQUERIMIENTO"
Case 48
    Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV, REMITOS_CUERPO.ANULADO, REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO,"
    Sql = Sql & " REMITOS_CUERPO.ID_CLIENTE, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, CONTENEDOR.ESTADO, CONTENEDOR.ESTANTERIA,"
    Sql = Sql & " CONTENEDOR.Horizontal , CONTENEDOR.Vertical"
    Sql = Sql & " FROM         REMITOS_CUERPO INNER JOIN"
    Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO INNER JOIN"
    Sql = Sql & " CONTENEDOR ON REMITOS_CUERPO.ID_CLIENTE = CONTENEDOR.COD_CLIENTE AND REMITOS_DETALLE.DESDE = CONTENEDOR.NRO_CAJA"
    Sql = Sql & " Where (REMITOS_CUERPO.TIPO = 3) And (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0)"
    Sql = Sql & " ORDER BY REMITOS_CUERPO.NRO_REMITO DESC"
Case 49
    Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV, REMITOS_CUERPO.TIPO, REMITOS_DETALLE.DESDE, REMITOS_CUERPO.ID_CLIENTE,"
    Sql = Sql & " CONVERT(char, REMITOS_CUERPO.FECHA ,103) as FECHA, CLIENTEUSUARIO.APELLIDO_NOMBRE"
    Sql = Sql & "  FROM         REMITOS_CUERPO INNER JOIN"
    Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO INNER JOIN"
    Sql = Sql & " CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
    Sql = Sql & "  WHERE     (REMITOS_CUERPO.ID_CLIENTE = 4) AND (REMITOS_CUERPO.FECHA > CONVERT(DATETIME, '2015-01-01 00:00:00', 102)) AND (REMITOS_CUERPO.TIPO = 0)"
    Sql = Sql & "  ORDER BY REMITOS_CUERPO.NRO_REMITO"
Case 50
    Sql = " SELECT     BARRA_PASO, SUBSTRING(CONVERT(char, NombreArchivoNumero), 1, 7) AS caja, COUNT(*) AS Cantidad"
    Sql = Sql & " From basasql.dbo.TELEFORM_BARRA"
    Sql = Sql & " Where  BatchNo  =" & InputBox("Ingrese el numero de BatchNo")
    Sql = Sql & " GROUP BY BARRA_PASO, SUBSTRING(CONVERT(char, NombreArchivoNumero), 1, 7)"
    Sql = Sql & " ORDER BY Cantidad"
Case 51
    Sql = " SELECT ID , NOMBRE_ARCHIVO , NRO_CAJA , CANTIDAD_IMAGEN, LOTEHORA"
    Sql = Sql & "  ,FK_CLIENTE , MESAÑO "
    Sql = Sql & " From CANTIDAD_IMAGEN "
    Sql = Sql & " WHERE  FK_CLIENTE = " & InputBox("Ingrese el numero de Cliente")
Case 52
    Sql = " SELECT LECTURACOLECTOR.CAJA AS CAJABAJA, DISCO_SUSURSALES_CAJAS.CLIENTE_CUSTODIA, LECTURACOLECTOR.NUMERO_LECTURA"
    Sql = Sql & " FROM DISCO_SUSURSALES_CAJAS RIGHT OUTER JOIN"
    Sql = Sql & " LECTURACOLECTOR ON DISCO_SUSURSALES_CAJAS.CAJA = LECTURACOLECTOR.CAJA"
    Sql = Sql & " Where LECTURACOLECTOR.NUMERO_LECTURA = " & InputBox("Ingrese la lectura")
    Sql = Sql & " GROUP BY LECTURACOLECTOR.CAJA, DISCO_SUSURSALES_CAJAS.CLIENTE_CUSTODIA, LECTURACOLECTOR.NUMERO_LECTURA"
    Sql = Sql & " ORDER BY DISCO_SUSURSALES_CAJAS.CLIENTE_CUSTODIA"
Case 53
    Sql = " SELECT        CONVERT(char, FECHA_INDEXACION, 103) AS fecha, DATEPART(HH, FECHA_INDEXACION) AS HORA, DATEPART(mi, FECHA_INDEXACION) AS Minuto, PERSONAL_INDEXACION,"
    Sql = Sql & " FK_DOCUMENTOS_DIGITALES_LOTE, FECHA_INDEXACION AS Expr1, ID, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA"
    Sql = Sql & " From DOCUMENTOS_DIGITALES"
    Sql = Sql & " WHERE (PERSONAL_INDEXACION IN (" & InputBox("Ingrese el usuario") & ")) "
    Sql = Sql & " AND (CONVERT(char, FECHA_INDEXACION, 103) = '" & InputBox("Ingrese la fecha", "fecha", date) & "')"
    Sql = Sql & " ORDER BY FECHA_INDEXACION "
Case 54
    Sql = "SELECT        DATEPART(HH, FECHA_INDEXACION) AS HORA , count(*) as cantidad_hora"
    Sql = Sql & " From DOCUMENTOS_DIGITALES"
    Sql = Sql & " WHERE   (PERSONAL_INDEXACION IN (" & InputBox("Ingrese el usuario") & ")) "
    Sql = Sql & " AND (CONVERT(char, FECHA_INDEXACION, 103) = '" & InputBox("Ingrese la fecha", "fecha", date) & "')"
    Sql = Sql & " GROUP BY DATEPART(HH, FECHA_INDEXACION)"
Case 55 'CARAGA EXTERNA
    Sql = " SELECT    FK_CLIENTES, FK_CAJAS, PERSONAL_PREPARACION_EXTERNO, NOMBRE, APELLIDO, AÑO, MES, DIA, CANTIDAD, COSTOINDEXACION, TOTAL,Descripcion"
    Sql = Sql & " From V_CONTROL_INDEX_EXTERNO"
    Sql = Sql & " WHERE AÑO = " & Mid(txtFechaDesde, 7, 4) & " AND  MES = " & Mid(txtFechaDesde, 4, 2)
    Sql = Sql & " ORDER BY PERSONAL_PREPARACION_EXTERNO, AÑO, MES, DIA"
Case 56 ' PREPARACION
    Sql = " SELECT        DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES,  DATEPART(YYYY, DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION) AS AÑO,  DATEPART(MM, DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION) AS MES, DATEPART(DD,"
    Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION) AS DIA, PERSONAL.NOMBRE, PERSONAL.APELLIDO, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
    Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.FK_PERSONAL_PREPARACION AS PERSONAL, INDICES.DESCRIPCION, INDICES.COSTOPREPARACION AS COSTO_UNITARIO,"
    Sql = Sql & " SUM(DOCUMENTOS_DIGITALES_LOTE.CANTIDAD_IMAGENES) AS CANTIDAD_IMAGENES,"
    Sql = Sql & " INDICES.COSTOPREPARACION * SUM(DOCUMENTOS_DIGITALES_LOTE.CANTIDAD_IMAGENES) AS PAGO"
    Sql = Sql & " FROM PERSONAL INNER JOIN"
    Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE ON PERSONAL.IDPERSONAL = DOCUMENTOS_DIGITALES_LOTE.FK_PERSONAL_PREPARACION RIGHT OUTER JOIN"
    Sql = Sql & " INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
    Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION >= '" & txtFechaDesde.Text & "' AND  DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION <= '" & txtFechaHasta.Text & "'"
    Sql = Sql & " GROUP BY DATEPART(YYYY, DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION), DATEPART(MM, DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION),"
    Sql = Sql & " DATEPART(DD, DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION), PERSONAL.NOMBRE, PERSONAL.APELLIDO, INDICES.ID,"
    Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES,"
    Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.FK_PERSONAL_PREPARACION , INDICES.Descripcion, INDICES.COSTOPREPARACION"
    Sql = Sql & " Having DOCUMENTOS_DIGITALES_LOTE.FK_PERSONAL_PREPARACION > 100 "
    Sql = Sql & " ORDER BY PERSONAL, MES, DIA, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS"
Case 57 ' DIGITALIZACION
    Sql = " SELECT DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DATEPART(YYYY, DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER) AS AÑO, DATEPART(MM,"
    Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER) AS MES, DATEPART(DD, DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER) AS DIA, PERSONAL.NOMBRE,"
    Sql = Sql & vbCrLf & " PERSONAL.APELLIDO, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES_LOTE.FK_PERSONAL_SCANNER AS PERSONAL_SCANER,"
    Sql = Sql & vbCrLf & " INDICES.DESCRIPCION, INDICES.COSTODIGITALIZACION AS COSTO_UNITARIO_DIGITALIZA, SUM(DOCUMENTOS_DIGITALES_LOTE.CANTIDAD_IMAGENES)"
    Sql = Sql & vbCrLf & " AS CANTIDAD_IMAGENES_PREPARACION, INDICES.COSTODIGITALIZACION * SUM(DOCUMENTOS_DIGITALES_LOTE.CANTIDAD_IMAGENES) AS PAGO"
    Sql = Sql & vbCrLf & " FROM PERSONAL INNER JOIN"
    Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE ON PERSONAL.IDPERSONAL = DOCUMENTOS_DIGITALES_LOTE.FK_PERSONAL_SCANNER RIGHT OUTER JOIN"
    Sql = Sql & vbCrLf & " INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
    Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER >= '" & txtFechaDesde.Text & "' AND  DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER <= '" & txtFechaHasta.Text & "'"
    Sql = Sql & vbCrLf & " GROUP BY DATEPART(YYYY, DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER), DATEPART(MM, DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER), DATEPART(DD,"
    Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER), PERSONAL.NOMBRE, PERSONAL.APELLIDO, INDICES.ID, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_PERSONAL_SCANNER,"
    Sql = Sql & vbCrLf & " INDICES.Descripcion , INDICES.COSTODIGITALIZACION"
    Sql = Sql & vbCrLf & " Having (Sum(DOCUMENTOS_DIGITALES_LOTE.Cantidad_Imagenes) > 0) And (DOCUMENTOS_DIGITALES_LOTE.FK_PERSONAL_SCANNER > 100)"
    Sql = Sql & vbCrLf & " ORDER BY PERSONAL_SCANER, MES, DIA, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS"
Case 58 ' Horarios
    Usuario = InputBox("Ingrese el usuario si es 0 son todo")
    Sql = " SELECT PERSONAL.NOMBRE, PERSONAL.APELLIDO, CONTROLHORARIOS.FK_PERSONAL,CONVERT( nvarchar, CONTROLHORARIOS.FECHA, 103) as fecha , CONVERT (nvarchar, DATEPART ( HH ,   CONTROLHORARIOS.HORA_INGRESO_1)) + ':' + CONVERT (nvarchar, DATEPART ( N ,   CONTROLHORARIOS.HORA_INGRESO_1))  as Hora_Ingreso  ,"
    Sql = Sql & vbCrLf & "  CONVERT (nvarchar, DATEPART ( HH ,   CONTROLHORARIOS.HORA_SALIDA_1)) + ':' + CONVERT (nvarchar, DATEPART ( N ,   CONTROLHORARIOS.HORA_SALIDA_1))  as HORA_SALIDA  , CONVERT (nvarchar, DATEPART ( HH , CONTROLHORARIOS.SUMA_1)) + ':' + CONVERT (nvarchar, DATEPART ( N , CONTROLHORARIOS.SUMA_1 ))  as DIferencia , CONTROLHORARIOS.TIPO_DIA, CONTROLHORARIOS.NOMBRE_DIA"
    Sql = Sql & vbCrLf & "  FROM CONTROLHORARIOS INNER JOIN     PERSONAL ON CONTROLHORARIOS.FK_PERSONAL = PERSONAL.IDPERSONAL"
    Sql = Sql & vbCrLf & "  WHERE CONTROLHORARIOS.FECHA BETWEEN '" & txtFechaDesde.Text & "' AND  '" & txtFechaHasta.Text & "'"
    If Usuario <> 0 Then
        Sql = Sql & vbCrLf & " AND CONTROLHORARIOS.FK_PERSONAL = " & Usuario
    End If
    Sql = Sql & vbCrLf & "  ORDER BY CONTROLHORARIOS.FK_PERSONAL, CONTROLHORARIOS.FECHA"
Case 59 '
    Sql = " SELECT REQUERIMIENTO.IDTIPOREQUERIMIENTO, TIPOREQUERIMIENTO.DESCRIPCION, CONVERT(char, REQUERIMIENTO.FECHARECEPCION, 103) AS FECHA,"
    Sql = Sql & vbCrLf & " REQUERIMIENTO.FK_SUCURSAL, REQUERIMIENTO.CANTIDAD, REQUERIMIENTO.IDESTADO, REQUERIMIENTO_ESTADO.DESCRIPCION AS descrp_estado,"
    Sql = Sql & vbCrLf & " REQUERIMIENTO.IDPERSONAL "
    Sql = Sql & vbCrLf & " FROM REQUERIMIENTO INNER JOIN"
    Sql = Sql & vbCrLf & " TIPOREQUERIMIENTO ON REQUERIMIENTO.IDTIPOREQUERIMIENTO = TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO INNER JOIN"
    Sql = Sql & vbCrLf & " REQUERIMIENTO_ESTADO ON REQUERIMIENTO.IDESTADO = REQUERIMIENTO_ESTADO.ID_ESTADO"
    Sql = Sql & vbCrLf & " WHERE REQUERIMIENTO.FECHARECEPCION BETWEEN " & FechaFormato(txtFechaDesde.Text) & "  AND " & FechaFormato(txtFechaHasta.Text)
    Sql = Sql & vbCrLf & " ORDER BY REQUERIMIENTO.FECHARECEPCION "
Case 60

 Sql = " SELECT REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV, CONVERT(CHAR, REMITOS_CUERPO.FECHA, 103) AS FECHA,"
 Sql = Sql & vbCrLf & " REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.CANTIDAD, REMITO_TIPO.DESCRIPCION AS TIPO_REMITO,"
 Sql = Sql & vbCrLf & " REMITO_OPERACION.DESCRIPCION AS TIPO_OPERACION, REMITO_ESTADOS.DESCRIPCION AS TIPO_ESTADO,TIPO_ALMACENAMIENTO.DESCRIPCION AS TIPO_ELEMENTO"
 Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO INNER JOIN"
 Sql = Sql & vbCrLf & " REMITO_TIPO ON REMITOS_CUERPO.TIPO = REMITO_TIPO.ID INNER JOIN"
 Sql = Sql & vbCrLf & " REMITO_OPERACION ON REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID INNER JOIN"
 Sql = Sql & vbCrLf & " REMITO_ESTADOS ON REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID INNER JOIN"
 Sql = Sql & vbCrLf & " TIPO_ALMACENAMIENTO ON REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = TIPO_ALMACENAMIENTO.ID"
 Sql = Sql & vbCrLf & " WHERE (REMITOS_CUERPO.ANULADO IS NULL) AND "
 Sql = Sql & vbCrLf & " (REMITOS_CUERPO.FECHA between '" & txtFechaDesde.Text & "' and '" & txtFechaHasta.Text & "'  )"

Case 61

    Sql = " SELECT ENTRADA.ID_ENTRADA, ENTRADA.COD_CLIENTE, ENTRADA.ELEMENTO, ENTRADA.TIPO, CONVERT(char, ENTRADA.FECHA, 103) AS FECHA_ENTRADA,"
    Sql = Sql & vbCrLf & " ENTRADA.COD_PERSONAL , ENTRADA.Cod_Estado, LEGAJOS.NRO_CAJA"
    Sql = Sql & vbCrLf & " FROM LEGAJOS RIGHT OUTER JOIN ENTRADA ON LEGAJOS.COD_CLIENTE = ENTRADA.COD_CLIENTE "
    Sql = Sql & vbCrLf & " AND LEGAJOS.ID_CLIENTE_LEGAJO = ENTRADA.ELEMENTO LEFT OUTER JOIN ORDEN_LEGAJOS INNER JOIN"
    Sql = Sql & vbCrLf & " ORDEN_LEGAJOS_DETALLE ON ORDEN_LEGAJOS.ID_ORDEN_LEGAJO = ORDEN_LEGAJOS_DETALLE.COD_ORDEN_LEGAJO ON "
    Sql = Sql & vbCrLf & " ENTRADA.COD_CLIENTE = ORDEN_LEGAJOS.COD_CLIENTE And ENTRADA.Elemento = ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO"
    Sql = Sql & vbCrLf & " WHERE (ENTRADA.FECHA > '" & txtFechaDesde.Text & "' ) AND (ENTRADA.TIPO = 3) AND (ORDEN_LEGAJOS.ID_ORDEN_LEGAJO IS NULL)"
    Sql = Sql & vbCrLf & "  ORDER BY ENTRADA.COD_CLIENTE, LEGAJOS.NRO_CAJA, ENTRADA.ELEMENTO"
    
Case 62

Sql = " SELECT        DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, CONVERT(char, DOCUMENTOS_DIGITALES.NRO_DESDE) AS nro_desde, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
 Sql = Sql & vbCrLf & "                        DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE , DOCUMENTOS_DIGITALES.ID"
Sql = Sql & vbCrLf & " FROM            DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & vbCrLf & "                         DOCUMENTOS_DIGITALES_LOTE ON DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & vbCrLf & " WHERE    FK_CLIENTES =     " & InputBox("Ingrese el cliente")
Sql = Sql & vbCrLf & " and (CONVERT(char, DOCUMENTOS_DIGITALES.NRO_DESDE) LIKE '%" & InputBox("Ingrese la parte del numero") & "%')"
   
Case 63
 Dim LoteDesdeImagenes As Double
 Dim LoteHastaImagenes As Double
 Dim ClientesImagenes As Integer
    
    ClientesImagenes = InputBox("Ingrese el cliente")
    LoteDesdeImagenes = InputBox("Ingrese el lote Desde")
    LoteHastaImagenes = InputBox("Ingrese el lote Hasta")

Sql = "  SELECT  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE AS LOTE, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.CANTIDAD_IMAGENES, DOCUMENTOS_DIGITALES_LOTE.CANTIDAD_ARCHIVOS, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.LETRA_HASTA,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.NRO_HASTA, DOCUMENTOS_DIGITALES.CANTIDAD_IMAGENES AS CANTIDAD_IMAGENES_ARCHIVO, DOCUMENTOS_DIGITALES.ESTADO , DOCUMENTOS_DIGITALES.ID AS ID_IMAGEN"
Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE ON DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & vbCrLf & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = " & ClientesImagenes
Sql = Sql & vbCrLf & " AND (DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE BETWEEN " & LoteDesdeImagenes & " AND " & LoteHastaImagenes & ")"
Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES.LOTE , ID_IMAGEN "

Case 64

MsgBox "verifique si ingreso las fechas "
Sql = "  SELECT    REMITO, CANTIDAD_IMAGENES, FK_CAJAS, ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & vbCrLf & " From DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & vbCrLf & " WHERE    FK_CLIENTES = " & InputBox("Ingrese el cliente")
Sql = Sql & vbCrLf & " AND FECHA_SCANNER BETWEEN '" & txtFechaDesde.Text & "' AND  '" & txtFechaHasta.Text & "'"

Case 65

Sql = " SELECT        REQUERIMIENTO.IDREQUERIMIENTO, TIPOREQUERIMIENTO.DESCRIPCION AS TIPO_REQUERIMIENTO_DESCRIPCION, REQUERIMIENTO.FECHARECEPCION, REQUERIMIENTO.CANTIDAD,"
Sql = Sql & vbCrLf & " REQUERIMIENTO.Cantidad_Imagenes , REQUERIMIENTO.HORA_ARCHIVISTA"
Sql = Sql & vbCrLf & " FROM REQUERIMIENTO LEFT OUTER JOIN  TIPOREQUERIMIENTO ON REQUERIMIENTO.IDTIPOREQUERIMIENTO = TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO"
Sql = Sql & vbCrLf & " WHERE (REQUERIMIENTO.IDTIPOREQUERIMIENTO IN (13)) "
Sql = Sql & vbCrLf & " AND (REQUERIMIENTO.ANULADO IS NULL) "
Sql = Sql & vbCrLf & " AND REQUERIMIENTO.FECHARECEPCION BETWEEN '" & txtFechaDesde.Text & "' AND  '" & txtFechaHasta.Text & "'"
Sql = Sql & vbCrLf & " ORDER BY TIPO_REQUERIMIENTO_DESCRIPCION, REQUERIMIENTO.IDREQUERIMIENTO"


End Select
        rs.Open Sql, strConBasa, adOpenDynamic, adLockReadOnly
        Set grdCargaLegajos.DataSource = rs.DataSource
Exit Sub
salir:
    MsgBox Err.Description
End Sub

Private Sub cmdInformeDetallado_Click()
 Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

        Dim Sql As String
        Sql = " SELECT     CONVERT(char, LEGAJOS.FECHA_CREACION, 103) AS FECHACARGA, LEGAJOS.FK_PERSONAL_CREACION AS Personal, LEGAJOS.COD_CLIENTE,"
        Sql = Sql & " LEGAJOS.NRO_CAJA, COUNT(*) AS CANTIDADLEGAJOS, SUM(CANTIDAD_CARACTERES) AS CARACTERES , PERSONAL.NOMBRE, PERSONAL.APELLIDO"
        Sql = Sql & " FROM         LEGAJOS INNER JOIN"
        Sql = Sql & "  PERSONAL ON LEGAJOS.FK_PERSONAL_CREACION = PERSONAL.IDPERSONAL"
        Sql = Sql & "  WHERE  LEGAJOS.FECHA_CREACION > '" & InputBox("Ingrese la fecha de inicio") & "'"
        Sql = Sql & "  GROUP BY CONVERT(char, LEGAJOS.FECHA_CREACION, 103), LEGAJOS.FK_PERSONAL_CREACION, LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA,"
        Sql = Sql & "  Personal.Nombre , Personal.Apellido"
        Sql = Sql & "  ORDER BY CONVERT(char, LEGAJOS.FECHA_CREACION, 103), LEGAJOS.FK_PERSONAL_CREACION, LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA"



rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
Set grdCargaLegajos.DataSource = rs.DataSource
End Sub

Private Sub Command3_Click()
CopiarDatosGrilla grdCargaLegajos
End Sub

Private Sub Command4_Click()


  Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

        Dim Sql As String
        
Sql = " SELECT     CONVERT(char, LEGAJOS.FECHA_CREACION, 103) AS FECHACARGA, LEGAJOS.FK_PERSONAL_CREACION AS Expr1, COUNT(*) AS CANTIDADLEGAJOS,"
Sql = Sql & "  Personal.Nombre , Personal.Apellido"
Sql = Sql & " FROM         LEGAJOS INNER JOIN"
                      Sql = Sql & " PERSONAL ON LEGAJOS.FK_PERSONAL_CREACION = PERSONAL.IDPERSONAL"
Sql = Sql & " WHERE     LEGAJOS.FECHA_CREACION > '" & InputBox("Ingrese la fecha de inicio") & "'"
Sql = Sql & " GROUP BY CONVERT(char, LEGAJOS.FECHA_CREACION, 103), LEGAJOS.FK_PERSONAL_CREACION, PERSONAL.NOMBRE, PERSONAL.APELLIDO"
Sql = Sql & " ORDER BY LEGAJOS.FK_PERSONAL_CREACION, CONVERT(char, LEGAJOS.FECHA_CREACION, 103)"



rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
Set grdCargaLegajos.DataSource = rs.DataSource

End Sub




Private Sub cmdInformeDetalladoIndice_Click()
Dim rs As New ADODB.Recordset

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

        Dim Sql As String
        
        
        Sql = "  SELECT    LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA, INDICES.DESCRIPCION, COUNT(*) AS CANTIDADLEGAJOS, SUM(LEGAJOS.CANTIDAD_CARACTERES)"
        Sql = Sql & "   AS CARACTERES, CONVERT(char, LEGAJOS.FECHA_CREACION, 103) AS FECHACARGA, LEGAJOS.FK_PERSONAL_CREACION AS Personal"
        Sql = Sql & "   FROM         LEGAJOS INNER JOIN"
        Sql = Sql & "                      PERSONAL ON LEGAJOS.FK_PERSONAL_CREACION = PERSONAL.IDPERSONAL INNER JOIN"
        Sql = Sql & "                    INDICES ON LEGAJOS.FK_INDICES = INDICES.ID"
        Sql = Sql & "   WHERE     (LEGAJOS.FECHA_CREACION BETWEEN CONVERT(DATETIME, '2009-01-01 00:00:00', 102) AND CONVERT(DATETIME, '2009-03-30 00:00:00', 102))"
        Sql = Sql & "   GROUP BY LEGAJOS.COD_CLIENTE, INDICES.DESCRIPCION, LEGAJOS.NRO_CAJA, CONVERT(char, LEGAJOS.FECHA_CREACION, 103),"
        Sql = Sql & "                      LEGAJOS.FK_PERSONAL_CREACION"
                      
                      

        Sql = " SELECT   CONVERT(char, LEGAJOS.FECHA_CREACION, 103) AS FECHACARGA, LEGAJOS.FK_PERSONAL_CREACION AS Personal, LEGAJOS.COD_CLIENTE,"
        Sql = Sql & "  LEGAJOS.NRO_CAJA, COUNT(*) AS CANTIDADLEGAJOS, SUM(LEGAJOS.CANTIDAD_CARACTERES) AS CARACTERES, PERSONAL.NOMBRE,"
        Sql = Sql & "  Personal.Apellido , INDICES.DESCRIPCION"
        Sql = Sql & "  FROM         LEGAJOS INNER JOIN"
        Sql = Sql & "  PERSONAL ON LEGAJOS.FK_PERSONAL_CREACION = PERSONAL.IDPERSONAL INNER JOIN"
        Sql = Sql & " INDICES ON LEGAJOS.FK_INDICES = INDICES.ID"
        Sql = Sql & "  WHERE     (LEGAJOS.FECHA_CREACION > '01/01/2009')"
        Sql = Sql & "  GROUP BY CONVERT(char, LEGAJOS.FECHA_CREACION, 103), LEGAJOS.FK_PERSONAL_CREACION, LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA,"
        Sql = Sql & "  Personal.Nombre , Personal.Apellido, INDICES.DESCRIPCION"
        Sql = Sql & "  ORDER BY CONVERT(char, LEGAJOS.FECHA_CREACION, 103), LEGAJOS.FK_PERSONAL_CREACION, LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA"


rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
Set grdCargaLegajos.DataSource = rs.DataSource
End Sub

Private Sub cmdRecalcularCaracteres_Click()
Legajos_RecalcularCaracteres_DescripcionRemito " where FECHA_CREACION >  '" & InputBox("Fecha de Inicio") & "'"
End Sub


Private Sub Command1_Click()
Legajos_RecalcularCaracteres_DescripcionRemito (" WHERE COD_CLIENTE = 231 AND DESCRIPCION_REMITO IS NULL ")
End Sub

Private Sub ControlCajas()


Dim ConBasa As New ADODB.Connection
Dim Sql As String

ConBasa.Open strConBasa


Dim rs As New ADODB.Recordset

Dim RsControl As New ADODB.Recordset

rs.CursorLocation = adUseClient




ExecutarSql "DELETE FROM TEM_CONTROL_CAJAS"
Sql = " INSERT INTO TEM_CONTROL_CAJAS"
Sql = Sql & " (FK_LECTURA, FK_CAJA, FK_CLIENTE, ORDEN)"
Sql = Sql & "  SELECT     NUMERO_LECTURA, CAJA, CLIENTE, ORDEN"
Sql = Sql & "  From LECTURACOLECTOR "
Sql = Sql & "  WHERE     (NUMERO_LECTURA IN ( " & InputBox("Ingrese los numeros de lectura separados ,", "", 0) & "))"
Sql = Sql & " ORDER BY NUMERO_LECTURA, ORDEN"

ExecutarSql Sql


Sql = " SELECT FK_LECTURA, FK_CLIENTE, FK_CAJA, ORDEN, REMITO_VACIAS, REMITOS_CUSTODIA, REFERENCIAS_RANGO, REFERENCIA_LEGAJOS,"
Sql = Sql & "  PERSONAL_ASIGNADO, TIPO_REFERENCIA  From TEM_CONTROL_CAJAS "


rs.Open Sql, ConActiva, adOpenKeyset, adLockOptimistic


Do While Not rs.EOF
    
    Sql = "  SELECT REMITOS_CUERPO.NRO_REMITO , ID_CLIENTE "
    Sql = Sql & " FROM REMITOS_CUERPO INNER JOIN "
    Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO "
    Sql = Sql & " WHERE     (REMITOS_CUERPO.TIPO = 0) AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) "
    If rs!FK_CAJA < 740000 Then
        Sql = Sql & " AND REMITOS_CUERPO.ID_CLIENTE = " & rs!FK_CLIENTE
    End If
       Sql = Sql & " AND  REMITOS_DETALLE.DESDE = " & rs!FK_CAJA
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, adOpenKeyset, adLockPessimistic
       If Not RsControl.EOF Then
            rs!REMITOS_CUSTODIA = RsControl!NRO_REMITO
            rs!FK_CLIENTE = RsControl!id_cliente
       End If
       
       
       
    Sql = "  SELECT REMITOS_CUERPO.NRO_REMITO "
    Sql = Sql & " FROM REMITOS_CUERPO INNER JOIN "
    Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO "
    Sql = Sql & " WHERE     (REMITOS_CUERPO.TIPO = 2) AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) "
    If rs!FK_CAJA < 740000 Then
        Sql = Sql & "   AND REMITOS_CUERPO.ID_CLIENTE = " & rs!FK_CLIENTE
    End If
        Sql = Sql & " AND  REMITOS_DETALLE.DESDE = " & rs!FK_CAJA
        
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, 0, 1
       If Not RsControl.EOF Then
        rs!REMITO_VACIAS = RsControl!NRO_REMITO
       End If
       
       
       
       Sql = " SELECT COUNT(*) AS CANTIDADREFERENCIA "
        Sql = Sql & " From REFERENCIAS "
        
        Sql = Sql & " Where NRO_CAJA = " & rs!FK_CAJA
       
       If rs!FK_CAJA < 740000 Then
         Sql = Sql & " And COD_CLIENTE = " & rs!FK_CLIENTE
    End If
        
       
       
       
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, 0, 1
       If Not RsControl.EOF Then
            rs!REFERENCIAS_RANGO = RsControl!CANTIDADREFERENCIA
       Else
            rs!REFERENCIAS_RANGO = 0
       End If
       
       
        Sql = " SELECT COUNT(*) AS CantidadLegajos "
        Sql = Sql & " From LEGAJOS"
        Sql = Sql & " Where NRO_CAJA = " & rs!FK_CAJA
         If rs!FK_CAJA < 740000 Then
         Sql = Sql & " And COD_CLIENTE = " & rs!FK_CLIENTE
    End If
        
        
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, 0, 1
       If Not RsControl.EOF Then
            rs!REFERENCIA_LEGAJOS = RsControl!CANTIDADLEGAJOS
       Else
            rs!REFERENCIA_LEGAJOS = 0
       End If
       
       
       
       
       
       
      
        Sql = " SELECT     CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_PERSONAL_ENTREGA, TIPO_REFERENCIA "
        Sql = Sql & " FROM CAJAS  "
          Sql = Sql & " Where "
        If rs!FK_CAJA < 740000 Then
            Sql = Sql & "   Cajas.FK_CLIENTE = " & rs!FK_CLIENTE & " AND "
         End If
        Sql = Sql & "  Cajas.NRO_CAJA = " & rs!FK_CAJA
       
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, 0, 1
       If Not RsControl.EOF Then
            rs!PERSONAL_ASIGNADO = RsControl!FK_PERSONAL_ENTREGA
            rs!TIPO_REFERENCIA = RsControl!TIPO_REFERENCIA
       Else
            rs!PERSONAL_ASIGNADO = 0
            rs!TIPO_REFERENCIA = "'0'"
       End If
       
       
       
    rs.Update
    rs.MoveNext
Loop






End Sub



Private Sub Command2_Click()
Dim i As Integer

For i = 1 To grdCargaLegajos.VisibleRows

MsgBox grdCargaLegajos.SelBookmarks(i)
Next


End Sub

Private Sub Form_Load()
txtFechaDesde.Text = Format(DateAdd("d", -10, Now), "dd/mm/yyyy")
txtFechaHasta.Text = Format(Now, "dd/mm/yyyy")
CargarTipoReferencias
cboTipoReferencia.ListIndex = 1
End Sub

Public Sub CargarTipoReferencias()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Sql = "SELECT      ID_PARAMETRO, DESCRIPCION, TABLA, CAMPO_NOMBRE"
    Sql = Sql & " From basasql.dbo.PARAMETROS "
    Sql = Sql & " WHERE     (CAMPO_NOMBRE = 'FK_TIPO_REFERENCIA')"
    Sql = Sql & " ORDER BY ID_PARAMETRO "
    
    rs.Open Sql, strConBasa
    
    
Do While Not rs.EOF
    cboTipoReferencia.AddItem rs!ID_PARAMETRO & "-" & Trim(rs!Descripcion)
    rs.MoveNext
Loop


End Sub



Public Sub InformeSupervielle()

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rsLegajos As New ADODB.Recordset
Dim rsRequerimiento As New ADODB.Recordset
Dim P As Integer
P = InputBox("Ingrese el nivel")


Dim FORMA As String
Dim REQUERIMIENTO As Long
Dim NRO_REMITO As Long
Dim NRO_REM_PROV As String
Dim TIPO As String
Dim SUBTIPO As String
Dim fecha As String
Dim OBSERVACIONES As String
Dim CANTIDADVACIAS As Long
Dim CANTIDADGUARDAYCUSTODIA As Long
Dim CANTIDADCONSULTAS As Long
Dim CANTIDADLEGAJOS As Long
Dim CANTIDADFLETESNORMALES As Long
Dim CANTIDADFLETESURGENTES As Long
Dim cantidadImagenes As Long
Dim CANTIDADRETIROS As Long
Dim CANTIDADHORASARCHIVISTA As Long
Dim APELLIDO_NOMBRE As String
Dim PROVINCIA As String
Dim Sucursal As String
Dim COBRAR As Integer
Dim FK_CLIENTE As Integer
Dim RETIROSFUERADERADIO As Integer
Dim ENVIOFUERADERADIO As Integer
Dim Precintos As Integer
Dim Retiro As Integer
Dim PASO_IMAGEN  As String


Dim Cliente As Integer
Dim FECHA_DESDE  As String


Cliente = InputBox("ingrese el cliente")

FECHA_DESDE = txtFechaDesde.Text


ExecutarSql "DELETE FROM TEM_SUPERVIELLE"

'_____________________ Remitos _____________________________________________
        Sql = " SELECT     REMITOS_CUERPO.ID_CLIENTE AS FK_CLIENTE , REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.IMAGEN AS PASO_IMAGEN ,  REMITOS_CUERPO.NRO_REM_PROV,CONVERT ( CHAR,  REMITOS_CUERPO.FECHA ,103)AS FECHA,"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.CANTIDAD, CLIENTEUSUARIO.APELLIDO_NOMBRE, INDICES_1.DESCRIPCION AS PROVINCIA, INDICES_2.DESCRIPCION AS SUCURSAL,"
        Sql = Sql & vbCrLf & " TIPO_ALMACENAMIENTO.DESCRIPCION + ' ' + TIPO_REMITO.DESCRIPCION + '  ' + REMITO_OPERACION.DESCRIPCION + ' ' + REMITO_ESTADOS.DESCRIPCION AS SERVICIO,"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO,  TIPO_REMITO.DESCRIPCION AS TIPO, REMITOS_CUERPO.OPERACION, REMITOS_CUERPO.ESTADO,"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.COBRAR_FLETE , REMITO_ESTADOS.DESCRIPCION AS ESTADO , OBSERVACIONES ,COBRAR_FLETE AS  COBRAR , HORASARCHIVISTA "
        Sql = Sql & vbCrLf & " FROM         REMITOS_CUERPO LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " REMITO_ESTADOS ON REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " TIPO_REMITO ON REMITOS_CUERPO.TIPO = TIPO_REMITO.ID LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " REMITO_OPERACION ON REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " TIPO_ALMACENAMIENTO ON REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = TIPO_ALMACENAMIENTO.ID LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " INDICES AS INDICES_1 INNER JOIN"
        Sql = Sql & vbCrLf & " INDICES AS INDICES_2 INNER JOIN"
        Sql = Sql & vbCrLf & " CLIENTEUSUARIO ON INDICES_2.INDICE = CLIENTEUSUARIO.COD_INDICE AND INDICES_2.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE ON"
        Sql = Sql & vbCrLf & " INDICES_1.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND INDICES_1.INDICE = SUBSTRING(CLIENTEUSUARIO.COD_INDICE, 1," & P & ") ON"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
        Sql = Sql & vbCrLf & " WHERE (REMITOS_CUERPO.ID_CLIENTE IN (" & Cliente & "))"
        Sql = Sql & vbCrLf & " AND REMITOS_CUERPO.FECHA BETWEEN " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text)
        Sql = Sql & vbCrLf & " AND (REMITOS_CUERPO.ANULADO IS NULL)"
        Sql = Sql & vbCrLf & " ORDER BY TIPO_REMITO.DESCRIPCION, REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO, REMITOS_CUERPO.TIPO, REMITOS_CUERPO.OPERACION,"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.estado "
         Set rs = New ADODB.Recordset
        rs.Open Sql, strConBasa, 0, 1
        Do While Not rs.EOF
            FORMA = "'Remito'"
            REQUERIMIENTO = 0
            NRO_REMITO = rs!NRO_REMITO
            Retiro = 0
            Precintos = 0
            RETIROSFUERADERADIO = 0
            ENVIOFUERADERADIO = 0
            CANTIDADFLETESNORMALES = 0
            CANTIDADFLETESURGENTES = 0
            
            
            If rs!NRO_REM_PROV = "0001-000_____" Then
                NRO_REM_PROV = 0
                If Not IsNull(rs!PASO_IMAGEN) Then
                    PASO_IMAGEN = "'\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REMITO & "\" & rs!PASO_IMAGEN & "'"
                 Else
                    PASO_IMAGEN = "NULL"
                 End If
            Else
                NRO_REM_PROV = "'" & Trim(rs!NRO_REM_PROV) & "'"
                 If Not IsNull(rs!PASO_IMAGEN) Then
                    PASO_IMAGEN = "'\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & Replace(NRO_REM_PROV, "'", "") & "\" & rs!PASO_IMAGEN & "'"
                 Else
                   PASO_IMAGEN = "NULL"
                 End If
            End If
            TIPO = "'" & rs!TIPO & "'"
            SUBTIPO = "'" & rs!SERVICIO & "'"
            fecha = FechaFormato(rs!fecha)
            OBSERVACIONES = "'" & rs!OBSERVACIONES & "'"
            If rs!TIPO = "VACIAS" Then
                CANTIDADVACIAS = rs!cantidad
                    If rs!estado = "NORMAL" And rs!COBRAR = 1 Then
                        CANTIDADFLETESNORMALES = 1
                    Else
                        CANTIDADFLETESNORMALES = 0
                    End If
                    If rs!estado = "URGENTE" And rs!COBRAR = 1 Then
                        CANTIDADFLETESURGENTES = 1
                    Else
                        CANTIDADFLETESURGENTES = 0
                    End If
                Else
                CANTIDADVACIAS = 0
            End If
            If rs!TIPO = "GUARDA Y CUSTODIA" Then
                CANTIDADGUARDAYCUSTODIA = rs!cantidad
                
            Else
                CANTIDADGUARDAYCUSTODIA = 0
            End If
            If rs!estado = "NORMAL" And rs!COBRAR = 1 Then
                    CANTIDADFLETESNORMALES = 1
                Else
                    CANTIDADFLETESNORMALES = 0
                End If
                If rs!estado = "URGENTE" And rs!COBRAR = 1 Then
                    CANTIDADFLETESURGENTES = 1
                Else
                    CANTIDADFLETESURGENTES = 0
                End If
            If rs!TIPO = "CONSULTA" Then
                CANTIDADCONSULTAS = rs!cantidad
            Else
                CANTIDADCONSULTAS = 0
            End If
            CANTIDADLEGAJOS = 0
            
            cantidadImagenes = 0
            CANTIDADRETIROS = 0
            If IsNull(rs!HORASARCHIVISTA) Then
                CANTIDADHORASARCHIVISTA = 0
            Else
                CANTIDADHORASARCHIVISTA = rs!HORASARCHIVISTA
            End If
            APELLIDO_NOMBRE = "'" & Trim(rs!APELLIDO_NOMBRE) & "'"
            PROVINCIA = "'" & Trim(rs!PROVINCIA) & "'"
            Sucursal = "'" & rs!Sucursal & "'"
            FK_CLIENTE = rs!FK_CLIENTE
            
            
             If rs!TIPO = "CONSULTA" And CANTIDADCONSULTAS = 0 Then
                MsgBox "DDDD"
             End If
             
            
            INSERTAR_FACTURA_SUPER FORMA, REQUERIMIENTO, NRO_REMITO _
            , NRO_REM_PROV, TIPO, SUBTIPO, fecha, OBSERVACIONES _
            , CANTIDADVACIAS, CANTIDADGUARDAYCUSTODIA, CANTIDADCONSULTAS, CANTIDADLEGAJOS _
            , CANTIDADFLETESNORMALES, CANTIDADFLETESURGENTES, cantidadImagenes _
            , CANTIDADHORASARCHIVISTA, APELLIDO_NOMBRE, PROVINCIA, Sucursal _
            , COBRAR, FK_CLIENTE, RETIROSFUERADERADIO, ENVIOFUERADERADIO, Precintos, Retiro, PASO_IMAGEN, "null"
            rs.MoveNext
        Loop
'--------------------------------fin remitos-----------------------------------
   
   
    
 '------------------------------CARGA DE LEGAJOS_______________________________________
        Sql = " SELECT  INDICES.DESCRIPCION, INDICES.INDICE, COUNT(DISTINCT LEGAJOS.NRO_CAJA) AS CANTIDAD_CAJAS, COUNT(*) AS CANTIDAD_LEGAJOS,"
        Sql = Sql & vbCrLf & " INDICES_1.DESCRIPCION AS PROVINCIA, INDICES_2.DESCRIPCION AS SUCURSAL "
        Sql = Sql & vbCrLf & " FROM LEGAJOS INNER JOIN "
        Sql = Sql & vbCrLf & " INDICES ON LEGAJOS.COD_CLIENTE = INDICES.COD_CLIENTE AND LEGAJOS.COD_INDICE = INDICES.INDICE LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " INDICES INDICES_2 ON SUBSTRING(LEGAJOS.COD_INDICE, 1, 6) = INDICES_2.INDICE AND"
        Sql = Sql & vbCrLf & " LEGAJOS.COD_CLIENTE = INDICES_2.COD_CLIENTE LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " INDICES INDICES_1 ON LEGAJOS.COD_CLIENTE = INDICES_1.COD_CLIENTE AND SUBSTRING(LEGAJOS.COD_INDICE, 1," & P & ") = INDICES_1.INDICE "
        Sql = Sql & vbCrLf & " WHERE   LEGAJOS.COD_CLIENTE in  ( " & Cliente & ")"
        Sql = Sql & vbCrLf & " AND LEGAJOS.FECHA_CREACION BETWEEN " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text)
        Sql = Sql & vbCrLf & " GROUP BY INDICES.DESCRIPCION, INDICES.INDICE, INDICES_1.DESCRIPCION, INDICES_2.DESCRIPCION, INDICES.TIPO_INDICE"
        Sql = Sql & vbCrLf & " HAVING      (INDICES.TIPO_INDICE LIKE 'LEGAJO') "
        Set rs = New ADODB.Recordset
        rs.Open Sql, strConBasa, 0, 1
        Do While Not rs.EOF
                   FORMA = "'CARGA DE LEGAJOS'"
                   REQUERIMIENTO = 0
                   NRO_REMITO = 0
                   NRO_REM_PROV = 0
                   TIPO = "'CARGA DE LEGAJOS'"
                   SUBTIPO = "'CARGA DE LEGAJOS'"
                   fecha = FechaFormato(FECHA_DESDE)
                   OBSERVACIONES = "NULL"
                   CANTIDADVACIAS = 0
                   Precintos = 0
                   CANTIDADGUARDAYCUSTODIA = 0
                   CANTIDADCONSULTAS = 0
                   CANTIDADLEGAJOS = rs!CANTIDAD_LEGAJOS
                   CANTIDADFLETESNORMALES = 0
                   CANTIDADFLETESURGENTES = 0
                   cantidadImagenes = 0
                   CANTIDADRETIROS = 0
                   CANTIDADHORASARCHIVISTA = 0
                   Retiro = 0
                   APELLIDO_NOMBRE = "NULL"
                   PROVINCIA = "'" & Trim(rs!PROVINCIA) & "'"
                   Sucursal = "'" & rs!Sucursal & "'"
                   FK_CLIENTE = Cliente
                   INSERTAR_FACTURA_SUPER FORMA, REQUERIMIENTO, NRO_REMITO _
                   , NRO_REM_PROV, TIPO, SUBTIPO, fecha, OBSERVACIONES _
                   , CANTIDADVACIAS, CANTIDADGUARDAYCUSTODIA, CANTIDADCONSULTAS, CANTIDADLEGAJOS _
                   , CANTIDADFLETESNORMALES, CANTIDADFLETESURGENTES, cantidadImagenes _
                   , CANTIDADHORASARCHIVISTA, APELLIDO_NOMBRE, PROVINCIA, Sucursal _
                   , COBRAR, FK_CLIENTE, RETIROSFUERADERADIO, ENVIOFUERADERADIO, Precintos, Retiro, "NULL", "NULL"
                   rs.MoveNext
               Loop
'---------------------------------------FIN CARGA DE LEGAJOS_____________________________

'________________________________________ Requerimiento _____________________________________
        
        
            Sql = "  SELECT   REQUERIMIENTO.ID_CLIENTE AS FK_CLIENTE ,   REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.IDREMITO, CONVERT ( CHAR ,  REQUERIMIENTO.FECHARECEPCION , 103) AS FECHA ,"
            Sql = Sql & " REQUERIMIENTO.DESCRIPCION AS OBSERVACIONES, INDICES_1.DESCRIPCION AS PROVINCIA, INDICES_2.DESCRIPCION AS SUCURSAL,"
            Sql = Sql & " TIPOREQUERIMIENTO.DESCRIPCION AS SERVICIO, CLIENTEUSUARIO.APELLIDO_NOMBRE, REQUERIMIENTO.CANTIDAD_IMAGENES,"
            Sql = Sql & " REQUERIMIENTO.CANTIDAD, REMITOS_CUERPO.OBSERVACIONES AS REMITO_DESCRIPCION,"
            Sql = Sql & " REQUERIMIENTO_ESTADO.DESCRIPCION AS ESTADO, REQUERIMIENTO.HORA_ARCHIVISTA, REQUERIMIENTO.FLETE,REQUERIMIENTO.COBRAR , REQUERIMIENTO.IDTIPOREQUERIMIENTO "
            Sql = Sql & vbCrLf & " FROM         TIPOREQUERIMIENTO RIGHT OUTER JOIN"
            Sql = Sql & " REMITOS_CUERPO RIGHT OUTER JOIN"
            Sql = Sql & " REQUERIMIENTO LEFT OUTER JOIN"
            Sql = Sql & " REQUERIMIENTO_ESTADO ON REQUERIMIENTO.IDESTADO = REQUERIMIENTO_ESTADO.ID_ESTADO ON"
            Sql = Sql & " REMITOS_CUERPO.NRO_REMITO = REQUERIMIENTO.IDREMITO ON"
            Sql = Sql & vbCrLf & " TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO = REQUERIMIENTO.IDTIPOREQUERIMIENTO LEFT OUTER JOIN"
            Sql = Sql & " INDICES INDICES_2 RIGHT OUTER JOIN"
            Sql = Sql & " INDICES INDICES_1 RIGHT OUTER JOIN"
            Sql = Sql & " CLIENTEUSUARIO ON INDICES_1.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND"
            Sql = Sql & " INDICES_1.INDICE = SUBSTRING(CLIENTEUSUARIO.COD_INDICE, 1," & P & ") ON INDICES_2.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND"
            Sql = Sql & " INDICES_2.INDICE = CLIENTEUSUARIO.COD_INDICE ON REQUERIMIENTO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
            Sql = Sql & " WHERE    (REQUERIMIENTO.ID_CLIENTE IN( " & Cliente & ")) "
            Sql = Sql & " AND REQUERIMIENTO. FECHAENTREGA BETWEEN " & FechaFormato(txtFechaDesde.Text) & " AND " & FechaFormato(txtFechaHasta.Text)
            Sql = Sql & " AND (REQUERIMIENTO.ANULADO IS NULL)  AND (REQUERIMIENTO.IDTIPOREQUERIMIENTO IN ( 5,6, 8,9,13,14,18,19,20,22,23,24)) "
            Sql = Sql & " ORDER BY INDICES_1.DESCRIPCION, INDICES_2.DESCRIPCION, REQUERIMIENTO.IDTIPOREQUERIMIENTO"
            Set rs = New ADODB.Recordset

'5   Pedido de Referencia y Trasvase
' 6 retiros NULL    NULL    NULL
'8   Busqueda de Documentación                                                                               NULL    NULL    NULL
'9   Consulta en Planta                                                                                      1   0   0
'13  Consulta Digital                                                                                        1   NULL    NULL
'14  Consulta Por Fax                                                                                        1   NULL    NULL
'18  Busqueda de documentos                                                                                  NULL    NULL    NULL
'19  Horas de archivista                                                                                     NULL    NULL    NULL
'20  Retiros especiales fuera de radio                                                                       NULL    NULL    NULL
'23  Venta de precintos                                                                                      NULL    NULL    NULL
'24  Envio especiales fuera de radio                                                                         NULL    NULL    NULL
            
            Dim Paso_Requerimiento_Imagenes As String
            
            rs.Open Sql, strConBasa, 0, 1
          Do While Not rs.EOF
          
          RETIROSFUERADERADIO = 0
           ENVIOFUERADERADIO = 0
          Retiro = 0
          Precintos = 0
            FORMA = "'Requerimientos'"
            REQUERIMIENTO = rs!IDREQUERIMIENTO
            
            
                PASO_IMAGEN = "NULL"
            
            
            
            
            NRO_REMITO = 0
            NRO_REM_PROV = 0
            TIPO = "'" & Trim(UCase(rs!SERVICIO)) & "'"
            SUBTIPO = "'" & Trim(UCase(rs!SERVICIO)) & "'"
            fecha = FechaFormato(rs!fecha)
            OBSERVACIONES = "'" & Mid(rs!OBSERVACIONES, 1, 50) & "'"
            CANTIDADVACIAS = 0
            CANTIDADGUARDAYCUSTODIA = 0
            CANTIDADCONSULTAS = 0
            
            Select Case rs!IDTIPOREQUERIMIENTO
            
            Case 5 Or 8 Or 9 Or 14
                CANTIDADCONSULTAS = rs!cantidad
            Case Else
              Rem   CANTIDADCONSULTAS = rs!cantidad
            End Select
            
             If rs!IDTIPOREQUERIMIENTO = 18 Then
             CANTIDADCONSULTAS = rs!cantidad
             End If
            
             If rs!IDTIPOREQUERIMIENTO = 6 Then
             Retiro = 1
             End If
             
            If rs!IDTIPOREQUERIMIENTO = 23 Then ' 23  Venta de precintos
                Precintos = rs!cantidad
            End If
            
            
            If rs!IDTIPOREQUERIMIENTO = 20 Then ' 20  Retiros especiales fuera de radio
                 RETIROSFUERADERADIO = rs!cantidad

            End If
            
            If rs!IDTIPOREQUERIMIENTO = 24 Then '24  Envio especiales fuera de radio
                 ENVIOFUERADERADIO = rs!cantidad
            End If
            

            
            
            If rs!IDTIPOREQUERIMIENTO = 13 Then ' CONSULTAS DIGITALES
                CANTIDADCONSULTAS = rs!cantidad
                If IsNull(rs!Cantidad_Imagenes) Then
                MsgBox "ATENCION EL REQUERIMIENTO: " & rs!IDREQUERIMIENTO & "NO TIENE IMAGENES", vbCritical
                 cantidadImagenes = 0
                Else
                cantidadImagenes = rs!Cantidad_Imagenes
                End If
            Else
                
                cantidadImagenes = 0
            End If
            CANTIDADLEGAJOS = 0
            
            If rs!IDTIPOREQUERIMIENTO = 6 Then ' RETIROS NORMALES
               CANTIDADFLETESNORMALES = rs!cantidad
            Else
              CANTIDADFLETESNORMALES = 0
            End If
            If rs!IDTIPOREQUERIMIENTO = 22 Or rs!IDTIPOREQUERIMIENTO = 21 Or rs!IDTIPOREQUERIMIENTO = 6 Then    ' RETIROS de cajas y legajos
               CANTIDADFLETESNORMALES = 1
            Else
              CANTIDADFLETESNORMALES = 0
            End If

            CANTIDADFLETESURGENTES = 0
            
            If rs!IDTIPOREQUERIMIENTO = 20 Then ' RETIROS ESPECIALES
               CANTIDADRETIROS = rs!cantidad
            Else
              CANTIDADRETIROS = 0
            End If
            
            
            If Not IsNull(rs!HORA_ARCHIVISTA) Then  ' HORAS DE ARCHIVISTA
                CANTIDADHORASARCHIVISTA = rs!HORA_ARCHIVISTA
            Else
              CANTIDADHORASARCHIVISTA = 0
            End If
            
            APELLIDO_NOMBRE = "'" & Trim(rs!APELLIDO_NOMBRE) & "'"
            PROVINCIA = "'" & Trim(rs!PROVINCIA) & "'"
            Sucursal = "'" & rs!Sucursal & "'"
            FK_CLIENTE = rs!FK_CLIENTE
            
            INSERTAR_FACTURA_SUPER FORMA, REQUERIMIENTO, NRO_REMITO _
            , NRO_REM_PROV, TIPO, SUBTIPO, fecha, OBSERVACIONES _
            , CANTIDADVACIAS, CANTIDADGUARDAYCUSTODIA, CANTIDADCONSULTAS, CANTIDADLEGAJOS _
            , CANTIDADFLETESNORMALES, CANTIDADFLETESURGENTES, cantidadImagenes _
            , CANTIDADHORASARCHIVISTA, APELLIDO_NOMBRE, PROVINCIA, Sucursal _
            , COBRAR, FK_CLIENTE, RETIROSFUERADERADIO, ENVIOFUERADERADIO, Precintos, Retiro, PASO_IMAGEN, "null"
            rs.MoveNext
            
          Loop
          
          
       '______________________ requerimiento ______________________
          
                      

End Sub
Public Sub INSERTAR_FACTURA_SUPER(FORMA As String, REQUERIMIENTO As Long _
, NRO_REMITO As Long, NRO_REM_PROV As String _
, TIPO As String, SUBTIPO As String _
, fecha As String, OBSERVACIONES As String _
 , CANTIDADVACIAS As Long, CANTIDADGUARDAYCUSTODIA As Long _
, CANTIDADCONSULTAS As Long, CANTIDADLEGAJOS As Long _
, CANTIDADFLETESNORMALES As Long, CANTIDADFLETESURGENTES As Long _
, cantidadImagenes As Long _
, CANTIDADHORASARCHIVISTA As Long, APELLIDO_NOMBRE As String _
, PROVINCIA As String, Sucursal As String _
, COBRAR As Integer, FK_CLIENTE As Integer _
, RETIROSFUERADERADIO As Integer, ENVIOFUERADERADIO As Integer _
, Precintos As Integer, Retiro As Integer, PASO_IMAGEN As String, USUARIO_REMITO As String)

Dim Sql As String
Sql = "  Insert INTO  TEM_SUPERVIELLE("
Sql = Sql & vbCrLf & " FORMA, REQUERIMIENTO"
Sql = Sql & vbCrLf & " , NRO_REMITO, NRO_REM_PROV"
Sql = Sql & vbCrLf & " , TIPO, SUBTIPO"
Sql = Sql & vbCrLf & " , FECHA, OBSERVACIONES"
Sql = Sql & vbCrLf & " , CANTIDADVACIAS , CANTIDADGUARDAYCUSTODIA"
Sql = Sql & vbCrLf & " , CANTIDADCONSULTAS, CANTIDADLEGAJOS"
Sql = Sql & vbCrLf & " , CANTIDADFLETESNORMALES, CANTIDADFLETESURGENTES"
Sql = Sql & vbCrLf & " , CANTIDADIMAGENES"
Sql = Sql & vbCrLf & " , CANTIDADHORASARCHIVISTA, APELLIDO_NOMBRE"
Sql = Sql & vbCrLf & " , PROVINCIA, SUCURSAL"
Sql = Sql & vbCrLf & " , COBRAR, FK_CLIENTE"
Sql = Sql & vbCrLf & " , RETIROSFUERADERADIO , ENVIOFUERADERADIO"
Sql = Sql & vbCrLf & " ,  PRECINTOS , RETIROS , PASO_IMAGEN , USUARIO_REMITO )"
Sql = Sql & vbCrLf & " VALUES ("
Sql = Sql & vbCrLf & FORMA & "," & REQUERIMIENTO
Sql = Sql & vbCrLf & "," & NRO_REMITO & "," & NRO_REM_PROV
Sql = Sql & vbCrLf & "," & TIPO & "," & SUBTIPO
Sql = Sql & vbCrLf & "," & fecha & "," & OBSERVACIONES
Sql = Sql & vbCrLf & " ," & CANTIDADVACIAS & "," & CANTIDADGUARDAYCUSTODIA
Sql = Sql & vbCrLf & "," & CANTIDADCONSULTAS & "," & CANTIDADLEGAJOS
Sql = Sql & vbCrLf & "," & CANTIDADFLETESNORMALES & "," & CANTIDADFLETESURGENTES
Sql = Sql & vbCrLf & "," & cantidadImagenes
Sql = Sql & vbCrLf & "," & CANTIDADHORASARCHIVISTA & "," & APELLIDO_NOMBRE
Sql = Sql & vbCrLf & "," & PROVINCIA & "," & Sucursal
Sql = Sql & vbCrLf & "," & COBRAR & "," & FK_CLIENTE
Sql = Sql & vbCrLf & "," & RETIROSFUERADERADIO & "," & ENVIOFUERADERADIO
Sql = Sql & vbCrLf & "," & Precintos & "," & Retiro & "," & PASO_IMAGEN & "," & USUARIO_REMITO
Sql = Sql & vbCrLf & ")"
ExecutarSql Sql
End Sub


Public Sub CajasSinReferencias(COD_CLIENTE As Integer, Sucursal As Integer)
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rsRef As New ADODB.Recordset
Dim Cliente As Integer
Dim Caja As Long
Dim RsControl As New ADODB.Recordset
Dim i As Long


Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV, REMITOS_CUERPO.TIPO, REMITOS_DETALLE.DESDE,"
Sql = Sql & " CLIENTEUSUARIO.APELLIDO_NOMBRE , INDICES.ID_CODIGO_DOCUMENTO, INDICES.Descripcion, REMITOS_CUERPO.id_cliente,     REMITOS_CUERPO.FECHA"
Sql = Sql & " FROM INDICES INNER JOIN"
Sql = Sql & " CLIENTEUSUARIO ON INDICES.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND INDICES.INDICE = CLIENTEUSUARIO.COD_INDICE RIGHT OUTER JOIN"
Sql = Sql & " REMITOS_CUERPO INNER JOIN"
Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO ON"
Sql = Sql & " CLIENTEUSUARIO.ID_CLIENTEUSUARIO = REMITOS_CUERPO.COD_USUARIO_CLIENTE"
Sql = Sql & "  Where "
Sql = Sql & " (REMITOS_CUERPO.ANULADO IS NULL) "
Sql = Sql & " And (REMITOS_CUERPO.id_cliente = " & COD_CLIENTE & ") "



If Sucursal <> 0 Then
    Sql = Sql & " And (REMITOS_CUERPO.TIPO = 2) "
    Sql = Sql & " And (INDICES.ID_CODIGO_DOCUMENTO =" & Sucursal & ")"
Else
    Sql = Sql & " And (REMITOS_CUERPO.TIPO = 0) "
End If
Sql = Sql & " Order by REMITOS_DETALLE.DESDE "

RsControl.Open Sql, strConBasa



Rem  Sql = "INSERT INTO CONTROL_REFERENCIAS   (ESTADO, FK_CLIENTE, FK_CAJA) SELECT     ESTADO, COD_CLIENTE, NRO_CAJA FROM         CONTENEDOR WHERE     (COD_CLIENTE = 231) AND (ESTANTERIA < 3000) ORDER BY NRO_CAJA  "
ExecutarSql "DELETE  CONTROL_REFERENCIAS "
ExecutarSql Sql


Sql = " SELECT  ID,   FK_CLIENTE, FK_CAJA,REFERENCIAS, ESTADO,  Remito_Manual ,  NRO_REMITO, FECHA, APELLIDO_NOMBRE "
Sql = Sql & " From CONTROL_REFERENCIAS"

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

rs.Open Sql, strConBasa, 3, 2



Do While Not RsControl.EOF

    Cliente = RsControl!id_cliente
    Caja = RsControl!Desde
    i = i + 1
rs.AddNew
rs!ID = i
    rs!FK_CLIENTE = Cliente
    rs!FK_CAJA = Caja
    rs!NRO_REMITO = RsControl!NRO_REMITO
    rs!Remito_Manual = RsControl!NRO_REM_PROV
    rs!fecha = RsControl!fecha
    rs!APELLIDO_NOMBRE = RsControl!APELLIDO_NOMBRE

    rs!REFERENCIAS = ""
    Sql = " SELECT     NRO_CAJA, COD_CLIENTE "
    Sql = Sql & " From REFERENCIAS "
    Sql = Sql & " WHERE  NRO_CAJA = " & Caja
    Sql = Sql & " AND COD_CLIENTE = " & Cliente
    
    Set rsRef = New ADODB.Recordset
    rsRef.Open Sql, ConActiva, 0, 1
    If Not rsRef.EOF Then
    
         rs!REFERENCIAS = "rango"
    
    
    End If
    

Sql = "  SELECT     NRO_CAJA, COD_CLIENTE"
Sql = Sql & " From LEGAJOS"
Sql = Sql & " Where NRO_CAJA =" & Caja
Sql = Sql & " And COD_CLIENTE =" & Cliente
    
    Set rsRef = New ADODB.Recordset
    rsRef.Open Sql, ConActiva, 0, 1
    If Not rsRef.EOF Then
    
         rs!REFERENCIAS = Trim(Trim(rs!REFERENCIAS) & " Legajo")
    
    
    End If
    
    
  Sql = " SELECT     FK_CLIENTES, FK_CAJAS"
Sql = Sql & "  From DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  Where FK_CLIENTES =" & Cliente
Sql = Sql & " AND FK_CAJAS = " & Caja
Sql = Sql & " ORDER BY FK_CAJAS DESC"
    Set rsRef = New ADODB.Recordset
    rsRef.Open Sql, ConActiva, 0, 1
    If Not rsRef.EOF Then
         rs!REFERENCIAS = Trim(Trim(rs!REFERENCIAS) & " IMAGEN")
    End If
    
    
    Sql = "  SELECT     COD_CLIENTE, COD_NRO_CAJA"
    Sql = Sql & " From ORDENAR_DOCUMENTACION_DETALLE"
    Sql = Sql & " Where COD_CLIENTE = " & Cliente
    Sql = Sql & " And Cod_Nro_Caja = " & Caja
    
    Set rsRef = New ADODB.Recordset
    rsRef.Open Sql, ConActiva, 0, 1
    If Not rsRef.EOF Then
         rs!REFERENCIAS = Trim(Trim(rs!REFERENCIAS) & " DOCUMENTO")
    End If
    
    
    
' sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, NRO_REM_PROV , CLIENTEUSUARIO.APELLIDO_NOMBRE, REMITOS_CUERPO.FECHA"
' sql = sql & " FROM   REMITOS_CUERPO INNER JOIN"
' sql = sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO LEFT OUTER JOIN"
' sql = sql & " CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
' sql = sql & " WHERE (REMITOS_CUERPO.TIPO = 0) "
' sql = sql & " AND REMITOS_CUERPO.ID_CLIENTE = " & cliente
' sql = sql & " AND REMITOS_DETALLE.DESDE = " & Caja
' sql = sql & " AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0)"
'
'

Sql = " SELECT     ESTADO"
Sql = Sql & " From basasql.dbo.CONTENEDOR"
Sql = Sql & " Where"
Sql = Sql & " cod_cliente = " & Cliente
Sql = Sql & " And NRO_CAJA = " & Caja

     Set rsRef = New ADODB.Recordset
    rsRef.Open Sql, ConActiva, 0, 1
    If Not rsRef.EOF Then
    rs!estado = rsRef!estado
    Else
    rs!estado = 0
    End If
    
     
    
    
    rs.Update
    
    

    RsControl.MoveNext
Loop
End Sub


Public Sub BUSCAR_REMITOS()


Dim Sql As String
    Dim rs As ADODB.Recordset
    
    Dim PROVINCIA As String
    Dim PasoOrigen As String
    Dim PasoFin As String
    Dim NombreImagen As String
    
  directorio = InputBox("Ingrese el directorio")
    
   If Dir("c:\" & directorio, vbDirectory) = "" Then
    
  FileSystem.MkDir ("c:\" & directorio)
  End If
    
    
    
Sql = " SELECT     TEM_SUPERVIELLE.PROVINCIA, TEM_SUPERVIELLE.NRO_REMITO, TEM_SUPERVIELLE.SUBTIPO, REMITOS_CUERPO.NRO_REM_PROV,"
Sql = Sql & vbCrLf & " REMITOS_CUERPO.IMAGEN, TEM_SUPERVIELLE.FECHA, LEN(REMITOS_CUERPO.IMAGEN) AS LARGO"
Sql = Sql & vbCrLf & " FROM         TEM_SUPERVIELLE INNER JOIN"
Sql = Sql & vbCrLf & " REMITOS_CUERPO ON TEM_SUPERVIELLE.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO"
Sql = Sql & vbCrLf & " Where (TEM_SUPERVIELLE.NRO_REMITO > 0)"
Sql = Sql & vbCrLf & " ORDER BY TEM_SUPERVIELLE.PROVINCIA, TEM_SUPERVIELLE.SUBTIPO"
Set rs = New ADODB.Recordset



Set rs = New ADODB.Recordset
rs.Open Sql, ConActiva

Dim rsImagen As New ADODB.Recordset

Do While Not rs.EOF
    If PROVINCIA <> rs!PROVINCIA Then
        PROVINCIA = rs!PROVINCIA
        If Dir("c:\" & directorio & "\" & PROVINCIA, vbDirectory) = "" Then
        FileSystem.MkDir ("c:\" & directorio & "\" & PROVINCIA)
        End If
        PasoFin = "c:\" & directorio & "\" & PROVINCIA & "\"
         
    Else
    
    End If
    
        If rs!largo = 15 Then
            PasoOrigen = "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REMITO & "\"
         Else
            PasoOrigen = "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REM_PROV & "\"
         End If
         
         PasoOrigen = "\\222.15.19.251\Imagenes\"
         NombreImagen = rs!NRO_REMITO
         
         
         
          If Trim((rs!NRO_REM_PROV)) = "0001-000_____" Then
          
                NombreImagen = rs!NRO_REMITO
                Sql = " SELECT     DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_DESDE"
                Sql = Sql & vbCrLf & ",DOCUMENTOS_DIGITALES.DIRECTORIO_PASO , DOCUMENTOS_DIGITALES.NRO_DESDE"
                Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES INNER JOIN DOCUMENTOS_DIGITALES_LOTE ON"
                Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
                Sql = Sql & vbCrLf & " WHERE (DOCUMENTOS_DIGITALES.NRO_DESDE =" & rs!NRO_REMITO & ")"
                Sql = Sql & vbCrLf & " AND (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 83)"
          
          
          Else
                
                
                NombreImagen = rs!NRO_REM_PROV
                Sql = " SELECT     DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_DESDE"
                Sql = Sql & vbCrLf & ",DOCUMENTOS_DIGITALES.DIRECTORIO_PASO , DOCUMENTOS_DIGITALES.NRO_DESDE"
                Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES INNER JOIN DOCUMENTOS_DIGITALES_LOTE ON"
                Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
                Sql = Sql & vbCrLf & " WHERE (DOCUMENTOS_DIGITALES.LETRA_DESDE = '" & rs!NRO_REM_PROV & "')"
                Sql = Sql & vbCrLf & " AND (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 83)"
                Sql = Sql & vbCrLf & " AND (DOCUMENTOS_DIGITALES_LOTE.FECHA_PREPARACION > CONVERT(DATETIME, '2014-01-01 00:00:00', 102))"
         
          
          
          End If
         
         
         NombreImagen = NombreImagen & ".tif"
         Set rsImagen = New ADODB.Recordset
         rsImagen.Open Sql, strConBasa
         Dim RemitosFaltantes As String
     PasoFin = "c:\" & directorio & "\"
         If Not rsImagen.EOF Then
                If Dir(PasoImagenes & rsImagen!DIRECTORIO_PASO & "\" & rsImagen!ID & ".tif") <> "" Then
                  Rem   FileCopy PasoImagenes & rsImagen!DIRECTORIO_PASO & "\" & rsImagen!ID & ".tif", PasoFin & rs!SUBTIPO & " Remito " & NombreImagen
                FileCopy PasoImagenes & rsImagen!DIRECTORIO_PASO & "\" & rsImagen!ID & ".tif", PasoFin & PROVINCIA & "\" & rs!SUBTIPO & " Remito " & NombreImagen
                
                Else
                    RemitosFaltantes = RemitosFaltantes & vbCrLf & rs!Imagen
                End If
         Else
         
         
         
         
         
         If rs!largo = 15 Then
            PasoOrigen = "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REMITO & "\"
         Else
            PasoOrigen = "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REM_PROV & "\"
         End If
         
         
         RemitosFaltantes = RemitosFaltantes & vbCrLf & NombreImagen
         
         End If
         
    

    rs.MoveNext
Loop

MsgBox "Remitos Faltantes " & RemitosFaltantes

MsgBox "Terminado"

End Sub

Public Sub BUSCAR_REMITOS_NUEVO()


Dim Sql As String
    Dim rs As ADODB.Recordset
    
    Dim PROVINCIA As String
    Dim PasoOrigen As String
    Dim PasoFin As String
    Dim NombreImagen As String
    
  directorio = InputBox("Ingrese el directorio")
    
   If Dir("c:\" & directorio, vbDirectory) = "" Then
    
  FileSystem.MkDir ("c:\" & directorio)
  End If
    
    
    
Sql = " SELECT     TEM_SUPERVIELLE.PROVINCIA, TEM_SUPERVIELLE.NRO_REMITO, TEM_SUPERVIELLE.SUBTIPO, REMITOS_CUERPO.NRO_REM_PROV,"
Sql = Sql & vbCrLf & " REMITOS_CUERPO.IMAGEN, TEM_SUPERVIELLE.FECHA, LEN(REMITOS_CUERPO.IMAGEN) AS LARGO , TEM_SUPERVIELLE.PASO_IMAGEN "
Sql = Sql & vbCrLf & " FROM         TEM_SUPERVIELLE INNER JOIN"
Sql = Sql & vbCrLf & " REMITOS_CUERPO ON TEM_SUPERVIELLE.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO"
Sql = Sql & vbCrLf & " Where (TEM_SUPERVIELLE.NRO_REMITO > 0)"
Sql = Sql & vbCrLf & " ORDER BY TEM_SUPERVIELLE.PROVINCIA, TEM_SUPERVIELLE.SUBTIPO"
Set rs = New ADODB.Recordset



Set rs = New ADODB.Recordset
rs.Open Sql, ConActiva

Dim rsImagen As New ADODB.Recordset

Do While Not rs.EOF
    If PROVINCIA <> rs!PROVINCIA Then
        PROVINCIA = rs!PROVINCIA
        If Dir("c:\" & directorio & "\" & PROVINCIA, vbDirectory) = "" Then
        FileSystem.MkDir ("c:\" & directorio & "\" & PROVINCIA)
        End If
        PasoFin = "c:\" & directorio & "\" & PROVINCIA & "\"
         
    Else
    
    End If
    
         
        If "0001-000_____" <> Trim(rs!NRO_REM_PROV) Then
             NombreImagen = rs!NRO_REM_PROV
        Else
            NombreImagen = rs!NRO_REMITO
        End If
        
         
         
        
         Dim RemitosFaltantes As String
         PasoFin = "c:\" & directorio & "\"
         If Not IsNull(rs!PASO_IMAGEN) Then
         
         If Dir(rs!PASO_IMAGEN) <> "" Then
                If PROVINCIA = "" Then
                  FileCopy rs!PASO_IMAGEN, PasoFin & rs!SUBTIPO & " Remito " & NombreImagen & ".TIF"
                Else
                  FileCopy rs!PASO_IMAGEN, PasoFin & PROVINCIA & "\" & rs!SUBTIPO & " Remito " & NombreImagen & ".TIF"
                End If
          Else
                If Dir(rs!PASO_IMAGEN & ".Tif") <> "" Then
                    If PROVINCIA = "" Then
                      FileCopy rs!PASO_IMAGEN & ".Tif", PasoFin & rs!SUBTIPO & " Remito " & NombreImagen & ".TIF"
                    Else
                      FileCopy rs!PASO_IMAGEN, PasoFin & PROVINCIA & "\" & rs!SUBTIPO & " Remito " & NombreImagen & ".TIF"
                    End If
                Else
                    Debug.Print rs!PASO_IMAGEN
                    MsgBox "error" & rs!PASO_IMAGEN
                End If
          End If
          
         Else
          RemitosFaltantes = RemitosFaltantes & vbCrLf & NombreImagen
         End If
         
         
    

    rs.MoveNext
Loop

MsgBox "Remitos Faltantes " & RemitosFaltantes

MsgBox "Terminado"

End Sub

Public Sub BUSCAR_Requerimientos()


Dim Sql As String
    Dim rs As ADODB.Recordset

    Dim PROVINCIA As String
    Dim PasoOrigen As String
    Dim PasoFin As String
    Dim NombreImagen As String
    Dim ErrorRequerimiento As String
    
 
    
 Dim ReqSinImangen As String
    
    
Sql = " SELECT    *"
Sql = Sql & vbCrLf & " FROM         TEM_SUPERVIELLE "
Sql = Sql & vbCrLf & "  Where (REQUERIMIENTO > 0) And (NRO_REMITO = 0)"

Sql = Sql & vbCrLf & " ORDER BY TEM_SUPERVIELLE.PROVINCIA, TEM_SUPERVIELLE.SUBTIPO"
Set rs = New ADODB.Recordset



Set rs = New ADODB.Recordset
rs.Open Sql, ConActiva
Do While Not rs.EOF
    If PROVINCIA <> rs!PROVINCIA Then
        PROVINCIA = rs!PROVINCIA
        If Dir("c:\" & directorio & "\" & PROVINCIA, vbDirectory) = "" Then
             FileSystem.MkDir ("c:\" & directorio & "\" & PROVINCIA)
        End If
        
        PasoFin = "c:\" & directorio & "\" & PROVINCIA & "\"
         
    Else
    
    End If
    

    Sql = " SELECT *  "
    Sql = Sql & vbCrLf & " FROM   V_REQUERIMIENTO_GENERICO"
    Sql = Sql & vbCrLf & " Where IDREQUERIMIENTO = " & rs!REQUERIMIENTO
    Sql = Sql & vbCrLf & " ORDER BY IDREQUERIMIENTO "
    If Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & rs!REQUERIMIENTO, vbDirectory) = "" Then
             FileSystem.MkDir ("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & rs!REQUERIMIENTO)
        End If
    
    frmReportes.Exportarpdf PasoReportes & "RequerimientoGenerico.rpt", Sql, "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & rs!REQUERIMIENTO & "\" & rs!REQUERIMIENTO & ".pdf"
    
    Rem frmReportes.Exportarpdf PasoReportes + "FacturacionSupervielle.rpt", sql, True
    
    NombreImagen = Dir("\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & rs!REQUERIMIENTO & "\*.*")
    
If NombreImagen = "" Then
    ReqSinImangen = ReqSinImangen & vbCrLf & "El requerimiento : " & rs!REQUERIMIENTO & " NO tiene Imagene"


Else
    PasoOrigen = "\\222.15.19.251\basa\Administracion\Imagenes_Internas\Requerimientos\" & rs!REQUERIMIENTO & "\" & NombreImagen
    FileCopy PasoOrigen, PasoFin & rs!SUBTIPO & " Requerimiento " & NombreImagen
End If


         
        
         
             
             

    rs.MoveNext
Loop

If ReqSinImangen <> "" Then
         Clipboard.Clear
         Clipboard.SetText ReqSinImangen
         MsgBox "Requerimiento sin imagen copiados a memoria" & ReqSinImangen, vbInformation
End If

 
MsgBox "Terminado"

End Sub


Public Sub IngresoDisco(FORMA As String, NRO_REMITO As Long, NRO_REM_PROV As String, TIPO As String, fecha As String, APELLIDO_NOMBRE As String, Sucursal As String, FK_CLIENTE As Integer, CANTIDADCAJAS As Double)

Dim Sql As String

Sql = " INSERT INTO basasql.dbo.TEM_DISCO ( "
Sql = Sql & "  FORMA "
Sql = Sql & " , NRO_REMITO"
Sql = Sql & " ,  NRO_REM_PROV"
Sql = Sql & " ,  TIPO"
Sql = Sql & " ,  FECHA"
Sql = Sql & " , APELLIDO_NOMBRE"
Sql = Sql & " , SUCURSAL"
Sql = Sql & " , FK_CLIENTE"
Sql = Sql & " , CANTIDADCAJAS)"
Sql = Sql & " VALUES ("
Sql = Sql & FORMA
Sql = Sql & " , " & NRO_REMITO
Sql = Sql & " , " & NRO_REM_PROV
Sql = Sql & " , " & TIPO
Sql = Sql & " , " & fecha
Sql = Sql & " , " & APELLIDO_NOMBRE
Sql = Sql & " , " & Sucursal
Sql = Sql & " , " & FK_CLIENTE
Sql = Sql & " ," & CANTIDADCAJAS & ")"

ExecutarSql Sql



End Sub

Public Sub InsertarDisco(Sql As String)
    Dim rs As ADODB.Recordset
    Dim cantidad As Double
    
    
    Rem FORMA As String, NRO_REMITO As Long, NRO_REM_PROV As String, TIPO As String, FECHA As String, APELLIDO_NOMBRE As String, SUCURSAL As String, FK_CLIENTE As Integer
    

    Set rs = New ADODB.Recordset
    rs.Open Sql, strConBasa
   Do While Not rs.EOF
     If rs!TIPOVALOR = 0 Then
        cantidad = rs!cantidad
     Else
         cantidad = rs!cantidad * -1
     End If
     
        IngresoDisco "'INGRESO Y EGRESO DE CAJAS'", rs!NRO_REMITO, "'" & rs!NRO_REM_PROV & "'", "'" & rs!TIPO & "'", FechaFormato(Trim(rs!fecha)) _
         , "'" & rs!APELLIDO_NOMBRE & "'", "'" & rs!Sucursal & "'", rs!FK_CLIENTE, cantidad
        rs.MoveNext
    Loop
    
    



    

End Sub


Public Sub IVADATA()

Dim conData As New ADODB.Connection
Dim RsFactura As New ADODB.Recordset
Dim rsTEM_IVA_DATA As New ADODB.Recordset

Dim Sql As String

    
conData.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=datas"

'Sql = "SELECT FacturaABC, NumeroFactura, FechaFacturacion, NombreCliente, "
'Sql = Sql & vbCrLf & "   CUIT,  Subtotal, TotalFacturado, MesFacturacion, AnoFacturacion, "
'Sql = Sql & vbCrLf & "    IVAInscripto , IvaNoInscripto, Impresa, IDCliente "
'Sql = Sql & vbCrLf & " From FACTURA "
'Sql = Sql & vbCrLf & " ORDER BY FechaFacturacion DESC "
'
'    Sql = "SELECT FacturaABC, NumeroFactura, FechaFacturacion, NombreCliente, "
'    Sql = Sql & vbCrLf & "   CUIT,  Subtotal, TotalFacturado , IDCliente"
'    Sql = Sql & vbCrLf & " From FACTURA "
'    Sql = Sql & vbCrLf & " where "
'SELECT CLIENTE.[IDCLIENTE], CLIENTE.[NOMBRE], FACTURA.FechaFacturacion, CLIENTE.CUIT
'FROM CLIENTE INNER JOIN FACTURA ON CLIENTE.IDCLIENTE = FACTURA.IDCliente;

Sql = "SELECT FacturaABC, NumeroFactura, FechaFacturacion, NombreCliente, "
    Sql = Sql & vbCrLf & "   CLIENTE.CUIT,  Subtotal, TotalFacturado ,CLIENTE.IDCliente"
    Sql = Sql & vbCrLf & " FROM FACTURA, CLIENTE  "
    Sql = Sql & vbCrLf & " where CLIENTE.IDCLIENTE = FACTURA.IDCliente AND "

    
  Sql = Sql & vbCrLf & " FechaFacturacion >= " & FECHADATA_Dias(txtFechaDesde.Text)
 Sql = Sql & vbCrLf & "  And FechaFacturacion <= " & FECHADATA_Dias(txtFechaHasta.Text)
 Rem Sql = Sql & vbCrLf & "   NumeroFactura BETWEEN 2 AND 20 "

Set RsFactura = New ADODB.Recordset

RsFactura.Open Sql, conData

 ExecutarSql ("DELETE FROM basasql.dbo.TEM_IVA_DATA")

Sql = " SELECT * "
 Sql = Sql & vbCrLf & " From [basasql].[dbo].[TEM_IVA_DATA]"


rsTEM_IVA_DATA.Open Sql, strConBasa, 2, 3
Dim lETRA As String
  

Do While Not RsFactura.EOF
rsTEM_IVA_DATA.AddNew
Select Case RsFactura!FacturaABC
Case "G"
    lETRA = "B"
Case "F"
    lETRA = "A"
Case Else
    lETRA = RsFactura!FacturaABC
End Select

If RsFactura!NumeroFactura < 50 Then
    MsgBox "DDDD"
    lETRA = "E"
End If


rsTEM_IVA_DATA!FacturaABC = RsFactura!FacturaABC
rsTEM_IVA_DATA!NumeroFactura = RsFactura!NumeroFactura
rsTEM_IVA_DATA!FACTURA = "_" & " 0001-" & Format(RsFactura!NumeroFactura, "00000000")
 rsTEM_IVA_DATA!FechaFacturacion = CStr(FECHADATA_Fecha(RsFactura!FechaFacturacion))
rsTEM_IVA_DATA!NombreCliente = RsFactura!NombreCliente
rsTEM_IVA_DATA!Cuit = RsFactura!Cuit
rsTEM_IVA_DATA!Subtotal = RsFactura!Subtotal
rsTEM_IVA_DATA!TotalFacturado = RsFactura!TotalFacturado
rsTEM_IVA_DATA!IVA = rsTEM_IVA_DATA!TotalFacturado - rsTEM_IVA_DATA!Subtotal
rsTEM_IVA_DATA!lETRA = lETRA


rsTEM_IVA_DATA.Update
    RsFactura.MoveNext
 Loop

'    Set RsFactura = New ADODB.Recordset
'
'    rsTEM_IVA_DATA.Open Sql, strConBasa, 2, 3
'
'    Sql = "SELECT FacturaABC, NumeroFactura, FechaFacturacion, NombreCliente, "
'    Sql = Sql & vbCrLf & "   CUIT,  Subtotal, TotalFacturado"
'    Sql = Sql & vbCrLf & " From FACTURA "
'    Sql = Sql & vbCrLf & " where FechaFacturacion >= " & FECHADATA_Dias(txtFechaDesde.Text)
'    Sql = Sql & vbCrLf & "  And FechaFacturacion <= " & FECHADATA_Dias(txtFechaHasta.Text)




End Sub


Public Sub ControlCajas5000SInMovimiento(Cliente As Integer)

ExecutarSql "DELETE FROM TEM_CONTROL_CAJAS_5000"
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim RS_UNITARIO As ADODB.Recordset
Sql = "SELECT     CAMBIOPOSICION.ESTANTERIA, CAMBIOPOSICION.HORIZONTAL, CAMBIOPOSICION.VERTICAL, CAMBIOPOSICION.ADELANTE_ATRAS,"
Sql = Sql & " CAMBIOPOSICION.NRO_ESTANTE, CAMBIOPOSICION.ESTADO, CAMBIOPOSICION.COD_CLIENTE, CAMBIOPOSICION.NRO_CAJA, CAMBIOPOSICION.FECHA,"
Sql = Sql & " CAMBIOPOSICION.ID_PERSONAL, V_CONTROL_HISTORICOS_5000.COD_CLIENTE AS CLIENTE , V_CONTROL_HISTORICOS_5000.CANT"
Sql = Sql & "  FROM         CAMBIOPOSICION INNER JOIN"
Sql = Sql & " V_CONTROL_HISTORICOS_5000 ON CAMBIOPOSICION.COD_CLIENTE = V_CONTROL_HISTORICOS_5000.COD_CLIENTE AND"
Sql = Sql & " CAMBIOPOSICION.NRO_CAJA = V_CONTROL_HISTORICOS_5000.NRO_CAJA"
Sql = Sql & "  Where (V_CONTROL_HISTORICOS_5000.COD_CLIENTE = " & Cliente & ") And (CAMBIOPOSICION.Estanteria > 5000)"
Sql = Sql & "  ORDER BY CAMBIOPOSICION.COD_CLIENTE, CAMBIOPOSICION.NRO_CAJA"

rs.Open Sql, strConBasa

Do While Not rs.EOF
    Set RS_UNITARIO = New ADODB.Recordset
    
    Sql = "SELECT     TOP (1) NRO_CAJA, ID_CLIENTE"
Sql = Sql & " From basasql.dbo.MOV_CAJAS2"
Sql = Sql & " Where id_cliente = " & rs!Cliente
Sql = Sql & " And Elemento = " & rs!NRO_CAJA
Sql = Sql & " And (TIPO_ELEMENTO = 0) "
Sql = Sql & " And (TIPO = 1)"
Sql = Sql & " ORDER BY FECHA_MOVIMIENTO"
    RS_UNITARIO.Open Sql, strConBasa
    If RS_UNITARIO.EOF Then
        Sql = " INSERT INTO TEM_CONTROL_CAJAS_5000"
        Sql = Sql & " (NRO_CAJA, COD_CLIENTE)"
        Sql = Sql & " VALUES     (" & rs!NRO_CAJA & "," & rs!Cliente & " )"
        ExecutarSql Sql
        Debug.Print rs!NRO_CAJA
    End If
    
    
    
    rs.MoveNext
Loop


End Sub




Public Sub EXPURGO_DISCO(Cliente As Integer, Nro_documento As Long, FechaHasta As String)
    Dim rs As New ADODB.Recordset
    Dim RsControl As ADODB.Recordset
    Dim Sql As String
    Dim Indice As String
    Dim FechaProceso As String
    Indice = BuscarIndice(Cliente, Nro_documento)
    
    

        Sql = " SELECT REFERENCIAS.COD_CLIENTE, REFERENCIAS.NRO_CAJA , INDICES.ID_CODIGO_DOCUMENTO "
        Sql = Sql & " FROM REFERENCIAS INNER JOIN "
        Sql = Sql & " INDICES ON REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE"
        Sql = Sql & " AND REFERENCIAS.INDICE = INDICES.INDICE "
        Sql = Sql & " WHERE (INDICES.INDICE LIKE '" & Indice & "%')"
        Sql = Sql & " AND (REFERENCIAS.FECHA_HASTA <= '" & FechaHasta & "')"
        Sql = Sql & " GROUP BY REFERENCIAS.COD_CLIENTE, REFERENCIAS.NRO_CAJA ,  INDICES.ID_CODIGO_DOCUMENTO"
        Sql = Sql & " Having REFERENCIAS.COD_CLIENTE = " & Cliente
        rs.Open Sql, strConBasa
        FechaProceso = SysDate
        
         ExecutarSql "DELETE FROM EXPURGO_DISCO"
        Do While Not rs.EOF
            Set RsControl = New ADODB.Recordset
            Sql = " SELECT COD_CLIENTE, NRO_CAJA, FECHA_HASTA "
            Sql = Sql & " From REFERENCIAS"
            Sql = Sql & " WHERE COD_CLIENTE =  " & rs!COD_CLIENTE
            Sql = Sql & " AND NRO_CAJA = " & rs!NRO_CAJA
            Sql = Sql & " AND FECHA_HASTA > '" & FechaHasta & "'"
            RsControl.Open Sql, strConBasa
            If RsControl.EOF Then
                Sql = "Insert Into EXPURGO_DISCO("
                Sql = Sql & "  NRO_CAJA"
                Sql = Sql & ", COD_CLIENTE"
                Sql = Sql & ", FECHA"
                Sql = Sql & ", ID_CODIGO_DOCUMENTO)"
                Sql = Sql & " VALUES ("
                Sql = Sql & rs!NRO_CAJA
                Sql = Sql & "," & rs!COD_CLIENTE
                Sql = Sql & "," & FechaProceso
                Sql = Sql & "," & rs!ID_CODIGO_DOCUMENTO
                Sql = Sql & ")"
                ExecutarSql Sql
            End If
            
            rs.MoveNext
        Loop
        
        Set rs = New ADODB.Recordset
       
        
Sql = " SELECT     EXPURGO_DISCO.NRO_CAJA, EXPURGO_DISCO.COD_CLIENTE, INDICES.ID_CODIGO_DOCUMENTO"
Sql = Sql & "  FROM         EXPURGO_DISCO INNER JOIN"
Sql = Sql & " REFERENCIAS ON EXPURGO_DISCO.NRO_CAJA = REFERENCIAS.NRO_CAJA AND EXPURGO_DISCO.COD_CLIENTE = REFERENCIAS.COD_CLIENTE INNER JOIN"
Sql = Sql & " INDICES ON REFERENCIAS.INDICE = INDICES.INDICE AND REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE"
Sql = Sql & "  where INDICES.ID_CODIGO_DOCUMENTO BETWEEN  80000 AND 89999"
Sql = Sql & "  ORDER BY EXPURGO_DISCO.NRO_CAJA"
        
        
        rs.Open Sql, strConBasa
        
        Do While Not rs.EOF
            Sql = " DELETE  "
            Sql = Sql & " From basasql.dbo.EXPURGO_DISCO"
            Sql = Sql & " Where NRO_CAJA = " & rs!NRO_CAJA
            Sql = Sql & " And COD_CLIENTE = " & rs!COD_CLIENTE
            ExecutarSql Sql
            rs.MoveNext
        Loop
        
MsgBox " terminados"
End Sub

