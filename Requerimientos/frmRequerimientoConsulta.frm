VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRequerimientoConsulta 
   Caption         =   "REQUERIMIENTO"
   ClientHeight    =   10650
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   13065
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
   MDIChild        =   -1  'True
   ScaleHeight     =   10650
   ScaleWidth      =   13065
   Begin VB.CommandButton cmdInsertOrdenBusqueda 
      Caption         =   "Orden Insert"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11280
      TabIndex        =   56
      Top             =   5700
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Excel"
      Height          =   315
      Left            =   8520
      TabIndex        =   55
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdMarcarComoVerificado 
      Caption         =   "Verificados"
      Height          =   315
      Left            =   9840
      TabIndex        =   54
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COPIAR"
      Height          =   375
      Left            =   11640
      TabIndex        =   53
      Top             =   7920
      Width           =   1215
   End
   Begin VB.ComboBox cboHoraDia 
      Height          =   345
      ItemData        =   "frmRequerimientoConsulta.frx":0000
      Left            =   4260
      List            =   "frmRequerimientoConsulta.frx":000A
      TabIndex        =   51
      Top             =   540
      Width           =   2595
   End
   Begin VB.CommandButton cmdCrearNuevo 
      Caption         =   "Crear Nuevo"
      Height          =   315
      Left            =   11160
      TabIndex        =   50
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CheckBox chkEstado 
      Alignment       =   1  'Right Justify
      Caption         =   "Estado"
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
      Left            =   7920
      TabIndex        =   49
      Top             =   2700
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.TextBox txtEstado 
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
      Left            =   9060
      TabIndex        =   48
      Top             =   2640
      Width           =   915
   End
   Begin VB.TextBox txtCantidad 
      Height          =   330
      Left            =   1320
      TabIndex        =   47
      Text            =   "0"
      Top             =   2220
      Width           =   555
   End
   Begin VB.TextBox txtImagenes 
      Height          =   315
      Left            =   3060
      TabIndex        =   45
      Text            =   "0"
      Top             =   2220
      Width           =   555
   End
   Begin VB.ComboBox cboSucursal 
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
      Left            =   6000
      TabIndex        =   43
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox chkFlete 
      Alignment       =   1  'Right Justify
      Caption         =   "Flete"
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
      Left            =   420
      TabIndex        =   42
      Top             =   2700
      Width           =   795
   End
   Begin VB.CheckBox chkCobrar 
      Alignment       =   1  'Right Justify
      Caption         =   "Cobrar Cajas"
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
      Left            =   3240
      TabIndex        =   41
      Top             =   2700
      Width           =   1455
   End
   Begin VB.TextBox txtHorasArchivista 
      Height          =   330
      Left            =   2700
      TabIndex        =   40
      Text            =   "0"
      Top             =   2640
      Width           =   435
   End
   Begin VB.CommandButton cmdfechaCompromiso 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10080
      TabIndex        =   37
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtFechaCompromiso 
      BackColor       =   &H00FFC0FF&
      Height          =   315
      Left            =   9060
      TabIndex        =   36
      Top             =   540
      Width           =   2415
   End
   Begin VB.CommandButton cmdLeida 
      Caption         =   "Leída"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   35
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdOrdenBusqueda 
      Caption         =   "O. Busq."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7800
      TabIndex        =   34
      Top             =   5700
      Width           =   1095
   End
   Begin VB.CheckBox chkEnvioPorCorreo 
      Caption         =   "Envio por correo"
      Height          =   255
      Left            =   9540
      TabIndex        =   33
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdBorrarLegajos 
      Caption         =   "Borrar"
      Height          =   315
      Left            =   7200
      TabIndex        =   32
      Top             =   2220
      Width           =   855
   End
   Begin VB.TextBox txtCargarLegajos 
      Height          =   315
      Left            =   5340
      TabIndex        =   31
      Top             =   2220
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdelante 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      TabIndex        =   29
      Top             =   5700
      Width           =   435
   End
   Begin VB.CommandButton cmdAtras 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5580
      TabIndex        =   28
      Top             =   5700
      Width           =   435
   End
   Begin VB.CommandButton cmdActual 
      Caption         =   "Actual"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9000
      TabIndex        =   27
      Top             =   5700
      Width           =   975
   End
   Begin VB.CommandButton cmdHistorico 
      Caption         =   "Historico"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4260
      TabIndex        =   26
      Top             =   5700
      Width           =   1335
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10140
      TabIndex        =   18
      Top             =   5700
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   1335
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   6180
      Width           =   12555
   End
   Begin MSDataGridLib.DataGrid grdDetalle 
      Height          =   2055
      Left            =   180
      TabIndex        =   15
      Top             =   3540
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
         Weight          =   700
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
   Begin VB.Label lblDescripcion 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label15"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   52
      Top             =   7740
      Width           =   11295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label13 
      Caption         =   "Sucursal:"
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
      Left            =   4740
      TabIndex        =   46
      Top             =   2700
      Width           =   915
   End
   Begin VB.Label Label11 
      Caption         =   "Imagenes:"
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
      Left            =   2100
      TabIndex        =   44
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label Label20 
      Caption         =   "H. Archivistas"
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
      Index           =   2
      Left            =   1440
      TabIndex        =   39
      Top             =   2700
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Hora"
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
      Left            =   3540
      TabIndex        =   38
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Agregar/Borrar"
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
      Left            =   3960
      TabIndex        =   30
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Fecha Carga:"
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
      TabIndex        =   25
      Top             =   5700
      Width           =   1155
   End
   Begin VB.Label lblFechaModificacion 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1440
      TabIndex        =   24
      Top             =   5700
      Width           =   2655
   End
   Begin VB.Label Label5 
      Caption         =   "Cantidad:"
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
      Left            =   360
      TabIndex        =   23
      Top             =   2280
      Width           =   915
   End
   Begin VB.Label lblSector 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6180
      TabIndex        =   22
      Top             =   1800
      Width           =   3915
   End
   Begin VB.Label Label20 
      Caption         =   "Suc./ Sector"
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
      Index           =   0
      Left            =   4740
      TabIndex        =   21
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Label lblSolicito 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1800
      TabIndex        =   20
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label18 
      Caption         =   "Solicito:"
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
      Left            =   420
      TabIndex        =   19
      Top             =   1860
      Width           =   1275
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Left            =   360
      TabIndex        =   16
      Top             =   3120
      Width           =   6915
   End
   Begin VB.Label lblPersonalAsignado 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8640
      TabIndex        =   14
      Top             =   1380
      Width           =   2895
   End
   Begin VB.Label Label14 
      Caption         =   "P.  Asignado:"
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
      Left            =   7320
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblPersonalCarga 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4800
      TabIndex        =   12
      Top             =   1380
      Width           =   2415
   End
   Begin VB.Label Label12 
      Caption         =   "P. Carga:"
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
      Left            =   3900
      TabIndex        =   11
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label lblEstado 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5880
      TabIndex        =   10
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "Estado:"
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
      Left            =   4980
      TabIndex        =   9
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label lblTipoRequerimiento 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1860
      TabIndex        =   8
      Top             =   960
      Width           =   2835
   End
   Begin VB.Label Label8 
      Caption         =   "Tipo:"
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
      Left            =   420
      TabIndex        =   7
      Top             =   1020
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Fecha Compromiso:"
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
      Left            =   7320
      TabIndex        =   6
      Top             =   600
      Width           =   1755
   End
   Begin VB.Label lblFechaCarga 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1860
      TabIndex        =   5
      Top             =   1380
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha Carga:"
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
      Left            =   420
      TabIndex        =   4
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label lblRequerimiento 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   3
      Top             =   540
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "Requerimiento:"
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
      Left            =   420
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label lblCliente 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   9675
   End
   Begin VB.Label Label1 
      Caption         =   "Cliente:"
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
      Left            =   420
      TabIndex        =   0
      Top             =   180
      Width           =   855
   End
   Begin VB.Menu mnuGrillaDetalle 
      Caption         =   "GrillaDetalle"
      Begin VB.Menu mnuIngresarLegajo 
         Caption         =   "Ingresar Legajos"
      End
      Begin VB.Menu mnuBorrarElementos 
         Caption         =   "Borrar Elementos "
      End
   End
End
Attribute VB_Name = "frmRequerimientoConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDescHistorico As ADODB.Recordset
 Dim RsDetalle As New ADODB.Recordset
Private Sub ReporteBusqueda(Lote_Busqueda As Integer)
  Dim sSQL As String
    
    
        sSQL = " SELECT  ID_LOTE_BUSQUEDA, FECHA, TIPO, COD_CLIENTE, NRO_CAJA, LEGAJO, DESCRIPCION, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
        sSQL = sSQL & " UB_PROVISORIA,DETALLE "
        sSQL = sSQL & " FROM  V_TEM_BUSQUEDA "
        sSQL = sSQL & " where ID_LOTE_BUSQUEDA = " & Lote_Busqueda
      
       frmReportes.ImprimirReporte PasoReportes + "BusquedaGeneral.rpt", sSQL, True


End Sub


Public Sub CargarRequerimiento(IDREQUERIMIENTO As String)
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
   
chkFlete.Enabled = True
   Rem hacer vista
   

    
   sql = "  SELECT     dbo.REQUERIMIENTO.IDREQUERIMIENTO, dbo.REQUERIMIENTO.ID_CLIENTE, dbo.REQUERIMIENTO.IDPERSONAL, dbo.REQUERIMIENTO.IDESTADO,"
    sql = sql & vbCrLf & "  dbo.REQUERIMIENTO.IDTIPOREQUERIMIENTO, dbo.CLIENTES.RAZON_SOCIAL, dbo.TIPOREQUERIMIENTO.DESCRIPCION AS DESC_TIPO,"
                       sql = sql & vbCrLf & "  dbo.REQUERIMIENTO.COD_USUARIO_CLIENTE, dbo.CLIENTEUSUARIO.APELLIDO_NOMBRE, dbo.INDICES.TITULOHERENCIA,"
                       sql = sql & vbCrLf & "  dbo.REQUERIMIENTO_ESTADO.DESCRIPCION AS DESC_ESTADO, dbo.REQUERIMIENTO.CANTIDAD, dbo.PERSONAL.APELLIDO,"
                       sql = sql & vbCrLf & "  dbo.PERSONAL.NOMBRE, dbo.REQUERIMIENTO.DESCRIPCION, dbo.REQUERIMIENTO.FECHAENTREGA,"
                       sql = sql & vbCrLf & "  dbo.Requerimiento.FECHARECEPCION , dbo.Requerimiento.CANTIDAD, dbo.REQUERIMIENTO.IDTIPOREQUERIMIENTO , IDESTADO ,  ENVIOPORCORREO , COMPROMISO_ENTREGA, FECHA_SISTEMA , HORA_ARCHIVISTA , FLETE,COBRAR ,   CANTIDAD_IMAGENES , dbo.REQUERIMIENTO.FK_SUCURSAL     "
 sql = sql & vbCrLf & "  FROM         dbo.REQUERIMIENTO_ESTADO INNER JOIN"
                       sql = sql & vbCrLf & " dbo.REQUERIMIENTO INNER JOIN"
                       sql = sql & vbCrLf & "  dbo.CLIENTES ON dbo.REQUERIMIENTO.ID_CLIENTE = dbo.CLIENTES.ID_CLIENTE INNER JOIN"
                       sql = sql & vbCrLf & "  dbo.TIPOREQUERIMIENTO ON dbo.REQUERIMIENTO.IDTIPOREQUERIMIENTO = dbo.TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO ON"
                       sql = sql & vbCrLf & "  dbo.REQUERIMIENTO_ESTADO.ID_ESTADO = dbo.REQUERIMIENTO.IDESTADO INNER JOIN"
                        sql = sql & vbCrLf & "  dbo.PERSONAL ON dbo.REQUERIMIENTO.IDPERSONAL = dbo.PERSONAL.IDPERSONAL LEFT OUTER JOIN"
                        sql = sql & vbCrLf & "  dbo.INDICES INNER JOIN"
                       sql = sql & vbCrLf & "  dbo.CLIENTEUSUARIO ON dbo.INDICES.INDICE = dbo.CLIENTEUSUARIO.COD_INDICE AND"
                       sql = sql & vbCrLf & "  dbo.INDICES.COD_CLIENTE = dbo.CLIENTEUSUARIO.COD_CLIENTE ON"
                       sql = sql & vbCrLf & " dbo.Requerimiento.COD_USUARIO_CLIENTE = dbo.CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
 sql = sql & vbCrLf & " WHERE dbo.REQUERIMIENTO.IDREQUERIMIENTO = " & IDREQUERIMIENTO
 
 
 sql = " SELECT * FROM V_REQUERIMIENTO_DETALLE WHERE IDREQUERIMIENTO = " & IDREQUERIMIENTO

 
 
    rs.CursorLocation = adUseClient
    
    
    rs.Open sql, ConActiva, 0, 1
    
    If Not rs.EOF Then
    
    If rs!IDESTADO > 3 Then
        txtDescripcion.Enabled = False
    Else
         txtDescripcion.Enabled = True
    End If
    
    If rs!IDESTADO > 4 Then
       chkFlete.Enabled = False
    End If
    
    If rs!IDTIPOREQUERIMIENTO = 6 Then
        txtCantidad.Enabled = True
    Else
        txtCantidad.Enabled = False
    End If
    
    If (rs!IDTIPOREQUERIMIENTO = 1 Or rs!IDTIPOREQUERIMIENTO = 3 Or rs!IDTIPOREQUERIMIENTO = 10 Or rs!IDTIPOREQUERIMIENTO = 11 Or rs!IDTIPOREQUERIMIENTO = 2 Or rs!IDTIPOREQUERIMIENTO = 4 Or rs!IDTIPOREQUERIMIENTO = 9) And rs!IDESTADO < 5 Then
        txtCargarLegajos.Enabled = True
    Else
        txtCargarLegajos.Enabled = False
    End If
        lblCliente.Caption = Format(rs!ID_CLIENTE, "0000") & "  " & Trim(rs!Razon_Social)
        lblRequerimiento.Caption = rs!IDREQUERIMIENTO
        If IsNull(rs!FECHA_SISTEMA) Then
        lblFechaCarga.Caption = ""
        Else
            lblFechaCarga.Caption = rs!FECHA_SISTEMA
        End If
        txtFechaCompromiso.Text = rs!FECHAENTREGA
        lblTipoRequerimiento.Caption = Format(rs!IDTIPOREQUERIMIENTO, "0000") & " " & rs!DESC_TIPO
        lblEstado.Caption = Format(rs!IDESTADO, "0000") & " " & rs!DESC_ESTADO
        lblPersonalAsignado.Caption = Trim(rs!Apellido) & " " & Trim(rs!Nombre)
        txtCantidad.Text = rs!CANTIDAD
        cboHoraDia.Text = Trim(rs!COMPROMISO_ENTREGA)
        
        
              
        txtHorasArchivista.Text = IIf(IsNull(Trim(rs!HORA_ARCHIVISTA)), "", Trim(rs!HORA_ARCHIVISTA))
        If IsNull(rs!Flete) Then
            chkFlete.Value = 0
        Else
            If Trim(rs!Flete) = "" Then
                chkFlete.Value = 0
            Else
                chkFlete.Value = rs!Flete
            End If
            
        End If
        
        
       
        
        If Trim(rs!Cobrar) = "" Then
            chkCobrar.Value = 0
        Else
            chkCobrar.Value = IIf(IsNull(Trim(rs!Cobrar)), "0", Trim(rs!Cobrar))
        End If
        
        If Not IsNull(rs!ENVIOPORCORREO) Then
            If rs!ENVIOPORCORREO = True Then
                chkEnvioPorCorreo.Value = 1
            Else
                chkEnvioPorCorreo.Value = 0
            End If
         Else
            chkEnvioPorCorreo.Value = 0
        End If
        
        If Not IsNull(rs!APELLIDO_NOMBRE) Then
            lblSolicito.Caption = rs!APELLIDO_NOMBRE
           If Not IsNull(rs!TITULOHERENCIA) Then
                    If Len(rs!TITULOHERENCIA) > 100 Then
                          lblsector.Caption = Mid(rs!TITULOHERENCIA, 50)
                    Else
                        lblsector.Caption = rs!TITULOHERENCIA
                        End If
                Else
                lblsector.Caption = ""
                
            End If
        Else
            lblSolicito.Caption = ""
            lblsector.Caption = ""
        End If
        txtDescripcion.Text = ""
        If Not IsNull(rs!DESCRIPCION) Then
           lblDescripcion.Caption = Trim(rs!DESCRIPCION)
        Else
           txtDescripcion.Text = ""
           lblDescripcion.Caption = ""
        End If
        
        If Not IsNull(rs!CANTIDAD_IMAGENES) Then
            txtImagenes.Text = rs!CANTIDAD_IMAGENES
        Else
            txtImagenes.Text = 0
        End If
        
        cboSucursal.Text = rs!FK_SUCURSAL
        
      End If

      
      
      Dim rsHistorico As New ADODB.Recordset
      If lblRequerimiento.Caption = "" Then
      lblRequerimiento.Caption = InputBox("Ingrese el requerimineto")
      End If
      
      
      
      sql = " SELECT    dbo.H_ESTADO_REQUE.IDREQUERIMIENTO, dbo.H_ESTADO_REQUE.IDESTADO, dbo.H_ESTADO_REQUE.IDPERSONAL, "
      sql = sql & " dbo.H_ESTADO_REQUE.CONTADOR , dbo.Personal.APELLIDO, dbo.Personal.NOMBRE "
      sql = sql & " FROM dbo.H_ESTADO_REQUE INNER JOIN "
      sql = sql & " dbo.PERSONAL ON dbo.H_ESTADO_REQUE.IDPERSONAL = dbo.PERSONAL.IDPERSONAL "
      sql = sql & " Where dbo.H_ESTADO_REQUE.IDREQUERIMIENTO = " & lblRequerimiento.Caption
      sql = sql & " ORDER BY dbo.H_ESTADO_REQUE.CONTADOR "
      rsHistorico.CursorLocation = adUseClient
      
      rsHistorico.Open sql, ConActiva, 0, 1
      
   If Not rsHistorico.EOF Then
        lblPersonalCarga.Caption = Trim(rsHistorico!Apellido) & " " & Trim(rsHistorico!Nombre)
   Else
        lblPersonalCarga.Caption = ""
   End If
      
      
      
    
   CargarGrillaDetalle rs!IDTIPOREQUERIMIENTO
   
    If CInt(Mid(lblEstado.Caption, 1, 4)) > 3 Then
        txtCargarLegajos.Text = "El estado > 1"
        txtCargarLegajos.Enabled = False
    Else
        txtCargarLegajos.Enabled = True
        txtCargarLegajos.Text = ""
    End If

End Sub

Private Sub chkEnvioPorCorreo_Click()

Dim sql As String

sql = " Update dbo.Requerimiento"
sql = sql & " SET ENVIOPORCORREO =" & chkEnvioPorCorreo.Value
sql = sql & " Where IDREQUERIMIENTO = " & lblRequerimiento.Caption

ExecutarSql sql

End Sub

Private Sub cmdAceptar_Click()
On Error GoTo salir:
    Dim sql As String
        If Trim(txtDescripcion.Text) <> "" Then
            sql = " Update dbo.Requerimiento "
            sql = sql & " SET DESCRIPCION ='" & Mid(Trim(MDIfrmInicio.StaInicio.Panels(3).Text) & " " & SysDate_DD_MM_YYYY_mm_ss & " " & Trim(txtDescripcion.Text) & vbCrLf & lblDescripcion.Caption, 1, 6999) & "'"
            sql = sql & " , DESCRIPCION_ACTUALIZADA = " & InputBox("Mensaje para 1 - Planta Mendoza , 2 - Pedidos , 3 - Cordoba ,  4 - San Luis ,  5 - Alsina , 6 - Gerencia ", "", 0)
            sql = sql & "  Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
            ExecutarSql sql
            Insert_Requerimiento_Historico_Descripcion lblRequerimiento.Caption, "'" & Trim(txtDescripcion.Text) & "'", 99, SysDate_mm_ss, ConActiva
            MsgBox "Grabado"
        Else
            MsgBox "Ingrese la descripción"
        End If
Exit Sub
salir:
 MsgBox Err.Description

End Sub

Private Sub cmdActual_Click()
On Error GoTo salir:
    cmdAceptar.Enabled = True
    txtDescripcion.Enabled = True
    Dim sql As String
    Dim rs   As New ADODB.Recordset
        sql = " SELECT     DESCRIPCION "
        sql = sql & " From dbo.Requerimiento"
        sql = sql & "  Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
        lblFechaModificacion.Caption = ""
        rs.Open sql, ConActiva, 0, 1
        If rs.EOF Then
            txtDescripcion.Text = ""
        Else
            txtDescripcion.Text = Trim(rs!DESCRIPCION)
        End If
Exit Sub

salir:

MsgBox Err.Description

End Sub

Private Sub cmdAdelante_Click()
On Error GoTo salir
    rsDescHistorico.MoveNext
    If Not rsDescHistorico.EOF Then
        lblDescripcion.Caption = rsDescHistorico!DESCRIPCION
        lblFechaModificacion.Caption = rsDescHistorico!Fecha
    End If
    
Exit Sub


salir:

MsgBox "Fin de archivo"
End Sub

Private Sub cmdAtras_Click()
On Error GoTo salir
    rsDescHistorico.MovePrevious
    If Not rsDescHistorico.EOF Then
       lblDescripcion.Caption = rsDescHistorico!DESCRIPCION
        lblFechaModificacion.Caption = rsDescHistorico!Fecha
    End If
Exit Sub


salir:

MsgBox "Fin de archivo"

End Sub

Private Sub cmdBorrarLegajos_Click()
    Dim sql As String
    Dim rs As ADODB.Recordset
    If txtCargarLegajos.Text = "" Then
    Exit Sub
    End If
    
   sql = "  DELETE FROM dbo.REQUELIBOSCAJAS "
sql = sql & "  WHERE   IDREQUERIMIENTOS = " & lblRequerimiento.Caption
sql = sql & " And CAJASLIBROS = " & txtCargarLegajos.Text
    
ExecutarSql sql

sql = " SELECT     COUNT(*) AS Cantidad "
sql = sql & " From dbo.REQUELIBOSCAJAS "
sql = sql & " Where IDREQUERIMIENTOS =" & lblRequerimiento.Caption
Set rs = New ADODB.Recordset
rs.Open sql, ConActiva, 0, 1

If Not rs.EOF Then
    txtCantidad.Text = rs!CANTIDAD
Else
    txtCantidad.Text = 0
End If

sql = " UPDATE    dbo.REQUERIMIENTO"
sql = sql & "  SET CANTIDAD =" & txtCantidad.Text
sql = sql & " Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
ExecutarSql sql
CargarGrillaDetalle Mid(lblTipoRequerimiento.Caption, 1, 4)

Dim Msg As String
Msg = InputBox("Ingrese el motivo de la mofidicacion")

sql = " Update dbo.Requerimiento "
sql = sql & " SET DESCRIPCION ='" & Trim(MDIfrmInicio.StaInicio.Panels(3).Text) & " " & SysDate_DD_MM_YYYY_mm_ss & " " & Msg & vbCrLf & lblDescripcion.Caption & "'"
sql = sql & " , DESCRIPCION_ACTUALIZADA = 1"
sql = sql & "  Where IDREQUERIMIENTO = " & lblRequerimiento.Caption


ExecutarSql sql


Insert_Requerimiento_Historico_Descripcion lblRequerimiento.Caption, "'" & Msg & "'", Trim(MDIfrmInicio.StaInicio.Panels(2).Text), SysDate_mm_ss, ConActiva




txtCargarLegajos.Text = ""

End Sub

Private Sub cmdBorrarNoEncontrados_Click()

Dim sql As String
Dim rs As ADODB.Recordset

sql = " SELECT     CAJASLIBROS"
sql = sql & " From REQUELIBOSCAJAS "
sql = sql & "  WHERE     IDREQUERIMIENTOS = " & lblRequerimiento.Caption
sql = sql & "  AND (ESTADO LIKE 'Para dar %')"

Set rs = New ADODB.Recordset


rs.Open sql, ConActiva, 0, 1

Do While Not rs.EOF
        sql = " Update LEGAJOS"
        sql = sql & " Set COD_ESTADO = 8"
        sql = sql & " Where Cod_cliente = " & Mid(lblCliente.Caption, 1, 4)
        sql = sql & " And ID_LEGAJO = " & rs!CAJASLIBROS
        ExecutarSql sql
        rs.MoveNext
Loop


sql = " DELETE FROM REQUELIBOSCAJAS"
sql = sql & "  WHERE     IDREQUERIMIENTOS = " & lblRequerimiento.Caption
sql = sql & "  AND ESTADO LIKE 'Para dar %'"
ExecutarSql sql

Set rs = New ADODB.Recordset
sql = " SELECT     COUNT(*) AS CANTIDAD"
sql = sql & " From REQUELIBOSCAJAS"
sql = sql & " Where IDREQUERIMIENTOS = " & lblRequerimiento.Caption
rs.Open sql, ConActiva, 0, 1

sql = " Update Requerimiento"
sql = sql & " SET  CANTIDAD =" & rs!CANTIDAD
sql = sql & " Where IDREQUERIMIENTO =  " & lblRequerimiento.Caption
ExecutarSql sql

Unload Me

End Sub

Private Sub cmdCrearNuevo_Click()
Dim sql As String

Dim conReq As New ADODB.Connection



Dim CantidadViejo As Integer
Dim CantidadNuevo As Integer
On Error GoTo salir
conReq.Open strConBasa

If InputBox("Ingrese la Password") <> "2338" Then
 MsgBox "CLAVE INCORRECTA"
 Exit Sub

Else

  
Dim REQUERIMAX As Long



sql = " INSERT INTO REQUERIMIENTO"
sql = sql & " ( ID_CLIENTE, IDPERSONAL, IDTIPORECEPCION, IDESTADO, IDTIPOREQUERIMIENTO, IDFAX, SECTOR, TELEFONO,"
sql = sql & " DESCRIPCION, SOLICITANTE, TOMO, FECHAENTREGA, FECHALIMITE, FECHARECEPCION, CANTIDAD, IDREMITO, TIEMPOTOTAL, PEDIDOCLIENTE,"
sql = sql & " ANULADO, COD_USUARIO_CLIENTE, COMPROMISO_ENTREGA, ENVIOPORCORREO, COD_HOJA_RUTA_TERMINADO, CANTIDAD_IMAGENES,"
sql = sql & " ENVIOTARDE, FECHA_SISTEMA, DESCRIPCION_ACTUALIZADA, FK_SUCURSAL, HORA_ARCHIVISTA, FLETE, COBRAR)"
sql = sql & " SELECT   ID_CLIENTE, IDPERSONAL, IDTIPORECEPCION, IDESTADO, IDTIPOREQUERIMIENTO, IDFAX, SECTOR, TELEFONO,"
sql = sql & "  DESCRIPCION, SOLICITANTE, TOMO, FECHAENTREGA, FECHALIMITE, FECHARECEPCION, CANTIDAD, IDREMITO, TIEMPOTOTAL, PEDIDOCLIENTE,"
sql = sql & " ANULADO, COD_USUARIO_CLIENTE, COMPROMISO_ENTREGA, ENVIOPORCORREO, COD_HOJA_RUTA_TERMINADO, CANTIDAD_IMAGENES,"
sql = sql & " EnvioTarde , FECHA_SISTEMA, DESCRIPCION_ACTUALIZADA, FK_SUCURSAL, HORA_ARCHIVISTA, FLETE, COBRAR"
sql = sql & " From Requerimiento"
sql = sql & " Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
conReq.Execute sql

REQUERIMAX = MaxIDRequerimiento

Dim i  As Integer

 Do While Not RsDetalle.EOF

    Select Case Trim(Mid(lblTipoRequerimiento.Caption, 1, 4))
    Case "0009"
    
        If IsNull(RsDetalle!ESTADO) Then
                  CantidadViejo = CantidadViejo + 1
               Else
                   sql = " INSERT INTO REQUELIBOSCAJAS "
                   sql = sql & " ( IDREQUERIMIENTOS ,CONTROL, CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO)"
                   sql = sql & "  SELECT  " & REQUERIMAX & "  ,CONTROL,  CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO"
                   sql = sql & " From REQUELIBOSCAJAS"
                   sql = sql & "  Where ID = " & RsDetalle!ID
                   conReq.Execute sql
                    sql = " DELETE FROM REQUELIBOSCAJAS  Where ID = " & RsDetalle!ID
                    conReq.Execute sql
                   CantidadNuevo = CantidadNuevo + 1
              End If
              RsDetalle.MoveNext
    
    
    
    Case "0001"
    
    
    
        If IsNull(RsDetalle!ESTADO) Then
                  CantidadViejo = CantidadViejo + 1
               Else
                   sql = " INSERT INTO REQUELIBOSCAJAS "
                   sql = sql & " ( IDREQUERIMIENTOS ,CONTROL, CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO)"
                   sql = sql & "  SELECT  " & REQUERIMAX & "  ,CONTROL,  CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO"
                   sql = sql & " From REQUELIBOSCAJAS"
                   sql = sql & "  Where ID = " & RsDetalle!ID
                   conReq.Execute sql
                   sql = " DELETE FROM REQUELIBOSCAJAS  Where ID = " & RsDetalle!ID
                   conReq.Execute sql
                   CantidadNuevo = CantidadNuevo + 1
              End If
              RsDetalle.MoveNext
    Case "0003"
               If IsNull(RsDetalle!ESTADO) Then
                  CantidadViejo = CantidadViejo + 1
               Else
                   sql = " INSERT INTO REQUELIBOSCAJAS "
                   sql = sql & " ( IDREQUERIMIENTOS ,CONTROL, CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO)"
                   sql = sql & "  SELECT  " & REQUERIMAX & "  ,CONTROL,  CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO"
                   sql = sql & " From REQUELIBOSCAJAS"
                   sql = sql & "  Where ID = " & RsDetalle!ID
                   conReq.Execute sql
                   sql = " DELETE FROM REQUELIBOSCAJAS  Where ID = " & RsDetalle!ID
                   conReq.Execute sql
                   CantidadNuevo = CantidadNuevo + 1
              End If
              RsDetalle.MoveNext
        
    Case "0010"
           If IsNull(RsDetalle!Control) Then
               CantidadViejo = CantidadViejo + 1
            Else
                sql = " INSERT INTO REQUELIBOSCAJAS "
                sql = sql & " ( IDREQUERIMIENTOS ,CONTROL, CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO)"
                sql = sql & "  SELECT  " & REQUERIMAX & "  ,CONTROL,  CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO"
                sql = sql & " From REQUELIBOSCAJAS"
                sql = sql & "  Where ID = " & RsDetalle!ID
                conReq.Execute sql
                sql = " DELETE FROM REQUELIBOSCAJAS  Where ID = " & RsDetalle!ID
                conReq.Execute sql
                CantidadNuevo = CantidadNuevo + 1
           End If
           RsDetalle.MoveNext
           
     Case "0011"
     
            If IsNull(RsDetalle!Control) Then
               CantidadViejo = CantidadViejo + 1
            Else
                sql = " INSERT INTO REQUELIBOSCAJAS "
                sql = sql & " ( IDREQUERIMIENTOS ,CONTROL, CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO)"
                sql = sql & "  SELECT  " & REQUERIMAX & "  ,CONTROL,  CAJASLIBROS,  FK_CAJAS, FK_LEGAJOS, FK_LIBROS, DEPOSITO, PERSONAL, ESTADO"
                sql = sql & " From REQUELIBOSCAJAS"
                sql = sql & "  Where ID = " & RsDetalle!ID
                conReq.Execute sql
                sql = " DELETE FROM REQUELIBOSCAJAS  Where ID = " & RsDetalle!ID
                conReq.Execute sql
                CantidadNuevo = CantidadNuevo + 1
           End If
           RsDetalle.MoveNext
         
     End Select
 Loop
 
    sql = " Update Requerimiento SET   CANTIDAD = " & CantidadViejo & " Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
    conReq.Execute sql
    sql = " Update Requerimiento SET   IDESTADO = 4 ,  CANTIDAD = " & CantidadNuevo & " Where IDREQUERIMIENTO = " & REQUERIMAX
    conReq.Execute sql
    MsgBox "Nuevo requerimiento es  " & REQUERIMAX


End If
Exit Sub

salir:

MsgBox Err.Description
End Sub

Private Sub cmdfechaCompromiso_Click()
Dim sql As String
Dim Msg As String


Msg = InputBox("Ingrese el motivo de a mofidicacion")

sql = " Update dbo.Requerimiento "
sql = sql & " SET DESCRIPCION ='" & Trim(MDIfrmInicio.StaInicio.Panels(3).Text) & " " & SysDate_DD_MM_YYYY_mm_ss & " " & Msg & vbCrLf & lblDescripcion.Caption & "'"
sql = sql & "  Where IDREQUERIMIENTO = " & lblRequerimiento.Caption

ExecutarSql sql


Insert_Requerimiento_Historico_Descripcion lblRequerimiento.Caption, "'" & Msg & "'", Trim(MDIfrmInicio.StaInicio.Panels(2).Text), SysDate_mm_ss, ConActiva


sql = " Update dbo.Requerimiento "
sql = sql & " SET FECHAENTREGA = " & FechaFormato(txtFechaCompromiso.Text)
sql = sql & ", COMPROMISO_ENTREGA = '" & Trim(cboHoraDia.Text) & "'"
If Trim(txtHorasArchivista.Text) = "" Then
    sql = sql & " , HORA_ARCHIVISTA = NULL "
Else
    sql = sql & " , HORA_ARCHIVISTA = '" & Trim(txtHorasArchivista.Text) & "'"
End If


    sql = sql & " , FLETE= '" & chkFlete.Value & "'"

    sql = sql & " , COBRAR = '" & chkCobrar.Value & "'"
    
     sql = sql & " , CANTIDAD_IMAGENES = " & txtImagenes.Text
     sql = sql & " ,FK_SUCURSAL='" & cboSucursal.Text & "'"
 sql = sql & " ,CANTIDAD  =" & txtCantidad.Text



sql = sql & " Where IDREQUERIMIENTO = " & lblRequerimiento.Caption

ExecutarSql sql
MsgBox "El REGISTRO actualizado", vbInformation

End Sub

Private Sub cmdHistorico_Click()
    Set rsDescHistorico = New ADODB.Recordset
    Dim sql As String
    sql = " SELECT     ID, FK_REQUERIMIENTO, DESCRIPCION, FK_USUARIO, FECHA"
    sql = sql & " From dbo.REQUERIMIENTO_DESCRIPCION_HISTORICO"
    sql = sql & " Where FK_REQUERIMIENTO = " & lblRequerimiento
    sql = sql & " ORDER BY ID"
    
    rsDescHistorico.Open sql, ConActiva, 0, 1
    cmdAceptar.Enabled = False
    txtDescripcion.Enabled = False
    
    If Not rsDescHistorico.EOF Then
        If Not IsNull(rsDescHistorico!DESCRIPCION) Then
    
        txtDescripcion.Text = rsDescHistorico!DESCRIPCION
        lblFechaModificacion.Caption = rsDescHistorico!Fecha
    Else
    
        txtDescripcion.Text = ""
        lblFechaModificacion.Caption = ""
    
    
    End If
    
        
    End If
    
    
    
End Sub

Public Sub CargarGrillaDetalle(IDTIPOREQUERIMIENTO As Integer)
Dim rs As New ADODB.Recordset
Set RsDetalle = New ADODB.Recordset

    RsDetalle.CursorLocation = adUseClient
    
    Dim sql  As String
    
        If IDTIPOREQUERIMIENTO = 10 Or IDTIPOREQUERIMIENTO = 11 Then
    
                sql = "SELECT REQUELIBOSCAJAS.ID, REQUELIBOSCAJAS.DEPOSITO, REQUELIBOSCAJAS.PERSONAL, REQUELIBOSCAJAS.ESTADO, REQUELIBOSCAJAS.CONTROL,"
                sql = sql & vbCrLf & " REQUERIMIENTO.ID_CLIENTE, 0 AS ID_CLIENTE_LEGAJO, '' AS DESCRIPCION_REMITO, REQUERIMIENTO.IDREQUERIMIENTO, REQUELIBOSCAJAS.DETALLE"
                sql = sql & vbCrLf & " FROM REQUELIBOSCAJAS INNER JOIN"
                sql = sql & vbCrLf & " REQUERIMIENTO ON REQUELIBOSCAJAS.IDREQUERIMIENTOS = REQUERIMIENTO.IDREQUERIMIENTO"
                sql = sql & vbCrLf & " Where (REQUELIBOSCAJAS.CAJASLIBROS Is Null) And "
                sql = sql & vbCrLf & " REQUERIMIENTO.IDREQUERIMIENTO = " & lblRequerimiento.Caption
                sql = sql & vbCrLf & " Union All"
                sql = sql & vbCrLf & " SELECT REQUELIBOSCAJAS_1.ID, REQUELIBOSCAJAS_1.DEPOSITO, REQUELIBOSCAJAS_1.PERSONAL, REQUELIBOSCAJAS_1.ESTADO, REQUELIBOSCAJAS_1.CONTROL,"
                sql = sql & vbCrLf & " REQUERIMIENTO_1.ID_CLIENTE, LEGAJOS.ID_CLIENTE_LEGAJO, LEGAJOS.DESCRIPCION_REMITO, REQUERIMIENTO_1.IDREQUERIMIENTO,"
                sql = sql & vbCrLf & " REQUELIBOSCAJAS_1.DETALLE"
                sql = sql & vbCrLf & " FROM REQUELIBOSCAJAS AS REQUELIBOSCAJAS_1 INNER JOIN"
                sql = sql & vbCrLf & " REQUERIMIENTO AS REQUERIMIENTO_1 ON REQUELIBOSCAJAS_1.IDREQUERIMIENTOS = REQUERIMIENTO_1.IDREQUERIMIENTO INNER JOIN"
                sql = sql & vbCrLf & " LEGAJOS ON REQUERIMIENTO_1.ID_CLIENTE = LEGAJOS.COD_CLIENTE AND REQUELIBOSCAJAS_1.CAJASLIBROS = LEGAJOS.ID_CLIENTE_LEGAJO"
                sql = sql & vbCrLf & " Where REQUERIMIENTO_1.IDREQUERIMIENTO =  " & lblRequerimiento.Caption
                sql = sql & vbCrLf & " ORDER BY  REQUELIBOSCAJAS.ESTADO desc , REQUELIBOSCAJAS.ID"
                 
                 
                 
                 
        Else
   
                    sql = "  SELECT     ID, DEPOSITO, PERSONAL, ESTADO ,CONTROL,  CAJASLIBROS , detalle"
                    sql = sql & vbCrLf & " From dbo.REQUELIBOSCAJAS"
                    sql = sql & vbCrLf & " Where IDREQUERIMIENTOS = " & lblRequerimiento.Caption
                    sql = sql & vbCrLf & " ORDER BY ESTADO DESC , CAJASLIBROS  "
                    
         End If
    
    
    RsDetalle.Open sql, ConActiva, adOpenKeyset, adLockPessimistic
Set grdDetalle.DataSource = RsDetalle.DataSource

grdDetalle.Refresh

End Sub

Private Sub cmdInsertOrdenBusqueda_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim DESCRIPCION As String
        sql = " SELECT     IDREQUERIMIENTOS"
        sql = sql & "  From basasql.dbo.REQUELIBOSCAJAS"
        sql = sql & "  Where IDREQUERIMIENTOS = " & lblRequerimiento.Caption
        Set rs = New ADODB.Recordset
        rs.Open sql, strConBasa
        If rs.EOF Then
                Set rs = New ADODB.Recordset
                sql = " SELECT      NRO_CAJA, ESTANTERIA, HORIZONTAL, VERTICAL,LEGAJO,  DETALLE"
                sql = sql & "  From V_TEM_BUSQUEDA"
                sql = sql & " Where ID_LOTE_BUSQUEDA = " & InputBox("Ingrese el numero de reporte", "", "0")
                Dim CANTIDAD As Integer
                rs.Open sql, strConBasa
                Do While Not rs.EOF
                    DESCRIPCION = rs!NRO_CAJA & " E:" & rs!ESTANTERIA & " H:" & rs!Horizontal & " V:" & rs!Vertical & " Busqueda:  " & rs!LEGAJO & "  " & rs!DETALLE
                    sql = " INSERT INTO dbo.REQUELIBOSCAJAS"
                    sql = sql & " (IDREQUERIMIENTOS, DETALLE , ESTADO)"
                    sql = sql & " VALUES     (" & lblRequerimiento.Caption & ",'" & Mid(DESCRIPCION, 1, 249) & "',NULL)"
                    ExecutarSql sql
                    CANTIDAD = CANTIDAD + 1
                    rs.MoveNext
                Loop
                sql = " Update dbo.Requerimiento"
                sql = sql & " SET CANTIDAD= " & CANTIDAD
                sql = sql & " Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
                ExecutarSql sql
        Else
            MsgBox "Ya existen registro "
        End If
End Sub




Private Sub cmdLeida_Click()
On Error GoTo salir:

Dim sql As String

sql = " Update dbo.Requerimiento"
sql = sql & " SET  DESCRIPCION_ACTUALIZADA = NULL "
sql = sql & " , DESCRIPCION ='Leido " & Trim(MDIfrmInicio.StaInicio.Panels(3).Text) & " " & SysDate_DD_MM_YYYY_mm_ss & " " & Trim(txtDescripcion.Text) & vbCrLf & lblDescripcion.Caption & "'"
sql = sql & " Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
ExecutarSql sql
Exit Sub
salir:
MsgBox Err.Description
End Sub

Private Sub cmdMarcarComoVerificado_Click()
Dim sql As String
sql = " Update REQUELIBOSCAJAS"
sql = sql & " SET              CONTROL ='AUTOMATICO'"
sql = sql & " WHERE     (ESTADO LIKE 'encontrado%')"
sql = sql & " AND IDREQUERIMIENTOS = " & lblRequerimiento.Caption
ExecutarSql sql
End Sub

Private Sub cmdOrdenBusqueda_Click()
On Error GoTo salir:
Dim Reporte As Integer
     Reporte = InputBox("Ingrese el numero de reporte", "", "0")
    ReporteBusqueda Reporte
    Exit Sub
salir:
MsgBox Err.Description


End Sub

Private Sub lblCantidad_Click()

End Sub

Private Sub Command1_Click()
 Clipboard.Clear
 Clipboard.SetText lblDescripcion.Caption
 MsgBox "COPIADO"
End Sub

Private Sub Command2_Click()
CopiarDatosGrilla grdDetalle
End Sub

Private Sub Form_Load()
'txtFechaCompromiso.Enabled = False
'If MDIfrmInicio.StaInicio.Panels(2).Text = 48 Or MDIfrmInicio.StaInicio.Panels(2).Text = 47 Then
'    txtFechaCompromiso.Enabled = True
'End If




End Sub

Private Sub grdDetalle_DblClick()
Dim rs As New ADODB.Recordset
Dim sql As String

If InputBox("Ingrese la clave") = "2338" Then
    sql = " SELECT  CANTIDAD From Requerimiento Where IDREQUERIMIENTO =" & lblRequerimiento.Caption
    rs.Open sql, ConActiva, 0, 1
    If Not rs.EOF Then
        ExecutarSql " UPDATE    REQUERIMIENTO SET CANTIDAD = " & rs!CANTIDAD - 1 & " Where IDREQUERIMIENTO =" & lblRequerimiento.Caption
        ExecutarSql " DELETE FROM REQUELIBOSCAJAS Where ID = " & grdDetalle.Text
        Unload Me
    End If
    
  Else
  MsgBox "CLave incorrecta"


End If
End Sub

Private Sub grdDetalle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu mnuGrillaDetalle

End If

End Sub

Private Sub mnuBorrarElementos_Click()
Dim sql As String
Dim rs As New ADODB.Recordset
grdDetalle.Col = 0
sql = " DELETE FROM basasql.dbo.REQUELIBOSCAJAS"
sql = sql & " Where ID = " & grdDetalle.Text
ExecutarSql sql

sql = " SELECT     COUNT(*) AS Cantidad "
sql = sql & " From dbo.REQUELIBOSCAJAS "
sql = sql & " Where IDREQUERIMIENTOS =" & lblRequerimiento.Caption
Set rs = New ADODB.Recordset
rs.Open sql, strConBasa
If Not rs.EOF Then
    sql = " UPDATE    dbo.REQUERIMIENTO "
    sql = sql & "  SET CANTIDAD =" & rs!CANTIDAD
    sql = sql & " Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
    ExecutarSql sql
 
End If
MsgBox "Actualizacion terminada Cargar nuevamente el requerimiento"

End Sub

Private Sub mnuIngresarLegajo_Click()
Dim sql As String
Dim Etiqueta As String
grdDetalle.Col = 0
Etiqueta = InputBox("Ingrese el legajo")
sql = " Update REQUELIBOSCAJAS"
sql = sql & " SET CAJASLIBROS =" & Etiqueta
sql = sql & " , FK_LEGAJOS =" & Etiqueta
sql = sql & " Where ID = " & grdDetalle.Text
ExecutarSql sql
MsgBox "Actualizacion terminada Cargar nuevamente el requerimiento"


End Sub

Private Sub txtCargarLegajos_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
        
Dim rs As New ADODB.Recordset
Dim sql As String

If Mid(txtCargarLegajos.Text, 1, 2) = "l2" Then
txtCargarLegajos.Text = Mid(txtCargarLegajos.Text, 3)
End If



sql = " SELECT     IDREQUERIMIENTOS, CAJASLIBROS"
sql = sql & " From dbo.REQUELIBOSCAJAS"
sql = sql & " Where IDREQUERIMIENTOS =" & lblRequerimiento.Caption
sql = sql & " And CAJASLIBROS = " & txtCargarLegajos.Text


Set rs = New ADODB.Recordset
rs.Open sql, ConActiva, 0, 1
If Not rs.EOF Then
    MsgBox "Legajos Cargado", vbInformation
    Exit Sub
End If

If CInt(Mid(lblTipoRequerimiento.Caption, 1, 5)) = 11 Or CInt(Mid(lblTipoRequerimiento.Caption, 1, 5)) = 10 Then
    sql = " SELECT ID_CLIENTE_LEGAJO, COD_CLIENTE, COD_ESTADO"
    sql = sql & " From dbo.LEGAJOS"
    sql = sql & " Where Cod_cliente =  " & Mid(lblCliente.Caption, 1, 4)
    sql = sql & " And ID_CLIENTE_LEGAJO =" & txtCargarLegajos.Text
Else
    sql = " SELECT ESTADO as COD_ESTADO, COD_CLIENTE, NRO_CAJA "
    sql = sql & "  From CONTENEDOR"
    sql = sql & "  Where Cod_cliente = " & Mid(lblCliente.Caption, 1, 4)
    sql = sql & "  And NRO_CAJA = " & txtCargarLegajos.Text
End If


Set rs = New ADODB.Recordset

rs.Open sql, ConActiva, 0, 1
If Not rs.EOF Then
    If rs!COD_ESTADO <> 2 Then
        MsgBox "Estado Incorrecto", vbCritical
        Exit Sub
    End If
  Else
  MsgBox "El legajo No existe", vbInformation
  Exit Sub
End If
    
    
sql = " INSERT INTO dbo.REQUELIBOSCAJAS"
sql = sql & " (IDREQUERIMIENTOS, CAJASLIBROS, FK_LEGAJOS, ESTADO)"
sql = sql & " VALUES     (" & lblRequerimiento.Caption & "," & txtCargarLegajos.Text & "," & txtCargarLegajos.Text & ",'" & "Encontrado " & SysDate_DD_MM_YYYY_mm_ss & "')"
ExecutarSql sql

sql = " SELECT     COUNT(*) AS Cantidad"
sql = sql & " From dbo.REQUELIBOSCAJAS"
sql = sql & " Where IDREQUERIMIENTOS =" & lblRequerimiento.Caption
Set rs = New ADODB.Recordset
rs.Open sql, ConActiva, 0, 1

If Not rs.EOF Then
  txtCantidad.Text = rs!CANTIDAD
Else
  txtCantidad.Text = 0
End If

sql = " UPDATE    dbo.REQUERIMIENTO "
sql = sql & "  SET CANTIDAD =" & txtCantidad.Text
sql = sql & " Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
ExecutarSql sql
CargarGrillaDetalle Mid(lblTipoRequerimiento.Caption, 1, 4)
txtCargarLegajos.Text = ""
End If


End Sub

Private Sub txtCobrar_Change()

End Sub



Private Sub txtEstado_KeyPress(KeyAscii As Integer)
Dim sql As String
Dim c As Integer
If KeyAscii = 13 Then
     
     If UCase(Mid(txtEstado.Text, 1, 2)) = "RD" Then
                sql = " UPDATE REQUELIBOSCAJAS "
                sql = sql & " Set ESTADO =NUll"
                sql = sql & " ,  Personal =NULL "
                sql = sql & "  Where ID = " & Mid(txtEstado.Text, 3)
               
                c = ExecutarSql(sql)
                If c <> 1 Then
                    MsgBox "No se actualizo"
                End If
                
                txtEstado.Text = ""
                Beep
                Exit Sub
         End If

    If CInt(Mid(lblTipoRequerimiento.Caption, 1, 4)) = 10 Or CInt(Mid(lblTipoRequerimiento.Caption, 1, 4)) = 11 Then
        
      
        If Trim(MDIfrmInicio.StaInicio.Panels(2).Text) = 17 Then
           If UCase(Mid(txtEstado.Text, 1, 2)) = "L2" Then
               sql = " Update REQUELIBOSCAJAS SET ESTADO = 'Encontrado " & SysDate_DD_MM_YYYY_mm_ss & "'"
               sql = sql & " Where IDREQUERIMIENTOS = " & lblRequerimiento.Caption
               sql = sql & "   And CAJASLIBROS = " & Mid(txtEstado.Text, 3)
               If ExecutarSql(sql) <> 1 Then
                   MsgBox "El Elemento " & CLng(Mid(txtEstado.Text, 3, 10)) & " No pertenece al requerimiento", vbCritical
               End If
           End If
            If UCase(Mid(txtEstado.Text, 1, 2)) = "12" Then
                sql = " Update REQUELIBOSCAJAS SET ESTADO = 'Encontrado " & SysDate_DD_MM_YYYY_mm_ss & "'"
                sql = sql & " Where IDREQUERIMIENTOS = " & lblRequerimiento.Caption
                sql = sql & "   And CAJASLIBROS = " & CLng(Mid(txtEstado.Text, 3, 10))
                If ExecutarSql(sql) <> 1 Then
                    MsgBox "El Elemento " & CLng(Mid(txtEstado.Text, 3, 10)) & " No pertenece al requerimiento", vbCritical
                End If
           End If
            
            
       Else
           If UCase(Mid(txtEstado.Text, 1, 2)) = "L1" Then
               sql = " Update REQUELIBOSCAJAS SET CONTROL ='" & SysDate_DD_MM_YYYY_mm_ss & "'"
               sql = sql & " Where IDREQUERIMIENTOS = " & lblRequerimiento.Caption
               sql = sql & "   And CAJASLIBROS = " & Mid(txtEstado.Text, 6)
               If ExecutarSql(sql) <> 1 Then
                   MsgBox "El Elemento " & CLng(Mid(txtEstado.Text, 3, 10)) & " No pertenece al requerimiento", vbCritical
               End If
           End If
       
            If UCase(Mid(txtEstado.Text, 1, 2)) = "L2" Then
               sql = " Update REQUELIBOSCAJAS SET CONTROL ='" & SysDate_DD_MM_YYYY_mm_ss & "'"
               sql = sql & " Where IDREQUERIMIENTOS = " & lblRequerimiento.Caption
               sql = sql & "   And CAJASLIBROS = " & Mid(txtEstado.Text, 3)
               If ExecutarSql(sql) <> 1 Then
                   MsgBox "El Elemento " & CLng(Mid(txtEstado.Text, 3, 10)) & " No pertenece al requerimiento", vbCritical
               End If
           End If
            If UCase(Mid(txtEstado.Text, 1, 2)) = "12" Then
        sql = " Update REQUELIBOSCAJAS SET CONTROL ='" & SysDate_DD_MM_YYYY_mm_ss & "'"
        sql = sql & " Where IDREQUERIMIENTOS = " & lblRequerimiento.Caption
        sql = sql & "   And CAJASLIBROS = " & CLng(Mid(txtEstado.Text, 3, 10))
        If ExecutarSql(sql) <> 1 Then
            MsgBox "El Elemento " & CLng(Mid(txtEstado.Text, 3, 10)) & " No pertenece al requerimiento", vbCritical
        End If
     
        End If
    
    End If
    End If
    
txtEstado.Text = ""
End If


End Sub

