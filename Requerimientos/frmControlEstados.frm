VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmControlEstados 
   Caption         =   "Control de estados"
   ClientHeight    =   7800
   ClientLeft      =   1740
   ClientTop       =   2025
   ClientWidth     =   12270
   FillColor       =   &H000000C0&
   ForeColor       =   &H80000013&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   12270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   12135
      Begin VB.CheckBox chkVerAnulados 
         Caption         =   "Ver Anulados"
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
         Left            =   3900
         TabIndex        =   35
         Top             =   720
         Width           =   1455
      End
      Begin Controles.cltGenerico ctlPersonal 
         Height          =   315
         Left            =   900
         TabIndex        =   34
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
      End
      Begin Controles.cltGenerico cltCliente 
         Height          =   315
         Left            =   900
         TabIndex        =   31
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
      End
      Begin Controles.cltGenerico ctlTipoRequerimiento 
         Height          =   315
         Left            =   900
         TabIndex        =   33
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
      End
      Begin Controles.ctlClienteUsuario ctlClienteUsuario 
         Height          =   315
         Left            =   5160
         TabIndex        =   30
         Top             =   1740
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
      End
      Begin VB.OptionButton optBaja 
         Caption         =   "Baja"
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
         Left            =   8220
         TabIndex        =   29
         Top             =   720
         Width           =   675
      End
      Begin VB.OptionButton OptNo 
         Caption         =   "NO"
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
         Left            =   9720
         TabIndex        =   27
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton optSI 
         Caption         =   "SI"
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
         Left            =   9060
         TabIndex        =   26
         Top             =   720
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.ComboBox cboDeposito 
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
         ItemData        =   "frmControlEstados.frx":0000
         Left            =   900
         List            =   "frmControlEstados.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtPersonal 
         BackColor       =   &H0080C0FF&
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
         Left            =   6420
         TabIndex        =   23
         Top             =   660
         Width           =   375
      End
      Begin VB.TextBox txtDias 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7680
         TabIndex        =   22
         Text            =   "5"
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdPegar 
         Caption         =   "..."
         Height          =   315
         Left            =   7980
         TabIndex        =   20
         Top             =   1200
         Width           =   315
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
         ItemData        =   "frmControlEstados.frx":0052
         Left            =   900
         List            =   "frmControlEstados.frx":005F
         TabIndex        =   18
         Top             =   720
         Width           =   2715
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excel"
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
         Left            =   8520
         TabIndex        =   17
         Top             =   1680
         Width           =   1035
      End
      Begin VB.TextBox txtRequerimientoEncontrado 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10500
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox chkHojaRutaSinAsignar 
         Caption         =   "Sin asignar hoja de ruta"
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
         Left            =   3720
         TabIndex        =   11
         Top             =   240
         Width           =   2355
      End
      Begin VB.CheckBox chkHojaRuta 
         Caption         =   "Selección"
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
         Left            =   10500
         TabIndex        =   9
         Top             =   300
         Width           =   1275
      End
      Begin VB.CommandButton cmdSalida 
         Caption         =   "Salida"
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
         Left            =   9660
         TabIndex        =   8
         Top             =   1680
         Width           =   1035
      End
      Begin VB.TextBox txtFiltro 
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
         Left            =   5160
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   2655
      End
      Begin VB.ComboBox cboFiltro 
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
         ItemData        =   "frmControlEstados.frx":007F
         Left            =   900
         List            =   "frmControlEstados.frx":00B9
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   2715
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refrescar"
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
         Left            =   10800
         TabIndex        =   5
         Top             =   1680
         Width           =   1035
      End
      Begin VB.CheckBox chkRefrescar 
         Caption         =   "Refrescar Automático"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   8280
         TabIndex        =   4
         Top             =   300
         Value           =   1  'Checked
         Width           =   2475
      End
      Begin VB.CheckBox chkRemitosPendientes 
         Caption         =   "Remitos Pendientes"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkPendientes 
         Caption         =   "Pendientes"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblFiltro 
         Caption         =   "Solicito:"
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
         Left            =   3840
         TabIndex        =   16
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Encontrado"
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
         Left            =   7140
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Personal"
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
         Left            =   5520
         TabIndex        =   24
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Ver los ultimos dias"
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
         Left            =   6180
         TabIndex        =   21
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "Sucursal:"
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
         TabIndex        =   19
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblUsuarioCliente 
         Caption         =   "Solicito:"
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
         Left            =   3840
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Filtro:"
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
         TabIndex        =   14
         Top             =   1200
         Width           =   675
      End
      Begin VB.Label lblCliente 
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
         TabIndex        =   13
         Top             =   1740
         Width           =   675
      End
   End
   Begin VB.Timer timRefrescarEstados 
      Interval        =   60000
      Left            =   540
      Top             =   6600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":01FF
            Key             =   "E2"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":0C11
            Key             =   "ENTRAMITE"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":0CE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":1001
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":1453
            Key             =   "E1"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":1E65
            Key             =   "E3"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":1FBF
            Key             =   "E4"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":2119
            Key             =   "SRojo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":256B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":2935
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":2A8F
            Key             =   "SAmarillo"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":2EE1
            Key             =   "SVerde"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":3333
            Key             =   "TIPOREQUERIMIENTO"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":3785
            Key             =   "E10"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":445F
            Key             =   "FUERATERMINO"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":45B9
            Key             =   "E17"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":4A0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":541D
            Key             =   "E7"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":5E2F
            Key             =   "E6"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":6841
            Key             =   "E5"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmControlEstados.frx":7253
            Key             =   "E8"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView trvEstado 
      CausesValidation=   0   'False
      Height          =   5295
      Left            =   60
      TabIndex        =   0
      Top             =   2400
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   5
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox cryRemito 
      Height          =   480
      Left            =   1980
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   10
      Top             =   2400
      Width           =   1200
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   555
      Left            =   3600
      TabIndex        =   32
      Top             =   2760
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   979
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
   Begin VB.Menu mnuHojaRuta 
      Caption         =   "Hoja de ruta"
      Begin VB.Menu mnuControlFaltantes 
         Caption         =   "Control Faltantes"
      End
      Begin VB.Menu mnuPasarSeleccion 
         Caption         =   "Pasar Seleccion"
      End
      Begin VB.Menu mnuPasarTodosRequerimientos 
         Caption         =   "Pasar todos los requerimientos"
      End
      Begin VB.Menu mnuExpanderDia 
         Caption         =   "Expander Dia"
      End
      Begin VB.Menu mnuCantidadElementos 
         Caption         =   "Cantidad de Elementos"
      End
   End
   Begin VB.Menu mnu1234 
      Caption         =   "1234"
      Visible         =   0   'False
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuEstado 
         Caption         =   "Imprimir C/Estado"
      End
   End
   Begin VB.Menu mnuBusquedaDocumentos 
      Caption         =   "BusquedaDocumentos"
      Visible         =   0   'False
      Begin VB.Menu mnuBuscarDocumentosImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuBusquedaDocumentoAsignacion 
         Caption         =   "Asignación Busqueda"
      End
      Begin VB.Menu mnuBusquedaDocumentoTerminado 
         Caption         =   "Terminado"
      End
      Begin VB.Menu mnuBusquedaDocumentoHojaRuta 
         Caption         =   "Hojas de rutas"
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "General"
      Begin VB.Menu mnuGeneralImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuGeneralCambioDeEstado 
         Caption         =   "Cambio de Estado"
      End
      Begin VB.Menu mnuGeneralTerminado 
         Caption         =   "Terminado"
      End
      Begin VB.Menu mnuHojasdeRutas 
         Caption         =   "Hoja de rutas"
      End
   End
   Begin VB.Menu mnuCajasVacias 
      Caption         =   "CajasVacias"
      Begin VB.Menu mnuCajasVaciasImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuCajasVaciasCambioEstado 
         Caption         =   "Imprimir c/Estado"
      End
      Begin VB.Menu mnuCajasVaciasReImprimirRotulos 
         Caption         =   "Re Imprimir Rotulos"
      End
      Begin VB.Menu mnuCajasVaciasImprimirRotuloEtiqueta 
         Caption         =   "Imprimir Rotulo Etiqueta"
      End
      Begin VB.Menu mnuCajasVaciasRemito 
         Caption         =   "Remitos"
      End
      Begin VB.Menu mnuPlanillas 
         Caption         =   "Planillas de referencias"
      End
      Begin VB.Menu mnuCajasVaciasReImprimirRemito 
         Caption         =   "Re Imprimir Remitos"
      End
      Begin VB.Menu mnuHRVacias 
         Caption         =   "Hoja de ruta"
      End
   End
   Begin VB.Menu mnuLegajos 
      Caption         =   "Legajos"
      Visible         =   0   'False
      Begin VB.Menu mnuLegajosImprimir 
         Caption         =   "Imprimir Legajos"
      End
      Begin VB.Menu mnuLegajosImprimirOrden 
         Caption         =   "Imprimir Orden"
      End
      Begin VB.Menu mnuLegajosBuscarLegajos 
         Caption         =   "Buscar Legajos"
      End
      Begin VB.Menu mnuLegajosEncontrados 
         Caption         =   "Legajos Encontrado"
      End
      Begin VB.Menu mnuLegajosRemitos 
         Caption         =   "Remito"
      End
      Begin VB.Menu mnuLegajosReImprimirRemito 
         Caption         =   "Re Imprimir Remito"
      End
      Begin VB.Menu mnuLegajosImprimirEtiquetas 
         Caption         =   "Imprimir Etiquetas"
      End
      Begin VB.Menu mnuLegajosHR 
         Caption         =   "Hoja de Ruta"
      End
      Begin VB.Menu mnuLegajosTerminado 
         Caption         =   "Terminado"
      End
   End
   Begin VB.Menu mnuCajas 
      Caption         =   "Cajas"
      Visible         =   0   'False
      Begin VB.Menu mnuCajasImprimir 
         Caption         =   "Imprimir "
      End
      Begin VB.Menu mnuCajasBuscarCajas 
         Caption         =   "Buscar Cajas"
      End
      Begin VB.Menu mnuCajasEncontradas 
         Caption         =   "Cajas Encontradas"
      End
      Begin VB.Menu mnuCajasRemito 
         Caption         =   "Remito de Cajas"
      End
      Begin VB.Menu mnuCajasReImprimirRemito 
         Caption         =   "Re Imprimir Remito Cajas"
      End
      Begin VB.Menu mnuCajasHojaRuta 
         Caption         =   "Hoja de Ruta"
      End
      Begin VB.Menu mnuCajasTerminado 
         Caption         =   "Terminado"
      End
   End
   Begin VB.Menu mnuLibros 
      Caption         =   "Libros"
      Visible         =   0   'False
      Begin VB.Menu mnuLibrosImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuLibrosBuscarLibros 
         Caption         =   "Buscar Libros"
      End
      Begin VB.Menu mnuLibrosRemitoDeLibros 
         Caption         =   "Remito de Libros"
      End
      Begin VB.Menu mnuLibrosReImprimirRemitoLibros 
         Caption         =   "Re Imprimir Remito Libros"
      End
      Begin VB.Menu mnuLibrosHojaderuta 
         Caption         =   "Hoja de ruta"
      End
   End
   Begin VB.Menu mnuReferenciaTrasvase 
      Caption         =   "ReferenciaTrasvase"
      Visible         =   0   'False
      Begin VB.Menu mnuReferenciaTrasvaseImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuReferenciaTrasvaseAsignacion 
         Caption         =   "Asignación"
      End
      Begin VB.Menu mnuReferenciaTrasvaseDiccSector 
         Caption         =   "Dicc Sector"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReferenciaTrasvaseDiccCompleto 
         Caption         =   "Dicc Completo"
      End
      Begin VB.Menu mnuReferenciaTrasvaseTerminado 
         Caption         =   "Terminado"
      End
      Begin VB.Menu mnuHojaderuta 
         Caption         =   "Hoja de ruta"
      End
   End
   Begin VB.Menu mnuTipoRequerimientoCajasVacias 
      Caption         =   "TipoRequerimientoCajasVacias"
      Visible         =   0   'False
      Begin VB.Menu mnuTipoRequerimientoCajasVaciasEtiquetas 
         Caption         =   "Re Imprimir Etiquetas"
      End
   End
   Begin VB.Menu mnuTipoRequerimientosLegajos 
      Caption         =   "TipoRequerimientosLegajos"
      Visible         =   0   'False
      Begin VB.Menu mnuTipoLegajosEtiquetas 
         Caption         =   "Etiquetas"
      End
      Begin VB.Menu mnuTipoRequerimientoLegajosImprimir 
         Caption         =   "Imprimir"
      End
   End
   Begin VB.Menu mnuTipoRequerimientoCajas 
      Caption         =   "TipoRequerimientoCajas"
      Visible         =   0   'False
      Begin VB.Menu mnuTipoRequerimientoCajasImprimirTodo 
         Caption         =   "Imprimir Todo Requerimiento"
      End
      Begin VB.Menu mnuTipoRequerimientoCajasImprimir 
         Caption         =   "Imprimir"
      End
   End
   Begin VB.Menu mnuConsultasDigitales 
      Caption         =   "Consultas Digitales"
      Visible         =   0   'False
      Begin VB.Menu mnuConsultas_digitales_Imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuConsultas_Digitales_Asignar_Tarea 
         Caption         =   "Asignar Tarea"
      End
      Begin VB.Menu mnuConsultas_Digitales_Imagenes_Encontradas 
         Caption         =   "ImagenesEncontradas"
      End
      Begin VB.Menu mnuConsultas_Digitales_Cantidad_Imagenes 
         Caption         =   "Cantidad de Imagenes"
      End
      Begin VB.Menu mnuConsultas_digitales_Finalizado 
         Caption         =   "Finalizado"
      End
   End
End
Attribute VB_Name = "frmControlEstados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Refrescar As Boolean
Dim FechaRequerimiento As String
Dim rsRequerimientos As ADODB.Recordset
Dim nodoSelecionado As String
Dim C_Fecha As String
Dim C_MañanaTarde As String
Dim C_Sucursal As String
Dim C_TipoRequerimiento As String


Dim KeyPadreTipo  As String
Dim KeyPadreMañana As String
Dim KeyPadreFecha As String
Dim KeyPadreSucursal As String
Dim KeyRequerimiento As String
Dim conreque As New ADODB.Connection
  


Private Sub cboFiltro_Click()
    ctlClienteUsuario.Valor = Null
    ctlClienteUsuario.Visible = False
    cltCliente.Valor = Null
    cltCliente.Visible = False
    txtFiltro.Visible = False
    cmdPegar.Visible = False
    lblUsuarioCliente.Visible = False
    lblCliente.Visible = False
    lblFiltro.Caption = ""
    txtFiltro = ""
    cboDeposito.Visible = False
    ctlTipoRequerimiento.Visible = False
    ctlPersonal.Visible = False
    
Select Case cboFiltro.Text
Case "Ninguno"
Case "Por Cliente"
    cltCliente.Visible = True
      lblCliente.Visible = True
Case "Por Remito"
    lblFiltro.Caption = "Remito"
    lblFiltro.Visible = True
    txtFiltro.Visible = True
    cmdPegar.Visible = True
Case "Fecha"
    lblFiltro.Caption = "Fecha"
    lblFiltro.Visible = True
    txtFiltro.Visible = True
    
Case "Requerimiento"
    lblFiltro.Caption = "Requerimiento"
    lblFiltro.Visible = True
    txtFiltro.Visible = True
    cmdPegar.Visible = True
Case "Por cliente y Solicitante"
    ctlClienteUsuario.Visible = True
    cltCliente.Visible = True
    lblUsuarioCliente.Visible = True
        lblCliente.Visible = True
Case "Requerimiento para Buscar"
Case "Hacer Remito"
Case "Por Cliente y Descripcion"
    lblCliente.Visible = True
    cltCliente.Visible = True
    lblFiltro.Caption = "Descrip"
    lblFiltro.Visible = True
    txtFiltro.Visible = True
Case "Por Cliente y elemento"
    cltCliente.Visible = True
    lblFiltro.Caption = "Caja/Legajo"
    lblFiltro.Visible = True
    txtFiltro.Visible = True
    lblCliente.Visible = True

Case "Deposito"
    cboDeposito.Visible = True
Case "Tipo Requerimiento"
ctlTipoRequerimiento.Visible = True
Case "Personal Responsable"
ctlPersonal.Visible = True
Case "Hoja de Ruta"
    lblFiltro.Caption = "Hoja de Ruta"
    lblFiltro.Visible = True
    txtFiltro.Visible = True
    cmdPegar.Visible = True


End Select


End Sub

Private Sub chkHojaRuta_Click()
    trvEstado.CheckBoxes = chkHojaRuta.Value
End Sub

Private Sub chkRefrescar_Click()
    timRefrescarEstados.Enabled = chkRefrescar.Value
End Sub



Private Sub chkUltimosdias_Click()

End Sub

Private Sub cltCliente_Click()
If Not IsNull(cltCliente.Valor) Then

Rem ctlClienteUsuario.LlenarConCliente cltCliente.Valor
End If
End Sub

Private Sub cltCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Not IsNull(ctlClienteUsuario.Valor) Then
        ctlClienteUsuario.LlenarConCliente cltCliente.Valor
    End If
End If
End Sub

Private Sub cmdPegar_Click()
Dim A As String
A = Trim(Clipboard.GetText)

A = Replace(A, vbCrLf, ",")
A = Replace(A, " ", "")
A = Mid(A, 1, Len(A) - 1)
txtFiltro.Text = A
End Sub

Private Sub Command1_Click()
    MousePointer = 11
txtPersonal.Text = ""
    Rem ExecutarSql "UPDATE    dbo.REQUERIMIENTO SET FK_SUCURSAL = 'MENDOZA' WHERE (FK_SUCURSAL IS NULL)"
    
    CargarTree
    txtDias.Text = 5
    MousePointer = 0
End Sub

Private Sub Command2_Click()
'    Dim Sql As String
'        Sql = " Update Requerimiento"
'        Sql = Sql & " SET FECHAENTREGA = CONVERT(char, FECHARECEPCION, 103), COMPROMISO_ENTREGA = 'TARDE'"
'        Sql = Sql & " Where (COMPROMISO_ENTREGA Is Null)"
'ExecutarSql Sql
CopiarDatosGrilla DataGrid1

End Sub

Private Sub ctlClienteUsuario_GotFocus()
If Not IsNull(cltCliente.Valor) Then

ctlClienteUsuario.LlenarConCliente cltCliente.Valor
End If
End Sub

Private Sub Form_Activate()
frmControlEstados.WindowState = 2
End Sub

Private Sub Form_Load()
cltCliente.TipoControl = Cliente
ctlClienteUsuario.Valor = Null
ctlTipoRequerimiento.TipoControl = Tipo_Requerimiento
ctlTipoRequerimiento.Visible = False
    ctlClienteUsuario.Visible = False
    cltCliente.Valor = Null
    cltCliente.Visible = False
    txtFiltro.Visible = False
    lblUsuarioCliente.Visible = False
    lblFiltro.Caption = ""
    txtFiltro = ""
    cboDeposito.Visible = False
    cmdPegar.Visible = False
    ctlPersonal.TipoControl = PERSONAL
    ctlPersonal.Visible = False
    
        lblCliente.Visible = False
        cboSucursal.Text = Sucursal
        If MDIfrmInicio.StaInicio.Panels(2).Text = 17 Or MDIfrmInicio.StaInicio.Panels(2).Text = 89 Or MDIfrmInicio.StaInicio.Panels(2).Text = 21 Or MDIfrmInicio.StaInicio.Panels(2).Text = 48 Or MDIfrmInicio.StaInicio.Panels(2).Text = 47 Or MDIfrmInicio.StaInicio.Panels(2).Text = 19 Or MDIfrmInicio.StaInicio.Panels(2).Text = 12 Then
          Rem    chkEncontrados.Enabled = True
        End If
        Set conreque = New ADODB.Connection
        conreque.CursorLocation = adUseClient
        conreque.Open strConBasa
        
End Sub

Private Sub Form_Resize()
On Error GoTo salir:

    If frmControlEstados.Height > 2000 Then
        trvEstado.Width = frmControlEstados.Width - 200
        trvEstado.Height = frmControlEstados.Height - 2700
    End If
salir:
    
End Sub

Private Sub mnuBuscarDocumentosImprimir_Click()
ImprimirRequerimientoGeneral CRequerimientos.Item(1).NumeroRequerimiento
End Sub

Private Sub mnuBusquedaDocumentoAsignacion_Click()
 frmPersonal.Show
End Sub

Private Sub mnuBusquedaDocumentoTerminado_Click()
    Dim rs As New ADODB.Recordset
    Dim sql As String

                CRequerimientos.CambioEstado InputBox("Ingrese el Nº de personal"), True, 5, 6, ConActiva
                 CargarTree
      
End Sub

Private Sub mnuCajasBuscarCajas_Click()
    EstadoFinal = 3
    frmPersonal.Show
End Sub

Private Sub mnuCajasEncontradas_Click()
    CRequerimientos.CambioEstado 0, False, 3, 4, ConActiva
    CargarTree
    
End Sub

Private Sub mnuCajasHojaRuta_Click()
HojasdeRutas
End Sub

Private Sub mnuCajasImprimir_Click()
    ImprimirRequerimientoCajas CRequerimientos.Item(1).NumeroRequerimiento
End Sub
Private Sub mnuCajasReImprimirRemito_Click()
    Dim Remito As Long
        Remito = TRequerimientos.Item(CStr(CRequerimientos.Item(1).NumeroRequerimiento)).IDREMITO
        frmRemitoEntrada.ImprimirRemitoCaja Remito
End Sub
Private Sub mnuCajasRemito_Click()
    On Error GoTo salir:

    frmRemitoEntrada.Show

    frmRemitoEntrada.CargarRequerimientoEnRemito CStr(CRequerimientos.Item(1).NumeroRequerimiento)

    frmRemitoEntrada.SetFocus
        
Exit Sub
salir:
MsgBox Err.Description
End Sub

Private Sub mnuCajasTerminado_Click()
    
    MsgBox "Recuerde Notificar al cliente", vbInformation
    CRequerimientos.CambioEstado InputBox("Ingrese el codigo de personal"), False, 0, 7, ConActiva
    CargarTree
End Sub

Private Sub mnuCajasVaciasCambioEstado_Click()
    frmCajasVacias.Show
'    On Error GoTo salir:
'    Dim RequerimientoVacias As Long
'    RequerimientoVacias = CRequerimientos.Item(1).NumeroRequerimiento
'    Dim rsContenedor As ADODB.Recordset
'    Dim rsCajas  As ADODB.Recordset
'    Dim RegistrosAfectados As Integer
'    Set rsContenedor = New ADODB.Recordset
'    Set rsCajas = New ADODB.Recordset
'    Dim conVacias As New ADODB.Connection
'    Set conVacias = strConBasa , 0 ,1
'    conVacias.CursorLocation = adUseClient
'    conVacias.BeginTrans
'    Dim sql As String
'        sql = " SELECT CAJASLIBROS, ID_CLIENTE,CANTIDAD "
'        sql = sql & " From REQUELIBOSCAJAS, Requerimiento "
'        sql = sql & " Where REQUELIBOSCAJAS.IDREQUERIMIENTOS = Requerimiento.IDRequerimiento"
'        sql = sql & " AND REQUERIMIENTO.IDREQUERIMIENTO = " & RequerimientoVacias
'        sql = sql & " ORDER BY CAJASLIBROS"
'            rsCajas.Open sql, strConBasa , 0 ,1
'
'
'        sql = " SELECT  TOP " & rsCajas!CANTIDAD & "  ESTANTERIA,  VERTICAL , HORIZONTAL,ADELANTE_ATRAS"
'        sql = sql & "  From CONTENEDOR"
'        sql = sql & "  WHERE ESTADO = 1 AND COD_CLIENTE IS NULL AND"
'        sql = sql & " NRO_CAJA IS NULL AND ESTANTERIA BETWEEN 150 AND 190 "
'
'
'
'
'
'        rsContenedor.Open sql, strConBasa , 0 ,1
'    Do While Not rsCajas.EOF
'        sql = " Update CONTENEDOR"
'        sql = sql & vbCrLf & "  SET COD_CLIENTE = " & rsCajas!ID_CLIENTE
'        sql = sql & vbCrLf & " , NRO_CAJA =" & rsCajas!CAJASLIBROS
'        sql = sql & vbCrLf & " , ESTADO = 4 "
'        sql = sql & vbCrLf & " , F_MODIFICACION =" & SysDate
'        sql = sql & vbCrLf & " , IDREQUERIMIENTO =" & RequerimientoVacias
'        sql = sql & vbCrLf & " WHERE ESTANTERIA = " & rsContenedor!ESTANTERIA
'        sql = sql & vbCrLf & " AND VERTICAL = " & rsContenedor!VERTICAL
'        sql = sql & vbCrLf & " AND ADELANTE_ATRAS = " & rsContenedor!ADELANTE_ATRAS
'        sql = sql & vbCrLf & " AND HORIZONTAL = " & rsContenedor!HORIZONTAL
'        conVacias.Execute sql, RegistrosAfectados
'        If RegistrosAfectados <> 1 Then
'            GoTo salir
'        End If
'          rsCajas.MoveNext
'          rsContenedor.MoveNext
'
'    Loop
'        sql = " Update Requerimiento "
'        sql = sql & vbCrLf & "  SET IDESTADO =" & 3
'        sql = sql & vbCrLf & "  Where IDREQUERIMIENTO = " & RequerimientoVacias
'        ExecutarSql sql
'        conVacias.CommitTrans
'        ImprimirRequerimientoGeneral RequerimientoVacias
'
'        CargarTree
'
'     Exit Sub
'salir:
'    conVacias.RollbackTrans
'    MsgBox "Error en la generacion de cajas"
End Sub
Private Sub mnuCajasVaciasImprimir_Click()
  ImprimirRequerimientoVacias CRequerimientos.Item(1).NumeroRequerimiento
End Sub
Private Sub mnuCajasVaciasImprimirRotuloEtiqueta_Click()
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open "Select CANTIDAD from requerimiento WHERE idrequerimiento = " & CRequerimientos.Item(1).NumeroRequerimiento, ConActiva, 0, 1
        If Not rs.EOF Then
           Rem If Rs!CANTIDAD = 36 Or Rs!CANTIDAD = 72 Or Rs!CANTIDAD = 108 Or Rs!CANTIDAD = 144 Or Rs!CANTIDAD = 180 Then
               ImprimirRotuloEtiqueta
            Rem Else
            Rem    MsgBox "La cantidad de cajas vacias debe ser 36-72-108-144-180" & vbCrLf & "Por favor usar la opcion de rotulos estandar", vbInformation
            Rem End If
        End If

End Sub

Private Sub mnuCajasVaciasReImprimirRemito_Click()
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open "Select * from requerimiento WHERE idrequerimiento = " & CRequerimientos.Item(1).NumeroRequerimiento, ConActiva, 0, 1
        If Not rs.EOF Then
             frmRemitoEntrada.ImprimirRemitoCaja rs!IDREMITO
        End If
End Sub

Private Sub mnuCajasVaciasReImprimirRotulos_Click()
    ImprimirRotulos
End Sub

Private Sub mnuCajasVaciasRemito_Click()
    frmRemitoEntrada.Show
    frmRemitoEntrada.CargarRequerimientoEnRemito (CRequerimientos.Item(1).NumeroRequerimiento)
End Sub

Private Sub mnuCantidadElementos_Click()
        Dim sql As String
        Dim Fecha As String
        Dim rs As New ADODB.Recordset
        Dim dato As String
        MousePointer = 11
        Fecha = Mid(FechaRequerimiento, 14)
            sql = " SELECT     REQUERIMIENTO.COMPROMISO_ENTREGA, SUM(REQUERIMIENTO.CANTIDAD) AS Cant"
            sql = sql & " FROM  REQUERIMIENTO INNER JOIN "
            sql = sql & " TIPOREQUERIMIENTO ON REQUERIMIENTO.IDTIPOREQUERIMIENTO = TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO"
            sql = sql & "  WHERE REQUERIMIENTO.FECHAENTREGA = '" & Fecha & "' AND (REQUERIMIENTO.ANULADO IS NULL)"
            sql = sql & "  and      (REQUERIMIENTO.IDTIPOREQUERIMIENTO IN (1, 3, 7, 10, 11, 13))"
            sql = sql & "  GROUP BY REQUERIMIENTO.COMPROMISO_ENTREGA"
            rs.Open sql, strConBasa
            
            Do While Not rs.EOF
                dato = dato & rs!COMPROMISO_ENTREGA & "  " & rs!Cant & vbCrLf
                rs.MoveNext
            Loop
            
            MsgBox dato
         
End Sub

Private Sub mnuConsultas_Digitales_Asignar_Tarea_Click()
frmPersonal.Show
End Sub

Private Sub mnuConsultas_Digitales_Cantidad_Imagenes_Click()
Dim sql As String

Dim CantidadImagenes As String

CantidadImagenes = InputBox("Cantidad de Imagenes")

If IsNumeric(CantidadImagenes) Then
    sql = " UPDATE    REQUERIMIENTO "
    sql = sql & " SET CANTIDAD_IMAGENES = " & CantidadImagenes
    sql = sql & " Where IDREQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
    ExecutarSql sql
    EstadoFinal = 5
    CRequerimientos.CambioEstado 99, False, 4, EstadoFinal, ConActiva
Else


    MsgBox "Error en el ingreso de la Cantidad"
End If
     CargarTree
End Sub

Private Sub mnuConsultas_digitales_Finalizado_Click()
      CRequerimientos.CambioEstado MDIfrmInicio.StaInicio.Panels(2).Text, True, 5, 6, ConActiva
       CargarTree
End Sub

Private Sub mnuConsultas_Digitales_Imagenes_Encontradas_Click()
 CRequerimientos.CambioEstado 99, True, 3, 4, ConActiva
     CargarTree
End Sub

Private Sub mnuConsultas_digitales_Imprimir_Click()
     ImprimirRequerimientoGeneral CRequerimientos.Item(1).NumeroRequerimiento
End Sub

Private Sub mnuControlFaltantes_Click()
    Dim sql As String
    MousePointer = 11
        sql = " SELECT * From V_REQUE_NO_RUTA"
        sql = sql & " WHERE (FECHARECEPCION > TO_DATE('" & FechaRequerimiento & "', 'DD/MM/YYYY'))"
        frmReportes.ImprimirReporte PasoReportes & "rptRequerimientoNoAsisgnadoRuta.rpt", sql, True
        MousePointer = 0
End Sub

Private Sub mnuEstado_Click()
    frmPersonal.Show vbModal
    ImprimirPosicionFechaTipoConsulta TipoConsulta
    CargarTree
End Sub

Private Sub mnuExpanderDia_Click()
Expander (trvEstado.SelectedItem.Index)

End Sub

Private Sub mnuGeneralCambioDeEstado_Click()
    frmPersonal.Show
End Sub

Private Sub mnuGeneralImprimir_Click()
    ImprimirRequerimientoGeneral CRequerimientos.Item(1).NumeroRequerimiento
End Sub

Private Sub mnuGeneralTerminado_Click()
    Dim rs As New ADODB.Recordset
    Dim sql As String

      
            
            CRequerimientos.CambioEstado InputBox("Ingrese el codigo de personal"), True, 5, 6, ConActiva
            CargarTree
      
End Sub


Private Sub HojasdeRutas()
 
 
 Dim rs As New ADODB.Recordset
 Dim sql As String
 Dim des As String
    sql = " SELECT HOJA_RUTA_CUERPO.ID_HOJA_RUTA, HOJA_RUTA_CUERPO.FECHA, HOJA_RUTA_DETALLE.COD_REQUERIMIENTO"
    sql = sql & " FROM   HOJA_RUTA_CUERPO INNER JOIN"
    sql = sql & " HOJA_RUTA_DETALLE ON HOJA_RUTA_CUERPO.ID_HOJA_RUTA = HOJA_RUTA_DETALLE.COD_HOJA_RUTA"
    sql = sql & "  Where HOJA_RUTA_DETALLE.COD_REQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
 
 rs.Open sql, ConActiva, 0, 1
 
 Do While Not rs.EOF
    
    des = des & vbCrLf & "Hoja " & rs!ID_HOJA_RUTA & "  Fecha:" & rs!Fecha
 
    rs.MoveNext
 Loop
 MsgBox des
 
End Sub

Private Sub mnuHojasdeRutas_Click()
HojasdeRutas
End Sub

Private Sub mnuHRVacias_Click()
HojasdeRutas
End Sub

Private Sub mnuImprimir_Click()
    ImprimirPosicionFechaTipoConsulta TipoConsulta
    CargarTree
End Sub

Private Sub ImprimirRotuloEtiqueta()
Dim sql As String
    sql = "  SELECT "
    sql = sql & vbCrLf & " CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, CONTENEDOR.ADELANTE_ATRAS, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA"
    sql = sql & vbCrLf & " From "
    sql = sql & vbCrLf & "CONTENEDOR  "
    sql = sql & vbCrLf & " Where "
    sql = sql & vbCrLf & " CONTENEDOR.COD_CLIENTE = " & TRequerimientos.Item(CStr(CRequerimientos.Item(1).NumeroRequerimiento)).ID_CLIENTE
    sql = sql & vbCrLf & " AND CONTENEDOR.NRO_CAJA in " & Detalle_Requerimiento_Filtro
    frmReportes.ImprimirReporte PasoReportes & "Rotulo_Etiqueta.rpt", sql, True
End Sub


Private Sub mnuLegajosBuscarLegajos_Click()
    ImprimirRequerimientoLegajos CRequerimientos.Item(1).NumeroRequerimiento
    frmPersonal.Show
End Sub

Private Sub mnuLegajosEncontrados_Click()
    CRequerimientos.CambioEstado 0, False, 3, 4, ConActiva
    CargarTree
End Sub

Private Sub mnuLegajosHR_Click()
HojasdeRutas
End Sub

Private Sub mnuLegajosImprimir_Click()
    ImprimirRequerimientoLegajos CRequerimientos.Item(1).NumeroRequerimiento
End Sub

Private Sub mnuLegajosImprimirEtiquetas_Click()
    ImprimirEtiquetasLegajos CRequerimientos.Item(1).NumeroRequerimiento
End Sub

Private Sub mnuLegajosImprimirOrden_Click()
ImprimirRequerimientoLegajosOrden CRequerimientos.Item(1).NumeroRequerimiento
End Sub

Private Sub mnuLegajosReImprimirRemito_Click()
    Dim Remito As Long
    Remito = TRequerimientos.Item(CStr(CRequerimientos.Item(1).NumeroRequerimiento)).IDREMITO
    frmRemitoEntrada.ImprimirRemitoLegajos Remito
End Sub

Private Sub mnuLegajosRemitos_Click()
      frmRemitoEntrada.Show
     frmRemitoEntrada.CargarRequerimientoEnRemito CStr(CRequerimientos.Item(1).NumeroRequerimiento)
     frmRemitoEntrada.SetFocus
End Sub

Private Sub mnuLegajosTerminado_Click()
    MsgBox "Recuerde Notificar al cliente", vbInformation
    CRequerimientos.CambioEstado InputBox("Ingrese el codigo de personal"), False, 0, 7, ConActiva
    CargarTree
End Sub

Private Sub mnuLibrosBuscarLibros_Click()
    frmPersonal.Show
    ImprimirRequerimientoLibros CRequerimientos.Item(1).NumeroRequerimiento
End Sub

Private Sub mnuLibrosHojaderuta_Click()
HojasdeRutas
End Sub

Private Sub mnuLibrosImprimir_Click()
ImprimirRequerimientoLibros CRequerimientos.Item(1).NumeroRequerimiento



End Sub

Private Sub mnuLibrosReImprimirRemitoLibros_Click()
    Dim Remito As Long
    Remito = TRequerimientos.Item(CStr(CRequerimientos.Item(1).NumeroRequerimiento)).IDREMITO
    frmRemitoEntrada.ImprimirRemitoLibros Remito
End Sub

Private Sub Timer2_Timer()
    Refrescar = True
End Sub
Private Sub mnuLibrosRemitoDeLibros_Click()
    frmRemitoEntrada.Show
    frmRemitoEntrada.CargarRequerimientoEnRemito (CRequerimientos.Item(1).NumeroRequerimiento)
    frmRemitoEntrada.SetFocus
End Sub

Private Sub mnuPasarSeleccion_Click()
    Dim i As Integer
    Dim Filtro As String
      MousePointer = 11
      On Error GoTo salir:
        For i = 1 To trvEstado.Nodes.Count
           
            If trvEstado.Nodes.Item(i).Checked = True Then
                Filtro = Filtro & CLng(Mid(trvEstado.Nodes.Item(i).Key, 2)) & ","
            End If
        Next
        MsgBox "paso select"
        If Filtro <> "" Then
            frmHojaRuta.AgregarHojaRuta Mid(Filtro, 1, Len(Filtro) - 1)
            frmHojaRuta.Show
        End If
       MousePointer = 0
       Exit Sub
salir:
       MsgBox Err.Description
       
End Sub

Private Sub mnuPasarTodosRequerimientos_Click()
    Dim i As Integer
    Dim Filtro As String
    Dim rs As New ADODB.Recordset
    Dim sql As String
     MousePointer = 11
        sql = "SELECT IDREQUERIMIENTO From Requerimiento"
        sql = sql & " WHERE " & FechaSolaString("FECHARECEPCION") & "'" & Trim(FechaRequerimiento) & "'"
        sql = sql & "  ORDER BY ID_CLIENTE, IDREQUERIMIENTO"
        rs.Open sql, ConActiva, 0, 1
        Do While Not rs.EOF
            Filtro = Filtro & rs!IDREQUERIMIENTO & ","
            rs.MoveNext
        Loop
        If Filtro <> "" Then
            frmHojaRuta.AgregarHojaRuta Mid(Filtro, 1, Len(Filtro) - 1)
            frmHojaRuta.Show
        End If
        MousePointer = 0
End Sub

Private Sub mnuPlanillas_Click()
Dim Usuario As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim CAJA_1, CAJA_2, CAJA_3, CAJA_4, CAJA_5, CAJA_6, CAJA_7 As String
    Dim DIG_1, DIG_2 As String
    Dim PER_1, PER_2, PER_3 As String
    Dim PERSONAL As String
    Dim CLI_1, CLI_2, CLI_3 As String
  MousePointer = 11
    



sql = " SELECT     REQUERIMIENTO.ID_CLIENTE, REQUELIBOSCAJAS.CAJASLIBROS, REQUERIMIENTO.IDREQUERIMIENTO, CAJAS.DIGITO_VERIFICADOR"
sql = sql & "  FROM         CAJAS INNER JOIN"
sql = sql & "                       REQUELIBOSCAJAS ON CAJAS.NRO_CAJA = REQUELIBOSCAJAS.CAJASLIBROS INNER JOIN"
sql = sql & "                       REQUERIMIENTO ON CAJAS.FK_CLIENTE = REQUERIMIENTO.ID_CLIENTE AND"
sql = sql & "                       REQUELIBOSCAJAS.IDREQUERIMIENTOS = Requerimiento.IDREQUERIMIENTO"
sql = sql & " Where Requerimiento.IDREQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
sql = sql & "  ORDER BY REQUELIBOSCAJAS.CAJASLIBROS"

    
    
    rs.Open sql, ConActiva, 0, 1


ExecutarSql "DELETE FROM IMPRESION_REFERENCIA"
PERSONAL = ""

Do While Not rs.EOF
CAJA_1 = "'" & Mid(rs!CAJASLIBROS, Len(rs!CAJASLIBROS), 1) & "'"
If Len(rs!CAJASLIBROS) > 1 Then
    CAJA_2 = "'" & Mid(rs!CAJASLIBROS, Len(rs!CAJASLIBROS) - 1, 1) & "'"
Else
     CAJA_2 = "'0'"
End If

If Len(rs!CAJASLIBROS) > 2 Then
    CAJA_3 = "'" & Mid(rs!CAJASLIBROS, Len(rs!CAJASLIBROS) - 2, 1) & "'"
Else
    CAJA_3 = "'0'"
End If


If Len(rs!CAJASLIBROS) > 3 Then
    CAJA_4 = "'" & Mid(rs!CAJASLIBROS, Len(rs!CAJASLIBROS) - 3, 1) & "'"
Else
     CAJA_4 = "'0'"
End If


If Len(rs!CAJASLIBROS) > 4 Then
    CAJA_5 = Mid(rs!CAJASLIBROS, Len(rs!CAJASLIBROS) - 4, 1)
Else
     CAJA_5 = "'0'"
End If


If Len(rs!CAJASLIBROS) > 5 Then
    CAJA_6 = Mid(rs!CAJASLIBROS, Len(rs!CAJASLIBROS) - 5, 1)
    Else
    CAJA_6 = "'0'"
End If
If Len(rs!CAJASLIBROS) > 6 Then

    CAJA_7 = Mid(rs!CAJASLIBROS, Len(rs!CAJASLIBROS) - 6, 1)
Else
    CAJA_7 = "'0'"
End If

DIG_1 = "'" & Mid(rs!Digito_Verificador, Len(rs!Digito_Verificador), 1) & "'"
DIG_2 = "'" & Mid(rs!Digito_Verificador, Len(rs!Digito_Verificador) - 1, 1) & "'"



PER_1 = "NULL"

 PER_2 = "NULL"
    PER_3 = "NULL"




CLI_1 = "'" & Mid(rs!ID_CLIENTE, Len(rs!ID_CLIENTE), 1) & "'"

If Len(rs!ID_CLIENTE) > 1 Then
    CLI_2 = "'" & Mid(rs!ID_CLIENTE, Len(rs!ID_CLIENTE) - 1, 1) & "'"
Else
    CLI_2 = "'0'"
End If

If Len(rs!ID_CLIENTE) > 2 Then
    CLI_3 = "'" & Mid(rs!ID_CLIENTE, Len(rs!ID_CLIENTE) - 2, 1) & "'"
Else
    CLI_3 = "'0'"
End If




sql = " INSERT INTO IMPRESION_REFERENCIA"
sql = sql & " (CAJA_1, CAJA_2, CAJA_3, CAJA_4, CAJA_5, CAJA_6, CAJA_7, DIG_1, DIG_2, PER_1, PER_2, PER_3, CLI_1, CLI_2, CLI_3, ORDEN)"
sql = sql & "  VALUES  "
sql = sql & "(" & CAJA_1 & "," & CAJA_2 & "," & CAJA_3 & "," & CAJA_4 & "," & CAJA_5 & "," & CAJA_6 & "," & CAJA_7 & "," & DIG_1 & "," & DIG_2 & "," & PER_1 & "," & PER_2 & "," & PER_3 & "," & CLI_1 & "," & CLI_2 & "," & CLI_3 & "," & rs!CAJASLIBROS & ")"
ExecutarSql sql
rs.MoveNext
Loop



  sql = " SELECT * "
 sql = sql & "  FROM  IMPRESION_REFERENCIA"
 sql = sql & " order by ORDEN desc "




frmReportes.ImprimirReporte PasoReportes & "ReferenciaManual.rpt", sql, True
MousePointer = 0

End Sub

Private Sub mnuReferenciaTrasvaseAsignacion_Click()
    frmPersonal.Show
    EstadoFinal = 5
End Sub

Private Sub mnuReferenciaTrasvaseDiccCompleto_Click()
   ImprimirIndiceRequerimiento CRequerimientos.Item(1).NumeroRequerimiento
End Sub

Private Sub mnuReferenciaTrasvaseImprimir_Click()
 ImprimirIndiceRequerimiento CRequerimientos.Item(1).NumeroRequerimiento
End Sub

Private Sub mnuReferenciaTrasvaseTerminado_Click()
Dim rs As New ADODB.Recordset
Dim sql As String


sql = " SELECT  IDPERSONAL From Requerimiento"
sql = sql & " Where IDREQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
rs.Open sql, ConActiva, 0, 1

If Not rs.EOF Then
   CRequerimientos.CambioEstado rs!IDPERSONAL, True, 5, 6, ConActiva
Else
MsgBox "Error "
End If
CargarTree

End Sub

Private Sub mnuTerminado_Click()
 
End Sub

Private Sub mnuTipoLegajosEtiquetas_Click()
Dim datos As String
datos = RequerimientoSelecion

 If datos <> "" Then
    ImprimirEtiquetasLegajos datos
 Else
 MsgBox "No existene elementos selecionador"
 End If
End Sub

Private Sub mnuTipoRequerimientoCajasImprimir_Click()
       
    MousePointer = 11
    Dim Seleccion As String
    Seleccion = RequerimientoSelecion
    If Seleccion <> "" Then
        ImprimirRequerimientoCajasTodos Seleccion
    Else
        MsgBox "No Exsiste seleccion"
    End If
    
    MousePointer = 0
End Sub
    
Private Sub mnuTipoRequerimientoCajasImprimirTodo_Click()
Dim dato As String
dato = RequerimientoSelecion
Dim sql As String
If dato <> "" Then
If MsgBox("Cambio de estado", vbYesNo + vbQuestion) = vbYes Then
sql = sql & " Update dbo.Requerimiento"
sql = sql & "  SET              IDESTADO =3, "
sql = sql & " IDPERSONAL = " & InputBox("Ingrese el resposable", , 99)
sql = sql & "  WHERE     (IDREQUERIMIENTO IN (" & dato
sql = sql & "  )) AND (IDESTADO = 2)"
ExecutarSql sql
End If

ImprimirRequerimientoCajasTodo dato
 MousePointer = 11
    CargarTree
    MousePointer = 0

Else
    MsgBox "FALTA SELECION"
End If

End Sub

Private Sub mnuTipoRequerimientoCajasVaciasEtiquetas_Click()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim Fecha As String
    Dim Filtro As String
        
      Dim FiltroRequerimiento As String
     Dim i As Integer
 
      MousePointer = 11
   
       
    If RequerimientoSelecion <> "" Then
    
    

    sql = " SELECT REQUELIBOSCAJAS.CAJASLIBROS, REQUERIMIENTO.ID_CLIENTE, Requerimiento.IDREQUERIMIENTO"
    sql = sql & "  From Requerimiento, REQUELIBOSCAJAS"
    sql = sql & "  Where Requerimiento.IDREQUERIMIENTO = REQUELIBOSCAJAS.IDREQUERIMIENTOS"
    sql = sql & " AND Requerimiento.IDREQUERIMIENTO in( " & RequerimientoSelecion & ")"
    sql = sql & " order by ID_CLIENTE,CAJASLIBROS "
    
    rs.Open sql, ConActiva, 0, 1
     Do While Not rs.EOF
        Filtro = Filtro & vbCrLf & "(NRO_CAJA = " & rs!CAJASLIBROS & " And COD_CLIENTE = " & rs!ID_CLIENTE & ") Or "
        rs.MoveNext
    Loop
    
    
    If Filtro <> "" Then
                Filtro = Mid(Filtro, 1, Len(Filtro) - 3)
                sql = "  SELECT "
                sql = sql & vbCrLf & " CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, CONTENEDOR.ADELANTE_ATRAS, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA"
                sql = sql & vbCrLf & " From "
                sql = sql & vbCrLf & "  CONTENEDOR "
                sql = sql & vbCrLf & " Where "
                sql = sql & vbCrLf & Filtro
                sql = sql & vbCrLf & " Order by COD_CLIENTE, NRO_CAJA"
                frmReportes.ImprimirReporte PasoReportes & "Rotulo_Etiqueta.rpt", sql, True
        End If
        
        
        End If
       MousePointer = 0
    
    
    
   
    
    
End Sub

Private Sub mnuTipoRequerimientoLegajosImprimir_Click()
 Dim datos As String
 datos = RequerimientoSelecion
 If datos <> "" Then
  Dim sql As String
    ImprimirRequerimientoLegajosTodos datos
        If MsgBox("Cambio de estado", vbYesNo + vbQuestion) = vbYes Then
            sql = sql & " Update dbo.Requerimiento"
            sql = sql & "  SET              IDESTADO =3, "
            sql = sql & " IDPERSONAL = " & InputBox("Ingrese el resposable", , 99)
            sql = sql & "  WHERE     (IDREQUERIMIENTO IN (" & datos
            sql = sql & "  )) AND (IDESTADO = 2)"
            ExecutarSql sql
        End If
    
    
    
    Else
    MsgBox "No existe selecion"
    
 End If
 
End Sub

Private Sub timRefrescarEstados_Timer()
    If Refrescar And CBool(chkRefrescar.Value) Then
        CargarTree
    End If
End Sub

Private Sub trvEstado_Click()

If Mid(trvEstado.SelectedItem.Tag, 1, 1) = "F" Then
   
   If chkHojaRuta.Value = 1 Then
           
    Expander (trvEstado.SelectedItem.Index)
   Else
        If trvEstado.SelectedItem.Expanded = True Then
            trvEstado.SelectedItem.Expanded = False
        Else
            Expander (trvEstado.SelectedItem.Index)
        End If
    End If
End If
If Mid(trvEstado.SelectedItem.Tag, 1, 1) = "S" Then
    If trvEstado.SelectedItem.Expanded = True Then
      trvEstado.SelectedItem.Expanded = False
    Else
        trvEstado.SelectedItem.Expanded = True
    End If
    
End If

If Mid(trvEstado.SelectedItem.Tag, 1, 1) = "M" Then
    If trvEstado.SelectedItem.Expanded = True Then
      trvEstado.SelectedItem.Expanded = False
    Else
        Expander (trvEstado.SelectedItem.Index)
    End If
    
    
End If
 If Mid(trvEstado.SelectedItem.Tag, 1, 1) = "T" Then
        If trvEstado.SelectedItem.Expanded = True Then
          trvEstado.SelectedItem.Expanded = True
        Else
            Expander (trvEstado.SelectedItem.Index)
        End If
    End If


 nodoSelecionado = ""
End Sub

Private Sub trvEstado_DblClick()
    If Mid(trvEstado.SelectedItem.Key, 1, 1) = "R" Then
        rsRequerimientos.Filter = "IDREQUERIMIENTO = " & CLng(Mid(trvEstado.SelectedItem.Key, 2))
        Load frmRequerimientoConsulta
        frmRequerimientoConsulta.CargarRequerimiento CLng(Mid(trvEstado.SelectedItem.Key, 2))
        frmRequerimientoConsulta.Show
    End If
End Sub

Private Sub trvEstado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim Fecha As Date
    Dim NRequerimiento As Long
    Dim ESTADO As Integer
    Dim TIPO As Integer
    Dim IndexInicio As Integer
    Dim IndexFin As Integer

On Error GoTo LUIS
 Rem Exit Sub
 
   If Button = 1 Then
'     If Mid(trvEstado.SelectedItem.Tag, 1, 1) = "T" Then
'        If trvEstado.SelectedItem.Expanded = True Then
'          trvEstado.SelectedItem.Expanded = False
'        Else
'            Expander (trvEstado.SelectedItem.Index)
'        End If
'    End If
 End If

        If Button = 2 Then
            Pedido = 0
             nodoSelecionado = ""
            For i = 1 To trvEstado.Nodes.Count
              If trvEstado.Nodes(i).Selected Then
                   
                    Select Case Mid(trvEstado.Nodes(i).Tag, 1, 1)
                    Case "F" ' Por Fecha
                        FechaRequerimiento = Mid(trvEstado.Nodes(i).Key, 2)
                        PopupMenu mnuHojaRuta
                     Case "M"
'                         TIPO = Mid(trvEstado.Nodes(i).Key, 15)
'                         Fecha = Mid(trvEstado.Nodes(i).Key, 2, 11)
'                        Select Case TIPO
'                        Case 1, 3
'                            PopupMenu mnuTipoRequerimientoCajas
'                        Case 7
'                            Menu_Tipo_Cajas_Vacias (Fecha)
'                         Case 10, 11
'                         PopupMenu mnuTipoRequerimientosLegajos
'                        End Select
                    
                    Case "T" 'seleciono por tipo de requerimiento
                         TIPO = Mid(trvEstado.Nodes(i).Tag, 2)
                         Fecha = trvEstado.Nodes(i).Parent.Parent.Text
                        Select Case TIPO
                        Case 1, 3
                            PopupMenu mnuTipoRequerimientoCajas
                        Case 7
                            Menu_Tipo_Cajas_Vacias (Fecha)
                         Case 10, 11
                         PopupMenu mnuTipoRequerimientosLegajos
                        End Select
                    Case "R"
                     nodoSelecionado = trvEstado.Nodes(i).Key
                        CRequerimientos.Clear
                        Fecha = CDate(Mid(trvEstado.Nodes(i).Parent.Parent.Parent.Tag, 3, 10))
                        NRequerimiento = Mid(trvEstado.Nodes(i).Tag, 2)
                        TIPO = CInt(Mid(trvEstado.Nodes(i).Parent.Tag, 2))
                        ESTADO = TRequerimientos.Item(CStr(NRequerimiento)).IDESTADO
                        CRequerimientos.Add Fecha, ESTADO, NRequerimiento, TIPO, CStr(NRequerimiento)
                        Select Case TIPO
                        Case 1, 3 'Cajas
                            Menu_Cajas ESTADO
                        Case 2, 4 'Libros
                            Menu_Libros ESTADO
                        Case 5 'PERIDO DE REFERENCIA Y TRASVASE
                            Menu_Referencia_Trasvase ESTADO
                        Case 6 'RETIRO DE CAJAS
                            Menu_General ESTADO
                        Case 7, 26  'PEDIDOS DE CAJAS VACIAS
                            Menu_Cajas_Vacias ESTADO
                        Case 8 'BUSQUEDA DE DOCUMENTACION
                            Menu_BusquedaDocumento ESTADO
                        Case 9  'Consulta en planta
                            Menu_Consulta_Planta ESTADO
                        Case 27  'Consulta INTERNA
                            Menu_Consulta_Planta ESTADO
                        Case 10, 11 'Legajos
                            Menu_Legajos ESTADO
                        Case 13
                            Menu_Consultas_Digitales ESTADO
                            Case 25
                            Menu_Consulta_Cupones ESTADO
                        Case Else
                            Menu_General ESTADO
                        End Select
                    End Select
                End If
            Next
        End If

LUIS:

End Sub


Public Sub ImprimirPosicionFechaTipoConsulta(dato As String)
Dim rs As ADODB.Recordset
Dim sql As String
Dim SQL1 As String
Dim FECHARECEPCION As Date
Dim IDTIPOREQUERIMIENTO As Integer
Dim Filtro As String
Dim ANTERIOR As Long
Dim Bandera As Boolean
Dim Responsables As String
Dim i As Integer
Bandera = False
If CRequerimientos.Count = 0 Then
     Exit Sub
End If

For i = 1 To CRequerimientos.Count
    If i = 1 Then
        Filtro = "(" & CRequerimientos.Item(i).NumeroRequerimiento
    Else
        Filtro = Filtro & "," & CRequerimientos.Item(i).NumeroRequerimiento
    End If
Next

Filtro = Filtro & ")"
    
    'REPORTE
    sql = " SELECT * "
    sql = sql & " From V_TIPO_REQUERIMIENTO "
    sql = sql & " Where IDREQUERIMIENTO IN (" & Filtro & ")"
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoTipo.rpt", sql, False
    CRequerimientos.Clear

End Sub

Public Sub CargarTreenAÑANAno()
        Dim rsRequerimientos As ADODB.Recordset
        Dim ConsFecha As Date
        Dim Tiporequerimiento As Integer
        Dim nodX As Node  ' Create variable.
        Dim ImagenX As Integer
        Dim sql As String
        Dim DESCRIPCION As String
        Dim Requerimiento As String
        Dim i As Long
        Dim Procimo As String
        Dim FechaInferior As Date
        Dim IDREQUERIMIENTO As Long
        Dim ID_CLIENTE As Integer
        Dim IDPERSONAL2 As Integer
        Dim IDTIPORECEPCION As Integer
        Dim IDESTADO As Integer
        Dim IDTIPOREQUERIMIENTO As Integer
        Dim IDFAX As Long
        Dim Sector As String
        Dim TELEFONO As String
        Dim DESCRIPCION2 As String
        Dim SOLICITANTE As String
        Dim TOMO As Integer
        Dim FECHAENTREGA As Date
        Dim FECHALIMITE As Date
        Dim FECHARECEPCION As Date
        Dim CANTIDAD As String
        Dim IDREMITO As Long
        Dim TIEMPOTOTAL As String
        Dim IDREQUERIMIENTOANT     As Long
        Dim PEDIDOCLIENTE As String
        Dim IconoRequerimiento As String





    sql = " SELECT * "
    sql = sql & vbCrLf & " From V_REQUERIMIENTO "
    sql = sql & vbCrLf & " WHERE ANULADO IS NULL "
    
  ' ----------- Filtros -----------------------------
    
            If cboFiltro.Text = "Fecha" Then
                If txtFiltro <> "" Then
                    sql = sql & vbCrLf & " AND  " & FechaSolaString("FECHARECEPCION") & " = '" & txtFiltro.Text & "'"
                End If
            Else
'                If CBool(chkUltimosdias.Value) Then
'                    FechaInferior = Format(DateAdd("D", -8, SysDateCompare), "DD/MM/YYYY")
'                    sql = sql & vbCrLf & " AND  FECHARECEPCION > '" & FechaInferior & "'"
'                End If
            End If
            If CBool(chkPendientes) Then
                sql = sql & vbCrLf & " AND IDEstado < 5 "
            End If
            If CBool(chkRemitosPendientes) Then
                sql = sql & vbCrLf & " AND IDEstado > 5 "
                sql = sql & vbCrLf & " AND TIEMPOTOTAL IS NULL  "
                sql = sql & vbCrLf & " AND IDTIPOREQUERIMIENTO IN (1,2,3,4,7,10,11)  "
            End If
            If cboFiltro.Text = "Por Cliente" Then
                If txtFiltro <> "" Then
                    sql = sql & vbCrLf & " AND ID_Cliente= " & txtFiltro
                End If
            End If
            If cboFiltro.Text = "Por Remito" Then
                If txtFiltro <> "" Then
                    sql = sql & vbCrLf & " AND idremito = " & txtFiltro
                End If
            End If
            
            If cboFiltro.Text = "Requerimiento" Then
                If txtFiltro <> "" Then
                    sql = sql & vbCrLf & " AND IDREQUERIMIENTO = " & txtFiltro
                End If
            End If
            
            If Not IsNull(ctlClienteUsuario.Valor) Then
                sql = sql & vbCrLf & " AND  COD_USUARIO_CLIENTE  =" & ctlClienteUsuario.Valor
                ctlClienteUsuario.Valor = Null
                txtFiltro.Text = ""
            End If
             sql = sql & vbCrLf & " order by CONVERT(char(8), FECHARECEPCION, 112) ,IDTIPOREQUERIMIENTO"
    
    Set rsRequerimientos = New ADODB.Recordset
    rsRequerimientos.Open sql, ConActiva, 0, 1
    trvEstado.Nodes.Clear
    trvEstado.Style = 5
    TRequerimientos.Clear
    Do While Not rsRequerimientos.EOF
        With rsRequerimientos
            IDREQUERIMIENTO = !IDREQUERIMIENTO
            ID_CLIENTE = !ID_CLIENTE
            IDPERSONAL2 = !IDPERSONAL
            IDTIPORECEPCION = !IDTIPORECEPCION
            IDESTADO = !IDESTADO
            IDTIPOREQUERIMIENTO = !IDTIPOREQUERIMIENTO
            IDFAX = !IDFAX
            If IsNull(!PEDIDOCLIENTE) Then
                PEDIDOCLIENTE = ""
            Else
                PEDIDOCLIENTE = !PEDIDOCLIENTE
            End If
            If IsNull(!Sector) Then
                Sector = ""
            Else
                Sector = !Sector
            End If
            If IsNull(!TELEFONO) Then
                TELEFONO = ""
            Else
                TELEFONO = !TELEFONO
            End If
            If IsNull(!DESCRIPCION) Then
                DESCRIPCION2 = ""
            Else
                DESCRIPCION2 = !DESCRIPCION
            End If
            If IsNull(!SOLICITANTE) Then
                SOLICITANTE = ""
            Else
                SOLICITANTE = !SOLICITANTE
            End If
            If IsNull(!TOMO) Then
                TOMO = 0
            Else
                TOMO = !TOMO
            End If
            If IsNull(!FECHAENTREGA) Then
                FECHAENTREGA = 0
            Else
                FECHAENTREGA = !FECHAENTREGA
            End If
            If IsNull(!FECHALIMITE) Then
                FECHALIMITE = 0
            Else
                FECHALIMITE = !FECHALIMITE
            End If
            FECHARECEPCION = !FECHARECEPCION
            CANTIDAD = !CANTIDAD
            If IsNull(!IDREMITO) Then
                IDREMITO = 0
            Else
                IDREMITO = !IDREMITO
            End If
            If IsNull(!TIEMPOTOTAL) Then
                TIEMPOTOTAL = ""
            Else
                TIEMPOTOTAL = !TIEMPOTOTAL
            End If
            If IDREQUERIMIENTOANT <> IDREQUERIMIENTO Then
                IDREQUERIMIENTOANT = IDREQUERIMIENTO
                TRequerimientos.Add IDREQUERIMIENTO, CLng(ID_CLIENTE), IDPERSONAL2, _
                IDTIPORECEPCION, IDESTADO, IDTIPOREQUERIMIENTO, IDFAX, Sector, _
                TELEFONO, DESCRIPCION2, SOLICITANTE, TOMO, FECHAENTREGA, FECHALIMITE, _
                FECHARECEPCION, CStr(CANTIDAD), IDREMITO, TIEMPOTOTAL, CStr(!IDREQUERIMIENTO)
            End If
            If IsNull(rsRequerimientos!TIPO_DESCRIPCION) Then
                DESCRIPCION = "0"
            Else
                DESCRIPCION = UCase(CStr(rsRequerimientos!TIPO_DESCRIPCION))
            End If
    End With
    
    
    Dim RazonSoc As String
    Dim RequeDes As String
        RazonSoc = Mid(Trim(UCase(CStr(rsRequerimientos!Razon_Social))), 1, 30)
        RequeDes = Format(CStr(rsRequerimientos!IDREQUERIMIENTO), "0000000") & "  //  " & Mid(Trim(UCase(CStr(rsRequerimientos!Razon_Social))), 1, 30) & " // Cant.:" & CStr(IIf(IsNull(rsRequerimientos!CANTIDAD), "0", rsRequerimientos!CANTIDAD)) & " // Hora Recep.:" & Trim(Format(CStr(rsRequerimientos!FECHARECEPCION), "DD/MM/YY hh:mm")) & "  //  Resp:" & Trim(UCase(rsRequerimientos!Apellido))
    
        Select Case rsRequerimientos!IDTIPOREQUERIMIENTO
        
     
        Case 13, 14
            Requerimiento = RequeDes & "  // Cantidad de Imagenes:" & rsRequerimientos!CANTIDAD_IMAGENES
        
        Case 5, 6, 8, 9, 12, 15, 16, 17, 18, 19
                If IsNull(rsRequerimientos!COD_HOJA_RUTA_TERMINADO) Then
                     Requerimiento = RequeDes
                Else
                
                     Requerimiento = RequeDes & "   // Hoja Ruta :" & rsRequerimientos!COD_HOJA_RUTA_TERMINADO
                End If
                      
        Case Else
                
                If CInt(rsRequerimientos!IDESTADO) < 5 Then
                     Requerimiento = RequeDes
                Else
                    If Not IsNull(rsRequerimientos!COD_HOJA_RUTA_TERMINADO) Then
                        Requerimiento = RequeDes & "  // Remito:" & Trim(CStr(rsRequerimientos!IDREMITO)) & "  // Hoja Ruta :" & rsRequerimientos!COD_HOJA_RUTA_TERMINADO
                     
                     Else
                        Requerimiento = RequeDes & "  // Remito:" & Trim(CStr(rsRequerimientos!IDREMITO)) & "  // Hoja Ruta : "
                     End If
                     
                End If
                
        End Select
        
        
        If CStr(rsRequerimientos!IDESTADO) = 2 Then
            If rsRequerimientos!EnvioTarde = True Then
                IconoRequerimiento = "E" & CStr(rsRequerimientos!IDESTADO) & "1"
            Else
                IconoRequerimiento = "E" & CStr(rsRequerimientos!IDESTADO)
            End If
        Else
            IconoRequerimiento = "E" & CStr(rsRequerimientos!IDESTADO)
        End If
        
        
        If ConsFecha = Format(CStr(rsRequerimientos!FECHARECEPCION), "DD/MM/YYYY") Then
            If Tiporequerimiento = CInt(rsRequerimientos!IDTIPOREQUERIMIENTO) Then
            
                 Set nodX = trvEstado.Nodes.Add("T " & ConsFecha & "  " & Format(CStr(Tiporequerimiento), "00"), tvwChild, "R " & Format(CStr(rsRequerimientos!IDREQUERIMIENTO), "0000000"), Requerimiento, IconoRequerimiento)
                 ColocarColorRequ "R " & Format(CStr(rsRequerimientos!IDREQUERIMIENTO), "0000000"), CInt(rsRequerimientos!IDESTADO), rsRequerimientos!IDTIPOREQUERIMIENTO, rsRequerimientos!FECHARECEPCION, trvEstado
                 Rem trvEstado.Nodes.Item("T " & ConsFecha & "  " & Format(CStr(Tiporequerimiento), "00"), tvwChild, "R " & Format(CStr(rsRequerimientos!IDREQUERIMIENTO), "0000000").Tag = Trim(rsRequerimientos!DESCRIPCION)
                If Not IsNull(rsRequerimientos!DESCRIPCION) Then
                 trvEstado.Nodes.Item(nodX.Key).ForeColor = &H8000&
                 trvEstado.Nodes.Item(nodX.Key).Tag = Trim(rsRequerimientos!DESCRIPCION)
                 End If
            Else
                Tiporequerimiento = CInt(rsRequerimientos!IDTIPOREQUERIMIENTO)
                Set nodX = trvEstado.Nodes.Add("F " & ConsFecha, tvwChild, "T " & ConsFecha & "  " & Format(CStr(Tiporequerimiento), "00"), DESCRIPCION, "E7")
                Set nodX = trvEstado.Nodes.Add("T " & ConsFecha & "  " & Format(CStr(Tiporequerimiento), "00"), tvwChild, "R " & Format(CStr(rsRequerimientos!IDREQUERIMIENTO), "0000000"), Requerimiento, "E" & CStr(rsRequerimientos!IDESTADO))
                ColocarColorRequ "R " & Format(CStr(rsRequerimientos!IDREQUERIMIENTO), "0000000"), CInt(rsRequerimientos!IDESTADO), rsRequerimientos!IDTIPOREQUERIMIENTO, rsRequerimientos!FECHARECEPCION, trvEstado
                If Not IsNull(rsRequerimientos!DESCRIPCION) Then
                 trvEstado.Nodes.Item(nodX.Key).ForeColor = &H8000&
                 trvEstado.Nodes.Item(nodX.Key).Tag = Trim(rsRequerimientos!DESCRIPCION)
                 End If
                
            End If
            rsRequerimientos.MoveNext
        Else
            ConsFecha = Format(CStr(rsRequerimientos!FECHARECEPCION), "DD/MM/YYYY")
            Tiporequerimiento = 0
            Set nodX = trvEstado.Nodes.Add(, , "F " & ConsFecha, ConsFecha, "SVerde")
        
        End If
    Loop
    
    On Error GoTo salir
    Expander (trvEstado.Nodes.Item("F " & Format(SysDateCompare, "DD/MM/YYYY")).Index)
salir:



End Sub




Public Sub CargarTree()

        Dim ConsFecha As Date
        Dim Tiporequerimiento As String
        Dim nodX As Node  ' Create variable.
        Dim ImagenX As Integer
        Dim sql As String
        Dim DESCRIPCION As String
        Dim Requerimiento As String
        Dim i As Long
        Dim Procimo As String
        Dim FechaInferior As Date
        Dim IDREQUERIMIENTO As Long
        Dim ID_CLIENTE As Integer
        Dim IDPERSONAL2 As Integer
        Dim IDTIPORECEPCION As Integer
        Dim IDESTADO As Integer
        Dim IDTIPOREQUERIMIENTO As Integer
        Dim IDFAX As Long
        Dim Sector As String
        Dim TELEFONO As String
        Dim DESCRIPCION2 As String
        Dim SOLICITANTE As String
        Dim TOMO As Integer
        Dim FECHAENTREGA As Date
        Dim FECHALIMITE As Date
        Dim FECHARECEPCION As Date
        Dim FECHA_SISTEMA As String
        Dim CANTIDAD As String
        Dim IDREMITO As Long
        Dim TIEMPOTOTAL As String
        Dim IDREQUERIMIENTOANT     As Long
        Dim PEDIDOCLIENTE As String
        Dim IconoRequerimiento As String
        Dim constMañanaTarde As String
        Dim KeyPadreFecha As String
        Dim KeyPadreMañana As String
        Dim KeyPadreTipo As String
        Dim KeyRequerimiento As String
        Dim RsElementos As ADODB.Recordset
        Dim PEDIDO_APELLIDO_NOMBRE As String
        Dim sqlTitulo As String
        Dim SQLP As String
        Dim rsPersonal As New ADODB.Recordset
        Dim ID_Personal_Responsable As String
        
        C_Fecha = ""
        C_MañanaTarde = ""
        C_Sucursal = ""
        C_TipoRequerimiento = ""
        
      On Error GoTo salir:

    sqlTitulo = " SELECT   * "
    sqlTitulo = sqlTitulo & vbCrLf & " From V_REQUERIMIENTO "
    sqlTitulo = sqlTitulo & vbCrLf & " WHERE "
    
    If chkVerAnulados.Value = 1 Then
    sqlTitulo = sqlTitulo & vbCrLf & "  NOT ANULADO IS NULL "
    Else
        
      sqlTitulo = sqlTitulo & vbCrLf & " ANULADO IS NULL "
     End If
     
     
    
  ' ----------- Filtros -----------------------------
    
           
           If cboFiltro.Text = "Personal Responsable" Then
                
                If Not IsNull(ctlPersonal.Valor) Then
                
                    SQLP = "SELECT  TOP 300    REQUERIMIENTO.IDREQUERIMIENTO, PERSONAL.NOMBRE, PERSONAL.APELLIDO, PERSONAL.IDPERSONAL"
                    SQLP = SQLP & vbCrLf & " FROM         REQUERIMIENTO INNER JOIN"
                    SQLP = SQLP & vbCrLf & "   H_ESTADO_REQUE ON REQUERIMIENTO.IDREQUERIMIENTO = H_ESTADO_REQUE.IDREQUERIMIENTO INNER JOIN"
                    SQLP = SQLP & vbCrLf & "   PERSONAL ON H_ESTADO_REQUE.IDPERSONAL = PERSONAL.IDPERSONAL"
                    SQLP = SQLP & vbCrLf & "  Where (H_ESTADO_REQUE.CONTADOR = 1) "
                    SQLP = SQLP & vbCrLf & "  And PERSONAL.IDPERSONAL = " & ctlPersonal.Valor
                    SQLP = SQLP & vbCrLf & "  ORDER BY REQUERIMIENTO.IDREQUERIMIENTO DESC"
                    
                    rsPersonal.Open SQLP, strConBasa
                    
                    Do While Not rsPersonal.EOF
                    
                    ID_Personal_Responsable = ID_Personal_Responsable & "," & rsPersonal!IDREQUERIMIENTO
                    
                      rsPersonal.MoveNext
                      
                    Loop
                    
                     sql = sql & vbCrLf & " AND IDREQUERIMIENTO in( " & Mid(ID_Personal_Responsable, 2) & ")"
                End If
                            
            End If
           
            
            
            
            If cboFiltro.Text = "Fecha" Then
                If txtFiltro <> "" Then
                    sql = sql & vbCrLf & " AND FECHARECEPCION   > " & FechaFormato(txtFiltro.Text)
                End If
            Else
                If txtDias.Text <> 0 Then
                    FechaInferior = Format(DateAdd("D", -txtDias.Text, SysDate_DD_MM_YYYY), "DD/MM/YYYY")
                    sql = sql & vbCrLf & " AND  FECHAENTREGA > " & FechaFormato(FechaInferior)
                End If
            End If
            
'            If chkpendienteDigi.Value = True Then
'                sql = sql & vbCrLf & " AND V_REQUERIMIENTO.IDEstado in(8 )"
'            End If
            
            
            
            If cboFiltro.Text = "Tipo Requerimiento" Then
             If Not IsNull(ctlTipoRequerimiento.Valor) Then
                sql = sql & vbCrLf & " AND IDTIPOREQUERIMIENTO=" & ctlTipoRequerimiento.Valor
             Else
                MsgBox "Ingrese el tipo de requerimiento"
             End If
             
            End If
            
           
            
             If cboSucursal.Text <> "" Then
            sql = sql & vbCrLf & " AND FK_SUCURSAL='" & Trim(cboSucursal.Text) & "'"
            
            End If
            
            If cboFiltro.Text = "Mensajes Alsina" Then
            
            sql = sql & vbCrLf & " AND  DESCRIPCION_ACTUALIZADA = 5 "
            End If
            
            
            If cboFiltro.Text = "Requerimiento para Buscar" Then
                sql = sql & vbCrLf & " AND V_REQUERIMIENTO.IDEstado in(1,2,3)  "
            End If
            
            If cboFiltro.Text = "Hacer Remito" Then
                    sql = sql & vbCrLf & " AND V_REQUERIMIENTO.IDEstado in(4) AND IDTIPOREQUERIMIENTO IN( 1,2,3,4,7,8,10,11,13,14,18) "
            End If
            
            
            If cboFiltro.Text = "Clientes Custodia" Then
                sql = sql & vbCrLf & " AND V_REQUERIMIENTO.ID_Cliente > 999 "
            End If
            
            
          If cboFiltro.Text = "Clientes Custodia planta" Then
            sql = sql & vbCrLf & " AND V_REQUERIMIENTO.ID_Cliente > 999 AND V_REQUERIMIENTO.IDEstado < 4  "
          End If
           
            
            If cboFiltro.Text = "Por Cliente y Descripcion" Then
               If Not IsNull(cltCliente.Valor) And Trim(txtFiltro.Text) <> "" Then
                    sql = sql & vbCrLf & " AND V_REQUERIMIENTO.ID_Cliente= " & cltCliente.Valor & " AND V_REQUERIMIENTO.DESCRIPCION LIKE '%" & Trim(txtFiltro.Text) & "%'"
               Else
                    MsgBox "Error en Filtro"
                    Exit Sub
               End If
            End If

           If cboFiltro.Text = "Deposito" Then




Dim sqlde As String
        
        Dim RsE As New ADODB.Recordset
        Dim FiltroE As String



sqlde = "  SELECT     REQUERIMIENTO.IDREQUERIMIENTO"
sqlde = sqlde & "  FROM         REQUERIMIENTO INNER JOIN"
sqlde = sqlde & " REQUELIBOSCAJAS ON REQUERIMIENTO.IDREQUERIMIENTO = REQUELIBOSCAJAS.IDREQUERIMIENTOS"
sqlde = sqlde & "  WHERE     (REQUELIBOSCAJAS.DEPOSITO = '" & Trim(cboDeposito.Text) & "')"
sqlde = sqlde & "  AND (REQUERIMIENTO.FECHARECEPCION > '" & DateAdd("d", -30, Format(SysDate, "dd/mm/yyyy")) & "')"
sqlde = sqlde & "  GROUP BY REQUERIMIENTO.IDREQUERIMIENTO"



        RsE.Open sqlde, ConActiva, 0, 1

        Do While Not RsE.EOF
            FiltroE = "," & RsE!IDREQUERIMIENTO & FiltroE
            RsE.MoveNext
        Loop

        If FiltroE <> "" Then
            sql = sql & vbCrLf & " AND IDREQUERIMIENTO in( " & Mid(FiltroE, 2) & ")"
            sql = sql & vbCrLf & " AND V_REQUERIMIENTO.IDEstado < 4 "
        Else
             MsgBox "No se encontraron registros"
            Exit Sub
        End If
End If

   
If cboFiltro.Text = "Por Cliente y elemento" Then


    If Not IsNull(cltCliente.Valor) And Trim(txtFiltro.Text) <> "" Then
        Dim SqlElemento As String
        Set RsElementos = New ADODB.Recordset
        Dim FiltroElemento As String
        SqlElemento = " SELECT     dbo.REQUERIMIENTO.IDREQUERIMIENTO "
        SqlElemento = SqlElemento & " FROM dbo.REQUERIMIENTO INNER JOIN "
        SqlElemento = SqlElemento & " dbo.REQUELIBOSCAJAS ON dbo.REQUERIMIENTO.IDREQUERIMIENTO = dbo.REQUELIBOSCAJAS.IDREQUERIMIENTOS"
        SqlElemento = SqlElemento & " Where dbo.Requerimiento.ID_CLIENTE = " & cltCliente.Valor
        SqlElemento = SqlElemento & "  And dbo.REQUELIBOSCAJAS.CAJASLIBROS = " & txtFiltro.Text
        
        RsElementos.Open SqlElemento, ConActiva, 0, 1
        
        Do While Not RsElementos.EOF
            FiltroElemento = "," & RsElementos!IDREQUERIMIENTO & FiltroElemento
            RsElementos.MoveNext
        Loop
        
        If FiltroElemento <> "" Then
            sql = sql & vbCrLf & " AND IDREQUERIMIENTO in( " & Mid(FiltroElemento, 2) & ")"
        Else
             MsgBox "No se encontraron registros"
            Exit Sub
        End If
        
        
        
        

    Else
    
    MsgBox "Error en filtro", vbCritical
    Exit Sub
End If




    cltCliente.Visible = True
    lblFiltro.Caption = "Caja/Legajo"
    lblFiltro.Visible = True
    txtFiltro.Visible = True

            End If
            
            
            
            
            If CBool(chkPendientes) Then
                sql = sql & vbCrLf & " AND V_REQUERIMIENTO.IDEstado < 5 "
            End If
            If CBool(chkHojaRutaSinAsignar) Then
            
'            1   Consulta normal Cajas
'2   Consulta normal Libros
'3   Consulta urgente Cajas
'4   Consulta urgente libros
'5   Pedido de Referencia y Trasvase
'6   Retiro de cajas por devolución
'7   Pedidos de cajas vacias
'8   Busqueda de Documentación
'9   Consulta en Planta
'10  Consulta Normal Legajos
'11  Consulta Urgente Legajos
'12  Consulta Telefonica
'13  Consulta Digital
'14  Consulta Por Fax
'15  Tramites Administrativos
'16  Pedidos de actualizacion de referencias
'17  Filtro de  referencia
'18  Busqueda de documentos
'19  Horas de archivista

            
                sql = sql & vbCrLf & " AND ((COD_HOJA_RUTA_TERMINADO IS NULL) OR (COD_HOJA_RUTA_TERMINADO = 0) )  AND  V_REQUERIMIENTO.IDEstado = 5 AND   ( ENVIOPORCORREO = 0 or  ENVIOPORCORREO is null)   "
            End If
            
            
            If CBool(chkRemitosPendientes) Then
                sql = sql & vbCrLf & " AND V_REQUERIMIENTO.IDEstado = 6 "
                sql = sql & vbCrLf & " AND TIEMPOTOTAL IS NULL  "
                sql = sql & vbCrLf & " AND IDTIPOREQUERIMIENTO IN (1,2,3,4,7,10,11)  "
            End If
            If cboFiltro.Text = "Por Cliente" Then
                If Not IsNull(cltCliente.Valor) Then
                    sql = sql & vbCrLf & " AND V_REQUERIMIENTO.ID_Cliente= " & cltCliente.Valor
                 Else
                    MsgBox "Error en cliente", vbCritical
                    Exit Sub
                End If
            End If
            If cboFiltro.Text = "Por Remito" Then
                If txtFiltro <> "" Then
                    sql = sql & vbCrLf & " AND idremito IN( " & txtFiltro & ")"
                End If
            End If
            
            If cboFiltro.Text = "Hoja de Ruta" Then
                If txtFiltro <> "" Then
                    sql = sql & vbCrLf & " AND COD_HOJA_RUTA_TERMINADO =  " & txtFiltro
                End If
            End If
            
            
            
            If cboFiltro.Text = "Requerimiento" Then
                If txtFiltro <> "" Then
                    sql = sql & vbCrLf & " AND IDREQUERIMIENTO in( " & txtFiltro & ")"
                    txtFiltro.Text = ""
                End If
            End If
            
            If Not IsNull(ctlClienteUsuario.Valor) Then
                sql = sql & vbCrLf & " AND  COD_USUARIO_CLIENTE  =" & ctlClienteUsuario.Valor
                ctlClienteUsuario.Valor = Null
                txtFiltro.Text = ""
            End If
            
            
            
             If cboFiltro.Text = "Mensajes sin Leer" Then
                 sql = sql & vbCrLf & " AND  NOT (DESCRIPCION_ACTUALIZADA IS NULL)"
             End If
             
            
            
          Dim TituloGrilla As String

             
             


             
            TituloGrilla = "  SELECT  top (1000)   dbo.V_REQUERIMIENTO.IDREQUERIMIENTO, dbo.CLIENTES.RAZON_SOCIAL, dbo.V_REQUERIMIENTO.FECHARECEPCION,"
             TituloGrilla = TituloGrilla & vbCrLf & "  dbo.V_REQUERIMIENTO.FECHAENTREGA, dbo.V_REQUERIMIENTO.COMPROMISO_ENTREGA, dbo.V_REQUERIMIENTO.TIPO_DESCRIPCION,"
             TituloGrilla = TituloGrilla & vbCrLf & "  dbo.V_REQUERIMIENTO.CANTIDAD , dbo.REQUERIMIENTO_ESTADO.DESCRIPCION,"
              TituloGrilla = TituloGrilla & vbCrLf & "  dbo.V_REQUERIMIENTO.APELLIDO_NOMBRE , dbo.V_REQUERIMIENTO.SECTOR_REQUERIMIENTO ,"
              TituloGrilla = TituloGrilla & vbCrLf & "  dbo.V_REQUERIMIENTO.COD_HOJA_RUTA_TERMINADO , dbo.V_REQUERIMIENTO.IDREMITO , FK_SUCURSAL"
             TituloGrilla = TituloGrilla & vbCrLf & "  FROM         dbo.CLIENTES INNER JOIN"
             TituloGrilla = TituloGrilla & vbCrLf & "  dbo.V_REQUERIMIENTO ON dbo.CLIENTES.ID_CLIENTE = dbo.V_REQUERIMIENTO.ID_CLIENTE INNER JOIN"
             TituloGrilla = TituloGrilla & vbCrLf & "  dbo.REQUERIMIENTO_ESTADO ON dbo.V_REQUERIMIENTO.IDESTADO = dbo.REQUERIMIENTO_ESTADO.ID_ESTADO"
             TituloGrilla = TituloGrilla & vbCrLf & "  Where (dbo.V_REQUERIMIENTO.ANULADO Is Null)"


             
             
        Dim rsGrilla As New ADODB.Recordset
             
        rsGrilla.CursorLocation = adUseClient
        
        
             
            sql = sql & vbCrLf & " order by FK_SUCURSAL,  FECHAENTREGA , COMPROMISO_ENTREGA,  IDTIPOREQUERIMIENTO"
    
   Rem  rsGrilla.Open TituloGrilla & sql, strConBasa , adOpenDynamic, adLockReadOnly
 Rem   Set DataGrid1.DataSource = rsGrilla.DataSource
    
   Set rsRequerimientos = New ADODB.Recordset
   rsRequerimientos.CursorLocation = adUseClient
   
   
  
    rsRequerimientos.Open sqlTitulo & sql, conreque, adOpenStatic, adLockReadOnly
    trvEstado.Nodes.Clear
    trvEstado.Style = 5
     
     
  Set TRequerimientos = New TRequerimientos
  Rem TRequerimientos.Clear
   
    Do While Not rsRequerimientos.EOF
        With rsRequerimientos
        Debug.Print
            IDREQUERIMIENTO = !IDREQUERIMIENTO
            Debug.Print IDREQUERIMIENTO
            ID_CLIENTE = !ID_CLIENTE
            IDPERSONAL2 = !IDPERSONAL
            IDTIPORECEPCION = !IDTIPORECEPCION
            IDESTADO = !IDESTADO
            IDTIPOREQUERIMIENTO = !IDTIPOREQUERIMIENTO

            If IsNull(!PEDIDOCLIENTE) Then
                PEDIDOCLIENTE = ""
            Else
                PEDIDOCLIENTE = !PEDIDOCLIENTE
            End If
            If IsNull(!SECTOR_REQUERIMIENTO) Then
                Sector = ""
            Else
                Sector = !SECTOR_REQUERIMIENTO
            End If
            If IsNull(!TELEFONO) Then
                TELEFONO = ""
            Else
                TELEFONO = !TELEFONO
            End If
            
            If Not IsNull(!APELLIDO_NOMBRE) Then
                PEDIDO_APELLIDO_NOMBRE = Trim(!APELLIDO_NOMBRE)
            Else
                PEDIDO_APELLIDO_NOMBRE = ""
            End If
            
            
            If IsNull(!DESCRIPCION) Then
                DESCRIPCION2 = ""
            Else
                DESCRIPCION2 = !DESCRIPCION
            End If
            If IsNull(!SOLICITANTE) Then
                SOLICITANTE = ""
            Else
                SOLICITANTE = !SOLICITANTE
            End If
            If IsNull(!TOMO) Then
                TOMO = 0
            Else
                TOMO = !TOMO
            End If
            If IsNull(!FECHAENTREGA) Then
                FECHAENTREGA = 0
            Else
                FECHAENTREGA = !FECHAENTREGA
            End If
            If IsNull(!FECHALIMITE) Then
                FECHALIMITE = 0
            Else
                FECHALIMITE = !FECHALIMITE
            End If
            FECHARECEPCION = Format(!FECHARECEPCION, "DD/MM/YY HH:mm")
            CANTIDAD = !CANTIDAD
            If IsNull(!IDREMITO) Then
                IDREMITO = 0
            Else
                IDREMITO = !IDREMITO
            End If
            If IsNull(!TIEMPOTOTAL) Then
                TIEMPOTOTAL = ""
            Else
                TIEMPOTOTAL = !TIEMPOTOTAL
            End If
            If IDREQUERIMIENTOANT <> IDREQUERIMIENTO Then
                IDREQUERIMIENTOANT = IDREQUERIMIENTO
                TRequerimientos.Add IDREQUERIMIENTO, CLng(ID_CLIENTE), IDPERSONAL2, _
                IDTIPORECEPCION, IDESTADO, IDTIPOREQUERIMIENTO, IDFAX, Sector, _
                TELEFONO, DESCRIPCION2, SOLICITANTE, TOMO, FECHAENTREGA, FECHALIMITE, _
                FECHARECEPCION, CStr(CANTIDAD), IDREMITO, TIEMPOTOTAL, CStr(!IDREQUERIMIENTO)
            End If
            If IsNull(rsRequerimientos!TIPO_DESCRIPCION) Then
                DESCRIPCION = "0"
            Else
                DESCRIPCION = UCase(CStr(rsRequerimientos!TIPO_DESCRIPCION))
            End If
            
            If IsNull(rsRequerimientos!FECHA_SISTEMA) Then
                FECHA_SISTEMA = ""
            Else
                FECHA_SISTEMA = rsRequerimientos!FECHA_SISTEMA
            End If
            
            
            
    End With
        ConstruirArbol rsRequerimientos.Fields
        rsRequerimientos.MoveNext

    Loop
    
salir:
 If Err.Number <> 0 Then
    MsgBox Err.Description
    Exit Sub
    
 End If
    
  On Error GoTo SALIR2:
  If nodoSelecionado <> "" Then
    
     ExpanderParent trvEstado.Nodes(nodoSelecionado).Index
      trvEstado.Nodes(nodoSelecionado).Bold = True
    
     End If
     
    trvEstado.Refresh
SALIR2:



End Sub

Public Function ColocarColorRequ(Item As String, ESTADO As Integer, Tiporequerimiento As Integer, Fecha As Date, TRV As TreeView)
 Dim DiferenciaH As Long
 Dim Bandera As Boolean
Bandera = False
 DiferenciaH = DateDiff("H", Fecha, SysDateCompare)
    
    
    If ESTADO < 6 Then
        If TRV.Nodes.Item("F " & Format(Fecha, "DD/MM/YYYY")).Image <> "SRojo" Then
            TRV.Nodes.Item("F " & Format(Fecha, "DD/MM/YYYY")).Image = "SAmarillo"
        End If
        If TRV.Nodes.Item("T " & Format(Fecha, "DD/MM/YYYY") & "  " & Format(CStr(Tiporequerimiento), "00")).ForeColor <> &HFF& Then
            TRV.Nodes.Item("T " & Format(Fecha, "DD/MM/YYYY") & "  " & Format(CStr(Tiporequerimiento), "00")).ForeColor = &H80FF&
        End If
        
        
        Select Case Tiporequerimiento
        Case 1, 2
             If DiferenciaH > 24 Then
                   TRV.Nodes(Item).ForeColor = &HFF&
                   TRV.Nodes.Item("T " & Format(Fecha, "DD/MM/YYYY") & "  " & Format(CStr(Tiporequerimiento), "00")).ForeColor = &HFF&
                   Bandera = True
             Else
                   TRV.Nodes(Item).ForeColor = &HFF0000
                   Bandera = False
             End If
        Case 3, 4
             If DiferenciaH > 2 Then
                   TRV.Nodes(Item).ForeColor = &HFF&
                   TRV.Nodes.Item("T " & Format(Fecha, "DD/MM/YYYY") & "  " & Format(CStr(Tiporequerimiento), "00")).ForeColor = &HFF&
                   Bandera = True
              Else
                   TRV.Nodes(Item).ForeColor = &HFF0000
                Bandera = False
             End If
        Case 5, 6, 7
             If DiferenciaH > 24 Then
                   TRV.Nodes(Item).ForeColor = &HFF&
                   TRV.Nodes.Item("T " & Format(Fecha, "DD/MM/YYYY") & "  " & Format(CStr(Tiporequerimiento), "00")).ForeColor = &HFF&
                   Bandera = True
             Else
                   TRV.Nodes(Item).ForeColor = &HFF0000
                   Bandera = False
             End If
        Case 8
          If DiferenciaH > 48 Then
                   TRV.Nodes(Item).ForeColor = &HFF&
                   TRV.Nodes.Item("T " & Format(Fecha, "DD/MM/YYYY") & "  " & Format(CStr(Tiporequerimiento), "00")).ForeColor = &HFF&
                   Bandera = True
           Else
                   TRV.Nodes(Item).ForeColor = &HFF0000
                   Bandera = False
           End If
        End Select
    Else
       ColocarColorRequ = &H8000&
            
        '--------------------------
            Select Case Tiporequerimiento
        Case 1, 2
           If TRequerimientos.Item(CStr(CLng(Mid(Item, 2)))).TIEMPOTOTAL <> "" Then
                If TRequerimientos.Item(CStr(CLng(Mid(Item, 2)))).TIEMPOTOTAL > 24 Then
                 TRV.Nodes(Item).Image = "FUERATERMINO"
                Else
                End If
          End If
             
        Case 3, 4
          If TRequerimientos.Item(CStr(CLng(Mid(Item, 2)))).TIEMPOTOTAL <> "" Then
                If TRequerimientos.Item(CStr(CLng(Mid(Item, 2)))).TIEMPOTOTAL > 2 Then
                TRV.Nodes(Item).Image = "FUERATERMINO"
                Else
                End If
          End If
        Case 8
          
        End Select
    
    
    
        '---------------------------
        
    
    End If
    
    
    
    If Bandera Then
        TRV.Nodes.Item("F " & Format(Fecha, "DD/MM/YYYY")).Image = "SRojo"
    End If
    If Format(Fecha, "DD/MM/YYYY") = Format(SysDate, "DD/MM/YYYY") Then
        TRV.Nodes.Item("F " & Format(Fecha, "DD/MM/YYYY")).Image = "SAmarillo"
    End If
    
    ' &H00FF0000& ' azul
    ' &H8000&     ' verde
    ' &HFF&       ' rojo
  
End Function
Public Function EstadoRequerimientoEvaluacion(ESTADO As Integer, Fecha As Date, TRV As TreeView)
    Dim DiferenciaH As Long
    Dim Bandera As Boolean
    Bandera = False
    
    Dim ColorAzul As ColorConstants
    ColorAzul = &HFF0000
    Dim ColorRojo As ColorConstants
    ColorRojo = &HFF&
    Dim ColorVerde As ColorConstants
    ColorVerde = &H8000&
    Dim ColorNegro As ColorConstants
    ColorNegro = &H80000012
    
    
    
'If TRV.Nodes.Item(KeyPadreFecha).Image <> "SRojo" Then
'            TRV.Nodes.Item(KeyPadreFecha).Image = "SAmarillo"
'            If TRV.Nodes.Item(KeyPadreMañana).Image <> "SRojo" Then
'                TRV.Nodes.Item(KeyPadreMañana).Image = "SAmarillo"
'            End If
'        End If
'        If TRV.Nodes.Item(KeyPadreTipo).ForeColor <> ColorRojo Then
'            TRV.Nodes.Item(KeyPadreTipo).ForeColor = &H80FF&
'        End If
    
    
    
    DiferenciaH = DateDiff("H", Fecha, SysDate_DD_MM_YYYY_mm_ss)

     
    If ESTADO < 6 Then
        Select Case Mid(C_TipoRequerimiento, 2)
        Case 1, 2, 10, 12, 13, 14   ' Consulta normal Cajas , Consulta normal Libros ,Consulta Normal Legajos ,Consulta Telefonica ,Consulta Digital ,Consulta Por Fax
             If DiferenciaH > 24 Then
                   TRV.Nodes.Item(KeyPadreSucursal).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreFecha).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreMañana).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreTipo).ForeColor = ColorRojo
                   TRV.Nodes.Item(KeyRequerimiento).ForeColor = ColorRojo
                   Bandera = True
              Else
                If TRV.Nodes.Item(KeyPadreTipo).ForeColor <> ColorRojo Then
                    TRV.Nodes.Item(KeyPadreTipo).ForeColor = ColorNegro
                    TRV.Nodes.Item(KeyRequerimiento).ForeColor = ColorNegro
                End If
                Bandera = False
             End If
        Case 3, 4, 11 ' Consulta urgente Cajas,Consulta urgente libros,Consulta Urgente Legajos
            If DiferenciaH > 4 Then
                   TRV.Nodes.Item(KeyPadreSucursal).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreFecha).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreMañana).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreTipo).ForeColor = ColorRojo
                   TRV.Nodes.Item(KeyRequerimiento).ForeColor = ColorRojo
                   Bandera = True
              Else
                If TRV.Nodes.Item(KeyPadreTipo).ForeColor <> ColorRojo Then
                    TRV.Nodes.Item(KeyPadreTipo).ForeColor = ColorNegro
                    TRV.Nodes.Item(KeyRequerimiento).ForeColor = ColorNegro
                End If
                Bandera = False
             End If
        
        Case 5, 6, 7   ' Pedido de Referencia y Trasvase ,Retiro de cajas por devolución ,Pedidos de cajas vacias

             If DiferenciaH > 48 Then
                   TRV.Nodes.Item(KeyPadreSucursal).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreFecha).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreMañana).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreTipo).ForeColor = ColorRojo
                   TRV.Nodes.Item(KeyRequerimiento).ForeColor = ColorRojo
                   Bandera = True
              Else
                If TRV.Nodes.Item(KeyPadreTipo).ForeColor <> ColorRojo Then
                    TRV.Nodes.Item(KeyPadreTipo).ForeColor = ColorNegro
                    TRV.Nodes.Item(KeyRequerimiento).ForeColor = ColorNegro
                End If
                Bandera = False
             End If
        
        Case Else
             If DiferenciaH > 48 Then
                   TRV.Nodes.Item(KeyPadreSucursal).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreFecha).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreMañana).Image = "SRojo"
                   TRV.Nodes.Item(KeyPadreTipo).ForeColor = ColorRojo
                   TRV.Nodes.Item(KeyRequerimiento).ForeColor = ColorRojo
                   Bandera = True
              Else
                If TRV.Nodes.Item(KeyPadreTipo).ForeColor <> ColorRojo Then
                    TRV.Nodes.Item(KeyPadreTipo).ForeColor = ColorNegro
                    TRV.Nodes.Item(KeyRequerimiento).ForeColor = ColorNegro
                End If
                Bandera = False
             End If
         End Select
   Else
   
   
   
   
   End If
      
     
   
End Function




Public Sub ImprimirRotulos()
    Dim sql As String
    sql = "  SELECT  * "
    sql = sql & vbCrLf & " From CONTENEDOR "
    sql = sql & vbCrLf & " Where COD_CLIENTE = " & TRequerimientos.Item(CStr(CRequerimientos.Item(1).NumeroRequerimiento)).ID_CLIENTE
    sql = sql & vbCrLf & " AND CONTENEDOR.NRO_CAJA in " & Detalle_Requerimiento_Filtro
    sql = sql & vbCrLf & " Order By nro_caja "
    frmReportes.ImprimirReporte PasoReportes & "Rotulo.rpt", sql, True
End Sub

Public Sub Imprimir_Ruta(Filtro As String)
    Dim SQL1 As String
    SQL1 = "   SELECT * "
    SQL1 = SQL1 & vbCrLf & " From V_HOJA_RUTA"
    SQL1 = SQL1 & vbCrLf & "WHERE IDREQUERIMIENTO IN(" & Filtro & ")"
    SQL1 = SQL1 & vbCrLf & " ORDER BY V_HOJA_RUTA.RAZON_SOCIAL"
    frmReportes.ImprimirReporte PasoReportes & "Hoja_Ruta.rpt", SQL1, True
End Sub

Public Sub ImprimirRequerimientoLegajos(IDREQUERIMIENTO As String)

Dim rs As New ADODB.Recordset
rs.Open "SELECT     ID_CLIENTE From Requerimiento WHERE  IDREQUERIMIENTO = " & IDREQUERIMIENTO, ConActiva, 0, 1


    Dim sql As String
    sql = " SELECT * "
    sql = sql & vbCrLf & " From V_REQUERIMIENTO_LEGAJOS"
    sql = sql & vbCrLf & " Where IDREQUERIMIENTO in(" & IDREQUERIMIENTO & ") AND COD_CLIENTE =" & rs!ID_CLIENTE
    sql = sql & vbCrLf & " ORDER BY  ESTANTERIA,  VERTICAL , HORIZONTAL"
    
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoLegajo.rpt", sql, True

End Sub

Public Sub ImprimirRequerimientoLegajosOrden(IDREQUERIMIENTO As String)

Dim rs As New ADODB.Recordset
rs.Open "SELECT     ID_CLIENTE From Requerimiento WHERE  IDREQUERIMIENTO = " & IDREQUERIMIENTO, ConActiva, 0, 1


    Dim sql As String
    sql = " SELECT * "
    sql = sql & vbCrLf & " From V_REQUERIMIENTO_LEGAJOS_ORDEN "
    sql = sql & vbCrLf & " Where IDREQUERIMIENTO in(" & IDREQUERIMIENTO & ") AND COD_CLIENTE =" & rs!ID_CLIENTE
    sql = sql & vbCrLf & " ORDER BY  ESTANTERIA,  VERTICAL , HORIZONTAL"
    
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoLegajo.rpt", sql, True

End Sub



Public Sub ImprimirRequerimientoLegajosTodos(IDREQUERIMIENTO As String)
    Dim sql As String
    sql = " SELECT * "
    sql = sql & vbCrLf & " From V_REQUERIMIENTO_LEGAJOS"
    sql = sql & vbCrLf & " Where IDREQUERIMIENTO in(" & IDREQUERIMIENTO & ")"
    sql = sql & vbCrLf & " ORDER BY  ESTANTERIA,  VERTICAL , HORIZONTAL"
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoLegajoTodos.rpt", sql, True

End Sub



Public Function Detalle_Requerimiento_Filtro() As String
Dim rs As ADODB.Recordset
Dim SQL1 As String
Dim Filtro As String

    SQL1 = " SELECT"
    SQL1 = SQL1 & " REQ.IDREQUERIMIENTOS , REQ.CAJASLIBROS "
    SQL1 = SQL1 & " From "
    SQL1 = SQL1 & "  REQUELIBOSCAJAS REQ "
    SQL1 = SQL1 & " Where REQ.IDREQUERIMIENTOS = " & CRequerimientos.Item(1).NumeroRequerimiento
    Set rs = New ADODB.Recordset
   rs.Open SQL1, ConActiva, 0, 1
   Filtro = ""
    Do While Not rs.EOF
        Filtro = Filtro & CStr(rs!CAJASLIBROS) & ","
        rs.MoveNext
    Loop
If Filtro <> "" Then
    Filtro = "(" & Mid(Filtro, 1, Len(Filtro) - 1) & ")"
    End If
    Detalle_Requerimiento_Filtro = Filtro
End Function

Public Sub Menu_Cajas(ESTADO As Integer)
   mnuCajasImprimir.Enabled = True
   mnuCajasBuscarCajas.Enabled = False
   mnuCajasReImprimirRemito.Enabled = False
   mnuCajasRemito.Enabled = False
   mnuCajasEncontradas.Enabled = False
   mnuCajasTerminado.Visible = True
   Select Case ESTADO
   Case 2
        mnuCajasBuscarCajas.Enabled = True
        EstadoFinal = 3
   Case 3
        mnuCajasEncontradas.Enabled = True
        EstadoFinal = 4
    
   Case 4
        mnuCajasRemito.Enabled = True
        EstadoFinal = 6
    Case 5, 6, 7, 8
        mnuCajasReImprimirRemito.Enabled = True
   End Select
   PopupMenu mnuCajas
End Sub

Public Sub Menu_General(ESTADO)
    mnuGeneralImprimir.Enabled = True
    mnuGeneralCambioDeEstado.Enabled = False
    mnuGeneralTerminado.Enabled = False
    Select Case ESTADO
    Case 2, 1
        mnuGeneralCambioDeEstado.Enabled = True
        EstadoFinal = 5
    Case 5
         mnuGeneralTerminado.Enabled = True
         EstadoFinal = 6
    End Select
    PopupMenu mnuGeneral
End Sub


Public Sub Menu_BusquedaDocumento(ESTADO)
    mnuBuscarDocumentosImprimir.Enabled = True
    mnuBusquedaDocumentoAsignacion.Enabled = False
        mnuBusquedaDocumentoTerminado.Enabled = False
    Select Case ESTADO
    Case 2, 1
        mnuBusquedaDocumentoAsignacion.Enabled = True
        EstadoFinal = 3
    Case 4
         mnuBusquedaDocumentoTerminado.Enabled = True
         EstadoFinal = 6
    End Select
    PopupMenu mnuBusquedaDocumentos
End Sub

Public Sub Menu_Consultas_Digitales(ESTADO)
    mnuConsultas_Digitales_Cantidad_Imagenes.Enabled = False
  
    mnuConsultas_digitales_Finalizado.Enabled = False
    mnuConsultas_digitales_Imprimir.Enabled = True
    mnuConsultas_Digitales_Asignar_Tarea.Enabled = False
    mnuConsultas_Digitales_Imagenes_Encontradas.Enabled = False
    Select Case ESTADO
    Case 2, 1
        mnuConsultas_Digitales_Asignar_Tarea.Enabled = True
        EstadoFinal = 3
    Case 3
            mnuConsultas_Digitales_Imagenes_Encontradas.Enabled = True
            EstadoFinal = 4
    Case 4
        mnuConsultas_Digitales_Cantidad_Imagenes.Enabled = True
       
    Case 5
         mnuConsultas_digitales_Finalizado.Enabled = True
         EstadoFinal = 6
    Case Is > 5
    End Select
    PopupMenu mnuConsultasDigitales
End Sub

Public Sub Menu_Libros(ESTADO As Integer)
    mnuLibrosImprimir.Enabled = True
    mnuLibrosBuscarLibros.Enabled = False
    mnuLibrosReImprimirRemitoLibros.Enabled = False
    mnuLibrosRemitoDeLibros.Enabled = False
        Select Case ESTADO
        Case 2
             mnuLibrosBuscarLibros.Enabled = True
             EstadoFinal = 3
        Case 4
             mnuLibrosRemitoDeLibros.Enabled = True
             EstadoFinal = 6
         Case 5, 6
           mnuLibrosReImprimirRemitoLibros.Enabled = True
         Case 7
            mnuLibrosReImprimirRemitoLibros.Enabled = True
        End Select
        PopupMenu mnuLibros
    
End Sub

Public Sub Menu_Referencia_Trasvase(ESTADO As Integer)
    mnuReferenciaTrasvaseImprimir.Enabled = True
    mnuReferenciaTrasvaseAsignacion.Enabled = False
    mnuReferenciaTrasvaseDiccCompleto.Enabled = True
    mnuReferenciaTrasvaseDiccSector.Enabled = False
    mnuReferenciaTrasvaseTerminado.Enabled = False
    Select Case ESTADO
    Case 2
        mnuReferenciaTrasvaseAsignacion.Enabled = True
        EstadoFinal = 4
    Case 4
         mnuReferenciaTrasvaseTerminado.Enabled = False
         EstadoFinal = 6
    Case 5
        mnuReferenciaTrasvaseTerminado.Enabled = True
        EstadoFinal = 6
        
    End Select
    PopupMenu mnuReferenciaTrasvase
    
End Sub

Public Sub Menu_Cajas_Vacias(ESTADO As Integer)
    mnuCajasVaciasImprimir.Enabled = True
    mnuCajasVaciasCambioEstado.Enabled = False
    mnuCajasVaciasRemito.Enabled = False
    mnuCajasVaciasReImprimirRemito.Enabled = False
    mnuCajasVaciasReImprimirRotulos.Enabled = False
    mnuCajasVaciasImprimirRotuloEtiqueta.Enabled = False
    Select Case ESTADO
    Case 2
        mnuCajasVaciasCambioEstado.Enabled = True
        EstadoFinal = 3
    
    Case 3
        mnuCajasVaciasReImprimirRotulos.Enabled = True
        mnuCajasVaciasImprimirRotuloEtiqueta.Enabled = True
        
        EstadoFinal = 4
        
    Case 4
        mnuCajasVaciasReImprimirRotulos.Enabled = True
        mnuCajasVaciasImprimirRotuloEtiqueta.Enabled = True
        mnuCajasVaciasRemito.Enabled = True
        EstadoFinal = 5
    Case 5, 6, 7, 8
        mnuCajasVaciasReImprimirRotulos.Enabled = True
        mnuCajasVaciasImprimirRotuloEtiqueta.Enabled = True
        mnuCajasVaciasReImprimirRemito.Enabled = True
        Case 5, 6
    End Select
        PopupMenu mnuCajasVacias
End Sub

Public Sub Menu_Consulta_Planta(ESTADO As Integer)
   mnuCajasImprimir.Enabled = True
   mnuCajasBuscarCajas.Enabled = False
   mnuCajasHojaRuta.Visible = False
   mnuCajasEncontradas.Visible = False
   mnuCajasReImprimirRemito.Visible = False
   mnuCajasRemito.Visible = False
   mnuCajasTerminado.Enabled = False
   Select Case ESTADO
   Case 2
        mnuCajasBuscarCajas.Enabled = True
        EstadoFinal = 3
   Case 4
        mnuCajasTerminado.Enabled = True
   End Select
   PopupMenu mnuCajas
End Sub

Public Sub Menu_Consulta_Cupones(ESTADO As Integer)
   mnuCajasImprimir.Enabled = True
   mnuCajasBuscarCajas.Enabled = False
   mnuCajasHojaRuta.Visible = False
   mnuCajasEncontradas.Visible = False
   mnuCajasReImprimirRemito.Visible = False
   mnuCajasRemito.Visible = False
   mnuCajasTerminado.Enabled = False
   Select Case ESTADO
   Case 2
        mnuCajasBuscarCajas.Enabled = True
        EstadoFinal = 3
   Case 4
        mnuCajasTerminado.Enabled = True
   End Select
   PopupMenu mnuCajas
End Sub


Public Sub Menu_Legajos(ESTADO As Integer)
            mnuLegajosImprimir.Enabled = True
            mnuLegajosBuscarLegajos.Enabled = False
            mnuLegajosImprimirEtiquetas.Enabled = False
            mnuLegajosReImprimirRemito.Enabled = False
            mnuLegajosEncontrados.Enabled = False
            mnuLegajosRemitos.Enabled = False
            Select Case ESTADO
            Case 2
                  mnuLegajosImprimirEtiquetas.Enabled = True
                  mnuLegajosBuscarLegajos.Enabled = True
                  EstadoFinal = 3
            Case 3
                mnuLegajosImprimirEtiquetas.Enabled = True
                mnuLegajosEncontrados.Enabled = True
                EstadoFinal = 4
            
            Case 4
                mnuLegajosImprimirEtiquetas.Enabled = True
                mnuLegajosRemitos.Enabled = True
                EstadoFinal = 5
            Case 5, 6, 7, 8
                mnuLegajosImprimirEtiquetas.Enabled = True
                mnuLegajosReImprimirRemito.Enabled = True
            End Select
            PopupMenu mnuLegajos
End Sub

Public Sub ImprimirRequerimientoCajas(IDREQUERIMIENTO As Long)
 Dim sql As String
    sql = " SELECT * "
    sql = sql & vbCrLf & "FROM V_REQUERIMIENTO_CAJA "
    sql = sql & vbCrLf & " Where IDREQUERIMIENTO =" & IDREQUERIMIENTO
    sql = sql & vbCrLf & "  ORDER BY IDREQUERIMIENTO, ESTANTERIA,VERTICAL , HORIZONTAL"
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoCaja.rpt", sql, True
End Sub

Public Sub ImprimirRequerimientoCajasTodo(requerimientos As String)
 Dim sql As String
    sql = " SELECT * "
    sql = sql & vbCrLf & "FROM V_REQUERIMIENTO_CAJA "
    sql = sql & vbCrLf & " Where IDREQUERIMIENTO in(" & requerimientos & ")"
    sql = sql & vbCrLf & "  ORDER BY IDREQUERIMIENTO, ESTANTERIA, HORIZONTAL"
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoCajatodo.rpt", sql, True
End Sub
Public Sub ImprimirRequerimientoCajasTodos(requerimientos As String)
 Dim sql As String
    
    
    sql = "  SELECT * "
    sql = sql & vbCrLf & " From V_REQUERIMIENTO_CAJA"
    sql = sql & vbCrLf & " WHERE (IDTIPOREQUERIMIENTO = 1) "
     Rem sql = sql & vbCrLf & " AND (IDESTADO < 5) "
    sql = sql & vbCrLf & "  AND IDREQUERIMIENTO in(" & requerimientos & ")"
     frmReportes.ImprimirReporte PasoReportes & "TodosRequerimientosCajas.rpt", sql, True
End Sub
Public Sub ImprimirRequerimientoLibros(IDREQUERIMIENTO As Long)
   Dim sql As String
    sql = " SELECT *  "
    sql = sql & vbCrLf & " FROM   V_REQUERIMIENTO_LIBROS "
    sql = sql & vbCrLf & " Where IDREQUERIMIENTO = " & IDREQUERIMIENTO
    sql = sql & vbCrLf & " ORDER BY IDREQUERIMIENTO "
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoLibro.rpt", sql, True
  
 
  
  
End Sub
Public Sub ImprimirRequerimientoGeneral(Requerimiento As Long)
    Dim sql As String
    sql = " SELECT *  "
    sql = sql & vbCrLf & " FROM   V_REQUERIMIENTO_GENERICO"
    sql = sql & vbCrLf & " Where IDREQUERIMIENTO = " & Requerimiento
    sql = sql & vbCrLf & " ORDER BY IDREQUERIMIENTO "
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoGenerico.rpt", sql, True
End Sub
Public Sub ImprimirRequerimientoVacias(Requerimiento As Long)
    Dim sql As String
    sql = " SELECT *  "
    sql = sql & vbCrLf & " FROM   V_REQUERIMIENTO_VACIAS"
    sql = sql & vbCrLf & " Where IDREQUERIMIENTO = " & Requerimiento
    sql = sql & vbCrLf & " ORDER BY IDREQUERIMIENTO "
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoVacias.rpt", sql, True
End Sub

Public Function Expander(IndiceNumerico As Integer)
    Dim indexs As Integer
    indexs = IndiceNumerico
    Dim ArbolMañana As Nodes
    Dim i As Integer
    Dim Indice As Integer
    Dim d As Node

    On Error Resume Next
            trvEstado.Nodes.Item(indexs).Expanded = True ' Fecha
            trvEstado.Nodes.Item(indexs).Child.Expanded = True ' mañana
            trvEstado.Nodes.Item(indexs).Child.Child.Expanded = True
'    luis        trvEstado.Nodes.Item(indexs).Child.Child.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Child.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Child.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Child.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Child.Next.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Child.Next.Next.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Child.Next.Next.Next.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Child.Next.Next.Next.Next.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Child.Next.Next.Next.Next.Next.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Child.Next.Next.Next.Next.Next.Next.Next.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Next.Expanded = True ' Tarde
'            trvEstado.Nodes.Item(indexs).Child.Next.Child.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Next.Child.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Next.Child.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Next.Child.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Next.Child.Next.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Next.Child.Next.Next.Next.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Next.Child.Next.Next.Next.Next.Next.Next.Next.Next.Next.Expanded = True
'            trvEstado.Nodes.Item(indexs).Child.Next.Child.Next.Next.Next.Next.Next.Next.Next.Next.Next.Next.Next.Expanded = True
                  
End Function

Public Function ExpanderParent(IndiceNumerico As Integer)
    Dim indexs As Integer
    indexs = IndiceNumerico
    Dim ArbolMañana As Nodes
    Dim i As Integer
    Dim Indice As Integer
    On Error Resume Next
    
    
    
      
      
    

    
            trvEstado.Nodes.Item(indexs).Parent.Parent.Parent.Parent.Expanded = True
            trvEstado.Nodes.Item(indexs).Parent.Parent.Parent.Expanded = True
            trvEstado.Nodes.Item(indexs).Parent.Parent.Expanded = True
            trvEstado.Nodes.Item(indexs).Parent.Expanded = True
            trvEstado.Nodes.Item(indexs).Expanded = True
                  
End Function
Public Sub ImprimirEtiquetasLegajos(requerimientos As String)



        Dim sql As String
        Dim ROLLO As String
        Dim intROLLO As String
        Dim LEGAJO_1   As String
        Dim LEGAJO_2 As String
        Dim BARRA_1 As String
        Dim BARRA_2  As String
        Dim DIGITO_1 As String
        Dim DIGITO_2 As String
        Dim rs As New ADODB.Recordset


sql = " SELECT     LEGAJOS.ID_CLIENTE_LEGAJO, REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.ID_CLIENTE, LEGAJOS.ID_LEGAJO, LEGAJOS.DIGITO_VERIFICADOR,"
sql = sql & vbCrLf & " LEGAJOS.NRO_DESDE , LEGAJOS.LETRA_DESDE, LEGAJOS.FECHA_DESDE, LEGAJOS.DESCRIPCION"
sql = sql & vbCrLf & " FROM REQUERIMIENTO INNER JOIN"
sql = sql & vbCrLf & " REQUELIBOSCAJAS ON REQUERIMIENTO.IDREQUERIMIENTO = REQUELIBOSCAJAS.IDREQUERIMIENTOS INNER JOIN"
sql = sql & vbCrLf & " LEGAJOS ON REQUELIBOSCAJAS.CAJASLIBROS = LEGAJOS.ID_CLIENTE_LEGAJO AND REQUERIMIENTO.ID_CLIENTE = LEGAJOS.COD_CLIENTE INNER JOIN"
sql = sql & vbCrLf & " CONTENEDOR ON LEGAJOS.NRO_CAJA = CONTENEDOR.NRO_CAJA AND LEGAJOS.COD_CLIENTE = CONTENEDOR.COD_CLIENTE"
sql = sql & vbCrLf & " Where Requerimiento.IDREQUERIMIENTO in( " & requerimientos & ")"
sql = sql & vbCrLf & " ORDER BY CONTENEDOR.ESTANTERIA, CONTENEDOR.VERTICAL, CONTENEDOR.HORIZONTAL"



Rem   '" & "L2" & Format(rs!ID_LEGAJO, "0000000") & "','" & CStr(rs!ID_LEGAJO) & "' ," & rs
  
  rs.Open sql, ConActiva, 0, 1
  Dim Orden As Integer
  Dim DESCRIPCION As String
  Orden = 1
  
   ExecutarSql "DELETE FROM TEM_LEGAJOS "
   
  Do While Not rs.EOF
    LEGAJO_1 = "'" & rs!ID_CLIENTE_LEGAJO & "'"
    BARRA_1 = "'" & "12" & Format(rs!ID_LEGAJO, "0000000000") & "'"
    DIGITO_1 = "'" & rs!Digito_Verificador & "'"
    DESCRIPCION = rs!NRO_DESDE & " " & Trim(rs!LETRA_DESDE) & " " & Format(rs!FECHA_DESDE, "yyyy") & " " & Trim(rs!DESCRIPCION)
    sql = " INSERT INTO TEM_LEGAJOS (ORDEN,  LEGAJO_1,  BARRA_1,  DIGITO_1, DESCRIPCION) "
    sql = sql & " VALUES     (" & Orden & "," & LEGAJO_1 & "," & BARRA_1 & "," & DIGITO_1 & ",'" & DESCRIPCION & "') "
    ExecutarSql sql
    Orden = Orden + 1
    rs.MoveNext
  Loop
   
   
   
 sql = " SELECT * "
  sql = sql & "  From TEM_LEGAJOS "
 sql = sql & "  ORDER BY ORDEN"
 frmReportes.ImprimirReporte PasoReportes & "Requerimiento_Etiquetas_Legajos.rpt", sql, True

End Sub

Public Sub ImprimirIndiceRequerimiento(IDREQUERIMIENTO As Long)
'    Dim sql As String
'    Dim rs As New ADODB.Recordset
'    Dim DESCRIPCION As String
'    MousePointer = 11
'    sql = "  SELECT CLIENTEUSUARIO.COD_INDICE,"
'    sql = sql & " REQUERIMIENTO.IDREQUERIMIENTO,REQUERIMIENTO.DESCRIPCION ,"
'    sql = sql & " REQUERIMIENTO.ID_CLIENTE, PERSONAL.APELLIDO,Personal.NOMBRE"
'    sql = sql & " From Requerimiento, CLIENTEUSUARIO, Personal"
'    sql = sql & " Where Requerimiento.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO  AND"
'    sql = sql & " REQUERIMIENTO.IDPERSONAL = PERSONAL.IDPERSONAL AND"
'    sql = sql & " REQUERIMIENTO.IDREQUERIMIENTO =" & IDREQUERIMIENTO
'    rs.Open sql, ConActiva, 0, 1
'     If Not rs.EOF Then
'        sql = " SELECT * "
'        sql = sql & "  From V_INDICES"
'        sql = sql & " Where COD_CLIENTE = " & rs!ID_CLIENTE
'        sql = sql & " AND INDICE l ike '" & rs!Cod_Indice & "%'"
'        sql = sql & " ORDER BY INDICE"
'        If IsNull(rs!DESCRIPCION) Then
'            DESCRIPCION = ""
'        Else
'            DESCRIPCION = " Descripcion " & Trim(rs!DESCRIPCION)
'        End If
'        frmReportes.ImprimirReporte PasoReportes + "rptindices.rpt", sql, False, "", "Requerimiento: " & rs!IDREQUERIMIENTO & " Referencia y Trasvase Responsable: " & Trim(rs!Apellido) & " " & Trim(rs!Nombre) & " " & DESCRIPCION
'    End If
'    M
    Dim sql As String
    sql = " SELECT *  "
    sql = sql & vbCrLf & " FROM   V_REQUERIMIENTO_GENERICO"
    sql = sql & vbCrLf & " Where IDREQUERIMIENTO = " & IDREQUERIMIENTO
    sql = sql & vbCrLf & " ORDER BY IDREQUERIMIENTO "
    frmReportes.ImprimirReporte PasoReportes & "RequerimientoGenerico.rpt", sql, True
 
    MousePointer = 0
End Sub

Public Sub Menu_Tipo_Cajas_Vacias(Fecha As String)
        PopupMenu mnuTipoRequerimientoCajasVacias
End Sub



Private Sub txtFiltro_LostFocus()
    If cboFiltro.Text = "Por cliente y Solicitante" Then
        ctlClienteUsuario.LlenarConCliente (txtFiltro.Text)
     End If

End Sub

Public Function RequerimientoSelecion() As String
Dim Filtro As String
Dim i As Integer

    For i = 1 To trvEstado.Nodes.Count
        If trvEstado.Nodes.Item(i).Checked = True Then
            Filtro = Filtro & CLng(Mid(trvEstado.Nodes.Item(i).Key, 2)) & ","
        End If
    Next
        
    If Filtro <> "" Then
            Filtro = Mid(Filtro, 1, Len(Filtro) - 1)
    Else
        Filtro = ""
    End If
    RequerimientoSelecion = Filtro
End Function

Private Sub txtRequerimientoEncontrado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim dato As Long
        Dim sql As String
        Dim c As Integer
        If Trim(txtPersonal.Text) = "" Then
            MsgBox "Ingrese el usuario"
            Exit Sub
        End If
        
        If UCase(Mid(txtRequerimientoEncontrado.Text, 1, 2)) = "RD" Then
                sql = " UPDATE REQUELIBOSCAJAS "
                If optSI.Value = True Then
                    sql = sql & " Set ESTADO = 'Encontrado " & SysDate_DD_MM_YYYY_mm_ss & "'"
                End If
               If optBaja.Value = True Then
                    sql = sql & " Set ESTADO = 'Para dar de baja " & MDIfrmInicio.StaInicio.Panels(2).Text & SysDate_DD_MM_YYYY & "'"
                    optSI.Value = True
               End If
                If OptNo.Value = True Then
                    sql = sql & " Set ESTADO ='NO Encontrado " & SysDate_DD_MM_YYYY_mm_ss & "'"
                    optSI.Value = True
                End If
                
                sql = sql & " ,  Personal = " & txtPersonal.Text
                sql = sql & "  Where ID = " & Mid(txtRequerimientoEncontrado.Text, 3)
                c = ExecutarSql(sql)
                If c <> 1 Then
                    MsgBox "No se actualizo"
                End If
                    txtRequerimientoEncontrado.Text = ""
                    Beep
                Exit Sub
         End If
    
        If UCase(Mid(txtRequerimientoEncontrado.Text, 1, 1)) = "D" Then
                Dim rs As New ADODB.Recordset
                dato = Mid(txtRequerimientoEncontrado.Text, 3)
                sql = "   SELECT     IDREQUERIMIENTO, IDTIPOREQUERIMIENTO"
                sql = sql & " From Requerimiento"
                sql = sql & " Where Requerimiento.IDREQUERIMIENTO =  " & dato
                
                Dardebaja CLng(dato)
                rs.Open sql, ConActiva, 0, 1
                
                
                If Not rs.EOF Then
                 If rs!IDTIPOREQUERIMIENTO = 1 Or rs!IDTIPOREQUERIMIENTO = 3 Or rs!IDTIPOREQUERIMIENTO = 10 Or rs!IDTIPOREQUERIMIENTO = 11 Then
                        sql = "   SELECT     IDTIPOREQUERIMIENTO, REQUELIBOSCAJAS.ESTADO "
                        sql = sql & " FROM         REQUELIBOSCAJAS INNER JOIN"
                        sql = sql & " REQUERIMIENTO ON REQUELIBOSCAJAS.IDREQUERIMIENTOS = REQUERIMIENTO.IDREQUERIMIENTO"
                        sql = sql & " Where Requerimiento.IDREQUERIMIENTO =  " & dato
                        sql = sql & " And  (REQUELIBOSCAJAS.ESTADO Is Null) "
                        Set rs = New ADODB.Recordset
                        rs.Open sql, ConActiva, 0, 1
                        If rs.EOF Then
                            sql = " Update dbo.Requerimiento"
                            sql = sql & " Set IDESTADO = 4"
                            sql = sql & " Where IDREQUERIMIENTO = " & dato
                            sql = sql & " And IDESTADO = 3 "
                            ExecutarSql sql
                            txtRequerimientoEncontrado.Text = ""
                            Beep
                        Else
                            MsgBox "Faltan elementos del requerimiento "
                        End If
                Else
                    sql = " Update dbo.Requerimiento"
                    sql = sql & " Set IDESTADO = 4"
                    sql = sql & " Where IDREQUERIMIENTO = " & dato
                    sql = sql & " And IDESTADO = 3 "
                    ExecutarSql sql
                    txtRequerimientoEncontrado.Text = ""
                    Beep
                End If
           End If
                
    End If
txtRequerimientoEncontrado.Text = ""
End If

End Sub

Public Sub ConstruirArbol(rs As ADODB.Fields)

Dim i  As Integer

    
      Rem     IDREQUERIMIENTO , IDTIPOREQUERIMIENTO, FECHAENTREGA, FK_SUCURSAL, COMPROMISO_ENTREGA, DESCRIPCION_ACTUALIZADA
        For i = 0 To 10
            If C_Sucursal = "S " & rs!FK_SUCURSAL Then
                If C_Fecha = "F " & Format(rs!FECHAENTREGA, "DD/MM/YYYY") Then
                    If C_MañanaTarde = "M " & rs!COMPROMISO_ENTREGA Then
                        If C_TipoRequerimiento = "T " & Format(rs!IDTIPOREQUERIMIENTO, "00") Then
                            InsertarRequerimientoArbol rs
                            Exit For
                        Else
                            C_TipoRequerimiento = "T " & Format(CStr(rs!IDTIPOREQUERIMIENTO), "00")
                            KeyPadreTipo = KeyPadreMañana & "  " & C_TipoRequerimiento
                             trvEstado.Nodes.Add KeyPadreMañana, tvwChild, KeyPadreTipo, Trim(rs!TIPO_DESCRIPCION), "E7"
                            trvEstado.Nodes.Item(KeyPadreTipo).Tag = C_TipoRequerimiento
                        End If
                    Else
                        C_MañanaTarde = "M " & Format(Trim(rsRequerimientos!COMPROMISO_ENTREGA), "      ")
                        KeyPadreMañana = KeyPadreFecha & " " & C_MañanaTarde
                        trvEstado.Nodes.Add KeyPadreFecha, tvwChild, KeyPadreMañana, Trim(rsRequerimientos!COMPROMISO_ENTREGA), "SVerde"
                        trvEstado.Nodes.Item(KeyPadreMañana).Tag = C_MañanaTarde
                        C_TipoRequerimiento = ""
                    End If
                Else
                    C_Fecha = "F " & Format(CStr(rs!FECHAENTREGA), "DD/MM/YYYY")
                    KeyPadreFecha = KeyPadreSucursal & C_Fecha
                    trvEstado.Nodes.Add C_Sucursal, tvwChild, KeyPadreFecha, Format(CStr(rs!FECHAENTREGA), "DD/MM/YYYY"), "SVerde"
                    trvEstado.Nodes.Item(KeyPadreFecha).Tag = C_Fecha
                    C_TipoRequerimiento = ""
                    C_MañanaTarde = ""
               End If
            Else
                C_Sucursal = "S " & rs!FK_SUCURSAL
                KeyPadreSucursal = C_Sucursal
                trvEstado.Nodes.Add , , C_Sucursal, rs!FK_SUCURSAL, "SVerde"
                trvEstado.Nodes.Item(KeyPadreSucursal).Tag = C_Sucursal
                C_Fecha = ""
                C_TipoRequerimiento = ""
                C_MañanaTarde = ""
            End If
        Next
End Sub

Public Sub InsertarRequerimientoArbol(rs As ADODB.Fields)
   
    Dim RequeDes As String
    Dim Remito As String
    Dim IconoRequerimiento As String
    Dim RazonSoc  As String
    Dim PEDIDO_APELLIDO_NOMBRE As String
    Dim Requerimiento  As String
    Dim Nodo As Integer
    Dim Sector As String
           If IsNull(rs!SECTOR_REQUERIMIENTO) Then
                Sector = ""
            Else
                Sector = rs!SECTOR_REQUERIMIENTO
            End If
           
    
    
    
    
    
    
    
    
    
                If Not IsNull(rs!APELLIDO_NOMBRE) Then
                    PEDIDO_APELLIDO_NOMBRE = Trim(rs!APELLIDO_NOMBRE)
                Else
                    PEDIDO_APELLIDO_NOMBRE = ""
                End If
    
    
    
    
    
        
        RazonSoc = Mid(Trim(UCase(CStr(rsRequerimientos!Razon_Social))), 1, 30)
        RequeDes = Format(CStr(rsRequerimientos!IDREQUERIMIENTO), "0000000") & "  // " & rsRequerimientos!ID_CLIENTE & "- " & Mid(Trim(UCase(CStr(rsRequerimientos!Razon_Social))), 1, 30) & " // Cant.:" & CStr(IIf(IsNull(rsRequerimientos!CANTIDAD), "0", rsRequerimientos!CANTIDAD)) & " // Pedido.:" & Sector & "**" & Trim(PEDIDO_APELLIDO_NOMBRE) & " //  Resp:" & Trim(UCase(rsRequerimientos!Apellido))
            
        
        Select Case rs!IDTIPOREQUERIMIENTO
        Case 13, 14  'Consulta  digital y por fax
               Requerimiento = RequeDes & "  // Cantidad de Imagenes:" & rsRequerimientos!CANTIDAD_IMAGENES
        Case 1, 2, 3, 4, 7, 10, 11 ' consultas con remito
            If IsNull(rsRequerimientos!IDREMITO) Then
                Remito = 0
            Else
                Remito = Trim(CStr(rsRequerimientos!IDREMITO))
            End If
            If CInt(rsRequerimientos!IDESTADO) < 5 Then
                Requerimiento = RequeDes
            Else
                If IsNull(rsRequerimientos!COD_HOJA_RUTA_TERMINADO) Then
                    Requerimiento = RequeDes & "  // Remito:" & Remito & "  // Hoja Ruta : "
                Else
                    Requerimiento = RequeDes & "  // Remito:" & Remito & "  // Hoja Ruta :" & rs!COD_HOJA_RUTA_TERMINADO
                End If
            End If
        Case Else
                  Requerimiento = RequeDes & "  // Hoja Ruta :" & rs!COD_HOJA_RUTA_TERMINADO
        End Select
        
       Select Case (rsRequerimientos!IDESTADO)
       Case Is = 2
                   IconoRequerimiento = "E" & CStr(rsRequerimientos!IDESTADO)
       Case Is = 5
            If IsNull(rsRequerimientos!ENVIOPORCORREO) Or rsRequerimientos!ENVIOPORCORREO = 0 Then
                IconoRequerimiento = "E" & CStr(rsRequerimientos!IDESTADO)
            Else
                IconoRequerimiento = "E17"
            End If
       Case Else
             IconoRequerimiento = "E" & CStr(rsRequerimientos!IDESTADO)
       End Select
       
       
    Rem    Debug.Assert rs!IDREQUERIMIENTO <> 82784
       
       KeyRequerimiento = "R " & Format(CStr(rs!IDREQUERIMIENTO), "0000000")
        trvEstado.Nodes.Add KeyPadreTipo, tvwChild, KeyRequerimiento, Requerimiento, IconoRequerimiento
       Rem trvEstado.Nodes.Add " F 05/08/2011", tvwChild, KeyRequerimiento, Requerimiento, IconoRequerimiento
         EstadoRequerimientoEvaluacion CInt(rs!IDESTADO), rs!FECHAENTREGA, trvEstado
        Nodo = trvEstado.Nodes.Count
        trvEstado.Nodes.Item(Nodo).Tag = KeyRequerimiento

    If rs!DESCRIPCION_ACTUALIZADA = 0 Then
        trvEstado.Nodes.Item(Nodo).BackColor = &HFFFF&
    End If
    If rs!DESCRIPCION_ACTUALIZADA = 1 Then
        trvEstado.Nodes.Item(Nodo).BackColor = &H80000003
    End If
    If rs!DESCRIPCION_ACTUALIZADA = 2 Then
        trvEstado.Nodes.Item(Nodo).BackColor = &HC0C0FF
    End If
 
    If rs!DESCRIPCION_ACTUALIZADA = 3 Then
        trvEstado.Nodes.Item(Nodo).BackColor = &HFF00&
    End If
    If rs!DESCRIPCION_ACTUALIZADA = 4 Then
        trvEstado.Nodes.Item(Nodo).BackColor = &HFFFF&
    End If
    
    If rs!DESCRIPCION_ACTUALIZADA = 5 Then
        trvEstado.Nodes.Item(Nodo).BackColor = &HFFFF00
    End If
    If rs!DESCRIPCION_ACTUALIZADA = 6 Then
        trvEstado.Nodes.Item(Nodo).BackColor = &HC000&
    End If


End Sub

Public Sub Dardebaja(IDREQUERIMIENTOS As Long)
Dim sql As String
Dim rs As ADODB.Recordset

sql = " SELECT     CAJASLIBROS"
sql = sql & " From REQUELIBOSCAJAS "
sql = sql & "  WHERE     IDREQUERIMIENTOS = " & IDREQUERIMIENTOS
sql = sql & "  AND (ESTADO LIKE 'Para dar %')"

sql = " SELECT     REQUELIBOSCAJAS.CAJASLIBROS, REQUERIMIENTO.ID_CLIENTE"
sql = sql & " FROM         REQUELIBOSCAJAS INNER JOIN"
sql = sql & " REQUERIMIENTO ON REQUELIBOSCAJAS.IDREQUERIMIENTOS = REQUERIMIENTO.IDREQUERIMIENTO"
sql = sql & "  Where Requerimiento.IDREQUERIMIENTO = " & IDREQUERIMIENTOS
sql = sql & "  AND (ESTADO LIKE 'Para dar %')"


Set rs = New ADODB.Recordset


rs.Open sql, ConActiva, 0, 1

Do While Not rs.EOF
        sql = " Update LEGAJOS"
        sql = sql & " Set COD_ESTADO = 8"
        sql = sql & " Where Cod_cliente = " & rs!ID_CLIENTE
        sql = sql & " And ID_LEGAJO = " & rs!CAJASLIBROS
        ExecutarSql sql
        rs.MoveNext
Loop


sql = " DELETE FROM REQUELIBOSCAJAS"
sql = sql & "  WHERE     IDREQUERIMIENTOS = " & IDREQUERIMIENTOS
sql = sql & "  AND ESTADO LIKE 'Para dar %'"
ExecutarSql sql

Set rs = New ADODB.Recordset
sql = " SELECT     COUNT(*) AS CANTIDAD"
sql = sql & " From REQUELIBOSCAJAS"
sql = sql & " Where IDREQUERIMIENTOS = " & IDREQUERIMIENTOS
rs.Open sql, ConActiva, 0, 1

sql = " Update Requerimiento"
sql = sql & " SET  CANTIDAD =" & rs!CANTIDAD
sql = sql & " Where IDREQUERIMIENTO =  " & IDREQUERIMIENTOS
ExecutarSql sql
End Sub
