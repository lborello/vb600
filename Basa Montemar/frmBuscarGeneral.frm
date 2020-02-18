VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmBuscarGenerico 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QUE BUSCAS?"
   ClientHeight    =   10605
   ClientLeft      =   1665
   ClientTop       =   1215
   ClientWidth     =   15030
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
   Picture         =   "frmBuscarGeneral.frx":0000
   ScaleHeight     =   10605
   ScaleWidth      =   15030
   Begin VB.CommandButton CmdBuscar 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   13080
      Picture         =   "frmBuscarGeneral.frx":0C42
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox txtLegajoEstado 
      Height          =   345
      Left            =   6360
      TabIndex        =   108
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   120
      TabIndex        =   87
      Top             =   120
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   661
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   1635
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   180
         TabIndex        =   81
         Top             =   1800
         Width           =   1155
      End
      Begin VB.CheckBox chkRearchivo 
         Caption         =   "Rearchivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   1020
         Width           =   1155
      End
      Begin VB.CheckBox ChkRearchivoDigital 
         Caption         =   "Digital"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   1440
         Width           =   915
      End
      Begin VB.CheckBox chkLegajos 
         Caption         =   "Legajos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox chkReferencias 
         Caption         =   "Referencias"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   180
         Width           =   1275
      End
   End
   Begin TabDlg.SSTab SSTBusqueda 
      Height          =   2415
      Left            =   1740
      TabIndex        =   44
      Top             =   600
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabHeight       =   406
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Búsqueda por Cajas o Etiquetas"
      TabPicture(0)   =   "frmBuscarGeneral.frx":1817
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdUnificarDigital"
      Tab(0).Control(1)=   "cmdBorrarReferencia"
      Tab(0).Control(2)=   "Command5"
      Tab(0).Control(3)=   "TxtOrden"
      Tab(0).Control(4)=   "Command4"
      Tab(0).Control(5)=   "Command3"
      Tab(0).Control(6)=   "cmdLimpiarMemoria"
      Tab(0).Control(7)=   "cmdLegajosEstado"
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(9)=   "cmdActualizarlegajos"
      Tab(0).Control(10)=   "txtCajaHasta"
      Tab(0).Control(11)=   "txtEtiqueta"
      Tab(0).Control(12)=   "txtCaja"
      Tab(0).Control(13)=   "txtEtiquetaHasta"
      Tab(0).Control(14)=   "Label14"
      Tab(0).Control(15)=   "Label10"
      Tab(0).Control(16)=   "Label4"
      Tab(0).Control(17)=   "LblCaja(0)"
      Tab(0).Control(18)=   "Label5"
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Campos"
      TabPicture(1)   =   "frmBuscarGeneral.frx":1833
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblTituloDescripcion"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblTituloLetraDesde"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblTituloLetraHasta"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblTituloNumeroDesde"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblTituloNumeroHasta"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblTituloFechaDesde"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblTituloFechaHasta"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label20"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Combo5"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Combo4"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cboFecha_Hasta"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtDescripcion"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtLetra_Entre"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtLetraDesde"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtLetraHasta"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtNro_Entre"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtNroDesde"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtNroHasta"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtFecha_Entre"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtFechaDesde"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "txtFechaHasta"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "cboFecha_Desde"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Combo6"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "cmdNºdesde"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cmdNºEntre"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "txtDocumento_fecha"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).ControlCount=   28
      TabCaption(2)   =   "Cambio de Indice"
      TabPicture(2)   =   "frmBuscarGeneral.frx":184F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdCambioIndice"
      Tab(2).Control(1)=   "chkCambioDescripcionAnterior"
      Tab(2).Control(2)=   "cmdCambioBuscarIndice"
      Tab(2).Control(3)=   "txtCambioIndiceDocumento"
      Tab(2).Control(4)=   "ctlCambioCliente"
      Tab(2).Control(5)=   "Label12"
      Tab(2).Control(6)=   "lblCambioIndiceID"
      Tab(2).Control(7)=   "lblCambioIndice"
      Tab(2).Control(8)=   "lblCambioDescripcion"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Unificar Legajos"
      TabPicture(3)   =   "frmBuscarGeneral.frx":186B
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtLecturaCodigo"
      Tab(3).Control(1)=   "Text1"
      Tab(3).Control(2)=   "cmdUnificar"
      Tab(3).Control(3)=   "TxtLegajosUnificarPadre"
      Tab(3).Control(4)=   "txtLegajosUnificarHijos"
      Tab(3).Control(5)=   "txtOrdenHijos"
      Tab(3).Control(6)=   "txtImagenesUnificarHijos"
      Tab(3).Control(7)=   "Label2(0)"
      Tab(3).Control(8)=   "Label19"
      Tab(3).Control(9)=   "Label18"
      Tab(3).Control(10)=   "Label16"
      Tab(3).ControlCount=   11
      Begin VB.TextBox txtLecturaCodigo 
         Height          =   330
         Left            =   -67260
         TabIndex        =   120
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H0080FF80&
         Height          =   375
         Left            =   -65940
         TabIndex        =   119
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdUnificar 
         Caption         =   "Unificar"
         Height          =   375
         Left            =   -66540
         TabIndex        =   118
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox TxtLegajosUnificarPadre 
         BackColor       =   &H00FFC0FF&
         Height          =   330
         Left            =   -73440
         TabIndex        =   113
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtLegajosUnificarHijos 
         Height          =   330
         Left            =   -73440
         MultiLine       =   -1  'True
         TabIndex        =   112
         Top             =   1020
         Width           =   5955
      End
      Begin VB.TextBox txtOrdenHijos 
         Height          =   330
         Left            =   -73440
         TabIndex        =   111
         Top             =   1440
         Width           =   5955
      End
      Begin VB.TextBox txtImagenesUnificarHijos 
         Height          =   330
         Left            =   -73440
         TabIndex        =   110
         Top             =   1860
         Width           =   5955
      End
      Begin VB.CommandButton cmdUnificarDigital 
         Caption         =   "Unificar Digital"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -66720
         TabIndex        =   107
         Top             =   1725
         Width           =   1455
      End
      Begin VB.CommandButton cmdBorrarReferencia 
         Caption         =   "Borrar referencia"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -68460
         TabIndex        =   106
         Top             =   1725
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
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
         Left            =   -65700
         TabIndex        =   104
         Top             =   1245
         Width           =   435
      End
      Begin VB.TextBox TxtOrden 
         Height          =   390
         Left            =   -67380
         TabIndex        =   103
         Top             =   1185
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "X"
         Height          =   315
         Left            =   -72240
         TabIndex        =   101
         Top             =   1200
         Width           =   315
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Documento Fecha"
         Height          =   375
         Left            =   -72180
         TabIndex        =   100
         Top             =   1725
         Width           =   1695
      End
      Begin VB.CommandButton cmdCambioIndice 
         Caption         =   "Cambio Indice"
         Height          =   375
         Left            =   -68400
         TabIndex        =   97
         Top             =   1785
         Width           =   1575
      End
      Begin VB.CheckBox chkCambioDescripcionAnterior 
         Caption         =   "Sumar la descripción del indice actual"
         Height          =   255
         Left            =   -74820
         TabIndex        =   96
         Top             =   1905
         Value           =   1  'Checked
         Width           =   4155
      End
      Begin VB.CommandButton cmdCambioBuscarIndice 
         Caption         =   "..."
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
         Left            =   -74160
         TabIndex        =   92
         Top             =   1245
         Width           =   375
      End
      Begin VB.TextBox txtCambioIndiceDocumento 
         Height          =   375
         Left            =   -74880
         TabIndex        =   91
         Top             =   1245
         Width           =   615
      End
      Begin VB.CommandButton cmdLimpiarMemoria 
         Caption         =   "Limpiar Memoria"
         Height          =   375
         Left            =   -70260
         TabIndex        =   90
         Top             =   1725
         Width           =   1575
      End
      Begin VB.TextBox txtDocumento_fecha 
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
         Left            =   1380
         TabIndex        =   89
         Top             =   1905
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdLegajosEstado 
         Caption         =   "Legajo Estado N Fecha"
         Height          =   375
         Left            =   -74760
         TabIndex        =   86
         Top             =   1725
         Width           =   2355
      End
      Begin VB.CommandButton Command2 
         Caption         =   "X"
         Height          =   315
         Left            =   -72240
         TabIndex        =   84
         Top             =   825
         Width           =   315
      End
      Begin VB.CommandButton cmdNºEntre 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         TabIndex        =   83
         Top             =   1485
         Width           =   315
      End
      Begin VB.CommandButton cmdNºdesde 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6840
         TabIndex        =   82
         Top             =   1005
         Width           =   315
      End
      Begin VB.CommandButton cmdActualizarlegajos 
         BackColor       =   &H8000000A&
         Caption         =   "ADMINISTRADOR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -67440
         TabIndex        =   80
         Top             =   645
         Width           =   1695
      End
      Begin VB.TextBox txtCajaHasta 
         BackColor       =   &H00C0C0FF&
         Height          =   330
         Left            =   -71160
         TabIndex        =   78
         Top             =   1185
         Width           =   1515
      End
      Begin VB.ComboBox Combo6 
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
         ItemData        =   "frmBuscarGeneral.frx":1887
         Left            =   1380
         List            =   "frmBuscarGeneral.frx":1891
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   585
         Width           =   675
      End
      Begin VB.ComboBox cboFecha_Desde 
         Height          =   345
         ItemData        =   "frmBuscarGeneral.frx":189F
         Left            =   1380
         List            =   "frmBuscarGeneral.frx":18AC
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   1065
         Width           =   675
      End
      Begin VB.TextBox txtFechaHasta 
         Height          =   330
         Left            =   2460
         TabIndex        =   63
         Top             =   1905
         Width           =   1035
      End
      Begin VB.TextBox txtFechaDesde 
         Height          =   330
         Left            =   2100
         TabIndex        =   62
         ToolTipText     =   "Se puede buscar por fecha Ej 01/01/2008 o por año 2008"
         Top             =   1065
         Width           =   1155
      End
      Begin VB.TextBox txtFecha_Entre 
         Height          =   330
         Left            =   1380
         TabIndex        =   61
         Top             =   1485
         Width           =   1035
      End
      Begin VB.TextBox txtNroHasta 
         Height          =   330
         Left            =   5040
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Top             =   1905
         Width           =   1755
      End
      Begin VB.TextBox txtNroDesde 
         Height          =   330
         Left            =   5040
         ScrollBars      =   2  'Vertical
         TabIndex        =   59
         Top             =   1065
         Width           =   1755
      End
      Begin VB.TextBox txtNro_Entre 
         Height          =   330
         Left            =   5040
         ScrollBars      =   2  'Vertical
         TabIndex        =   58
         Top             =   1485
         Width           =   1755
      End
      Begin VB.TextBox txtLetraHasta 
         Height          =   330
         Left            =   9360
         TabIndex        =   57
         Top             =   1905
         Width           =   1335
      End
      Begin VB.TextBox txtLetraDesde 
         Height          =   330
         Left            =   9360
         TabIndex        =   56
         Top             =   1065
         Width           =   1335
      End
      Begin VB.TextBox txtLetra_Entre 
         Height          =   330
         Left            =   8640
         TabIndex        =   55
         Top             =   1485
         Width           =   2055
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   330
         Left            =   2160
         TabIndex        =   54
         Top             =   585
         Width           =   8355
      End
      Begin VB.ComboBox cboFecha_Hasta 
         Height          =   345
         ItemData        =   "frmBuscarGeneral.frx":18B9
         Left            =   1380
         List            =   "frmBuscarGeneral.frx":18C6
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1905
         Width           =   675
      End
      Begin VB.ComboBox Combo4 
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
         ItemData        =   "frmBuscarGeneral.frx":18D3
         Left            =   8640
         List            =   "frmBuscarGeneral.frx":18DD
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1065
         Width           =   675
      End
      Begin VB.ComboBox Combo5 
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
         ItemData        =   "frmBuscarGeneral.frx":18EB
         Left            =   8640
         List            =   "frmBuscarGeneral.frx":18F5
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1905
         Width           =   675
      End
      Begin VB.TextBox txtEtiqueta 
         BackColor       =   &H00C0FFC0&
         Height          =   330
         Left            =   -73800
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   825
         Width           =   1515
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H00C0C0FF&
         Height          =   330
         Left            =   -73800
         TabIndex        =   46
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox txtEtiquetaHasta 
         BackColor       =   &H00C0FFC0&
         Height          =   330
         Left            =   -71160
         TabIndex        =   45
         Top             =   825
         Width           =   1515
      End
      Begin Controles.cltGenerico ctlCambioCliente 
         Height          =   375
         Left            =   -73440
         TabIndex        =   98
         Top             =   705
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   661
      End
      Begin VB.Label Label20 
         Caption         =   "Entre:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   121
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Legajos Padre:"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   117
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Legajos Hijos:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   116
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Orden Hijos:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   115
         Top             =   1500
         Width           =   1035
      End
      Begin VB.Label Label16 
         Caption         =   "Imagenes Hijos:"
         Height          =   315
         Left            =   -74880
         TabIndex        =   114
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label14 
         Caption         =   "Orden"
         Height          =   255
         Left            =   -68040
         TabIndex        =   102
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Nuevo Cliente"
         Height          =   255
         Left            =   -74820
         TabIndex        =   99
         Top             =   705
         Width           =   1215
      End
      Begin VB.Label lblCambioIndiceID 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -66540
         TabIndex        =   95
         Top             =   1245
         Width           =   1095
      End
      Begin VB.Label lblCambioIndice 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -68460
         TabIndex        =   94
         Top             =   1245
         Width           =   1815
      End
      Begin VB.Label lblCambioDescripcion 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -73620
         TabIndex        =   93
         Top             =   1245
         Width           =   5115
      End
      Begin VB.Label Label10 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   -71760
         TabIndex        =   79
         Top             =   1245
         Width           =   615
      End
      Begin VB.Label lblTituloFechaHasta 
         Caption         =   "Fecha desde :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   73
         Top             =   1965
         Width           =   1095
      End
      Begin VB.Label lblTituloFechaDesde 
         Caption         =   "Fecha desde :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   1125
         Width           =   1155
      End
      Begin VB.Label lblTituloNumeroHasta 
         Caption         =   "Hasta Numero"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   71
         Top             =   1845
         Width           =   1215
      End
      Begin VB.Label lblTituloNumeroDesde 
         Caption         =   "Desde Numero:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   70
         Top             =   1125
         Width           =   1395
      End
      Begin VB.Label lblTituloLetraHasta 
         Caption         =   "Letra Hasta:"
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
         Left            =   7320
         TabIndex        =   69
         Top             =   1965
         Width           =   1335
      End
      Begin VB.Label lblTituloLetraDesde 
         Caption         =   "Letra Desde:"
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
         Left            =   7320
         TabIndex        =   68
         Top             =   1125
         Width           =   1275
      End
      Begin VB.Label lblTituloDescripcion 
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   67
         Top             =   645
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Entre:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   66
         Top             =   1545
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Entre:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -71100
         TabIndex        =   65
         Top             =   1485
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Entre:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   64
         Top             =   1605
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Etiqueta:"
         Height          =   255
         Left            =   -74580
         TabIndex        =   50
         Top             =   885
         Width           =   735
      End
      Begin VB.Label LblCaja 
         Caption         =   "Caja:"
         Height          =   255
         Index           =   0
         Left            =   -74520
         TabIndex        =   49
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta"
         Height          =   255
         Left            =   -71760
         TabIndex        =   48
         Top             =   825
         Width           =   675
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      Left            =   9000
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscarDocumento 
      Caption         =   "F12"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5580
      TabIndex        =   23
      Top             =   120
      Width           =   675
   End
   Begin VB.TextBox txtIndice_Nro_Documento 
      BackColor       =   &H00C0E0FF&
      Height          =   330
      Left            =   4740
      TabIndex        =   22
      Top             =   120
      Width           =   675
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7395
      Left            =   60
      TabIndex        =   6
      Top             =   3060
      Width           =   14835
      Begin TabDlg.SSTab SSTab1 
         Height          =   3255
         Left            =   120
         TabIndex        =   7
         Top             =   4020
         Width           =   14535
         _ExtentX        =   25638
         _ExtentY        =   5741
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   406
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Legajos"
         TabPicture(0)   =   "frmBuscarGeneral.frx":1903
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdSeleccionLegajos"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdPasarTodosLegajos"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdLimpiarLegajos"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdBuscarLegajos"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdRotulos"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdBuscarRearchivo"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdBorrar"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdEntrada"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdLecturaLegajo"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmdcopiarExcel"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Búsqueda"
         TabPicture(1)   =   "frmBuscarGeneral.frx":191F
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdPasarTodos"
         Tab(1).Control(1)=   "cmdInsertarBusqueda"
         Tab(1).Control(2)=   "cmdReporteBusqueda"
         Tab(1).Control(3)=   "cmdLimpiar"
         Tab(1).Control(4)=   "grdVarios"
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Detalles"
         TabPicture(2)   =   "frmBuscarGeneral.frx":193B
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label6"
         Tab(2).Control(1)=   "lblDetalleCaja"
         Tab(2).Control(2)=   "Label7"
         Tab(2).Control(3)=   "lblDetalleIndiceDescripcion"
         Tab(2).Control(4)=   "Label9"
         Tab(2).Control(5)=   "lblDetalle_fecha_desde"
         Tab(2).Control(6)=   "Label11"
         Tab(2).Control(7)=   "lblDetalle_fecha_hasta"
         Tab(2).Control(8)=   "Label13"
         Tab(2).Control(9)=   "lblDetalle_nro_desde"
         Tab(2).Control(10)=   "Label15"
         Tab(2).Control(11)=   "lblDetalle_Nro_Hasta"
         Tab(2).Control(12)=   "lbl1(0)"
         Tab(2).Control(13)=   "lblDetalle_Letra_desde"
         Tab(2).Control(14)=   "lbl"
         Tab(2).Control(15)=   "lblDetalle_Letra_hasta"
         Tab(2).Control(16)=   "Label17(1)"
         Tab(2).Control(17)=   "lblDetalle_descripcion"
         Tab(2).Control(18)=   "lblDetalleLote"
         Tab(2).Control(19)=   "lbl1(1)"
         Tab(2).ControlCount=   20
         TabCaption(3)   =   "Imágenes"
         TabPicture(3)   =   "frmBuscarGeneral.frx":1957
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label8"
         Tab(3).Control(1)=   "ctlVerImagenes1"
         Tab(3).ControlCount=   2
         Begin VB.CommandButton cmdcopiarExcel 
            Caption         =   "Copiar Excel"
            Height          =   315
            Left            =   6240
            TabIndex        =   105
            Top             =   2400
            Width           =   1395
         End
         Begin Controles.ctlVerImagenes ctlVerImagenes1 
            Height          =   2295
            Left            =   -74760
            TabIndex        =   88
            Top             =   540
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   4048
         End
         Begin VB.CommandButton cmdPasarTodos 
            Caption         =   "Pasar todos"
            Height          =   345
            Left            =   -70140
            TabIndex        =   85
            Top             =   2820
            Width           =   1400
         End
         Begin VB.CommandButton cmdLecturaLegajo 
            Caption         =   "Lectura "
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
            Left            =   300
            TabIndex        =   18
            Top             =   2400
            Width           =   1335
         End
         Begin VB.CommandButton cmdEntrada 
            Caption         =   "Entrada"
            Height          =   315
            Left            =   1740
            TabIndex        =   17
            Top             =   2400
            Width           =   1395
         End
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "Borrar"
            Height          =   315
            Left            =   3240
            TabIndex        =   16
            Top             =   2400
            Width           =   1395
         End
         Begin VB.CommandButton cmdBuscarRearchivo 
            Caption         =   "Buscar Rear."
            Height          =   315
            Left            =   4740
            TabIndex        =   15
            Top             =   2400
            Width           =   1395
         End
         Begin VB.CommandButton cmdRotulos 
            Caption         =   "Rotulos"
            Height          =   315
            Left            =   4740
            TabIndex        =   14
            Top             =   2820
            Width           =   1395
         End
         Begin VB.CommandButton cmdBuscarLegajos 
            Caption         =   "Buscar"
            Height          =   315
            Left            =   3240
            TabIndex        =   13
            Top             =   2820
            Width           =   1395
         End
         Begin VB.CommandButton cmdLimpiarLegajos 
            Caption         =   "Limpiar"
            Height          =   315
            Left            =   1740
            TabIndex        =   12
            Top             =   2820
            Width           =   1395
         End
         Begin VB.CommandButton cmdPasarTodosLegajos 
            Caption         =   "Pasar"
            Height          =   315
            Left            =   300
            TabIndex        =   11
            Top             =   2820
            Width           =   1335
         End
         Begin VB.CommandButton cmdInsertarBusqueda 
            Caption         =   "Busqueda"
            Height          =   345
            Left            =   -74760
            TabIndex        =   10
            Top             =   2820
            Width           =   1400
         End
         Begin VB.CommandButton cmdReporteBusqueda 
            Caption         =   "Reporte"
            Height          =   345
            Left            =   -73260
            TabIndex        =   9
            Top             =   2820
            Width           =   1400
         End
         Begin VB.CommandButton cmdLimpiar 
            Caption         =   "Limpiar"
            Height          =   345
            Left            =   -71700
            TabIndex        =   8
            Top             =   2820
            Width           =   1400
         End
         Begin MSFlexGridLib.MSFlexGrid grdSeleccionLegajos 
            Height          =   1935
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Width           =   13935
            _ExtentX        =   24580
            _ExtentY        =   3413
            _Version        =   393216
            BackColorSel    =   12632064
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid grdVarios 
            Height          =   2055
            Left            =   -74820
            TabIndex        =   20
            Top             =   540
            Width           =   12795
            _ExtentX        =   22569
            _ExtentY        =   3625
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   12640511
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbl1 
            Caption         =   "Lote:"
            Height          =   255
            Index           =   1
            Left            =   -74820
            TabIndex        =   77
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label lblDetalleLote 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   -73740
            TabIndex        =   76
            Top             =   2640
            Width           =   10695
         End
         Begin VB.Label Label8 
            Caption         =   "En desarrollo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   -72060
            TabIndex        =   43
            Top             =   660
            Width           =   6915
         End
         Begin VB.Label lblDetalle_descripcion 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   -73740
            TabIndex        =   42
            Top             =   1800
            Width           =   10695
         End
         Begin VB.Label Label17 
            Caption         =   "Descripcion:"
            Height          =   315
            Index           =   1
            Left            =   -74880
            TabIndex        =   41
            Top             =   1860
            Width           =   1095
         End
         Begin VB.Label lblDetalle_Letra_hasta 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   -67620
            TabIndex        =   40
            Top             =   1320
            Width           =   4575
         End
         Begin VB.Label lbl 
            Caption         =   "Letra Hasta:"
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
            Left            =   -68640
            TabIndex        =   39
            Top             =   1380
            Width           =   1395
         End
         Begin VB.Label lblDetalle_Letra_desde 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   -73740
            TabIndex        =   38
            Top             =   1320
            Width           =   4575
         End
         Begin VB.Label lbl1 
            Caption         =   "Letra desde:"
            Height          =   315
            Index           =   0
            Left            =   -74880
            TabIndex        =   37
            Top             =   1380
            Width           =   1155
         End
         Begin VB.Label lblDetalle_Nro_Hasta 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   -64680
            TabIndex        =   36
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label Label15 
            Caption         =   "Nº Hasta:"
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
            Left            =   -65640
            TabIndex        =   35
            Top             =   900
            Width           =   795
         End
         Begin VB.Label lblDetalle_nro_desde 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   -67620
            TabIndex        =   34
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label Label13 
            Caption         =   "Nº Desde:"
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
            Left            =   -68700
            TabIndex        =   33
            Top             =   900
            Width           =   1155
         End
         Begin VB.Label lblDetalle_fecha_hasta 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   -70800
            TabIndex        =   32
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha Desde:"
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
            Left            =   -72000
            TabIndex        =   31
            Top             =   900
            Width           =   1155
         End
         Begin VB.Label lblDetalle_fecha_desde 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   -73740
            TabIndex        =   30
            Top             =   840
            Width           =   1635
         End
         Begin VB.Label Label9 
            Caption         =   "Fecha Desde:"
            Height          =   315
            Left            =   -74880
            TabIndex        =   29
            Top             =   900
            Width           =   1155
         End
         Begin VB.Label lblDetalleIndiceDescripcion 
            BackColor       =   &H000080FF&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   -70800
            TabIndex        =   28
            Top             =   360
            Width           =   5115
         End
         Begin VB.Label Label7 
            Caption         =   "Documento"
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
            Left            =   -71880
            TabIndex        =   27
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblDetalleCaja 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   -73740
            TabIndex        =   26
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label6 
            Caption         =   "Caja:"
            Height          =   315
            Left            =   -74880
            TabIndex        =   25
            Top             =   420
            Width           =   615
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdResultadoBusqueda 
         Height          =   3615
         Left            =   180
         TabIndex        =   21
         Top             =   240
         Width           =   14475
         _ExtentX        =   25532
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   8
         BackColorFixed  =   -2147483638
         BackColorSel    =   8454016
         ForeColorSel    =   -2147483642
         AllowBigSelection=   0   'False
         SelectionMode   =   1
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label lblIndice_Descripcion 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   7500
      TabIndex        =   109
      Top             =   120
      Width           =   3015
   End
   Begin VB.Menu mnuImagenes 
      Caption         =   "Imagenes"
      Begin VB.Menu mnuImagenesver 
         Caption         =   "Ver Imagen"
      End
      Begin VB.Menu mnuCopiarPaso 
         Caption         =   "Copiar paso"
      End
      Begin VB.Menu Impresosi 
         Caption         =   "Impreso SI"
      End
      Begin VB.Menu mnuCopiarImagenC 
         Caption         =   "Copiar Imagen al c"
      End
      Begin VB.Menu Impresono 
         Caption         =   "Impreso No"
      End
      Begin VB.Menu mnuCarpetaSincronizada 
         Caption         =   "Copiar Imagen a carpeta sincronizada "
      End
   End
   Begin VB.Menu mnuLegajos 
      Caption         =   "Legajos"
      Begin VB.Menu mnuModificar_Legajos 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnuBorrarEtiquetaVirtual 
         Caption         =   "Borrar Etiqueta Virtural"
      End
   End
   Begin VB.Menu mnuVarios 
      Caption         =   "Varios"
      Visible         =   0   'False
      Begin VB.Menu mnuBorrarVarios 
         Caption         =   "Borrar"
      End
   End
   Begin VB.Menu mnuReferencias 
      Caption         =   "Refrerencias"
      Begin VB.Menu mnuReferenciasNuevasCopiar 
         Caption         =   "Referencias Nuevas Copiar"
      End
      Begin VB.Menu mnuReferenciasModificar 
         Caption         =   "Modificar referencias"
      End
      Begin VB.Menu mnuReferenciasNuevas 
         Caption         =   "Referencias Nuevas"
      End
      Begin VB.Menu mnuCrearLegajo 
         Caption         =   "Crear Legajo"
      End
      Begin VB.Menu mnuVerImagen 
         Caption         =   "Ver Imagen"
      End
   End
   Begin VB.Menu mnuUnificar 
      Caption         =   "Unificar"
      Visible         =   0   'False
      Begin VB.Menu mnuCargarPadre 
         Caption         =   "CargarPadre"
      End
      Begin VB.Menu mnuCargarHijos 
         Caption         =   "CargarHijos"
      End
   End
End
Attribute VB_Name = "frmBuscarGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CadenaBusqueda As String


Public Sub TitulosSeleccionLegajos()
    grdSeleccionLegajos.Cols = 5
    grdSeleccionLegajos.Rows = 1
    grdSeleccionLegajos.ColAlignment(1) = 0
    grdSeleccionLegajos.ColAlignment(2) = 0
    grdSeleccionLegajos.ColAlignment(3) = 0
    grdSeleccionLegajos.ColAlignment(4) = 0
    grdSeleccionLegajos.ColWidth(0) = 100
    grdSeleccionLegajos.ColWidth(1) = 2000
    grdSeleccionLegajos.ColWidth(2) = 2000
    grdSeleccionLegajos.ColWidth(3) = 2000
    grdSeleccionLegajos.ColWidth(4) = 2000
    grdSeleccionLegajos.TextMatrix(0, 1) = "Cliente"
    grdSeleccionLegajos.TextMatrix(0, 2) = "Etiqueta"
    grdSeleccionLegajos.TextMatrix(0, 3) = "Legajo Cliente"
    grdSeleccionLegajos.TextMatrix(0, 4) = "Caja"
Rem grdSeleccionLegajos.
End Sub

Private Sub cmdActualizarlegajos_Click()
    If MDIfrmInicio.StaInicio.Panels(2).Text = 19 Or MDIfrmInicio.StaInicio.Panels(2).Text = 48 Then
        If InputBox("Ingrese la Password") = "1907" Then
            frmActualizacion_legajos.Show
        End If
    Else
        MsgBox "Usuario No Permitido", vbCritical
    End If
End Sub

Private Sub cmdBorrarReferencia_Click()
Dim i As Integer
Dim Sql As String

For i = 0 To grdResultadoBusqueda.Rows - 1
    If grdResultadoBusqueda.TextMatrix(i, 1) = "Referencias" Then
        Sql = "  DELETE FROM basasql.dbo.REFERENCIAS"
        Sql = Sql & "  Where COD_ID_REFERENCIA = " & grdResultadoBusqueda.TextMatrix(i, 0)
        ExecutarSql Sql
    End If
Next

MsgBox "Terminado"
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo salir
    Dim rsbusqueda As ADODB.Recordset
    Dim strIndice As String
    Dim Des_indice As String
    Dim Sql As String
    Dim REARCHIVO_CAJA As String
    




        TituloGrilla
        If chkLegajos = 0 And chkRearchivo = 0 And ChkRearchivoDigital.value = 0 And chkReferencias.value = 0 Then
            MsgBox "Selecione una busqueda"
            Exit Sub
        End If
            If chkLegajos.value = 1 Then
                Dim DesVarioslegajos As String
                Dim INdice_Des As String
                Sql = "   SELECT  ID_LEGAJO,  LEGAJOS.COD_CLIENTE,  LEGAJOS.ID_CLIENTE_LEGAJO , LEGAJOS.NRO_CAJA , LEGAJOS.COD_ESTADO,"
                Sql = Sql & vbCrLf & " LEGAJOS.NRO_DESDE, LEGAJOS.NRO_HASTA, LEGAJOS.LETRA_DESDE, LEGAJOS.LETRA_HASTA, LEGAJOS.FECHA_DESDE,LEGAJOS.PASOARCHIVO , "
                Sql = Sql & vbCrLf & " LEGAJOS.FECHA_HASTA, LEGAJOS.DESCRIPCION, INDICES.DESCRIPCION AS DES_INDICE,CLIENTE_LEGAJO , DESCRIPCION_REMITO , FECHA_CREACION ,  REARCHIVO_CAJA  "
                Sql = Sql & vbCrLf & " FROM LEGAJOS LEFT OUTER JOIN INDICES ON LEGAJOS.FK_INDICES = INDICES.ID"
                If txtCaja.Text <> "" Then
                    If txtCajaHasta.Text <> "" Then
                        Sql = Sql & vbCrLf & " where NRO_CAJA   between  " & txtCaja.Text & " and " & txtCajaHasta.Text & " AND LEGAJOS.COD_CLIENTE =  " & ctlCliente.Valor
                    Else
                        Sql = Sql & vbCrLf & " where NRO_CAJA  in(  " & txtCaja.Text & "  ) AND LEGAJOS.COD_CLIENTE =  " & ctlCliente.Valor
                    End If
                Else
                    If txtEtiqueta.Text = "" Then
                        Sql = Sql & vbCrLf & " WHERE LEGAJOS.COD_CLIENTE = " & ctlCliente.Valor
                    If txtIndice_Nro_Documento.Text <> "" Then
                        Sql = Sql & vbCrLf & "  and  INDICES.INDICE LIKE '" & BuscarIndiceDocumento_Indice(txtIndice_Nro_Documento.Text, ctlCliente.Valor) & "%'"
                    End If
                    Sql = Sql & vbCrLf & Replace(ArmarConsulta, "DESCRIPCION", "LEGAJOS.DESCRIPCION")
                    
                    If txtLegajoEstado.Text <> "" Then
                            Sql = Sql & vbCrLf & " AND ( " & txtLegajoEstado.Text & ")"
                           txtLegajoEstado.Text = ""
                End If
                    
                    
                    If ArmarConsulta = "" And txtLegajoEstado.Text = "" Then
                        If MsgBox("Solo se traeran los primeros 3000 registros  ¿Usted desea Continuar? ", vbCritical + vbYesNo) = vbYes Then
                            Sql = Replace(Sql, "SELECT", "SELECT TOP 3000 ")
                        Else
                            Exit Sub
                        End If
                     End If
                Else
                
                
                
                    If Trim(txtEtiquetaHasta.Text) <> "" Then
                        Sql = Sql & vbCrLf & " WHERE "
                        Sql = Sql & vbCrLf & " ID_CLIENTE_LEGAJO BETWEEN  " & txtEtiqueta.Text & " and " & txtEtiquetaHasta.Text
                        Sql = Sql & vbCrLf & " ORDER BY ID_LEGAJO"
                    Else
                        Sql = Sql & vbCrLf & " WHERE "
                        Sql = Sql & vbCrLf & " ID_CLIENTE_LEGAJO IN( " & txtEtiqueta.Text & ")"
                        If Not IsNull(ctlCliente.Valor) Then
                        Sql = Sql & vbCrLf & " AND LEGAJOS.COD_CLIENTE = " & ctlCliente.Valor
                        End If
                        
                        Sql = Sql & vbCrLf & " ORDER BY ID_LEGAJO"
                    End If
            
                
                End If
                End If
                
                
                Dim clienteAnterior As Integer
                Set rsbusqueda = New ADODB.Recordset
                Dim con As New ADODB.Connection
                con.Open strConBasa
                con.CommandTimeout = 6000
                
                rsbusqueda.Open Sql, con, adOpenStatic, adLockReadOnly
                Do While Not rsbusqueda.EOF
                If txtEtiqueta.Text <> "" Then
                If IsNull(rsbusqueda!COD_CLIENTE) Then
                MsgBox "LA ETIQUETA NO ESTA EN USO"
                Exit Sub
                End If
                If clienteAnterior = rsbusqueda!COD_CLIENTE Then
                Else
                clienteAnterior = rsbusqueda!COD_CLIENTE
                MsgBox "Cliente : " & rsbusqueda!COD_CLIENTE
                End If
                End If
                If IsNull(rsbusqueda!DESCRIPCION_REMITO) Then
                    If Not IsNull(rsbusqueda!CLIENTE_LEGAJO) Then
                    DesVarioslegajos = rsbusqueda!CLIENTE_LEGAJO
                    End If
                Else
                    DesVarioslegajos = rsbusqueda!DESCRIPCION_REMITO
                End If
                
                If IsNull(rsbusqueda!Des_indice) Then
                    Des_indice = "No tiene Indice"
                Else
                    Des_indice = rsbusqueda!Des_indice
                End If
                
                If IsNull(rsbusqueda!REARCHIVO_CAJA) Then
                   REARCHIVO_CAJA = ""
                Else
                     REARCHIVO_CAJA = rsbusqueda!REARCHIVO_CAJA
                End If
                
                CargarGrillaBusqueda rsbusqueda!ID_LEGAJO, "Legajos", Replace(Des_indice, "/", " ** "), rsbusqueda!ID_CLIENTE_LEGAJO, rsbusqueda!Cod_Estado, rsbusqueda!NRO_CAJA, "", rsbusqueda!NRO_DESDE, rsbusqueda!NRO_HASTA, rsbusqueda!FECHA_DESDE, rsbusqueda!FECHA_HASTA, rsbusqueda!LETRA_DESDE, rsbusqueda!LETRA_HASTA, rsbusqueda!Descripcion, DesVarioslegajos, rsbusqueda!FECHA_CREACION, rsbusqueda!PASOARCHIVO, REARCHIVO_CAJA
                rsbusqueda.MoveNext
                Loop
                txtEtiqueta.Text = ""
            End If







'' Referencia

        If chkReferencias.value = 1 Then


                    Sql = "  SELECT INDICES.DESCRIPCION as DES_INDICE, COD_ID_REFERENCIA , REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA, REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA, "
                    Sql = Sql & vbCrLf & " REFERENCIAS.LETRA_DESDE, REFERENCIAS.LETRA_HASTA, REFERENCIAS.DESCRIPCION , REFERENCIAS.NRO_CAJA, "
                    Sql = Sql & vbCrLf & " REFERENCIAS.cod_cliente , REFERENCIAS.Indice ,REFERENCIAS.FECHA_MODIFICACION , REFERENCIAS.PASOARCHIVO "
                    Sql = Sql & vbCrLf & " FROM REFERENCIAS INNER JOIN "
                    Sql = Sql & vbCrLf & " INDICES ON REFERENCIAS.INDICE = INDICES.INDICE AND REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE"
                    Sql = Sql & vbCrLf & " WHERE REFERENCIAS.COD_CLIENTE = " & ctlCliente.Valor
                    
                    If txtIndice_Nro_Documento.Text = "" Then
                        MsgBox "La busqueda de la referencia no es la adecuada ", vbCritical
                    Else
                        Sql = Sql & vbCrLf & " AND REFERENCIAS.INDICE LIKE '" & BuscarIndiceDocumento_Indice(txtIndice_Nro_Documento.Text, ctlCliente.Valor) & "%'"
                    End If
                    
                    If txtCaja.Text <> "" Then
                            If txtCajaHasta.Text <> "" Then
                                 Sql = Sql & vbCrLf & " AND  NRO_CAJA   between  " & txtCaja.Text & " and " & txtCajaHasta.Text
                             Else
                                 Sql = Sql & vbCrLf & " AND NRO_CAJA  in(  " & txtCaja.Text & ")"
                             End If
                    Else
                        If ArmarConsulta = "" Then
                            
                            If txtDocumento_fecha <> "" Then
                                Sql = Sql & " AND " & txtDocumento_fecha
                                txtDocumento_fecha.Text = ""
                            Else
                            
'                            If MsgBox("Solo se traeran los primeros 3000 registros  ¿Usted desea Continuar? ", vbCritical + vbYesNo) = vbYes Then
'                                SQL = Replace(SQL, "SELECT", "SELECT TOP 3000 ")
'                            Else
'                                Exit Sub
'                            End If
                          End If
                        Else
                             Sql = Sql & vbCrLf & Replace(ArmarConsulta, "DESCRIPCION", "REFERENCIAS.DESCRIPCION")
                        End If
                    
                    End If
                                       
                    Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.NRO_DESDE DESC, REFERENCIAS.FECHA_DESDE DESC"
                    Set rsbusqueda = New ADODB.Recordset
                        rsbusqueda.Open Sql, ConActiva, 0, 1
                        Do While Not rsbusqueda.EOF
                            CargarGrillaBusqueda rsbusqueda!COD_ID_REFERENCIA, "Referencias", rsbusqueda!Des_indice, "", 2, rsbusqueda!NRO_CAJA, "", rsbusqueda!NRO_DESDE, rsbusqueda!NRO_HASTA, rsbusqueda!FECHA_DESDE, rsbusqueda!FECHA_HASTA, rsbusqueda!LETRA_DESDE, rsbusqueda!LETRA_HASTA, rsbusqueda!Descripcion, "", rsbusqueda!FECHA_MODIFICACION, rsbusqueda!PASOARCHIVO
                            rsbusqueda.MoveNext
                        Loop
                       
         End If
                        

        If chkRearchivo.value = 1 Then


                        Sql = "  SELECT    ORDENAR_DOCUMENTACION_DETALLE.ID , ORDENAR_DOCUMENTACION_DETALLE.COD_DOCUMENTACION, ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE,"
                        Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.COD_INDICE, ORDENAR_DOCUMENTACION_DETALLE.COD_NRO_CAJA AS NRO_CAJA,"
                        Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.NRO_DESDE, ORDENAR_DOCUMENTACION_DETALLE.NRO_HASTA,"
                        Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.LETRA_DESDE, ORDENAR_DOCUMENTACION_DETALLE.LETRA_HASTA,"
                        Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.FECHA_DESDE, ORDENAR_DOCUMENTACION_DETALLE.FECHA_HASTA,"
                        Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.COD_TIPO_ORDEN, ORDENAR_DOCUMENTACION_DETALLE.COD_ESTADO, ORDENAR_DOCUMENTACION_DETALLE.CONTENEDOR_PROV, "
                        Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.DESCRIPCION, INDICES.DESCRIPCION AS DES_INDICE"
                        Sql = Sql & vbCrLf & " FROM ORDENAR_DOCUMENTACION_DETALLE LEFT OUTER JOIN"
                        Sql = Sql & vbCrLf & " INDICES ON ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = INDICES.COD_CLIENTE AND"
                        Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Indice = INDICES.Indice"
                        
                        Sql = Sql & vbCrLf & " WHERE ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = " & ctlCliente.Valor
                        If txtCaja.Text <> "" Then
                             If txtCajaHasta.Text <> "" Then
                                 Sql = Sql & vbCrLf & " AND ORDENAR_DOCUMENTACION_DETALLE.COD_NRO_CAJA   between  " & txtCaja.Text & " and " & txtCajaHasta.Text
                             Else
                                 Sql = Sql & vbCrLf & " AND ORDENAR_DOCUMENTACION_DETALLE.COD_NRO_CAJA  =  " & txtCaja.Text
                             End If
                        Else
                          Sql = Sql & vbCrLf & Replace(ArmarConsulta, "DESCRIPCION", "dbo.ORDENAR_DOCUMENTACION_DETALLE.DESCRIPCION")
                        End If
                        
                      
                        
                        Sql = Sql & vbCrLf & " ORDER BY NRO_DESDE, FECHA_DESDE"
                        
                        
                         Set rsbusqueda = New ADODB.Recordset
                            rsbusqueda.Open Sql, ConActiva, 0, 1
                            Do While Not rsbusqueda.EOF
                                 CargarGrillaBusqueda rsbusqueda!ID, "Rearchivo", rsbusqueda!Des_indice, "Cod_doc: " & Trim(rsbusqueda!COD_DOCUMENTACION), rsbusqueda!Cod_Estado, rsbusqueda!NRO_CAJA, "", rsbusqueda!NRO_DESDE, rsbusqueda!NRO_HASTA, rsbusqueda!FECHA_DESDE, rsbusqueda!FECHA_HASTA, rsbusqueda!LETRA_DESDE, rsbusqueda!LETRA_HASTA, rsbusqueda!Descripcion, "PROV:" & rsbusqueda!Contenedor_Prov, "", ""
                                rsbusqueda.MoveNext
                            Loop
            End If
            
            
            Dim Impreso As String
            
              If ChkRearchivoDigital.value = 1 Then


                    
                        Sql = "  SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS AS NRO_CAJA, DOCUMENTOS_DIGITALES.LETRA_DESDE,"
                        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES.LETRA_HASTA, DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.NRO_HASTA,"
                        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.FECHA_HASTA,"
                        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION AS DESC_LOTE, DOCUMENTOS_DIGITALES.COD_ESTADO,"
                        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN , DOCUMENTOS_DIGITALES.DESCRIPCION , DOCUMENTOS_DIGITALES.IMPRESO, DOCUMENTOS_DIGITALES.Lote"
                        Sql = Sql & vbCrLf & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
                        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES_LOTE ON"
                        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
                        Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = " & ctlCliente.Valor
                        Sql = Sql & vbCrLf & ArmarConsulta
                    
                    
                  Sql = "  SELECT     dbo.DOCUMENTOS_DIGITALES.ID, dbo.DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS AS NRO_CAJA, dbo.DOCUMENTOS_DIGITALES.LETRA_DESDE,"
                  Sql = Sql & vbCrLf & "     dbo.DOCUMENTOS_DIGITALES.LETRA_HASTA, dbo.DOCUMENTOS_DIGITALES.NRO_DESDE, dbo.DOCUMENTOS_DIGITALES.NRO_HASTA,"
                   Sql = Sql & vbCrLf & "    dbo.DOCUMENTOS_DIGITALES.FECHA_DESDE, dbo.DOCUMENTOS_DIGITALES.FECHA_HASTA,"
                  Sql = Sql & vbCrLf & "     dbo.DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION AS DESC_LOTE, dbo.DOCUMENTOS_DIGITALES.COD_ESTADO, FK_ID_LEGAJO ,"
                   Sql = Sql & vbCrLf & "    dbo.DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, dbo.DOCUMENTOS_DIGITALES.DESCRIPCION, dbo.DOCUMENTOS_DIGITALES.IMPRESO,"
                   Sql = Sql & vbCrLf & "    dbo.DOCUMENTOS_DIGITALES.LOTE, dbo.INDICES.DESCRIPCION AS Des_Indice "
Sql = Sql & vbCrLf & "  FROM         dbo.DOCUMENTOS_DIGITALES INNER JOIN"
                      Sql = Sql & vbCrLf & "  dbo.DOCUMENTOS_DIGITALES_LOTE ON"
                      Sql = Sql & vbCrLf & "  dbo.DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = dbo.DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE LEFT"
                       Sql = Sql & vbCrLf & "  Outer Join"
                      Sql = Sql & vbCrLf & "  dbo.INDICES ON dbo.DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = dbo.INDICES.ID"
Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = " & ctlCliente.Valor
                        Sql = Sql & vbCrLf & Replace(ArmarConsulta, "DESCRIPCION", "dbo.DOCUMENTOS_DIGITALES.DESCRIPCION")
                    
                    
                        Set rsbusqueda = New ADODB.Recordset
                        rsbusqueda.Open Sql, ConActiva, 0, 1
                        Do While Not rsbusqueda.EOF
                        If IsNull(rsbusqueda!FK_ID_LEGAJO) Then
                            Impreso = "impreso:" & rsbusqueda!Impreso & " LOTE: " & Trim(rsbusqueda!DESC_LOTE) & " POS:" & Trim(rsbusqueda!IMAGEN_ORIGEN)
                        
                        Else
                            Impreso = "Unificado Etiqueta " & rsbusqueda!FK_ID_LEGAJO
                        End If
                        
                             CargarGrillaBusqueda rsbusqueda!ID, "Digital", Trim(rsbusqueda!Des_indice), rsbusqueda!ID, rsbusqueda!Cod_Estado, rsbusqueda!NRO_CAJA, Impreso, rsbusqueda!NRO_DESDE, rsbusqueda!NRO_HASTA, rsbusqueda!FECHA_DESDE, rsbusqueda!FECHA_HASTA, rsbusqueda!LETRA_DESDE, rsbusqueda!LETRA_HASTA, rsbusqueda!Descripcion, "Lote: " & rsbusqueda!lote & " Pos:" & rsbusqueda!IMAGEN_ORIGEN, "", ""
                             rsbusqueda.MoveNext
                        Loop
                    
                End If
                txtNroDesde.Text = ""
  txtNro_Entre = ""
    Exit Sub
salir:
txtNroDesde.Text = ""
txtNro_Entre.Text = ""
txtEtiqueta.Text = ""
txtLegajoEstado.Text = ""
MsgBox Err.Description
    MsgBox "Verifique los datos ingresados son incorrectos", vbCritical
    
End Sub

Public Function ArmarConsulta() As String
Dim Sql As String
CadenaBusqueda = ""

If txtFechaDesde.Text <> "" Then
   CadenaBusqueda = "Fecha desde: " & txtFechaDesde.Text
    If Len(txtFechaDesde.Text) = 4 Then
        Sql = Sql & " AND  YEAR(FECHA_DESDE) = " & txtFechaDesde.Text
    Else
        Sql = Sql & " AND  FECHA_DESDE = " & FechaFormato(txtFechaDesde.Text)
    End If
End If

If txtFechaHasta.Text <> "" Then
     CadenaBusqueda = CadenaBusqueda & " Fecha Hasta " & txtFechaHasta.Text
    If Len(txtFechaHasta.Text) = 4 Then
        Sql = Sql & " AND  YEAR(FECHA_HASTA) = " & txtFechaHasta.Text
    Else
        Sql = Sql & " AND  FECHA_HASTA = '" & txtFechaHasta.Text & "'"
    End If
End If

If txtFecha_Entre.Text <> "" Then
    CadenaBusqueda = CadenaBusqueda & " Fecha " & txtFecha_Entre.Text
     Sql = Sql & " AND  (" & FechaFormato(txtFecha_Entre.Text) & " BETWEEN FECHA_DESDE AND FECHA_HASTA)"
End If



If txtNroDesde.Text <> "" Then
    CadenaBusqueda = CadenaBusqueda & " Nro_desde " & txtNroDesde.Text
'    If cboNro_Desde.Text = "CONTIENE" Then
'      Sql = Sql & "  AND ( CONVERT(CHAR, NRO_DESDE) LIKE '%" & txtNroDesde.Text & "%')"
'
'    Else
'
'   End If
    Sql = Sql & " AND  NRO_DESDE in( " & txtNroDesde.Text & ")"
End If

If txtNroHasta.Text <> "" Then
    CadenaBusqueda = CadenaBusqueda & "  Nro_Hasta " & txtNroHasta.Text
    Sql = Sql & " AND  NRO_HASTA in( " & txtNroHasta.Text & ")"
End If
Dim Coma As Integer
Dim Coma_Pos As Integer
Dim Coma_Inicio As Integer
Dim Coma_Caja As String
Dim Coma_i As Integer
            If txtNro_Entre.Text <> "" Then
                Coma_Inicio = 1
                Coma_Caja = ""
                  If InStr(1, txtNro_Entre.Text, ",") <> 0 Then
                         CadenaBusqueda = txtNro_Entre.Text
                        For Coma_i = 1 To Len(txtNro_Entre.Text)
                            If Mid(txtNro_Entre.Text, Coma_i, 1) = "," Then
                            Sql = Sql & " OR  (" & Coma_Caja & "  BETWEEN NRO_DESDE AND NRO_HASTA )"
                            Coma_Caja = ""
                            
                            Else
                                Coma_Caja = Coma_Caja & Mid(txtNro_Entre.Text, Coma_i, 1)
                            
                            End If
                            
                        Next
                        Sql = Sql & " OR  (" & Coma_Caja & "  BETWEEN NRO_DESDE AND NRO_HASTA )"
                      Sql = " AND ( " & Mid(Sql, 4) & ")"
                     
                Else
                      CadenaBusqueda = CadenaBusqueda & " Nro: " & txtNro_Entre.Text
                     Sql = Sql & " AND  (" & txtNro_Entre.Text & "  BETWEEN NRO_DESDE AND NRO_HASTA )"
                 End If
      
End If
If txtLetraDesde.Text <> "" Then
   CadenaBusqueda = CadenaBusqueda & "  Letra desde : " & txtLetraDesde.Text
   Sql = Sql & " AND  LETRA_DESDE LIKE '%" & txtLetraDesde.Text & "%'"
End If

If txtLetraHasta.Text <> "" Then
    CadenaBusqueda = CadenaBusqueda & "  LETRA_HASTA : " & txtLetraHasta.Text
   Sql = Sql & "  AND  LETRA_HASTA LIKE '%" & txtLetraHasta.Text & "%'"
End If

If txtLetra_Entre.Text <> "" Then
    CadenaBusqueda = CadenaBusqueda & "  Letra  " & txtLetra_Entre.Text
     Sql = Sql & " AND  ('" & txtLetra_Entre.Text & "' BETWEEN LETRA_DESDE AND LETRA_HASTA )"
End If

If txtDescripcion.Text <> "" Then
    CadenaBusqueda = CadenaBusqueda & "  Desc.:  " & txtDescripcion.Text
    Sql = Sql & " AND  DESCRIPCION LIKE '%" & txtDescripcion.Text & "%'"
End If
CadenaBusqueda = Trim(CadenaBusqueda)
ArmarConsulta = Sql

End Function

Private Sub cmdBuscarDocumento_Click()
On Error GoTo salir
frmIndice.COD_CLIENTE = ctlCliente.Valor
frmIndice.Actualizar ctlCliente.Valor, Nulo, 0
frmIndice.Show
 
salir:
End Sub

Public Sub Configurar_Carga(Cliente As Integer, Nro_documento As Long)
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim ColorHabilitado As ColorConstants
    Dim ColorDesaHabilitado As ColorConstants
    ColorHabilitado = &HC0FFFF
    ColorDesaHabilitado = &HC0C0FF

    Sql = "SELECT     COD_CLIENTE, ID_CODIGO_DOCUMENTO, TIPO_INDICE,DESCRIPCION, TITULO_FECHA_DESDE, TITULO_FECHA_HASTA, TITULO_LETRA_DESDE,"
    Sql = Sql & " TITULO_LETRA_HASTA, TITULO_NRO_DESDE, TITULO_NRO_HASTA, TITULO_DESCRIPCION, HABILITAR_FECHA_DESDE, HABILITAR_FECHA_HASTA,"
    Sql = Sql & " HABILITAR_LETRA_DESDE, HABILITAR_LETRA_HASTA, HABILITAR_NRO_DESDE, HABILITAR_NRO_HASTA, HABILITAR_DESCRIPCION,"
    Sql = Sql & " REQUERIR_FECHA_DESDE, REQUERIR_FECHA_HASTA, REQUERIR_LETRA_DESDE, REQUERIR_LETRA_HASTA, REQUERIR_NRO_DESDE,"
    Sql = Sql & " REQUERIR_NRO_HASTA , REQUERIR_DESCRIPCION, DESCRIPCION, COPIAR_FECHA, COPIAR_LETRA, COPIAR_NRO "
    Sql = Sql & "  From INDICES "
    Sql = Sql & "  Where COD_CLIENTE = " & Cliente
    Sql = Sql & "  And ID_CODIGO_DOCUMENTO = " & Nro_documento
    
    Set rs = New ADODB.Recordset
    
    rs.Open Sql, ConActiva, 0, 1

    
    
    If Not rs.EOF Then
        lblIndice_Descripcion.Caption = rs!Descripcion
    
       If IsNull(rs!TITULO_FECHA_DESDE) Then
            lblTituloFechaDesde.Caption = "Fecha Desde:"
       Else
            lblTituloFechaDesde.Caption = rs!TITULO_FECHA_DESDE
       End If
       
       If IsNull(rs!TITULO_FECHA_HASTA) Then
            lblTituloFechaHasta.Caption = "Fecha Hasta:"
       Else
            lblTituloFechaHasta.Caption = rs!TITULO_FECHA_HASTA
       End If
               
       If IsNull(rs!TITULO_LETRA_DESDE) Then
            lblTituloLetraDesde.Caption = "Letra Desde:"
       Else
           lblTituloLetraDesde.Caption = rs!TITULO_LETRA_DESDE
       End If
        
        If IsNull(rs!TITULO_LETRA_HASTA) Then
            lblTituloLetraHasta.Caption = "Letra desde:"
        Else
           lblTituloLetraHasta.Caption = rs!TITULO_LETRA_HASTA
        End If
         
         If IsNull(rs!TITULO_NRO_DESDE) Then
            lblTituloNumeroDesde.Caption = "Nro Desde:"
         Else
             lblTituloNumeroDesde.Caption = rs!TITULO_NRO_DESDE
         End If
         
         If IsNull(rs!TITULO_NRO_HASTA) Then
            lblTituloNumeroHasta.Caption = "Nro Hasta:"
         Else
            lblTituloNumeroHasta.Caption = rs!TITULO_NRO_HASTA
         End If
         
         If IsNull(rs!TITULO_DESCRIPCION) Then
            lblTituloDescripcion.Caption = "Descripción:"
         Else
            lblTituloDescripcion.Caption = rs!TITULO_DESCRIPCION
         End If
         
'       If IsNull(rs!HABILITAR_FECHA_DESDE) Then
'            txtFechaDesde.Enabled = False
'            txtFechaDesde.BackColor = ColorDesaHabilitado
'        Else
'            If rs!HABILITAR_FECHA_DESDE = True Then
'                txtFechaDesde.Enabled = True
'                txtFechaDesde.BackColor = ColorHabilitado
'            Else
'                txtFechaDesde.Enabled = False
'                txtFechaDesde.BackColor = ColorDesaHabilitado
'            End If
'       End If
'
'
'       If IsNull(rs!HABILITAR_FECHA_HASTA) Then
'            txtFechaHasta.Enabled = False
'            txtFechaHasta.BackColor = ColorDesaHabilitado
'       Else
'            If rs!HABILITAR_FECHA_HASTA = True Then
'                txtFechaHasta.Enabled = True
'                txtFechaHasta.BackColor = ColorHabilitado
'            Else
'                txtFechaHasta.Enabled = False
'                txtFechaHasta.BackColor = ColorDesaHabilitado
'            End If
'       End If
'
'
'
'       If IsNull(rs!HABILITAR_LETRA_DESDE) Then
'            txtLetraDesde.Enabled = False
'            txtLetraDesde.BackColor = ColorDesaHabilitado
'       Else
'            If rs!HABILITAR_LETRA_DESDE = True Then
'                txtLetraDesde.Enabled = True
'                txtLetraDesde.BackColor = ColorHabilitado
'            Else
'                txtLetraDesde.Enabled = False
'                txtLetraDesde.BackColor = ColorDesaHabilitado
'            End If
'       End If
'
'
'
'       If IsNull(rs!HABILITAR_LETRA_HASTA) Then
'            txtLetraHasta.Enabled = False
'            txtLetraHasta.BackColor = ColorDesaHabilitado
'       Else
'            If rs!HABILITAR_LETRA_HASTA = True Then
'                txtLetraHasta.Enabled = True
'                txtLetraHasta.BackColor = ColorHabilitado
'            Else
'                txtLetraHasta.Enabled = False
'                txtLetraHasta.BackColor = ColorDesaHabilitado
'            End If
'
'       End If
'
'       If IsNull(rs!HABILITAR_NRO_DESDE) Then
'            txtNroDesde.Enabled = False
'            txtNroDesde.BackColor = ColorDesaHabilitado
'       Else
'            If rs!HABILITAR_NRO_DESDE = True Then
'                txtNroDesde.Enabled = True
'                txtNroDesde.BackColor = ColorHabilitado
'            Else
'                 txtNroDesde.Enabled = False
'                txtNroDesde.BackColor = ColorDesaHabilitado
'            End If
'       End If
'
'      If IsNull(rs!HABILITAR_NRO_HASTA) Then
'           txtNroHasta.Enabled = False
'           txtNroHasta.BackColor = ColorDesaHabilitado
'      Else
'            If rs!HABILITAR_NRO_HASTA = True Then
'                txtNroHasta.Enabled = True
'                txtNroHasta.BackColor = ColorHabilitado
'             Else
'                txtNroHasta.Enabled = False
'                txtNroHasta.BackColor = ColorDesaHabilitado
'             End If
'      End If
'
'      If IsNull(rs!HABILITAR_DESCRIPCION) Then
'         txtDescripcion.Enabled = False
'         txtDescripcion.BackColor = ColorDesaHabilitado
'      Else
'            If rs!HABILITAR_DESCRIPCION = True Then
'                txtDescripcion.Enabled = True
'                txtDescripcion.BackColor = ColorHabilitado
'            Else
'                txtDescripcion.Enabled = False
'                txtDescripcion.BackColor = ColorDesaHabilitado
'            End If
'      End If
      
      Else
      LimpiarCampos
     
    End If
 
 End Sub


Private Sub ReporteBusqueda(Lote_Busqueda As Integer)
  Dim sSQL As String
    
    
        sSQL = " SELECT  ID_LOTE_BUSQUEDA, FECHA, TIPO, COD_CLIENTE, NRO_CAJA, LEGAJO, DESCRIPCION, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
        sSQL = sSQL & " UB_PROVISORIA, DETALLE "
        sSQL = sSQL & " FROM  V_TEM_BUSQUEDA "
        sSQL = sSQL & " where ID_LOTE_BUSQUEDA = " & Lote_Busqueda
      
       frmReportes.ImprimirReporte PasoReportes + "BusquedaGeneral.rpt", sSQL, True


End Sub

Private Sub cmdBuscarLegajos_Click()
 Dim sSQL As String
    Dim i As Integer
        
        
    sSQL = "  SELECT *  From V_BUSCAR_LEGAJOS "
    sSQL = sSQL & " Where  "
    sSQL = sSQL & vbCrLf & "  ( COD_CLIENTE =" & grdSeleccionLegajos.TextMatrix(1, 1) & " AND ID_CLIENTE_LEGAJO =" & grdSeleccionLegajos.TextMatrix(1, 2) & ")"
    For i = 2 To grdSeleccionLegajos.Rows - 1
        sSQL = sSQL & vbCrLf & "  OR ( COD_CLIENTE =" & grdSeleccionLegajos.TextMatrix(i, 1) & " AND ID_CLIENTE_LEGAJO =" & grdSeleccionLegajos.TextMatrix(i, 2) & "  )  "
    Next
    sSQL = sSQL & vbCrLf & " ORDER BY ESTANTERIA, Horizontal , Vertical"


'    For i = 1 To grdSeleccionLegajos.Rows - 1
'       If Not IsNull(ctlPersonal.Valor) Then
'            InsertarProducion ctlPersonal.Valor, 12, grdSeleccionLegajos.TextMatrix(i, 1) & " " & grdSeleccionLegajos.TextMatrix(i, 2), 1, grdSeleccionLegajos.TextMatrix(i, 1)
'       Else
'           MsgBox "Falta el responsable", vbInformation
'           Exit Sub
'       End If
'    Next
    frmReportes.ImprimirReporte PasoReportes + "rptBuscarLegajos.rpt", sSQL, True
End Sub

Private Sub cmdCambioBuscarIndice_Click()
On Error GoTo salir
frmIndice.COD_CLIENTE = ctlCambioCliente.Valor
frmIndice.Actualizar ctlCambioCliente.Valor, Nulo, 0
frmIndice.Show
 
salir:
End Sub

Private Sub cmdCambioIndice_Click()
Dim i As Integer
Dim Descripcion As String
Dim Sql As String

For i = 1 To grdResultadoBusqueda.Rows - 1
  
  If chkCambioDescripcionAnterior.value = 1 Then
    Descripcion = "'" & Trim(UCase(grdResultadoBusqueda.TextMatrix(i, 2))) & " // " & Trim(grdResultadoBusqueda.TextMatrix(i, 13)) & "'"
  Else
    If Trim(grdResultadoBusqueda.TextMatrix(i, 13)) <> "" Then
  
    Descripcion = "'" & Trim(grdResultadoBusqueda.TextMatrix(i, 13)) & "'"
    Else
    Descripcion = "Null"
    End If
    
  End If
  
  
  
  If grdResultadoBusqueda.TextMatrix(i, 1) = "Referencias" Then
    Sql = " UPDATE    REFERENCIAS"
    Sql = Sql & " SET INDICE ='" & Trim(lblCambioIndice.Caption) & "'"
    Sql = Sql & "  , FK_INDICES =" & lblCambioIndiceID
    Sql = Sql & "  , DESCRIPCION =" & Descripcion
    Sql = Sql & "  , COD_CLIENTE = " & ctlCambioCliente.Valor
    Sql = Sql & " Where COD_ID_REFERENCIA = " & grdResultadoBusqueda.TextMatrix(i, 0)
    ExecutarSql Sql
  
  
  End If
  
  
  If grdResultadoBusqueda.TextMatrix(i, 1) = "Legajos" Then
  
  
  Sql = " Update LEGAJOS"
Sql = Sql & " SET  FK_INDICES = " & lblCambioIndiceID
Sql = Sql & " , COD_INDICE =  '" & Trim(lblCambioIndice.Caption) & "'"
Sql = Sql & " , DESCRIPCION =" & Descripcion
Sql = Sql & " ,COD_CLIENTE =" & ctlCambioCliente.Valor
Sql = Sql & " Where ID_CLIENTE_LEGAJO = " & grdResultadoBusqueda.TextMatrix(i, 0)
Sql = Sql & " And COD_CLIENTE =" & ctlCliente.Valor
   ExecutarSql Sql
  
  
  End If
  
Next

MsgBox "Operacion terminada"
TituloGrilla

End Sub

Private Sub cmdCopiarExcel_Click()
CopiarDatosGrillaMSg grdSeleccionLegajos
End Sub

Private Sub cmdEntrada_Click()
Dim FechaEntrada As String
Dim lote As Integer
Dim Sql As String
Dim rs As ADODB.Recordset
Dim Valor As String
  Dim rsbusqueda As ADODB.Recordset
    Dim strIndice As String



If IsNull(ctlCliente.Valor) Then
        MsgBox "Ingrese el codigo del cliente"
    Exit Sub
End If

FechaEntrada = InputBox("Ingrese la fecha de entrada", "Fecha", Format(Now, "DD/MM/YYYY"))
lote = InputBox("Ingrese el lote el 0 solo tomara la fecha", "Lote", 0)
Sql = " SELECT     ELEMENTO, TIPO, COD_CLIENTE"
Sql = Sql & " From ENTRADA"
Sql = Sql & "  Where TIPO = 3"
Sql = Sql & "  AND COD_CLIENTE =" & ctlCliente.Valor
Sql = Sql & "  AND FECHA = '" & FechaEntrada & "'"
If lote <> 0 Then
    Sql = Sql & "  AND LOTE = " & lote
End If
Sql = Sql & "  ORDER BY ELEMENTO"

Set rs = New ADODB.Recordset

rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
    
    Valor = Valor & rs!Elemento & ","
    
    rs.MoveNext
    
Loop
Clipboard.Clear
Clipboard.SetText Valor
txtEtiqueta.Text = Valor

If Valor <> "" Then
    txtEtiqueta.Text = Mid(Valor, 1, Len(Valor) - 1)
End If


Sql = "   SELECT  LEGAJOS.COD_CLIENTE,  LEGAJOS.ID_CLIENTE_LEGAJO , LEGAJOS.NRO_CAJA , LEGAJOS.COD_ESTADO,"
                Sql = Sql & vbCrLf & " LEGAJOS.NRO_DESDE, LEGAJOS.NRO_HASTA, LEGAJOS.LETRA_DESDE, LEGAJOS.LETRA_HASTA, LEGAJOS.FECHA_DESDE,"
                Sql = Sql & vbCrLf & " LEGAJOS.FECHA_HASTA, LEGAJOS.DESCRIPCION, INDICES.DESCRIPCION AS DES_INDICE,CLIENTE_LEGAJO , DESCRIPCION_REMITO"
                Sql = Sql & vbCrLf & " FROM LEGAJOS LEFT OUTER JOIN INDICES ON LEGAJOS.FK_INDICES = INDICES.ID"
                
                

                    Sql = Sql & vbCrLf & " WHERE LEGAJOS.COD_CLIENTE = " & ctlCliente.Valor
                    Sql = Sql & vbCrLf & " AND  ID_CLIENTE_LEGAJO IN(" & txtEtiqueta.Text & ")"
                    Sql = Sql & vbCrLf & " ORDER BY ID_LEGAJO"
                
            
             Set rsbusqueda = New ADODB.Recordset
                rsbusqueda.Open Sql, ConActiva, 0, 1
                Do While Not rsbusqueda.EOF
                    
                    Rem  CargarGrillaBusqueda "Legajos", rsbusqueda!Des_Indice, rsbusqueda!ID_CLIENTE_LEGAJO, rsbusqueda!Cod_Estado, rsbusqueda!NRO_CAJA, "", rsbusqueda!NRO_DESDE, rsbusqueda!NRO_HASTA, rsbusqueda!FECHA_DESDE, rsbusqueda!FECHA_HASTA, rsbusqueda!LETRA_DESDE, rsbusqueda!LETRA_HASTA, rsbusqueda!DESCRIPCION, ""
                    rsbusqueda.MoveNext
                Loop


txtEtiqueta.Text = ""
End Sub

Private Sub cmdExcel_Click()
CopiarDatosGrillaMSg grdResultadoBusqueda
End Sub

Private Sub cmdInsertarBusqueda_Click()
        Dim Sql  As String
        Dim i As Integer
        Dim ID_LOTE_BUSQUEDA, TIPO, fecha, COD_CLIENTE, NRO_CAJA, Legajo, Descripcion, detalle As String
        Dim MaxLote As Integer
        Dim rs As New ADODB.Recordset
        
        On Error GoTo salir:
        
        rs.Open " SELECT MAX(ID_LOTE_BUSQUEDA) AS maxLote From TEM_BUSQUEDA", ConActiva, 0, 1
        ID_LOTE_BUSQUEDA = rs!MaxLote + 1
        
        For i = 1 To grdVarios.Rows - 1
            
            TIPO = "'" & grdVarios.TextMatrix(i, 1) & "'"
            fecha = SysDateMinutoSegundo
            COD_CLIENTE = grdVarios.TextMatrix(i, 2)
            NRO_CAJA = grdVarios.TextMatrix(i, 3)
            Legajo = "'" & grdVarios.TextMatrix(i, 4) & "'"
            Descripcion = "'" & grdVarios.TextMatrix(i, 5) & "'"
            detalle = "'" & grdVarios.TextMatrix(i, 6) & "'"
            
            Sql = " INSERT INTO TEM_BUSQUEDA "
            Sql = Sql & " (ID_LOTE_BUSQUEDA,TIPO "
            Sql = Sql & "  , FECHA, COD_CLIENTE "
            Sql = Sql & "  , NRO_CAJA, LEGAJO "
            Sql = Sql & "  , DESCRIPCION ,DETALLE) "
            Sql = Sql & "  VALUES "
            Sql = Sql & " (" & ID_LOTE_BUSQUEDA & "," & TIPO
            Sql = Sql & " ," & fecha & "," & COD_CLIENTE
            Sql = Sql & " ," & NRO_CAJA & "," & Legajo
            Sql = Sql & "," & Descripcion & "," & detalle & ")"
            ExecutarSql Sql
             
        Next
        
        MsgBox "Lote de busqueda " & ID_LOTE_BUSQUEDA & " Copiado a menoria"
        Clipboard.Clear
        Clipboard.SetText ID_LOTE_BUSQUEDA
        
        TitulosVarios
        ReporteBusqueda CInt(ID_LOTE_BUSQUEDA)
        Exit Sub
salir:
        MsgBox Err.Description
        
End Sub



Private Sub cmdLecturaLegajo_Click()
On Error GoTo salir:

 Dim rs As New ADODB.Recordset
    Dim MaxLectura As Long
    Dim Sql As String
    Dim i As Integer
    
    rs.Open "SELECT     MAX(ID_Lectura_Legajo) AS MaxLectura FROM         LECTURA_LEGAJO", ConActiva, 0, 1
    
    
    MaxLectura = rs!MaxLectura + 1
    
    
    If grdSeleccionLegajos.Rows = 2 Then
            Sql = " INSERT INTO LECTURA_LEGAJO"
            Sql = Sql & "(ID_Lectura_Legajo, Cod_Legajo_cliente, Cliente)"
            Sql = Sql & "  VALUES     (" & MaxLectura & "," & grdSeleccionLegajos.TextMatrix(1, 2) & "," & grdSeleccionLegajos.TextMatrix(1, 1) & ")"
            ExecutarSql Sql
     End If
     
    
    
    For i = 1 To grdSeleccionLegajos.Rows - 1
           Sql = " INSERT INTO LECTURA_LEGAJO"
           Sql = Sql & "(ID_Lectura_Legajo, Cod_Legajo_cliente, Cliente)"
            Sql = Sql & "  VALUES     (" & MaxLectura & "," & grdSeleccionLegajos.TextMatrix(i, 2) & "," & grdSeleccionLegajos.TextMatrix(i, 1) & ")"
            ExecutarSql Sql
           
    Next
 
MsgBox "Lectura es :" & MaxLectura & "  Esta copiada en memoria"
Clipboard.Clear
Clipboard.SetText MaxLectura

Exit Sub

salir:
MsgBox Err.Description


End Sub

Private Sub cmdLegajosEstado_Click()
Dim DATO As String
DATO = Clipboard.GetText
On Error GoTo salir:


Rem dato = Replace(dato, Chr(9), "@1@2")
DATO = Replace(DATO, vbCrLf, ") OR (NRO_DESDE = ")
DATO = Replace(DATO, vbTab, " AND YEAR(FECHA_DESDE) =")
Rem dato = Replace(dato, "@2", vbCrLf)
DATO = "(NRO_DESDE= " & DATO
txtLegajoEstado.Text = Mid(DATO, 1, Len(DATO) - 16)

MsgBox "NUMERO Y AÑO"

Exit Sub


salir:

MsgBox "NUMERO Y AÑO"


End Sub

Private Sub cmdLimpiar_Click()
TitulosVarios
End Sub

Private Sub cmdLimpiarLegajos_Click()
 TitulosSeleccionLegajos
End Sub

Private Sub cmdLimpiarMemoria_Click()
    Clipboard.Clear
End Sub

Private Sub cmdNºdesde_Click()
 Dim DATO As String
 On Error GoTo salir:
   DATO = Clipboard.GetText
   DATO = Replace(DATO, vbCrLf, ",")
        txtNroDesde.Text = Trim(DATO)
        If Trim(Mid(DATO, Len(DATO), 1)) = "," Then
        txtNroDesde.Text = Trim(Mid(txtNroDesde.Text, 1, Len(txtNroDesde.Text) - 1))
        
        End If
        Exit Sub
salir:
        MsgBox Err.Description
End Sub

Private Sub cmdNºEntre_Click()
Dim DATO As String
On Error GoTo salir:
   DATO = Clipboard.GetText
   If DATO <> "" Then
   DATO = Replace(DATO, vbCrLf, ",")
        txtNro_Entre.Text = DATO
        If Mid(DATO, Len(DATO), 1) = "," Then
        txtNro_Entre.Text = Mid(txtNro_Entre.Text, 1, Len(txtNro_Entre.Text) - 1)
        
        End If
    End If
            Exit Sub
salir:
        MsgBox Err.Description
End Sub


Private Sub cmdPasarTodos_Click()

    Dim TIPO As String
    Dim Cliente As String
    Dim Caja As String
    Dim Elemento As String
    Dim Descripcion, detalle As String
    Dim i As Integer
 
      With grdResultadoBusqueda
            For i = 1 To .Rows - 1
                    TIPO = grdResultadoBusqueda.TextMatrix(i, 1)
                    Cliente = ctlCliente.Valor
                    Caja = grdResultadoBusqueda.TextMatrix(i, 5)
                     If grdResultadoBusqueda.TextMatrix(i, 4) = 3 Then
                        If MsgBox("El legajo esta en consulta quiere continuar", vbYesNo + vbInformation) = vbNo Then
                            Exit Sub
                        End If
                    End If
                    detalle = "Nº Desde =" & grdResultadoBusqueda.TextMatrix(i, 7) & "   "
                    detalle = detalle & "Nº Hasta =" & grdResultadoBusqueda.TextMatrix(i, 8) & "   "
                    detalle = detalle & "Fecha Desde=" & grdResultadoBusqueda.TextMatrix(i, 9) & "   "
                    detalle = detalle & "Fecha Hasta=" & grdResultadoBusqueda.TextMatrix(i, 10) & "   "
                    detalle = detalle & "Letra Desde =" & grdResultadoBusqueda.TextMatrix(i, 11) & "   "
                    detalle = detalle & "Letra Hasta =" & grdResultadoBusqueda.TextMatrix(i, 12) & "   "
                    detalle = detalle & "desc.=" & grdResultadoBusqueda.TextMatrix(i, 13) & "   "
                    
                    Elemento = CadenaBusqueda
                    
                    Select Case grdResultadoBusqueda.TextMatrix(i, 1)
                    Case "Legajos"
                       Descripcion = "Etiqueta :" & grdResultadoBusqueda.TextMatrix(i, 3)
                       
                    Case "Referencias"
                         Descripcion = grdResultadoBusqueda.TextMatrix(i, 2)
                    Case "Rearchivo Digital"
                         Descripcion = Trim(grdResultadoBusqueda.TextMatrix(i, 8))
                    Case "Rearchivo"
                          Descripcion = Trim(grdResultadoBusqueda.TextMatrix(i, 3))
                    Case "Digital"
                          Descripcion = Trim(grdResultadoBusqueda.TextMatrix(i, 6))
                    
                    End Select
                
                grdVarios.AddItem "" & vbTab & TIPO & vbTab & Cliente & vbTab & Caja & vbTab & Elemento & vbTab & Descripcion & vbTab & detalle
                SSTab1.Tab = 1
                
                
                If grdResultadoBusqueda.TextMatrix(i, 1) = "Legajos" Then
                    If grdResultadoBusqueda.TextMatrix(i, 4) <> 3 Then
                        grdSeleccionLegajos.AddItem "" & vbTab & ctlCliente.Valor & vbTab & grdResultadoBusqueda.TextMatrix(i, 3) & vbTab & grdResultadoBusqueda.TextMatrix(i, 14) & vbTab & grdResultadoBusqueda.TextMatrix(i, 5)
                    Else
                         If MsgBox("El legajo esta en consulta Usted queiere ingresarlo Igual", vbYesNo) = vbYes Then
                            grdSeleccionLegajos.AddItem "" & vbTab & ctlCliente.Valor & vbTab & grdResultadoBusqueda.TextMatrix(i, 3) & vbTab & grdResultadoBusqueda.TextMatrix(i, 14) & vbTab & grdResultadoBusqueda.TextMatrix(i, 5)
                         End If
                    End If
                    SSTab1.Tab = 0
                End If
 
            Next
 End With
End Sub

Private Sub cmdPasarTodosLegajos_Click()
Dim i As Integer
On Error GoTo salir
  With grdResultadoBusqueda
For i = 1 To .Rows - 1
     If .TextMatrix(i, 1) = "Legajos" Then
            If .TextMatrix(i, 4) <> 2 Then
                If MsgBox("El legajo se encuentra en consulta" & vbCrLf & "Quiere pasarlo igual", vbYesNo + vbCritical) = vbYes Then
                    grdSeleccionLegajos.AddItem "" & vbTab & ctlCliente.Valor & vbTab & grdResultadoBusqueda.TextMatrix(i, 3) & vbTab & grdResultadoBusqueda.TextMatrix(i, 14) & vbTab & grdResultadoBusqueda.TextMatrix(i, 5)
               End If
        Else
            grdSeleccionLegajos.AddItem "" & vbTab & ctlCliente.Valor & vbTab & grdResultadoBusqueda.TextMatrix(i, 3) & vbTab & grdResultadoBusqueda.TextMatrix(i, 14) & vbTab & grdResultadoBusqueda.TextMatrix(i, 5)
        End If
    
    End If
    
 Next
   

   
   End With
   Exit Sub
salir:
MsgBox Err.Description

End Sub

Private Sub cmdReporteBusqueda_Click()
 Dim Reporte As Integer
     Reporte = InputBox("Ingrese el numero de reporte")
    ReporteBusqueda Reporte
End Sub

Private Sub cmdRotulos_Click()
 Dim sSQL As String
    Dim i As Integer
    sSQL = " SELECT  * FROM   LEGAJOS "
    sSQL = sSQL & " Where  "
    sSQL = sSQL & vbCrLf & "  ( COD_CLIENTE =" & grdSeleccionLegajos.TextMatrix(1, 1) & " AND ID_CLIENTE_LEGAJO =" & grdSeleccionLegajos.TextMatrix(1, 2) & ")"
    For i = 2 To grdSeleccionLegajos.Rows - 1
        sSQL = sSQL & vbCrLf & "  OR ( COD_CLIENTE =" & grdSeleccionLegajos.TextMatrix(i, 1) & " AND ID_CLIENTE_LEGAJO =" & grdSeleccionLegajos.TextMatrix(i, 2) & "  )  "
    Next
    sSQL = sSQL & vbCrLf & " ORDER BY ID_CLIENTE_LEGAJO "
    frmReportes.ImprimirReporte PasoReportes + "rptLegajos_Etiqueta.rpt", sSQL, True

End Sub

Private Sub CmdUnificar_Click()
    Dim Sql As String

       If Trim(txtLegajosUnificarHijos.Text) <> "" Then
                Sql = " Update LEGAJOS "
                Sql = Sql & " SET UNIFICACION_ID_LEGAJOS = " & TxtLegajosUnificarPadre.Text
                Sql = Sql & " ,FECHA_UNIFICACION = " & SysDate
                Sql = Sql & " ,COD_ESTADO = 9"
                Sql = Sql & " Where ID_LEGAJO in ( " & Mid(txtLegajosUnificarHijos.Text, 1, Len(txtLegajosUnificarHijos.Text) - 1) & ")"
                ExecutarSql Sql
                txtLegajosUnificarHijos.Text = ""
        End If
        
        
        If Trim(txtOrdenHijos.Text) <> "" Then
            Sql = " Update ORDENAR_DOCUMENTACION_DETALLE "
            Sql = Sql & " SET COD_ESTADO =9"
            Sql = Sql & " , UNIFICACION_ID_LEGAJOS =" & TxtLegajosUnificarPadre.Text
            Sql = Sql & " , FECHA_UNIFICACION =" & SysDate
            Sql = Sql & " Where ID in ( " & Mid(txtOrdenHijos.Text, 1, Len(txtOrdenHijos.Text) - 1) & ")"
            ExecutarSql Sql
        End If
        
        If Trim(txtImagenesUnificarHijos.Text) <> "" Then
            Sql = "  Update basasql.dbo.DOCUMENTOS_DIGITALES"
            Sql = Sql & " SET UNIFICACION_ID_LEGAJOS =" & TxtLegajosUnificarPadre.Text
            Sql = Sql & " , FECHA_UNIFICACION =" & SysDate
            Sql = Sql & " , COD_ESTADO =9"
            Sql = Sql & " WHERE ID IN (" & Mid(txtImagenesUnificarHijos.Text, 1, Len(txtImagenesUnificarHijos.Text) - 1) & ")"
            ExecutarSql Sql
       End If
    MsgBox "TERMINADO"

End Sub

Private Sub cmdUnificarDigital_Click()
Dim i As Integer
Dim Etiqueta As Long
Dim ID_imagen As Long
Dim Sql As String

For i = 1 To grdResultadoBusqueda.Rows - 1
    If grdResultadoBusqueda.TextMatrix(i, 1) = "Legajos" Then
    Etiqueta = grdResultadoBusqueda.TextMatrix(i, 3)
    End If
    If grdResultadoBusqueda.TextMatrix(i, 1) = "Digital" Then
    If MsgBox("Usted desea Unificar esta imagen " & grdResultadoBusqueda.TextMatrix(i, 6), vbYesNo) = vbYes Then
    
                Sql = "Update basasql.dbo.DOCUMENTOS_DIGITALES"
            Sql = Sql & " Set FK_ID_LEGAJO = " & Etiqueta
            Sql = Sql & " Where ID = " & grdResultadoBusqueda.TextMatrix(i, 3)
            ExecutarSql Sql
End If
    End If

Next


End Sub

Private Sub Command1_Click()
Dim i As Integer
For i = 0 To grdResultadoBusqueda.Cols - 1
    Debug.Print ".ColWidth(" & i & ") = " & grdResultadoBusqueda.ColWidth(i)
 Next
End Sub

Private Sub Command2_Click()
 Dim DATO As String
 On Error GoTo salir:
   DATO = Clipboard.GetText
   DATO = Replace(DATO, vbCrLf, ",")
        txtEtiqueta.Text = DATO
        If Mid(DATO, Len(DATO), 1) = "," Then
        txtEtiqueta.Text = Mid(txtEtiqueta.Text, 1, Len(txtEtiqueta.Text) - 1)
        
        End If
        
        Exit Sub
salir:
        MsgBox Err.Description & "  " & txtEtiqueta.Text
        
End Sub

Private Sub Command3_Click()
Dim DATO As String
DATO = Clipboard.GetText
Dim Indice As String


chkLegajos.value = 0
chkRearchivo.value = 0
ChkRearchivoDigital = 0
chkReferencias.value = 1

On Error GoTo salir:

Dim Doc(1000) As String
Dim fecha(1000) As String
Dim i As Integer
Dim C As Integer
Dim j As Integer
Dim DatoCom(1000)  As String

DATO = Replace(Replace(DATO, vbCrLf, "@"), vbTab, "&")



For i = 1 To Len(DATO)
    If InStr(i, DATO, "@") = 0 Then
        MsgBox "EL dato est mal"
        Exit Sub
    End If
    DatoCom(C) = Mid(DATO, i, 17)
    C = C + 1
    i = InStr(i, DATO, "@")
    If i > Len(DATO) Then
        Exit Sub
    End If

Next

MsgBox " NUMERO DOCUMENTO  Y FECHA"
For i = 1 To 1000
    
If Trim(DatoCom(i)) <> "" Then
   Doc(i) = Mid(DatoCom(i), 1, 5)
   fecha(i) = Mid(DatoCom(i), 7, 10)
End If



Next




Dim Sql As String

For i = 0 To 99

Dim Fecha1 As String
Dim Indice1 As String


 If Doc(i) <> "" Then
 Fecha1 = Replace(Replace(Trim(fecha(i)), vbCrLf, ""), vbTab, "")
 Indice1 = BuscarIndiceDocumento_Indice(Doc(i), ctlCliente.Valor)
    Sql = Sql & " (  REFERENCIAS.INDICE = '" & Indice1 & "' AND ( " & FechaFormato(Fecha1) & " BETWEEN FECHA_DESDE AND FECHA_HASTA) )  OR  " & vbCrLf
End If


Next




If Sql <> "" Then
Sql = Mid(Sql, 1, Len(Sql) - 6)

txtDocumento_fecha.Text = Sql

End If

Exit Sub
salir:
MsgBox Err.Description

End Sub

Private Sub Command4_Click()
 Dim DATO As String
 On Error GoTo salir:
   DATO = Clipboard.GetText
   DATO = Replace(DATO, vbCrLf, ",")
        txtCaja.Text = DATO
        If Mid(DATO, Len(DATO), 1) = "," Then
        txtCaja.Text = Mid(txtCaja.Text, 1, Len(txtCaja.Text) - 1)
        
        End If
        
        Exit Sub
        
salir:
txtCaja.Text = ""
        MsgBox Err.Description & "  " & txtCaja.Text
End Sub

Private Sub Command5_Click()
Dim Sql As String
Dim rsbusqueda As New ADODB.Recordset
TituloGrilla
Sql = " SELECT INDICES.DESCRIPCION as DES_INDICE, COD_ID_REFERENCIA , REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA, REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA,"
Sql = Sql & "   REFERENCIAS.LETRA_DESDE, REFERENCIAS.LETRA_HASTA, REFERENCIAS.DESCRIPCION , REFERENCIAS.NRO_CAJA,"
Sql = Sql & " REFERENCIAS.COD_CLIENTE , REFERENCIAS.Indice, REFERENCIAS.FECHA_MODIFICACION, REFERENCIAS.PASOARCHIVO"
Sql = Sql & " FROM REFERENCIAS INNER JOIN"
Sql = Sql & " INDICES ON REFERENCIAS.INDICE = INDICES.INDICE AND REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE"
Sql = Sql & "  Where  PASOARCHIVO LIKE '%" & TxtOrden.Text & "%'"
Sql = Sql & " ORDER BY NRO_CAJA, COD_ID_REFERENCIA "
 
                        
                        Set rsbusqueda = New ADODB.Recordset
                        rsbusqueda.Open Sql, ConActiva, 0, 1
                        Do While Not rsbusqueda.EOF
                            CargarGrillaBusqueda rsbusqueda!COD_ID_REFERENCIA, "Referencias", rsbusqueda!Des_indice, "", 2, rsbusqueda!NRO_CAJA, "", rsbusqueda!NRO_DESDE, rsbusqueda!NRO_HASTA, rsbusqueda!FECHA_DESDE, rsbusqueda!FECHA_HASTA, rsbusqueda!LETRA_DESDE, rsbusqueda!LETRA_HASTA, rsbusqueda!Descripcion, "", rsbusqueda!FECHA_MODIFICACION, rsbusqueda!PASOARCHIVO
                            rsbusqueda.MoveNext
                        Loop
End Sub

Private Sub ctlCliente_Click()
    LimpiarCampos
    TituloGrilla
    txtEtiqueta.Text = ""
    txtCaja.Text = ""
    txtIndice_Nro_Documento.Text = ""
    ctlCambioCliente.Valor = ctlCliente.Valor

End Sub

Private Sub DataGrid1_ButtonClick(ByVal ColIndex As Integer)

End Sub

Private Sub DataGrid1_Click()

End Sub

Private Sub ctlCliente_LostFocus()
ctlCambioCliente.Valor = ctlCliente.Valor
End Sub

Private Sub Form_Activate()
On Error GoTo salir:
 frmAgregarDocumentos.Top = 0
 frmAgregarDocumentos.Left = 0
 Dim IndiceAnterior As String
 If Nro_documento <> 0 And SSTBusqueda.Tab = 2 Then
  IndiceAnterior = BuscarIDDocumento(CLng(Nro_documento), ctlCambioCliente.Valor)
  If Len(IndiceAnterior) > 3 Then
    IndiceAnterior = Mid(IndiceAnterior, 1, Len(IndiceAnterior) - 3)
    txtCambioIndiceDocumento.Text = Buscar_ID_CODIGO_DOCUMENTO(IndiceAnterior, ctlCambioCliente.Valor)
    End If
  Else
     frmBuscarGenerico.txtIndice_Nro_Documento = Nro_documento
 End If
  SSTBusqueda.Tab = 1


  
salir:
End Sub

Private Sub Form_Load()
ctlCliente.TipoControl = Cliente
ctlCambioCliente.TipoControl = Cliente
TitulosVarios
chkLegajos.value = 0
chkRearchivo.value = 0
ChkRearchivoDigital.value = 0
chkReferencias.value = 0
TitulosSeleccionLegajos
Nro_documento = ""
txtIndice_Nro_Documento.Text = ""
 If MDIfrmInicio.StaInicio.Panels(2) = 48 Or MDIfrmInicio.StaInicio.Panels(2) = 17 Or MDIfrmInicio.StaInicio.Panels(2) = 48 Or MDIfrmInicio.StaInicio.Panels(2) = 48 Then
   Rem  SSTBusqueda.TabEnabled(2) = True
   cmdBorrarReferencia.Enabled = True
   
 Else
   Rem  SSTBusqueda.TabEnabled(2) = False
    cmdBorrarReferencia.Enabled = False
 End If
   If MDIfrmInicio.StaInicio.Panels(2) = 46 Or MDIfrmInicio.StaInicio.Panels(2) = 17 Then
        cmdUnificarDigital.Enabled = True
    End If

End Sub

Private Sub TitulosVarios()


With grdVarios
    .Clear
    .Cols = 7
    .Rows = 1
    .ColAlignment(1) = 0
    .ColAlignment(2) = 0
    .ColAlignment(3) = 0
    .ColAlignment(4) = 0
     .ColAlignment(5) = 0
    
    .ColWidth(0) = 100
    .ColWidth(1) = 2000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1500
    .ColWidth(5) = 3500
    
    .TextMatrix(0, 1) = "Proceso"
    .TextMatrix(0, 2) = "Cliente"
    .TextMatrix(0, 3) = "Caja"
    .TextMatrix(0, 4) = "Elemento"
    .TextMatrix(0, 5) = "Descripcion"
    .TextMatrix(0, 6) = "Detalle"
    End With

End Sub



Private Sub grdResultadoBusqueda_Click()
Detalle_ver
'If grdResultadoBusqueda.Row = 1 Then
'grdResultadoBusqueda.Sort = grdResultadoBusqueda.Col
'End If

End Sub

Private Sub grdResultadoBusqueda_DblClick()

Dim TIPO As String
Dim Cliente As String
Dim Caja As String
Dim Elemento As String
Dim Descripcion, detalle As String

 
 


    TIPO = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 1)
    Cliente = ctlCliente.Valor
    Caja = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 5)
    
    
    If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 4) = 3 Then
        If MsgBox("El legajo esta en consulta quiere continuar", vbYesNo + vbInformation) = vbNo Then
            Exit Sub
        End If
    End If
    

    
    detalle = "Nº Desde =" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 7) & "   "
    detalle = detalle & "Nº Hasta =" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 8) & "   "
    detalle = detalle & "Fecha Desde=" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 9) & "   "
    detalle = detalle & "Fecha Hasta=" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 10) & "   "
    detalle = detalle & "Letra Desde =" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 11) & "   "
    detalle = detalle & "Letra Hasta =" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 12) & "   "
    detalle = detalle & "desc.=" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 13) & "   "
    
    Elemento = CadenaBusqueda
    
    Select Case grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 1)
    Case "Legajos"
       Descripcion = "Etiqueta :" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 3)
       
    Case "Referencias"
         Descripcion = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 2)
    Case "Rearchivo Digital"
         Descripcion = Trim(grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 8))
    Case "Rearchivo"
          Descripcion = Trim(grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 3))
    Case "Digital"
          Descripcion = Trim(grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 6))
    
    End Select

grdVarios.AddItem "" & vbTab & TIPO & vbTab & Cliente & vbTab & Caja & vbTab & Elemento & vbTab & Descripcion & vbTab & detalle
SSTab1.Tab = 1

'grdSeleccionLegajos.TextMatrix(0, 1) = "Cliente"
'    grdSeleccionLegajos.TextMatrix(0, 2) = "Etiqueta"
'    grdSeleccionLegajos.TextMatrix(0, 3) = "Legajo Cliente"
'    grdSeleccionLegajos.TextMatrix(0, 4) = "Caja"
'
'.TextMatrix(0, 1) = "Proceso"
'    .TextMatrix(0, 2) = "Tipo Doc"
'    .TextMatrix(0, 3) = "Etiqueta"
'    .TextMatrix(0, 4) = "Estado"
'    .TextMatrix(0, 5) = "Caja"
'    .TextMatrix(0, 6) = "Lote"
'    .TextMatrix(0, 7) = "Nro Desde"
'    .TextMatrix(0, 8) = "Nro Hasta"
'    .TextMatrix(0, 9) = "Fecha Desde"
'    .TextMatrix(0, 10) = "Fecha Hasta"
'    .TextMatrix(0, 11) = "Letra Desde"
'    .TextMatrix(0, 12) = "letra Hasta"
'    .TextMatrix(0, 13) = "Descripcion"
'    .TextMatrix(0, 13) = "Varios"


If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 1) = "Legajos" Then
    If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 4) <> 3 Then

        grdSeleccionLegajos.AddItem "" & vbTab & ctlCliente.Valor & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 3) & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 14) & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 5)
    Else
    
         If MsgBox("El legajo esta en consulta Usted queiere ingresarlo Igual", vbYesNo) = vbYes Then
            grdSeleccionLegajos.AddItem "" & vbTab & ctlCliente.Valor & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 3) & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 14) & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 5)
         End If
    End If
    SSTab1.Tab = 0
End If

End Sub

Private Sub grdResultadoBusqueda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If SSTBusqueda.Tab = 3 Then
        If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Legajos" Then
             mnuCargarPadre.Enabled = True
             Else
             mnuCargarPadre.Enabled = False
        End If
        
        PopupMenu mnuUnificar
       
    
    Else
        If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Digital" Then
            PopupMenu mnuImagenes
        End If
        If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Legajos" Then
            PopupMenu mnuLegajos
        End If
        If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Referencias" Then
            PopupMenu mnuReferencias
        End If
        
    End If
End If


End Sub

Private Sub grdResultadoBusqueda_RowColChange()
Detalle_ver
End Sub

Private Sub grdVarios_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuVarios

End If

End Sub

Private Sub Impresono_Click()
Dim Sql As String
Sql = " Update DOCUMENTOS_DIGITALES Set IMPRESO = 0 Where ID = " & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 0)
ExecutarSql Sql
End Sub

Private Sub Impresosi_Click()
Dim Sql As String

Sql = " Update DOCUMENTOS_DIGITALES Set IMPRESO = 1 Where ID = " & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 0)
ExecutarSql Sql


End Sub

Private Sub lblDetalle_descripcion_DblClick()
Clipboard.Clear
Clipboard.SetText lblDetalle_descripcion.Caption
MsgBox "Las descricipcion fue copiado a memoria"
End Sub

Private Sub mnuBorrarEtiquetaVirtual_Click()
 Dim Sql As String
 If MDIfrmInicio.StaInicio.Panels(2) = 48 Or MDIfrmInicio.StaInicio.Panels(2) = 17 Or MDIfrmInicio.StaInicio.Panels(2) = 48 Then
        
     If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Legajos" Then
            Sql = " UPDATE    LEGAJOS"
            Sql = Sql & "  SET LETRA_DESDE = NULL, LETRA_HASTA = NULL, NRO_DESDE = NULL, NRO_HASTA = NULL, FECHA_DESDE = NULL, FECHA_HASTA = NULL,"
            Sql = Sql & " DESCRIPCION = NULL, NRO_CAJA = 0, COD_CLIENTE = 0, ID_PERSONAL = NULL, FK_PERSONAL_CREACION = NULL,"
            Sql = Sql & " FECHA_ACTUALIZACION = NULL, FECHA_CREACION = NULL,COD_ESTADO=NULL, COD_INDICE = NULL, FK_INDICES = NULL"
            Sql = Sql & " Where    ID_CLIENTE_LEGAJO = " & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 3)
            Sql = Sql & " AND COD_CLIENTE = " & ctlCliente.Valor
            If MsgBox("Esta usted seguro de borrar el registro", vbCritical + vbYesNo) = vbYes Then
                ExecutarSql Sql
            End If
        End If
        
        Else
        MsgBox "El Usuario No puede realiaza esta operacion"
     End If
End Sub

Private Sub mnuBorrarVarios_Click()
 grdVarios.RemoveItem grdVarios.Row
End Sub

Private Sub mnuCargarHijos_Click()
'         If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Digital" Then
'            txtLegajosUnificarHijos
'        End If
        If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Legajos" Then
            txtLegajosUnificarHijos.Text = txtLegajosUnificarHijos.Text & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 0) & ","
        End If
        
        If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Rearchivo" Then
            txtOrdenHijos.Text = txtOrdenHijos.Text & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 0) & ","
        End If
        If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Digital" Then
            txtImagenesUnificarHijos.Text = txtImagenesUnificarHijos & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 0) & ","
            
        End If
End Sub

Private Sub mnuCargarPadre_Click()
    TxtLegajosUnificarPadre.Text = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 0)
End Sub

Private Sub mnuCarpetaSincronizada_Click()
Clipboard.Clear
Dim ID As Long
Dim N_desde As String
Dim LETRA_DESDE As String
 Dim Nombre  As String

ID = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 3)
N_desde = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 7)


LETRA_DESDE = Trim(grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 11))


 
 Nombre = N_desde & " " & LETRA_DESDE & " " & ID


FileSystem.FileCopy PasoImagenes & BuscarDirectorioPaso(ID) & "\" & CStr(ID) & ".TIF", "\\222.15.19.251\basa\Planta\Imagenes Sistema Basa\" & Nombre & ".TIF"
MsgBox "Informacion copiada Imagen " & Nombre, vbInformation
End Sub

Private Sub mnuCopiarImagenC_Click()
Clipboard.Clear
Dim ID As Long
Dim Nombre As String
    ID = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 3)
    Nombre = txtNroDesde.Text & txtNroHasta.Text & txtFechaDesde.Text & txtFechaHasta.Text & txtLetraDesde.Text & txtLetraHasta.Text & txtNro_Entre.Text & txtLetra_Entre.Text & "_" & CStr(ID)
    FileCopy PasoImagenes & BuscarDirectorioPaso(ID) & "\" & CStr(ID) & ".TIF", "C:\IMAGENES\" & Trim(Nombre) & ".TIF"
    MsgBox "Terminado"
End Sub

Private Sub mnuCopiarPaso_Click()
Clipboard.Clear
Dim ID As Long
ID = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 3)

Rem Shell  PasoImagenes & BuscarDirectorioPaso(ID) & "\" & CStr(ID) & ".TIF", vbMaximizedFocus



Clipboard.SetText PasoImagenes & BuscarDirectorioPaso(ID) & "\" & CStr(ID) & ".TIF"
MsgBox "Informacion copiada", vbInformation


End Sub

Private Sub mnuCrearLegajo_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim IDLegajo As Long


Sql = "  SELECT    TOP 1 ID_LEGAJO"
Sql = Sql & " From LEGAJOS"
Sql = Sql & " WHERE    ID_LEGAJO BETWEEN  7333582 AND 7344482   And  (COD_CLIENTE IS NULL)  "
Sql = Sql & " ORDER BY ID_LEGAJO"


rs.CursorLocation = adUseClient

rs.Open Sql, strConBasa, adOpenForwardOnly, adLockReadOnly

If rs.EOF Then
    MsgBox "err"
    Exit Sub

Else
IDLegajo = rs!ID_LEGAJO

End If



Dim rsReferencias As New ADODB.Recordset
Dim sqlR As String

Dim NUMERO As Long
Dim lETRA As String
Dim FK_Indice As Long

 Dim COD_CLIENTE As String
 Dim NRO_CAJA As String
 Dim Indice As String
 Dim Descripcion As String
 Dim FECHA_DESDE As String
 Dim FECHA_HASTA As String
 Dim FECHA_MODIFICACION As String
 Dim FECHA_CREACION As String
 Dim USUARIO_MODIFICACION As String
sqlR = sqlR & " FK_PERSONAL_CREACION , FK_PERSONAL_MODIFICACION, COD_ID_REFERENCIA"



sqlR = " SELECT     COD_CLIENTE, NRO_CAJA, INDICE, DESCRIPCION, FECHA_DESDE, FECHA_HASTA, FECHA_MODIFICACION, FECHA_CREACION, USUARIO_MODIFICACION,"
sqlR = sqlR & " FK_PERSONAL_CREACION , FK_PERSONAL_MODIFICACION, COD_ID_REFERENCIA"
sqlR = sqlR & " From REFERENCIAS"
sqlR = sqlR & " Where COD_ID_REFERENCIA = " & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 0)

rsReferencias.Open sqlR, ConActiva, 0, 1


If Not rsReferencias.EOF Then

NUMERO = InputBox("Ingrese el numero", , "0")
lETRA = InputBox("Ingrese la letra", "", "Null")
If lETRA = "Null" Then

Else
    lETRA = "'" & UCase(Trim(lETRA)) & "'"
End If


FK_Indice = Buscar_ID_Indice_Por_indice(rsReferencias!Indice, rsReferencias!COD_CLIENTE)

 COD_CLIENTE = rsReferencias!COD_CLIENTE
 NRO_CAJA = rsReferencias!NRO_CAJA
 Indice = "'" & rsReferencias!Indice & "'"
 Descripcion = "'" & UCase(Trim(rsReferencias!Indice)) & "'"
 
 If IsNull(rsReferencias!FECHA_DESDE) Then
     FECHA_DESDE = "Null"
 
 Else
     FECHA_DESDE = FechaFormato(rsReferencias!FECHA_DESDE)
 End If
 
 If IsNull(rsReferencias!FECHA_HASTA) Then
 FECHA_HASTA = "Null"
 Else
  
 FECHA_HASTA = FechaFormato(rsReferencias!FECHA_HASTA)
 End If
 
 FECHA_MODIFICACION = "'" & rsReferencias!FECHA_MODIFICACION & "'"
 FECHA_CREACION = "'" & rsReferencias!FECHA_CREACION & "'"
 USUARIO_MODIFICACION = MDIfrmInicio.StaInicio.Panels(2).Text

Sql = " Update LEGAJOS "
Sql = Sql & vbCrLf & " SET COD_INDICE =" & Indice
Sql = Sql & vbCrLf & " , FK_INDICES =" & CStr(FK_Indice)
Sql = Sql & vbCrLf & ", LETRA_DESDE =" & lETRA
Sql = Sql & vbCrLf & ", LETRA_HASTA =" & lETRA
Sql = Sql & vbCrLf & ", NRO_DESDE =" & NUMERO
Sql = Sql & vbCrLf & ", NRO_HASTA =" & NUMERO
Sql = Sql & vbCrLf & ", FECHA_DESDE =" & FECHA_DESDE
Sql = Sql & vbCrLf & ", FECHA_HASTA =" & FECHA_HASTA
Sql = Sql & vbCrLf & ", NRO_CAJA =" & NRO_CAJA
Sql = Sql & vbCrLf & ", DESCRIPCION =" & Descripcion
Sql = Sql & vbCrLf & ", COD_CLIENTE =" & COD_CLIENTE
Sql = Sql & vbCrLf & ", COD_ESTADO  =2"
Sql = Sql & "  Where ID_LEGAJO = " & IDLegajo
If ExecutarSql(Sql) = 1 Then
    MsgBox "Su numero de legajos es " & IDLegajo
    Clipboard.SetText (IDLegajo & vbCrLf & Clipboard.GetText)
 Else
    MsgBox "ERROR "
 End If
        
End If





End Sub

Private Sub mnuImagenesver_Click()
    Dim ID As Long
    ID = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 3)
   ctlVerImagenes1.PonerImagen PasoImagenes & BuscarDirectorioPaso(ID) & "\" & CStr(ID) & ".TIF"
   
  
    SSTab1.Tab = 3
End Sub

Private Sub mnuModificar_Legajos_Click()
       
  If Not IsNull(ctlCliente.Valor) Then
        frmAgregarDocumentos.Modificar_Legajos ctlCliente.Valor, grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 3)
        frmAgregarDocumentos.Show
        frmAgregarDocumentos.SetFocus
  Else
        MsgBox "Ingrese el cliente"
  End If
End Sub

Private Sub mnuReferenciasModificar_Click()
frmAgregarDocumentos.CargarReferencias ctlCliente.Valor, grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 0), "Modificar"
frmAgregarDocumentos.Show
frmAgregarDocumentos.SetFocus
End Sub

Private Sub mnuReferenciasNuevas_Click()
frmAgregarDocumentos.CargarReferencias ctlCliente.Valor, grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 0), "Nuevo"
frmAgregarDocumentos.Show
frmAgregarDocumentos.SetFocus
End Sub

Private Sub mnuReferenciasNuevasCopiar_Click()
frmAgregarDocumentos.CargarReferencias ctlCliente.Valor, grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 0), "NuevoCopiar"
frmAgregarDocumentos.Show
frmAgregarDocumentos.SetFocus
End Sub

Private Sub mnuVerImagen_Click()
Dim Imagen As String
Dim Paso As String
Paso = "Z:\Administracion\Imagenes_Internas\Cajas\" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 5)

Imagen = Dir(Paso & "\*.tif")
If Imagen <> "" Then
ctlVerImagenes1.PonerImagen Paso & "\" & Imagen
SSTab1.Tab = 3

 
Else
 MsgBox "No tiene Imagen"
End If

End Sub

Private Sub txtCambioIndiceDocumento_Change()
If IsNumeric(txtCambioIndiceDocumento) Then
lblCambioIndice = BuscarIDDocumento(txtCambioIndiceDocumento, ctlCambioCliente.Valor)
lblCambioDescripcion = BuscarIndiceDescripcion(lblCambioIndice, ctlCambioCliente.Valor)
lblCambioIndiceID = Buscar_ID_Indice(txtCambioIndiceDocumento, ctlCambioCliente.Valor)
Else
lblCambioIndice = ""
lblCambioDescripcion = ""
lblCambioIndiceID = ""


End If


End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If

End Sub

Private Sub txtEtiqueta_Change()
chkLegajos = 1
chkRearchivo = 0
ChkRearchivoDigital.value = 0
chkReferencias.value = 0
End Sub

Private Sub txtFecha_Entre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If

End Sub

Private Sub txtFechaDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If

End Sub

Private Sub txtFechaHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If

End Sub

Private Sub txtIndice_Nro_Documento_Change()
On Error GoTo salir:
lblIndice_Descripcion.Caption = ""
LimpiarCampos
If txtIndice_Nro_Documento.Text <> "" And txtIndice_Nro_Documento.Text <> "0" Then
    Configurar_Carga ctlCliente.Valor, txtIndice_Nro_Documento.Text
 Else
  LimpiarCampos
  lblIndice_Descripcion.Caption = ""
 End If
    TituloGrilla
    Exit Sub
    
salir:
  MsgBox Err.Description
 
End Sub

Public Sub CargarGrillaBusqueda(ID, Proceso, TipoDoc, Etiqueta, estado, Caja, lote, NRO_DESDE, NRO_HASTA _
 , FECHA_DESDE, FECHA_HASTA, LETRA_DESDE, LETRA_HASTA, Descripcion, VARIOS, FechaCarga, PASOARCHIVO, Optional REARCHIVO_CAJA As String)

Dim Sql As String
VARIOS = Replace(VARIOS, vbTab, "")
If IsNull(Descripcion) Then
Descripcion = ""
End If

Sql = Sql & ID & vbTab
Sql = Sql & Proceso & vbTab
Sql = Sql & TipoDoc & vbTab
Sql = Sql & Etiqueta & vbTab
Sql = Sql & estado & vbTab
Sql = Sql & Caja & vbTab
Sql = Sql & lote & vbTab
Sql = Sql & NRO_DESDE & vbTab
Sql = Sql & NRO_HASTA & vbTab
Sql = Sql & FECHA_DESDE & vbTab
Sql = Sql & FECHA_HASTA & vbTab
Sql = Sql & LETRA_DESDE & vbTab
Sql = Sql & LETRA_HASTA & vbTab
Sql = Sql & Descripcion & vbTab
Sql = Sql & VARIOS & vbTab
Sql = Sql & FechaCarga & vbTab
Sql = Sql & Trim(PASOARCHIVO) & vbTab
Sql = Sql & Trim(REARCHIVO_CAJA)
grdResultadoBusqueda.AddItem Sql
If estado = 3 Then
    grdResultadoBusqueda.Col = 4
    grdResultadoBusqueda.Row = grdResultadoBusqueda.Rows - 1
    grdResultadoBusqueda.CellBackColor = &H8080FF
    grdResultadoBusqueda.Refresh
End If

If estado = 2 Then
    grdResultadoBusqueda.Col = 4
    grdResultadoBusqueda.Row = grdResultadoBusqueda.Rows - 1
    grdResultadoBusqueda.CellBackColor = &H80FF80
    grdResultadoBusqueda.Refresh
End If

If estado = 9 Then
    grdResultadoBusqueda.Col = 4
    grdResultadoBusqueda.Row = grdResultadoBusqueda.Rows - 1
    grdResultadoBusqueda.CellBackColor = &HFFC0FF
    grdResultadoBusqueda.Refresh
End If


End Sub

Public Sub TituloGrilla()
With grdResultadoBusqueda
    .Clear
    .Cols = 18
    .Rows = 1
    .ColAlignment(1) = 0
    .ColAlignment(2) = 0
    .ColAlignment(3) = 0
    .ColAlignment(4) = 0
    .ColAlignment(5) = 0
    .ColAlignment(6) = 0
    .ColAlignment(7) = 0
    .ColAlignment(8) = 0
    .ColAlignment(9) = 0
    .ColAlignment(10) = 0
    .ColAlignment(11) = 0
    .ColAlignment(12) = 0
    .ColAlignment(13) = 0
     .ColAlignment(14) = 0
     .ColAlignment(15) = 0
     .ColAlignment(16) = 0
     .ColAlignment(17) = 0
    
        .ColWidth(0) = 105
        .ColWidth(1) = 990
        .ColWidth(2) = 1605
        .ColWidth(3) = 945
        .ColWidth(4) = 750
        .ColWidth(5) = 780
        .ColWidth(6) = 930
        .ColWidth(7) = 960
        .ColWidth(8) = 945
        .ColWidth(9) = 1095
        .ColWidth(10) = 1095
        .ColWidth(11) = 1500
        .ColWidth(12) = 1365
        .ColWidth(13) = 1995
        .ColWidth(14) = 990
         .ColWidth(15) = 990
          .ColWidth(16) = 1000
           .ColWidth(17) = 1000
    
    .TextMatrix(0, 1) = "Proceso"
    .TextMatrix(0, 2) = "Tipo Doc"
    .TextMatrix(0, 3) = "Etiqueta"
    .TextMatrix(0, 4) = "Estado"
    .TextMatrix(0, 5) = "Caja"
    .TextMatrix(0, 6) = "Lote"
    .TextMatrix(0, 7) = "Nro Desde"
    .TextMatrix(0, 8) = "Nro Hasta"
    .TextMatrix(0, 9) = "Fecha Desde"
    .TextMatrix(0, 10) = "Fecha Hasta"
    .TextMatrix(0, 11) = "Letra Desde"
    .TextMatrix(0, 12) = "Letra Hasta"
    .TextMatrix(0, 13) = "Descripcion"
    .TextMatrix(0, 14) = "Varios"
    .TextMatrix(0, 15) = "Fecha_Modificacion"
    .TextMatrix(0, 16) = "Archivo"
    .TextMatrix(0, 17) = "Caja Rearch"
    
   End With

End Sub

Public Sub LimpiarCampos()

    txtFecha_Entre.Text = ""
    txtNro_Entre.Text = ""
    txtLetra_Entre.Text = ""

    txtDescripcion.Text = ""
    txtDescripcion.BackColor = &H80000005
    txtDescripcion.Enabled = True
    lblTituloDescripcion.Caption = "Descripcion"
    
    
    txtFechaDesde.Text = ""
    txtFechaDesde.BackColor = &H80000005
    txtFechaDesde.Enabled = True
    lblTituloFechaDesde.Caption = "Fecha Desde"

    txtFechaHasta.Text = ""
    txtFechaHasta.BackColor = &H80000005
    txtFechaHasta.Enabled = True
    lblTituloFechaHasta.Caption = "Fecha Hasta"
    
    

    txtLetraDesde.Text = ""
    txtLetraDesde.BackColor = &H80000005
    txtLetraDesde.Enabled = True
    lblTituloLetraDesde.Caption = "Letra Desde"
    
    
    txtLetraHasta.Text = ""
    txtLetraHasta.BackColor = &H80000005
    txtLetraHasta.Enabled = True
    lblTituloLetraHasta.Caption = "Letra Hasta"
    


    txtNroDesde.Text = ""
    txtNroDesde.BackColor = &H80000005
    txtNroDesde.Enabled = True
    lblTituloNumeroDesde.Caption = "Nro Desde"
    

    txtNroHasta.Text = ""
    txtNroHasta.BackColor = &H80000005
    txtNroHasta.Enabled = True
    lblTituloNumeroHasta.Caption = "Nro Hasta"
    




End Sub

Private Sub txtLecturaCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
txtLegajosUnificarHijos.Text = txtLegajosUnificarHijos.Text & CLng(Mid(txtLecturaCodigo.Text, 3, 10)) & ","
txtLecturaCodigo.Text = ""
End If

End Sub

Private Sub txtLetra_Entre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If

End Sub

Private Sub txtLetraDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If
End Sub

Private Sub txtLetraHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If

End Sub

Private Sub txtNro_Entre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If

End Sub

Private Sub txtNroDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If

End Sub

Private Sub txtNroHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If

End Sub

Public Sub Detalle_ver()
Dim R As Integer
R = grdResultadoBusqueda.Row

lblDetalleIndiceDescripcion.Caption = grdResultadoBusqueda.TextMatrix(R, 2)
lblDetalleCaja.Caption = grdResultadoBusqueda.TextMatrix(R, 5)
 lblDetalleLote.Caption = grdResultadoBusqueda.TextMatrix(R, 6)
lblDetalle_nro_desde.Caption = grdResultadoBusqueda.TextMatrix(R, 7)
lblDetalle_Nro_Hasta.Caption = grdResultadoBusqueda.TextMatrix(R, 8)
lblDetalle_fecha_desde.Caption = grdResultadoBusqueda.TextMatrix(R, 9)
lblDetalle_fecha_hasta.Caption = grdResultadoBusqueda.TextMatrix(R, 10)
lblDetalle_Letra_desde.Caption = grdResultadoBusqueda.TextMatrix(R, 11)
lblDetalle_Letra_hasta.Caption = grdResultadoBusqueda.TextMatrix(R, 12)
lblDetalle_descripcion.Caption = grdResultadoBusqueda.TextMatrix(R, 13)

SSTab1.Tab = 2

End Sub
