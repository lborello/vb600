VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmOrdenDocumentacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
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
   ScaleHeight     =   8430
   ScaleWidth      =   11115
   Begin TabDlg.SSTab tabOrdenDocumentacion 
      Height          =   7935
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   5
      TabHeight       =   520
      TabCaption(0)   =   "Orden de documentación"
      TabPicture(0)   =   "frmOrdenDocumentacion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOrdenDocuementacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "-"
      TabPicture(1)   =   "frmOrdenDocumentacion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOrdenLegajos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Buscar Documentos"
      TabPicture(2)   =   "frmOrdenDocumentacion.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraBuscarOrdenamiento"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Administración"
      TabPicture(3)   =   "frmOrdenDocumentacion.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(3)=   "fraReparacionOrden"
      Tab(3).Control(4)=   "fraOrdenControl"
      Tab(3).Control(5)=   "fraOrdenTerminado"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Orden de Documentacion Osep"
      TabPicture(4)   =   "frmOrdenDocumentacion.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraAfiliadoOsep"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Cambio de Ubicacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   -69240
         TabIndex        =   108
         Top             =   3000
         Width           =   3375
         Begin VB.TextBox txtCambioUbicacionUbicacion 
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
            HideSelection   =   0   'False
            Left            =   1200
            TabIndex        =   113
            Top             =   660
            Width           =   1395
         End
         Begin VB.TextBox txtCambioUbicacionOrden 
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
            HideSelection   =   0   'False
            Left            =   1200
            TabIndex        =   110
            Top             =   300
            Width           =   1395
         End
         Begin VB.CommandButton cmdCambioUbicacion 
            Caption         =   "Aceptar"
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
            Left            =   1980
            TabIndex        =   109
            Top             =   1140
            Width           =   1200
         End
         Begin VB.Label Label31 
            Caption         =   "Nº de Orden"
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
            Left            =   120
            TabIndex        =   112
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Ubicacion :"
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
            Index           =   6
            Left            =   120
            TabIndex        =   111
            Top             =   660
            Width           =   1275
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Reparacion Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   -69360
         TabIndex        =   102
         Top             =   1020
         Width           =   3375
         Begin VB.CommandButton cmdActualizarOrdenRemito 
            Caption         =   "Aceptar"
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
            Left            =   1980
            TabIndex        =   104
            Top             =   1140
            Width           =   1200
         End
         Begin VB.TextBox txtRepararOrdenRemito 
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
            HideSelection   =   0   'False
            Left            =   1200
            TabIndex        =   103
            Top             =   300
            Width           =   1395
         End
         Begin MSMask.MaskEdBox mskRemitoOrden 
            Height          =   375
            Left            =   1200
            TabIndex        =   107
            Top             =   660
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   13
            Mask            =   "0001-000#####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            Caption         =   "Ubicacion :"
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
            Index           =   5
            Left            =   120
            TabIndex        =   106
            Top             =   660
            Width           =   1275
         End
         Begin VB.Label Label30 
            Caption         =   "Nº de Orden"
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
            Left            =   120
            TabIndex        =   105
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Anular Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -74700
         TabIndex        =   98
         Top             =   5520
         Width           =   5235
         Begin VB.CommandButton cmdAnular 
            Caption         =   "Aceptar"
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
            Left            =   3840
            TabIndex        =   100
            Top             =   300
            Width           =   1200
         End
         Begin VB.TextBox txtOrdenAnular 
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
            HideSelection   =   0   'False
            Left            =   1200
            TabIndex        =   99
            Top             =   300
            Width           =   2535
         End
         Begin VB.Label Label29 
            Caption         =   "Nº de Orden"
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
            Left            =   120
            TabIndex        =   101
            Top             =   300
            Width           =   975
         End
      End
      Begin VB.Frame fraOrdenDocuementacion 
         Caption         =   "Orden de documentación"
         Height          =   6480
         Left            =   180
         TabIndex        =   66
         Top             =   840
         Width           =   10125
         Begin VB.ComboBox cboTipo_Orden 
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
            ItemData        =   "frmOrdenDocumentacion.frx":008C
            Left            =   5460
            List            =   "frmOrdenDocumentacion.frx":0096
            TabIndex        =   115
            Text            =   "Combo1"
            Top             =   360
            Width           =   3135
         End
         Begin VB.CommandButton cmdImpresionRearchivoDigital 
            Caption         =   "Rearchivo Digital"
            Height          =   435
            Left            =   6960
            TabIndex        =   114
            Top             =   2400
            Width           =   1995
         End
         Begin VB.CommandButton cmdCaracteresEstadistica 
            Caption         =   "Caracteres Est"
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
            Left            =   120
            TabIndex        =   97
            Top             =   5940
            Width           =   1260
         End
         Begin VB.CommandButton cmdCantCaracteres 
            Caption         =   "Caracteres"
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
            Left            =   1440
            TabIndex        =   78
            Top             =   5940
            Width           =   1200
         End
         Begin VB.CommandButton cmdRetiroDocumentacion 
            Caption         =   "Retiro"
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
            Left            =   5220
            TabIndex        =   77
            Top             =   5940
            Width           =   1200
         End
         Begin VB.ComboBox cboDescripcion 
            Height          =   345
            ItemData        =   "frmOrdenDocumentacion.frx":00A8
            Left            =   5400
            List            =   "frmOrdenDocumentacion.frx":00B5
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1200
            Width           =   3495
         End
         Begin VB.CommandButton cmdImprimirOrden 
            Caption         =   "Imprimir"
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
            Left            =   2700
            TabIndex        =   75
            Top             =   5940
            Width           =   1200
         End
         Begin VB.CommandButton cmdEnviarInformacion 
            Caption         =   "Acuses"
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
            Left            =   3960
            TabIndex        =   74
            Top             =   5940
            Width           =   1200
         End
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "Borrar"
            Height          =   375
            Left            =   3000
            TabIndex        =   73
            Top             =   2460
            Width           =   1200
         End
         Begin VB.TextBox txtSql 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8700
            TabIndex        =   72
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox txtCodigo 
            Height          =   330
            Left            =   1320
            TabIndex        =   71
            Top             =   1620
            Width           =   2775
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   4320
            TabIndex        =   70
            Top             =   2460
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cancelar"
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
            Left            =   6480
            TabIndex        =   69
            Top             =   5940
            Width           =   1200
         End
         Begin VB.CommandButton cmdAceptarDocumentacion 
            BackColor       =   &H80000018&
            Caption         =   "Aceptar"
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
            Left            =   7740
            TabIndex        =   68
            Top             =   5940
            Width           =   1200
         End
         Begin VB.TextBox TXTCAJA 
            Height          =   330
            Left            =   5400
            TabIndex        =   67
            Top             =   780
            Width           =   3435
         End
         Begin Controles.cltGenerico ctlClientesDocumento 
            Height          =   375
            Left            =   1320
            TabIndex        =   79
            Top             =   780
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
         End
         Begin Controles.cltGenerico ctlPersonalDOcumento 
            Height          =   375
            Left            =   1320
            TabIndex        =   80
            Top             =   360
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
         End
         Begin MSMask.MaskEdBox mskRemito 
            Height          =   330
            Left            =   1320
            TabIndex        =   81
            Top             =   1200
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   13
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "0001-000#####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskBuscar 
            Height          =   375
            Left            =   1320
            TabIndex        =   82
            Top             =   2460
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid grdOrdenDocumentacion 
            Height          =   2895
            Left            =   120
            TabIndex        =   83
            Top             =   2940
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   5106
            _Version        =   393216
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
         Begin VB.Label Label13 
            Caption         =   " Estado : "
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
            Index           =   4
            Left            =   4380
            TabIndex        =   94
            Top             =   1260
            Width           =   915
         End
         Begin VB.Label lblIndicador 
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
            Left            =   4200
            TabIndex        =   93
            Top             =   2640
            Width           =   2475
         End
         Begin VB.Label Label13 
            Caption         =   "Remito:"
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   92
            Top             =   1260
            Width           =   855
         End
         Begin VB.Label lblDescripcionCodigo 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1320
            TabIndex        =   91
            Top             =   2040
            Width           =   7635
         End
         Begin VB.Label lblCodigo 
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   4260
            TabIndex        =   90
            Top             =   1620
            Width           =   4695
         End
         Begin VB.Label Label12 
            Caption         =   "Desc.:"
            Height          =   315
            Left            =   180
            TabIndex        =   89
            Top             =   2160
            Width           =   915
         End
         Begin VB.Label Label11 
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
            Left            =   4380
            TabIndex        =   88
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label10 
            Caption         =   "Personal:"
            Height          =   315
            Left            =   180
            TabIndex        =   87
            Top             =   420
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Codigo:"
            Height          =   315
            Left            =   180
            TabIndex        =   86
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label Label4 
            Caption         =   "Cliente:"
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   85
            Top             =   900
            Width           =   1035
         End
         Begin VB.Label Label3 
            Caption         =   "Provisorio:"
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
            Left            =   4380
            TabIndex        =   84
            Top             =   900
            Width           =   915
         End
      End
      Begin VB.Frame fraOrdenLegajos 
         Caption         =   "Orden Legajos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6300
         Left            =   -74900
         TabIndex        =   59
         Top             =   700
         Width           =   9100
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
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
            Left            =   6660
            TabIndex        =   61
            Top             =   5880
            Width           =   1095
         End
         Begin VB.TextBox txtLecturaLegajos 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5880
            TabIndex        =   60
            Top             =   300
            Width           =   3135
         End
         Begin Controles.cltGenerico ctlPersonalOrden 
            Height          =   375
            Left            =   900
            TabIndex        =   62
            Top             =   300
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   661
         End
         Begin MSFlexGridLib.MSFlexGrid grdOrdenLegajos 
            Height          =   5115
            Left            =   180
            TabIndex        =   63
            Top             =   720
            Width           =   8835
            _ExtentX        =   15584
            _ExtentY        =   9022
            _Version        =   393216
         End
         Begin VB.Label Label2 
            Caption         =   "Personal:"
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
            Left            =   180
            TabIndex        =   65
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label1 
            Caption         =   "Lectura Etiqueta :"
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
            Left            =   4500
            TabIndex        =   64
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame fraBuscarOrdenamiento 
         Caption         =   "Buscar Documento"
         Height          =   6300
         Left            =   -74820
         TabIndex        =   52
         Top             =   780
         Width           =   9100
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "Actualizar"
            Height          =   375
            Left            =   6360
            TabIndex        =   116
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtElementoBuscar 
            Height          =   375
            Left            =   1080
            TabIndex        =   54
            Top             =   720
            Width           =   5175
         End
         Begin VB.CommandButton cmdCopiarExcel 
            Caption         =   "Copiar Excel"
            Height          =   375
            Left            =   7680
            TabIndex        =   53
            Top             =   720
            Width           =   1335
         End
         Begin Controles.cltGenerico ctlClientesBuscarDocumento 
            Height          =   375
            Left            =   1080
            TabIndex        =   55
            Top             =   300
            Width           =   7875
            _ExtentX        =   13891
            _ExtentY        =   661
         End
         Begin MSDataGridLib.DataGrid grdBuscarOrdenamiento 
            Bindings        =   "frmOrdenDocumentacion.frx":00D2
            Height          =   4995
            Left            =   120
            TabIndex        =   56
            Top             =   1200
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   8811
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   16
            FormatLocked    =   -1  'True
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
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DataMember      =   "Command1"
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "DESCRIPCION"
               Caption         =   "DESCRIPCIÓN"
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
               DataField       =   "ELEMENTO_NUMERO"
               Caption         =   "ELEMENTO"
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
            BeginProperty Column02 
               DataField       =   "COD_ESTADO"
               Caption         =   "ESTADO"
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
            BeginProperty Column03 
               DataField       =   "COD_DOCUMENTACION"
               Caption         =   "ORDEN"
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
            BeginProperty Column04 
               DataField       =   "CONTENEDOR_PROV"
               Caption         =   "PROV"
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
            BeginProperty Column05 
               DataField       =   "COD_NRO_CAJA"
               Caption         =   "Nº CAJA"
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
            BeginProperty Column06 
               DataField       =   "ORDEN"
               Caption         =   "POSI"
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
                  ColumnWidth     =   2505,26
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   705,26
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   705,26
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   794,835
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   900,284
               EndProperty
               BeginProperty Column06 
               EndProperty
            EndProperty
         End
         Begin VB.Label Label15 
            Caption         =   "Cliente:"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label16 
            Caption         =   "Elemento:"
            Height          =   375
            Left            =   120
            TabIndex        =   57
            Top             =   780
            Width           =   975
         End
      End
      Begin VB.Frame fraAfiliadoOsep 
         Caption         =   "Orden de Documentación Osep Legajos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6300
         Left            =   -74900
         TabIndex        =   23
         Top             =   700
         Width           =   9100
         Begin VB.CommandButton cmdControl 
            Caption         =   "Control"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4020
            TabIndex        =   35
            Top             =   5880
            Width           =   1200
         End
         Begin VB.TextBox txtCajaAfil 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5940
            TabIndex        =   34
            Top             =   780
            Width           =   3015
         End
         Begin VB.TextBox txtMax_COD_DOCUMENTACION 
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
            Left            =   900
            TabIndex        =   33
            Top             =   5880
            Width           =   1155
         End
         Begin VB.CommandButton cmdOsepActualizar 
            Caption         =   "Actualizar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   32
            Top             =   5880
            Width           =   1200
         End
         Begin VB.TextBox txtubicaprovosep 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3900
            TabIndex        =   31
            Top             =   780
            Width           =   1335
         End
         Begin VB.CommandButton cmdAceptarOsep 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   6540
            TabIndex        =   30
            Top             =   5880
            Width           =   1200
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   7800
            TabIndex        =   29
            Top             =   5880
            Width           =   1200
         End
         Begin VB.CommandButton cmdAceptarLegajos 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5160
            TabIndex        =   28
            Top             =   2040
            Width           =   1200
         End
         Begin VB.TextBox txtCodigoAfiliado 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   27
            Text            =   "002008002"
            Top             =   1200
            Width           =   2775
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5940
            TabIndex        =   26
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Borrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3900
            TabIndex        =   25
            Top             =   2040
            Width           =   1200
         End
         Begin VB.CommandButton cmdOsepImprimir 
            Caption         =   "Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5280
            TabIndex        =   24
            Top             =   5880
            Width           =   1200
         End
         Begin Controles.cltGenerico ctlPersonalafiadosOsep 
            Height          =   375
            Left            =   1080
            TabIndex        =   36
            Top             =   360
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   661
         End
         Begin MSMask.MaskEdBox mskRemitoOsep 
            Height          =   375
            Left            =   1080
            TabIndex        =   37
            Top             =   780
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   13
            Mask            =   "0001-000#####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDocumento 
            Height          =   375
            Left            =   1080
            TabIndex        =   38
            Top             =   2040
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSFlexGridLib.MSFlexGrid grdAfiliados 
            Height          =   3315
            Left            =   480
            TabIndex        =   39
            Top             =   2880
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   5847
            _Version        =   393216
         End
         Begin VB.Label Label25 
            Caption         =   "Caja:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   51
            Top             =   840
            Width           =   675
         End
         Begin VB.Label Label24 
            Caption         =   "Provisorio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   50
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "Codigo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   49
            Top             =   1260
            Width           =   915
         End
         Begin VB.Label Label22 
            Caption         =   "Personal:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   48
            Top             =   420
            Width           =   975
         End
         Begin VB.Label Label21 
            Caption         =   "SQL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            TabIndex        =   47
            Top             =   420
            Width           =   495
         End
         Begin VB.Label Label20 
            Caption         =   "Desc.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   46
            Top             =   1680
            Width           =   915
         End
         Begin VB.Label Label19 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3900
            TabIndex        =   45
            Top             =   1200
            Width           =   5055
         End
         Begin VB.Label Label18 
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1080
            TabIndex        =   44
            Top             =   1620
            Width           =   7875
         End
         Begin VB.Label Label13 
            Caption         =   "Remito:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   43
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label17 
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
            Left            =   6300
            TabIndex        =   42
            Top             =   2040
            Width           =   2475
         End
         Begin VB.Label Label26 
            Caption         =   "Item:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   41
            Top             =   2040
            Width           =   915
         End
         Begin VB.Label Label27 
            Caption         =   "Nº Doc.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   5880
            Width           =   915
         End
      End
      Begin VB.Frame fraReparacionOrden 
         Caption         =   "Reparacion Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -74700
         TabIndex        =   16
         Top             =   900
         Width           =   5235
         Begin VB.TextBox txtUbicacion 
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
            Left            =   1200
            TabIndex        =   20
            Top             =   660
            Width           =   3855
         End
         Begin VB.TextBox txtOrdenReparacion 
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
            HideSelection   =   0   'False
            Left            =   1200
            TabIndex        =   19
            Top             =   300
            Width           =   3855
         End
         Begin VB.TextBox txtReparacionElemento 
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
            Left            =   1200
            TabIndex        =   18
            Top             =   1020
            Width           =   2595
         End
         Begin VB.CommandButton cmdAceptarReparacionOrden 
            BackColor       =   &H000000FF&
            Caption         =   "Aceptar"
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
            Left            =   3840
            TabIndex        =   17
            Top             =   1020
            Width           =   1200
         End
         Begin VB.Label Label28 
            Caption         =   "Nº de Orden"
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
            Left            =   120
            TabIndex        =   96
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Ubicacion :"
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
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   660
            Width           =   1275
         End
         Begin VB.Label Label13 
            Caption         =   "Elemento:"
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
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   1020
            Width           =   1275
         End
      End
      Begin VB.Frame fraOrdenControl 
         Caption         =   "Control de ordenamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -74700
         TabIndex        =   8
         Top             =   4020
         Width           =   5235
         Begin VB.CommandButton cmdControlTerminado 
            Caption         =   "Control Terminado"
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
            Left            =   2280
            TabIndex        =   12
            Top             =   1020
            Width           =   1455
         End
         Begin VB.CommandButton cmdPendientes 
            Caption         =   "Pendientes"
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
            Left            =   3900
            TabIndex        =   11
            Top             =   1020
            Width           =   1155
         End
         Begin VB.CommandButton cmdGenerar 
            Caption         =   "Generar"
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
            Left            =   1080
            TabIndex        =   10
            Top             =   1020
            Width           =   1095
         End
         Begin VB.TextBox txtOrdenControl 
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
            Left            =   1140
            TabIndex        =   9
            Top             =   660
            Width           =   3915
         End
         Begin Controles.cltGenerico ctlPersonalControl 
            Height          =   315
            Left            =   1140
            TabIndex        =   13
            Top             =   300
            Width           =   3915
            _ExtentX        =   6906
            _ExtentY        =   556
         End
         Begin VB.Label Label9 
            Caption         =   "Personal"
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
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label8 
            Caption         =   "Nº de Orden"
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
            Left            =   120
            TabIndex        =   14
            Top             =   660
            Width           =   975
         End
      End
      Begin VB.Frame fraOrdenTerminado 
         Caption         =   "Asignacion de responsable de ordenamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -74700
         TabIndex        =   1
         Top             =   2460
         Width           =   5235
         Begin VB.TextBox txtFechaOrdenTerminado 
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
            Left            =   1140
            TabIndex        =   4
            Top             =   600
            Width           =   3975
         End
         Begin VB.CommandButton cmdAceptarOrdenTerminado 
            Caption         =   "Aceptar"
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
            Left            =   4080
            TabIndex        =   3
            Top             =   960
            Width           =   1035
         End
         Begin VB.TextBox txtOrdenOrdenamientoTerminado 
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
            Left            =   1140
            TabIndex        =   2
            Top             =   960
            Width           =   2895
         End
         Begin Controles.cltGenerico ctlPersonalOrdenTerminado 
            Height          =   375
            Left            =   1140
            TabIndex        =   5
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   661
         End
         Begin VB.Label Label7 
            Caption         =   "Nº de Orden"
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
            Left            =   180
            TabIndex        =   95
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Fecha :"
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
            Left            =   180
            TabIndex        =   7
            Top             =   600
            Width           =   675
         End
         Begin VB.Label Label6 
            Caption         =   "Personal"
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
            Left            =   180
            TabIndex        =   6
            Top             =   300
            Width           =   795
         End
      End
   End
End
Attribute VB_Name = "frmOrdenDocumentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsBuscar As ADODB.Recordset
    Dim ConOsep As ADODB.Connection
Private Sub txtBuscarLegajo_Change()

End Sub

Public Sub TitulosGrillaOrdenLegajos()
     
    With grdOrdenLegajos
            .Cols = 4
            .Rows = 1
            .ColAlignment(1) = 4
            .ColAlignment(2) = 4
            .ColAlignment(3) = 4
            .ColWidth(0) = 100
            .ColWidth(1) = 2000
            .ColWidth(2) = 2500
            .ColWidth(3) = 2500
            .TextMatrix(0, 1) = "Cliente"
            .TextMatrix(0, 2) = "Etiqueta"
    End With

End Sub
Public Sub TitulosGrillaOrdenDocumentacion()
     
    With grdOrdenDocumentacion
            .Cols = 7
            .Rows = 1
            .ColAlignment(1) = 4
            .ColAlignment(2) = 4
            .ColAlignment(3) = 4
            .ColWidth(0) = 400
            .ColWidth(1) = 2000
            .ColWidth(2) = 2000
            .ColWidth(3) = 2000
            .ColWidth(4) = 4000
            .ColWidth(5) = 2000
            .TextMatrix(0, 1) = "Cliente"
            .TextMatrix(0, 2) = "Caja"
            .TextMatrix(0, 3) = "Elemento"
            .TextMatrix(0, 4) = "Descripcion"
            .TextMatrix(0, 5) = "Provi"
            .TextMatrix(0, 6) = "Fecha"
    End With

End Sub
Public Sub TitulosGrillaAfiliados()
    With grdAfiliados
            .Cols = 7
            .Rows = 1
            .ColAlignment(1) = 4
            .ColAlignment(2) = 4
            .ColAlignment(3) = 4
            .ColWidth(0) = 500
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 2000
            .ColWidth(4) = 3000
            .ColWidth(5) = 2000
            .TextMatrix(0, 1) = "Cliente"
            .TextMatrix(0, 2) = "Caja"
            .TextMatrix(0, 3) = "Elemento"
            .TextMatrix(0, 4) = "Nombre"
            .TextMatrix(0, 5) = "Provi"
            .TextMatrix(0, 5) = "Fecha"
    End With

End Sub
Private Sub cndBuscar_Click()

End Sub

Private Sub cboDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If

End Sub

Private Sub cmdAceptarDocumentacion_Click()

    Dim Sql As String
    Dim COD_CLIENTE, Cod_Indice As String
    Dim Cod_Nro_Caja, Elemento, Elemento_String, Cod_Tipo_Orden, Descripcion, Orden, Contenedor_Prov As String
    Dim Max_Cod_Documentacion As Long
    Dim Elemento_Numero As Long
    Dim CANTIDAD_CARACTERES As Integer
    Dim errorDes As Stream
    On Error GoTo salir:

    
    Dim rs As New ADODB.Recordset
    If Trim(txtCaja.Text) = "" Then
        MsgBox "Ingrese el ubicación provisoria"
        Exit Sub
    End If
    If cboDescripcion.ListIndex = -1 Then
        MsgBox "Ingrese el Tipo de Ingreso ", vbInformation
        Exit Sub
    End If
    If IsNull(ctlPersonalDOcumento.Valor) Then
        MsgBox "Ingrese el personal", vbInformation
        Exit Sub
    End If
    If mskRemito.Text = "0001-000_____" Then
            MsgBox "Ingrese el Nº de Remito", vbInformation
            Exit Sub
    End If
    
    If cboTipo_Orden.Text = "" Then
            MsgBox "TIPO DE ORDEN", vbInformation
            Exit Sub
    End If
    Sql = " SELECT MAX(ID_ORDENAR_DOCUMENTACION) AS MaxDocu From ORDENAR_DOCUMENTACION "
    rs.Open Sql, ConActiva, 0, 1
    Max_Cod_Documentacion = rs!MaxDocu + 1
    
    
Dim i As Integer
With grdOrdenDocumentacion
    For i = 1 To .Rows - 1
        COD_CLIENTE = .TextMatrix(i, 1)
        Cod_Indice = "'" & lblCodigo.Caption & "'"
        Cod_Nro_Caja = .TextMatrix(i, 2)
        Elemento = "'" & Trim(.TextMatrix(i, 3)) & "'"
        If IsNumeric(.TextMatrix(i, 3)) Then
            Elemento_Numero = CLng(.TextMatrix(i, 3))
            Elemento_String = "Null"
        Else
            Elemento_String = "'" & (.TextMatrix(i, 3)) & "'"
            Elemento_Numero = Null
        End If
        CANTIDAD_CARACTERES = Len(Trim(.TextMatrix(i, 3)))
        Cod_Tipo_Orden = 1
        Descripcion = .TextMatrix(i, 4)
        
        Orden = .TextMatrix(i, 0)
        Contenedor_Prov = "'" & .TextMatrix(i, 5) & "'"
        Sql = " INSERT INTO ORDENAR_DOCUMENTACION_DETALLE"
        Sql = Sql & vbCrLf & " (COD_DOCUMENTACION, COD_CLIENTE, COD_INDICE,"
        Sql = Sql & vbCrLf & " COD_NRO_CAJA, ELEMENTO, COD_TIPO_ORDEN,DESCRIPCION,CONTENEDOR_PROV,ORDEN ,COD_ESTADO,ELEMENTO_STRING, ELEMENTO_NUMERO)"
        Sql = Sql & vbCrLf & "  VALUES ( "
        Sql = Sql & vbCrLf & Max_Cod_Documentacion & "," & COD_CLIENTE & "," & Cod_Indice & ","
        Sql = Sql & vbCrLf & Cod_Nro_Caja & "," & Elemento & "," & Cod_Tipo_Orden & ",'" & Descripcion & "'," & Contenedor_Prov & "," & Orden & " ,0," & Elemento_String & "," & Elemento_Numero & " )"
        ExecutarSql Sql
    Next
        Sql = " INSERT INTO ORDENAR_DOCUMENTACION  (COD_CLIENTE , ID_ORDENAR_DOCUMENTACION, FECHA, COD_REMITO_PRO , COD_RESPONSABLE_CARGA, COD_ESTADO,DESCRIPCION,CANTIDAD,COD_TIPO_ORDEN)"
        Sql = Sql & vbCrLf & " VALUES (" & ctlClientesDocumento.Valor & "," & Max_Cod_Documentacion & "," & SysDate & ",'" & mskRemito.Text & "'," & ctlPersonalDOcumento.Valor & "," & 0 & ",'" & Trim(cboDescripcion.Text) & "'," & i & ",'" & cboTipo_Orden.Text & "')"
        ExecutarSql Sql
        cboDescripcion.ListIndex = -1

End With
    InsertarProducion ctlPersonalDOcumento.Valor, 9, "CARGA ORDEN:" & Max_Cod_Documentacion, grdOrdenDocumentacion.Rows - 1, ctlClientesDocumento.Valor
    MsgBox "Orden Numero : " & Max_Cod_Documentacion
    ImprimirOrdenDocumentacion Max_Cod_Documentacion
    TitulosGrillaOrdenDocumentacion
    mskBuscar.Mask = ""
    LimpiarMask mskBuscar
    lblCodigo.Caption = ""
    lblDescripcionCodigo = ""
    MousePointer = 0
    Exit Sub
salir:
    MsgBox Err.Description

End Sub

Private Sub cmdAceptarLegajos_Click()
    Dim rs As ADODB.Recordset
    On Error GoTo salir
If IsNull(ctlPersonalafiadosOsep.Valor) Then
        MsgBox "Ingerese el personal", vbInformation
    Exit Sub
End If
        Set rs = New ADODB.Recordset
        Dim Sql As String
        Sql = " SELECT IDAFILIADO, TIPO_DOC_AFI, NUMERO_AFI, VINCULO,"
        Sql = Sql & vbCrLf & " APELLIDO_NOMBRE, TIPO_DOC, DOCUMENTO, SITUACION,"
        Sql = Sql & vbCrLf & "DESCRIPCION "
        Sql = Sql & vbCrLf & " From OSEPAFILI"
        Sql = Sql & vbCrLf & " Where (VINCULO = 0) And NUMERO_AFI = " & mskDocumento.Text
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic
        rs.Open Sql, ConOsep
        Select Case rs.RecordCount
        Case 0
           MsgBox "No se encontro Registro"
        Case 1
           grdAfiliados.AddItem grdAfiliados.Rows & vbTab & 20 & vbTab & txtCajaAfil.Text & vbTab & mskDocumento.Text & vbTab & rs!APELLIDO_NOMBRE & vbTab & txtubicaprovosep.Text & vbTab & Now
           grdAfiliados.TopRow = grdAfiliados.Rows - 1
           grdAfiliados_Scroll
        Case Is > 1
           grdAfiliados.AddItem grdAfiliados.Rows & vbTab & 20 & vbTab & txtCajaAfil.Text & vbTab & mskDocumento.Text & vbTab & rs!APELLIDO_NOMBRE & vbTab & txtubicaprovosep.Text & Now
           grdAfiliados.TopRow = grdAfiliados.Rows - 1
           grdAfiliados_Scroll
        End Select
        LimpiarMask mskDocumento
        mskDocumento.SetFocus
        Exit Sub
salir:
        MsgBox Err.Description
       
End Sub

Private Sub cmdAceptarOrdenTerminado_Click()
        Dim Sql As String
        Dim Registros As Integer
        Dim rs As New ADODB.Recordset
        Dim cantidad As Integer

            Sql = " Update ORDENAR_DOCUMENTACION SET COD_RESPONSABLE_ORDEN =" & ctlPersonalOrdenTerminado.Valor & " , Cod_Estado = 2"
            Sql = Sql & " Where (Cod_Estado = 0 and  ID_ORDENAR_DOCUMENTACION =" & txtOrdenOrdenamientoTerminado.Text & ")"
            ExecutarSql Sql
            Sql = " Update ORDENAR_DOCUMENTACION_DETALLE Set Cod_Estado = 2"
            Sql = Sql & " Where (Cod_Estado = 0 and  COD_DOCUMENTACION =" & txtOrdenOrdenamientoTerminado.Text & ")"
            Registros = ExecutarSql(Sql)
            cantidad = Registros
            Set rs = New ADODB.Recordset
            rs.Open "SELECT COD_CLIENTE From ORDENAR_DOCUMENTACION_DETALLE Where COD_DOCUMENTACION = " & txtOrdenOrdenamientoTerminado.Text, ConActiva, 0, 1
            If cantidad <> 0 Then
                InsertarProducion ctlPersonalOrdenTerminado.Valor, 15, "ORDEN Nº:" & txtOrdenOrdenamientoTerminado.Text, cantidad, rs!COD_CLIENTE, txtFechaOrdenTerminado.Text
            Else
                MsgBox "Actualizacion no relizada", vbInformation
                Exit Sub
            End If
            txtOrdenOrdenamientoTerminado.Text = ""
            MsgBox "La asignacion fue realizada", vbInformation
           
           Rem  ImprimirOrdenDocumentacion txtOrdenOrdenamientoTerminado.Text
           
            
End Sub

Private Sub cmdAceptarOsep_Click()
Dim Sql As String
    Dim COD_CLIENTE, Cod_Indice As String
    Dim Cod_Nro_Caja, Elemento, Elemento_String, Cod_Tipo_Orden, Descripcion, Orden, Contenedor_Prov, FechaCarga As String
    Dim Max_Cod_Documentacion As Long
    Dim Elemento_Numero As Long
    Dim rs As New ADODB.Recordset
    Dim CANTIDAD_CARACTERES As Integer
    On Error GoTo salir:

    If Trim(txtubicaprovosep.Text) = "" Then
        MsgBox "Ingrese el ubicación provisoria"
        Exit Sub
    End If
    Sql = " SELECT MAX(ID_ORDENAR_DOCUMENTACION) AS MaxDocu From ORDENAR_DOCUMENTACION "
    rs.Open Sql, ConActiva, 0, 1
    Max_Cod_Documentacion = rs!MaxDocu + 1
    
Dim i As Integer
With grdAfiliados
    For i = 1 To .Rows - 1
        COD_CLIENTE = .TextMatrix(i, 1)
        
        Cod_Nro_Caja = .TextMatrix(i, 2)
        Elemento = "'" & Trim(.TextMatrix(i, 3)) & "'"
        If IsNumeric(.TextMatrix(i, 3)) Then
            Elemento_Numero = CLng(.TextMatrix(i, 3))
            Elemento_String = "Null"
        Else
            Elemento_String = "'" & (.TextMatrix(i, 3)) & "'"
            Elemento_Numero = Null
        End If
        CANTIDAD_CARACTERES = Len(Trim(.TextMatrix(i, 3)))
        Cod_Tipo_Orden = 1
        Descripcion = Trim(Replace(.TextMatrix(i, 4), "'", "´"))
        Orden = .TextMatrix(i, 0)
        Contenedor_Prov = "'" & .TextMatrix(i, 5) & "'"
        FechaCarga = FechaFormato(.TextMatrix(i, 6))
        
        
        
        Sql = " INSERT INTO ORDENAR_DOCUMENTACION_DETALLE"
        Sql = Sql & vbCrLf & " (COD_DOCUMENTACION, COD_CLIENTE, COD_INDICE,"
        Sql = Sql & vbCrLf & " COD_NRO_CAJA, ELEMENTO, COD_TIPO_ORDEN,DESCRIPCION,CONTENEDOR_PROV,ORDEN ,COD_ESTADO,ELEMENTO_STRING, ELEMENTO_NUMERO,CANTIDAD_CARACTERES, FECHA_ACTUALIZACION )"
        Sql = Sql & vbCrLf & "  VALUES ( "
        Sql = Sql & vbCrLf & Max_Cod_Documentacion & "," & COD_CLIENTE & ",'002008002',"
        Sql = Sql & vbCrLf & txtCajaAfil.Text & "," & Elemento & "," & Cod_Tipo_Orden & ",'" & Mid(Descripcion, 1, 50) & "'," & Contenedor_Prov & "," & Orden & " ,0," & Elemento_String & "," & Elemento_Numero & ", " & CANTIDAD_CARACTERES
        Sql = Sql & vbCrLf & "," & FechaCarga & ")"
        ExecutarSql Sql
    Next
    Sql = " INSERT INTO ORDENAR_DOCUMENTACION  (COD_CLIENTE,ID_ORDENAR_DOCUMENTACION, FECHA, COD_REMITO_PRO , COD_RESPONSABLE_CARGA, COD_ESTADO,CANTIDAD,COD_TIPO_ORDEN)"
    Sql = Sql & vbCrLf & " VALUES ( 20 , " & Max_Cod_Documentacion & "," & SysDate & ",'" & mskRemitoOsep.Text & "'," & ctlPersonalafiadosOsep.Valor & "," & 0 & "," & i & ",'LOTE')"
    ExecutarSql Sql
End With

InsertarProducion ctlPersonalafiadosOsep.Valor, 9, "CARGA ORDEN:" & Max_Cod_Documentacion, grdAfiliados.Rows - 1, 20
MsgBox "Orden Numero : " & Max_Cod_Documentacion
MousePointer = 11

ImprimirOrdenDocumentacion Max_Cod_Documentacion



TitulosGrillaAfiliados
mskDocumento.Mask = ""
LimpiarMask mskDocumento
LimpiarMask mskRemitoOsep
txtubicaprovosep.Text = ""
txtCajaAfil.Text = ""
MousePointer = 0
Exit Sub
salir:
MousePointer = 0
MsgBox Err.Description & "Para el doc " & Elemento & " en la pos" & i

MousePointer = 0
End Sub

Private Sub cmdAceptarReparacionOrden_Click()
    Dim Sql As String
    Dim Afec As Integer
    If IsNumeric(txtReparacionElemento.Text) Then
        Sql = " Update ORDENAR_DOCUMENTACION_DETALLE"
        Sql = Sql & " SET ELEMENTO = '" & Trim(txtReparacionElemento.Text) & "'"
        Sql = Sql & " ,ELEMENTO_NUMERO= '" & Trim(txtReparacionElemento.Text) & "'"
        Sql = Sql & " Where COD_DOCUMENTACION =" & txtOrdenReparacion.Text & " And ORDEN = " & txtUbicacion.Text
    Else
        Sql = " Update ORDENAR_DOCUMENTACION_DETALLE"
        Sql = Sql & " SET ELEMENTO = '" & Trim(txtReparacionElemento.Text) & "'"
        Sql = Sql & " Where COD_DOCUMENTACION =" & txtOrdenReparacion.Text & " And ORDEN = " & txtUbicacion.Text
    End If
    
    Afec = ExecutarSql(Sql)
    If Afec = 0 Then
        MsgBox "No se relizo la actualizacion", vbCritical
    End If
    txtOrdenReparacion.Text = ""
    txtUbicacion.Text = ""
    txtReparacionElemento.Text = ""
End Sub

Private Sub cmdActualizar_Click()
Dim Sql As String
Sql = " Update dbo.ORDENAR_DOCUMENTACION_DETALLE"
Sql = Sql & " SET              NRO_DESDE = ELEMENTO, NRO_HASTA = ELEMENTO"
Sql = Sql & "  Where (NRO_DESDE Is Null) And (Not (ELEMENTO Is Null))"

ExecutarSql Sql
End Sub

Private Sub cmdActualizarOrdenRemito_Click()
    Dim conR As New ADODB.Connection
    Set conR = ConBasa
    Dim Sql As String

If txtRepararOrdenRemito.Text <> "" Then
    Sql = " Update ORDENAR_DOCUMENTACION "
    Sql = Sql & " SET COD_REMITO_PRO ='" & mskRemitoOrden.Text & "'"
    Sql = Sql & "  Where ID_ORDENAR_DOCUMENTACION = " & txtRepararOrdenRemito.Text
    conR.Execute Sql
    MsgBox "La actualizacion se realizo con Exito", vbInformation
    End If
End Sub


Private Sub cmdAnular_Click()
    Dim Sql As String
    Sql = " Update ORDENAR_DOCUMENTACION  SET ANULADO = '1'"
    Sql = Sql & " Where ID_ORDENAR_DOCUMENTACION = " & txtOrdenAnular.Text
    ExecutarSql Sql
    MsgBox "La orden a sido anulada", vbInformation
End Sub

Private Sub cmdBorrar_Click()
    Dim NUMERO As Integer
        NUMERO = InputBox("Ingrese el Numero de fila a borrar")
        grdOrdenDocumentacion.RemoveItem NUMERO
End Sub

Private Sub cmdCambioUbicacion_Click()


Dim Sql As String

Sql = " Update ORDENAR_DOCUMENTACION_DETALLE"
Sql = Sql & " SET CONTENEDOR_PROV ='" & txtCambioUbicacionUbicacion.Text & "'"
Sql = Sql & " Where COD_DOCUMENTACION =" & txtCambioUbicacionOrden.Text
ExecutarSql Sql
MsgBox "La actualización de completo"
End Sub

Private Sub cmdCantCaracteres_Click()
    Dim fecha As String
    Dim Sql As String
        
       If Not IsNull(ctlPersonalDOcumento.Valor) Then
            fecha = InputBox("Ingrese la Fecha", "Control de Carga", Format(Now, "dd/mm/yyyy"))
            MousePointer = 11
       

                Sql = " SELECT *"
                Sql = Sql & " FROM V_ORDEN_DOCUMENTACION_CARGA "
                Sql = Sql & " where COD_RESPONSABLE_CARGA = " & ctlPersonalDOcumento.Valor
                Sql = Sql & " AND FECHA > " & FechaServerTipo(fecha)
                Sql = Sql & " AND FECHA < " & FechaServerTipo(DateAdd("d", 1, fecha))
                Sql = Sql & " ORDER BY  COD_RESPONSABLE_CARGA , FECHA ,ID_ORDENAR_DOCUMENTACION"
            frmReportes.ImprimirReporte PasoReportes & "rptOrdenarDocumentacionCantidadCaracteres.RPT", Sql, True
            MousePointer = 0
        Else
            MsgBox "Ingrese el responsable", vbInformation
        End If
End Sub

Private Sub cmdCaracteresEstadistica_Click()
 Dim Sql As String
    Dim FechaInicio As String
    Dim FechaFin As String
        If Not IsNull(ctlPersonalDOcumento.Valor) Then
            FechaInicio = InputBox("Ingrese la Fecha Inicio de Control ", "Control Carga", DateAdd("d", -7, Format(Now, "DD/mm/yyyy")))
            FechaFin = InputBox("Ingrese la Fecha de Control Fin  ", "Control Carga", Format(Now, "DD/mm/yyyy"))
                MousePointer = 11
                Sql = " SELECT * "
                Sql = Sql & " FROM V_ORDEN_DOCUMENTACION_CARGA "
                Sql = Sql & " where COD_RESPONSABLE_CARGA = " & ctlPersonalDOcumento.Valor
                Sql = Sql & " and FECHA > " & FechaServerTipo(FechaInicio)
                Sql = Sql & " AND FECHA < " & FechaServerTipo(FechaFin)
                Sql = Sql & " ORDER BY  COD_RESPONSABLE_CARGA , FECHA ,ID_ORDENAR_DOCUMENTACION"
                frmReportes.ImprimirReporte PasoReportes & "rptOrdenarDocumentacionCantidadCaracteresEstadistica.RPT", Sql, True
                MousePointer = 0
        Else
            MsgBox "Ingrese el responsable", vbInformation
        End If
End Sub

Private Sub cmdControl_Click()
    Dim Sql As String
 If mskRemitoOsep.Text <> "" Then
    Sql = Sql & vbCrLf & " SELECT *"
    Sql = Sql & vbCrLf & "  From V_RETIRO_OSEP"
    Sql = Sql & vbCrLf & "  where COD_REMITO_PRO = '" & mskRemitoOsep.Text & "'"
    Sql = Sql & vbCrLf & " Order By COD_INDICE ASC, ID_ORDENAR_DOCUMENTACION ASC,ELEMENTO_NUMERO Asc "
    frmReportes.ImprimirReporte PasoReportes & "rptRetiroDocumentacionOsep.rpt", Sql, True
 Else
    MsgBox "Ingrese el Remito"
 End If
    
    
End Sub

Private Sub cmdControlTerminado_Click()
        Dim Clave As Integer
        Dim Sql As String
        Dim rs As ADODB.Recordset
        Rem Clave = InputBox("Ingrese Clave")
        Dim cantidad As Integer
        Dim Cliente As Integer
        Clave = 0
        If Clave = 0 Then
          Sql = " Update ORDENAR_DOCUMENTACION SET COD_ESTADO = 6,  COD_RESPONSABLE_CONTROL =" & ctlPersonalControl.Valor
          Sql = Sql & " Where ( ID_ORDENAR_DOCUMENTACION = " & txtOrdenControl.Text & ") And (Cod_Estado = 4)"
          ExecutarSql Sql
          Set rs = New ADODB.Recordset
            Sql = " UPDATE ORDENAR_DOCUMENTACION_DETALLE Set Cod_Estado = 6 WHERE (COD_ESTADO = 4) AND COD_DOCUMENTACION = " & txtOrdenControl.Text
             cantidad = ExecutarSql(Sql)
            If cantidad = 0 Then
                MsgBox "ERROR ACTUALIZACION", vbCritical
                Exit Sub
            End If
            InsertarProducion ctlPersonalControl.Valor, 14, "CONTROL ORDEN:" & txtOrdenControl.Text, cantidad, Cliente
            txtOrdenControl.Text = ""
            MsgBox "La orden esta terminada ", vbInformation
        Else
          MsgBox "Clave incorrecta", vbCritical
        End If
End Sub

Private Sub cmdCopiarExcel_Click()
    CopiarDatosGrilla grdBuscarOrdenamiento
End Sub

Private Sub cmdEnviarInformacion_Click()

   Dim rsbasa  As New ADODB.Recordset
   Dim Sql As String
   On Error GoTo er
   MousePointer = 11
        Sql = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE,DESCRIPCION"
        Sql = Sql & "  From INDICES"
        Sql = Sql & "  Where (COD_CLIENTE = 40)"
        Sql = Sql & "  ORDER BY INDICE"
        
        
    
    Dim DATO As String
    rsbasa.Open Sql, ConActiva, 0, 1
    
    
    
    Dim oConn As ADODB.Connection
    Set oConn = New ADODB.Connection
    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=c:\Acuses Montemar " & Format(date, "dd_mm_yyyy") & ".xls;" & _
               "Extended Properties=""Excel 8.0;HDR=NO;"""
    
    'Create a new table (or worksheet in the workbook)
    oConn.Execute "create table Indice  (Uno Char(255),dos Char(255),tres Char(255),cuatro Char(255),cinco Char(255), SEIS Char(255),SIETE Char(255),OCHO Char(255))"
    oConn.Execute "create table referencia  (Uno Char(255),dos Char(255),tres Char(255),cuatro Char(255),cinco Char(255))"
    Dim oRS As ADODB.Recordset
    Set oRS = New ADODB.Recordset
    oRS.Open "Select * from Indice", oConn, adOpenKeyset, adLockOptimistic
Dim i As Integer
oRS.MoveFirst

   Do While Not rsbasa.EOF
        
        Select Case Len(rsbasa!Indice) / 3
        Case 1
         oRS.Fields(0).value = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & " " & rsbasa!Descripcion
        Case 2
          oRS.Fields(0).value = "-"
          oRS.Fields(1).value = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & " " & rsbasa!Descripcion
          Case 3
          oRS.Fields(0).value = "-"
          oRS.Fields(1).value = "-"
          oRS.Fields(2).value = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & " " & rsbasa!Descripcion
          Case 4
          oRS.Fields(0).value = "-"
          oRS.Fields(1).value = "-"
          oRS.Fields(2).value = "-"
          oRS.Fields(3).value = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & " " & rsbasa!Descripcion
          Case 5
                oRS.Fields(0).value = "-"
                oRS.Fields(1).value = "-"
                oRS.Fields(2).value = "-"
                oRS.Fields(3).value = "-"
                oRS.Fields(4).value = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & " " & rsbasa!Descripcion
          Case 6
                oRS.Fields(0).value = "-"
                oRS.Fields(1).value = "-"
                oRS.Fields(2).value = "-"
                oRS.Fields(3).value = "-"
                oRS.Fields(4).value = "-"
                oRS.Fields(5).value = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & " " & rsbasa!Descripcion
          Case 7
           
                oRS.Fields(0).value = "-"
                oRS.Fields(1).value = "-"
                oRS.Fields(2).value = "-"
                oRS.Fields(3).value = "-"
                oRS.Fields(4).value = "-"
                oRS.Fields(5).value = "-"
                oRS.Fields(6).value = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & " " & rsbasa!Descripcion
           
          Case 8
                oRS.Fields(0).value = "-"
                oRS.Fields(1).value = "-"
                oRS.Fields(2).value = "-"
                oRS.Fields(3).value = "-"
                oRS.Fields(4).value = "-"
                oRS.Fields(5).value = "-"
                oRS.Fields(5).value = "-"
                oRS.Fields(6).value = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & " " & rsbasa!Descripcion
          End Select
          
        oRS.Update
        oRS.MoveNext
        rsbasa.MoveNext
    Loop
rsbasa.Close
oConn.Close
MousePointer = 0
Exit Sub
er:
MousePointer = 0
MsgBox "VERIFICAR LA EXISTENCIA DEL ARCHIVO"
End Sub

Private Sub cmdGenerar_Click()
        Dim rs As New ADODB.Recordset
        Dim i As Integer
        Dim Porcentaje As Integer
        Dim Valor As String
        Dim Sql As String
         rs.Open "SELECT MAX(ORDEN)as MaxOrden From ORDENAR_DOCUMENTACION_DETALLE WHERE ( COD_ESTADO = 2 AND  COD_DOCUMENTACION = " & txtOrdenControl.Text & ")", ConActiva, 0, 1
        If Not rs.EOF Then
            If IsNull(rs!MaxOrden) Then
            MsgBox "ERROR ORDEN"
            Exit Sub
            End If
            Porcentaje = 0.1 * rs!MaxOrden
            For i = 0 To Porcentaje
                Valor = Valor & Int((rs!MaxOrden * Rnd) + 1) & ","
            Next
        End If
        Valor = Mid(Valor, 1, Len(Valor) - 1)
        Sql = " Update ORDENAR_DOCUMENTACION_DETALLE Set Cod_Estado = 4 "
        Sql = Sql & " WHERE (COD_DOCUMENTACION = " & txtOrdenControl.Text & ") AND (ORDEN IN (" & Valor & "))"
        ExecutarSql Sql
        
        Sql = " Update ORDENAR_DOCUMENTACION"
        Sql = Sql & vbCrLf & " SET COD_ESTADO = 4, COD_RESPONSABLE_CONTROL =" & ctlPersonalControl.Valor
        Sql = Sql & vbCrLf & " WHERE (COD_ESTADO = 2) AND  ID_ORDENAR_DOCUMENTACION = " & txtOrdenControl.Text
        ExecutarSql Sql
        
        Sql = " SELECT * "
        Sql = Sql & " FROM V_ORDEN_DOCUMENTACION "
        Sql = Sql & " Where ID_ORDENAR_DOCUMENTACION =" & txtOrdenControl.Text
        Sql = Sql & " AND DETALLE_COD_ESTADO = 4"
        Sql = Sql & " ORDER BY ID_ORDENAR_DOCUMENTACION"
        
        frmReportes.ImprimirReporte PasoReportes & "rptOrdenarDocumentacion.rpt", Sql, True

        
        

End Sub

Private Sub cmdImpresionRearchivoDigital_Click()
    Dim Sql As String
    
    MousePointer = 11
        Sql = "  SELECT V_REARCHIVO_DIGITAL.COD_REMITO, "
        Sql = Sql & vbCrLf & " V_REARCHIVO_DIGITAL.ID_CODIGO_DOCUMENTO,"
        Sql = Sql & vbCrLf & " V_REARCHIVO_DIGITAL.DESCRIPCION, V_REARCHIVO_DIGITAL.ELEMENTO"
        Sql = Sql & vbCrLf & " FROM   BASA.V_REARCHIVO_DIGITAL V_REARCHIVO_DIGITAL"
        Sql = Sql & vbCrLf & " where V_REARCHIVO_DIGITAL.COD_REMITO ='" & InputBox("Ingrese el numero de legajo") & "'"
        Sql = Sql & vbCrLf & " ORDER BY V_REARCHIVO_DIGITAL.COD_REMITO, "
        Sql = Sql & vbCrLf & " V_REARCHIVO_DIGITAL.ID_CODIGO_DOCUMENTO"
        frmReportes.ImprimirReporte PasoReportes & "rptRearchivoDigital.rpt", Sql, True
        MousePointer = 0
End Sub

Private Sub cmdImprimirOrden_Click()
    Dim Orden As Long
        Orden = InputBox("Ingrese el Numero de Orden", "Orden", 0)
        ImprimirOrden Orden
        
End Sub

Private Sub cmdOsepActualizar_Click()
        Dim Sql As String
        Dim COD_CLIENTE, Cod_Indice As String
        Dim Cod_Nro_Caja, Elemento, Elemento_String As String
        Dim Cod_Tipo_Orden, Descripcion, Orden As String
        Dim Contenedor_Prov As String
        Dim Max_Cod_Documentacion As Long
        Dim Elemento_Numero As Long
        Dim i As Integer
        Dim rs As New ADODB.Recordset
            If Not IsNumeric(txtMax_COD_DOCUMENTACION.Text) Then
                MsgBox "La orden es incorrecta", vbCritical
                Exit Sub
            End If
            Sql = " SELECT MAX(ORDEN) AS maxOrden "
            Sql = Sql & "  From ORDENAR_DOCUMENTACION_DETALLE "
            Sql = Sql & " Where COD_DOCUMENTACION = " & txtMax_COD_DOCUMENTACION.Text
            rs.Open Sql, ConActiva, 0, 1
            If Not rs.EOF Then
                Orden = rs!MaxOrden
            Else
                MsgBox "La orden No existe", vbCritical
                Exit Sub
            End If
            If Trim(txtubicaprovosep.Text) = "" Then
                MsgBox "Ingrese el ubicación provisoria"
                Exit Sub
            End If
            Max_Cod_Documentacion = txtMax_COD_DOCUMENTACION.Text
            With grdAfiliados
                For i = 1 To .Rows - 1
                    COD_CLIENTE = .TextMatrix(i, 1)
                    Cod_Nro_Caja = .TextMatrix(i, 2)
                    Elemento = "'" & Trim(.TextMatrix(i, 3)) & "'"
                    If IsNumeric(.TextMatrix(i, 3)) Then
                        Elemento_Numero = CLng(.TextMatrix(i, 3))
                        Elemento_String = "Null"
                    Else
                        Elemento_String = "'" & (.TextMatrix(i, 3)) & "'"
                        Elemento_Numero = Null
                    End If
                    Cod_Tipo_Orden = 1
                    Descripcion = .TextMatrix(i, 4)
                    Orden = CInt(Orden) + CInt(.TextMatrix(i, 0))
                    Contenedor_Prov = "'" & .TextMatrix(i, 5) & "'"
                    Descripcion = Replace(Descripcion, "'", "´")
                    Sql = " INSERT INTO ORDENAR_DOCUMENTACION_DETALLE"
                    Sql = Sql & vbCrLf & " (COD_DOCUMENTACION, COD_CLIENTE, COD_INDICE,"
                    Sql = Sql & vbCrLf & " COD_NRO_CAJA, ELEMENTO, COD_TIPO_ORDEN,DESCRIPCION,CONTENEDOR_PROV,ORDEN ,COD_ESTADO,ELEMENTO_STRING, ELEMENTO_NUMERO )"
                    Sql = Sql & vbCrLf & "  VALUES ( "
                    Sql = Sql & vbCrLf & Max_Cod_Documentacion & "," & COD_CLIENTE & ",'002008002',"
                    Sql = Sql & vbCrLf & txtCajaAfil.Text & "," & Elemento & "," & Cod_Tipo_Orden & ",'" & Descripcion & "'," & Contenedor_Prov & "," & Orden & " ,0," & Elemento_String & "," & Elemento_Numero & ")"
                    ExecutarSql Sql
                Next
            End With
            InsertarProducion ctlPersonalafiadosOsep.Valor, 9, "CARGA ORDEN:" & Max_Cod_Documentacion, grdAfiliados.Rows - 1, 20
            MsgBox "Orden Numero : " & Max_Cod_Documentacion
            MousePointer = 11
            ImprimirOrdenDocumentacion Max_Cod_Documentacion
            TitulosGrillaAfiliados
            mskDocumento.Mask = ""
            txtMax_COD_DOCUMENTACION.Text = ""
            LimpiarMask mskDocumento
            LimpiarMask mskRemitoOsep
            txtubicaprovosep.Text = ""
            MousePointer = 0
End Sub

Private Sub ImprimirOrdenDocumentacion(Orden As Long)
    Dim Sql As String
        Sql = " SELECT * "
        Sql = Sql & " FROM V_ORDEN_DOCUMENTACION "
        Sql = Sql & " Where ID_ORDENAR_DOCUMENTACION =" & Orden
        Sql = Sql & " ORDER BY ID_ORDENAR_DOCUMENTACION"
        frmReportes.ImprimirReporte PasoReportes & "rptOrdenarDocumentacion.rpt", Sql, True
End Sub

Private Sub cmdOsepImprimir_Click()
    On Error GoTo salir
    ImprimirOrdenDocumentacion InputBox("Ingrese el Nº de Orden")
salir:
End Sub

Private Sub cmdPendientes_Click()
    Dim Sql As String
        Sql = " SELECT *"
        Sql = Sql & "  From V_ORDEN_DOCUMENTACION_PEN"
        Sql = Sql & "  Where Cod_Estado < 6"
        Sql = Sql & " ORDER BY ID_ORDENAR_DOCUMENTACION"
        frmReportes.ImprimirReporte PasoReportes & "rptOrdenarDocumentacionPendientes.RPT", Sql, True

End Sub

Private Sub cmdRetiroDocumentacion_Click()
Dim Sql As String
    If mskRemito.Text = "0001-0000____" Then
        MsgBox "Ingrese el Numero de remito"
        Exit Sub
    End If
    MousePointer = 11
    Sql = "  SELECT *"
    Sql = Sql & vbCrLf & " From V_RETIRO_OSEP"
    Sql = Sql & vbCrLf & " Where COD_REMITO_PRO = '" & mskRemito.Text & "'"
    Sql = Sql & vbCrLf & " ORDER BY COD_INDICE ASC, ID_ORDENAR_DOCUMENTACION ASC ,ELEMENTO_NUMERO Asc"
    frmReportes.ImprimirReporte PasoReportes & "rptREtiroDocumentacion.rpt", Sql, True
    MousePointer = 0
End Sub


Private Sub Command1_Click()
 TitulosGrillaOrdenDocumentacion
End Sub



Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command2_Click()
 Dim fecha As String
    Dim Sql As String
        
       If Not IsNull(ctlPersonalDOcumento.Valor) Then
            fecha = InputBox("Ingrese la Fecha", "Control de Carga", Format(Now, "dd/mm/yyyy"))
            MousePointer = 11
                Sql = " SELECT *"
                Sql = Sql & " FROM V_ORDEN_DOCUMENTACION "
                Sql = Sql & " where COD_RESPONSABLE_CARGA = " & ctlPersonalDOcumento.Valor
                Sql = Sql & " and FECHA > " & FechaServerTipo(fecha)
                Sql = Sql & " AND FECHA < " & FechaServerTipo(DateAdd("d", 1, fecha))
                Sql = Sql & " ORDER BY FECHADDMMYYYY,ID_ORDENAR_DOCUMENTACION, FECHA"
            frmReportes.ImprimirReporte PasoReportes & "rptOrdenarDocumentacionCantidadCaracteres.RPT", Sql, True
            MousePointer = 0
        Else
            MsgBox "Ingrese el responsable", vbInformation
        End If
End Sub

'Private Sub Command5_Click()
'     Dim Sql As String
'    Dim COD_CLIENTE, COD_INDICE As String
'    Dim COD_NRO_CAJA, ELEMENTO, ELEMENTO_STRING, COD_TIPO_ORDEN, DESCRIPCION, ORDEN, CONTENEDOR_PROV As String
'    Dim Max_COD_DOCUMENTACION As Long
'    Dim ELEMENTO_NUMERO As Long
'    Dim i As Integer
'    Dim Rs As New ADODB.Recordset
'    If Trim(txtubicaprovosep.Text) = "" Then
'        MsgBox "Ingrese el ubicación provisoria"
'        Exit Sub
'    End If
'    Max_COD_DOCUMENTACION = txtMax_COD_DOCUMENTACION.Text
'    txtMax_COD_DOCUMENTACION.Text = ""
'With grdAfiliados
'    For i = 1 To .Rows - 1
'       COD_CLIENTE = .TextMatrix(i, 1)
'       COD_NRO_CAJA = .TextMatrix(i, 2)
'        ELEMENTO = "'" & Trim(.TextMatrix(i, 3)) & "'"
'        If IsNumeric(.TextMatrix(i, 3)) Then
'            ELEMENTO_NUMERO = CLng(.TextMatrix(i, 3))
'            ELEMENTO_STRING = "Null"
'        Else
'            ELEMENTO_STRING = "'" & (.TextMatrix(i, 3)) & "'"
'            ELEMENTO_NUMERO = Null
'        End If
'        COD_TIPO_ORDEN = 1
'        DESCRIPCION = .TextMatrix(i, 4)
'        ORDEN = .TextMatrix(i, 0)
'        CONTENEDOR_PROV = "'" & .TextMatrix(i, 5) & "'"
'        DESCRIPCION = Replace(DESCRIPCION, "'", "´")
'        Sql = " INSERT INTO ORDENAR_DOCUMENTACION_DETALLE"
'        Sql = Sql & vbCrLf & " (COD_DOCUMENTACION, COD_CLIENTE, COD_INDICE,"
'        Sql = Sql & vbCrLf & " COD_NRO_CAJA, ELEMENTO, COD_TIPO_ORDEN,DESCRIPCION,CONTENEDOR_PROV,ORDEN ,COD_ESTADO,ELEMENTO_STRING, ELEMENTO_NUMERO )"
'        Sql = Sql & vbCrLf & "  VALUES ( "
'        Sql = Sql & vbCrLf & Max_COD_DOCUMENTACION & "," & COD_CLIENTE & ",'002008002',"
'        Sql = Sql & vbCrLf & 0 & "," & ELEMENTO & "," & COD_TIPO_ORDEN & ",'" & DESCRIPCION & "'," & CONTENEDOR_PROV & "," & ORDEN & " ,0," & ELEMENTO_STRING & "," & ELEMENTO_NUMERO & ")"
'        ExecutarSql Sql
'    Next
'End With
'InsertarProducion ctlPersonalafiadosOsep.Valor, 9, "CARGA ORDEN:" & Max_COD_DOCUMENTACION, grdAfiliados.Rows - 1, 20
'MsgBox "Orden Numero : " & Max_COD_DOCUMENTACION
'MousePointer = 11
'Sql = " SELECT * "
'Sql = Sql & vbCrLf & " From ORDENAR_DOCUMENTACION, ORDENAR_DOCUMENTACION_DETALLE"
'Sql = Sql & vbCrLf & " Where ORDENAR_DOCUMENTACION.ID_ORDENAR_DOCUMENTACION = ORDENAR_DOCUMENTACION_DETALLE.COD_DOCUMENTACION"
'Sql = Sql & vbCrLf & " AND ORDENAR_DOCUMENTACION.ID_ORDENAR_DOCUMENTACION =" & Max_COD_DOCUMENTACION
'Sql = Sql & vbCrLf & " Order By ORDENAR_DOCUMENTACION_DETALLE.ORDEN ASC"
'    CrystalReport1.DiscardSavedData = True
'    CrystalReport1.Connect = "DSN = bpdc;UID = basa ;PWD = 1742"
'    rptORDENARDocumentacioN
'    CrystalReport1.SQLQuery = Sql
'    CrystalReport1.Destination = 0
'    CrystalReport1.Action = 1
'
'TitulosGrillaAfiliados
'mskDocumento.Mask = ""
'LimpiarMask mskDocumento
'LimpiarMask mskRemitoOsep
'txtubicaprovosep.Text = ""
'MousePointer = 0
'End Sub

Private Sub Command6_Click()
On Error GoTo salir
Dim NUMERO As Integer
  NUMERO = InputBox("Ingrese el Numero A Borrar")
  If grdAfiliados.Rows = 2 Then
    grdAfiliados.Clear
    grdAfiliados.Rows = 1
  End If
  
        grdAfiliados.RemoveItem (NUMERO)
        Exit Sub
salir:
MsgBox "Error"
        
End Sub

Private Sub Command7_Click()

End Sub


Private Sub Command9_Click()

End Sub

Private Sub ctlPersonalafiadosOsep_Click()
    Set ConOsep = New ADODB.Connection
    ConOsep.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ClienteOsep & ";Persist Security Info=False"

End Sub

Private Sub Form_Load()
    ctlClientesBuscarDocumento.TipoControl = Cliente
    ctlClientesDocumento.TipoControl = Cliente
    ctlPersonalafiadosOsep.TipoControl = Personal
    ctlPersonalControl.TipoControl = Personal
    ctlPersonalDOcumento.TipoControl = Personal
    ctlPersonalOrden.TipoControl = Personal
    ctlPersonalOrdenTerminado.TipoControl = Personal
    TitulosGrillaOrdenDocumentacion
    TitulosGrillaAfiliados
    tabOrdenDocumentacion.TabCaption(0) = "Orden de documentación"
    
 End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
End Sub

Private Sub MaskEdBox2_Change()

End Sub

Private Sub grdAfiliados_Scroll()
If 1 = 1 Then
End If
End Sub

Private Sub grdOrdenDocumentacion_Scroll()
Rem MsgBox grdOrdenDocumentacion.TopRow
End Sub

Private Sub mskBuscar_GotFocus()
        mskBuscar.SelStart = 0
End Sub

Private Sub mskBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    VerificarDato (20)
End If

End Sub


Private Sub mskDocumento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not IsNumeric(mskDocumento.Text) Then
           MsgBox "NO ES UN NUMERO"
           Exit Sub
        End If
        If mskRemitoOsep.Text = "0001-0000____" Then
           MsgBox "INGRESAR REMITO"
           Exit Sub
        End If
        If Trim(txtubicaprovosep.Text) = "" Then
           MsgBox "INGRESAR UBICACION PROVISORIA"
           Exit Sub
        End If
        
        If Trim(txtCajaAfil.Text) = "" Then
           MsgBox "INGRESAR LA CAJA"
           Exit Sub
        End If
        
        cmdAceptarLegajos_Click
    
    End If

End Sub

Private Sub mskRemito_GotFocus()
    mskRemito.SelStart = 8
End Sub

Private Sub mskRemito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If

End Sub

Private Sub TabStrip2_Click()
fraAfiliadoOsep.Visible = True
End Sub

Private Sub txtBuscar1_Change()

End Sub

Private Sub txtCaja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            If IsNumeric(txtCodigo.Text) Then
                SendKeys vbTab
            Else
                MsgBox "NO es un Nº Valido"
            End If
        End If
End Sub

Private Sub txtCodigo_LostFocus()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String
    If Trim(txtCodigo.Text) = "" Then
    Exit Sub
    End If
    If Mid(txtCodigo.Text, 1, 1) = 0 Then
       Sql = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE,MASK_EXPEDIENTE , DESCRIPCION "
       Sql = Sql & " From INDICES WHERE (COD_CLIENTE =" & ctlClientesDocumento.Valor & ") AND (INDICE = '" & txtCodigo.Text & "')"
    Else
        
        If Not IsNull(ctlClientesDocumento.Valor) Then
            Sql = "SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE,MASK_EXPEDIENTE , DESCRIPCION "
            Sql = Sql & " From INDICES WHERE (COD_CLIENTE = " & ctlClientesDocumento.Valor & ") AND (ID_CODIGO_DOCUMENTO =" & txtCodigo.Text & ")"
        Else
            MsgBox "INGRESE EL CLIENTE"
            Exit Sub
        End If
    End If
    rs.Open Sql, ConActiva, 0, 1
    If Not rs.EOF Then
        If Not IsNull(rs!MASK_EXPEDIENTE) Then
          Rem   mskBuscar.Mask = RS!MASK_EXPEDIENTE
            lblDescripcionCodigo.Caption = rs!Descripcion
            lblCodigo.Caption = rs!Indice
        Else
            mskBuscar.Mask = ""
           lblDescripcionCodigo.Caption = rs!Descripcion
           lblCodigo.Caption = rs!Indice
        End If
    Else
         mskBuscar.Mask = "NO CARGAR"
         lblDescripcionCodigo.Caption = ""
         lblCodigo.Caption = ""
    End If
    txtCodigo.Text = ""
    
End Sub

Private Sub txtElementoBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not IsNull(ctlClientesBuscarDocumento.Valor) Then
    
    If IsNumeric(txtElementoBuscar) Then
        
       Dim rs As ADODB.Recordset
        Dim Sql As String
        
      Sql = " SELECT INDICES.DESCRIPCION as DESCRIPCION,ELEMENTO_NUMERO,"
    Sql = Sql & vbCrLf & " Cod_Estado , COD_DOCUMENTACION, CONTENEDOR_PROV, COD_NRO_CAJA,ORDEN"
    Sql = Sql & vbCrLf & " From ORDENAR_DOCUMENTACION_DETALLE, INDICES"
    Sql = Sql & vbCrLf & " Where ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = INDICES.COD_CLIENTE"
    Sql = Sql & vbCrLf & " AND ORDENAR_DOCUMENTACION_DETALLE.COD_INDICE = INDICES.INDICE"
    Sql = Sql & vbCrLf & " AND (ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO_NUMERO =" & txtElementoBuscar.Text & ")"
  Sql = Sql & vbCrLf & " AND (ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE =" & ctlClientesBuscarDocumento.Valor & ")"
    Sql = Sql & vbCrLf & " ORDER BY INDICES.DESCRIPCION, ORDENAR_DOCUMENTACION_DETALLE.COD_DOCUMENTACION"
    
    
    Sql = " SELECT     ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO_NUMERO, ORDENAR_DOCUMENTACION_DETALLE.COD_ESTADO,"
    Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.COD_DOCUMENTACION, ORDENAR_DOCUMENTACION_DETALLE.CONTENEDOR_PROV,"
    Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Nro_Caja , ORDENAR_DOCUMENTACION_DETALLE.Orden, ORDENAR_DOCUMENTACION.ANULADO"
    Sql = Sql & vbCrLf & " FROM ORDENAR_DOCUMENTACION_DETALLE INNER JOIN"
    Sql = Sql & vbCrLf & " INDICES ON ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = INDICES.COD_CLIENTE AND"
    Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.COD_INDICE = INDICES.INDICE INNER JOIN"
    Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION ON "
    Sql = Sql & vbCrLf & " ORDENAR_DOCUMENTACION_DETALLE.COD_DOCUMENTACION = ORDENAR_DOCUMENTACION.ID_ORDENAR_DOCUMENTACION"
 Sql = Sql & vbCrLf & " Where   (ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO_NUMERO =" & txtElementoBuscar.Text & ")"
  Sql = Sql & vbCrLf & " AND (ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE =" & ctlClientesBuscarDocumento.Valor & ")"
Sql = Sql & vbCrLf & " ORDER BY INDICES.DESCRIPCION, ORDENAR_DOCUMENTACION_DETALLE.COD_DOCUMENTACION"
    
     
    
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Sql, ConActiva, 0, 1
    
    Set grdBuscarOrdenamiento.DataSource = rs.DataSource
    grdBuscarOrdenamiento.DataMember = rs.DataMember
    grdBuscarOrdenamiento.Rebind
    grdBuscarOrdenamiento.Refresh
    Else
        MsgBox "no es Un Numero", vbCritical
    End If
    
End If
End If
End Sub

Private Sub txtLecturaLegajos_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    If UCase(Mid(txtLecturaLegajos.Text, 1, 2)) = "L1" Then
       grdOrdenLegajos.AddItem "" & vbTab & CLng(Mid(txtLecturaLegajos, 3, 3)) & vbTab & CLng(Mid(txtLecturaLegajos, 6, 6))
       txtLecturaLegajos.Text = ""
    End If
  End If
End Sub

Private Sub txtSql_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Public Function VerificarDato(COD_CLIENTE As Integer) As Boolean
    Dim Sql As String
    Set rsBuscar = New ADODB.Recordset
    If Trim(txtCaja.Text) = "" Then
        MsgBox "Por favor ingrese la ubicacion prov"
        Exit Function
    End If
    If Not IsNumeric(mskBuscar.Text) Then
        MsgBox "Verifique el dato"
      End If
        Select Case COD_CLIENTE
        Case 65
              grdOrdenDocumentacion.AddItem grdOrdenDocumentacion.Rows & vbTab & ctlClientesDocumento.Valor & vbTab & txtCaja.Text & vbTab & mskBuscar.ClipText & vbTab & mskBuscar.ClipText & vbTab & txtCaja.Text & vbTab & Now
              LimpiarMask mskBuscar
              If grdOrdenDocumentacion.Rows = 25 Or grdOrdenDocumentacion.Rows = 50 Or grdOrdenDocumentacion.Rows = 75 Or grdOrdenDocumentacion.Rows = 100 Or grdOrdenDocumentacion.Rows = 125 Or grdOrdenDocumentacion.Rows = 150 Then
                    lblIndicador.BackColor = &HFF00&
                    Beep
                    Beep
              Else
                    lblIndicador.BackColor = &H8000000F
              End If
              grdOrdenDocumentacion.TopRow = grdOrdenDocumentacion.Rows - 1
              grdOrdenDocumentacion_Scroll
              LimpiarMask mskBuscar
        
'            SQL = " SELECT SOCNRO, SOCNOM, TIPDOC, SOCNUMDOC, ESTACOD,DATETRANSACTION "
'            SQL = SQL & " From AsistirSocios Where  SOCNUMDOC =" & mskBuscar.ClipText
'
'
'           rsBuscar.Open SQL, strConBasa , 0 ,1
'           If Not rsBuscar.EOF Then
'                grdOrdenDocumentacion.AddItem grdOrdenDocumentacion.Rows & vbTab & ctlCliente.Valor & vbTab & rsBuscar!SOCNRO & vbTab & mskBuscar.ClipText & vbTab & Format(rsBuscar!SOCNRO, "000000") & " " & rsBuscar!SOCNOM & vbTab & TXTCAJA.Text
'                If grdOrdenDocumentacion.Rows = 25 Or grdOrdenDocumentacion.Rows = 50 Or grdOrdenDocumentacion.Rows = 75 Or grdOrdenDocumentacion.Rows = 100 Or grdOrdenDocumentacion.Rows = 125 Or grdOrdenDocumentacion.Rows = 150 Then
'                    lblIndicador.BackColor = &HFF00&
'                Else
'                    lblIndicador.BackColor = &H8000000F
'                End If
'                grdOrdenDocumentacion.TopRow = grdOrdenDocumentacion.Rows - 1
'                grdOrdenDocumentacion_Scroll
'                LimpiarMask mskBuscar
'           Else
'              MsgBox "Atencion NO fue encontrado en la referencia", vbCritical
'              LimpiarMask mskBuscar
'           End If
'
        Case Else
                If IsNumeric(mskBuscar.Text) Then
                    Sql = " SELECT NRO_CAJA, NRO_DESDE, NRO_HASTA, ITEM, INDICE,LETRA_DESDE ,LETRA_HASTA "
                    Sql = Sql & vbCrLf & " From REFERENCIAS WHERE (COD_CLIENTE = " & ctlClientesDocumento.Valor & ") AND (INDICE = '" & lblCodigo.Caption & "')"
                    Sql = Sql & vbCrLf & " AND (" & mskBuscar.ClipText & " BETWEEN   NRO_DESDE AND NRO_HASTA)"
                   If Trim(txtSql.Text) <> "" Then
                        Sql = Sql & txtSql.Text
                   End If
                   rsBuscar.Open Sql, ConActiva, 0, 1
                   If Not rsBuscar.EOF Then
                        grdOrdenDocumentacion.AddItem grdOrdenDocumentacion.Rows & vbTab & ctlClientesDocumento.Valor & vbTab & rsBuscar!NRO_CAJA & vbTab & mskBuscar.ClipText & vbTab & " " & vbTab & txtCaja.Text & vbTab & Now
                        If grdOrdenDocumentacion.Rows = 25 Or grdOrdenDocumentacion.Rows = 50 Or grdOrdenDocumentacion.Rows = 75 Or grdOrdenDocumentacion.Rows = 100 Or grdOrdenDocumentacion.Rows = 125 Or grdOrdenDocumentacion.Rows = 150 Then
                            lblIndicador.BackColor = &HFF00&
                        Else
                            lblIndicador.BackColor = &H8000000F
                        End If
                        grdOrdenDocumentacion.TopRow = grdOrdenDocumentacion.Rows - 1
                        grdOrdenDocumentacion_Scroll
                        LimpiarMask mskBuscar
                   Else
                      MsgBox "Atencion NO fue encontrado en la referencia", vbCritical
                      Rem grdOrdenDocumentacion.AddItem grdOrdenDocumentacion.Rows & vbTab & 0 & vbTab & 0 & vbTab & mskBuscar.ClipText & vbTab & 0 & vbTab & TXTCAJA.Text
                      LimpiarMask mskBuscar
                   End If
                Else
                    MsgBox "No es un numero", vbCritical
                    LimpiarMask mskBuscar
                End If
        
        End Select
        


End Function

Public Sub ImprimirOrden(Orden As Long)
    Dim Sql As String
        MousePointer = 11
        Sql = Sql & vbCrLf & " SELECT * "
        Sql = Sql & vbCrLf & " From V_ORDEN_DOCUMENTACION"
        Sql = Sql & vbCrLf & " Where ID_ORDENAR_DOCUMENTACION = " & Orden
        Sql = Sql & vbCrLf & " Order By ORDEN ASC"
        frmReportes.ImprimirReporte PasoReportes & "rptOrdenarDocumentacion.rpt", Sql, True
        MousePointer = 0
End Sub


