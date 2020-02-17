VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAgregarDocumentos 
   AutoRedraw      =   -1  'True
   Caption         =   "Agregar Documentos"
   ClientHeight    =   11010
   ClientLeft      =   645
   ClientTop       =   20445
   ClientWidth     =   14550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   14550
   WindowState     =   2  'Maximized
   Begin VB.Frame fraLegajo 
      Caption         =   "LEGAJOS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   25
      Top             =   600
      Width           =   11715
      Begin VB.CommandButton cmdsolicituddecaja 
         Caption         =   "Solicitar de Caja"
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
         Left            =   7680
         TabIndex        =   90
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox txtLectura 
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
         Left            =   6960
         TabIndex        =   89
         Top             =   300
         Width           =   495
      End
      Begin VB.CommandButton cmdModificar 
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
         Left            =   2640
         Picture         =   "frmReferencias.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton cmdBorrarEtiqueta 
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
         Left            =   3000
         Picture         =   "frmReferencias.frx":0256
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdRegistroVerificado 
         Caption         =   "Nuevo"
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
         Left            =   3420
         Picture         =   "frmReferencias.frx":04BB
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdCajasConLegajos 
         Caption         =   "Caja Con Legajos"
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
         TabIndex        =   61
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox txtEtiquetaDigitoVerificador 
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
         Height          =   330
         Left            =   2220
         TabIndex        =   58
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdCargaCompletaCajaLegajo 
         Caption         =   "Carga Completa"
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
         Left            =   9480
         TabIndex        =   54
         Top             =   300
         Width           =   1575
      End
      Begin VB.CheckBox chkAutoIncrementar 
         Caption         =   "Auto Incrementar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4920
         TabIndex        =   28
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtEtiqueta 
         BackColor       =   &H00C0E0FF&
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
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Etiqueta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   255
         Left            =   180
         TabIndex        =   26
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame fraCampos 
      Height          =   3195
      Left            =   180
      TabIndex        =   0
      Top             =   2100
      Width           =   11715
      Begin VB.Frame fraDescripcion 
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
         Left            =   120
         TabIndex        =   20
         Top             =   2400
         Width           =   10935
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
            Left            =   9540
            TabIndex        =   10
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkNoBuscar 
            Caption         =   "F11 No buscar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   80
            Top             =   180
            Value           =   1  'Checked
            Width           =   1635
         End
         Begin VB.CheckBox chkFijarDescripcion 
            Caption         =   "F9"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   180
            Width           =   615
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
            Left            =   3900
            TabIndex        =   9
            ToolTipText     =   "Buscar (*) Todos //  (-)  Tiene en cuenta tipo doc F11 No Busca"
            Top             =   180
            Width           =   5535
         End
         Begin VB.Label lblTituloDescripcion 
            Caption         =   "Descripcion:"
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
            TabIndex        =   22
            Top             =   180
            Width           =   1215
         End
      End
      Begin VB.Frame fraLetraDesde 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   120
         TabIndex        =   17
         Top             =   1380
         Width           =   10935
         Begin VB.CheckBox chk_Copiar_Letra 
            Caption         =   "Copiar"
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
            Left            =   9120
            TabIndex        =   52
            Top             =   240
            Width           =   975
         End
         Begin VB.CheckBox chkFijarLetraHasta 
            Caption         =   "F8"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtLetraHasta 
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
            Left            =   2640
            TabIndex        =   8
            Top             =   600
            Width           =   6375
         End
         Begin VB.CheckBox chkFijarLetraDesde 
            Caption         =   "F7"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   300
            Width           =   495
         End
         Begin VB.TextBox txtLetraDesde 
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
            Left            =   2640
            TabIndex        =   7
            ToolTipText     =   "Para copiar +"
            Top             =   180
            Width           =   6375
         End
         Begin VB.Label lblTituloLetraHasta 
            Caption         =   "Letra Hasta:"
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
            TabIndex        =   51
            Top             =   660
            Width           =   1515
         End
         Begin VB.Label lblTituloLetraDesde 
            Caption         =   "Letra Desde:"
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
            TabIndex        =   19
            Top             =   300
            Width           =   1575
         End
      End
      Begin VB.Frame fraNumero 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   10935
         Begin VB.CheckBox chk_Copiar_Nro 
            Caption         =   "Copiar"
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
            Left            =   9060
            TabIndex        =   42
            Top             =   180
            Width           =   975
         End
         Begin VB.TextBox txtNroHasta 
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
            Left            =   7200
            TabIndex        =   6
            Top             =   180
            Width           =   1755
         End
         Begin VB.CheckBox chkFijarNumeroHasta 
            Caption         =   "F6"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5100
            TabIndex        =   36
            Top             =   240
            Width           =   555
         End
         Begin VB.CheckBox chkFijarNumeroDesde 
            Caption         =   "F5"
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
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtNroDesde 
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
            Left            =   2640
            TabIndex        =   5
            ToolTipText     =   "Para copiar +"
            Top             =   180
            Width           =   1695
         End
         Begin VB.Label lblTituloNumeroHasta 
            Caption         =   "Hasta Numero"
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
            Left            =   5700
            TabIndex        =   37
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label lblTituloNumeroDesde 
            Caption         =   "Desde Numero:"
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
            TabIndex        =   16
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fra_Fecha_Desde 
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   10935
         Begin VB.CheckBox chk_Copiar_Fecha 
            Caption         =   "Copiar"
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
            TabIndex        =   41
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox txtFechaHasta 
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
            Left            =   7200
            TabIndex        =   4
            ToolTipText     =   "Para mes 4 año 2 dia +"
            Top             =   180
            Width           =   1635
         End
         Begin VB.CheckBox chkFijarFechaHasta 
            Caption         =   "F4"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5100
            TabIndex        =   34
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox txtFechaDesde 
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
            Left            =   2640
            TabIndex        =   3
            ToolTipText     =   "Para mes 4 año 2 dia +"
            Top             =   180
            Width           =   1695
         End
         Begin VB.CheckBox chkFijarFechaDesde 
            Caption         =   "F3"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblTituloFechaHasta 
            Caption         =   "Fecha desde :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5700
            TabIndex        =   35
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lblTituloFechaDesde 
            Caption         =   "Fecha desde :"
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
            TabIndex        =   13
            Top             =   240
            Width           =   1515
         End
      End
   End
   Begin VB.ComboBox cboTipoCarga 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "frmReferencias.frx":070E
      Left            =   9000
      List            =   "frmReferencias.frx":071B
      Style           =   2  'Dropdown List
      TabIndex        =   82
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command1"
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
      Left            =   9300
      TabIndex        =   66
      Top             =   960
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   62
      Top             =   10710
      Width           =   14550
      _ExtentX        =   25665
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Enabled         =   0   'False
            Key             =   "EstadoAplicacion"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "ID"
         EndProperty
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
   End
   Begin VB.CommandButton Command7 
      Caption         =   "84"
      Height          =   375
      Left            =   2160
      TabIndex        =   57
      Top             =   11160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "UJAM"
      Height          =   375
      Left            =   360
      TabIndex        =   56
      Top             =   11160
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton Command5 
      Caption         =   "INDICES"
      Height          =   375
      Left            =   3420
      TabIndex        =   55
      Top             =   11160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtUsuarioCarga 
      BackColor       =   &H00C0FFC0&
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
      Left            =   7380
      TabIndex        =   44
      Top             =   120
      Width           =   615
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   1380
      TabIndex        =   43
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   180
      TabIndex        =   68
      Top             =   5340
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   9869
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   4
      Tab             =   3
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Datos Ingr."
      TabPicture(0)   =   "frmReferencias.frx":073B
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grdDatos"
      Tab(0).Control(1)=   "fraBotones"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Imágenes"
      TabPicture(1)   =   "frmReferencias.frx":0757
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CommonDialog1"
      Tab(1).Control(1)=   "cmdCargarImagen"
      Tab(1).Control(2)=   "ctlVerImagenes1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Descripción"
      TabPicture(2)   =   "frmReferencias.frx":0773
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CMGGR"
      Tab(2).Control(1)=   "grdDescripcion"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Video"
      TabPicture(3)   =   "frmReferencias.frx":078F
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Video"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "cmdCargarVideo"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "txtPasoVideo"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cmdPlay_Pausa"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Timer1"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtVideoLugar"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      Begin VB.TextBox txtVideoLugar 
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
         Left            =   2100
         TabIndex        =   95
         Top             =   900
         Width           =   1035
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   780
         Top             =   1800
      End
      Begin VB.CommandButton cmdPlay_Pausa 
         Caption         =   "Espera"
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
         Left            =   3180
         TabIndex        =   94
         Top             =   900
         Width           =   915
      End
      Begin VB.TextBox txtPasoVideo 
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
         Left            =   720
         TabIndex        =   93
         Text            =   "C:\Video\"
         Top             =   420
         Width           =   3375
      End
      Begin VB.CommandButton cmdCargarVideo 
         Caption         =   "Cargar Video"
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
         Left            =   720
         TabIndex        =   92
         Top             =   900
         Width           =   1275
      End
      Begin Controles.ctlVerImagenes ctlVerImagenes1 
         Height          =   4695
         Left            =   -74220
         TabIndex        =   81
         Top             =   240
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   8281
      End
      Begin VB.CommandButton cmdCargarImagen 
         Caption         =   "Imagen"
         Height          =   315
         Left            =   -67140
         TabIndex        =   78
         Top             =   240
         Width           =   1035
      End
      Begin VB.Frame fraBotones 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74220
         TabIndex        =   69
         Top             =   4140
         Width           =   10815
         Begin VB.CommandButton cmddescripcion 
            Caption         =   "Descripción"
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
            Left            =   5100
            TabIndex        =   88
            Top             =   240
            Width           =   1215
         End
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
            Height          =   375
            Left            =   6480
            TabIndex        =   74
            Top             =   240
            Width           =   1275
         End
         Begin VB.TextBox txtID_Referencia 
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
            Left            =   780
            MultiLine       =   -1  'True
            TabIndex        =   73
            Top             =   300
            Width           =   2955
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
            Height          =   375
            Left            =   9300
            TabIndex        =   72
            Top             =   240
            Width           =   1215
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
            Height          =   375
            Left            =   7920
            TabIndex        =   71
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdBorrarFiltro 
            Caption         =   "X"
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
            Left            =   3780
            TabIndex        =   70
            Top             =   300
            Width           =   315
         End
         Begin VB.Label Label3 
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
            Height          =   255
            Left            =   180
            TabIndex        =   75
            Top             =   360
            Width           =   495
         End
      End
      Begin MSDataGridLib.DataGrid grdDescripcion 
         Height          =   4095
         Left            =   -74220
         TabIndex        =   77
         Top             =   180
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7223
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
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
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -65760
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid grdDatos 
         Height          =   3615
         Left            =   -74220
         TabIndex        =   79
         Top             =   240
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   6376
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
      Begin VB.CommandButton CMGGR 
         Caption         =   "Command1"
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
         Left            =   -66360
         TabIndex        =   76
         Top             =   4500
         Visible         =   0   'False
         Width           =   1755
      End
      Begin WMPLibCtl.WindowsMediaPlayer Video 
         Height          =   5115
         Left            =   4260
         TabIndex        =   91
         Top             =   300
         Width           =   7035
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "mini"
         stretchToFit    =   -1  'True
         windowlessVideo =   0   'False
         enabled         =   0   'False
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   12409
         _cy             =   9022
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   57
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":07AB
            Key             =   "Ver+"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":0BA5
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":0E45
            Key             =   "Ver-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":1243
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":1657
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":1A1F
            Key             =   "Punto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":1B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":1F57
            Key             =   "RotarI"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":2368
            Key             =   "Vertical"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":2787
            Key             =   "Sig"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":2B88
            Key             =   "Ant"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":2F86
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":339C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":3452
            Key             =   "RotarD"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":385C
            Key             =   "Cargar"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":3C37
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":3FFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":40F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":44E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":48F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":4CB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":4D51
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":4DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":51E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":55D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":59AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":5DC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":5F6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":634B
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":6458
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":685F
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":6C6F
            Key             =   "Fin"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":709B
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":71D2
            Key             =   "Aceptar"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":758B
            Key             =   "Control"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":7676
            Key             =   "Esp. Fax"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":7AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":7BE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":7FF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":8169
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":8570
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":89A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":8D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":918C
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":958D
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":995B
            Key             =   "Anular"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":9D29
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":A132
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":A4FE
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":A932
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":AD77
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":B00F
            Key             =   "Casa"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":B416
            Key             =   "Bandera"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":B7F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":BC51
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":C8A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmReferencias.frx":D17D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCajaDocumento 
      Caption         =   "CAJA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   45
      Top             =   1320
      Width           =   11715
      Begin VB.CommandButton cmdBuscarCaja 
         Caption         =   "..."
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
         Left            =   2940
         TabIndex        =   60
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtCajaDigitoVerificador 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2460
         TabIndex        =   59
         Top             =   300
         Width           =   375
      End
      Begin VB.TextBox txtCaja 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1200
         TabIndex        =   1
         Top             =   300
         Width           =   1215
      End
      Begin VB.CheckBox chkFijarCaja 
         Caption         =   "F1 Caja"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   180
         TabIndex        =   49
         Top             =   300
         Width           =   915
      End
      Begin VB.TextBox txtIndice_Nro_Documento 
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
         Left            =   5160
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
      Begin VB.CommandButton cmdBuscarDocumento 
         Caption         =   "F12"
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
         Left            =   6000
         TabIndex        =   47
         Top             =   300
         Width           =   435
      End
      Begin VB.CheckBox chkFijarTipoDocumento 
         Caption         =   "F2 Tipo Doc."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   3720
         TabIndex        =   46
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblIndice_Descripcion 
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
         Height          =   330
         Left            =   6600
         TabIndex        =   48
         Top             =   300
         Width           =   4455
      End
   End
   Begin VB.Frame fraRearchivo 
      Caption         =   "Rearchivo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   180
      TabIndex        =   29
      Top             =   7200
      Width           =   11355
      Begin VB.CommandButton cmdNuevoLote 
         Caption         =   "Nuevo Lote"
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
         Left            =   7500
         TabIndex        =   40
         Top             =   780
         Width           =   1695
      End
      Begin VB.TextBox txtLote 
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
         Left            =   780
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   660
         Width           =   2715
      End
      Begin VB.TextBox txtRearchivoUbicacionProvisoria 
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
         Left            =   5520
         TabIndex        =   33
         Top             =   240
         Width           =   3675
      End
      Begin VB.ComboBox cboTipoRearchivo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   780
         TabIndex        =   31
         Text            =   "Combo2"
         Top             =   300
         Width           =   2715
      End
      Begin VB.Label Label9 
         Caption         =   "Lote"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Ubicación provisoria"
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
         TabIndex        =   32
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraReferencias 
      Caption         =   "REFERENCIAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   240
      TabIndex        =   63
      Top             =   1320
      Width           =   11715
      Begin VB.CommandButton Command11 
         Caption         =   "Command1"
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
         Left            =   7200
         TabIndex        =   67
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtLoteReferencia 
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
         Left            =   2700
         TabIndex        =   65
         Text            =   "0"
         Top             =   300
         Width           =   1155
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   450
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   794
         ButtonWidth     =   714
         ButtonHeight    =   688
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Nuevo"
               ImageIndex      =   57
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Borrar"
               ImageIndex      =   56
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Imprimir"
               ImageIndex      =   54
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PorCaja"
                     Text            =   "Por Caja"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PorIndice"
                     Text            =   "Por Indice"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Diccionario"
                     Text            =   "Dicionario de documento"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "PorIndiceyCaja"
                     Text            =   "Por Indice y Caja"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         Begin VB.Label Label2 
            Caption         =   "Modificar Caja : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   5460
            TabIndex        =   84
            Top             =   0
            Width           =   1395
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Orden:"
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
         Index           =   0
         Left            =   2040
         TabIndex        =   64
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario:"
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
      Left            =   6540
      TabIndex        =   53
      Top             =   180
      Width           =   675
   End
   Begin VB.Label Label5 
      Caption         =   "Tipo Carga:"
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
      Left            =   8400
      TabIndex        =   24
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "CLIENTES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   180
      Width           =   1035
   End
   Begin VB.Menu mnuLegajos 
      Caption         =   "Legajos"
      Visible         =   0   'False
      Begin VB.Menu mnuFiltroUsuarioDia 
         Caption         =   "Fitro por Usuario y dia"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmAgregarDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim grdAnchoDatos(20)  As Long
Dim APELLIDO_NOMBRE As String
Dim Nombre_Farmacia As String
Public NºDOCUMENTO As Long
Dim SqlReferencia As String
Dim ValorAnteVideo As String

Private Sub txtTipoDocumento_Change()

End Sub

Public Function Configurar_Carga(Cliente As Integer, Optional Nro_documento As Long, Optional Indice As String) As Boolean
    Dim Sql As String
    Dim rs As ADODB.Recordset
    
On Error GoTo salir

    
    
    Sql = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, TIPO_INDICE,DESCRIPCION, TITULO_FECHA_DESDE, TITULO_FECHA_HASTA, TITULO_LETRA_DESDE,"
    Sql = Sql & " TITULO_LETRA_HASTA, TITULO_NRO_DESDE, TITULO_NRO_HASTA, TITULO_DESCRIPCION, HABILITAR_FECHA_DESDE, HABILITAR_FECHA_HASTA,"
    Sql = Sql & " HABILITAR_LETRA_DESDE, HABILITAR_LETRA_HASTA, HABILITAR_NRO_DESDE, HABILITAR_NRO_HASTA, HABILITAR_DESCRIPCION,"
    Sql = Sql & " REQUERIR_FECHA_DESDE, REQUERIR_FECHA_HASTA, REQUERIR_LETRA_DESDE, REQUERIR_LETRA_HASTA, REQUERIR_NRO_DESDE,"
    Sql = Sql & " REQUERIR_NRO_HASTA , REQUERIR_DESCRIPCION, DESCRIPCION, COPIAR_FECHA, COPIAR_LETRA, COPIAR_NRO "
    Sql = Sql & " From INDICES "
    Sql = Sql & " Where COD_CLIENTE = " & Cliente
    Sql = Sql & " And ID_CODIGO_DOCUMENTO = " & Nro_documento
    
    Set rs = New ADODB.Recordset
    
    rs.Open Sql, ConActiva, 0, 1

   
     lblIndice_Descripcion.Caption = ""
    
    If Not rs.EOF Then
        Configurar_Carga = True
     If cboTipoCarga.Text = "Legajos" Then
        If Trim(rs!Tipo_Indice) <> "Legajo" Then
          LimpiarLegajos
           Exit Function
        End If
     
    End If
    
        lblIndice_Descripcion.Caption = rs!Descripcion
    
       If IsNull(rs!TITULO_FECHA_DESDE) Then
            lblTituloFechaDesde.Caption = "Fecha Desde:"
       Else
            lblTituloFechaDesde.Caption = rs!TITULO_FECHA_DESDE
       End If
       
       txtFechaDesde.Tag = ISNULLFALSE(rs!REQUERIR_FECHA_DESDE)
       
       
       If IsNull(rs!TITULO_FECHA_HASTA) Then
            lblTituloFechaHasta.Caption = "Fecha Hasta:"
       Else
            lblTituloFechaHasta.Caption = rs!TITULO_FECHA_HASTA
       End If
       txtFechaHasta.Tag = ISNULLFALSE(rs!REQUERIR_FECHA_HASTA)
       
               
       If IsNull(rs!TITULO_LETRA_DESDE) Then
            lblTituloLetraDesde.Caption = "Letra Desde:"
       Else
           lblTituloLetraDesde.Caption = rs!TITULO_LETRA_DESDE
       End If
       txtLetraDesde.Tag = ISNULLFALSE(rs!REQUERIR_LETRA_DESDE)
        
        If IsNull(rs!TITULO_LETRA_HASTA) Then
            lblTituloLetraHasta.Caption = "Letra desde:"
        Else
           lblTituloLetraHasta.Caption = rs!TITULO_LETRA_HASTA
        End If
        txtLetraHasta.Tag = ISNULLFALSE(rs!REQUERIR_LETRA_HASTA)
         
         If IsNull(rs!TITULO_NRO_DESDE) Then
            lblTituloNumeroDesde.Caption = "Nro Desde:"
         Else
             lblTituloNumeroDesde.Caption = rs!TITULO_NRO_DESDE
         End If
         txtNroDesde.Tag = ISNULLFALSE(rs!REQUERIR_NRO_DESDE)
         
         
         
         If IsNull(rs!TITULO_NRO_HASTA) Then
            lblTituloNumeroHasta.Caption = "Nro Hasta:"
         Else
            lblTituloNumeroHasta.Caption = rs!TITULO_NRO_HASTA
         End If
         txtNroHasta.Tag = ISNULLFALSE(rs!REQUERIR_NRO_HASTA)
         
         
         If IsNull(rs!TITULO_DESCRIPCION) Then
            lblTituloDescripcion.Caption = "Descripción:"
         Else
            lblTituloDescripcion.Caption = rs!TITULO_DESCRIPCION
         End If
         
         
         txtDescripcion.Tag = ISNULLFALSE(rs!REQUERIR_DESCRIPCION)
         
         
       If IsNull(rs!HABILITAR_FECHA_DESDE) Then
            txtFechaDesde.Enabled = False
            txtFechaDesde.BackColor = ColorDesaHabilitado
        Else
            If rs!HABILITAR_FECHA_DESDE = True Then
                txtFechaDesde.Enabled = True
                
                txtFechaDesde.BackColor = ColorHabilitado
            Else
                txtFechaDesde.Enabled = False
                txtFechaDesde.BackColor = ColorDesaHabilitado
            End If
       End If
       
       
       If IsNull(rs!HABILITAR_FECHA_HASTA) Then
          Rem   txtFechaHasta.Enabled = False
            txtFechaHasta.BackColor = ColorDesaHabilitado
       Else
            If rs!HABILITAR_FECHA_HASTA = True Then
                txtFechaHasta.Enabled = True
                txtFechaHasta.BackColor = ColorHabilitado
            Else
                Rem txtFechaHasta.Enabled = False
                txtFechaHasta.BackColor = ColorDesaHabilitado
            End If
       End If
        
        
        
       If IsNull(rs!HABILITAR_LETRA_DESDE) Then
            Rem txtLetraDesde.Enabled = False
            txtLetraDesde.BackColor = ColorDesaHabilitado
       Else
            If rs!HABILITAR_LETRA_DESDE = True Then
                txtLetraDesde.Enabled = True
                txtLetraDesde.BackColor = ColorHabilitado
            Else
                Rem txtLetraDesde.Enabled = False
                txtLetraDesde.BackColor = ColorDesaHabilitado
            End If
       End If
       
       
       
       If IsNull(rs!HABILITAR_LETRA_HASTA) Then
            Rem txtLetraHasta.Enabled = False
            txtLetraHasta.BackColor = ColorDesaHabilitado
       Else
            If rs!HABILITAR_LETRA_HASTA = True Then
                txtLetraHasta.Enabled = True
                txtLetraHasta.BackColor = ColorHabilitado
            Else
               Rem  txtLetraHasta.Enabled = False
                txtLetraHasta.BackColor = ColorDesaHabilitado
            End If
            
       End If
       
       If IsNull(rs!HABILITAR_NRO_DESDE) Then
           Rem  txtNroDesde.Enabled = False
            txtNroDesde.BackColor = ColorDesaHabilitado
       Else
            If rs!HABILITAR_NRO_DESDE = True Then
                txtNroDesde.Enabled = True
                txtNroDesde.BackColor = ColorHabilitado
            Else
               Rem txtNroDesde.Enabled = False
                txtNroDesde.BackColor = ColorDesaHabilitado
            End If
       End If
       
      If IsNull(rs!HABILITAR_NRO_HASTA) Then
          Rem txtNroHasta.Enabled = False
           txtNroHasta.BackColor = ColorDesaHabilitado
      Else
            If rs!HABILITAR_NRO_HASTA = True Then
                txtNroHasta.Enabled = True
                txtNroHasta.BackColor = ColorHabilitado
             Else
                Rem txtNroHasta.Enabled = False
                txtNroHasta.BackColor = ColorDesaHabilitado
             End If
      End If
      
      If IsNull(rs!HABILITAR_DESCRIPCION) Then
         Rem  txtDescripcion.Enabled = False
         txtDescripcion.BackColor = ColorDesaHabilitado
      Else
            If rs!HABILITAR_DESCRIPCION = True Then
                txtDescripcion.Enabled = True
                txtDescripcion.BackColor = ColorHabilitado
            Else
             Rem txtDescripcion.Enabled = False
                txtDescripcion.BackColor = ColorDesaHabilitado
            End If
      End If
      
      
      If rs!COPIAR_FECHA = True Then
            chk_Copiar_Fecha.value = 1
      Else
            chk_Copiar_Fecha.value = 0
      End If
      
     If rs!COPIAR_LETRA = True Then
        chk_Copiar_Letra.value = 1
     Else
        chk_Copiar_Letra.value = 0
     End If
     
     
     If rs!COPIAR_NRO = True Then
          chk_Copiar_Nro.value = 1
     Else
          chk_Copiar_Nro.value = 0
     End If
Else
    Configurar_Carga = False
    MsgBox "NO EXISTE EL DOCUMENTO"
        lblIndice_Descripcion.Caption = ""
            txtFechaDesde.BackColor = ColorDesaHabilitado
            txtFechaDesde.Enabled = False
            txtFechaDesde.Text = ""
            
            txtFechaHasta.BackColor = ColorDesaHabilitado
            txtFechaHasta.Enabled = False
            txtFechaHasta.Text = ""
            
            
            
            txtNroDesde.BackColor = ColorDesaHabilitado
            txtNroDesde.Enabled = False
            txtNroDesde.Text = ""
            
            
            txtNroHasta.BackColor = ColorDesaHabilitado
            txtNroHasta.Enabled = False
            txtNroHasta.Text = ""
            
            
            txtLetraDesde.BackColor = ColorDesaHabilitado
            txtLetraDesde.Enabled = False
            txtLetraDesde.Text = ""
            
            txtLetraHasta.BackColor = ColorDesaHabilitado
            txtLetraHasta.Enabled = False
            txtLetraHasta.Text = ""
            
            
            txtDescripcion.BackColor = ColorDesaHabilitado
            txtDescripcion.Enabled = False
            txtDescripcion.Text = ""
            
        




     
    End If
    
  Exit Function
salir:
    MsgBox "Error en el indice", vbCritical
End Function

Private Sub txtTipoDocumento_LostFocus()
Configurar_Carga ctlCliente.Valor, txtIndice_Nro_Documento.Text
End Sub

Private Sub Check1_Click()

End Sub

Private Sub cboTipoCarga_Click()
cmdAceptar.Visible = True
cmdInforme.Visible = True
cmdActualizar.Visible = True
fraLegajo.Visible = False
fraCajaDocumento.Visible = False
fraRearchivo.Visible = False
grdDatos.Visible = False
fraCampos.Visible = False
fraReferencias.Visible = False
SSTab1.Visible = True
Select Case cboTipoCarga.Text
Case "Legajos"

fraLegajo.Visible = True
fraCajaDocumento.Visible = True
fraRearchivo.Visible = False
grdDatos.Visible = True
fraCampos.Visible = True
StatusBar.Panels(1).Text = "Nuevo"
Case "Referencia"
        fraReferencias.Visible = True
        fraCajaDocumento.Visible = True
        fraRearchivo.Visible = False
        grdDatos.Visible = True
        fraCampos.Visible = True
        StatusBar.Panels(1).Text = "Nuevo"

Case "Rearchivo"

End Select
End Sub

Private Sub cboTipoCarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        If cboTipoCarga.Text = "Legajos" Then
        
         txtEtiqueta.SetFocus
        End If
        If cboTipoCarga.Text = "Referencia" Then
           txtCaja.SetFocus
        End If
    
    End If

End Sub

Private Sub cmdAceptar_Click()
    Dim UsuarioCarga As Integer
    Dim FK_INDICES As String
    Dim Indice As String
    Dim LETRA_DESDE As String
    Dim LETRA_HASTA As String
    Dim NRO_DESDE As String
    Dim NRO_HASTA As String
    Dim FECHA_DESDE As String
    Dim FECHA_HASTA As String
    Dim Descripcion As String
    Dim FECHA_ACTUALIZACION As String
    Dim NRO_CAJA As String
    Dim FK_CLIENTE As String
    Dim RemitoProv As String
    Dim PLANILLA As String
    Dim rs As New ADODB.Recordset
    Dim Sql As String
        On Error GoTo salir:
    
    If StatusBar.Panels(1).Text = "" Then
        MsgBox "Error en el estado", vbCritical
        Exit Sub
    End If
    
    If Trim(txtCaja.Text) = "" Then
        MsgBox "Ingrese la caja ", vbCritical
        Exit Sub
    End If
    
    
    If IsNull(ctlCliente.Valor) Then
        MsgBox "Ingrese el cliente", vbCritical
        Exit Sub
    Else
        FK_CLIENTE = ctlCliente.Valor
    End If
    
    
    

    If txtIndice_Nro_Documento.Text = "" Then
        MsgBox "Ingrese el cliente", vbCritical
        Exit Sub
    Else
        Sql = " SELECT     ID, COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE"
        Sql = Sql & " From INDICES "
        Sql = Sql & "  Where Cod_cliente = " & ctlCliente.Valor
        Sql = Sql & "  And ID_CODIGO_DOCUMENTO = " & txtIndice_Nro_Documento.Text
  
        
        rs.Open Sql, strConBasa, 0, 1
        If rs.EOF Then
            MsgBox "Error en indice"
            Exit Sub
        Else
            FK_INDICES = rs!ID
            Indice = "'" & rs!Indice & "'"
        
        End If
        
      
  

    End If
    
    If Trim(txtCaja.Text) = "" Then
        MsgBox "Ingrese la caja", vbCritical
        Exit Sub
    Else
        NRO_CAJA = txtCaja.Text
    End If


    
    
    

    
    
    If txtUsuarioCarga.Text = "" Then
        MsgBox "Ingrese el usuario", vbCritical
        Exit Sub
    Else
        UsuarioCarga = txtUsuarioCarga.Text
    End If
    
    

    If Trim(txtLetraDesde.Text) <> "" Then
        LETRA_DESDE = "'" & UCase(Trim(txtLetraDesde.Text)) & "'"
    Else
    If txtLetraDesde.Tag <> "" Then
        If txtLetraDesde.Tag = True Then
            MsgBox "El dato Letra desde es requerido", vbCritical
            Exit Sub
        Else
           LETRA_DESDE = "Null"
        End If
     End If
    End If
    
    If Trim(txtLetraHasta.Text) <> "" Then
        LETRA_HASTA = "'" & UCase(Trim(txtLetraHasta.Text)) & "'"
    Else
       If txtLetraHasta.Tag <> "" Then
        If txtLetraHasta.Tag = True Then
            MsgBox "El dato Letra hasta es requerido", vbCritical
            Exit Sub
        Else
                If LETRA_DESDE = "Null" Then
                    LETRA_HASTA = "Null"
                Else
            End If
        End If
        End If
    End If
    
    
    If Trim(txtNroDesde.Text) <> "" Then
       If IsNumeric(txtNroDesde.Text) Then
           NRO_DESDE = Trim(txtNroDesde.Text)
       Else
            txtNroDesde.Text = ""
            MsgBox "NO ES UN NUMERO"
            Exit Sub
       End If
       
           
    Else
        If txtNroDesde.Tag <> "" Then
        If txtNroDesde.Tag = True Then
            MsgBox "El dato Nro desde es requerido", vbCritical
            Exit Sub
        Else
            NRO_DESDE = "Null"
        End If
        End If
        
        
    End If
    
    
    If Trim(txtNroHasta.Text) <> "" Then
        If IsNumeric(txtNroHasta.Text) Then
            NRO_HASTA = Trim(txtNroHasta.Text)
        Else
            txtNroHasta.Text = ""
            MsgBox "NO ES UN NUMERO"
            Exit Sub
        End If
        
         
            
    Else
    If txtNroHasta.Tag <> "" Then
        If txtNroHasta.Tag = True Then
            MsgBox "El dato Nro hasta es requerido", vbCritical
            Exit Sub
        Else
         If NRO_DESDE = "Null" Then
              NRO_HASTA = "Null"
         Else
            MsgBox "El dato Numero Desde puede se nulo ", vbCritical
            Exit Sub
         End If
        End If
        End If
        
    End If
    
    
    If Trim(txtFechaDesde.Text) <> "" Then
        If IsDate(Trim(txtFechaDesde.Text)) Then
                FECHA_DESDE = FechaFormato(Trim(txtFechaDesde.Text))
        Else
            MsgBox "La Fecha Ingresada es incorrecta", vbCritical
            Exit Sub
        End If
        
    Else
     If txtFechaDesde.Tag <> "" Then
        If txtFechaDesde.Tag = True Then
           MsgBox "El dato Fecha Desde es requerido", vbCritical
           Exit Sub
        Else
            FECHA_DESDE = "Null"
        End If
        End If
    End If
    
    If Trim(txtFechaHasta.Text) <> "" Then
        FECHA_HASTA = FechaFormato(Trim(txtFechaHasta.Text))
    Else
       If txtFechaHasta.Tag <> "" Then
        If txtFechaHasta.Tag = True Then
            MsgBox "El dato Fecha hasta es requerido", vbCritical
            Exit Sub
        Else
            If FECHA_DESDE = "Null" Then
                FECHA_HASTA = "Null"
            Else
                MsgBox "La Fecha Hasta no puede ser Nula", vbCritical
                Exit Sub
            End If
            End If
            
        End If
    End If
    
    If Trim(txtDescripcion.Text) <> "" Then
        Descripcion = "'" & Trim(txtDescripcion.Text) & "'"
    Else
        If txtDescripcion.Tag = True Then
            MsgBox "El dato descripcion es requerido", vbCritical
            Exit Sub
        Else
            Descripcion = "Null"
        End If
    
    End If
    
    FECHA_ACTUALIZACION = SysDateMinutoSegundo
    

    If cboTipoCarga = "Legajos" Then
            If txtEtiqueta.Text = "" Then
                MsgBox "Ingrese la etiqueta cliente", vbCritical
                Exit Sub
             End If
             
             If txtEtiquetaDigitoVerificador.Text = "" Then
                MsgBox "INGRESE EL DIGITO VERIFICADOR"
                Exit Sub
             Else
             
             
             If txtEtiqueta.Text > 4794261 Then
             
                If BuscarDigitoVerificador(txtEtiqueta.Text) = txtEtiquetaDigitoVerificador.Text Then
                    
                    Else
                       MsgBox "Error en el numero de Etiqueta"
                       Exit Sub
                    
                    End If
             
             
             Else
             
                 If Digito_Verificador(txtEtiqueta.Text) = txtEtiquetaDigitoVerificador.Text Then
                 
                 Else
                    MsgBox "Error en el numero de Etiqueta"
                    Exit Sub
                 
                 End If
               End If
               
             End If
             
             
'            If mskRemitoProv.Text = "0001-000_____" Then
'                MsgBox "Ingrese el Numero de Remito", vbCritical
'                Exit Sub
'             Else
'                RemitoProv = "'" & Trim(mskRemitoProv.Text) & "'"
'             End If
             
             RemitoProv = "NULL"
             
            If cboTipoCarga.Text = "Legajos" Then
                ActualizarLegajos FK_INDICES, Indice, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, _
                FECHA_DESDE, FECHA_HASTA, Descripcion, FECHA_ACTUALIZACION, NRO_CAJA, FK_CLIENTE, CStr(UsuarioCarga), RemitoProv, CLng(txtEtiqueta.Text)
                If chkAutoIncrementar.value = 1 Then
                    txtEtiqueta.Text = txtEtiqueta.Text + 1
                    txtEtiquetaDigitoVerificador.Text = Digito_Verificador(txtEtiqueta.Text)
                End If
            End If
    
    
  
    
     LimpiarCampos True
If chkFijarCaja.value = 0 Then
    txtCaja.SetFocus
 End If
 If chkFijarTipoDocumento.value = 0 Then
    txtIndice_Nro_Documento.SetFocus
 End If
 
 If chkFijarLetraDesde.value = 0 Then
     If txtLetraDesde.Enabled = True Then
        txtLetraDesde.SetFocus
     End If
 End If
 
 If chkFijarNumeroDesde.value = 0 Then
  If txtNroDesde.Enabled = True Then
    txtNroDesde.SetFocus
  End If
    
 End If
 
 
 If chkFijarFechaDesde.value = 0 Then
  If txtFechaDesde.Enabled = True Then
    txtFechaDesde.SetFocus
  End If
 End If
 
    End If
  If cboTipoCarga.Text = "Referencia" Then
  If Trim(txtLoteReferencia.Text) = "" Then
    MsgBox "Ingrese lote"
    Exit Sub
  End If
        GrabarReferencias Indice
 End If
    

    
    Exit Sub
salir:
MsgBox "Error en dato", vbCritical


End Sub


Private Sub cmdActualizar_Click()
    Dim rsGrilla As ADODB.Recordset
Set rsGrilla = New ADODB.Recordset
    rsGrilla.CursorLocation = adUseClient
    Dim Sql As String
    Dim C As Integer
    On Error GoTo salir
 Select Case cboTipoCarga.Text
 Case "Legajos"
 
 
        For C = 0 To grdDatos.Columns.Count - 1
            grdAnchoDatos(C) = grdDatos.Columns(C).Width
        Next
        
         
           
           
           Sql = " SELECT  TOP 200   ID_LEGAJO, NRO_CAJA AS CAJA , COD_CLIENTE AS CLIENTE , LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION, "
           Sql = Sql & " FECHA_ACTUALIZACION AS FECHA , FK_PERSONAL_ACTUALIZACION   as PERSONAL"
           Sql = Sql & " From LEGAJOS"
           Sql = Sql & " WHERE  FK_PERSONAL_ACTUALIZACION   = " & txtUsuarioCarga.Text
           Sql = Sql & " AND FECHA_ACTUALIZACION > " & FechaFormato(DateAdd("D", -10, Format(Now, "DD/MM/YYYY"))) & ""
           Sql = Sql & " ORDER BY  FECHA_ACTUALIZACION DESC"
           rsGrilla.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
           
           
           
           
           
            Set grdDatos.DataSource = rsGrilla.DataSource
            grdDatos.Columns("ID_LEGAJO").Locked = True
            grdDatos.Columns("FECHA").Locked = True
            grdDatos.Columns("PERSONAL").Locked = True
            
             grdDatos.Columns.Item(0).Width = 990
            grdDatos.Columns.Item(1).Width = 855
            grdDatos.Columns.Item(2).Width = 960
            grdDatos.Columns.Item(3).Width = 2310
            grdDatos.Columns.Item(4).Width = 1230
            grdDatos.Columns.Item(5).Width = 1095
            grdDatos.Columns.Item(6).Width = 1080
            grdDatos.Columns.Item(7).Width = 1245
            grdDatos.Columns.Item(8).Width = 1290
            grdDatos.Columns.Item(9).Width = 1200
            grdDatos.Columns.Item(10).Width = 1740
            grdDatos.Columns.Item(11).Width = 1005
            
  
  Case "Referencia"
        
    If txtID_Referencia.Text = "" Then
     
        Sql = SqlReferencia
        Sql = Sql & "  WHERE     REFERENCIAS.FK_PERSONAL_MODIFICACION =" & txtUsuarioCarga.Text
        Sql = Sql & "  AND  REFERENCIAS.FECHA_MODIFICACION > " & FechaFormato(DateAdd("D", -10, Now))
        Sql = Sql & "  ORDER BY REFERENCIAS.FECHA_MODIFICACION DESC "
   Else
   
        Sql = SqlReferencia
        Sql = Sql & "  WHERE     COD_ID_REFERENCIA in( " & txtID_Referencia.Text & ")"
        Sql = Sql & "  ORDER BY FECHA_MODIFICACION DESC "
    End If
   
        
        rsGrilla.Open Sql, ConActiva, 0, 1
        Set grdDatos.DataSource = rsGrilla.DataSource
         grdDatos.Columns.Item(0).Width = 780
 grdDatos.Columns.Item(1).Width = 510
 grdDatos.Columns.Item(2).Width = 2294
 grdDatos.Columns.Item(3).Width = 1305
 grdDatos.Columns.Item(4).Width = 1289
 grdDatos.Columns.Item(5).Width = 1170
 grdDatos.Columns.Item(6).Width = 1124
 grdDatos.Columns.Item(7).Width = 1275
 grdDatos.Columns.Item(8).Width = 1319
 grdDatos.Columns.Item(9).Width = 764
 grdDatos.Columns.Item(10).Width = 1604
 grdDatos.Columns.Item(11).Width = 945

  
         
 End Select
 txtID_Referencia.Text = ""
 
 Exit Sub
salir:
  MsgBox Err.Description
End Sub

Private Sub cmdBorrarEtiqueta_Click()

            Dim Sql As String
            If Digito_Verificador(txtEtiqueta.Text) = txtEtiquetaDigitoVerificador.Text Then
                Sql = " UPDATE    LEGAJOS"
                Sql = Sql & "  SET LETRA_DESDE = NULL, LETRA_HASTA = NULL, NRO_DESDE = NULL, NRO_HASTA = NULL, FECHA_DESDE = NULL, FECHA_HASTA = NULL,"
                Sql = Sql & " DESCRIPCION = NULL, NRO_CAJA = NULL, COD_CLIENTE = NULL, ID_PERSONAL = NULL, FK_PERSONAL_CREACION = NULL,"
                Sql = Sql & " FECHA_ACTUALIZACION = NULL, FECHA_CREACION = NULL,COD_ESTADO=NULL, COD_INDICE = NULL, FK_INDICES = NULL"
                Sql = Sql & " Where    ID_CLIENTE_LEGAJO = " & txtEtiqueta.Text
                Sql = Sql & " AND COD_CLIENTE = " & ctlCliente.Valor
                If MsgBox("Esta usted seguro de borrar el registro", vbCritical + vbYesNo) = vbYes Then
                    ExecutarSql Sql
                End If
            Else
                MsgBox "Ingrese el digito verificador correcto"
            End If

    
End Sub

Private Sub cmdBorrarFiltro_Click()
txtID_Referencia.Text = ""
End Sub

Private Sub cmdBuscarCaja_Click()
  Dim rsGrilla As ADODB.Recordset
Set rsGrilla = New ADODB.Recordset
    rsGrilla.CursorLocation = adUseClient
    Dim Sql As String
    Dim C As Integer
    On Error GoTo salir
 Select Case cboTipoCarga.Text
 Case "Legajos"
 
 
        For C = 0 To grdDatos.Columns.Count - 1
            grdAnchoDatos(C) = grdDatos.Columns(C).Width
        Next
        
         
           
           
           Sql = " SELECT  TOP 200   ID_LEGAJO, NRO_CAJA AS CAJA , COD_CLIENTE AS CLIENTE , LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION, "
           Sql = Sql & " FECHA_ACTUALIZACION AS FECHA , FK_PERSONAL_ACTUALIZACION   as PERSONAL"
           Sql = Sql & " From LEGAJOS"
           Sql = Sql & " WHERE  COD_CLIENTE  = " & ctlCliente.Valor
           Sql = Sql & " AND NRO_CAJA= " & txtCaja.Text
           rsGrilla.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
           
           
           
           
           
            Set grdDatos.DataSource = rsGrilla.DataSource
            grdDatos.Columns("ID_LEGAJO").Locked = True
            grdDatos.Columns("FECHA").Locked = True
            grdDatos.Columns("PERSONAL").Locked = True
            
             grdDatos.Columns.Item(0).Width = 990
            grdDatos.Columns.Item(1).Width = 855
            grdDatos.Columns.Item(2).Width = 960
            grdDatos.Columns.Item(3).Width = 2310
            grdDatos.Columns.Item(4).Width = 1230
            grdDatos.Columns.Item(5).Width = 1095
            grdDatos.Columns.Item(6).Width = 1080
            grdDatos.Columns.Item(7).Width = 1245
            grdDatos.Columns.Item(8).Width = 1290
            grdDatos.Columns.Item(9).Width = 1200
            grdDatos.Columns.Item(10).Width = 1740
            grdDatos.Columns.Item(11).Width = 1005
            
  
  Case "Referencia"
        
   
    Sql = SqlReferencia
    Sql = Sql & "  WHERE     REFERENCIAS.COD_CLIENTE= " & ctlCliente.Valor
    Sql = Sql & " AND REFERENCIAS.NRO_CAJA =" & txtCaja.Text
    Sql = Sql & "  ORDER BY REFERENCIAS.FECHA_MODIFICACION DESC "
        
        rsGrilla.Open Sql, ConActiva, 0, 1
        Set grdDatos.DataSource = rsGrilla.DataSource
         grdDatos.Columns.Item(0).Width = 780
 grdDatos.Columns.Item(1).Width = 510
 grdDatos.Columns.Item(2).Width = 2294
 grdDatos.Columns.Item(3).Width = 1305
 grdDatos.Columns.Item(4).Width = 1289
 grdDatos.Columns.Item(5).Width = 1170
 grdDatos.Columns.Item(6).Width = 1124
 grdDatos.Columns.Item(7).Width = 1275
 grdDatos.Columns.Item(8).Width = 1319
 grdDatos.Columns.Item(9).Width = 764
 grdDatos.Columns.Item(10).Width = 1604
 grdDatos.Columns.Item(11).Width = 945

  
         
 End Select
 txtID_Referencia.Text = ""
 
 Exit Sub
salir:
  MsgBox Err.Description
End Sub

Private Sub cmdBuscarDocumento_Click()
    frmIndice.COD_CLIENTE = ctlCliente.Valor
    frmIndice.Actualizar ctlCliente.Valor, Nulo, 0
    frmAgregarDocumentos.WindowState = 0
    frmIndice.Show
    frmIndice.SetFocus


End Sub

Private Sub cmdBuscarRemito_Click()
If IsNull(ctlCliente.Valor) Then
 MsgBox "Ingrese el cliente"
 Exit Sub
End If

If txtCaja.Text = "" Then
 MsgBox "Ingrese la caja"
 Exit Sub
End If

Dim rs As New ADODB.Recordset
Dim Sql As String


Sql = " SELECT     REMITOS_CUERPO.NRO_REM_PROV, REMITOS_CUERPO.TIPO, REMITOS_CUERPO.ID_CLIENTE, REMITOS_DETALLE.DESDE,"
Sql = Sql & " REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO"
Sql = Sql & " FROM         REMITOS_CUERPO INNER JOIN"
Sql = Sql & "  REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO"
Sql = Sql & " Where(REMITOS_CUERPO.Tipo = 0) And (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0)"
Sql = Sql & " AND REMITOS_CUERPO.ID_CLIENTE =  " & ctlCliente.Valor
Sql = Sql & " AND REMITOS_DETALLE.DESDE = " & txtCaja.Text

rs.Open Sql, ConActiva, 0, 1


End Sub

Private Sub cmdCajasConLegajos_Click()
 If txtCaja.Text <> "" And Not IsNull(ctlCliente.Valor) Then
        Dim Sql As String
            Sql = " Update dbo.Cajas "
            Sql = Sql & " Set FK_TIPO_REFERENCIA = 1015 "
            Sql = Sql & " , FK_TIPO_REFERENCIA_PERSONAL =" & MDIfrmInicio.StaInicio.Panels(2).Text
            Sql = Sql & " Where FK_CLIENTE =" & ctlCliente.Valor
            Sql = Sql & " And NRO_CAJA = " & txtCaja.Text
            Rem SQL = SQL & " AND FK_TIPO_REFERENCIA is null "
            ExecutarSql Sql
            MsgBox "La caja se registro como para carga de legajos", vbInformation
    End If

End Sub

Private Sub cmdCargaCompletaCajaLegajo_Click()
Dim rs As New ADODB.Recordset
    If txtCaja.Text <> "" And Not IsNull(ctlCliente.Valor) Then
        Dim Sql As String
        
        Sql = "SELECT     COUNT(*) as cantidad"
        Sql = Sql & " From basasql.dbo.LEGAJOS"
        Sql = Sql & " Where NRO_CAJA =  " & txtCaja.Text
        Sql = Sql & " And COD_CLIENTE = " & ctlCliente.Valor
        rs.Open Sql, strConBasa
        
    If Not rs.EOF Then
        If rs!cantidad > 0 Then
            Sql = "Update dbo.Cajas "
            Sql = Sql & " Set FK_TIPO_REFERENCIA = 1020"
            Sql = Sql & " Where FK_CLIENTE =" & ctlCliente.Valor
            Sql = Sql & " and NRO_CAJA = " & txtCaja.Text
            ExecutarSql Sql
            MsgBox "La caja se registro como carga de legajos completa", vbInformation
            Rem cmdCargaCompletaCajaLegajo.Enabled = False
            Unload Me
        Else
            MsgBox "La caja NO TIENE LEGAJOS", vbCritical
        End If
   End If
            
    End If
    txtCaja.Text = ""
    txtCaja.Enabled = False
End Sub

Private Sub cmdCargarVideo_Click()

Video.URL = txtPasoVideo.Text
Video.Controls.play
ValorAnteVideo = 100
Timer1.Enabled = True

cmdPlay_Pausa.Caption = "Pausa"
End Sub

Private Sub cmdCopiarExcel_Click()
    CopiarDatosGrilla grdDatos
End Sub

Private Sub cmddescripcion_Click()
    Dim sql1 As String
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim DATO As String
        txtID_Referencia.Text = ""
        txtID_Referencia.Text = ""
        MousePointer = 11
        If txtCaja.Text = "" Then
            MsgBox "Ingrese la caja"
            MousePointer = 0
            Exit Sub
        End If
        sql1 = " SELECT ID_LEGAJO,  ID_PERSONAL,CANTIDAD_CARACTERES,"
        sql1 = sql1 & " ID_CLIENTE_LEGAJO,NRO_CAJA, DESCRIPCION_REMITO ,"
        sql1 = sql1 & " COD_CLIENTE , NRO_DESDE, LETRA_DESDE, FECHA_DESDE"
        sql1 = sql1 & " ,CLIENTE_LEGAJO, DESCRIPCION, NOMBRE,DESCRIPCION_REMITO "
        sql1 = sql1 & " From LEGAJOS"
        sql1 = sql1 & " Where  Cod_cliente = " & ctlCliente.Valor
        sql1 = sql1 & " AND NRO_CAJA = " & txtCaja.Text
        sql1 = sql1 & " ORDER BY  ID_LEGAJO  "
        Rem Dim rs As New ADODB.Recordset
        Dim detalle As String
        rs.Open sql1, strConBasa
        Do While Not rs.EOF
            detalle = rs!NRO_DESDE & "-" & Trim(rs!LETRA_DESDE) & "-" & Format(rs!FECHA_DESDE, "YYYY") & "  " & Trim(rs!Descripcion)
            sql1 = "   Update basasql.dbo.LEGAJOS"
            sql1 = sql1 & " SET DESCRIPCION_REMITO = '" & Trim(detalle) & "'"
            sql1 = sql1 & " Where ID_LEGAJO = " & rs!ID_LEGAJO
            ExecutarSql sql1
            rs.MoveNext
        Loop
        MousePointer = 0
        MsgBox "Terminado"
End Sub

Private Sub cmde_Click()
Debug.Print Video.currentMedia.attributeCount
End Sub

Private Sub cmdInforme_Click()
    Dim sql1 As String
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim DATO As String
    txtID_Referencia.Text = ""
    MousePointer = 11
    If txtCaja.Text = "" Then
        MsgBox "Ingrese la caja"
        Exit Sub
    End If
'    If cmdCargaCompletaCajaLegajo.Enabled = True Then
''        If MsgBox("¿Usted confirmo la carga completa de la caja?" & vbCrLf & " ¿Quiere continuar?", vbYesNo) = vbYes Then
''            Sql = " Update dbo.Cajas "
''            Sql = Sql & " Set FK_TIPO_REFERENCIA = 1020"
''            Sql = Sql & " Where FK_CLIENTE =" & ctlCliente.Valor
''            Sql = Sql & " And NRO_CAJA = " & txtCaja.Text
''            ExecutarSql Sql
''            MsgBox "La caja se registro como carga de legajos completa", vbInformation
''        Else
''          Rem  Exit Sub
''        End If
'     End If
    sql1 = " SELECT  ID_PERSONAL,CANTIDAD_CARACTERES,"
    sql1 = sql1 & " ID_CLIENTE_LEGAJO,NRO_CAJA, DESCRIPCION_REMITO ,"
    sql1 = sql1 & " COD_CLIENTE , NRO_DESDE, LETRA_DESDE, FECHA_DESDE"
    sql1 = sql1 & " ,CLIENTE_LEGAJO, DESCRIPCION, NOMBRE,DESCRIPCION_REMITO,ID_LEGAJO "
    sql1 = sql1 & " From LEGAJOS"
    sql1 = sql1 & " Where  Cod_cliente = " & ctlCliente.Valor
    sql1 = sql1 & " AND nro_caja = " & txtCaja.Text
    sql1 = sql1 & " ORDER BY  ID_LEGAJO  "
    txtCaja.Text = ""
    frmReportes.ImprimirReporte PasoReportes & "rptLegajosControl2.rpt", sql1, True
    MousePointer = 0

End Sub

Public Sub Modificar_Legajos(FK_CLIENTES As Integer, Etiqueta As Long)
 Dim rsGrilla As ADODB.Recordset
Set rsGrilla = New ADODB.Recordset
    rsGrilla.CursorLocation = adUseClient
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    
 ctlCliente.Valor = FK_CLIENTES
 
  cboTipoCarga.ListIndex = 1
 
 StatusBar.Panels(1).Text = "Modificar"
  StatusBar.Panels(2).Text = txtEtiqueta.Text
'    Sql = " SELECT  TOP 40   ID_LEGAJO, NRO_CAJA, COD_CLIENTE, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION, NRO_REM_PROV ,"
'    Sql = Sql & " FECHA_ACTUALIZACION AS FECHA , ID_Personal as PERSONAL"
'    Sql = Sql & " From LEGAJOS"
'    Sql = Sql & " WHERE ID_CLIENTE_LEGAJO  > " & txtEtiqueta.Text - 10
'    Sql = Sql & " AND  COD_CLIENTE = " & ctlCliente.Valor
'    Sql = Sql & " ORDER BY ID_LEGAJO "
'    rsGrilla.Open Sql,ConActiva, adOpenDynamic, adLockOptimistic
'    Set grdDatos.DataSource = rsGrilla.DataSource
'    grdDatos.Columns("ID_LEGAJO").Locked = True
'    grdDatos.Columns("FECHA").Locked = True
'    grdDatos.Columns("PERSONAL").Locked = True
    
    Sql = "  SELECT    LEGAJOS.COD_CLIENTE ,  LEGAJOS.ID_CLIENTE_LEGAJO, LEGAJOS.FK_INDICES, LEGAJOS.LETRA_DESDE, LEGAJOS.LETRA_HASTA, LEGAJOS.NRO_DESDE,"
    Sql = Sql & " LEGAJOS.NRO_HASTA, LEGAJOS.FECHA_DESDE, LEGAJOS.FECHA_HASTA, LEGAJOS.DESCRIPCION, "
    Sql = Sql & " INDICES.ID_CODIGO_DOCUMENTO , LEGAJOS.NRO_CAJA,LEGAJOS.COD_CLIENTE,  FK_PERSONAL_CREACION , FK_PERSONAL_ACTUALIZACION "
    Sql = Sql & " FROM         LEGAJOS LEFT OUTER JOIN"
    Sql = Sql & " INDICES ON LEGAJOS.FK_INDICES = INDICES.ID"
    Sql = Sql & " WHERE ID_CLIENTE_LEGAJO =  " & Etiqueta
    If Etiqueta < 300000 Then
    Sql = Sql & " AND  LEGAJOS.COD_CLIENTE = " & FK_CLIENTES
    End If
    
   
   
   Set rs = New ADODB.Recordset
   rs.Open Sql, ConActiva, 0, 1
   
   If Not rs.EOF Then
  If IsNull(rs!FK_PERSONAL_ACTUALIZACION) Then
  txtUsuarioCarga.Text = ""
  Else
 txtUsuarioCarga.Text = rs!FK_PERSONAL_ACTUALIZACION
  End If
   If Not IsNull(rs!NRO_CAJA) Then
   ctlCliente.Valor = rs!COD_CLIENTE
   txtCaja.Text = rs!NRO_CAJA
   txtEtiqueta.Text = rs!ID_CLIENTE_LEGAJO
   txtEtiquetaDigitoVerificador.Text = Digito_Verificador(rs!ID_CLIENTE_LEGAJO)
   Else
   txtCaja.Text = ""
   End If
   
   
   If Not IsNull(rs!ID_CODIGO_DOCUMENTO) Then
        txtIndice_Nro_Documento.Text = rs!ID_CODIGO_DOCUMENTO
        
   Else
        txtIndice_Nro_Documento.Text = ""
    End If
    
    If Not IsNull(rs!LETRA_DESDE) Then
        txtLetraDesde.Text = rs!LETRA_DESDE
    Else
        txtLetraDesde.Text = ""
    End If
    
    If Not IsNull(rs!LETRA_HASTA) Then
        txtLetraHasta.Text = rs!LETRA_HASTA
    Else
        txtLetraHasta.Text = ""
    End If
    
    If Not IsNull(rs!NRO_DESDE) Then
        txtNroDesde.Text = rs!NRO_DESDE
    Else
        txtNroDesde.Text = ""
    End If
    
    If Not IsNull(rs!NRO_HASTA) Then
        txtNroHasta.Text = rs!NRO_HASTA
    Else
        txtNroHasta.Text = ""
    End If
    
    If Not IsNull(rs!FECHA_DESDE) Then
        txtFechaDesde.Text = rs!FECHA_DESDE
    Else
        txtFechaDesde.Text = ""
    End If
    
     If Not IsNull(rs!FECHA_HASTA) Then
        txtFechaHasta.Text = rs!FECHA_HASTA
    Else
        txtFechaHasta.Text = ""
    End If
    
    If Not IsNull(rs!Descripcion) Then
        txtDescripcion.Text = rs!Descripcion
    Else
        txtDescripcion.Text = ""
    End If
  Else
  
  LimpiarCampos False
   End If
    
   
    
    

End Sub


Public Sub CargarReferencias(FK_CLIENTES As Integer, ID As Long, Accion As String)
    Dim rsGrilla As ADODB.Recordset
    Dim Sql As String
    Dim rs As New ADODB.Recordset

Set rsGrilla = New ADODB.Recordset
rsGrilla.CursorLocation = adUseClient
ctlCliente.Valor = FK_CLIENTES
cboTipoCarga.ListIndex = 0



   StatusBar.Panels(1).Text = Accion
   If ID <> 0 Then
      StatusBar.Panels(2).Text = ID
    Else
        StatusBar.Panels(2).Text = ""
    End If
    
    If Accion = "Nuevo" Then
        ctlCliente.Valor = FK_CLIENTES
       StatusBar.Panels(1).Text = "Nuevo"
       StatusBar.Panels(2).Text = ""
       LimpiarCampos True
       Exit Sub
    End If
    
    Sql = " SELECT dbo.REFERENCIAS.COD_CLIENTE, dbo.REFERENCIAS.NRO_CAJA, dbo.REFERENCIAS.INDICE, dbo.REFERENCIAS.COD_DOCUMENTO, "
    Sql = Sql & " dbo.REFERENCIAS.DESCRIPCION, dbo.REFERENCIAS.FECHA_DESDE, dbo.REFERENCIAS.FECHA_HASTA, dbo.REFERENCIAS.NRO_DESDE, "
    Sql = Sql & " dbo.REFERENCIAS.NRO_HASTA, dbo.REFERENCIAS.LETRA_DESDE, dbo.REFERENCIAS.LETRA_HASTA, dbo.REFERENCIAS.COD_ID_REFERENCIA, "
    Sql = Sql & " dbo.INDICES.ID_CODIGO_DOCUMENTO, dbo.INDICES.INDICE AS Expr1, dbo.INDICES.DESCRIPCION AS Expr2 , dbo.REFERENCIAS.FK_PERSONAL_MODIFICACION , dbo.REFERENCIAS.ID_IMAGEN  "
    Sql = Sql & " FROM dbo.REFERENCIAS INNER JOIN "
    Sql = Sql & " dbo.INDICES ON dbo.REFERENCIAS.INDICE = dbo.INDICES.INDICE AND dbo.REFERENCIAS.COD_CLIENTE = dbo.INDICES.COD_CLIENTE "
    Sql = Sql & " Where dbo.REFERENCIAS.COD_ID_REFERENCIA = " & ID
   
   Set rs = New ADODB.Recordset
   rs.Open Sql, ConActiva, 0, 1
   chkFijarTipoDocumento.value = 0
   
   If Not IsNull(rs!ID_imagen) Then
        If Accion <> "Descripcion" Then
           Rem ViewImg1.MostrarImagen PasoImagenes & BuscarDirectorioPaso(rs!ID_imagen) & "\" & Trim(rs!ID_imagen) & ".tif"
        End If
    End If
   If Not rs.EOF Then
   If IsNull(rs!FK_PERSONAL_MODIFICACION) Then
    txtUsuarioCarga.Text = "99"
   Else
    
      txtUsuarioCarga.Text = rs!FK_PERSONAL_MODIFICACION
    End If
    
   If Not IsNull(rs!NRO_CAJA) Then
   txtCaja.Text = rs!NRO_CAJA
  
  
   Else
   txtCaja.Text = ""
   End If
   
   
   If Not IsNull(rs!ID_CODIGO_DOCUMENTO) Then
        txtIndice_Nro_Documento.Text = rs!ID_CODIGO_DOCUMENTO
       Rem  txtIndice_Nro_Documento_Change
        chkFijarTipoDocumento.value = 1
   Else
        txtIndice_Nro_Documento.Text = ""
    End If

    If Not IsNull(rs!LETRA_DESDE) Then
        txtLetraDesde.Text = rs!LETRA_DESDE
        chkFijarLetraDesde.value = 0
    Else
        txtLetraDesde.Text = ""
        chkFijarLetraDesde.value = 1
    End If
    
    If Not IsNull(rs!LETRA_HASTA) Then
        txtLetraHasta.Text = rs!LETRA_HASTA
        chkFijarLetraHasta.value = 0
    Else
        txtLetraHasta.Text = ""
        chkFijarLetraHasta.value = 1
    End If
    
    If Not IsNull(rs!NRO_DESDE) Then
        txtNroDesde.Text = rs!NRO_DESDE
        chkFijarNumeroDesde.value = 0
    Else
        txtNroDesde.Text = ""
        chkFijarNumeroDesde.value = 1
    End If
    
    If Not IsNull(rs!NRO_HASTA) Then
        txtNroHasta.Text = rs!NRO_HASTA
        chkFijarNumeroHasta.value = 0
    Else
        txtNroHasta.Text = ""
        chkFijarNumeroHasta.value = 1
    End If
    
    If Not IsNull(rs!FECHA_DESDE) Then
        txtFechaDesde.Text = rs!FECHA_DESDE
        chkFijarFechaDesde.value = 0
    Else
        txtFechaDesde.Text = ""
        chkFijarFechaDesde.value = 1
    End If
    
     If Not IsNull(rs!FECHA_HASTA) Then
        txtFechaHasta.Text = rs!FECHA_HASTA
        chkFijarFechaHasta.value = 0
    Else
        txtFechaHasta.Text = ""
        chkFijarFechaHasta.value = 1
    End If
    
    If Not IsNull(rs!Descripcion) Then
        txtDescripcion.Text = rs!Descripcion
        chkFijarDescripcion.value = 0
    Else
        txtDescripcion.Text = ""
        chkFijarDescripcion.value = 1
    End If
  Else
  
  LimpiarCampos False
   End If
    
   If Accion = "NuevoCopiar" Then
        txtCaja.Text = ""
        
        StatusBar.Panels(1).Text = "Nuevo"
        StatusBar.Panels(2).Text = ""
    End If
    
    If Accion = "Descripcion" Then
        txtCaja.Text = ""
        txtCaja.SetFocus
        StatusBar.Panels(1).Text = "Nuevo"
        StatusBar.Panels(2).Text = ""
    End If

End Sub

Private Sub cmdModificar_Click()
Modificar_Legajos ctlCliente.Valor, txtEtiqueta.Text
End Sub

Private Sub cmdPausa_Click()
Video.Controls.pause
End Sub

Private Sub cmdPlay_Click()
Video.Controls.play
Timer1.Enabled = True
End Sub

Private Sub cmdPlay_Pausa_Click()
  Play_Pausa cmdPlay_Pausa.Caption

End Sub

Private Sub cmdRegistroVerificado_Click()
'Dim sql As String
'
'If txtEtiqueta.Text <> "" Then
'
'sql = " UPDATE    LEGAJOS "
'sql = sql & " Set REGISTRO_VERIFICADO = 1"
'sql = sql & " Where ID_LEGAJO = " & txtEtiqueta.Text
'ExecutarSql sql
'Else
'MsgBox "Ingrese la Etiqueta", vbCritical
'
'End If
StatusBar.Panels(1).Text = "Nuevo"
        StatusBar.Panels(2).Text = ""

End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
Dim NUMERO As Long
Dim lETRA As String
Dim fecha  As String
Dim Año As String
Dim i As Integer
Dim Legajo As String
Dim Pos1 As Integer
Dim Pos2 As Integer
ConBasa.CommandTimeout = 300000
Sql = " SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_CLIENTE, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA,NUMERO_LEGAJO_CLIENTE, CLIENTE_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE, ID_LEGAJO, COD_CLIENTE,Nombre, descripcion"
Sql = Sql & " From LEGAJOS"
Sql = Sql & "  WHERE     (COD_CLIENTE =118)"
Sql = Sql & " and   (FECHA_DESDE IS NULL) "
Sql = Sql & " ORDER BY ID_LEGAJO "
rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
On Error GoTo salir
Do While Not rs.EOF
    
    If Not IsNull(rs!CLIENTE_LEGAJO) Then
    lETRA = ""
    NUMERO = 0
    fecha = ""
    Legajo = Replace(Trim(rs!CLIENTE_LEGAJO), "_", "")
     Legajo = Replace(Trim(rs!CLIENTE_LEGAJO), "--", "-?-")
    If Not IsNumeric(Legajo) And Legajo <> "" Then
        
        Pos1 = InStr(1, Legajo, "-")
        
        If Pos1 <> 0 Then
                    If IsNumeric(Mid(Legajo, 1, Pos1 - 1)) Then
                    NUMERO = Mid(Legajo, 1, Pos1 - 1)
                    rs!NRO_DESDE = NUMERO
                     rs!NRO_HASTA = NUMERO
                    End If
                    Pos2 = InStr(Pos1 + 1, Legajo, "-")
                    
                    If Pos2 - (Pos1 + 1) > 0 Then
                    lETRA = Mid(Legajo, Pos1 + 1, Pos2 - (Pos1 + 1))
                    End If
                    fecha = Mid(Legajo, Pos2 + 1)
                
                If lETRA <> "" Then
                rs!LETRA_DESDE = lETRA
                rs!LETRA_HASTA = lETRA
                Else
                rs!LETRA_DESDE = Null
                rs!LETRA_HASTA = Null
                End If
                
                If fecha <> "" And Len(fecha) = 4 Then
                rs!FECHA_DESDE = "01/01/" & fecha
                rs!FECHA_HASTA = "31/12/" & fecha
                Else
                rs!FECHA_DESDE = Null
                rs!FECHA_HASTA = Null
                End If
    
                If Not IsNull(rs!Nombre) Then
                rs!Descripcion = Trim(rs!Descripcion & " " & Trim(rs!Nombre))
                
                End If
                
    
    End If
    Else
        
        If Legajo <> "" Then
        rs!NRO_DESDE = Legajo
        rs!NRO_HASTA = Legajo
        End If
    End If
    End If
    rs.Update
    
salir:
    If Err.Number <> 0 Then
    rs.CancelUpdate
    Err.Clear
    End If
    rs.MoveNext
Loop


End Sub

Private Sub Command2_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset


Sql = " SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO, LETRA_DESDE , LETRA_hasta ,CONTROL_EXPORT, COD_CLIENTE, CLIENTE_LEGAJO, COD_CLIENTE, NOMBRE, DESCRIPCION, NUMERO_LEGAJO_CLIENTE,  COD_INDICE, NRO_DESDE , NRO_HASTA"
Sql = Sql & " From LEGAJOS "
Sql = Sql & " Where (Cod_cliente = 06) And  nro_desde is null "
Rem sql = sql & "  AND  CONTROL_EXPORT IS NULL"
Sql = Sql & " ORDER BY ID_LEGAJO "

Sql = " SELECT     ID, ELEMENTO, NRO_DESDE, NRO_HASTA"
Sql = Sql & "  From dbo.ORDENAR_DOCUMENTACION_DETALLE"
Sql = Sql & "  Where (NRO_DESDE Is Null)"


rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
Dim Legajo As String
    Do While Not rs.EOF
        If IsNumeric(rs!Elemento) Then
                rs!NRO_DESDE = rs!Elemento
                rs!NRO_HASTA = rs!Elemento
                
        End If
        
                
                rs.Update
                rs.MoveNext
    Loop

End Sub

Private Sub Command3_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset


Sql = "SELECT    ID_LEGAJO, NOMBRE,  LETRA_DESDE,LETRA_HASTA,CONTROL_EXPORT "


Sql = Sql & " From LEGAJOS"
Sql = Sql & " Where (Cod_cliente = 20) And cod_indice= '002008002'"
Sql = Sql & " ORDER BY ID_LEGAJO "

Sql = " SELECT     FECHA_DESDE, FECHA_HASTA, COD_CLIENTE, LETRA_HASTA, LETRA_DESDE, COD_INDICE, ID_CLIENTE_LEGAJO"
Sql = Sql & " From dbo.LEGAJOS"
Sql = Sql & " WHERE     (NOT (FECHA_DESDE IS NULL)) AND (FECHA_HASTA IS NULL) AND (NOT (COD_CLIENTE IN (4, 20)))"
Sql = Sql & " ORDER BY FECHA_DESDE"

Sql = "  SELECT     ID_CLIENTE_LEGAJO, LEN(LETRA_DESDE) AS Expr2, FECHA_DESDE, DATEPART(MM, FECHA_HASTA) AS Expr1, NRO_DESDE, LETRA_DESDE,"
Sql = Sql & " FECHA_HASTA"
Sql = Sql & "  From dbo.LEGAJOS"
Sql = Sql & " WHERE     (DATEPART(MM, FECHA_HASTA) = 1) AND (LEN(LETRA_DESDE) = 1) AND (FECHA_HASTA > CONVERT(DATETIME, '1950-01-01 00:00:00', 102))"

Sql = "  SELECT     ID_LEGAJO, FECHA_HASTA, ID_CLIENTE_LEGAJO, COD_INDICE, LETRA_DESDE, LETRA_HASTA, NRO_HASTA, FECHA_DESDE, DATEPART(MM, FECHA_HASTA)"
Sql = Sql & " AS Expr1, CLIENTE_LEGAJO, COD_CLIENTE, FK_INDICES"
Sql = Sql & " From dbo.LEGAJOS"
Sql = Sql & " WHERE     (DATEPART(MM, FECHA_HASTA) = 1) AND (FK_INDICES IN (457, 1195, 1464, 1548, 1602, 2074, 2121, 4774)) OR"
Sql = Sql & " (DATEPART(dd, FECHA_HASTA) = 1)"


rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic

Do While Not rs.EOF
    If Mid(rs!FECHA_DESDE, 1, 5) = "01/01" Then
        rs!FECHA_HASTA = "31/12/" & Format(rs!FECHA_DESDE, "YYYY")
    Else
        rs!FECHA_HASTA = rs!FECHA_DESDE
    End If

        rs.Update
    rs.MoveNext
Loop

End Sub

Private Sub Command4_Click()

Dim Sql As String
Dim Sqlc As String
Dim rsContenedor As New ADODB.Recordset
Dim rsCajas As New ADODB.Recordset
Dim Legajo As String
Dim RSESTANTERIA  As New ADODB.Recordset



Sqlc = " SELECT     ID_CONTENEDOR,COD_CLIENTE, NRO_CAJA, FK_CAJAS "
Sqlc = Sqlc & " From CONTENEDOR "
Sqlc = Sqlc & " Where (FK_CAJAS Is Null) "
Sqlc = Sqlc & " And (Not (COD_CLIENTE Is Null))"
Sqlc = Sqlc & " ORDER BY COD_CLIENTE, NRO_CAJA "


Sql = "  SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_CONTENEDOR"
Sql = Sql & "  From Cajas "
Sql = Sql & "  Where (FK_CLIENTE Is Null)"
Sql = Sql & "  ORDER BY ID_CAJA "

rsContenedor.CursorLocation = adUseClient

rsContenedor.Open Sqlc, ConActiva, adOpenKeyset, adLockOptimistic
rsCajas.Open Sql, ConActiva, adOpenKeyset, adLockOptimistic

Do While Not rsContenedor.EOF
    rsContenedor!FK_CAJAS = rsCajas!ID_CAJA
    rsContenedor.Update
    rsCajas!NRO_CAJA = rsContenedor!NRO_CAJA
    rsCajas!FK_CLIENTE = rsContenedor!COD_CLIENTE
    rsCajas!FK_CONTENEDOR = rsContenedor!ID_CONTENEDOR
    rsCajas.Update
    rsCajas.MoveNext
    rsContenedor.MoveNext
Loop



Exit Sub



Dim i As Long




'For i = 42932 To 43291
'
' Set rs = New ADODB.Recordset
'    rs.Open "SELECT * FROM  CONTENEDOR WHERE  NRO_CAJA = " & i & " AND (COD_CLIENTE = 231)", strConBasa , 0 ,1
'        If rs.EOF Then
'
'
'        RSESTANTERIA!NRO_CAJA = i
'        RSESTANTERIA!COD_CLIENTE = 231
'        RSESTANTERIA!Estado = 5
'        RSESTANTERIA.Update
'        RSESTANTERIA.MoveNext
'       Else
'      Rem MsgBox ""
'        End If
'
'Next
'
'ConBasa.CommitTrans

'Dim con As New ADODB.Connection
'
'
'con.Open "Provider=SQLOLEDB.1;Password=21877471;Persist Security Info=True;User ID=usuario1;Initial Catalog=BASE_SOPORTE;Data Source=Serverbasa1"
'
'
'
'Dim rstablas As New ADODB.Recordset
'Dim RSPADRON As ADODB.Recordset
'Dim Sql As String
'
'Sql = " SELECT     name"
'Sql = Sql & " From dbo.sysobjects"
'Sql = Sql & "  WHERE     (xtype = 'u') AND name <> 'PADRON' AND name <> 'chacofem' "
'Sql = Sql & " AND name <> 'chacomas' AND name <> 'CHUFEM' AND  name <> 'CHUMAS'"
'Sql = Sql & "  ORDER BY name desc "
'
'Sql = " SELECT     name"
'Sql = Sql & " From dbo.sysobjects"
'Sql = Sql & "  WHERE     (xtype = 'u') AND name = 'mendofem' or  name = 'mendomas'"
'Sql = Sql & "  ORDER BY name desc "
'
'
'
'Dim DOCUMENTO As Double
'Dim Apellido As String
'rstablas.Open Sql, con
'Dim POS As Integer
'
'Dim RSCON As New ADODB.Recordset
'
'
''Do While Not rstablas.EOF
''
''    Set RSCON = New ADODB.Recordset
''   Sql = " SELECT     COUNT(*) AS CANTIDAD"
''Sql = Sql & " From  " & rstablas!Name
''
''RSCON.Open Sql, con
'' Debug.Print rstablas!Name & vbTab & RSCON!CANTIDAD
''
''    rstablas.MoveNext
''Loop
'
'
'
'On Error Resume Next
'
'
'Do While Not rstablas.EOF
'    Sql = "SELECT     MATRICULA, LINEAPAD"
'    Sql = Sql & "  From " & rstablas!Name
'
'
'    Set RSPADRON = New ADODB.Recordset
'        Dim TABLA As String
'    RSPADRON.Open Sql, con
'        Do While Not RSPADRON.EOF
'
'            DOCUMENTO = 0
'            Apellido = ""
'            DOCUMENTO = CDbl(RSPADRON!MATRICULA)
'
'            POS = InStr(1, RSPADRON!LINEAPAD, ",")
'
'            Apellido = Mid(RSPADRON!LINEAPAD, 1, POS - 1)
'
'            Apellido = UCase(Replace(Apellido, "Ð", "Ñ"))
'            Apellido = UCase(Replace(Apellido, "'", "´"))
'
''            If DOCUMENTO = 0 Then
''                MsgBox "QUE PASO"
''            End If
''
'
'    Sql = " INSERT INTO PADRON  (DOCUMENTO, APELLIDO_NOMBRE, ARCHIVO)"
'Sql = Sql & " VALUES     (" & DOCUMENTO & ",'" & Apellido & "','" & Trim(rstablas!Name) & "')"
'
'            con.Execute Sql
'        RSPADRON.MoveNext
'      Loop
'
'
'    rstablas.MoveNext
'Loop




End Sub

Private Sub cmdsolicituddecaja_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
        Sql = " SELECT TOP (1) ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_TIPO_REFERENCIA, FK_TIPO_REFERENCIA_PERSONAL, TIPO_REFERENCIA_FECHA, FK_PERSONAL_LEGAJO,"
        Sql = Sql & vbCrLf & " ORDEN_CARGA"
        Sql = Sql & vbCrLf & " From CAJAS "
        Sql = Sql & vbCrLf & "  Where FK_TIPO_REFERENCIA = 1015 "
        Sql = Sql & vbCrLf & "  And FK_PERSONAL_LEGAJO = " & MDIfrmInicio.StaInicio.Panels(2)
        Sql = Sql & vbCrLf & "  ORDER BY ORDEN_CARGA DESC , ORDEN_CARRO "
        rs.Open Sql, strConBasa
        If Not rs.EOF Then
            MsgBox "Usted ya tiene una caja asignada : " & rs!NRO_CAJA
            txtCaja.Enabled = False
            txtCaja.Text = rs!NRO_CAJA
            ctlCliente.Valor = rs!FK_CLIENTE
            
            Exit Sub
        End If
        
            Sql = " SELECT TOP (1) ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_TIPO_REFERENCIA, FK_TIPO_REFERENCIA_PERSONAL, TIPO_REFERENCIA_FECHA, FK_PERSONAL_LEGAJO,"
            Sql = Sql & vbCrLf & " ORDEN_CARGA "
            Sql = Sql & vbCrLf & " From CAJAS "
            Sql = Sql & vbCrLf & " Where FK_TIPO_REFERENCIA = 1015 "
            Sql = Sql & vbCrLf & " AND (NOT (FK_CLIENTE IS NULL))"
            Sql = Sql & " AND  (CAJAS.FK_TIPO_REFERENCIA_PERSONAL IN (19, 69, 17, 47)) "
            Sql = Sql & " AND FK_PERSONAL_LEGAJO is null "
            Sql = Sql & vbCrLf & "  AND (NOT (ORDEN_CARGA IS NULL))"
            'Sql = Sql & " AND FK_CLIENTE = " & InputBox("Ingrese el cliente ")
            'Sql = Sql & " AND NRO_CAJA = " & InputBox("Ingrese la Caja")
            Sql = Sql & "  ORDER BY CAJAS.ORDEN_CARGA desc,FK_TIPO_REFERENCIA DESC "
            Set rs = New ADODB.Recordset
            rs.Open Sql, strConBasa
        If Not rs.EOF Then
            Sql = " Update basasql.dbo.CAJAS"
            Sql = Sql & vbCrLf & " Set FK_PERSONAL_LEGAJO = " & MDIfrmInicio.StaInicio.Panels(2)
            Sql = Sql & vbCrLf & " Where ID_CAJA = " & rs!ID_CAJA
            ExecutarSql Sql
            MsgBox "Su caja asignada es : " & rs!NRO_CAJA
            txtCaja.Enabled = False
            txtCaja.Text = rs!NRO_CAJA
            ctlCliente.Valor = rs!FK_CLIENTE
            
        Else
            MsgBox "Caja no disponible "
        End If
            
    End Sub

Private Sub cmdVideoAtras_Click()
Video.Controls.fastForward
End Sub

Private Sub CMGGR_Click()
Dim i As Integer
 For i = 0 To grdDescripcion.Columns.Count - 1
    
    
    
    Debug.Print " grdDescripcion.Columns.Item(" & i & ").Width= " & CInt(grdDescripcion.Columns.Item(i).Width)
 Next
 
 
End Sub

Private Sub Command10_Click()
MsgBox DateAdd("d", 68651, "28/12/1800")
End Sub

Private Sub Command11_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
        Sql = " SELECT     COD_CLIENTE, NRO_CAJA"
        Sql = Sql & " From dbo.REFERENCIAS"
        Sql = Sql & " GROUP BY COD_CLIENTE, NRO_CAJA"
        Sql = Sql & " ORDER BY COD_CLIENTE, NRO_CAJA"
        rs.Open Sql, ConActiva, 0, 1
        Do While Not rs.EOF
            Sql = "  Update dbo.Cajas"
            Sql = Sql & vbCrLf & "   Set FK_TIPO_REFERENCIA = 1000"
            Sql = Sql & vbCrLf & " Where FK_CLIENTE =  " & rs!COD_CLIENTE
            Sql = Sql & vbCrLf & " And NRO_CAJA = " & rs!NRO_CAJA
            Sql = Sql & vbCrLf & " And (FK_TIPO_REFERENCIA Is Null)"
            ExecutarSql (Sql)
            rs.MoveNext
        Loop
    


End Sub

Private Sub Command5_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String

Sql = " SELECT     COD_CLIENTE, INDICE, ID"
Sql = Sql & "  From INDICES"
Sql = Sql & " WHERE     (TIPO_INDICE = 'Legajo')"
Rem sql = sql & " AND  COD_CLIENTE = 99 "
Sql = Sql & " ORDER BY COD_CLIENTE"

rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF

    Sql = " Update LEGAJOS"
Sql = Sql & "  SET  FK_INDICES =" & rs!ID
Sql = Sql & "  WHERE     COD_INDICE = '" & rs!Indice & "'"
Sql = Sql & "  AND COD_CLIENTE =" & rs!COD_CLIENTE
ExecutarSql Sql
    rs.MoveNext
Loop


End Sub

Private Sub Command6_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Legajo As String

Sql = " SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO,  LETRA_HASTA, LETRA_DESDE, NOMBRE, CONTROL_EXPORT, COD_CLIENTE, CLIENTE_LEGAJO, COD_CLIENTE, NOMBRE, DESCRIPCION, NUMERO_LEGAJO_CLIENTE,  COD_INDICE, NRO_DESDE , NRO_HASTA"
Sql = Sql & " From LEGAJOS "
Sql = Sql & " Where (COD_CLIENTE IN (82)) And   (NRO_DESDE IS NULL) "
Sql = Sql & "  "
Sql = Sql & " ORDER BY ID_LEGAJO "

ConBasa.CommandTimeout = 600000
rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic

Do While Not rs.EOF
rs!CONTROL_EXPORT = "2302200914:33"
If Not IsNull(rs!CLIENTE_LEGAJO) Then

If Mid(rs!CLIENTE_LEGAJO, 3, 1) = "-" Then



Legajo = Replace(rs!CLIENTE_LEGAJO, "_", "")
Legajo = Replace(Legajo, ".", "")
Legajo = Mid(Legajo, 4)
rs!Cod_Indice = "0010010" & Mid(rs!CLIENTE_LEGAJO, 1, 2)

    If IsNumeric(Legajo) Then
            rs!NRO_DESDE = CDbl(Legajo)
            rs!NRO_HASTA = CDbl(Legajo)
          
    End If
    

    If Not IsNull(rs!Nombre) Then
        rs!LETRA_DESDE = Trim(UCase(rs!Nombre))
        rs!LETRA_HASTA = Trim(UCase(rs!Nombre))
    End If
    
    
      rs.Update
 Else
 End If
 End If
rs.MoveNext
Loop

 

End Sub

Private Sub Command7_Click()
Dim Sql As String
Dim NUMERO As Long
Dim lETRA As String
Dim Año As String
Dim fecha  As String
Dim i As Integer
Dim Legajo As String
Dim Pos1 As Integer
Dim Pos2 As Integer
Dim rs As New ADODB.Recordset
ConBasa.CommandTimeout = 300000
Sql = " SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_CLIENTE, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA,NUMERO_LEGAJO_CLIENTE, CLIENTE_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE, ID_LEGAJO, COD_CLIENTE"
Sql = Sql & " From LEGAJOS"
Sql = Sql & "  WHERE     (COD_CLIENTE =84)"

Sql = Sql & " ORDER BY ID_LEGAJO "
rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
On Error GoTo salir
Do While Not rs.EOF
    
    If Not IsNull(rs!CLIENTE_LEGAJO) Then
    lETRA = ""
    NUMERO = 0
    fecha = ""
    Legajo = Replace(Trim(rs!CLIENTE_LEGAJO), "_", "")
     Legajo = Replace(Trim(Legajo), "--", "-?-")
    If Not IsNumeric(Legajo) And Legajo <> "" Then
        
        Pos1 = InStr(1, Legajo, "-")
        
        If Pos1 <> 0 Then
                    If IsNumeric(Mid(Legajo, 1, Pos1 - 1)) Then
                    NUMERO = Mid(Legajo, 1, Pos1 - 1)
                    rs!NRO_DESDE = NUMERO
                     rs!NRO_HASTA = NUMERO
                    End If
                   
                    
                    
                
                
                rs!LETRA_DESDE = Legajo
                rs!LETRA_HASTA = Legajo
                
                
                If fecha <> "" And Len(fecha) = 4 Then
                rs!FECHA_DESDE = "01/01/" & fecha
                rs!FECHA_HASTA = "31/12/" & fecha
                Else
                rs!FECHA_DESDE = Null
                rs!FECHA_HASTA = Null
                End If
    Else
    rs!LETRA_DESDE = Legajo
                rs!LETRA_HASTA = Legajo
    End If
    Else
        
        If Legajo <> "" Then
        rs!NRO_DESDE = Legajo
        rs!NRO_HASTA = Legajo
        End If
    End If
    End If
    rs.Update
    
salir:
    If Err.Number <> 0 Then
    rs.CancelUpdate
    Err.Clear
    End If
    rs.MoveNext
Loop


End Sub

Private Sub Command8_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
Dim NUMERO As Long
Dim lETRA As String
Dim Año As String
Dim i As Integer
Dim Legajo As String
Dim Pos1 As Integer
Dim Pos2 As Integer
ConBasa.CommandTimeout = 300000
Sql = " SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_CLIENTE, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA,NUMERO_LEGAJO_CLIENTE, CLIENTE_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE, ID_LEGAJO, COD_CLIENTE,Nombre, descripcion"
Sql = Sql & " From LEGAJOS"
Sql = Sql & "  WHERE     (COD_CLIENTE =128)"
Sql = Sql & " ORDER BY ID_LEGAJO "
rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
On Error GoTo salir
Do While Not rs.EOF
    
    
    
                If Not IsNull(rs!CLIENTE_LEGAJO) Then
                rs!Descripcion = Trim(rs!Descripcion & "   #" & Trim(rs!CLIENTE_LEGAJO)) & "#"
                
                End If
                
 
    rs.Update
    
salir:
    If Err.Number <> 0 Then
    rs.CancelUpdate
    Err.Clear
    End If
    rs.MoveNext
Loop
End Sub

Private Sub Command9_Click()

'Dim Sql As String
'Dim rs As New ADODB.Recordset
'Dim Legajo As String
'Dim C As Double
'Sql = "  SELECT     CONTENEDOR.ID_CONTENEDOR, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, CONTENEDOR.ESTANTERIA, "
' Sql = Sql & "                     CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, CONTENEDOR.ADELANTE_ATRAS, CONTENEDOR.NRO_ESTANTE, CONTENEDOR.NUEVA,"
' Sql = Sql & "                      CONTENEDOR.Estado"
'Sql = Sql & "  FROM         CONTENEDOR INNER JOIN"
'                      Sql = Sql & "  CAJAS_DUPLICADAS ON CONTENEDOR.COD_CLIENTE = CAJAS_DUPLICADAS.COD_CLIENTE AND"
'                      Sql = Sql & "  CONTENEDOR.NRO_CAJA = CAJAS_DUPLICADAS.NRO_CAJA"
'Sql = Sql & "  Where (CONTENEDOR.Estanteria = 110) And (CONTENEDOR.NUEVA = 1)"
'Sql = Sql & "  ORDER BY CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, CONTENEDOR.NUEVA DESC"
'
'ConBasa.CommandTimeout = 600000
'rs.CursorLocation = adUseClient
'rs.Open Sql,ConActiva, adOpenDynamic, adLockOptimistic
'
'Do While Not rs.EOF
'Sql = " Update CONTENEDOR"
'Sql = Sql & "  SET              COD_CLIENTE = NULL, NRO_CAJA = NULL"
'Sql = Sql & "  WHERE     ID_CONTENEDOR = " & rs!ID_CONTENEDOR
' ExecutarSql Sql
'rs.MoveNext
'Loop


    Dim Sql As String
    Dim rs As New ADODB.Recordset
    
    Sql = "  SELECT     FK_DOCUMENTOS_DIGITALES_LOTE, COUNT(*) AS CANTIDAD"
    Sql = Sql & "  From DOCUMENTOS_DIGITALES"
    Sql = Sql & "  GROUP BY FK_DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & "  Having (FK_DOCUMENTOS_DIGITALES_LOTE > 1)"
    Sql = Sql & "  ORDER BY FK_DOCUMENTOS_DIGITALES_LOTE"


rs.Open Sql, ConActiva, 0, 1
Do While Not rs.EOF
    Sql = " Update DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & "  SET  CANTIDAD_ARCHIVOS =" & rs!cantidad
    Sql = Sql & "  Where (Cantidad_Archivos Is Null) "
    Sql = Sql & "  And ID_DOCUMENTOS_DIGITALES_LOTE =" & rs!FK_DOCUMENTOS_DIGITALES_LOTE
    ExecutarSql Sql
    rs.MoveNext
Loop




End Sub

Private Sub ctlCliente_Click()
If txtCaja.Enabled = False Then
Else

    LimpiarCampos True
    End If
    
      txtCaja.Text = ""

End Sub

Private Sub ctlCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub Form_Activate()
On Error GoTo salir:
Timer1.Enabled = False
If MDIfrmInicio.StaInicio.Panels(2) = 29 Or MDIfrmInicio.StaInicio.Panels(2) = 23 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 37 Or MDIfrmInicio.StaInicio.Panels(2) = 35 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 27 Or MDIfrmInicio.StaInicio.Panels(2) = 49 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 82 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 46 Or MDIfrmInicio.StaInicio.Panels(2) = 31 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 47 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 70 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 48 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 19 _
        Then
        
        cmdCajasConLegajos.Visible = True
        cmdsolicituddecaja.Visible = False
        
        txtCaja.Enabled = True
    chkFijarCaja.value = 0
    
    Else
        cmdsolicituddecaja.Visible = True
    End If

If NºDOCUMENTO <> 0 Then
    frmAgregarDocumentos.txtIndice_Nro_Documento = NºDOCUMENTO
    txtIndice_Nro_Documento.SetFocus

End If
Rem txtCaja.Text = ""
Rem txtCaja.Enabled = False
txtIndice_Nro_Documento.SetFocus
chkFijarTipoDocumento.value = 0
frmAgregarDocumentos.Top = 0
frmAgregarDocumentos.Left = 0
Rem frmAgregarDocumentos.Width = MDIfrmInicio.Width - 300
Rem frmAgregarDocumentos.Height = MDIfrmInicio.Height - 600
frmAgregarDocumentos.WindowState = 2
SSTab1.Tab = 0
Rem cmdCargaCompletaCajaLegajo.Enabled = False
If MDIfrmInicio.StaInicio.Panels(2) = 55 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 29 Or MDIfrmInicio.StaInicio.Panels(2) = 23 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 37 Or MDIfrmInicio.StaInicio.Panels(2) = 35 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 27 Or MDIfrmInicio.StaInicio.Panels(2) = 49 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 82 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 46 Or MDIfrmInicio.StaInicio.Panels(2) = 31 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 70 Then
        cmdCajasConLegajos.Visible = True
        txtCaja.Enabled = True
    chkFijarCaja.value = 0
    
    Else
        cmdsolicituddecaja.Visible = True
    End If

salir:
End Sub

Public Sub ExportarExcelReferencia(Filtro As String, DocSolo As Boolean, PorCaja As Boolean, PorIndice As Boolean)
   
   
    Dim Sql As String
    Dim SqlIndice As String
    Dim rsbasa As New ADODB.Recordset
    Dim PasoPlanilla As String
    
   
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim ExcelIndice As Excel.Worksheet
    Dim ExcelReferenciaIndice As Excel.Worksheet
    Dim ExcelReferenciaCaja As Excel.Worksheet
    
    
    Dim SqlBase As String
    SqlBase = " SELECT  COD_ID_REFERENCIA, INDICES.TITULOHERENCIA,INDICES.DESCRIPCION AS DESCRIPCIONINDICE , REFERENCIAS.NRO_CAJA, REFERENCIAS.ITEM,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.INDICE, REFERENCIAS.DESCRIPCION,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.COD_CLIENTE,INDICES.ID_CODIGO_DOCUMENTO,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.LETRA_DESDE, REFERENCIAS.LETRA_HASTA,"
    SqlBase = SqlBase & vbCrLf & " REFERENCIAS.EXPEDIENTE, REFERENCIAS.APELLIDO_NOMBRE, REFERENCIAS.BORRADO"
    SqlBase = SqlBase & vbCrLf & " From REFERENCIAS, INDICES"


   On Error GoTo er
       
       'Plantilla Base
       MousePointer = 11
        If Dir("C:\Referencias", vbDirectory) = "" Then
    FileSystem.MkDir "C:\Referencias"
    
    End If
       PasoPlanilla = "C:\Referencias\referencias " & ctlCliente.Valor & Format(Now, "ddmmyyyy") & ".xls"
      
       FileCopy strPasoPlanillas & "\" & "Referencia Envio.xls", PasoPlanilla
    
    'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open(PasoPlanilla)
        Set ExcelIndice = libroEx.Worksheets.Item(1)
        Set ExcelReferenciaIndice = libroEx.Worksheets.Item(2)
        Set ExcelReferenciaCaja = libroEx.Worksheets.Item(3)
    
    Dim i As Integer
    For i = 1 To 10
    Rem MsgBox ExcelIndice.Cells(2, I).NumberFormat
    Next
    
    'Creacion del Indice
            Sql = " SELECT * "
            Sql = Sql & vbCrLf & "  From INDICES "
            Sql = Sql & vbCrLf & "  WHERE COD_CLIENTE = " & ctlCliente.Valor
            Sql = Sql & vbCrLf & Filtro
            Sql = Sql & vbCrLf & "  ORDER BY INDICE"
            Set rsbasa = New ADODB.Recordset
            rsbasa.Open Sql, ConActiva, 0, 1
            ExcelIndice.Cells(1, 2) = " DICCIONARIO DE DOCUMENTOS "
            ExcelIndice.Cells(2, 2) = ctlCliente.Descripcion
            Rem  Apellido
            IndiceExcel rsbasa, ExcelIndice
            
            If DocSolo = True Then
            
            libroEx.Worksheets.Item(2).Delete
                ExcelReferenciaIndice.Delete
                libroEx.Worksheets.Item(3).Delete
                ExcelReferenciaCaja.Delete
                libroEx.Save
            End If
    
   If DocSolo = False Then
        If PorIndice = True Then
        ' Creacion de referencia por Indice
                Sql = SqlBase & vbCrLf & " WHERE REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND"
                Sql = Sql & vbCrLf & " REFERENCIAS.INDICE = INDICES.INDICE "
                Sql = Sql & vbCrLf & " AND REFERENCIAS.COD_CLIENTE =" & ctlCliente.Valor
                Sql = Sql & vbCrLf & Filtro
                Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.INDICE, REFERENCIAS.FECHA_DESDE,REFERENCIAS.NRO_DESDE"
                Set rsbasa = New ADODB.Recordset
                rsbasa.Open Sql, ConActiva, 0, 1
                 ExcelReferenciaIndice.Name = "Ref Indice"
                 ProcesarPorIndices rsbasa, ExcelReferenciaIndice
            End If
     End If
     
    If DocSolo = False Then
        If PorCaja = True Then
   'Referencia por Caja
            Sql = SqlBase & vbCrLf & " WHERE REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND"
            Sql = Sql & vbCrLf & " REFERENCIAS.INDICE = INDICES.INDICE "
            Sql = Sql & vbCrLf & " AND REFERENCIAS.COD_CLIENTE =" & ctlCliente.Valor
            Sql = Sql & vbCrLf & Filtro
            Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.NRO_CAJA, REFERENCIAS.FECHA_DESDE,REFERENCIAS.NRO_DESDE "
            Set rsbasa = New ADODB.Recordset
            rsbasa.Open Sql, ConActiva, adOpenDynamic
            ProcesarPorCajas rsbasa, ExcelReferenciaCaja
         End If
     End If

    rsbasa.Close
    
   
    libroEx.Save
    
    
    libroEx.Close
    ApExcel.Quit
    Set ApExcel = Nothing
    Set libroEx = Nothing
    Set ExcelIndice = Nothing
    Set ExcelReferenciaIndice = Nothing
    Set ExcelReferenciaCaja = Nothing
    
MousePointer = 0
MsgBox "Terminado"
Exit Sub
er:
    Set ApExcel = Nothing
    Set libroEx = Nothing
    Set ExcelIndice = Nothing
    Set ExcelReferenciaIndice = Nothing
    Set ExcelReferenciaCaja = Nothing
MsgBox Err.Description

End Sub

Private Sub Form_GotFocus()
frmAgregarDocumentos.WindowState = 2

End Sub

Public Sub GrabarReferencias(StringIndice As String)
    Dim COD_CLIENTE, NRO_CAJA, Item, Indice, Descripcion     As String
    Dim FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA As String
    Dim LETRA_DESDE, LETRA_HASTA, EXPEDIENTE, APELLIDO_NOMBRE As String
    Dim FECHA_MODIFICACION, USUARIO_MODIFICACION, PLANILLA As String
    Dim ID_imagen As String
    Dim ID_UNITER As Long
    Dim sSQL As String

On Error GoTo salir:



        If IsNull(ctlCliente.Valor) Then
            MsgBox "Usted Debe ingresar el Cliente"
            Exit Sub
        Else
            COD_CLIENTE = ctlCliente.Valor
        End If

        If txtLoteReferencia.Text <> "" Then
            PLANILLA = "'" & txtLoteReferencia.Text & "'"
        Else
            PLANILLA = "'0'"
        End If



        If Not IsNumeric(txtCaja.Text) Then
            MsgBox "Usted Debe ingresar la Caja"
            Exit Sub
        Else
            NRO_CAJA = txtCaja.Text
        End If

        If StringIndice = "" Then
            MsgBox "Error indice"
            Exit Sub
         Else
            Indice = StringIndice
        End If

        If txtDescripcion.Text = "" Then
          Descripcion = "NULL"
        Else
           Descripcion = "'" & Replace(UCase(Trim(Replace(txtDescripcion.Text, vbCrLf, " "))), vbCrLf, " ") & "'"
        End If

        If txtFechaDesde.Text = "" Then
            FECHA_DESDE = "NULL"
        Else
            FECHA_DESDE = FechaFormato(txtFechaDesde.Text)
        End If

        If txtFechaHasta.Text = "" Then
            FECHA_HASTA = "NULL"
        Else
            FECHA_HASTA = FechaFormato(txtFechaHasta.Text)
        End If

        If Not IsNumeric(txtNroDesde.Text) Then
            NRO_DESDE = "NULL"
        Else
            NRO_DESDE = txtNroDesde.Text
        End If

        If Not IsNumeric(txtNroHasta.Text) Then
            NRO_HASTA = "NULL"
        Else
            NRO_HASTA = txtNroHasta.Text
        End If

        If txtLetraDesde.Text = "" Then
            LETRA_DESDE = "NULL"
        Else
            LETRA_DESDE = "'" & UCase(Trim(txtLetraDesde.Text)) & "'"
        End If

        If txtLetraHasta.Text = "" Then
            LETRA_HASTA = "NULL"
        Else
            LETRA_HASTA = "'" & UCase(Trim(txtLetraHasta.Text)) & "'"
        End If

        FECHA_MODIFICACION = SysDateMinutoSegundo

         ID_UNITER = 0


        If txtUsuarioCarga.Text <> "" Then
            Usuario = txtUsuarioCarga.Text
        Else
            MsgBox "Ingrese quien carga", vbInformation
            Exit Sub
        End If

'        If lbl_ID_imagen.Caption = "" Then
'            ID_imagen = "Null"
'        Else
'            ID_imagen = lbl_ID_imagen.Caption
'             InsertarImagenes CLng(ID_imagen), ctlCliente.Valor, CLng(NRO_CAJA), 1, SysDate2
'        End If

Select Case StatusBar.Panels(1).Text
Case "Nuevo"

        sSQL = "    INSERT INTO REFERENCIAS"
        sSQL = sSQL & vbCrLf & "        ( COD_CLIENTE, NRO_CAJA, ITEM, INDICE, DESCRIPCION,"
        sSQL = sSQL & vbCrLf & "        FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA,"
        sSQL = sSQL & vbCrLf & "        LETRA_DESDE, LETRA_HASTA, "
        sSQL = sSQL & vbCrLf & "         FECHA_MODIFICACION,"
        sSQL = sSQL & vbCrLf & "        FK_PERSONAL_CREACION, FK_PERSONAL_MODIFICACION,borrado, ESTADO ,ID_IMAGEN , PASOARCHIVO)"
        sSQL = sSQL & vbCrLf & "    Values"
        sSQL = sSQL & vbCrLf & "  (" & COD_CLIENTE & "," & NRO_CAJA & ",0," & Indice & "," & Descripcion & ","
        sSQL = sSQL & vbCrLf & FECHA_DESDE & "," & FECHA_HASTA & "," & NRO_DESDE & "," & NRO_HASTA & ","
        sSQL = sSQL & vbCrLf & LETRA_DESDE & "," & LETRA_HASTA & ","
        sSQL = sSQL & vbCrLf & FECHA_MODIFICACION & ","
        sSQL = sSQL & vbCrLf & Usuario & "," & Usuario & ", 0,2," & 0 & "," & PLANILLA & ")"
        ExecutarSql (sSQL)

          sSQL = "  Update dbo.Cajas"
          sSQL = sSQL & vbCrLf & "   Set FK_TIPO_REFERENCIA = 1000"
          sSQL = sSQL & vbCrLf & " Where FK_CLIENTE =  " & COD_CLIENTE
          sSQL = sSQL & vbCrLf & " And NRO_CAJA = " & NRO_CAJA
          sSQL = sSQL & vbCrLf & " And (FK_TIPO_REFERENCIA Is Null)"
          ExecutarSql (sSQL)

               LimpiarCampos True
Case "Modificar"

If IsNumeric(StatusBar.Panels(2).Text) Then
    sSQL = "    Update REFERENCIAS"
    sSQL = sSQL & vbCrLf & "   SET "
    sSQL = sSQL & vbCrLf & "  COD_CLIENTE = " & COD_CLIENTE
    sSQL = sSQL & vbCrLf & " , NRO_CAJA = " & NRO_CAJA
    sSQL = sSQL & vbCrLf & " , INDICE = " & Indice
    sSQL = sSQL & vbCrLf & " , DESCRIPCION = " & Descripcion
    sSQL = sSQL & vbCrLf & " , FECHA_DESDE =" & FECHA_DESDE
    sSQL = sSQL & vbCrLf & " , FECHA_HASTA =" & FECHA_HASTA
    sSQL = sSQL & vbCrLf & " , NRO_DESDE =" & NRO_DESDE
    sSQL = sSQL & vbCrLf & " , NRO_HASTA =" & NRO_HASTA
    sSQL = sSQL & vbCrLf & " , LETRA_DESDE =" & LETRA_DESDE
    sSQL = sSQL & vbCrLf & " , LETRA_HASTA =" & LETRA_HASTA
    sSQL = sSQL & vbCrLf & " , FECHA_MODIFICACION=" & FECHA_MODIFICACION
    sSQL = sSQL & vbCrLf & " , USUARIO_MODIFICACION =" & Usuario
    sSQL = sSQL & vbCrLf & " Where COD_ID_REFERENCIA = " & StatusBar.Panels(2).Text
    ExecutarSql sSQL
     MsgBox "Se realizo la actualizacion"
Else
    MsgBox "Error en la actualizacion" & vbCrLf & "verifique si el estado de la aplicacion es modicacion", vbInformation
End If
End Select


    txtCaja.SetFocus
Exit Sub
salir:

MsgBox Err.Description

End Sub


Public Sub BuscarEcogas(Pig As Long)
Dim rs As New ADODB.Recordset
Dim Sql As String
Dim CONlEGAJOS As ADODB.Connection

Sql = " SELECT num "
Sql = Sql & vbCrLf & " , [calle] & '  ' & [inmueble_puerta_num] & '  ' & "
Sql = Sql & vbCrLf & " [localidad_nombre] & ' ' & Pig.pcia_nombre AS descripcion, "
Sql = Sql & vbCrLf & " 'Bº ' & [barrio] & '  ' & [inmueble_torre_des] & '   ' & "
Sql = Sql & vbCrLf & " [inmueble_dpto_des] & '  ' & [inmueble_piso_des] AS Nombre"
Sql = Sql & vbCrLf & " From Pig  "
Sql = Sql & vbCrLf & " WHERE num = " & txtNroDesde.Text

    
     
     
            Set CONlEGAJOS = New ADODB.Connection
            CONlEGAJOS.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ClienteEcogas
    
    rs.Open Sql, CONlEGAJOS
    If Not rs.EOF Then
        If Trim(txtLetraDesde.Text) = "" Then
           txtLetraDesde.Text = Trim(rs!Descripcion)
           txtLetraHasta.Text = Trim(rs!Nombre)
        End If
        
    Else
        MsgBox "NO se encontro el pig"
    End If
End Sub




Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 42 Then ' Video
            KeyAscii = 0
            If cmdPlay_Pausa.Caption = "Play" Then
                Play_Pausa "Play"
                Exit Sub
            End If
            If cmdPlay_Pausa.Caption = "Pausa" Then
                Play_Pausa "Pausa"
            End If
 End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Debug.Print KeyCode

If KeyCode = 122 Then
'    If chkNoBuscar.value = 1 Then
'            chkNoBuscar.value = 0
'    Else
'        chkNoBuscar.value = 1
'    End If
'


End If

If KeyCode = 112 Then ' CAJA
        If chkFijarCaja.value = 1 Then
           chkFijarCaja.value = 0
           txtCaja.SetFocus
        Else
           chkFijarCaja.value = 1
        End If
    End If

    If KeyCode = 113 Then ' Tipo Documento
        If chkFijarTipoDocumento.value = 1 Then
           chkFijarTipoDocumento.value = 0
           txtIndice_Nro_Documento.SetFocus
        Else
           chkFijarTipoDocumento.value = 1
        End If
    End If

    If KeyCode = 114 Then ' Fecha Desde
        If chkFijarFechaDesde.value = 1 Then
           chkFijarFechaDesde.value = 0
           txtFechaDesde.SetFocus
        Else
           chkFijarFechaDesde.value = 1
        End If
    End If
    
    If KeyCode = 115 Then ' Fecha Hasta
        If chkFijarFechaHasta.value = 1 Then
           chkFijarFechaHasta.value = 0
           txtFechaHasta.SetFocus
        Else
           chkFijarFechaHasta.value = 1
        End If
    End If
    
    If KeyCode = 116 Then ' Nro Desde
        If chkFijarNumeroDesde.value = 1 Then
           chkFijarNumeroDesde.value = 0
           txtNroDesde.SetFocus
        Else
           chkFijarNumeroDesde.value = 1
        End If
    End If
    
    If KeyCode = 117 Then ' Nro Desde
       If chkFijarNumeroHasta.value = 1 Then
           chkFijarNumeroHasta.value = 0
           txtNroHasta.SetFocus
        Else
           chkFijarNumeroHasta.value = 1
        End If
    End If
    
    If KeyCode = 118 Then ' Letra Desde
       If chkFijarLetraDesde.value = 1 Then
           chkFijarLetraDesde.value = 0
           txtLetraDesde.SetFocus
        Else
           chkFijarLetraDesde.value = 1
        End If
    End If
    
    
    If KeyCode = 119 Then ' Letra Hasta
       If chkFijarLetraHasta.value = 1 Then
           chkFijarLetraHasta.value = 0
           txtLetraHasta.SetFocus
        Else
           chkFijarLetraHasta.value = 1
        End If
    End If
        If KeyCode = 120 Then ' Letra Hasta
       If chkFijarDescripcion.value = 1 Then
           chkFijarDescripcion.value = 0
           txtDescripcion.SetFocus
        Else
           chkFijarDescripcion.value = 1
        End If
    End If
    If KeyCode = 123 Then ' Buscar Indice
        cmdBuscarDocumento_Click
    End If
        
       
     
    

End Sub

Private Sub Form_Load()
On Error GoTo salir:

fraReferencias.Left = 120
fraReferencias.Top = 600
fraLegajo.Left = 120
fraLegajo.Top = 600

frmAgregarDocumentos.Top = 0
frmAgregarDocumentos.Left = 0
frmAgregarDocumentos.Width = MDIfrmInicio.Width - 300
frmAgregarDocumentos.Height = MDIfrmInicio.Height - 600
salir:

ctlCliente.TipoControl = Cliente
fraLegajo.Visible = False
fraCajaDocumento.Visible = False
fraRearchivo.Visible = False
grdDatos.Visible = False
fraCampos.Visible = False
fraReferencias.Visible = False
NºDOCUMENTO = 0
SSTab1.Visible = False
txtUsuarioCarga.Text = MDIfrmInicio.StaInicio.Panels(2).Text
cmdCajasConLegajos.Visible = False
cmdsolicituddecaja.Visible = False
Rem Or MDIfrmInicio.StaInicio.Panels(2) = 17

    If MDIfrmInicio.StaInicio.Panels(2) = 55 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 29 Or MDIfrmInicio.StaInicio.Panels(2) = 23 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 37 Or MDIfrmInicio.StaInicio.Panels(2) = 35 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 27 Or MDIfrmInicio.StaInicio.Panels(2) = 49 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 82 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 46 Or MDIfrmInicio.StaInicio.Panels(2) = 31 Or _
        MDIfrmInicio.StaInicio.Panels(2) = 70 Then
        cmdCajasConLegajos.Visible = True
        txtCaja.Enabled = True
    
    Else
        cmdsolicituddecaja.Visible = True
    End If
     



SqlReferencia = "  SELECT Top 1000 REFERENCIAS.COD_CLIENTE AS CLIENTE, REFERENCIAS.NRO_CAJA AS CAJA, INDICES.DESCRIPCION AS INDICE, REFERENCIAS.DESCRIPCION,"
SqlReferencia = SqlReferencia & " REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA, REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA,"
SqlReferencia = SqlReferencia & " REFERENCIAS.LETRA_DESDE, REFERENCIAS.LETRA_HASTA, REFERENCIAS.COD_ID_REFERENCIA AS CODIGO,"
SqlReferencia = SqlReferencia & " REFERENCIAS.FECHA_MODIFICACION AS FECHA, REFERENCIAS.FK_PERSONAL_MODIFICACION AS PERSONAL, REFERENCIAS.ID_IMAGEN ,  COD_ID_REFERENCIA AS CODIGO "
SqlReferencia = SqlReferencia & " FROM REFERENCIAS INNER JOIN"
SqlReferencia = SqlReferencia & " INDICES ON REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND REFERENCIAS.INDICE = INDICES.INDICE"
txtCaja.Text = ""


End Sub

Private Sub Form_Resize()
On Error GoTo salir


SSTab1.Width = frmAgregarDocumentos.Width - 300
grdDatos.Width = SSTab1.Width - 500
Rem ViewImg1.Width = SSTab1.Width - 200
Rem   grdDescripcion.Width = ViewImg1.Width
Rem grdDatos.Width = ViewImg1.Width



SSTab1.Height = frmAgregarDocumentos.Height - SSTab1.Top - 900
Rem ViewImg1.Height = SSTab1.Height - 400
Rem grdDescripcion.Height = ViewImg1.Height
grdDatos.Height = grdDescripcion.Height - fraBotones.Height - 300
fraBotones.Top = grdDatos.Height + 500
ctlVerImagenes1.Width = SSTab1.Width - 400
Video.Top = 200
Video.Height = SSTab1.Height - 400
salir:

End Sub

Private Sub grdDatos_DblClick()
    On Error GoTo salir:
   If cboTipoCarga = "Legajos" Then
    grdDatos.Col = 0
    txtEtiqueta.Text = grdDatos.Text
    LimpiarLegajos
    StatusBar.Panels(1).Text = "Modificar"
    cmdModificar_Click
   End If
   
   If cboTipoCarga = "Referencia" Then
    Dim Cliente As Integer
    Dim ID_referencia As Long
    grdDatos.Col = 0
    Cliente = grdDatos.Text
    
    grdDatos.Col = 14
    ID_referencia = grdDatos.Text
    CargarReferencias Cliente, ID_referencia, "Modificar"
    
    
   End If
   
   Exit Sub
   
   
salir:
   
   MsgBox "Error "
End Sub

Private Sub grdDescripcion_DblClick()

'Dim Descripcion As String
'Dim Indice As String
'grdDescripcion.Col = 0
'Descripcion = grdDescripcion.Text
'
'grdDescripcion.Col = 3
'Indice = grdDescripcion.Text
'
'Dim rs As New ADODB.Recordset
'Dim sql As String
'
'sql = " SELECT     COD_ID_REFERENCIA ,Descripcion"
'sql = sql & " From dbo.REFERENCIAS"
'sql = sql & " Where COD_CLIENTE = " & ctlCliente.Valor
'sql = sql & " AND DESCRIPCION = '" & Descripcion & "'"
'sql = sql & " AND INDICE = '" & Indice & "'"
'rs.Open sql, strConBasa , 0 ,1


Rem Dim B As Bookmark

Dim C As Column
chkFijarDescripcion.value = 1
txtDescripcion.Text = grdDescripcion.Columns("Descripcion").Text
End Sub

Private Sub mnuFiltroUsuarioDia_Click()

 If mnuFiltroUsuarioDia.Checked = True Then
    mnuFiltroUsuarioDia.Checked = False
 Else
    mnuFiltroUsuarioDia.Checked = True
 End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_LostFocus()

End Sub

Private Sub Timer1_Timer()
    txtVideoLugar.Text = Video.Controls.currentPosition
    If ValorAnteVideo = txtVideoLugar.Text Then
        Timer1.Enabled = False
        cmdPlay_Pausa.Caption = "Espera"
    Else
        ValorAnteVideo = txtVideoLugar.Text
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim Sql As String
Select Case Button.Key
    
   Case "Nuevo"
        StatusBar.Panels(1).Text = "Nuevo"
        StatusBar.Panels(2).Text = ""
        LimpiarCampos True
    Case "Borrar"
        If StatusBar.Panels(1).Text = "Modificar" Then
            If IsNumeric(StatusBar.Panels(2).Text) Then
                If MsgBox("Esta Usted Seguro de Borrar el registro", vbInformation + vbYesNo) = vbYes Then
                    Sql = "DELETE FROM dbo.REFERENCIAS "
                    Sql = Sql & " Where COD_ID_REFERENCIA = " & StatusBar.Panels(2).Text
                    ExecutarSql Sql
                    MsgBox "El regitro fue borrado", vbInformation
                    StatusBar.Panels(1).Text = "Nuevo"
                    StatusBar.Panels(2).Text = ""
                    LimpiarCampos False
                End If
            End If
        End If
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

Dim Filtro As String

If txtIndice_Nro_Documento.Text = "" Then
    Filtro = ""
Else
    Filtro = " AND INDICES.INDICE like '" & BuscarIDDocumento(txtIndice_Nro_Documento.Text, ctlCliente.Valor) & "%'"
End If




Select Case ButtonMenu.Key
Case "PorIndice"
    ExportarExcelReferencia Filtro, False, False, True
Case "PorIndiceyCaja"
    ExportarExcelReferencia Filtro, False, True, True

Case "PorCaja"

    ExportarExcelReferencia Filtro, False, True, False
Case "Diccionario"
    ExportarExcelReferencia Filtro, True, False, False
End Select


End Sub


Private Sub txtCaja_GotFocus()
    If chkFijarCaja.value = 1 Then
        SendKeys vbTab
    End If
    
     cmdCargaCompletaCajaLegajo.Enabled = True
    
    
    

End Sub

Private Sub txtCaja_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 16 And Shift = 1 Then
   Buscar_Inidice_Por_caja
End If
End Sub

Private Sub txtCaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If



End Sub

Private Sub txtCaja_LostFocus()

On Error GoTo salir:
  
    Dim rs As New ADODB.Recordset
            Dim Sql As String
    If txtCaja.Text <> "" And Not IsNull(ctlCliente.Valor) Then
            Sql = " SELECT  NRO_CAJA, FK_CLIENTE, FK_TIPO_REFERENCIA "
            Sql = Sql & " From dbo.Cajas "
            Sql = Sql & " Where FK_CLIENTE = " & ctlCliente.Valor
            Sql = Sql & " And NRO_CAJA =  " & txtCaja.Text
            Set rs = New ADODB.Recordset
            rs.Open Sql, ConActiva, 0, 1
            If rs.EOF Then
                MsgBox " Atencion el cliente no tiene esta caja NO se puede cargar ", vbCritical
                txtCaja.Text = ""
                txtCaja.SetFocus
                Exit Sub
            End If


            Sql = " SELECT  ESTADO "
            Sql = Sql & " From CONTENEDOR "
            Sql = Sql & " Where COD_CLIENTE = " & ctlCliente.Valor
            Sql = Sql & " And NRO_CAJA =" & txtCaja.Text
            Sql = Sql & " and  estado in( 2 , 3 ) "


              Set rs = New ADODB.Recordset
              rs.Open Sql, strConBasa, 0, 1
            If rs.EOF Then
                MsgBox "Atencion la caja no tiene Guarda y custodia", vbCritical
                txtCaja.Text = ""
                txtCaja.SetFocus
                Exit Sub
            End If


         End If
  
  
  
  
  
 If cboTipoCarga.Text = "Legajos" Then
        If txtCaja.Text <> "" And Not IsNull(ctlCliente.Valor) Then
            Sql = " SELECT     NRO_CAJA, FK_CLIENTE, FK_TIPO_REFERENCIA"
            Sql = Sql & " From dbo.Cajas"
            Sql = Sql & " Where FK_CLIENTE = " & ctlCliente.Valor
            Sql = Sql & " And NRO_CAJA =  " & txtCaja.Text
            Sql = Sql & " And FK_TIPO_REFERENCIA IN (1015 )"
            Set rs = New ADODB.Recordset
            rs.Open Sql, ConActiva, 0, 1
            If rs.EOF Then
                MsgBox "Atencion el cliente no tiene marcada esta caja como legajos", vbCritical
                Rem txtCaja.Text = ""
                Exit Sub
            End If
         End If
    End If
    
 If cboTipoCarga.Text = "Referencia" Then
        If txtCaja.Text <> "" And Not IsNull(ctlCliente.Valor) Then
            Sql = " SELECT     NRO_CAJA, FK_CLIENTE, FK_TIPO_REFERENCIA"
            Sql = Sql & " From dbo.Cajas"
            Sql = Sql & " Where FK_CLIENTE = " & ctlCliente.Valor
            Sql = Sql & " And NRO_CAJA =  " & txtCaja.Text
             Set rs = New ADODB.Recordset
            rs.Open Sql, ConActiva, 0, 1
            If rs.EOF Then
                MsgBox "Atencion el cliente no tiene esta caja", vbCritical
                txtCaja.Text = ""
                txtCaja.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Ingrese la Caja", vbInformation
        End If
    End If
    Exit Sub
    
salir:
    MsgBox Err.Description
    
End Sub

Private Sub txtCajaDigitoVerificador_Validate(Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    
    If cboTipoCarga.Text = "Legajos" Then
        If MDIfrmInicio.StaInicio.Panels(2) = 31 Or _
            MDIfrmInicio.StaInicio.Panels(2) = 82 Or MDIfrmInicio.StaInicio.Panels(2) = 47 Or _
            MDIfrmInicio.StaInicio.Panels(2) = 49 Or MDIfrmInicio.StaInicio.Panels(2) = 46 Or MDIfrmInicio.StaInicio.Panels(2) = 48 _
            Then
        
        Else
            If IsNumeric(txtCajaDigitoVerificador.Text) Then
                    Sql = "SELECT     TOP (1) ID_CAJA, FK_CLIENTE, NRO_CAJA, DIGITO_VERIFICADOR"
                    Sql = Sql & " From CAJAS"
                    Sql = Sql & " Where FK_CLIENTE =" & ctlCliente.Valor
                    Sql = Sql & " And NRO_CAJA = " & txtCaja.Text
                    Sql = Sql & " And Digito_Verificador = " & txtCajaDigitoVerificador.Text
                    rs.Open Sql, strConBasa
                    If rs.EOF Then
                        MsgBox "INGRESE EL NUMERO VERIFICADOR CORRECTO", vbCritical
                        Cancel = True
                    End If
                Else
                    MsgBox "INGRESE EL NUMERO VERIFICADOR"
                    Cancel = True
                End If
         End If
    End If
       

End Sub

Private Sub txtDescripcion_Change()
'If Len(txtDescripcion.Text) > 4 And chkFijarDescripcion.value = 0 And Mid(txtDescripcion.Text, 1, 1) <> " " And chkNoBuscar.value = 0 Then
' BuscarDescripcion
'End If

End Sub

Private Sub txtDescripcion_GotFocus()
If chkFijarDescripcion.value = 1 Then
        SendKeys vbTab
    End If

End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
If KeyAscii = 45 Or KeyAscii = 42 Then
    If KeyAscii = 42 Then
     txtIndice_Nro_Documento.Text = ""
     KeyAscii = 0
    End If
    BuscarDescripcion
End If

End Sub

Private Sub txtEtiqueta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEtiquetaDigitoVerificador.SetFocus
    Else
'        LimpiarLegajos
'
'        txtCaja.Text = ""
'        txtIndice_Nro_Documento.Text = ""
'        lblIndice_Descripcion.Caption = ""
    
    End If
End Sub

Private Sub txtEtiqueta_LostFocus()
    If txtEtiqueta.Text <> "" Then
        If txtEtiqueta.Text < 750000 Then
          
        Else
        
        End If
    End If
    
    If StatusBar.Panels(1).Text = "Modificar" Then
    
        LimpiarLegajos
    End If
    
    
    
    
End Sub

Private Sub txtEtiquetaDigitoVerificador_DblClick()

Dim Sql As String
Dim rs As New ADODB.Recordset
If ctlCliente.Valor = 279 Then
    If txtEtiqueta.Text <> "" Then
        Sql = " SELECT     DIGITO_VERIFICADOR "
        Sql = Sql & " From LEGAJOS "
        Sql = Sql & " Where ID_LEGAJO =" & txtEtiqueta
        rs.Open Sql, strConBasa
        If Not rs.EOF Then
            txtEtiquetaDigitoVerificador.Text = rs!Digito_Verificador
        
        End If
        
    End If
    
    
    
End If

End Sub

Private Sub txtEtiquetaDigitoVerificador_KeyPress(KeyAscii As Integer)
On Error GoTo salir
 If KeyAscii = 13 Then
 
 If txtEtiqueta.Text <> "" Then
 
        If Digito_Verificador(txtEtiqueta.Text) = txtEtiquetaDigitoVerificador.Text Then
            Rem txtCaja.SetFocus
        Else
            MsgBox "Error etiqueta"
        End If
        
 End If
 
 End If
Exit Sub
salir:
 MsgBox Err.Description
End Sub


Private Sub txtFechaDesde_GotFocus()
If chkFijarFechaDesde.value = 1 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtFechaDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If

 If KeyAscii = 43 Then
    KeyAscii = 0
    txtFechaHasta.Text = txtFechaDesde.Text
    SendKeys vbTab
End If

End Sub

Private Sub txtFechaDesde_LostFocus()
Dim Año As String
If Len(Trim(txtFechaDesde.Text)) = 5 Then
    AgregarQuincena Trim(txtFechaDesde.Text)
    Exit Sub
   End If
   
If Len(txtFechaDesde.Text) = 6 Then
    AgregarMes (txtFechaDesde.Text)
    Exit Sub
End If

If Len(txtFechaDesde.Text) = 4 Then
    Año = txtFechaDesde.Text

    txtFechaDesde.Text = "01/01/" & Año
    txtFechaHasta.Text = "31/12/" & Año
Else
   
  If Trim(txtFechaDesde.Text) <> "" Then
        If InStr(1, txtFechaDesde.Text, "/") = 0 Then
             txtFechaDesde.Text = Mid(txtFechaDesde.Text, 1, 2) & "/" & Mid(txtFechaDesde.Text, 3, 2) & "/" & Mid(txtFechaDesde.Text, 5)
        End If
        
        If chk_Copiar_Fecha.value = 1 Then
             txtFechaHasta.Text = txtFechaDesde.Text
        End If
   End If

End If




End Sub

Private Sub txtFechaHasta_GotFocus()
If chkFijarFechaHasta.value = 1 Then
    SendKeys vbTab
Else
    If txtFechaHasta.Text <> "" Then
        txtFechaHasta.SelStart = 2
    
    End If
    
End If

End Sub

Private Sub txtFechaHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub txtFechaHasta_LostFocus()
   If Trim(txtFechaHasta.Text) <> "" Then
    
    If InStr(1, txtFechaHasta.Text, "/") = 0 Then
        txtFechaHasta.Text = Mid(txtFechaHasta.Text, 1, 2) & "/" & Mid(txtFechaHasta.Text, 3, 2) & "/" & Mid(txtFechaHasta.Text, 5)
    End If
    End If
End Sub

Private Sub txtID_Referencia_Change()
txtID_Referencia.Text = Replace(txtID_Referencia.Text, vbCrLf, " ")
End Sub

Private Sub txtIndice_Nro_Documento_GotFocus()
    If chkFijarTipoDocumento.value = 1 Then
        SendKeys vbTab
    End If

End Sub

Private Sub txtIndice_Nro_Documento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    On Error GoTo salir:
        
      
        If IsNumeric(txtIndice_Nro_Documento.Text) Then

                If MDIfrmInicio.StaInicio.Panels(2) = 31 Or _
            MDIfrmInicio.StaInicio.Panels(2) = 82 Or MDIfrmInicio.StaInicio.Panels(2) = 47 Or _
            MDIfrmInicio.StaInicio.Panels(2) = 49 Or MDIfrmInicio.StaInicio.Panels(2) = 46 Or MDIfrmInicio.StaInicio.Panels(2) = 48 _
            Then

            Else
                
               If cboTipoCarga.Text = "Legajos" Then
                    If Trim(txtCajaDigitoVerificador.Text) = "" Then
                        MsgBox "Ingrese el digito Verificador Caja", vbCritical
                        txtCajaDigitoVerificador.SetFocus
                        Exit Sub
                    End If
            End If
        End If
            
            If Configurar_Carga(ctlCliente.Valor, txtIndice_Nro_Documento.Text) Then
                SendKeys vbTab
             End If
             
          End If
        Exit Sub
        
salir:
        
        MsgBox Err.Description
        
       
    End If

End Sub

Private Sub txtLectura_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If UCase(Mid(txtLectura.Text, 1, 2)) = "C5" Then
         txtCaja.Text = Mid(txtLectura.Text, 3, 7)
    End If
    If UCase(Mid(txtLectura.Text, 1, 2)) = "L2" Then
         txtEtiqueta.Text = Mid(txtLectura.Text, 3)
         txtEtiquetaDigitoVerificador.Text = Digito_Verificador(txtEtiqueta.Text)
    End If
 If IsNumeric(txtLectura.Text) Then
    txtNroDesde.Text = txtLectura.Text
   txtNroDesde.SetFocus
 End If
 txtLectura.Text = ""
End If

End Sub

Private Sub txtLetraDesde_Change()
If UCase(lblTituloNumeroDesde.Caption) = "DOCUMENTO" Then
         If Len(txtLetraDesde.Text) = 2 Then
            If UCase(Mid(APELLIDO_NOMBRE, 1, 2)) = UCase(Trim(txtLetraDesde.Text)) Then
                txtLetraDesde.Text = Trim(APELLIDO_NOMBRE)
                txtLetraHasta.Text = "DNI"
            End If
        
        End If
        
    End If
    If txtIndice_Nro_Documento <> "" Then
     If ctlCliente.Valor = 20 And txtIndice_Nro_Documento = 2279 And Len(txtLetraDesde.Text) = 2 Then
         If Nombre_Farmacia <> "" Then
            If Mid(UCase(Nombre_Farmacia), 1, 2) = Mid(UCase(Trim(txtLetraDesde.Text)), 1, 2) Then
                txtLetraHasta.Text = Trim(Nombre_Farmacia)
                txtLetraDesde.Text = Trim(Nombre_Farmacia)
                
            End If
         End If
         
     End If
    End If
End Sub


Private Sub txtLetraDesde_GotFocus()
    If chkFijarLetraDesde.value = 1 Then
        SendKeys vbTab
       Rem FechaFormato
        
    End If

End Sub

Private Sub txtLetraDesde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
    
    If KeyAscii = 43 Then
        KeyAscii = 0
     txtLetraHasta.Text = txtLetraDesde.Text
    SendKeys vbTab
End If
    
End Sub

Private Sub txtLetraDesde_LostFocus()
    If chk_Copiar_Letra.value = 1 Then
        txtLetraHasta.Text = txtLetraDesde.Text
    End If
    
    txtLetraDesde.Text = Replace(txtLetraDesde.Text, "'", " ")


End Sub

Private Sub txtLetraHasta_GotFocus()
If chkFijarLetraHasta.value = 1 Then
        SendKeys vbTab
    End If

End Sub

Private Sub txtLetraHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtLetraHasta_LostFocus()

txtLetraHasta.Text = Replace(txtLetraHasta.Text, "'", " ")
End Sub

Private Sub txtNroDesde_GotFocus()
    If chkFijarNumeroDesde.value = 1 Then
            SendKeys vbTab
    End If
    

End Sub

Private Sub txtNroDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
If KeyAscii = 43 Then
    KeyAscii = 0
     txtNroHasta.Text = txtNroDesde.Text
    SendKeys vbTab
End If
End Sub

Private Sub txtNroDesde_LostFocus()
On Error GoTo salir:

If txtIndice_Nro_Documento.Text = "" Then
    Exit Sub
End If


txtLetraDesde.ToolTipText = ""
    
'If ctlCliente.Valor = 4 And txtIndice_Nro_Documento = 79 Then
'    If txtNroDesde.Text <> "" Then
'        BuscarEcogas txtNroDesde.Text
'    End If
'End If
'If ctlCliente.Valor = 4 And txtIndice_Nro_Documento = 7763 Then
'    If txtNroDesde.Text <> "" Then
'        BuscarEcogas txtNroDesde.Text
'    End If
'End If
If ctlCliente.Valor = 20 And txtIndice_Nro_Documento = 61 Then
    If txtNroDesde.Text <> "" Then
        BuscarOsep txtNroDesde.Text
    End If

End If

APELLIDO_NOMBRE = ""
If UCase(lblTituloNumeroDesde.Caption) = "DOCUMENTO" And txtNroDesde.Text <> "" Then
    Dim con As New ADODB.Connection
    Dim BackColor As ColorConstants
     con.Open strConBasa
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    
 
    
    
 Sql = "SELECT APELLIDO_NOMBRE From PADRON Where DOCUMENTO = " & txtNroDesde.Text
   
   rs.Open Sql, con, adOpenStatic, adLockReadOnly
    If rs.EOF Then
        APELLIDO_NOMBRE = ""
        txtLetraDesde.BackColor = ColorHabilitado
    Else
    
        txtLetraDesde.BackColor = &H80FF80
        APELLIDO_NOMBRE = Trim(rs!APELLIDO_NOMBRE)
        txtLetraDesde.ToolTipText = APELLIDO_NOMBRE
    End If
      con.Close
 End If

        
        
        
        If ctlCliente.Valor = 20 And txtIndice_Nro_Documento = 2279 Then
           Rem Dim sql As String
                Dim conFar As New ADODB.Connection
            Nombre_Farmacia = ""
            If txtNroDesde.Text <> "" Then
                conFar.Open strConBasa
                Dim rsfar As New ADODB.Recordset
                
                If Len(txtNroDesde.Text) < 8 Then
                    Sql = " SELECT     FARMACIA, PAMI, FARMALINK, OSEP"
                    Sql = Sql & " From FARMACIAS "
                    Sql = Sql & "  Where OSEP = " & txtNroDesde.Text
                    Sql = Sql & "  OR FARMALINK = " & txtNroDesde.Text
                    
                Else
                    Sql = " SELECT     FARMACIA, PAMI, FARMALINK, OSEP"
                    Sql = Sql & " From FARMACIAS "
                    Sql = Sql & "  Where PAMI = " & txtNroDesde.Text
                
                End If
                rsfar.Open Sql, conFar
                
                If Not rsfar.EOF Then
                If Not IsNull(rsfar!PAMI) And Not IsNull(rsfar!FARMACIA) Then
                    txtNroDesde.Text = rsfar!PAMI
                    txtNroHasta.Text = txtNroDesde.Text
                    Nombre_Farmacia = Trim(rsfar!FARMACIA)
                    txtLetraDesde.Text = Nombre_Farmacia
                    txtLetraHasta.Text = Nombre_Farmacia
                 Else
                    txtLetraDesde.Text = ""
                    txtLetraHasta.Text = ""
                    txtNroHasta.Text = txtNroDesde.Text
                    Nombre_Farmacia = ""
                 End If
                Else
                                    Nombre_Farmacia = ""
                End If

                End If
                
        End If


If ctlCliente.Valor = 20 And (txtIndice_Nro_Documento = 21 Or txtIndice_Nro_Documento = 222) And Len(txtNroDesde.Text) = 16 Then
    
    txtLetraDesde.Text = UCase(Replace(Mid(txtNroDesde.Text, 7, 2), ".", ""))
    txtLetraHasta.Text = txtLetraDesde.Text
    txtFechaDesde.Text = "01/01/" & Mid(txtNroDesde.Text, 9, 4)
    txtFechaHasta.Text = "31/12/" & Mid(txtNroDesde.Text, 9, 4)
    txtNroDesde.Text = CLng(Mid(txtNroDesde.Text, 1, 6))
    txtNroHasta.Text = txtNroDesde.Text
    
End If
If chk_Copiar_Nro.value = 1 Then
        txtNroHasta.Text = txtNroDesde.Text
    End If

Exit Sub
salir:

MsgBox Err.Description

End Sub

Private Sub BuscarOsep(Doc As String)

Dim rs As New ADODB.Recordset
Dim Sql As String
Dim CONlEGAJOS As ADODB.Connection

 Set CONlEGAJOS = New ADODB.Connection
            CONlEGAJOS.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Serverbasa1\SistemasBasa\ClientesBases\osep.mdb"
       Rem  ConLegajos.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=\\Serverbasa1\SistemasBasa\Etiquetas\CAJAS.mdb"

Sql = "  SELECT   NUMERO_AFI, APELLIDO_NOMBRE AS Nombre"
Sql = Sql & vbCrLf & " , 'Doc:  ' & [NUMERO_AFI] AS Descripcion"
Sql = Sql & vbCrLf & "  From OSEPAFILI "
Sql = Sql & vbCrLf & "  WHERE NUMERO_AFI = " & Doc
rs.Open Sql, CONlEGAJOS
    If Not rs.EOF Then
       Rem txtDescripcion.Text = rs!DESCRIPCION
        txtLetraDesde.Text = rs!Nombre
        txtLetraHasta.Text = rs!Nombre
    Else
            txtLetraHasta.Text = ""
            txtLetraDesde.Text = ""
            txtDescripcion.Text = ""
    End If
End Sub


Private Sub txtNroHasta_GotFocus()
If chkFijarNumeroHasta.value = 1 Then
    SendKeys vbTab
End If

End Sub

Private Sub txtNroHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If

End Sub


Public Sub ActualizarLegajos(FK_Indice As String, Cod_Indice As String, LETRA_DESDE As String, LETRA_HASTA As String, NRO_DESDE As String, NRO_HASTA As String _
, FECHA_DESDE As String, FECHA_HASTA As String, Descripcion As String _
, FECHA_ACTUALIZACION As String, NRO_CAJA As String, COD_CLIENTE As String, ID_Personal As String, NRO_REM_PROV As String, ID_LEGAJO As Long)


Dim Sql As String
Dim Registros As Integer
Dim Des As String
Dim CONlEGAJOS As New ADODB.Connection
CONlEGAJOS.Open strConBasa


On Error GoTo salir:

    Sql = " Update LEGAJOS "
    Sql = Sql & vbCrLf & " SET FK_INDICES =" & FK_Indice
    Sql = Sql & vbCrLf & " , COD_INDICE =" & Cod_Indice
    
    If LETRA_DESDE <> "" Then
        Sql = Sql & vbCrLf & " , LETRA_DESDE =" & LETRA_DESDE
    End If
    If LETRA_HASTA <> "" Then
        Sql = Sql & vbCrLf & " , LETRA_HASTA =" & LETRA_HASTA
    End If
    If NRO_DESDE <> "" Then
        Sql = Sql & vbCrLf & " , NRO_DESDE =" & NRO_DESDE
    End If
    If NRO_HASTA <> "" Then
        Sql = Sql & vbCrLf & " , NRO_HASTA =" & NRO_HASTA
    End If
    If FECHA_DESDE <> "" Then
        Sql = Sql & vbCrLf & " , FECHA_DESDE =" & (FECHA_DESDE)
        Sql = Sql & vbCrLf & " , FECHA_HASTA =" & (FECHA_HASTA)
    End If
    
    Sql = Sql & vbCrLf & " , DESCRIPCION =" & Descripcion
    
    If APELLIDO_NOMBRE = Trim(Replace(LETRA_DESDE, "'", "")) Then
        Sql = Sql & vbCrLf & " , CONTROL_PADRON =1"
    End If
    
    If StatusBar.Panels(1).Text = "Modificar" Then
        Sql = Sql & vbCrLf & " , FK_PERSONAL_ACTUALIZACION   = " & ID_Personal
        Sql = Sql & vbCrLf & " , FECHA_ACTUALIZACION =" & FECHA_ACTUALIZACION
    Else
        
        
        Sql = Sql & vbCrLf & " , FK_PERSONAL_CREACION = " & ID_Personal
        Sql = Sql & vbCrLf & " , FECHA_CREACION =" & FECHA_ACTUALIZACION
        Sql = Sql & vbCrLf & " , FECHA_ACTUALIZACION = " & FECHA_ACTUALIZACION
         Sql = Sql & vbCrLf & " , FK_PERSONAL_ACTUALIZACION   = " & ID_Personal
        
        Sql = Sql & vbCrLf & " , COD_ESTADO  = 2 "
    End If
    
    Sql = Sql & vbCrLf & " , NRO_CAJA =" & NRO_CAJA
    Sql = Sql & vbCrLf & " , COD_CLIENTE =" & COD_CLIENTE
 Rem    Sql = Sql & vbCrLf & " , NRO_REM_PROV = " & NRO_REM_PROV
    Sql = Sql & vbCrLf & " , DESCRIPCION_REMITO = '" & Des & "'"
    Sql = Sql & vbCrLf & "  Where ID_CLIENTE_LEGAJO = " & ID_LEGAJO
    If StatusBar.Panels(1).Text <> "Modificar" Then
        Sql = Sql & vbCrLf & "  AND COD_CLIENTE is null"
    Else
         Sql = Sql & vbCrLf & " AND COD_CLIENTE =" & ctlCliente.Valor
         StatusBar.Panels(1).Text = "Nuevo"
         StatusBar.Panels(2).Text = ""
    End If
    CONlEGAJOS.Execute Sql, Registros
    Legajos_RecalcularCaracteres_DescripcionRemito " Where ID_LEGAJO = " & ID_LEGAJO
      Clipboard.SetText Clipboard.GetText & vbCrLf & ID_LEGAJO

    If Registros = 0 Then
        MsgBox "Atención No se ingreso el Legajo", vbCritical
    End If
            Beep
CONlEGAJOS.Close
Exit Sub


salir:
 MsgBox "No se actualizo el registro verifique los datos", vbCritical
 CONlEGAJOS.Close
 
 
 


End Sub

Public Sub LimpiarCampos(RespetarFijarValor As Boolean)
 
If chkFijarCaja.value = 0 Then
    txtCaja.Text = ""
End If

If chkFijarTipoDocumento.value = 0 Then
    txtIndice_Nro_Documento.Text = ""
    lblIndice_Descripcion.Caption = ""
End If

If chkFijarDescripcion.value = 0 Then
    txtDescripcion.Text = ""
End If

If chkFijarFechaDesde.value = 0 Then
    txtFechaDesde.Text = ""
End If

If chkFijarFechaHasta.value = 0 Then
    txtFechaHasta.Text = ""
End If

If chkFijarLetraDesde.value = 0 Then
    txtLetraDesde.Text = ""
End If

If chkFijarLetraHasta.value = 0 Then
    txtLetraHasta.Text = ""
End If

If chkFijarNumeroDesde.value = 0 Then
    txtNroDesde.Text = ""
End If

If chkFijarNumeroHasta.value = 0 Then
    txtNroHasta.Text = ""
End If





End Sub



Private Sub txtUsuarioCarga_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If

End Sub

Public Sub LimpiarLegajos()
   
    
    chkFijarDescripcion.value = 0
    txtDescripcion.Text = ""
    txtDescripcion.Enabled = False
    txtDescripcion.BackColor = ColorDesaHabilitado
    
    chkFijarFechaDesde.value = 0
    txtFechaDesde.Text = ""
    txtFechaDesde.Enabled = False
    txtFechaDesde.BackColor = ColorDesaHabilitado
    
    chkFijarFechaHasta.value = 0
    txtFechaHasta.Text = ""
    txtFechaHasta.Enabled = False
    txtFechaHasta.BackColor = ColorDesaHabilitado
    
    chkFijarLetraDesde.value = 0
    txtLetraDesde.Text = ""
    txtLetraDesde.Enabled = False
    txtLetraDesde.BackColor = ColorDesaHabilitado
    
    chkFijarLetraHasta.value = 0
    txtLetraHasta.Text = ""
    txtLetraHasta.Enabled = False
    txtLetraHasta.BackColor = ColorDesaHabilitado
    
    chkFijarNumeroDesde.value = 0
    txtNroDesde.Text = ""
    txtNroDesde.Enabled = False
    txtNroDesde.BackColor = ColorDesaHabilitado
    
    chkFijarNumeroHasta.value = 0
    txtNroHasta.Text = ""
    txtNroHasta.Enabled = False
    txtNroHasta.BackColor = ColorDesaHabilitado
     
    
End Sub

Public Function ISNULLFALSE(DATO) As Boolean

If IsNull(DATO) Then
    ISNULLFALSE = False
Else
ISNULLFALSE = DATO
End If


End Function

Public Sub AgregarQuincena(DATO As String)
    Dim Mes As String
    Dim MesNumero As String
    Dim Quincena As String
    Dim DiaInicio As String
    Dim DiaFin As String
    Dim Año As String
    On Error GoTo salir:
If txtIndice_Nro_Documento.Text = 2279 And Len(Trim(txtFechaDesde.Text)) = 5 Then
Quincena = Mid(DATO, 1, 1)
Mes = UCase(Trim(Mid(DATO, 2, 2)))
Año = UCase(Trim(Mid(DATO, 4, 2)))
Select Case Mes
Case "01"
    MesNumero = 1
Case "02"
    MesNumero = 2
Case "03"
    MesNumero = 3
Case "04"
    MesNumero = 4
Case "05"
    MesNumero = 5
Case "06"
    MesNumero = 6
Case "07"
    MesNumero = 7
Case "08"
    MesNumero = 8
Case "09"
    MesNumero = 9
Case "10"
    MesNumero = 10
Case "11"
    MesNumero = 11
Case "12"
    MesNumero = 12
End Select


Select Case Quincena
Case 1
    DiaInicio = "01"
    DiaFin = "15"
Case 2
    DiaInicio = "16"
    Select Case MesNumero
    Case 1, 3, 5, 7, 8, 10, 12
        DiaFin = "31"
    Case 2
        DiaFin = "28"
    Case Else
        DiaFin = "30"
    End Select

End Select
    
        txtFechaDesde.Text = DiaInicio & "/" & Format(MesNumero, "00") & "/20" & Año
        txtFechaHasta.Text = DiaFin & "/" & Format(MesNumero, "00") & "/20" & Año
    End If
    Exit Sub
salir:
    MsgBox Err.Description

End Sub

Public Sub AgregarMes(DATO As String)
    Dim Mes As String
    Dim MesNumero As String
    Dim Quincena As String
    Dim DiaInicio As String
    Dim DiaFin As String
    Dim Año As String


Mes = UCase(Trim(Mid(DATO, 1, 2)))
Año = UCase(Trim(Mid(DATO, 3)))
If Mes = "01" Or Mes = "03" Or Mes = "05" Or Mes = "07" Or Mes = "08" Or Mes = "10" Or Mes = "12" Then
    txtFechaDesde.Text = "01/" & Mes & "/" & Año
   txtFechaHasta.Text = "31/" & Mes & "/" & Año
End If

If Mes = "02" Then
    txtFechaDesde.Text = "01/" & Mes & "/" & Año
    txtFechaHasta.Text = "29/" & Mes & "/" & Año
End If

If Mes = "04" Or Mes = "06" Or Mes = "09" Or Mes = "11" Then
    txtFechaDesde.Text = "01/" & Mes & "/" & Año
    txtFechaHasta.Text = "30/" & Mes & "/" & Año
End If
    


End Sub

Public Sub BuscarDescripcion()
Dim rs As New ADODB.Recordset
    Dim Sql As String
        Sql = " SELECT      dbo.REFERENCIAS.DESCRIPCION, COUNT(*) AS Cantidad, dbo.INDICES.DESCRIPCION AS Sector,"
        Sql = Sql & " dbo.INDICES.Indice , dbo.INDICES.ID_CODIGO_DOCUMENTO "
        Sql = Sql & " FROM         dbo.REFERENCIAS INNER JOIN"
        Sql = Sql & " dbo.INDICES ON dbo.REFERENCIAS.COD_CLIENTE = dbo.INDICES.COD_CLIENTE"
        Sql = Sql & " AND dbo.REFERENCIAS.INDICE = dbo.INDICES.INDICE "
        Sql = Sql & " GROUP BY dbo.REFERENCIAS.COD_CLIENTE, dbo.REFERENCIAS.DESCRIPCION, dbo.INDICES.DESCRIPCION, dbo.REFERENCIAS.COD_DOCUMENTO,"
        Sql = Sql & " dbo.INDICES.ID_CODIGO_DOCUMENTO , dbo.INDICES.Indice, dbo.INDICES.ID_CODIGO_DOCUMENTO "
        Sql = Sql & " HAVING  dbo.REFERENCIAS.COD_CLIENTE = " & ctlCliente.Valor
        Sql = Sql & " AND (dbo.REFERENCIAS.DESCRIPCION LIKE '%" & txtDescripcion.Text & "%')"
        If txtIndice_Nro_Documento.Text <> "" Then
            Sql = Sql & " AND (dbo.INDICES.INDICE LIKE '" & BuscarIndiceDocumento_Indice(txtIndice_Nro_Documento.Text, ctlCliente.Valor) & "%')"
        Else

        End If
        Sql = Sql & " ORDER BY COUNT(*) DESC"

 rs.CursorLocation = adUseClient
        rs.Open Sql, ConActiva, 0, 1
        Set grdDescripcion.DataSource = rs.DataSource
        grdDescripcion.Columns.Item(0).Width = 5715
        grdDescripcion.Columns.Item(1).Width = 915
        grdDescripcion.Columns.Item(2).Width = 2295
        grdDescripcion.Columns.Item(3).Width = 1740



        SSTab1.Tab = 2
    

End Sub


Public Sub Buscar_Inidice_Por_caja()
    Dim rsIndice As New ADODB.Recordset
  
  Dim sSQL As String
  Dim Item As Integer
        
        
        sSQL = " SELECT CLIENTEUSUARIO.COD_INDICE"
        sSQL = sSQL & vbCrLf & " FROM REMITOS_CUERPO, REMITOS_DETALLE, CLIENTEUSUARIO"
        sSQL = sSQL & vbCrLf & " Where REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO  AND"
        sSQL = sSQL & vbCrLf & " REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
        sSQL = sSQL & vbCrLf & " AND REMITOS_CUERPO.ID_CLIENTE = " & ctlCliente.Valor
        sSQL = sSQL & vbCrLf & " AND REMITOS_DETALLE.DESDE =" & txtCaja.Text
        sSQL = sSQL & vbCrLf & " AND REMITOS_CUERPO.TIPO = 0"

        rsIndice.Open sSQL, ConActiva, 0, 1
        If Not rsIndice.EOF Then
        
        frmIndice.COD_CLIENTE = ctlCliente.Valor
        frmIndice.Actualizar ctlCliente.Valor, Nulo, 0, rsIndice!Cod_Indice
        frmAgregarDocumentos.WindowState = 0
        frmIndice.Show
        frmIndice.SetFocus
         Else
                MsgBox "No se encontro el remito"
                txtIndice_Nro_Documento.Text = 1000
                txtIndice_Nro_Documento.SetFocus
                
                
        
        End If
End Sub







Private Sub txtVideoLugar_GotFocus()
Timer1.Enabled = False
End Sub

Private Sub txtVideoLugar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Video.Controls.currentPosition = txtVideoLugar.Text
         Play_Pausa "Play"
        Timer1.Enabled = True
            Play_Pausa "Pausa"
    End If

End Sub

Public Sub Play_Pausa(Valor As String)
If Valor = "Play" Then
    Video.Controls.play
    Timer1.Enabled = True
    ValorAnteVideo = 0

    cmdPlay_Pausa.Caption = "Pausa"
    Exit Sub
End If
If Valor = "Pausa" Then
    Video.Controls.pause
    Timer1.Enabled = False
    ValorAnteVideo = 0
    cmdPlay_Pausa.Caption = "Play"
End If
End Sub

