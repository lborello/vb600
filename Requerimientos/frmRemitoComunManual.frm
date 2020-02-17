VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{40CE97D1-1C1F-47E7-B2C4-A9B643CAAFFD}#2.0#0"; "Controles.ocx"
Begin VB.Form frmRemitoManualNuevo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remito"
   ClientHeight    =   7080
   ClientLeft      =   105
   ClientTop       =   210
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9540
   Begin VB.Frame fraCliente 
      Height          =   1695
      Left            =   0
      TabIndex        =   12
      Top             =   660
      Width           =   4695
      Begin Controles.cltGenerico ctlPersonal 
         Height          =   375
         Left            =   960
         TabIndex        =   29
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
      End
      Begin Controles.cltGenerico ctlCliente 
         Height          =   375
         Left            =   960
         TabIndex        =   28
         Top             =   120
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
      End
      Begin MSMask.MaskEdBox mskFechaRemito 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   900
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskNumeroRemitoProv 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   1260
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "0001-0000####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCantidad 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3780
         TabIndex        =   21
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2820
         TabIndex        =   20
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label11 
         Caption         =   "Remito:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         TabIndex        =   17
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   795
      End
      Begin VB.Label Label12 
         Caption         =   "Resp.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   675
      End
   End
   Begin VB.Frame fraRemito 
      Height          =   1695
      Left            =   4800
      TabIndex        =   7
      Top             =   660
      Width           =   4695
      Begin Controles.cltGenerico ctlTipo_Almacenamiento 
         Height          =   375
         Left            =   1320
         TabIndex        =   33
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
      End
      Begin Controles.cltGenerico ctlTipo_Estado 
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   840
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
      End
      Begin Controles.cltGenerico ctlTipo_Operacion 
         Height          =   375
         Left            =   1320
         TabIndex        =   31
         Top             =   480
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
      End
      Begin Controles.cltGenerico ctlTipo_Remito 
         Height          =   375
         Left            =   1320
         TabIndex        =   30
         Top             =   120
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
      End
      Begin VB.CommandButton cmdColector 
         Caption         =   "..."
         Height          =   315
         Left            =   4260
         TabIndex        =   19
         Top             =   1260
         Width           =   315
      End
      Begin VB.TextBox txtElemento 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   360
         Left            =   2880
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Elemento:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Operación:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraGrilla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   0
      TabIndex        =   5
      Top             =   2280
      Width           =   9495
      Begin Controles.ctlClienteUsuario ctlClienteUsuario1 
         Height          =   375
         Left            =   720
         TabIndex        =   34
         Top             =   3600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
      End
      Begin VB.TextBox txtDescripcionConsulta 
         Height          =   675
         Left            =   60
         TabIndex        =   4
         Top             =   4020
         Width           =   9375
      End
      Begin MSFlexGridLib.MSFlexGrid grdElementos 
         Height          =   3315
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   6
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblSector 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4980
         TabIndex        =   26
         Top             =   3600
         Width           =   4455
      End
      Begin VB.Label Label6 
         Caption         =   "Solito: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   60
         TabIndex        =   25
         Top             =   3660
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "Sector:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4140
         TabIndex        =   24
         Top             =   3660
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3540
         TabIndex        =   18
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   13200
      TabIndex        =   0
      Top             =   4380
      Width           =   1200
   End
   Begin Crystal.CrystalReport CryRemito 
      Left            =   6900
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   690
      Left            =   780
      TabIndex        =   16
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   1217
      ButtonWidth     =   1191
      ButtonHeight    =   1058
      Appearance      =   1
      ImageList       =   "ImageList2020"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ant"
            Key             =   "Ant"
            Object.ToolTipText     =   "Pagina anterior"
            ImageKey        =   "Ant"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sig"
            Key             =   "Sig"
            Object.ToolTipText     =   "Pagina siguiente"
            ImageKey        =   "Sig"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimir la imagen"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo Requerimiento"
            ImageKey        =   "Nuevo"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "Modif."
            Key             =   "Modif."
            Object.ToolTipText     =   "Modificación de un Requerimiento"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anular"
            Key             =   "Anular"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            Key             =   "Aceptar"
            ImageKey        =   "Aceptar"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            ImageKey        =   "Buscar"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Control"
            Key             =   "Control"
            ImageKey        =   "Control"
         EndProperty
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   8460
         TabIndex        =   23
         Top             =   180
         Width           =   795
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   7500
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   0
         Width           =   1275
      End
      Begin VB.Label lblEstado 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5700
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSComctlLib.ImageList ImageList2020 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   51
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":0000
            Key             =   "Ver+"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":03FA
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":069A
            Key             =   "Ver-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":0A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":0EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":1274
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":1389
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":17AC
            Key             =   "RotarI"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":1BBD
            Key             =   "Vertical"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":1FDC
            Key             =   "Sig"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":23DD
            Key             =   "Ant"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":27DB
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":2BF1
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":2CA7
            Key             =   "RotarD"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":30B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":348C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":3854
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":3946
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":3D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":4149
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":4509
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":45A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":4649
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":4A39
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":4E25
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":5200
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":561A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":57C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":5BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":5CAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":60B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":64C4
            Key             =   "Fin"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":68F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":6A27
            Key             =   "Aceptar"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":6DE0
            Key             =   "Control"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":6ECB
            Key             =   "Esp. Fax"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":72FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":7435
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":7846
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":79BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":7DC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":81F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":85D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":89E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":8DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":91B0
            Key             =   "Anular"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":957E
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":9987
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":9D53
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":A187
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRemitoComunManual.frx":A5CC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRemitoManualNuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ctlCliente.TipoControl = Cliente
    ctlPersonal.TipoControl = Personal
    ctlTipo_Almacenamiento.TipoControl = Tipo_Remito_almacenamiento
    ctlTipo_Estado.TipoControl = Tipo_Remito_Estados
    ctlTipo_Operacion.TipoControl = Tipo_Remito_Operacion
    Rem ctlTipo_Remito.TipoControl = Tipo_Remito
End Sub

Private Sub mskNumeroRemito_Change()

End Sub

Public Sub Remito_Nuevo()
Dim RemitoNuevo As New clsRemitos
RemitoNuevo.Remto_ADD mskNumeroRemitoProv, ctlTipo_Remito, ctltipi

End Sub
