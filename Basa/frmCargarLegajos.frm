VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{C30A4F2E-16E3-4694-9920-512C55E5C51A}#1.0#0"; "Controles4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCargarLegajos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargar Legajos"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   12210
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   12210
   Begin VB.Frame fraCargarLegajos 
      Height          =   7035
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   11955
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   495
         Left            =   4260
         TabIndex        =   36
         Top             =   4740
         Width           =   1635
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   2040
         TabIndex        =   35
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtCantidadElementos 
         Height          =   315
         Left            =   4140
         TabIndex        =   32
         Top             =   1140
         Width           =   555
      End
      Begin MSDataGridLib.DataGrid grdLegajos 
         Height          =   2895
         Left            =   60
         TabIndex        =   28
         Top             =   1920
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   5106
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
      Begin VB.CheckBox chkDescripcion 
         Alignment       =   1  'Right Justify
         Caption         =   "Descrip."
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
         Left            =   5820
         TabIndex        =   27
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CheckBox chkNombre 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre"
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
         Left            =   60
         TabIndex        =   26
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   555
         Left            =   60
         TabIndex        =   25
         Top             =   4740
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   435
         Left            =   60
         TabIndex        =   24
         Top             =   5280
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox chkRotulo 
         Alignment       =   1  'Right Justify
         Caption         =   "Incre. Aut."
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
         Left            =   4740
         TabIndex        =   20
         Top             =   1140
         Width           =   1215
      End
      Begin VB.CheckBox chkPegado 
         Alignment       =   1  'Right Justify
         Caption         =   "Pegado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1980
         TabIndex        =   4
         Top             =   720
         Width           =   1035
      End
      Begin Controles.cltIndice ctlIndice 
         Height          =   1935
         Left            =   120
         TabIndex        =   19
         Top             =   4980
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   3413
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Controles.cltGenerico ctlPersonal 
         Height          =   315
         Left            =   6600
         TabIndex        =   1
         Top             =   300
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
      End
      Begin MSMask.MaskEdBox mskRemitoProv 
         Height          =   315
         Left            =   10380
         TabIndex        =   2
         Top             =   300
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   13
         Mask            =   "0001-000#####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtDocumento 
         Height          =   315
         Left            =   4140
         TabIndex        =   5
         Top             =   720
         Width           =   555
      End
      Begin VB.TextBox txtID_Cliente_Legajos 
         BackColor       =   &H80000009&
         Height          =   315
         Left            =   6120
         TabIndex        =   6
         Top             =   1140
         Width           =   1875
      End
      Begin VB.TextBox txtCaja 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1260
         TabIndex        =   8
         Top             =   1500
         Width           =   4395
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   7080
         TabIndex        =   9
         Top             =   1500
         Width           =   4755
      End
      Begin MSMask.MaskEdBox mskLegajo_Cliente 
         Height          =   315
         Left            =   9060
         TabIndex        =   7
         Top             =   1140
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   180
         Top             =   2580
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   46
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":0279
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":0637
               Key             =   "Print"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":0A71
               Key             =   "Borrar1"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":0E74
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":129B
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":1642
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":1A00
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":1DC6
               Key             =   "Salvar2"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":2042
               Key             =   "Nuevo"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":22C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":268D
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":2A4A
               Key             =   "Modificar"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":2CCB
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":2F3C
               Key             =   "Casa"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":330F
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":36E7
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":3964
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":3B35
               Key             =   "Atras2"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":3EEF
               Key             =   "Inicio"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":3FD4
               Key             =   "Fin"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":40B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":4484
               Key             =   "Adelante2"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":4838
               Key             =   "Correo2"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":4C2B
               Key             =   "Bandera"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":4E8E
               Key             =   "trvt2"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":524F
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":5B29
               Key             =   "Buscar"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":5DC9
               Key             =   "Cancelar1"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":60E3
               Key             =   "Aceptar1"
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":63FD
               Key             =   "trvt"
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":64D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":65A9
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":69D0
               Key             =   "Atras3"
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":70CA
               Key             =   "Atras"
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":77C4
               Key             =   "Adelante"
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":7EBE
               Key             =   "Correo3"
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":85B8
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":8CB2
               Key             =   "Correo4"
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":904C
               Key             =   "Correo"
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":9D26
               Key             =   "Borrar"
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":A600
               Key             =   "Punto"
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":AEDA
               Key             =   "Cancelar2"
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":B2A8
               Key             =   "Aceptar2"
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":B661
               Key             =   "Aceptar"
            EndProperty
            BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCargarLegajos.frx":BF3B
               Key             =   "Cancelar"
            EndProperty
         EndProperty
      End
      Begin Controles.cltGenerico ctlCliente 
         Height          =   375
         Left            =   960
         TabIndex        =   0
         Top             =   300
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   661
      End
      Begin VB.Label Label11 
         Caption         =   "Legajo:"
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
         Left            =   8220
         TabIndex        =   34
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad:"
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
         Left            =   3240
         TabIndex        =   33
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Etiqueta"
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
         Left            =   60
         TabIndex        =   31
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Inicio"
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
         Left            =   960
         TabIndex        =   30
         Top             =   1140
         Width           =   555
      End
      Begin VB.Label lblInicial 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1920
         TabIndex        =   29
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label9 
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
         Left            =   9660
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Cliente:"
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
         Left            =   60
         TabIndex        =   17
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label5 
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
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Resp.:"
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
         Left            =   5880
         TabIndex        =   15
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Nº Doc"
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
         Left            =   3240
         TabIndex        =   14
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblCodigo_Indice 
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
         Height          =   315
         Left            =   4680
         TabIndex        =   11
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label lblCodigo_Nombre 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   8040
         TabIndex        =   12
         Top             =   720
         Width           =   3855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   630
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   1111
      ButtonWidth     =   1667
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            ImageKey        =   "Nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar F9"
            Key             =   "Aceptar"
            ImageKey        =   "Aceptar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            ImageKey        =   "Print"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   5
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Control Carga por cliente"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Control Carga"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Control carga estadistico"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir comprobante"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Control"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Grilla"
            Description     =   "Grilla"
            ImageIndex      =   42
         EndProperty
      EndProperty
      Begin VB.Label Label6 
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
         Left            =   5460
         TabIndex        =   21
         Top             =   0
         Width           =   1395
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   53
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":C815
            Key             =   "Ver+"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":CC0F
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":CEAF
            Key             =   "Ver-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":D2AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":D6C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":DA89
            Key             =   "Punto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":DB9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":DFC1
            Key             =   "RotarI"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":E3D2
            Key             =   "Vertical"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":E7F1
            Key             =   "Sig"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":EBF2
            Key             =   "Ant"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":EFF0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":F406
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":F4BC
            Key             =   "RotarD"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":F8C6
            Key             =   "Cargar"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":FCA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":10069
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":1015B
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":1054B
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":1095E
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":10D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":10DBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":10E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":1124E
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":1163A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":11A15
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":11E2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":11FD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":123B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":124C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":128C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":12CD9
            Key             =   "Fin"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":13105
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":1323C
            Key             =   "Aceptar"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":135F5
            Key             =   "Control"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":136E0
            Key             =   "Esp. Fax"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":13B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":13C4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":1405B
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":141D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":145DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":14A0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":14DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":151F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":155F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":159C5
            Key             =   "Anular"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":15D93
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":1619C
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":16568
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":1699C
            Key             =   "grilla"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":16DE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":17079
            Key             =   "Casa"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCargarLegajos.frx":17480
            Key             =   "Bandera"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Left            =   6720
      TabIndex        =   23
      Top             =   180
      Width           =   1995
   End
   Begin VB.Label Label10 
      Caption         =   "Estado:"
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
      Left            =   6000
      TabIndex        =   22
      Top             =   180
      Width           =   735
   End
   Begin VB.Menu mnuArbol 
      Caption         =   "Arbol"
      Begin VB.Menu mnuBuscarIndice 
         Caption         =   "Buscar Indice"
      End
      Begin VB.Menu mnuBuscarLegajos 
         Caption         =   "Buscar Legajos"
      End
      Begin VB.Menu mnuRefrescar 
         Caption         =   "Refrescar"
      End
   End
End
Attribute VB_Name = "frmCargarLegajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CONlEGAJOS As ADODB.Connection
'Public Sub CargarIndices(rsIndices As ADODB.Recordset)
'    Dim Indice0 As String
'    Dim KeyTreeView1 As String
'    Dim Indice1 As String
'    Dim DESCRIPCION As String
'    Dim nodX As Node
'        trvIndices.Nodes.Clear
'        Set nodX = trvIndices.Nodes.Add(, , "RAIZ", " TODAS LAS CATEGORIAS", "Casa") ' Root
'        trvIndices.Nodes.Item("RAIZ").Tag = "TODOS"
'        Do While Not rsIndices.EOF
'            If ExisItem("R" & Mid(rsIndices!Indice, 1, Len(rsIndices!Indice) - 3)) Then
'                KeyTreeView1 = "R" & Mid(rsIndices!Indice, 1, Len(rsIndices!Indice) - 3)
'                DESCRIPCION = rsIndices!Indice & " - " & rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!DESCRIPCION)
'                Set nodX = trvIndices.Nodes.Add(KeyTreeView1, tvwChild, "R" & rsIndices!Indice, DESCRIPCION, "Punto", "Bandera")
'                trvIndices.Nodes.Item("R" & rsIndices!Indice).Tag = rsIndices!Indice
'            Else
'                DESCRIPCION = rsIndices!Indice & " - " & rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!DESCRIPCION)
'                Set nodX = trvIndices.Nodes.Add(, , "R" & rsIndices!Indice, DESCRIPCION, "Punto", "Bandera")  ' Root
'                trvIndices.Nodes.Item("R" & rsIndices!Indice).Tag = rsIndices!Indice
'            End If
'            rsIndices.MoveNext
'        Loop
'End Sub
'Public Function ExisItem(Dato As String) As Boolean
'    Dim s As String
'    On Error GoTo ErrorHandler
'        ExisItem = True
'        s = trvIndices.Nodes.Item(Dato)
'    Exit Function
'ErrorHandler:
'    ExisItem = False
'End Function



Private Sub Nuevo()
    Dim sql As String
    Dim NUMERO_LEGAJO_CLIENTE As Long
    Dim Cod_Indice, CLIENTE_LEGAJO, Descripcion As String
    Dim NRO_CAJA, COD_CLIENTE, Cod_Estado As Integer
    Dim Nombre As String
    Dim Control As Integer
    Dim FECHA_ACTUALIZACION As String
    Dim ID_Personal As Integer
    Dim CANTIDAD_CARACTERES As Integer
    
    
'    If InStr(1, mskLegajo_Cliente.Text, "_", vbTextCompare) <> 0 Then
'        MsgBox "Error en el formato"
'        Exit Sub
'    End If
'

    If lblCodigo_Indice.Caption <> "" Then
        Cod_Indice = "'" & lblCodigo_Indice.Caption & "'"
    Else
        MsgBox "Falta el Indice"
        Exit Sub
    End If
    If mskLegajo_Cliente.ClipText = "" Then
         MsgBox "Falta el cliente legajo"
        Exit Sub
    Else
        CLIENTE_LEGAJO = "'" & mskLegajo_Cliente.Text & "'"
        
        
        If IsNumeric(Replace(mskLegajo_Cliente.Text, "_", "")) Then
            NUMERO_LEGAJO_CLIENTE = Replace(mskLegajo_Cliente.Text, "_", "")
        Else
            NUMERO_LEGAJO_CLIENTE = 0
        End If
    End If
    If mskRemitoProv.Text = "0001-000_____" Then
        MsgBox "Falta Nuemro de Remito Provisorio"
        Exit Sub
    End If
    
    If Trim(txtDescripcion.Text) = "" Then
        Descripcion = "NULL"
    Else
        Descripcion = "'" & UCase(Trim(txtDescripcion.Text)) & "'"
    End If
    NRO_CAJA = txtCaja.Text
    COD_CLIENTE = ctlCliente.Valor
    Cod_Estado = "2"
    If Trim(txtNombre.Text) = "" Then
        Nombre = "NULL"
    Else
        Nombre = "'" & UCase(Trim(txtNombre.Text)) & "'"
    End If
    FECHA_ACTUALIZACION = FechaServerTipo(date)
    If IsNumeric(ctlPersonal.Valor) Then
        ID_Personal = ctlPersonal.Valor
    Else
        MsgBox "El responsable no es correcto"
        Exit Sub
    End If
    Select Case ctlCliente.Valor
    Case 4
        CANTIDAD_CARACTERES = Len(txtID_Cliente_Legajos.Text) + Len(mskLegajo_Cliente.Text)
    Case Else
        CANTIDAD_CARACTERES = Len(Trim(txtID_Cliente_Legajos.Text)) + Len(Trim(mskLegajo_Cliente.Text)) + Len(Trim(txtNombre.Text)) + Len(Trim(txtDescripcion.Text))
    End Select
    
'    If ctlCliente.Valor = 20 And Cod_Indice = "'002002002008'" Then
'
'        Sql = " UPDATE OSEP_LEGAJOS_ARCHIVO"
'        Sql = Sql & vbCrLf & " Set COD_CLIENTE_LEGAJO = " & txtID_Cliente_Legajos.Text
'        Sql = Sql & vbCrLf & " WHERE "
'        Sql = Sql & vbCrLf & "  LEGAJO = " & CLIENTE_LEGAJO
'        ExecutarSql Sql, Control
'        If Control <> 1 Then
'            MsgBox "El legajo no fue encontrado", vbCritical
'            Exit Sub
'        End If
'    End If
            
            
            sql = " Update LEGAJOS "
            sql = sql & vbCrLf & " SET COD_INDICE =" & Cod_Indice & ", CLIENTE_LEGAJO =" & CLIENTE_LEGAJO
            sql = sql & vbCrLf & " , DESCRIPCION =" & Descripcion & " , NRO_CAJA = " & NRO_CAJA
            sql = sql & vbCrLf & " , COD_CLIENTE =" & COD_CLIENTE & " , COD_ESTADO = 2"
            sql = sql & vbCrLf & " , NOMBRE =" & Nombre & ", FECHA_ACTUALIZACION =" & SysDateMinutoSegundo
            sql = sql & vbCrLf & " ,  ID_PERSONAL = " & ID_Personal & ", NRO_REM_PROV ='" & mskRemitoProv.Text & "'"
            sql = sql & vbCrLf & " , NUMERO_LEGAJO_CLIENTE =" & NUMERO_LEGAJO_CLIENTE
            sql = sql & vbCrLf & " , PEGADOETIQUETA='" & chkPegado.value & "',CANTIDAD_CARACTERES = " & CANTIDAD_CARACTERES
            sql = sql & vbCrLf & " WHERE ID_CLIENTE_LEGAJO = " & txtID_Cliente_Legajos.Text
            
           If CLng(txtID_Cliente_Legajos.Text) < 750000 Then
                sql = sql & vbCrLf & " AND COD_CLIENTE = " & COD_CLIENTE
            End If
                sql = sql & vbCrLf & " AND CLIENTE_LEGAJO IS NULL "
            Dim CantidadAfectada As Long
            
             CantidadAfectada = ExecutarSql(sql)
            If CantidadAfectada <> 1 Then
                MsgBox "La  etiqueta ya esta utilizada", vbCritical
                Exit Sub
            Else
                InsertarProducion ctlPersonal.Valor, 10, txtID_Cliente_Legajos.Text, 1, ctlCliente.Valor
            End If
            If chkRotulo.value <> 1 Then
                 txtID_Cliente_Legajos.Text = ""
            Else
                  txtID_Cliente_Legajos.Text = CLng(txtID_Cliente_Legajos.Text) + 1
            End If
            LimpiarMask mskLegajo_Cliente
            txtDocumento.Text = ""
            txtDescripcion.Text = ""
            txtNombre.Text = ""
            If chkRotulo.value = 1 Then
                mskLegajo_Cliente.SetFocus
            Else
                txtID_Cliente_Legajos.SetFocus
            End If
            
           
            
            Beep

End Sub


Private Sub cmdAceptar_Click()
Select Case lblEstado.Caption
Case "Nuevo"
    Nuevo
Case "Modificar"
    Modificar
Case Else
    MsgBox "Error en el estado", vbInformation
End Select


End Sub

Private Sub ImprimirComprobante()
    Dim sql As String
            MousePointer = 11
            sql = " SELECT * "
            sql = sql & "  From  LEGAJOS "
            sql = sql & "  Where NRO_REM_PROV = '" & mskRemitoProv.Text & "'"
            sql = sql & "  Order By ID_CLIENTE_LEGAJO Asc"
            frmReportes.ImprimirReporte PasoReportes & "rptComprobanteRetiroLegajos.rpt", sql, True
           MousePointer = 0
End Sub

Private Sub ControlCaracteres()
    Dim sql As String
    Dim fecha As String
        If Not IsNull(ctlPersonal.Valor) Then
            fecha = InputBox("Ingrese la Fecha de Control ", "Control Carga", Format(Now, "DD/mm/yyyy"))
            MousePointer = 11
            sql = " SELECT * "
            sql = sql & "  From LEGAJOS "
            sql = sql & "  WHERE  ID_Personal = " & ctlPersonal.Valor
            sql = sql & " AND FECHA_ACTUALIZACION > " & FechaServerTipo(fecha)
            sql = sql & " AND FECHA_ACTUALIZACION < " & FechaServerTipo(DateAdd("d", 1, fecha))
            sql = sql & "  ORDER BY ID_PERSONAL, FECHA_ACTUALIZACION "
            frmReportes.ImprimirReporte PasoReportes & "rptControlCargaLegajos.rpt", sql, True
            MousePointer = 0
        Else
            MsgBox "Ingrese el personal", vbInformation
        End If
End Sub

Private Sub ControlCargaCliente()
Dim sql As String
     Dim fecha As String
     
     If Not IsNull(ctlPersonal.Valor) Then
            fecha = InputBox("Ingrese la Fecha de Control ", "Control Carga", Format(Now, "DD/mm/yyyy"))
            MousePointer = 11
            sql = " SELECT * "
            sql = sql & "  From LEGAJOS_LARGO "
            sql = sql & "  WHERE  ID_Personal = " & ctlPersonal.Valor
            sql = sql & " and cod_cliente = " & ctlCliente.Valor
            sql = sql & " and FECHA_ACTUALIZACION > " & FechaServerTipo(fecha)
            sql = sql & " AND FECHA_ACTUALIZACION < " & FechaServerTipo(DateAdd("d", 1, fecha))
            sql = sql & "  ORDER BY ID_PERSONAL, FECHA_ACTUALIZACION "
            sql = " SELECT * "
            sql = sql & "  From LEGAJOS_LARGO "
            sql = sql & "  WHERE  ID_Personal = " & ctlPersonal.Valor
            sql = sql & " and cod_cliente = " & ctlCliente.Valor
            sql = sql & " and FECHA_ACTUALIZACION > " & FechaServerTipo("04/05/2006")
            sql = sql & " AND FECHA_ACTUALIZACION < " & FechaServerTipo("01/06/2006")
            sql = sql & "  ORDER BY ID_PERSONAL, FECHA_ACTUALIZACION "
            
            frmReportes.ImprimirReporte PasoReportes & "rptControlCargaLegajos.rpt", sql, True
            MousePointer = 0
    Else
        MsgBox "Ingrese el personal", vbInformation
    End If
End Sub

Private Sub ControlCargaEstadistico()
Dim sql As String
    Dim FechaInicio As String
    Dim FechaFin As String
    Dim ClienteLegajo As String
        If Not IsNull(ctlPersonal.Valor) Then
            FechaInicio = InputBox("Ingrese la Fecha Inicio de Control ", "Control Carga", DateAdd("d", -7, Format(Now, "DD/mm/yyyy")))
            FechaFin = InputBox("Ingrese la Fecha de Control Fin  ", "Control Carga", Format(Now, "DD/mm/yyyy"))
            ClienteLegajo = InputBox("INGRESE EL CLIENTE  " & " SI el cliente es 0 incluye todos los clientes ", "Control Carga", 0)
            MousePointer = 11
            sql = " SELECT * "
            sql = sql & "  From LEGAJOS "
            sql = sql & "  WHERE  ID_Personal = " & ctlPersonal.Valor
            sql = sql & " and FECHA_ACTUALIZACION > " & FechaServerTipo(FechaInicio)
            sql = sql & " AND FECHA_ACTUALIZACION < " & FechaServerTipo(FechaFin)
            If ClienteLegajo <> "0" Then
                   sql = sql & " AND COD_CLIENTE=" & ClienteLegajo
            End If
            sql = sql & "  ORDER BY ID_PERSONAL, FECHA_ACTUALIZACION "
            frmReportes.ImprimirReporte PasoReportes & "rptControlCargaLegajosEstadistica.rpt", sql, True
            MousePointer = 0
        Else
            MsgBox "Ingrese el personal", vbInformation
        End If
End Sub

Private Sub ImprimirCOntrolLegajos()
        Dim sql As String
  If Not IsNull(ctlCliente.Valor) And txtCaja.Text <> "" Then
        sql = " SELECT *"
        sql = sql & "  From PERSONAL, LEGAJOS "
        sql = sql & "  Where PERSONAL.IDPERSONAL = LEGAJOS.ID_PERSONAL"
        sql = sql & "  AND NRO_CAJA = " & txtCaja
        sql = sql & "  AND COD_CLIENTE = " & ctlCliente.Valor
        sql = sql & "  ORDER BY ID_CLIENTE_LEGAJO "
        frmReportes.ImprimirReporte PasoReportes & "rptLegajosControl.rpt", sql, True
   Else
       MsgBox "Ingrese la caja y el cliente ", vbInformation
   End If
   
End Sub

Private Sub cmdComprobante_Click()

End Sub

Private Sub cmdModificar_Click()
 
    
End Sub

Private Sub cmdNuevo_Click()
    
End Sub

Private Sub chkDescripcion_Click()
If chkDescripcion.value = 1 Then

    txtDescripcion.Enabled = True
Else
txtDescripcion.Enabled = False


End If

End Sub

Private Sub chkNombre_Click()
 If chkNombre.value = 1 Then
    txtNombre.Enabled = True
 Else
    txtNombre.Enabled = False
 End If
 
End Sub

Private Sub Command1_Click()
    Dim rs As New ADODB.Recordset
    Dim RsSql As New ADODB.Recordset
    Dim sql As String
    Dim CANTIDAD_CARACTERES As Integer
    Dim ConSql As New ADODB.Connection
    Dim s_ConSql As String
    
    s_ConSql = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Basa;Data Source=server001"
    ConSql.Open s_ConSql
        
        sql = "  SELECT ID_PERSONAL, FECHA_ACTUALIZACION,Descripcion, "
        sql = sql & " CANTIDAD_CARACTERES, COD_CLIENTE,"
        sql = sql & "ID_CLIENTE_LEGAJO , CLIENTE_LEGAJO, Nombre"
        sql = sql & " From LEGAJOS"
        sql = sql & " WHERE FECHA_ACTUALIZACION > " & FechaSegundoServerTipo("14/09/2006 16:30:00")
         sql = sql & " AND(CANTIDAD_CARACTERES IS NULL) AND"
        sql = sql & " (NOT (CLIENTE_LEGAJO IS NULL))"
        sql = sql & "ORDER BY FECHA_ACTUALIZACION"
        RsSql.Open sql, ConActiva, 0, 1
    
    Do While Not RsSql.EOF
    CANTIDAD_CARACTERES = 0
    Select Case RsSql!COD_CLIENTE
    Case 4
        CANTIDAD_CARACTERES = Len(Trim(RsSql!ID_CLIENTE_LEGAJO)) + Len(Trim(RsSql!CLIENTE_LEGAJO))
    Case Else
        CANTIDAD_CARACTERES = Len(Trim(RsSql!ID_CLIENTE_LEGAJO)) + Len(Trim(RsSql!CLIENTE_LEGAJO))
        If Not IsNull(RsSql!Descripcion) Then
            CANTIDAD_CARACTERES = CANTIDAD_CARACTERES + Len(Trim(RsSql!Descripcion))
        End If
        If Not IsNull(RsSql!Nombre) Then
            CANTIDAD_CARACTERES = CANTIDAD_CARACTERES + Len(Trim(RsSql!Nombre))
        End If
    End Select
        sql = " Update LEGAJOS"
        sql = sql & " SET CANTIDAD_CARACTERES =" & CANTIDAD_CARACTERES
       Rem sql = sql & ", ID_PERSONAL =" & rsSql!ID_Personal
        sql = sql & " WHERE ID_CLIENTE_LEGAJO = " & RsSql!ID_CLIENTE_LEGAJO
        sql = sql & "  AND COD_CLIENTE =" & RsSql!COD_CLIENTE
        ExecutarSql sql
        RsSql.MoveNext
    Loop
    
    
    
    
    sql = " SELECT ID_PERSONAL, COD_CLIENTE, FECHA_ACTUALIZACION, "
    sql = sql & " CANTIDAD_CARACTERES, ID_LEGAJO, CLIENTE_LEGAJO,"
    sql = sql & " ID_CLIENTE_LEGAJO , Nombre, Descripcion"
    sql = sql & "  From LEGAJOS "
    sql = sql & "  WHERE FECHA_ACTUALIZACION > " & FechaServerTipo("14/07/2006")
    rs.Open sql, ConActiva, 0, 1
    
    
     
    
    
    
End Sub

Private Sub Command2_Click()
Dim rs As New ADODB.Recordset
rs.CursorLocation = adUseClient
Dim sql As String
Dim Legajo As String

sql = " SELECT      ID_LEGAJO, ID_CLIENTE_LEGAJO, CLIENTE_LEGAJO, NUMERO_LEGAJO_CLIENTE"
sql = sql & " From LEGAJOS Where (Not (CLIENTE_LEGAJO Is Null))"
sql = sql & "  ORDER BY ID_LEGAJO"

sql = " SELECT     ID_CLIENTE_LEGAJO, COD_INDICE, CLIENTE_LEGAJO, COD_CLIENTE"
sql = sql & " From LEGAJOS"
sql = sql & " WHERE     (COD_CLIENTE = 49) AND (CLIENTE_LEGAJO LIKE '%_%') AND (COD_INDICE = '004001002001')"


rs.Open sql, ConActiva, adOpenKeyset, adLockOptimistic


Do While Not rs.EOF
  Debug.Print rs!CLIENTE_LEGAJO
Legajo = Replace(rs!CLIENTE_LEGAJO, ".-", "")
Legajo = Replace(Legajo, "_", "")

Legajo = Replace(Legajo, "_", "")

Legajo = Replace(Legajo, " ", "")
Legajo = Replace(Legajo, ".", "")
rs!CLIENTE_LEGAJO = UCase(Trim(Legajo))
Debug.Print rs!CLIENTE_LEGAJO
On Error GoTo salir

If rs!CLIENTE_LEGAJO <> "102-" Then
If IsNumeric(rs!CLIENTE_LEGAJO) Then

rs!NUMERO_LEGAJO_CLIENTE = rs!CLIENTE_LEGAJO
End If
End If
salir:

Rem   Debug.Print rs!ID_LEGAJO
rs.Update
rs.MoveNext
Loop


End Sub

Private Sub Command3_Click()
Dim rs As New ADODB.Recordset
rs.CursorLocation = adUseClient
Dim sql As String
Dim Legajo As String


sql = " SELECT     ID_CLIENTE_LEGAJO, NUMERO_LEGAJO_CLIENTE, COD_INDICE, CLIENTE_LEGAJO, COD_CLIENTE"
sql = sql & " From LEGAJOS"
sql = sql & " WHERE     (COD_CLIENTE = 101) "


rs.Open sql, ConActiva, adOpenKeyset, adLockOptimistic


Do While Not rs.EOF
 
 If IsNumeric(rs!CLIENTE_LEGAJO) Then
    rs!NUMERO_LEGAJO_CLIENTE = CLng(rs!CLIENTE_LEGAJO)
 
 Else
    rs!NUMERO_LEGAJO_CLIENTE = 0
 End If
 
rs.Update
rs.MoveNext
Loop
End Sub

Private Sub Command4_Click()


 Dim rs As New ADODB.Recordset
 Dim sql As String
 
 rs.Open " SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_CLIENTE From idlegajos$ ", ConActiva, 0, 1
 
 Do While Not rs.EOF
 
 sql = " Update LEGAJOS"
sql = sql & " Set ID_LEGAJO =" & rs!id_legajo
sql = sql & "  Where ID_CLIENTE_LEGAJO =" & rs!ID_CLIENTE_LEGAJO
sql = sql & "  And Cod_cliente = " & rs!COD_CLIENTE
 ExecutarSql sql
    rs.MoveNext
Loop

End Sub

Private Sub ctlCliente_Click()
        ctlIndice.Actualizar ctlCliente.Valor, Nulo, 0
      On Error GoTo salir:
        If ctlCliente.Valor = 4 Then
            Set CONlEGAJOS = New ADODB.Connection
            CONlEGAJOS.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ClienteEcogas & ";Persist Security Info=False"
        End If
        If ctlCliente.Valor = 20 Then
            Set CONlEGAJOS = New ADODB.Connection
            CONlEGAJOS.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ClienteOsep & " ;Persist Security Info=False"
        End If
salir:
        
        lblCodigo_Indice.Caption = ""
        lblCodigo_Nombre.Caption = ""
End Sub

Private Sub ctlIndice_DblClick()
Dim rsLegajo As ADODB.Recordset
Set rsLegajo = New ADODB.Recordset
Dim Codigo As String
Dim sql As String
Codigo = ctlIndice.Item_Selecionado

sql = " SELECT Expediente, MASK_EXPEDIENTE,  TOOLTIPEXPEDIENTE , TIPO_INDICE"
sql = sql & " From INDICES "
sql = sql & "  WHERE (COD_CLIENTE =" & ctlCliente.Valor & ") "
sql = sql & "  AND (INDICE = '" & Codigo & "')"

rsLegajo.Open sql, ConActiva, 0, 1
    mskLegajo_Cliente.Mask = ""
    mskLegajo_Cliente.ToolTipText = ""
    mskLegajo_Cliente.Text = ""
If Not rsLegajo.EOF Then
    If Trim(rsLegajo!Tipo_Indice) <> "Legajo" Then
        lblCodigo_Indice.Caption = ""
        lblCodigo_Nombre.Caption = ""
        mskLegajo_Cliente.Mask = ""
        mskLegajo_Cliente.ToolTipText = ""
        mskLegajo_Cliente.Text = ""
        MsgBox "El tipo de documento NO es un legajo", vbCritical
        Exit Sub
    End If
lblCodigo_Indice.Caption = Codigo
lblCodigo_Nombre.Caption = ctlIndice.Descripcion
    If Not IsNull(rsLegajo!EXPEDIENTE) Then
        If Not IsNull(rsLegajo!MASK_EXPEDIENTE) Then
        mskLegajo_Cliente.Mask = rsLegajo!MASK_EXPEDIENTE
        If IsNull(rsLegajo!TOOLTIPEXPEDIENTE) Then
              mskLegajo_Cliente.ToolTipText = ""
        Else
              mskLegajo_Cliente.ToolTipText = rsLegajo!TOOLTIPEXPEDIENTE
        End If
        End If
    End If
Else
    MsgBox "No se encontro marca de legajo"
End If
End Sub

Private Sub ctlIndice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = 2 Then

          PopupMenu mnuArbol
    End If


End Sub

Private Sub Form_Load()
ctlCliente.TipoControl = Cliente
ctlPersonal.TipoControl = PERSONAL
txtNombre.Enabled = False
txtDescripcion.Enabled = False
End Sub

Private Sub grdCantidad_Click()

End Sub

Private Sub lblCodigo_Indice_Change()
Dim rs As New ADODB.Recordset
Dim sql As String
        lblCodigo_Nombre.Caption = ""
        If lblCodigo_Indice.Caption = "" Then
            Exit Sub
        End If
        
        sql = " SELECT     TIPO_INDICE From INDICES"
        sql = sql & " WHERE COD_CLIENTE = " & ctlCliente.Valor
        sql = sql & "  AND indice = '" & lblCodigo_Indice.Caption & "'"

        rs.Open sql, ConActiva, 0, 1
        If Not rs.EOF Then
            If Trim(rs!Tipo_Indice) <> "Legajo" Then
                MsgBox "Asignacion Incorrecta", vbCritical
                lblCodigo_Indice.Caption = ""
                txtDocumento.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "El documento No existe"
            lblCodigo_Indice.Caption = ""
            txtDocumento.SetFocus
            Exit Sub
        End If
   lblCodigo_Nombre.Caption = BuscarIndiceDescripcion(lblCodigo_Indice.Caption, ctlCliente.Valor)
    mskLegajo_Cliente.Mask = BuscarMaskExpediente(lblCodigo_Indice.Caption, ctlCliente.Valor)
End Sub

Private Sub mnuBuscarIndice_Click()

ctlIndice.BuscarIndice InputBox("Ingrese el indice"), True
End Sub

Private Sub mnuBuscarLegajos_Click()
 ctlIndice.BuscarTipoIndice "Legajo", True
End Sub

Private Sub mnuRefrescar_Click()
    ctlIndice.Actualizar ctlCliente.Valor, Nulo, 0
End Sub

Private Sub mskLegajo_Cliente_GotFocus()
    mskLegajo_Cliente.SelStart = 0
End Sub

Private Sub mskLegajo_Cliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    If ctlCliente.Valor = 4 And lblCodigo_Indice.Caption = "001003001001001001" Then
        BuscarEcogas mskLegajo_Cliente.ClipText
    End If
'    If ctlCliente.Valor = 20 And lblCodigo_Indice.Caption = "002008002" Then
        BuscarOsep mskLegajo_Cliente.ClipText
'    End If
If txtDescripcion.Enabled = False Then
    Select Case lblEstado.Caption
    Case "Nuevo"
        Nuevo
    Case "Modificar"
        Modificar
    Case Else
        MsgBox "Error en el estado", vbInformation
    End Select

Else
    SendKeys vbTab
End If



 End If
End Sub

Private Sub mskLegajo_Cliente_LostFocus()
 Dim rs As ADODB.Recordset
 Set rs = New ADODB.Recordset
 Dim sql As String
 Dim n As Double
  If IsNumeric(mskLegajo_Cliente.Text) Then
    n = CDbl(mskLegajo_Cliente.ClipText)
    sql = " SELECT ORDENAR_DOCUMENTACION_DETALLE.Cod_Estado as estado"
    sql = sql & vbCrLf & " From ORDENAR_DOCUMENTACION_DETALLE, ORDENAR_DOCUMENTACION"
    sql = sql & vbCrLf & "  Where ORDENAR_DOCUMENTACION_DETALLE.COD_DOCUMENTACION = ORDENAR_DOCUMENTACION.ID_ORDENAR_DOCUMENTACION"
    sql = sql & vbCrLf & "  AND (ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = " & ctlCliente.Valor & ")"
    sql = sql & vbCrLf & " AND (ORDENAR_DOCUMENTACION_DETALLE.COD_INDICE = '" & lblCodigo_Indice.Caption & "')"
    sql = sql & vbCrLf & "  AND(ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO = '" & n & "')"
  Else
    sql = " SELECT   ORDENAR_DOCUMENTACION_DETALLE.Cod_Estado as estado "
    sql = sql & vbCrLf & "  From ORDENAR_DOCUMENTACION_DETALLE, ORDENAR_DOCUMENTACION"
    sql = sql & vbCrLf & "  Where ORDENAR_DOCUMENTACION_DETALLE.COD_DOCUMENTACION = ORDENAR_DOCUMENTACION.ID_ORDENAR_DOCUMENTACION"
    sql = sql & vbCrLf & "  AND (ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = " & ctlCliente.Valor & ")"
    sql = sql & vbCrLf & "  AND (ORDENAR_DOCUMENTACION_DETALLE.COD_INDICE = '" & lblCodigo_Indice.Caption & "')"
    sql = sql & vbCrLf & "  AND(ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO = '" & mskLegajo_Cliente.ClipText & "')"
  End If
Rem Rs.Open Sql, strConBasa , 0 ,1
' If Not Rs.EOF Then
'
'    MsgBox "Legajos para archivar"
' End If
 
 
End Sub

Private Sub mskRemitoProv_GotFocus()
    mskRemitoProv.SelStart = 8
End Sub

Private Sub mskRemitoProv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub trvIndices_DblClick()


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Caption
Case "Aceptar F9"
'    Select Case lblEstado.Caption
'    Case "Nuevo"
'        Nuevo
'    Case "Modificar"
'        Modificar
'    Case Else
'        MsgBox "Error en el estado", vbInformation
'    End Select
Case "Nuevo"
    lblEstado.Caption = "Nuevo"
    txtID_Cliente_Legajos.Enabled = True
    txtID_Cliente_Legajos.Text = ""
    LimpiarMask mskLegajo_Cliente
    txtDocumento.Text = ""
    Rem lblCodigo_Indice.Caption = ""
    Rem lblCodigo_Nombre.Caption = ""
    txtDescripcion.Text = ""
    txtNombre.Text = ""
Case "Grilla"
    ActualizarGrilla
Case "Modificar"
        lblEstado.Caption = "Modificar"
        
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        Dim sql As String
           If IsNull(ctlCliente.Valor) Then
               MsgBox "Ingresar el cliente"
               Exit Sub
           End If
             
           If txtID_Cliente_Legajos.Text = "" Then
               MsgBox "Ingresar el Rotulo"
               Exit Sub
           End If
           
           sql = "  SELECT ID_CLIENTE_LEGAJO,COD_INDICE, CLIENTE_LEGAJO, DESCRIPCION,"
           sql = sql & vbCrLf & "  Nombre , COD_CLIENTE, NRO_CAJA"
           sql = sql & vbCrLf & "  From LEGAJOS"
           sql = sql & vbCrLf & " Where COD_CLIENTE = " & ctlCliente.Valor
           sql = sql & vbCrLf & " And ID_CLIENTE_LEGAJO =" & txtID_Cliente_Legajos.Text
           
        rs.Open sql, ConActiva, 0, 1
        If Not rs.EOF Then
        
        If IsNull(rs!Cod_Indice) Then
           MsgBox "El legajo no esta cargado", vbCritical
           Exit Sub
        End If
               lblCodigo_Indice.Caption = rs!Cod_Indice
               LimpiarMask mskLegajo_Cliente
               mskLegajo_Cliente.Mask = ""
               mskLegajo_Cliente.Text = rs!CLIENTE_LEGAJO
         If Not IsNull(rs!Nombre) Then
               txtNombre.Text = Trim(rs!Nombre)
           End If
           If Not IsNull(rs!Descripcion) Then
               txtDescripcion.Text = Trim(rs!Descripcion)
           End If
           txtCaja.Text = rs!NRO_CAJA
         End If
         txtID_Cliente_Legajos.Enabled = False
    
End Select



End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Debug.Print ButtonMenu.Text
Select Case ButtonMenu.Text
Case "Control Carga por cliente"
    ControlCargaCliente
Case "Control Carga"
    ControlCaracteres
Case "Control carga estadistico"
    ControlCargaEstadistico
Case "Imprimir comprobante"
    ImprimirComprobante
Case "Imprimir Control"
    ImprimirCOntrolLegajos

End Select

End Sub

Private Sub txtCaja_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtCaja_LostFocus()
If lblEstado.Caption = "" Then
    MsgBox "Verifique el estado", vbCritical
End If
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Select Case lblEstado.Caption
    Case "Nuevo"
        Nuevo
    Case "Modificar"
        Modificar
    Case Else
        MsgBox "Error en el estado", vbInformation
    End Select
        
    End If
End Sub

Private Sub txtDocumento_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
     If IsNumeric(txtDocumento.Text) Then
        lblCodigo_Indice = BuscarIDDocumento(txtDocumento.Text, ctlCliente.Valor)
        SendKeys vbTab
    End If
End If

End Sub

Private Sub txtID_Cliente_Legajos_Change()
Dim largo As Integer
Dim cantidad As Long
largo = Len(txtID_Cliente_Legajos.Text)
    If largo > 3 Then
        If Mid(txtID_Cliente_Legajos.Text, largo) = "0" Then
            txtID_Cliente_Legajos.BackColor = &H80000013
        Else
            txtID_Cliente_Legajos.BackColor = &H80000005
        End If
            If Trim(txtCantidadElementos.Text) <> "" Then
               If lblInicial.Caption <> "" Then
                   cantidad = txtID_Cliente_Legajos.Text - lblInicial.Caption
                   If cantidad = txtCantidadElementos.Text Then
                       MsgBox "Usted llego a la cantidad de :" & cantidad
                   End If
               End If
     End If

    End If
End Sub

Private Sub txtID_Cliente_Legajos_DblClick()
lblInicial.Caption = txtID_Cliente_Legajos.Text
End Sub

Private Sub txtID_Cliente_Legajos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Trim(txtID_Cliente_Legajos.Text) <> "" Then
        If UCase(Mid(txtID_Cliente_Legajos.Text, 1, 2)) = "L2" Then
            txtID_Cliente_Legajos.Text = CLng(Mid(txtID_Cliente_Legajos.Text, 3))
        End If
        
        SendKeys vbTab
    End If
    
    If KeyAscii = 43 And Trim(txtID_Cliente_Legajos.Text) <> "" Then
        
        txtID_Cliente_Legajos.Text = CLng(txtID_Cliente_Legajos.Text) + 1
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        SendKeys vbTab
        KeyAscii = 0
    End If
End Sub

Private Sub txtResponsable_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Public Sub Modificar()

Dim sql As String
Dim NUMERO_LEGAJO_CLIENTE As Long
Dim CANTIDAD_CARACTERES As Integer
Dim CantidadRegistro As Integer
If IsNull(ctlPersonal.Valor) Then
    MsgBox "Se debe ingresar el responsable"
    Exit Sub
End If


If mskLegajo_Cliente.Text = "" Then
    MsgBox "Se debe ingresar el legajo"
    Exit Sub
End If

If IsNumeric(mskLegajo_Cliente) Then
    NUMERO_LEGAJO_CLIENTE = CLng(mskLegajo_Cliente)
Else
    NUMERO_LEGAJO_CLIENTE = 0
End If
    Select Case ctlCliente.Valor
    Case 4
        CANTIDAD_CARACTERES = Len(txtID_Cliente_Legajos.Text) + Len(mskLegajo_Cliente.Text)
    Case Else
        CANTIDAD_CARACTERES = Len(Trim(txtID_Cliente_Legajos.Text)) + Len(Trim(mskLegajo_Cliente.Text)) + Len(Trim(txtNombre.Text)) + Len(Trim(txtDescripcion.Text))
    End Select

    sql = " Update LEGAJOS "
    sql = sql & vbCrLf & " SET CLIENTE_LEGAJO = '" & mskLegajo_Cliente.Text & "', DESCRIPCION = '" & Trim(txtDescripcion.Text) & "'"
    sql = sql & vbCrLf & " ,  NOMBRE = '" & Trim(txtNombre.Text) & "', COD_INDICE = '" & lblCodigo_Indice.Caption & "'"
    sql = sql & vbCrLf & " , NRO_CAJA =" & txtCaja.Text & ", FECHA_ACTUALIZACION = " & SysDateMinutoSegundo
    sql = sql & vbCrLf & " , ID_PERSONAL = " & ctlPersonal.Valor & ",NUMERO_LEGAJO_CLIENTE = " & NUMERO_LEGAJO_CLIENTE
    sql = sql & vbCrLf & " ,CANTIDAD_CARACTERES =" & CANTIDAD_CARACTERES
    sql = sql & vbCrLf & "  Where COD_CLIENTE = " & ctlCliente.Valor
    sql = sql & vbCrLf & " And ID_CLIENTE_LEGAJO = " & txtID_Cliente_Legajos.Text
     CantidadRegistro = ExecutarSql(sql)
    If CantidadRegistro <> 1 Then
        MsgBox "La actualizacion no se realizo "
    End If
    
    
    txtID_Cliente_Legajos.Text = ""
    LimpiarMask mskLegajo_Cliente
    LimpiarMask mskRemitoProv
    txtDescripcion.Text = ""
    txtNombre.Text = ""
    txtCaja.Text = ""
    txtID_Cliente_Legajos.Text = ""
    txtID_Cliente_Legajos.Enabled = True
    txtID_Cliente_Legajos.SetFocus

End Sub

Public Sub BuscarEcogas(Pig As Long)
Dim rs As New ADODB.Recordset
Dim sql As String

sql = " SELECT num "
sql = sql & vbCrLf & " , [calle] & '  ' & [inmueble_puerta_num] & '  ' & "
sql = sql & vbCrLf & " [localidad_nombre] & ' ' & Pig.pcia_nombre AS descripcion, "
sql = sql & vbCrLf & " 'Bº ' & [barrio] & '  ' & [inmueble_torre_des] & '   ' & "
sql = sql & vbCrLf & " [inmueble_dpto_des] & '  ' & [inmueble_piso_des] AS Nombre"
sql = sql & vbCrLf & " From Pig  "
sql = sql & vbCrLf & " WHERE num = " & Pig
rs.Open sql, CONlEGAJOS
    If Not rs.EOF Then
        txtDescripcion.Text = rs!Descripcion
        txtNombre.Text = rs!Nombre
    Else
        txtDescripcion.Text = ""
        txtNombre.Text = ""
    End If
End Sub
Public Sub BuscarOsep(Doc As Long)
Dim rs As New ADODB.Recordset
Dim sql As String


sql = "  SELECT   NUMERO_AFI, APELLIDO_NOMBRE AS Nombre"
sql = sql & vbCrLf & " , 'Doc:  ' & [NUMERO_AFI] AS Descripcion"
sql = sql & vbCrLf & "  From OSEPAFILI "
sql = sql & vbCrLf & "  WHERE NUMERO_AFI = " & Doc
rs.Open sql, CONlEGAJOS
    If Not rs.EOF Then
        txtDescripcion.Text = rs!Descripcion
        txtNombre.Text = rs!Nombre
    Else
        txtDescripcion.Text = ""
        txtNombre.Text = ""
    End If
End Sub



Public Sub ActualizarGrilla()
 Dim rs As New ADODB.Recordset
 Dim cantidad As Integer
 
 cantidad = InputBox("Ingrese la cantidad de Registros", "", 50)
 Dim sql As String
            sql = " SELECT     TOP " & cantidad & " ID_CLIENTE_LEGAJO,NRO_CAJA, CLIENTE_LEGAJO, NOMBRE, DESCRIPCION, FECHA_ACTUALIZACION"
            sql = sql & vbCrLf & " From LEGAJOS"
            sql = sql & vbCrLf & " Where ID_Personal = " & ctlPersonal.Valor
            sql = sql & vbCrLf & " and FECHA_ACTUALIZACION > '" & Format(Now, "DD/mm/yyyy") & "'"
            sql = sql & vbCrLf & " ORDER BY FECHA_ACTUALIZACION DESC"
            rs.CursorLocation = adUseClient
            
            rs.Open sql, ConActiva, adOpenDynamic, adLockOptimistic
            
            Set grdLegajos.DataSource = rs
            grdLegajos.DataMember = rs.DataMember
    grdLegajos.Refresh
    grdLegajos.Columns(0).Locked = True
    grdLegajos.Columns(5).Locked = True
    
    
End Sub
