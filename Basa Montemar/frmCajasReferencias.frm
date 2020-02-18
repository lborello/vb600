VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajasReferencias 
   Caption         =   "Referencias Cajas"
   ClientHeight    =   8865
   ClientLeft      =   1230
   ClientTop       =   240
   ClientWidth     =   13650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   13650
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picCargar 
      Height          =   3870
      Left            =   5520
      ScaleHeight     =   3810
      ScaleWidth      =   7695
      TabIndex        =   69
      Top             =   840
      Width           =   7755
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Acep."
         Height          =   315
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   615
      End
      Begin VB.CheckBox chkCaja 
         Alignment       =   1  'Right Justify
         Caption         =   "Caja"
         Height          =   315
         Left            =   60
         TabIndex        =   80
         Top             =   60
         Width           =   615
      End
      Begin VB.TextBox txtDescripcion 
         DataField       =   "DESCRIPCION"
         DataMember      =   "Command1"
         Height          =   315
         Left            =   1320
         MaxLength       =   2000
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox chkIndice 
         Alignment       =   1  'Right Justify
         Caption         =   "Indice"
         Height          =   315
         Left            =   2520
         TabIndex        =   79
         Top             =   60
         Width           =   840
      End
      Begin VB.TextBox txtIndice 
         DataField       =   "INDICE"
         Height          =   315
         Left            =   3420
         TabIndex        =   1
         Top             =   60
         Width           =   1470
      End
      Begin VB.TextBox txtNro_Caja 
         DataField       =   "NRO_CAJA"
         DataMember      =   "Command1"
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         ToolTipText     =   "Buscar Indice F12"
         Top             =   60
         Width           =   1110
      End
      Begin VB.CheckBox chkDescripcion 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   315
         Left            =   60
         TabIndex        =   78
         Top             =   960
         Width           =   1155
      End
      Begin VB.CheckBox chkFecha_Desde 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Desde"
         Height          =   255
         Left            =   4860
         TabIndex        =   75
         Top             =   420
         Width           =   1455
      End
      Begin VB.CheckBox chkFechaHasta 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Hasta"
         Height          =   255
         Left            =   4860
         TabIndex        =   74
         Top             =   780
         Width           =   1455
      End
      Begin VB.CheckBox chkNº_Hasta 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Hasta"
         Height          =   255
         Left            =   5220
         TabIndex        =   73
         Top             =   1500
         Width           =   1035
      End
      Begin VB.CheckBox chkNº_Desde 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Desde"
         Height          =   255
         Left            =   5220
         TabIndex        =   72
         Top             =   1140
         Width           =   1035
      End
      Begin VB.CheckBox chkLetra_Desde 
         Alignment       =   1  'Right Justify
         Caption         =   "Letra Desde"
         Height          =   255
         Left            =   60
         TabIndex        =   71
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CheckBox chkLetra_Hasta 
         Alignment       =   1  'Right Justify
         Caption         =   "Letra Hasta"
         Height          =   255
         Left            =   60
         TabIndex        =   70
         Top             =   1920
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid grdDescripcionRepetida 
         Height          =   1995
         Left            =   360
         TabIndex        =   76
         Top             =   2400
         Visible         =   0   'False
         Width           =   12675
         _ExtentX        =   22357
         _ExtentY        =   3519
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid grdModificacion 
         Height          =   1515
         Left            =   240
         TabIndex        =   77
         Top             =   2340
         Width           =   12675
         _ExtentX        =   22357
         _ExtentY        =   2672
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
      Begin MSMask.MaskEdBox mskFecha_Hasta 
         Bindings        =   "frmCajasReferencias.frx":0000
         DataField       =   "FECHA_HASTA"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   3
         EndProperty
         DataMember      =   "Command1"
         DataSource      =   "DataEnvironment1"
         Height          =   315
         Left            =   6360
         TabIndex        =   3
         Top             =   780
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskLetra_Desde 
         DataField       =   "Letra_Desde"
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   1500
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
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
      Begin MSMask.MaskEdBox mskFecha_Desde 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   6360
         TabIndex        =   2
         Top             =   360
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskLetra_Hasta 
         DataField       =   "Letra_Hasta"
         Height          =   315
         Left            =   1380
         TabIndex        =   8
         Top             =   1860
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox mskNro_hasta 
         Height          =   315
         Left            =   6360
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskNro_desde 
         Height          =   315
         Left            =   6360
         TabIndex        =   5
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Index           =   13
         Left            =   60
         TabIndex        =   84
         Top             =   480
         Width           =   555
      End
      Begin VB.Label lblEstadoReferencia 
         Caption         =   "Label11"
         Height          =   15
         Left            =   60
         TabIndex        =   83
         Top             =   2580
         Width           =   135
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   82
         Top             =   420
         Width           =   3255
      End
      Begin VB.Label lblCod_Referencia 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   5280
         TabIndex        =   81
         Top             =   60
         Width           =   855
      End
   End
   Begin VB.TextBox txtReferenciaLote 
      Height          =   315
      Left            =   6600
      TabIndex        =   63
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdTerminada 
      Caption         =   "Hoja Terminada"
      Height          =   315
      Left            =   8340
      TabIndex        =   62
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdProxReferencia 
      Caption         =   ">>"
      Height          =   315
      Left            =   7860
      TabIndex        =   61
      Top             =   120
      Width           =   435
   End
   Begin VB.CommandButton cmdImagenUna 
      Height          =   315
      Left            =   7620
      TabIndex        =   60
      Top             =   120
      Width           =   195
   End
   Begin VB.CommandButton cmdBuscarIndice 
      Caption         =   "F10"
      Height          =   255
      Left            =   4980
      TabIndex        =   59
      ToolTipText     =   "Buscar Indice"
      Top             =   1140
      Width           =   495
   End
   Begin VB.CommandButton cmdControlReferencias 
      Caption         =   ">>"
      Height          =   375
      Left            =   12480
      TabIndex        =   57
      Top             =   360
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.CommandButton cmdControlCarga 
      Caption         =   "Control Carga"
      Height          =   375
      Left            =   11280
      TabIndex        =   56
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdBorrarReferencia 
      Caption         =   "Borar Lote"
      Height          =   435
      Left            =   12960
      TabIndex        =   53
      Top             =   360
      Width           =   615
   End
   Begin Controles.cltIndice cltIndice1 
      Height          =   3255
      Left            =   120
      TabIndex        =   52
      Top             =   1500
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5741
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
   Begin TabDlg.SSTab sstReferencia 
      Height          =   4275
      Left            =   180
      TabIndex        =   24
      Top             =   4920
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   7541
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Carga"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picFiltro"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Buscar"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picBuscar"
      Tab(1).ControlCount=   1
      Begin VB.PictureBox picBuscar 
         Height          =   3255
         Left            =   -74880
         ScaleHeight     =   3195
         ScaleWidth      =   12675
         TabIndex        =   25
         Top             =   840
         Width           =   12735
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   315
            Left            =   8700
            TabIndex        =   85
            Top             =   780
            Width           =   855
         End
         Begin MSDataGridLib.DataGrid grdReferencias 
            Height          =   2055
            Left            =   120
            TabIndex        =   37
            Top             =   1080
            Width           =   12405
            _ExtentX        =   21881
            _ExtentY        =   3625
            _Version        =   393216
            AllowUpdate     =   -1  'True
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   2
            WrapCellPointer =   -1  'True
            FormatLocked    =   -1  'True
            AllowDelete     =   -1  'True
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
            ColumnCount     =   13
            BeginProperty Column00 
               DataField       =   "INDICE"
               Caption         =   "Indice"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "NRO_CAJA"
               Caption         =   "Caja"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column02 
               DataField       =   "ITEM"
               Caption         =   "Item"
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
               DataField       =   "DESCRIPCION"
               Caption         =   "Descripción"
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
               DataField       =   "FECHA_DESDE"
               Caption         =   "F. Desde"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "FECHA_HASTA"
               Caption         =   "F. Hasta"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   3
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "NRO_DESDE"
               Caption         =   "N. Desde"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "NRO_HASTA"
               Caption         =   "N. Hasta"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11274
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column08 
               DataField       =   "LETRA_DESDE"
               Caption         =   "L. Desde"
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
            BeginProperty Column09 
               DataField       =   "LETRA_HASTA"
               Caption         =   "L. Hasta"
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
            BeginProperty Column10 
               DataField       =   "EXPEDIENTE"
               Caption         =   "Expediente"
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
            BeginProperty Column11 
               DataField       =   "APELLIDO_NOMBRE"
               Caption         =   "Apellido Nombre"
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
            BeginProperty Column12 
               DataField       =   "COD_ID_REFERENCIA"
               Caption         =   "ID"
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
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column01 
                  ColumnAllowSizing=   0   'False
               EndProperty
               BeginProperty Column02 
                  ColumnAllowSizing=   0   'False
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column03 
                  Locked          =   -1  'True
                  Object.Visible         =   -1  'True
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
               BeginProperty Column06 
               EndProperty
               BeginProperty Column07 
               EndProperty
               BeginProperty Column08 
                  Object.Visible         =   -1  'True
               EndProperty
               BeginProperty Column09 
                  Object.Visible         =   -1  'True
               EndProperty
               BeginProperty Column10 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column11 
                  Object.Visible         =   0   'False
               EndProperty
               BeginProperty Column12 
                  Locked          =   -1  'True
                  Object.Visible         =   0   'False
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmdBuscarErrores 
            Caption         =   "Errores"
            Height          =   330
            Left            =   9000
            TabIndex        =   36
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox txtCodigo 
            Height          =   330
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   6195
         End
         Begin VB.TextBox txtdescripcionBuscar 
            Height          =   285
            Left            =   9720
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   540
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.CommandButton cmdCambioIndice 
            Caption         =   "Cam/Indice"
            Height          =   330
            Left            =   6480
            TabIndex        =   33
            Top             =   120
            Width           =   975
         End
         Begin VB.Frame Frame1 
            Caption         =   "Reemplazar"
            Height          =   915
            Left            =   9960
            TabIndex        =   27
            Top             =   120
            Width           =   2715
            Begin VB.CommandButton cmdReplazar 
               Caption         =   "..."
               Height          =   315
               Left            =   2280
               TabIndex        =   30
               Top             =   540
               Width           =   315
            End
            Begin VB.TextBox txtReplazo 
               Height          =   315
               Left            =   960
               TabIndex        =   29
               Top             =   540
               Width           =   1275
            End
            Begin VB.TextBox txtElementoExistente 
               Height          =   315
               Left            =   960
               TabIndex        =   28
               Top             =   180
               Width           =   1635
            End
            Begin VB.Label Label6 
               Caption         =   "Buscar"
               Height          =   255
               Left            =   60
               TabIndex        =   32
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "Reemplazar"
               Height          =   255
               Left            =   60
               TabIndex        =   31
               Top             =   540
               Width           =   915
            End
         End
         Begin VB.CommandButton cmdArbol 
            Caption         =   "Arbol"
            Height          =   330
            Left            =   8520
            TabIndex        =   26
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lblCodigo 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   570
            Left            =   120
            TabIndex        =   40
            Top             =   480
            Width           =   8415
         End
         Begin VB.Label lblCantidadRegistro 
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
            Height          =   315
            Left            =   8280
            TabIndex        =   39
            Top             =   120
            Width           =   675
         End
         Begin VB.Label Label8 
            Caption         =   "Cant."
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
            Left            =   7680
            TabIndex        =   38
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.PictureBox picFiltro 
         Height          =   3135
         Left            =   300
         ScaleHeight     =   3075
         ScaleWidth      =   3675
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   3735
         Begin VB.ComboBox cboCampoFiltro 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   420
            Width           =   2775
         End
         Begin VB.ComboBox cboFiltro 
            Height          =   315
            Left            =   840
            TabIndex        =   46
            Text            =   "Combo1"
            Top             =   780
            Width           =   2775
         End
         Begin VB.TextBox txtDato 
            Height          =   315
            Left            =   840
            TabIndex        =   45
            Top             =   1140
            Width           =   2775
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cerrar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2520
            TabIndex        =   44
            Top             =   60
            Width           =   1095
         End
         Begin VB.CommandButton cmdFiltro 
            Caption         =   "Filtro"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1260
            TabIndex        =   43
            Top             =   60
            Width           =   1095
         End
         Begin VB.CommandButton cmdFiltroBorrar 
            Caption         =   "Borrar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   42
            Top             =   60
            Width           =   1095
         End
         Begin VB.Label Label9 
            Caption         =   "Campo "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   51
            Top             =   420
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Criterio "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   50
            Top             =   780
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Dato"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   49
            Top             =   1140
            Width           =   975
         End
         Begin VB.Label lblFiltro 
            BorderStyle     =   1  'Fixed Single
            Height          =   1455
            Left            =   120
            TabIndex        =   48
            Top             =   1500
            Width           =   3555
         End
      End
   End
   Begin Controles.cltGenerico ctlPersonal 
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   661
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   375
      Left            =   720
      TabIndex        =   11
      Top             =   1080
      Width           =   4275
      _ExtentX        =   7964
      _ExtentY        =   661
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   8550
      Width           =   13650
      _ExtentX        =   24077
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Nuevo"
            TextSave        =   "Nuevo"
            Key             =   "EstadoAplicacion"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Ayuda"
            TextSave        =   "Ayuda"
            Key             =   "Ayuda"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picSectorRequerimiento 
      Height          =   2415
      Left            =   6360
      ScaleHeight     =   2355
      ScaleWidth      =   6975
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   7035
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   315
         Left            =   6000
         TabIndex        =   18
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox txtAgregarReferencia 
         Height          =   375
         Left            =   780
         TabIndex        =   17
         Top             =   1800
         Width           =   4275
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   315
         Left            =   5100
         TabIndex        =   16
         Top             =   1860
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid grdSector 
         Height          =   1395
         Left            =   60
         TabIndex        =   19
         Top             =   0
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   2461
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Sector Requerimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -2700
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   6555
      End
      Begin VB.Label Label4 
         Caption         =   "Sector:"
         Height          =   375
         Left            =   180
         TabIndex        =   20
         Top             =   1860
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   660
      Left            =   0
      TabIndex        =   13
      Top             =   60
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   1164
      ButtonWidth     =   1588
      ButtonHeight    =   1058
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            ImageKey        =   "Nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar F9"
            Key             =   "Aceptar"
            ImageKey        =   "Aceptar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Control"
            Key             =   "Control"
            ImageIndex      =   50
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cargar"
            Key             =   "Cargar"
            ImageKey        =   "Cargar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Borrar"
            ImageIndex      =   46
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            ImageKey        =   "Print"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   10
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PorCajaTodo"
                  Text            =   "Por Caja Todo"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PorCajaFiltro"
                  Text            =   "Por Caja Filtro"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PorIndiceTodo"
                  Text            =   "Por Indice Todo"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PorIndiceFiltro"
                  Text            =   "Por Indice Filtro"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Por indice con index"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Por indice filtro con index"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Doc solo Filtro"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Control referencia"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Control carga"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Cajas sin referencias"
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
         Left            =   5460
         TabIndex        =   14
         Top             =   0
         Width           =   1395
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   120
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
            Picture         =   "frmCajasReferencias.frx":0013
            Key             =   "Ver+"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0071
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":00CF
            Key             =   "Ver-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":012D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":018B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":01E9
            Key             =   "Punto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0247
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":02A5
            Key             =   "RotarI"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0303
            Key             =   "Vertical"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0361
            Key             =   "Sig"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":03BF
            Key             =   "Ant"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":041D
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":047B
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":04D9
            Key             =   "RotarD"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0537
            Key             =   "Cargar"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0595
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":05F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0651
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":06AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":070D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":076B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":07C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0827
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0885
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":08E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0941
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":099F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":09FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0A5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0AB9
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0B17
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0B75
            Key             =   "Fin"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0BD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0C31
            Key             =   "Aceptar"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0C8F
            Key             =   "Control"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0CED
            Key             =   "Esp. Fax"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0D4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0DA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0E07
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0E65
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0EC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0F21
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0F7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":0FDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":103B
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":1099
            Key             =   "Anular"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":10F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":1155
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":11B3
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":1211
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":126F
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":12CD
            Key             =   "Casa"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajasReferencias.frx":132B
            Key             =   "Bandera"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Pagina"
      Height          =   315
      Left            =   9780
      TabIndex        =   68
      Top             =   180
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "Lote:"
      Height          =   315
      Left            =   6180
      TabIndex        =   67
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lblId_Referencia 
      Height          =   315
      Left            =   12240
      TabIndex        =   66
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lbl_ID_imagen 
      Height          =   315
      Left            =   11220
      TabIndex        =   65
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblPasoImagen 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   10680
      TabIndex        =   64
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblImagenActual 
      Height          =   255
      Left            =   7680
      TabIndex        =   58
      Top             =   4980
      Width           =   615
   End
   Begin VB.Label lbImagenInicio 
      Height          =   255
      Left            =   6840
      TabIndex        =   55
      Top             =   4980
      Width           =   735
   End
   Begin VB.Label lblImagenFin 
      Height          =   255
      Left            =   8400
      TabIndex        =   54
      Top             =   4980
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   5
      X1              =   0
      X2              =   11400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   5
      X1              =   0
      X2              =   11400
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label13 
      Caption         =   "Usuario: "
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   780
      Width           =   615
   End
   Begin VB.Label Label12 
      Caption         =   "Cliente : "
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   615
   End
   Begin VB.Menu mnuArbol 
      Caption         =   "Arbol"
      Begin VB.Menu mnuBuscarReferencia 
         Caption         =   "BuscarReferencia"
         Begin VB.Menu mnuNumero 
            Caption         =   "Numero"
         End
         Begin VB.Menu mnunro_desde 
            Caption         =   "Nro_desde"
         End
         Begin VB.Menu mnuBuscarFecha 
            Caption         =   "Fecha"
            Begin VB.Menu mnuFechaBuscar 
               Caption         =   "Buscar"
            End
            Begin VB.Menu mnuFechaOrdenar 
               Caption         =   "Ordenar"
            End
            Begin VB.Menu mnuFechaDescripcion 
               Caption         =   "Descripcion"
            End
         End
         Begin VB.Menu mnuLetra 
            Caption         =   "Letra"
         End
         Begin VB.Menu mnuBuscarDescripcion 
            Caption         =   "Descripcion"
         End
         Begin VB.Menu mnuTodos 
            Caption         =   "Todos"
         End
         Begin VB.Menu mnuBuscarCaja 
            Caption         =   "Caja"
         End
         Begin VB.Menu mnuBuscarCajas 
            Caption         =   "Cajas"
         End
         Begin VB.Menu MNUNOMBREYAPELLIDO 
            Caption         =   "Nombre y Apellido"
         End
         Begin VB.Menu mnuFiltro 
            Caption         =   "Filtro"
         End
         Begin VB.Menu mnuBuscarindice 
            Caption         =   "Buscar Indice"
         End
         Begin VB.Menu mnuDescripcionIndiceFijo 
            Caption         =   "Descripcion Indice Fijo"
         End
         Begin VB.Menu mnuIDReferencia 
            Caption         =   "ID Referencia"
         End
         Begin VB.Menu mnuInconsistenciasNivel 
            Caption         =   "Inconsistencias Nivel"
         End
      End
      Begin VB.Menu mnuCampos 
         Caption         =   "Campos"
         Begin VB.Menu mnuHabilitarFecha 
            Caption         =   "Habilitar Fecha"
         End
         Begin VB.Menu mnuDeshabilitarfecha 
            Caption         =   "Deshabilitar Fecha"
         End
         Begin VB.Menu mnuHabilitarNumero 
            Caption         =   "Habilitar Numero"
         End
         Begin VB.Menu mnuDeshabilitarNumero 
            Caption         =   "Deshabilitar Numero"
         End
         Begin VB.Menu mnuHabilitarLetra 
            Caption         =   "Habilitar Letra"
         End
         Begin VB.Menu mnuDeshabilitarLetra 
            Caption         =   "Deshabilitar Letra"
         End
      End
      Begin VB.Menu mnuCantidadcajas 
         Caption         =   "Cantidad de Cajas"
      End
      Begin VB.Menu mnuResponsables 
         Caption         =   "Responsables"
      End
   End
   Begin VB.Menu MNUGRDBUSQUEDA 
      Caption         =   "GRDBUSQUEDA"
      Begin VB.Menu mnuCajasposicionTodo 
         Caption         =   "Cajas posicion Todo"
      End
      Begin VB.Menu MNUCAJAPOSICION 
         Caption         =   "Caja posición"
      End
      Begin VB.Menu mnutodalacaj 
         Caption         =   "Toda la caja"
      End
      Begin VB.Menu mnuVerRequerimientos 
         Caption         =   "Ver requerimientos"
      End
      Begin VB.Menu mnuCopiarDatos 
         Caption         =   "Copiar Datos"
      End
      Begin VB.Menu mnuVerImagen 
         Caption         =   "Ver Imagen"
      End
      Begin VB.Menu mnuModificar 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnuVerLote 
         Caption         =   "ver Lote"
      End
   End
   Begin VB.Menu mnugrdmodificar 
      Caption         =   "Modificar Datos"
      Visible         =   0   'False
      Begin VB.Menu mnuModificarDatos 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnuBorrarDatos 
         Caption         =   "Borrar"
      End
   End
   Begin VB.Menu mnuBorrarCajas 
      Caption         =   "Borrar Cajas"
   End
End
Attribute VB_Name = "frmCajasReferencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents rsReferencias As ADODB.Recordset
Attribute rsReferencias.VB_VarHelpID = -1
Dim Sql_Filtro_Referencia As String
Dim Filtro_Indice_Reporte As String
Dim Filtro_Indice As String
Dim Filtro_Reporte As String
Dim Sql_Filtro_Orden As String


Private Sub cboCampoFiltro_Click()
    cboFiltro.Clear
    If cboCampoFiltro.ListIndex = -1 Then
        Exit Sub
    End If
    Select Case cboCampoFiltro.ItemData(cboCampoFiltro.ListIndex)
    Case 0
        cboFiltro.AddItem ">"
        cboFiltro.AddItem "<"
        cboFiltro.AddItem "="
    Case 1
        cboFiltro.AddItem "="
        cboFiltro.AddItem "Not Like"
        cboFiltro.AddItem "Like"
    Case 2
        cboFiltro.AddItem "ENTRE FECHA"
    Case 4
        cboFiltro.AddItem "ENTRE NUMERO"
    End Select

End Sub

Private Sub cltIndice1_DblClick()
    TxtIndice.Text = cltIndice1.Item_Selecionado
    TxtIndice.SetFocus
End Sub


Private Sub cltIndice1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
       If Not IsNull(ctlPersonal.Valor) Then
          PopupMenu mnuArbol
          InsertarProducion ctlPersonal.Valor, 18, "Buscar Refrencia", 1, ctlCliente.Valor
        Else
          MsgBox "Insertar personal", vbInformation
        End If
    End If

End Sub


Private Sub cmdAceptar_Click()
     Toolbar1_ButtonClick Toolbar1.Buttons.Item("Aceptar")
End Sub

Private Sub AddFiltro()
    
Select Case cboFiltro
 Case "Not Like"
        lblFiltro.Caption = lblFiltro & " AND REFERENCIAS." & cboCampoFiltro & " " & cboFiltro & "'%" & txtDato.Text & "%'"
Case "Like"
        lblFiltro.Caption = lblFiltro & " AND REFERENCIAS." & cboCampoFiltro & " " & cboFiltro & "'%" & txtDato.Text & "%'"
Case "ENTRE FECHA"
        lblFiltro.Caption = lblFiltro & " AND FechaServerTipo ( txtDato.Text) BETWEEN REFERENCIAS.FECHA_DESDE AND REFERENCIAS.FECHA_HASTA "
Case "ENTRE NUMERO"
        lblFiltro.Caption = lblFiltro & " AND " & txtDato.Text & "  BETWEEN REFERENCIAS.NRO_DESDE AND REFERENCIAS.NRO_HASTA "
Case ">"
    lblFiltro.Caption = lblFiltro & " AND REFERENCIAS." & cboCampoFiltro & " > '" & txtDato.Text & "'"
Case "<"
    lblFiltro.Caption = lblFiltro & " AND REFERENCIAS." & cboCampoFiltro & " < '" & txtDato.Text & "'"
Case "="
    lblFiltro.Caption = lblFiltro & " AND REFERENCIAS." & cboCampoFiltro & " = '" & txtDato.Text & "'"
Case Else
    lblFiltro.Caption = lblFiltro & cboCampoFiltro & " " & cboFiltro & txtDato.Text
End Select
End Sub

Private Sub cmdAgregar_Click()
Dim Sql As String
Dim rsclon As New ADODB.Recordset
MousePointer = 13
Set rsclon = rsReferencias.Clone
rsclon.Filter = "nro_caja= " & grdReferencias.Columns(1).Text

Do While Not rsclon.EOF
    rsclon!Descripcion = "SECTOR REQUE.:" & txtAgregarReferencia.Text & " " & rsclon!Descripcion
    rsclon.Update
    rsclon.MoveNext
Loop

rsclon.Close

rsReferencias.Requery
grdReferencias.Refresh
MousePointer = 0
picSectorRequerimiento.Visible = False


End Sub

Private Sub cmdArbol_Click()
    BuscarIndice txtCodigo.Text, True
End Sub

Private Sub cmdBorrarReferencia_Click()
Dim Sql As String
Dim Registros As Integer
Dim ConSqlBasa  As New ADODB.Connection

ConSqlBasa.Open strConBasa
If Len(txtReferenciaLote.Text) = 7 Then
    MousePointer = 11
    Sql = " DELETE FROM REFERENCIAS WHERE PASOARCHIVO LIKE '%" & txtReferenciaLote.Text & "%'"
    ExecutarSql Sql
    Sql = " Update Documentos Set ESTADO = 0 Where COD_CLIENTE = 83 AND Lote = '" & txtReferenciaLote.Text & "'"
    ConSqlBasa.Execute Sql, Registros
    MousePointer = 0
    MsgBox "Cantidad de registros " & Registros
Else
    MsgBox "El Largo No es el correcto"
End If

End Sub

Private Sub cmdBuscarErrores_Click()
    Sql_Filtro_Referencia = " AND (FECHA_DESDE IS NULL) AND (NRO_DESDE IS NULL) "
    Sql_Filtro_Orden = " ORDER BY NRO_CAJA"
    FiltroReferencia True
End Sub

Private Sub cmdBuscarIndice_Click()



cltIndice1.BuscarIndice InputBox("Ingrese el indice"), True
End Sub

Private Sub cmdCambioIndice_Click()
Dim Clave As String
Clave = InputBox("Ingrese la clave", "Clave")
If Trim(Clave) = "21877471" Then
    If MsgBox("Esto Afectara a todos los registro", vbCritical + vbYesNo) = vbYes Then
        Dim Valor As String
            Valor = Trim(txtCodigo.Text)
            rsReferencias.MoveFirst
            Do While Not rsReferencias.EOF
                rsReferencias!USUARIO_MODIFICACION = ctlPersonal.Valor
                rsReferencias!FECHA_MODIFICACION = Now
                rsReferencias!Indice_Anterior = rsReferencias!Indice
                rsReferencias!Indice = Valor
                rsReferencias.Update
                rsReferencias.MoveNext
            Loop
            Set grdReferencias.DataSource = rsReferencias.DataSource
            grdReferencias.DataMember = rsReferencias.DataMember
            txtdescripcionBuscar.DataMember = rsReferencias.DataMember
            Set txtdescripcionBuscar.DataSource = rsReferencias.DataSource
            txtdescripcionBuscar.DataField = "Descripcion"
            txtCodigo.DataMember = rsReferencias.DataMember
            Set txtCodigo.DataSource = rsReferencias.DataSource
            txtCodigo.DataField = "Indice"
            grdReferencias.Refresh
            MsgBox "Operacion terminada ", vbInformation
  End If
 End If
End Sub

Private Sub cmdCerrar_Click()
picSectorRequerimiento.Visible = False
End Sub

Private Sub cmdContraer_Click()
  cltIndice1.Width = MDIfrmInicio.Width - picBuscar.Width - MDIfrmInicio.dxSideBar1.Width - 300
  ResizePic
   
    
End Sub

Private Sub cmdExpander_Click()
    cltIndice1.Width = 9000
    ResizePic
End Sub



Private Sub ControlCarga()

Dim rs As New ADODB.Recordset
Dim Sql As String

If IsNull(ctlCliente.Valor) Then
    MsgBox "Ingrese el cliente"
    Exit Sub
End If

Sql = " SELECT MIN(ID_IMAGEN) AS ImagenMinima, MAX(ID_IMAGEN)AS ImagenMaxima"
Sql = Sql & " From REFERENCIAS"
Sql = Sql & "  WHERE PASOARCHIVO LIKE '%" & txtReferenciaLote.Text & "%'"
Sql = Sql & " AND COD_CLIENTE =" & ctlCliente.Valor
rs.Open Sql, ConActiva, 0, 1

If Not rs.EOF Then
    lbImagenInicio.Caption = rs!ImagenMinima
    lblImagenActual.Caption = rs!ImagenMinima
    lblImagenFin.Caption = rs!ImagenMaxima
End If


End Sub

Private Sub cmdControlReferencias_Click()
      Dim DATO As String
        Sql_Filtro_Referencia = ""
         Sql_Filtro_Orden = ""
         Sql_Filtro_Referencia = " AND ID_IMAGEN =" & lblImagenActual.Caption
       Sql_Filtro_Orden = " order by NRO_CAjA "
       
        FiltroReferencia True
        sstReferencia.Tab = 1
      Rem   imgReferencia.MostrarImagen PasoImagenes & lblImagenActual.Caption & ".tif"
        If CLng(lblImagenActual.Caption) + 1 < CLng(lblImagenFin.Caption) Then
            lblImagenActual.Caption = lblImagenActual.Caption + 1
        End If
End Sub

Private Sub Informe_Carga()

        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Sql = "   SELECT     Lote, cod_Estado, COUNT(*) AS cantidad"
        Sql = Sql & "  FROM DOCUMENTOS_DIGITALES"
        Sql = Sql & "  Where (COD_CLIENTE = 83)"
        Sql = Sql & "  GROUP BY Lote, cod_Estado"
        Sql = Sql & "  Having (Not (Lote Is Null))"
        Sql = Sql & " Order by Lote "
        rs.CursorLocation = adUseClient
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        frmInforme.CargarInforme "Estado de las referencias", rs
        frmInforme.Show
        MousePointer = 0
End Sub

Private Sub Informe_Carga_Operador()
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim FechaControl As String
        
        FechaControl = InputBox("Ingrese la fecha de control")
        Sql = " SELECT USUARIO_MODIFICACION, FECHA_MODIFICACION,COD_ID_REFERENCIA"
        Sql = Sql & vbCrLf & " From REFERENCIAS"
        Sql = Sql & vbCrLf & " WHERE FECHA_MODIFICACION >" & FechaServerTipo(FechaControl)
        Sql = Sql & vbCrLf & " AND (USUARIO_MODIFICACION = '" & ctlPersonal.Valor & "')"
        Sql = Sql & vbCrLf & " ORDER BY COD_ID_REFERENCIA"
        rs.CursorLocation = adUseClient
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        frmInforme.CargarInforme "Estado de las Carga", rs
        frmInforme.Show
        MousePointer = 0
End Sub

Private Sub cmdImagenUna_Click()
Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim nuneroImagen As String

nuneroImagen = InputBox("Ingrese el Numero de Imagen")
 Sql = " SELECT     id, Lote,  COD_ESTADO, ArchivoNombre, PasoOrigen, IMAGEN_ORIGEN"
Sql = Sql & "  FROM DOCUMENTOS_DIGITALES"
Sql = Sql & "  WHERE     (COD_CLIENTE = 83) "
Sql = Sql & "  AND LOTE = '" & txtReferenciaLote.Text & "'"
Sql = Sql & "  AND IMAGEN_ORIGEN = " & nuneroImagen

      
      
      
        rs.Open Sql, ConActiva, 0, 1
    If Not rs.EOF Then
        lbl_ID_imagen.Caption = rs!ID
        Rem imgReferencia.MostrarImagen PasoImagenes & BuscarDirectorioPaso(rs!ID) & "/" & rs!ID & ".tif"
        lblPasoImagen.Caption = rs!IMAGEN_ORIGEN
        
    Else
        MsgBox "No existen mas imagenes"
        lbl_ID_imagen.Caption = ""
    End If
End Sub

Private Sub cmdProxReferencia_Click()
        Dim rs As New ADODB.Recordset
        Dim Sql As String

        Sql = "  SELECT     id, Lote,  COD_ESTADO ,IMAGEN_ORIGEN,  ArchivoNombre, IMAGEN_ORIGEN"
        Sql = Sql & "  FROM DOCUMENTOS_DIGITALES"
        Sql = Sql & " WHERE  (COD_CLIENTE = 83) AND  Lote ='" & txtReferenciaLote.Text & "'"
        Sql = Sql & "  AND  COD_ESTADO = 0 "
        Sql = Sql & " ORDER BY id"
        
         Sql = " SELECT   IMAGEN_ORIGEN,  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION,"
         Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.COD_ESTADO, DOCUMENTOS_DIGITALES.ID,"
         Sql = Sql & " DOCUMENTOS_DIGITALES.DIRECTORIO_PASO"
         Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON"
         Sql = Sql & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
         Sql = Sql & "  WHERE     (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 83) "
         Sql = Sql & "  AND (DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION = '" & txtReferenciaLote.Text & "')"
         Sql = Sql & "  AND  COD_ESTADO = 0 "
         Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES.ID  "
        
        
        
        rs.Open Sql, ConActiva, 0, 1
    If Not rs.EOF Then
        lbl_ID_imagen.Caption = rs!ID
        Rem imgReferencia.MostrarImagen PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".TIF"
        lblPasoImagen.Caption = rs!IMAGEN_ORIGEN
    Else
        MsgBox "no existen mas imagenes"
        lbl_ID_imagen.Caption = ""
    End If
    

End Sub

Private Sub cmdReplazar_Click()
 rsReferencias.MoveFirst
 Do While Not rsReferencias.EOF
            rsReferencias!USUARIO_MODIFICACION = ctlPersonal.Valor
            rsReferencias!FECHA_MODIFICACION = SysDate2
            rsReferencias!Descripcion = Replace(rsReferencias!Descripcion, txtElementoExistente, txtReplazo)
            rsReferencias.Update
            rsReferencias.MoveNext
        Loop
        Set grdReferencias.DataSource = rsReferencias.DataSource
        grdReferencias.DataMember = rsReferencias.DataMember
        txtdescripcionBuscar.DataMember = rsReferencias.DataMember
        Set txtdescripcionBuscar.DataSource = rsReferencias.DataSource
        txtdescripcionBuscar.DataField = "Descripcion"
        txtCodigo.DataMember = rsReferencias.DataMember
        Set txtCodigo.DataSource = rsReferencias.DataSource
        txtCodigo.DataField = "Indice"
        grdReferencias.Refresh
        MsgBox "Operacion terminada ", vbInformation


End Sub



Private Sub cmdTerminada_Click()
Dim conA As New ADODB.Connection
conA.Open strConBasa
On Error GoTo salir
Dim Sql As String
Sql = " Update   DOCUMENTOS_DIGITALES "
Sql = Sql & " Set COD_Estado = 100 "
Sql = Sql & " , PERSONAL_INDEXACION = " & ctlPersonal.Valor
Sql = Sql & " Where Id = " & lbl_ID_imagen
conA.Execute Sql
cmdProxReferencia_Click
salir:
End Sub

Private Sub Command1_Click()
     txtDato.Text = ""
     cboCampoFiltro.ListIndex = -1
     cboFiltro.Clear
     picFiltro.Visible = False
     
End Sub

Private Sub cmdFiltro_Click()
   Sql_Filtro_Referencia = lblFiltro.Caption
   Sql_Filtro_Orden = " ORDER BY FECHA_DESDE , NRO_DESDE"
   FiltroReferencia True
   picFiltro.Visible = False
End Sub



Public Sub CargarIndices(rsIndices As ADODB.Recordset)
    Dim Indice0 As String
    Dim KeyTreeView1 As String
    Dim Indice1 As String
    Dim Descripcion As String
    Dim nodX As Node
'        trvIndices.Nodes.Clear
'        Set nodX = trvIndices.Nodes.Add(, , "RAIZ", "TODAS LAS CATEGORIAS", "Casa") ' Root
'        trvIndices.Nodes.Item("RAIZ").Tag = "TODOS"
'        Do While Not rsIndices.EOF
'            If ExisItem("R" & Mid(rsIndices!Indice, 1, Len(rsIndices!Indice) - 3)) Then
'                KeyTreeView1 = "R" & Mid(rsIndices!Indice, 1, Len(rsIndices!Indice) - 3)
'                Descripcion = rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!Descripcion)
'                Set nodX = trvIndices.Nodes.Add(KeyTreeView1, tvwChild, "R" & rsIndices!Indice, Descripcion, "Punto", "Bandera")
'                trvIndices.Nodes.Item("R" & rsIndices!Indice).Tag = rsIndices!Indice
'            Else
'                Descripcion = rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!Descripcion)
'                Set nodX = trvIndices.Nodes.Add(, , "R" & rsIndices!Indice, Descripcion, "Punto", "Bandera")   ' Root
'                trvIndices.Nodes.Item("R" & rsIndices!Indice).Tag = rsIndices!Indice
'            End If
'            rsIndices.MoveNext
'        Loop
'
        
'
'        trvIndices.Nodes.Clear
'        Set nodX = trvIndices.Nodes.Add(, , "RAIZ", "TODAS LAS CATEGORIAS", "Casa") ' Root
'        trvIndices.Nodes.Item("RAIZ").Tag = "TODOS"
'        Do While Not rsIndices.EOF
'
'            If ExisItem("R" & Mid(rsIndices!INDICE, 1, Len(rsIndices!INDICE) - 3)) Then
'                KeyTreeView1 = "R" & Mid(rsIndices!INDICE, 1, Len(rsIndices!INDICE) - 3)
'                DESCRIPCION = rsIndices!INDICE & " - " & rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!DESCRIPCION)
'                Set nodX = trvIndices.Nodes.Add(KeyTreeView1, tvwChild, "R" & rsIndices!INDICE, DESCRIPCION, "Punto", "Bandera")
'                trvIndices.Nodes.Item("R" & rsIndices!INDICE).Tag = rsIndices!INDICE
'            Else
'                DESCRIPCION = rsIndices!INDICE & " - " & rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!DESCRIPCION)
'                Set nodX = trvIndices.Nodes.Add(, , "R" & rsIndices!INDICE, DESCRIPCION, "Punto", "Bandera")   ' Root
'                trvIndices.Nodes.Item("R" & rsIndices!INDICE).Tag = rsIndices!INDICE
'            End If
'            rsIndices.MoveNext
'        Loop
End Sub
'Public Function ExisItem(dato As String) As Boolean
'    Dim s As String
'    On Error GoTo ErrorHandler
'        ExisItem = True
'        s = trvIndices.Nodes.Item(dato)
'    Exit Function
'ErrorHandler:
'    ExisItem = False
'End Function

Private Sub ctlCliente_Click()
    Dim rsIndice As New ADODB.Recordset
    Dim sSQL As String
      If Not IsNull(ctlPersonal.Valor) Then
        sSQL = "Select * From Indices Where Cod_Cliente =" & ctlCliente.Valor & "Order by indice        "
            rsIndice.Open sSQL, ConActiva, 0, 1
        If Not rsIndice.EOF Then
            CargarIndices rsIndice
            cltIndice1.Actualizar ctlCliente.Valor, Nulo, 0
            LimpiarTodos
           
             Set grdModificacion.DataSource = Nothing
              grdModificacion.ClearFields
            grdModificacion.Refresh
         
        Else
'            cltIndice1.
'             trvIndices.Nodes.Clear
        End If
     Else
        MsgBox "Ingrese el personal", vbCritical
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 120 Then
       Toolbar1_ButtonClick Toolbar1.Buttons.Item(4)
    End If
    If KeyCode = 121 Then
       cmdBuscarIndice_Click
    End If
    
    
End Sub


Private Sub Form_Load()
    ctlCliente.TipoControl = Cliente
    ctlPersonal.TipoControl = Personal

'    Rem frmCajasReferencias.StartUpPosition = vbStartUpScreen
'    frmCajasReferencias.WindowState = 2
'    picCargar.Visible = False
'    picBuscar.Visible = False
'    picSectorRequerimiento.Visible = False
'    ResizePic
'    picCargar.Visible = False
'    picBuscar.Visible = True
'    StatusBar1.Panels("EstadoAplicacion").Text = "Buscar"
'    picCargar.Top = 1140
'    picCargar.Left = 3360
'    picBuscar.Top = 1140
'    picBuscar.Left = 3360
'    MousePointer = 0
'    frmCajasReferencias.WindowState = vbMaximized
'        cmdContraer_Click
'        ResizePic
'        picCargar.Visible = False
'        picBuscar.Visible = True
'        StatusBar1.Panels("EstadoAplicacion").Text = "Buscar"

End Sub

Private Sub Form_LostFocus()
    ResizePic
End Sub

Private Sub Form_Resize()
    On Error GoTo salir
    Rem sstReferencia.Height = frmCajasReferencias.Height - 200
     Rem imgReferencia.Height = frmCajasReferencias.Height - (sstReferencia.Height + Toolbar1.Height + StatusBar1.Height + 1000)
        
salir:
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
' If Not oWordL Is Nothing Then
'    MousePointer = 11
'   If Not oTmpDocL Is Nothing Then
'    oTmpDocL.Close

'    End If
'    oWordL.Quit
'    Set oWordL = Nothing
'    MousePointer = 0
' End If

End Sub

Private Sub grdDescripcionRepetida_DblClick()
txtDescripcion.Text = grdDescripcionRepetida.TextMatrix(grdDescripcionRepetida.RowSel, 1)
grdDescripcionRepetida.Visible = False
End Sub

Private Sub grdModificacion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
 PopupMenu mnugrdmodificar
 End If


End Sub


Private Sub grdReferencias_Click()
grdReferencias.Col = 0
txtCodigo.Text = grdReferencias.Text
End Sub

Private Sub grdReferencias_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
        PopupMenu MNUGRDBUSQUEDA
    End If

End Sub

Private Sub grdSector_DblClick()
 If Trim(grdSector.Columns.Item(3).Text) <> "" Then
    txtAgregarReferencia.Text = Trim(grdSector.Columns.Item(3).Text)
    Else
    If Trim(grdSector.Columns.Item(4).Text) <> "" Then
        txtAgregarReferencia.Text = Trim(grdSector.Columns.Item(4).Text)
    End If
 End If
 
End Sub

Private Sub mnuBorrarDatos_Click()
    Dim Sql As String
    If MsgBox("Esta Ud seguro de borrar el registro", vbInformation + vbYesNo) = vbYes Then
        Sql = " DELETE FROM REFERENCIAS"
        Sql = Sql & "  Where COD_ID_REFERENCIA = " & grdModificacion.Columns.Item(grdModificacion.Columns.Count - 1).Text
        ExecutarSql Sql
    End If
    ActualizarGrillaCarga
End Sub

Private Sub mnuBuscarCaja_Click()
        Dim DATO As String
        Sql_Filtro_Referencia = ""
        DATO = InputBox("Ingrese la busqueda de cajas Ej. 15,16,23,25 ")
        If (DATO) <> "" Then
            Sql_Filtro_Orden = ""
            Sql_Filtro_Referencia = " AND NRO_CAJA in  (" & DATO & ")"
        Else
            MsgBox "NO SE INGRESARON LOS DATOS CORRECTOS"
            Exit Sub
        End If
        Sql_Filtro_Orden = ""
        FiltroReferencia True
        sstReferencia.Tab = 1



End Sub

Private Sub mnuBuscarCajas_Click()
 Dim DATO As String
        Sql_Filtro_Referencia = ""
        DATO = InputBox("Ingrese la busqueda CAJA INICIO AND CAJA FINAL ")
        If (DATO) <> "" Then
            Sql_Filtro_Orden = ""
            Sql_Filtro_Referencia = " AND NRO_CAJA  BETWEEN " & DATO
        Else
            MsgBox "NO SE INGRESARON LOS DATOS CORRECTOS"
            Exit Sub
        End If
        Sql_Filtro_Orden = " ORDER BY NRO_CAJA"
        FiltroReferencia True
        sstReferencia.Tab = 1
End Sub

Private Sub mnuBuscarDescripcion_Click()
    Dim DATO As String
        Sql_Filtro_Referencia = ""
        DATO = InputBox("Ingrese la busqueda")
        If (DATO) <> "" Then
            Sql_Filtro_Orden = ""
            Sql_Filtro_Referencia = " AND Referencias.DESCRIPCION like '%" & UCase(DATO) & "%' "
        Else
            MsgBox "NO SE INGRESARON LOS DATOS CORRECTOS"
            Exit Sub
        End If
        Sql_Filtro_Orden = " ORDER BY FECHA_DESDE , NRO_DESDE"
        FiltroReferencia True
        sstReferencia.Tab = 1
End Sub

Private Sub mnuBuscarIndice_Click()
'Dim Datos As String
'Datos = InputBox("Ingrese la Letra a Buscar")
'cltIndice1.BuscarIndice Datos, True

End Sub

Private Sub MNUCAJAPOSICION_Click()
    frmCajasUbicacion.Show
    frmCajasUbicacion.InsertarGrilla grdReferencias.Columns(1).Text, ctlCliente.Valor, False
    frmCajasUbicacion.SetFocus
End Sub

Private Sub mnuCajasposicionTodo_Click()
    frmCajasUbicacion.Show
       rsReferencias.MoveFirst
    Do While Not rsReferencias.EOF
        frmCajasUbicacion.InsertarGrilla rsReferencias!NRO_CAJA, ctlCliente.Valor, False
        rsReferencias.MoveNext
    Loop
    frmCajasUbicacion.SetFocus
End Sub

Private Sub mnuCantidadcajas_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String

Sql = " SELECT     COUNT(DISTINCT NRO_CAJA) AS Cantidad"
Sql = Sql & " From REFERENCIAS"
Sql = Sql & "  WHERE COD_CLIENTE = " & ctlCliente.Valor
Sql = Sql & "  AND INDICE LIKE '" & cltIndice1.Item_Selecionado & "%'"

rs.Open Sql, ConActiva, 0, 1
If Not rs.EOF Then
    MsgBox "La cantidad de cajas es " & rs!cantidad
End If


End Sub

Private Sub mnuCopiarDatos_Click()
    CopiarDatosGrilla grdReferencias
End Sub

Private Sub mnuDescripcionIndiceFijo_Click()
Dim DATO As String
        Sql_Filtro_Referencia = ""
        DATO = InputBox("Ingrese la busqueda")
        If (DATO) <> "" Then
            Sql_Filtro_Orden = ""
            Sql_Filtro_Referencia = " AND Referencias.DESCRIPCION like '%" & UCase(DATO) & "%' "
        Else
            MsgBox "NO SE INGRESARON LOS DATOS CORRECTOS"
        End If
        Sql_Filtro_Orden = " ORDER BY FECHA_DESDE , NRO_DESDE"
        FiltroReferencia True, True
End Sub

Private Sub mnuDeshabilitarfecha_Click()
Dim Indice As String
Dim Sql As String
    Indice = cltIndice1.Item_Selecionado
    Sql = " Update INDICES"
    Sql = Sql & " SET  FECHA =Null "
    Sql = Sql & " WHERE COD_CLIENTE = " & ctlCliente.Valor
    Sql = Sql & "  AND INDICE ='" & Indice & "'"
    ExecutarSql Sql

End Sub

Private Sub mnuDeshabilitarLetra_Click()
Dim Indice As String
Dim Sql As String
    Indice = cltIndice1.Item_Selecionado
    Sql = " Update INDICES"
    Sql = Sql & " SET  LETRA = NULL "
    Sql = Sql & " WHERE COD_CLIENTE = " & ctlCliente.Valor
    Sql = Sql & "  AND INDICE ='" & Indice & "'"
    ExecutarSql Sql
End Sub

Private Sub mnuDeshabilitarNumero_Click()
Dim Indice As String
Dim Sql As String
    Indice = cltIndice1.Item_Selecionado
    Sql = " Update INDICES"
    Sql = Sql & " SET  Numero =Null "
    Sql = Sql & " WHERE COD_CLIENTE = " & ctlCliente.Valor
    Sql = Sql & "  AND INDICE ='" & Indice & "'"
    ExecutarSql Sql
End Sub

Private Sub mnuFechaBuscar_Click()
 Dim DATO As String
        Sql_Filtro_Referencia = ""
        DATO = InputBox("Ingrese la Fecha a Buscar")
        If IsDate(DATO) Then
            Rem Sql_Filtro_Orden = Orden_Referencia(trvIndices.SelectedItem.Text, ctlCliente.Valor)
            Sql_Filtro_Referencia = " AND (" & FechaServerTipo(DATO) & " BETWEEN FECHA_DESDE AND FECHA_HASTA)"
        Else
            MsgBox "NO SE INGRESARON LOS DATOS CORRECTOS"
            Exit Sub
        End If
        Sql_Filtro_Orden = Orden_Referencia(ctlCliente.Valor)
        FiltroReferencia True
        sstReferencia.Tab = 1
End Sub

Private Sub mnuFechaDescripcion_Click()
Dim DATO As String
Dim Desc As String
        Sql_Filtro_Referencia = ""
        DATO = InputBox("Ingrese la Fecha a Buscar")
        Desc = InputBox("Ingrese parte de la descripcion")
        If IsDate(DATO) Then
            Rem Sql_Filtro_Orden = Orden_Referencia(trvIndices.SelectedItem.Text, ctlCliente.Valor)
            Sql_Filtro_Referencia = " AND (" & FechaServerTipo(DATO) & " BETWEEN FECHA_DESDE AND FECHA_HASTA)"
            Sql_Filtro_Referencia = Sql_Filtro_Referencia & " AND DESCRIPCION  like  '%" & Desc & "%'"
        Else
            MsgBox "NO SE INGRESARON LOS DATOS CORRECTOS"
            Exit Sub
        End If
        Sql_Filtro_Orden = Orden_Referencia(ctlCliente.Valor)
        FiltroReferencia True
        sstReferencia.Tab = 1
End Sub

Private Sub mnuFechaOrdenar_Click()
 Dim DATO As String
        Sql_Filtro_Orden = " order by fecha_desde"
        Sql_Filtro_Referencia = ""
        FiltroReferencia True
        sstReferencia.Tab = 1
End Sub

Private Sub mnuFiltro_Click()
    picFiltro.Visible = True
    picFiltro.Left = picBuscar.Left
    picFiltro.Top = picBuscar.Top
    lblFiltro.Caption = ""
End Sub

Private Sub mnuHabilitarFecha_Click()
Dim Indice As String
Dim Sql As String
    Indice = cltIndice1.Item_Selecionado
    Sql = " Update INDICES"
    Sql = Sql & " SET  FECHA ='1'"
    Sql = Sql & " WHERE COD_CLIENTE = " & ctlCliente.Valor
    Sql = Sql & "  AND INDICE = '" & Indice & "'"
    ExecutarSql Sql
   
End Sub

Private Sub mnuHabilitarLetra_Click()
Dim Indice As String
Dim Sql As String
    Indice = cltIndice1.Item_Selecionado
    Sql = " Update INDICES"
    Sql = Sql & " SET  Letra =1 "
    Sql = Sql & " WHERE COD_CLIENTE = " & ctlCliente.Valor
    Sql = Sql & "  AND INDICE ='" & Indice & "'"
    ExecutarSql Sql
End Sub

Private Sub mnuHabilitarNumero_Click()
Dim Indice As String
Dim Sql As String
    Indice = cltIndice1.Item_Selecionado
    Sql = " Update INDICES"
    Sql = Sql & " SET  Numero =1 "
    Sql = Sql & " WHERE COD_CLIENTE = " & ctlCliente.Valor
    Sql = Sql & "  AND INDICE ='" & Indice & "'"
    ExecutarSql Sql
End Sub

Private Sub mnuIDReferencia_Click()
        Dim DATO As String
        Sql_Filtro_Referencia = ""
        DATO = InputBox("Ingrese la busqueda")
        If (DATO) <> "" Then
            Sql_Filtro_Orden = ""
            Sql_Filtro_Referencia = " AND Referencias.COD_ID_REFERENCIA IN (  " & DATO & ") "
        Else
            MsgBox "NO SE INGRESARON LOS DATOS CORRECTOS"
            Exit Sub
        End If
        Sql_Filtro_Orden = " ORDER BY FECHA_DESDE , NRO_DESDE"
        FiltroReferencia True
End Sub

Private Sub mnuInconsistenciasNivel_Click()

        Dim DATO As String
        
'        SELECT REFERENCIAS.COD_CLIENTE, REFERENCIAS.NRO_CAJA,
'    REFERENCIAS.INDICE, INDICES.ID_CODIGO_DOCUMENTO,
'    INDICES.Indice , INDICES.TIPO_INDICE
'From REFERENCIAS, INDICES
'WHERE REFERENCIAS.INDICE = INDICES.INDICE AND
'    (REFERENCIAS.COD_CLIENTE = 04) AND
'    (REFERENCIAS.INDICE LIKE '001001003%') AND
'    (INDICES.TIPO_INDICE <> 'Documento')
'ORDER BY REFERENCIAS.INDICE
        
        If (DATO) <> "" Then
            Sql_Filtro_Orden = ""
            Rem Sql_Filtro_Referencia = " AND Referencias.DESCRIPCION like '%" & UCase(Dato) & "%' "
            Sql_Filtro_Referencia = ""
        End If
        Sql_Filtro_Orden = " ORDER BY FECHA_DESDE , NRO_DESDE"
        FiltroReferencia True, True
sstReferencia.Tab = 1


End Sub

Private Sub mnuLetra_Click()
'    Dim dato As String
'    Dim Orden As String
'        dato = InputBox("Ingrese la Letra a Buscar")
'        Sql_Filtro_Referencia = ""
'        Sql_Filtro_Referencia = " AND LETRA_DESDE ='" & dato & "' "
'        Sql_Filtro_Orden = Orden_Referencia(trvIndices.SelectedItem.Text, ctlCliente.Valor)
'        FiltroReferencia True
'        sstReferencia.Tab = 1
End Sub

Private Sub mnuModificar_Click()

       Set rsReferencias = New ADODB.Recordset
       Dim sSQL As String
       Dim Item As Integer
           
           rsReferencias.CursorLocation = adUseClient
           sSQL = "Select * from Referencias where cod_Cliente =" & ctlCliente.Valor
           sSQL = sSQL & vbCrLf & " AND NRO_CAJA =" & grdReferencias.Columns(1).Text
           sSQL = sSQL & vbCrLf & " ORDER BY INDICE"
           rsReferencias.Open sSQL, ConActiva, adOpenDynamic, adLockOptimistic
           Set grdModificacion.DataSource = rsReferencias.DataSource
           grdModificacion.DataMember = rsReferencias.DataMember
     sstReferencia.Tab = 0
End Sub

Private Sub mnuModificarDatos_Click()
        LimpiarTodos
        StatusBar1.Panels("EstadoAplicacion").Text = "Modificar"
        lblCod_Referencia.Caption = grdModificacion.Columns.Item(grdModificacion.Columns.Count - 1).Text
        txtNro_Caja.Text = grdModificacion.Columns.Item(0).Text
        TxtIndice.Text = grdModificacion.Columns.Item(1).Text
        txtDescripcion.Text = grdModificacion.Columns.Item(2).Text
        If grdModificacion.Columns.Item(3).Text <> "" Then
            mskFecha_Desde.Text = grdModificacion.Columns.Item(3).Text
        End If
        If grdModificacion.Columns.Item(4).Text <> "" Then
            mskFecha_Hasta.Text = grdModificacion.Columns.Item(4).Text
        End If
        If grdModificacion.Columns.Item(5).Text <> "" Then
            mskNro_desde.Text = grdModificacion.Columns.Item(5).Text
        End If
        If grdModificacion.Columns.Item(6).Text <> "" Then
            mskNro_hasta.Text = grdModificacion.Columns.Item(6).Text
        End If
        If grdModificacion.Columns.Item(7).Text <> "" Then
            mskLetra_Desde.Text = grdModificacion.Columns.Item(7).Text
        End If
        If grdModificacion.Columns.Item(8).Text <> "" Then
            mskLetra_Hasta.Text = grdModificacion.Columns.Item(8).Text
        End If
        txtNro_Caja.SetFocus

End Sub

Private Sub MNUNOMBREYAPELLIDO_Click()
' Dim dato As String
'        Sql_Filtro_Referencia = ""
'        dato = InputBox("Ingrese Nombre a Buscar")
'        Sql_Filtro_Orden = Orden_Referencia(trvIndices.SelectedItem.Text, ctlCliente.Valor)
'        Sql_Filtro_Referencia = " and APELLIDO_NOMBRE like '%" & dato & "%'"
'        Sql_Filtro_Orden = Orden_Referencia(trvIndices.SelectedItem.Text, ctlCliente.Valor)
'        FiltroReferencia True
End Sub

Private Sub mnunro_desde_Click()
'Dim NUMERO As String
'    Dim Orden As String
'        NUMERO = InputBox("Ingrese el numero a Buscar")
'        Sql_Filtro_Referencia = ""
'        If IsNumeric(NUMERO) Then
'            Sql_Filtro_Referencia = " AND  NRO_DESDE = " & CLng(NUMERO)
'            Sql_Filtro_Orden = " ORDER BY NRO_DESDE , FECHA_DESDE "
'        Else
'            Sql_Filtro_Orden = Orden_Referencia(trvIndices.SelectedItem.Text, ctlCliente.Valor)
'            MsgBox "NO SE INGRESARON LOS DATOS CORRECTOS"
'        End If
'        Sql_Filtro_Orden = Orden_Referencia(ctlCliente.Valor)
'        FiltroReferencia True
'        sstReferencia.Tab = 1
End Sub

Private Sub mnuNumero_Click()
    Dim NUMERO As String
    Dim Orden As String
    On Error GoTo salir:
        NUMERO = InputBox("Ingrese el numero a Buscar")
        Sql_Filtro_Referencia = ""
         
        If CLng(NUMERO) = 0 Then
        
         Sql_Filtro_Orden = " order by nro_desde "
        Else
        If IsNumeric(NUMERO) Then
            Sql_Filtro_Referencia = " AND (" & CLng(NUMERO) & " BETWEEN NRO_DESDE AND NRO_HASTA)"
        Else
            Sql_Filtro_Orden = Orden_Referencia(ctlCliente.Valor)
            MsgBox "NO SE INGRESARON LOS DATOS CORRECTOS"
            Exit Sub
        End If
        End If
        
        FiltroReferencia True
        sstReferencia.Tab = 1
salir:
End Sub

Private Sub mnuResponsables_Click()
 Dim rs As New ADODB.Recordset
        Dim Sql As String
        
   Sql = " SELECT   ID_CLIENTEUSUARIO,  COD_INDICE, APELLIDO_NOMBRE, CORREO, TELEFONOS ,REFERENCIAS"
Sql = Sql & "  From CLIENTEUSUARIO"
Sql = Sql & "  Where COD_CLIENTE = " & ctlCliente.Valor
If Not (cltIndice1.Item_Selecionado = "AIZ") Then
  Sql = Sql & "  AND (COD_INDICE LIKE '" & cltIndice1.Item_Selecionado & "%')"
Else
    
End If

Sql = Sql & "  ORDER BY APELLIDO_NOMBRE"
       
        rs.CursorLocation = adUseClient
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        frmInforme.CargarInforme "Estado de las referencias", rs
        frmInforme.Show
        MousePointer = 0
End Sub

Private Sub mnutodalacaj_Click()
      Sql_Filtro_Orden = ""
      Sql_Filtro_Referencia = " AND NRO_CAJA  = " & grdReferencias.Columns(1).Text
      FiltroReferencia False
End Sub

Private Sub mnuTodos_Click()
    Dim i As Integer
    Sql_Filtro_Referencia = ""
    Sql_Filtro_Orden = Orden_Referencia(ctlCliente.Valor)
    FiltroReferencia True
    sstReferencia.Tab = 1
    
End Sub

Private Sub mnuVerImagen_Click()
 Dim Sql As String
  Dim i As Integer
  ReDim a(50) As String
  Dim rs As New ADODB.Recordset
        
        
        Sql = "  SELECT ID_SQL  From IMAGENES"
        Sql = Sql & " Where TIPO_DOCUMENTO = 1"
        Sql = Sql & " And Elemento = " & grdReferencias.Columns(1).Text
        Sql = Sql & " and COD_CLIENTE= " & ctlCliente.Valor
        rs.Open Sql, ConActiva, 0, 1
        i = 0
           Do While Not rs.EOF
            
            a(i) = PasoImagenes & BuscarDirectorioPaso(rs!ID_SQL) & "/" & rs!ID_SQL & ".tif"
            i = i + 1
            rs.MoveNext
        Loop
        
        If i = 0 Then
            MsgBox "NO existen imagenes"
             Rem imgReferencia.MostrarImagen ""
        End If
        
        If i > 0 Then
            ReDim Preserve a(i - 1)
            Rem imgReferencia.MostrarImagenes a
        End If
End Sub

Private Sub mnuVerLote_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
        
        Sql = " SELECT PASOARCHIVO"
        Sql = Sql & " From REFERENCIAS"
        Sql = Sql & " Where COD_CLIENTE = " & ctlCliente.Valor
        Sql = Sql & " And NRO_CAJA = " & grdReferencias.Columns(1).Text
        rs.Open Sql, ConActiva, 0, 1
        If Not rs.EOF Then
            MsgBox "Lote :" & rs!PASOARCHIVO
        Else
            MsgBox "No exsite lote"
        End If



End Sub

Private Sub mnuVerRequerimientos_Click()
Dim rsReque As ADODB.Recordset
        Set rsReque = New ADODB.Recordset
        Dim sSQL As String
            rsReque.CursorLocation = adUseClient
            sSQL = "SELECT IDREQUERIMIENTO,CAJASLIBROS AS CAJA, TO_CHAR(FECHARECEPCION, 'dd/mm/yyyy') as FECHA,SECTOR,Solicitante"
            sSQL = sSQL & " From REQUERIMIENTO, REQUELIBOSCAJAS"
            sSQL = sSQL & " Where REQUERIMIENTO.IDRequerimiento = REQUELIBOSCAJAS.IDREQUERIMIENTOS"
            sSQL = sSQL & "  AND REQUERIMIENTO.ID_CLIENTE = " & ctlCliente.Valor & " AND  REQUELIBOSCAJAS.CAJASLIBROS = " & grdReferencias.Columns.Item(1).Text
            sSQL = sSQL & " ORDER BY FECHARECEPCION DESC"
            rsReque.Open sSQL, ConActiva, 0, 1
            If Not rsReque.EOF Then
               DATOSGRILLA grdSector, rsReque
               grdSector.Columns(0).Width = 700
               grdSector.Columns(1).Width = 700
               grdSector.Columns(2).Width = 1000
               grdSector.Columns(3).Width = 2000
               grdSector.Columns(4).Width = 2000
               picSectorRequerimiento.Visible = True
               picSectorRequerimiento.Top = picBuscar.Top
               picSectorRequerimiento.Left = picBuscar.Left
               txtAgregarReferencia.Text = ""
            Else
                MsgBox "NO SE ENCONTRARON MOVIMIENTO", vbInformation
                picSectorRequerimiento.Visible = False
                txtAgregarReferencia.Text = ""
            End If
            
End Sub

Private Sub mskExpediente_GotFocus()
   Rem StatusBar1.Panels("Ayuda").Text = mskExpediente.ToolTipText
End Sub

Private Sub mskExpediente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub mskFecha_Desde_GotFocus()
    StatusBar1.Panels("Ayuda").Text = mskFecha_Desde.ToolTipText
    mskFecha_Desde.SelStart = 0
End Sub

Private Sub mskFecha_Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If KeyAscii = 42 Then
         Select Case CInt(Mid(mskFecha_Desde.Text, 4, 2))
         Case 1, 3, 5, 7, 8, 10, 12
            mskFecha_Desde.Text = "01" & Mid(mskFecha_Desde.Text, 3)
            mskFecha_Hasta.Text = "31" & Mid(mskFecha_Desde.Text, 3)
         Case 2
            mskFecha_Desde.Text = "01" & Mid(mskFecha_Desde.Text, 3)
            mskFecha_Hasta.Text = "28" & Mid(mskFecha_Desde.Text, 3)
         Case 4, 6, 9, 11
            mskFecha_Desde.Text = "01" & Mid(mskFecha_Desde.Text, 3)
            mskFecha_Hasta.Text = "30" & Mid(mskFecha_Desde.Text, 3)
         End Select
          SendKeys vbTab
          SendKeys vbTab
     End If
     If KeyAscii = 45 Then
         Select Case CInt(Mid(mskFecha_Desde.Text, 4, 2))
         Case 1, 3, 5, 7, 8, 10, 12
            mskFecha_Desde.Text = "01" & Mid(mskFecha_Desde.Text, 3)
            mskFecha_Hasta.Text = "31/12" & Mid(mskFecha_Desde.Text, 6)
         Case 2
            mskFecha_Desde.Text = "01" & Mid(mskFecha_Desde.Text, 3)
            mskFecha_Hasta.Text = "31/12" & Mid(mskFecha_Desde.Text, 6)
         Case 4, 6, 9, 11
            mskFecha_Desde.Text = "01" & Mid(mskFecha_Desde.Text, 3)
            mskFecha_Hasta.Text = "31/12" & Mid(mskFecha_Desde.Text, 6)
         End Select
          
          
          SendKeys vbTab
          SendKeys vbTab
     End If
    If KeyAscii = 43 Then
         mskFecha_Hasta.Text = mskFecha_Desde.Text
          SendKeys vbTab
          SendKeys vbTab
     End If
End Sub

Private Sub mskFecha_Desde_LostFocus()
On Error GoTo salir
   Select Case Mid(mskFecha_Desde.Text, 4, 2)
   Case 1, 3, 5, 7, 8, 10, 12 ' 31 dias
        If Mid(mskFecha_Desde.Text, 1, 2) > 31 Then
            MsgBox "Error en fecha mes de 31 dias", vbCritical
            mskFecha_Desde.SetFocus
            Exit Sub
        End If
   Case 2 ' 28 dias
   If Mid(mskFecha_Desde.Text, 1, 2) > 28 Then
            MsgBox "Error en fecha mes de 28 dias", vbCritical
            mskFecha_Desde.SetFocus
            Exit Sub
        End If
   Case 4, 6, 9, 11
        If Mid(mskFecha_Desde.Text, 1, 2) > 30 Then
            MsgBox "Error en fecha mes de 30 dias", vbCritical
            mskFecha_Desde.SetFocus
            Exit Sub
        End If
   End Select
   If mskFecha_Desde.ClipText <> "" Then
   If Mid(mskFecha_Desde.Text, 4, 2) > 12 Then
        MsgBox "Error en fecha desde mes ", vbCritical
        mskFecha_Desde.SetFocus
        Exit Sub
    End If
   End If
   Exit Sub
salir:

MsgBox "Error en fecha"
End Sub

Private Sub mskFecha_Hasta_GotFocus()
    StatusBar1.Panels("Ayuda").Text = mskFecha_Hasta.ToolTipText
    mskFecha_Hasta.SelStart = 0
End Sub

Private Sub mskFecha_Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    If KeyAscii = 42 Then
         Select Case CInt(Mid(mskFecha_Desde.Text, 4, 2))
         Case 1, 3, 5, 7, 8, 10, 12
            mskFecha_Desde.Text = "01" & Mid(mskFecha_Desde.Text, 3)
            mskFecha_Hasta.Text = "31" & Mid(mskFecha_Desde.Text, 3)
         Case 2
            mskFecha_Desde.Text = "01" & Mid(mskFecha_Desde.Text, 3)
            mskFecha_Hasta.Text = "28" & Mid(mskFecha_Desde.Text, 3)
         Case 4, 6, 9, 11
            mskFecha_Desde.Text = "01" & Mid(mskFecha_Desde.Text, 3)
            mskFecha_Hasta.Text = "30" & Mid(mskFecha_Desde.Text, 3)
         End Select
          SendKeys vbTab
          End If
End Sub

Private Sub mskFecha_Hasta_LostFocus()
On Error GoTo salir
   Select Case Mid(mskFecha_Hasta.Text, 4, 2)
   Case 1, 3, 5, 7, 8, 10, 12 ' 31 dias
        If Mid(mskFecha_Hasta.Text, 1, 2) > 31 Then
            MsgBox "Error en fecha mes de 31 dias", vbCritical
            mskFecha_Hasta.SetFocus
            Exit Sub
        End If
   Case 2 ' 28 dias
   If Mid(mskFecha_Hasta.Text, 1, 2) > 28 Then
            MsgBox "Error en fecha mes de 28 dias", vbCritical
            mskFecha_Hasta.SetFocus
            Exit Sub
        End If
   Case 4, 6, 9, 11
        If Mid(mskFecha_Hasta.Text, 1, 2) > 30 Then
            MsgBox "Error en fecha mes de 30 dias", vbCritical
            mskFecha_Hasta.SetFocus
            Exit Sub
        End If
   End Select
   If Mid(mskFecha_Hasta.ClipText, 4, 2) <> "" Then
    If Mid(mskFecha_Hasta.Text, 4, 2) > 12 Then
        MsgBox "Error en fecha hasta mes ", vbCritical
        mskFecha_Hasta.SetFocus
        Exit Sub
    End If
    
    If Mid(mskFecha_Desde.Text, 4, 2) > 12 Then
        MsgBox "Error en fecha desde mes ", vbCritical
        mskFecha_Desde.SetFocus
        Exit Sub
    End If
   End If
   
   
   If mskFecha_Desde.ClipText <> "" And mskFecha_Hasta.ClipText <> "" Then
        If CDate(mskFecha_Desde.Text) > CDate(mskFecha_Hasta.Text) Then
         MsgBox "Fecha desde mayor que fecha hasta", vbInformation
         mskFecha_Desde.SetFocus
        End If
   End If
   
   If mskFecha_Desde.ClipText <> "" And mskFecha_Hasta.ClipText = "" Then
        MsgBox "Fecha hasta es obligatoria", vbInformation
        mskFecha_Desde.SetFocus
   End If
   
   Exit Sub
salir:

MsgBox "Error en fecha"

   
   
End Sub

Private Sub mskLetra_Desde_GotFocus()
    StatusBar1.Panels("Ayuda").Text = mskLetra_Desde.ToolTipText
     
End Sub

Private Sub mskLetra_Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub


Private Sub mskLetra_Hasta_GotFocus()
    StatusBar1.Panels("Ayuda").Text = mskLetra_Hasta.ToolTipText
End Sub

Private Sub mskLetra_Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub


Private Sub mskNro_desde_GotFocus()
    StatusBar1.Panels("Ayuda").Text = mskNro_desde.ToolTipText
End Sub

Private Sub mskNro_desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub


Private Sub mskNro_hasta_GotFocus()
    StatusBar1.Panels("Ayuda").Text = mskNro_hasta.ToolTipText
End Sub

Private Sub mskNro_hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub




Private Sub mskNro_hasta_LostFocus()

If mskNro_desde.ClipText <> "" And mskNro_hasta.ClipText = "" Then
    MsgBox "El numero hasta en obligatorio", vbInformation
    mskNro_desde.SetFocus
    Exit Sub
End If

If mskNro_desde.ClipText <> "" And mskNro_hasta.ClipText <> "" Then
    If CLng(mskNro_desde.Text) > CLng(mskNro_hasta.Text) Then
        MsgBox "El numero hasta en mayor que el numero hasta", vbInformation
        mskNro_desde.SetFocus
    End If
End If

End Sub

Private Sub picBuscar_Resize()
'    grdReferencias.Width = picBuscar.Width - 250
'    grdReferencias.Height = picBuscar.Height - 2000
'    txtdescripcionBuscar.Width = grdReferencias.Width - 1200
'    cmdBuscarErrores.left = txtdescripcionBuscar.Width
'    cmdCambioIndice.left = cmdBuscarErrores.left
'    lblCodigo.Width = txtdescripcionBuscar.Width - txtCodigo.Width
End Sub



Private Sub rsReferencias_RecordChangeComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, ByVal pError As ADODB.ERROR, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  If Not IsNull(ctlPersonal.Valor) Then
  InsertarProducion ctlPersonal.Valor, 3, "Reparacion referencia ", "0,01", ctlCliente.Valor
   
   Dim Sql As String
'    SQL = " Update REFERENCIAS"
'    SQL = SQL & " SET FECHA_MODIFICACION = " & SysDate
'    SQL = SQL & ", USUARIO_MODIFICACION = '" & CTLPERSONAL.Valor & "'"
'    SQL = SQL & " Where COD_ID_REFERENCIA =" & rsReferencias.Fields("COD_ID_REFERENCIA").Value
'    ExecutarSql SQL

   
  Else
    MsgBox "Ingrese el personal"
  End If
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Aceptar"
       If Not IsNull(ctlPersonal.Valor) Then
           GrabarReferencias
           InsertarProducion ctlPersonal.Valor, 3, "Carga Ref Caja:" & txtNro_Caja.Text, 1, ctlCliente.Valor
        Else
            MsgBox "Ingrese Carga ", vbInformation
        End If
    Case "Nuevo"
        StatusBar1.Panels("EstadoAplicacion").Text = "Nuevo"
   Case "Cargar"
        Informe_Carga_Operador
   Case "Buscar"
        ResizePic
        sstReferencia.Tab = 1
        StatusBar1.Panels("EstadoAplicacion").Text = "Buscar"
    Case "Control"
        Informe_Carga
        Exit Sub
    Case "ImprimirTodo"
    Case "Borrar"
    
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim Sql As String
Select Case ButtonMenu.Text
    Case "Por Caja Filtro"
        ExportarExcelporCaja True, True
    Case "Por Caja Todo"
        ExportarExcelporCaja False, True
    Case "Por Indice Todo"
        ExportarExcelReferencia False, False, False
    Case "Por Indice Filtro"
        ExportarExcelReferencia True, False, False
    Case "Por indice con index"
        ExportarExcelReferencia False, True, False
    Case "Por indice filtro con index"
        ExportarExcelReferencia True, True, False
    Case "Doc solo Filtro"
     ExportarExcelReferencia True, False, True
     Case "Control referencia"
            ImprimirControlReferencia
     Case "Control carga"
            ImprimirControlCarga
      Case "Cajas sin referencias"
            ImprimirControlCajasSinReferencias
End Select
Debug.Print ButtonMenu.Text
End Sub

Private Sub trvIndices_Click()
'  Dim I As Integer
'  If picCargar.Visible Then
'        With trvIndices.Nodes
'           For I = 1 To .Count
'              If .Item(I).Selected Then
'                  txtIndice.Text = Mid(.Item(I).Key, 2)
'                  txtIndice.SetFocus
'                  Exit Sub
'              End If
'           Next
'        End With
'  End If
End Sub

Private Sub FiltroReferencia(Indice As Boolean, Optional INDICEFIJO As Boolean)
 
    Dim Filtro As String
    Rem MsgBox trvIndices.SelectedItem.Text
    MousePointer = 11
    If Indice = True Then
        If cltIndice1.Item_Selecionado = "AIZ" Then
            Filtro = ""
        Else
            Filtro = cltIndice1.Item_Selecionado
        End If
    Else
        Filtro = ""
    End If
    
    Set rsReferencias = New ADODB.Recordset
       Dim sSQL As String
       Dim Orden As String
       Dim Item As Integer
           
           rsReferencias.CursorLocation = adUseClient
           sSQL = "Select * from Referencias where cod_Cliente =" & ctlCliente.Valor
           Filtro_Reporte = " and REFERENCIAS.cod_Cliente =" & ctlCliente.Valor
           If Filtro = "" Then
            
           Else
              If INDICEFIJO = True Then
                    sSQL = sSQL & vbCrLf & " AND indice LIKE '" & Filtro & "'"
                    Filtro_Reporte = Filtro_Reporte & " AND REFERENCIAS.indice LIKE '" & Filtro & "'"
                    Filtro_Indice_Reporte = " AND indice LIKE '" & Filtro & "'"
                    Filtro_Indice = " AND indice LIKE '" & Filtro & "%'"
               Else
                    sSQL = sSQL & vbCrLf & " AND indice LIKE '" & Filtro & "%'"
                    Filtro_Reporte = Filtro_Reporte & " AND REFERENCIAS.indice LIKE '" & Filtro & "%'"
                    Filtro_Indice_Reporte = " AND indice LIKE '" & Filtro & "%'"
                    Filtro_Indice = " AND indice LIKE '" & Filtro & "%'"
               End If
           End If
           If Sql_Filtro_Referencia <> "" Then
                sSQL = sSQL & Sql_Filtro_Referencia
                Filtro_Reporte = Filtro_Reporte & Sql_Filtro_Referencia
           End If
           If Sql_Filtro_Orden <> "" Then
            sSQL = sSQL & vbCrLf & Sql_Filtro_Orden
           Else
           
           End If
           grdReferencias.Refresh
           rsReferencias.Open sSQL, ConActiva, adOpenDynamic, adLockOptimistic
           lblCantidadRegistro.Caption = rsReferencias.RecordCount
           Set grdReferencias.DataSource = rsReferencias.DataSource
           grdReferencias.DataMember = rsReferencias.DataMember
     Rem     ConfigurarGrilla Filtro, ctlCliente.Valor
           txtdescripcionBuscar.DataMember = rsReferencias.DataMember
           Set txtdescripcionBuscar.DataSource = rsReferencias.DataSource
           txtdescripcionBuscar.DataField = "Descripcion"
           txtCodigo.DataMember = rsReferencias.DataMember
           Set txtCodigo.DataSource = rsReferencias.DataSource
           txtCodigo.DataField = "Indice"
           grdReferencias.Refresh
           
    MousePointer = 0

End Sub

Private Sub trvIndices_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
       If Not IsNull(ctlPersonal.Valor) Then
          PopupMenu mnuArbol
          InsertarProducion ctlPersonal.Valor, 18, "Buscar Refrencia", 1, ctlCliente.Valor
        Else
          MsgBox "Insertar personal", vbInformation
        End If
    End If

End Sub

Private Sub txtApellido_Nombre_GotFocus()
 Rem   StatusBar1.Panels("Ayuda").Text = txtApellido_Nombre.ToolTipText
End Sub

Private Sub txtApellido_Nombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub txtBuscarCajas_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'       Set rsReferencias = New ADODB.Recordset
'       Dim sSQL As String
'       Dim Item As Integer
'           If txtBuscarCajas = "" Then
'               Exit Sub
'           End If
'           rsReferencias.CursorLocation = adUseClient
'           sSQL = "Select * from Referencias where cod_Cliente =" & ctlCliente.Valor
'           sSQL = sSQL & vbCrLf & " AND NRO_CAJA =" & txtBuscarCajas.Text
'           sSQL = sSQL & vbCrLf & " ORDER BY INDICE"
'           rsReferencias.Open sSQL,ConActiva, adOpenDynamic, adLockOptimistic
'           Set grdModificacion.DataSource = rsReferencias.DataSource
'           grdModificacion.DataMember = rsReferencias.DataMember
'     End If
End Sub

Private Sub txtCodigo_Change()
Dim rs As ADODB.Recordset


lblCodigo.Caption = ""
If Len(txtCodigo.Text) > 5 Then
    Set rs = New ADODB.Recordset
    rs.Open " SELECT * from INDICES WHERE COD_CLIENTE =" & ctlCliente.Valor & " AND INDICE = '" & txtCodigo.Text & "'", ConActiva, 0, 1
     If Not rs.EOF Then
        lblCodigo.Caption = rs!ID_CODIGO_DOCUMENTO & " - " & Trim(rs!Descripcion)
     End If
    Set rs = New ADODB.Recordset
    rs.Open " SELECT * from INDICES WHERE COD_CLIENTE =" & ctlCliente.Valor & " AND INDICE = '" & txtCodigo.Text & "'", ConActiva, 0, 1
     If Not rs.EOF Then
        lblCodigo.Caption = lblCodigo.Caption & " // " & rs!Descripcion
     End If
Else
Set rs = New ADODB.Recordset
    rs.Open " SELECT * from INDICES WHERE COD_CLIENTE =" & ctlCliente.Valor & " AND INDICE = '" & txtCodigo.Text & "'", ConActiva, 0, 1
     If Not rs.EOF Then
        lblCodigo.Caption = rs!Descripcion
     End If
End If



End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Mid(txtCodigo.Text, 1, 1) <> "0" Then
        txtCodigo.Text = BuscarIDDocumento(txtCodigo.Text, ctlCliente.Valor)
    End If
End If

End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 123 Then
If Mid(txtCodigo.Text, 1, 1) <> "" Then
            Dim strIndice As String
            strIndice = BuscarIDDocumento(txtCodigo.Text, ctlCliente.Valor)
            If strIndice = "ERROR" Then
                MsgBox "ERROR EN EL NUMERO DE DOCUMENTO"
                txtCodigo.Text = ""
            Else
                txtCodigo.Text = strIndice
            End If
        End If
End If
End Sub

Private Sub txtDato_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 AddFiltro
End If
End Sub

Private Sub txtDescripcion_GotFocus()
    StatusBar1.Panels("Ayuda").Text = txtDescripcion.ToolTipText
    
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
        If KeyAscii = 46 Then
         MsgBox "No esta permitido Usar . ni abreviaciones ", vbInformation
         KeyAscii = 0
        End If
        
        If KeyAscii = 45 Then
            Dim rs As New ADODB.Recordset
            Dim Sql As String
            Dim i As Integer
                KeyAscii = 0
                Sql = " SELECT DESCRIPCION, COUNT(*) AS CANTIDAD"
                Sql = Sql & " From REFERENCIAS "
                Sql = Sql & "  WHERE COD_CLIENTE = " & ctlCliente.Valor
                Sql = Sql & " AND INDICE ='" & TxtIndice & "'"
                Sql = Sql & " GROUP BY DESCRIPCION"
                Sql = Sql & "  HAVING DESCRIPCION LIKE '%" & UCase(txtDescripcion.Text) & "%'"
                Sql = Sql & " ORDER BY COUNT(*) DESC"
                rs.Open Sql, ConActiva, 0, 1
                i = 1
                grdDescripcionRepetida.Clear
                grdDescripcionRepetida.Rows = 1
                grdDescripcionRepetida.Cols = 3
                grdDescripcionRepetida.ColWidth(0) = 300
                grdDescripcionRepetida.ColWidth(1) = 5000
                grdDescripcionRepetida.ColWidth(2) = 500
                    Do While Not rs.EOF
                        grdDescripcionRepetida.AddItem i & vbTab & rs!Descripcion & vbTab & rs!cantidad
                        rs.MoveNext
                        i = i + 1
                    Loop
            
                grdDescripcionRepetida.Visible = True
            End If
        
        If KeyAscii = 13 Then
'            If Trim(txtDescripcion.Text) <> "" Then
'                    Dim BANDERA As Boolean
'                    MousePointer = 11
'                    BANDERA = False
'                    If oWordL.Documents.Count = 0 Then
'                        BANDERA = False
'                        Set oTmpDocL = New Word.Document
'                        Set oTmpDocL = oWordL.Documents.Add
'                    Else
'                        BANDERA = True
'                        Set oTmpDocL = oWordL.Documents.Item(1)
'                    End If
'                    txtDescripcion.SelStart = 0
'                    txtDescripcion.SelLength = Len(txtDescripcion.Text)
'                    Clipboard.Clear
'                    Clipboard.SetText txtDescripcion.SelText
'                    With oTmpDocL
'                        .Content.Paste
'                        .Activate
'                        oWordL.WindowState = wdWindowStateNormal
'                        oWordL.top = 3000
'                        oWordL.left = 4000
'                        oWordL.Visible = True
'                        If BANDERA = True Then
'                            oWordL.Activate
'                        End If
'                        Rem Tasks.Item("Microsoft Word").Activate
'
'                        MousePointer = 0
'                        .CheckSpelling
'
'                        .Content.Copy
'                        txtDescripcion.Text = Clipboard.GetText(vbCFText)
'                        .Saved = True
'                    End With
'                    Dim I As Integer
'                    '                For i = 1 To Tasks.Count
'                    '                 Debug.Print Tasks.Item(i).Name
'                    '                Next
'                    Rem Tasks.Item("Sistema Basa").Activate
'            End If
            KeyAscii = 0
            cmdAceptar.SetFocus
        End If
End Sub



Private Sub txtIndice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(TxtIndice.Text) = "" Then
            MsgBox "Error Indice"
            Exit Sub
            
        End If
        
        If Mid(TxtIndice.Text, 1, 1) <> 0 Then
            Dim strIndice As String
            strIndice = BuscarIDDocumento(TxtIndice.Text, ctlCliente.Valor)
            If strIndice = "ERROR" Then
                MsgBox "ERROR EN EL NUMERO DE DOCUMENTO"
                TxtIndice.Text = ""
            Else
                TxtIndice.Text = strIndice
            End If
        Else
        
        End If
        SendKeys vbTab
    End If
End Sub

Private Sub INDICE_CONFIG()
    Dim rsIndice As New ADODB.Recordset
    Dim sSQL As String
    Dim Descripcion As String
      If TxtIndice = "" Then
        Exit Sub
      End If
      'Exit Sub
        sSQL = "  SELECT COD_CLIENTE, INDICE, DESCRIPCION, FECHA, NUMERO,LETRA, EXPEDIENTE, APELLIDO_NOMBRE,"
        sSQL = sSQL & vbCrLf & "  MASK_EXPEDIENTE , MASK_LETRA, TOOLTIPFECHA, TOOLTIPNUMERO, TOOLTIPLETRA, TOOLTIPEXPEDIENTE, TOOLTIPAPELLIDO_NOMBRE , TOOLTIPDESCRIPCION"
        sSQL = sSQL & vbCrLf & " FROM INDICES "
        sSQL = sSQL & vbCrLf & "  Where Cod_Cliente =" & ctlCliente.Valor & " AND indice  = '" & TxtIndice.Text & "'"
        Set rsIndice = New ADODB.Recordset
        rsIndice.Open sSQL, ConActiva, 0, 1
        With rsIndice
        If .EOF Then
                lblNombre.Caption = ""
                txtDescripcion.BackColor = &HC0C0FF
                mskFecha_Desde.Enabled = False
                mskFecha_Desde.BackColor = &HC0C0FF
                mskFecha_Hasta.Enabled = False
                mskFecha_Hasta.BackColor = &HC0C0FF
                mskNro_desde.Enabled = False
                mskNro_desde.BackColor = &HC0C0FF
                mskNro_hasta.Enabled = False
                mskNro_hasta.BackColor = &HC0C0FF
                mskLetra_Desde.Enabled = False
                mskLetra_Desde.BackColor = &HC0C0FF
                mskLetra_Hasta.Enabled = False
                mskLetra_Hasta.BackColor = &HC0C0FF
               Rem mskExpediente.Enabled = False
               Rem mskExpediente.BackColor = &HC0C0FF
               Rem txtApellido_Nombre.Enabled = False
               Rem txtApellido_Nombre.BackColor = &HC0C0FF
         Exit Sub
        End If
            
            
         
            lblNombre.Caption = Trim(!Descripcion)
            If StatusBar1.Panels.Item("EstadoAplicacion").Text = "Nuevo" Then
               Rem  BuscarIndice txtIndice.Text, True
            End If
            If IsNull(!TOOLTIPDESCRIPCION) Then
                txtDescripcion.ToolTipText = "NO tiene Ayuda"
                txtDescripcion.BackColor = &H80000005
            Else
                txtDescripcion.ToolTipText = Trim(!TOOLTIPDESCRIPCION)
                txtDescripcion.BackColor = &HF2FFEA
            End If
            
            If !fecha = 1 Then
                mskFecha_Desde.Enabled = True
                mskFecha_Desde.BackColor = &HF2FFEA
                mskFecha_Hasta.Enabled = True
                mskFecha_Hasta.BackColor = &HF2FFEA
                If Not IsNull(!TOOLTIPFECHA) Then
                    mskFecha_Desde.ToolTipText = !TOOLTIPFECHA
                    mskFecha_Hasta.ToolTipText = !TOOLTIPFECHA
                Else
                    mskFecha_Desde.ToolTipText = "NO tiene Ayuda"
                    mskFecha_Hasta.ToolTipText = "NO tiene Ayuda"
                End If
             Else
                mskFecha_Desde.Enabled = False
                mskFecha_Desde.BackColor = &HC0C0FF
                mskFecha_Hasta.Enabled = False
                mskFecha_Hasta.BackColor = &HC0C0FF
            End If
            
            If !NUMERO = 1 Then
                mskNro_desde.Enabled = True
                mskNro_desde.BackColor = &HF2FFEA
                mskNro_hasta.Enabled = True
                mskNro_hasta.BackColor = &HF2FFEA
                If Not IsNull(!TOOLTIPNUMERO) Then
                    mskNro_desde.ToolTipText = !TOOLTIPNUMERO
                    mskNro_hasta.ToolTipText = !TOOLTIPNUMERO
                Else
                    mskNro_desde.ToolTipText = "NO tiene Ayuda"
                    mskNro_hasta.ToolTipText = "NO tiene Ayuda"
                End If
                    
             Else
                mskNro_desde.Enabled = False
                mskNro_desde.BackColor = &HC0C0FF
                mskNro_hasta.Enabled = False
                mskNro_hasta.BackColor = &HC0C0FF
            End If
            
            If !lETRA = 1 Then
                mskLetra_Desde.Enabled = True
                mskLetra_Desde.BackColor = &HF2FFEA
                mskLetra_Hasta.Enabled = True
                mskLetra_Hasta.BackColor = &HF2FFEA
                If IsNull(rsIndice!MASK_LETRA) Then
                    mskLetra_Desde.Mask = ""
                    mskLetra_Desde.Text = ""
                    mskLetra_Hasta.Mask = ""
                    mskLetra_Hasta.Text = ""
                    mskLetra_Desde.ToolTipText = ""
                    mskLetra_Hasta.ToolTipText = ""
                    If IsNull(rsIndice!TOOLTIPLETRA) Then
                        mskLetra_Desde.ToolTipText = "NO tiene Ayuda"
                        mskLetra_Hasta.ToolTipText = "NO tiene Ayuda"
                    Else
                        mskLetra_Desde.ToolTipText = Trim(rsIndice!TOOLTIPLETRA)
                        mskLetra_Hasta.ToolTipText = Trim(rsIndice!TOOLTIPLETRA)
                    End If
                Else
                    mskLetra_Desde.Mask = rsIndice!MASK_LETRA
                    mskLetra_Hasta.Mask = rsIndice!MASK_LETRA
                    If IsNull(rsIndice!TOOLTIPLETRA) Then
                        mskLetra_Desde.ToolTipText = "NO tiene Ayuda"
                        mskLetra_Hasta.ToolTipText = "NO tiene Ayuda"
                    Else
                        mskLetra_Desde.ToolTipText = Trim(rsIndice!TOOLTIPLETRA)
                        mskLetra_Hasta.ToolTipText = Trim(rsIndice!TOOLTIPLETRA)
                    End If
                End If
             Else
                mskLetra_Desde.Enabled = False
                mskLetra_Desde.BackColor = &HC0C0FF
                mskLetra_Hasta.Enabled = False
                mskLetra_Hasta.BackColor = &HC0C0FF
            End If
            
'            If !EXPEDIENTE = 1 Then
'                mskExpediente.Enabled = True
'                mskExpediente.BackColor = &HF2FFEA
'                If IsNull(rsIndice!MASK_EXPEDIENTE) Then
'                    mskExpediente.Mask = ""
'                    mskExpediente.Text = ""
'                    mskExpediente.ToolTipText = ""
'                Else
'                    mskExpediente.Mask = rsIndice!MASK_EXPEDIENTE
'                    If IsNull(rsIndice!TOOLTIPEXPEDIENTE) Then
'                        mskExpediente.ToolTipText = "No tienen Ayuda"
'                    Else
'                        mskExpediente.ToolTipText = Trim(rsIndice!TOOLTIPEXPEDIENTE)
'                    End If
'                End If
'            Else
'                mskExpediente.Enabled = False
'                mskExpediente.BackColor = &HC0C0FF
'            End If
              
'            If !APELLIDO_NOMBRE = 1 Then
'                txtApellido_Nombre.Enabled = True
'                txtApellido_Nombre.BackColor = &HF2FFEA
'                If IsNull(rsIndice!TOOLTIPAPELLIDO_NOMBRE) Then
'                    txtApellido_Nombre.ToolTipText = "No tienen Ayuda"
'                Else
'                    txtApellido_Nombre.ToolTipText = Trim(rsIndice!TOOLTIPAPELLIDO_NOMBRE)
'                End If
'            Else
'                txtApellido_Nombre.Enabled = False
'                txtApellido_Nombre.BackColor = &HC0C0FF
'            End If
        
        
             
        
        End With
        
End Sub

Private Sub txtIndice_LostFocus()
    INDICE_CONFIG
End Sub

Private Sub txtNro_Caja_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 123 Then
    Buscar_Inidice_Por_caja
 End If
 
End Sub

Private Sub txtNro_Caja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNro_Caja.Text) = "" Then
            MsgBox "Error en caja", vbInformation
            Exit Sub
        End If
        INDICE_CONFIG
        SendKeys vbTab
    End If
End Sub

Private Sub txtNro_Caja_LostFocus()
  Set rsReferencias = New ADODB.Recordset
  Dim rsIndice As New ADODB.Recordset
  
  Dim sSQL As String
  Dim Item As Integer
         If txtNro_Caja = "" Then
            Exit Sub
        End If
        rsReferencias.CursorLocation = adUseClient
    Dim rs As New ADODB.Recordset
        sSQL = "  SELECT ESTADO"
        sSQL = sSQL & vbCrLf & "  From CONTENEDOR"
        sSQL = sSQL & vbCrLf & "  Where NRO_CAJA = " & txtNro_Caja
        sSQL = sSQL & vbCrLf & " And COD_CLIENTE = " & ctlCliente.Valor
        rs.Open sSQL, ConActiva, 0, 1
        
        
        
        If rs.EOF Then
            MsgBox "El cliente No tiene esta caja", vbCritical
        Else
            If rs!estado = 2 Then
            
            Else
             MsgBox "El Estado es incorrecto", vbInformation
            End If
        End If
        
        ActualizarGrillaCarga

       
       
End Sub

Public Sub GrabarReferencias()
    Dim COD_CLIENTE, NRO_CAJA, Item, Indice, Descripcion     As String
    Dim FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA As String
    Dim LETRA_DESDE, LETRA_HASTA, EXPEDIENTE, APELLIDO_NOMBRE As String
    Dim FECHA_MODIFICACION, USUARIO_MODIFICACION As String
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
        
        If Not IsNumeric(txtNro_Caja.Text) Then
            MsgBox "Usted Debe ingresar la Caja"
            Exit Sub
        Else
            NRO_CAJA = txtNro_Caja.Text
        End If
        
        If Mid(TxtIndice.Text, 1, 1) <> "0" Then
            MsgBox "Error indice"
            Exit Sub
         Else
            Indice = "'" & TxtIndice & "'"
        End If
        
        If txtDescripcion.Text = "" Then
          Descripcion = "NULL"
        Else
           Descripcion = "'" & Replace(UCase(Trim(Replace(txtDescripcion.Text, vbCrLf, " "))), vbCrLf, " ") & "'"
        End If
        
        If mskFecha_Desde.ClipText = "" Then
            FECHA_DESDE = "NULL"
        Else
            FECHA_DESDE = "'" & mskFecha_Desde.Text & "'"
        End If
        
        If mskFecha_Hasta.ClipText = "" Then
            FECHA_HASTA = "NULL"
        Else
            FECHA_HASTA = "'" & mskFecha_Hasta.Text & "'"
        End If
      
        If Not IsNumeric(mskNro_desde.Text) Then
            NRO_DESDE = "NULL"
        Else
            NRO_DESDE = mskNro_desde.Text
        End If
        
        If Not IsNumeric(mskNro_hasta.Text) Then
            NRO_HASTA = "NULL"
        Else
            NRO_HASTA = mskNro_hasta.Text
        End If
        
        If mskLetra_Desde.Text = "" Then
            LETRA_DESDE = "NULL"
        Else
            LETRA_DESDE = "'" & mskLetra_Desde.Text & "'"
        End If
        
        If mskLetra_Hasta.Text = "" Then
            LETRA_HASTA = "NULL"
        Else
            LETRA_HASTA = "'" & mskLetra_Hasta.Text & "'"
        End If
          
'        If mskExpediente.Text = "" Then
'          EXPEDIENTE = "NULL"
'        Else
'           EXPEDIENTE = "'" & mskExpediente.Text & "'"
'        End If
''
'        If txtApellido_Nombre.Text = "" Then
'            APELLIDO_NOMBRE = "NULL"
'        Else
'            APELLIDO_NOMBRE = "'" & Trim(UCase(txtApellido_Nombre.Text)) & "'"
'        End If
        FECHA_MODIFICACION = SysDateMinutoSegundo
         
         ID_UNITER = 0

        
        If Not IsNull(ctlPersonal.Valor) Then
            Usuario = ctlPersonal.Valor
        Else
            MsgBox "Ingrese quien carga", vbInformation
            Exit Sub
        End If
        
        If lbl_ID_imagen.Caption = "" Then
            ID_imagen = "Null"
        Else
            ID_imagen = lbl_ID_imagen.Caption
             InsertarImagenes CLng(ID_imagen), ctlCliente.Valor, CLng(NRO_CAJA), 1, SysDate2
        End If
        
Select Case StatusBar1.Panels("EstadoAplicacion").Text
Case "Nuevo"
        
        sSQL = "    INSERT INTO REFERENCIAS"
        sSQL = sSQL & vbCrLf & "        (COD_ID_REFERENCIA, COD_CLIENTE, NRO_CAJA, ITEM, INDICE, DESCRIPCION,"
        sSQL = sSQL & vbCrLf & "        FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA,"
        sSQL = sSQL & vbCrLf & "        LETRA_DESDE, LETRA_HASTA, "
        sSQL = sSQL & vbCrLf & "         FECHA_MODIFICACION,"
        sSQL = sSQL & vbCrLf & "        USUARIO_MODIFICACION,borrado,ID_UNITER, ESTADO ,ID_IMAGEN )"
        sSQL = sSQL & vbCrLf & "    Values"
        sSQL = sSQL & vbCrLf & "  (" & Max_Cod_Id_Referencia & "," & COD_CLIENTE & "," & NRO_CAJA & ",0," & Indice & "," & Descripcion & ","
        sSQL = sSQL & vbCrLf & FECHA_DESDE & "," & FECHA_HASTA & "," & NRO_DESDE & "," & NRO_HASTA & ","
        sSQL = sSQL & vbCrLf & LETRA_DESDE & "," & LETRA_HASTA & ","
        sSQL = sSQL & vbCrLf & FECHA_MODIFICACION & ","
        sSQL = sSQL & vbCrLf & "'" & Usuario & "',0," & ID_UNITER & ",2," & ID_imagen & ")"
        ExecutarSql (sSQL)
          
Case "Modificar"

If IsNumeric(lblCod_Referencia.Caption) Then
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
    sSQL = sSQL & vbCrLf & " Where COD_ID_REFERENCIA = " & lblCod_Referencia.Caption
    ExecutarSql sSQL
Else
    MsgBox "Error en la actualizacion" & vbCrLf & "verifique si el estado de la aplicacion es modicacion", vbInformation
End If
End Select
    Refrescar
    LimpiarTodos
    txtNro_Caja.SetFocus
Exit Sub
salir:

MsgBox Err.Description

End Sub

Public Sub LimpiarTodos()
   Rem txtApellido_Nombre.Text = ""
   lblCod_Referencia.Caption = ""
    If chkDescripcion.value <> 1 Then
        txtDescripcion.Text = ""
    End If
    If chkIndice.value <> 1 Then
        TxtIndice.Text = ""
    End If
   
    If chkCaja.value <> 1 Then
        txtNro_Caja.Text = ""
    End If
        
    Rem txtUnit.Text = ""
    If chkNº_Desde.value <> 1 Then
        mskNro_desde.Mask = ""
        mskNro_desde.Text = ""
    End If
    
    If chkNº_Hasta.value <> 1 Then
        mskNro_hasta.Mask = ""
        mskNro_hasta.Text = ""
    End If
    
    
    If chkLetra_Desde.value <> 1 Then
        mskLetra_Desde.Mask = ""
        mskLetra_Desde = ""
    End If
    
    If chkLetra_Hasta.value <> 1 Then
        mskLetra_Hasta.Mask = ""
        mskLetra_Hasta = ""
    End If
    
'    mskExpediente.Mask = ""
'    mskExpediente.Text = ""
    
    If chkFecha_Desde.value <> 1 Then
        mskFecha_Desde.PromptInclude = False
        mskFecha_Desde.Text = ""
        mskFecha_Desde.PromptInclude = True
    End If
    
    If chkFechaHasta.value <> 1 Then
        mskFecha_Hasta.PromptInclude = False
        mskFecha_Hasta.Text = ""
        mskFecha_Hasta.PromptInclude = True
    End If
    
    lblEstadoReferencia.Caption = ""
    

End Sub

Public Sub Refrescar()
    Set rsReferencias = New ADODB.Recordset
    Dim sSQL As String
    Dim Item As Integer
            rsReferencias.CursorLocation = adUseClient
            sSQL = "Select * from Referencias where cod_Cliente =" & ctlCliente.Valor
            sSQL = sSQL & vbCrLf & " AND NRO_CAJA =" & txtNro_Caja
            sSQL = sSQL & vbCrLf & " ORDER BY INDICE"
            rsReferencias.Open sSQL, ConActiva, 0, 1
            Set grdReferencias.DataSource = rsReferencias.DataSource
            grdReferencias.DataMember = rsReferencias.DataMember
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub





Public Sub ConfigurarGrilla(FiltroIndice As String, COD_CLIENTE As Integer)
Dim sSQL As String
Dim rsIndiceConfig As ADODB.Recordset
    Set rsIndiceConfig = New ADODB.Recordset
   
        sSQL = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE,"
        sSQL = sSQL & vbCrLf & " DESCRIPCION, FECHA, NUMERO, LETRA, EXPEDIENTE, APELLIDO_NOMBRE"
        sSQL = sSQL & vbCrLf & " From INDICES"
        sSQL = sSQL & vbCrLf & " WHERE (COD_CLIENTE = " & COD_CLIENTE & ") AND (INDICE =  '" & FiltroIndice & "')"
        rsIndiceConfig.Open sSQL, ConActiva, 0, 1
        With rsIndiceConfig
        If Not .EOF Then
            If IsNull(!fecha) Then
                grdReferencias.Columns.Item(4).Visible = False
                grdReferencias.Columns.Item(5).Visible = False
            Else
                grdReferencias.Columns.Item(4).Visible = True
                grdReferencias.Columns.Item(5).Visible = True
            End If
            If IsNull(!NUMERO) Then
                grdReferencias.Columns.Item(6).Visible = False
                grdReferencias.Columns.Item(7).Visible = False
            Else
                grdReferencias.Columns.Item(6).Visible = True
                grdReferencias.Columns.Item(7).Visible = True
            End If
            If IsNull(!lETRA) Then
                grdReferencias.Columns.Item(8).Visible = False
                grdReferencias.Columns.Item(9).Visible = False
            Else
                grdReferencias.Columns.Item(8).Visible = True
                grdReferencias.Columns.Item(9).Visible = True
            End If
            If IsNull(!EXPEDIENTE) Then
                grdReferencias.Columns.Item(10).Visible = False
            Else
                grdReferencias.Columns.Item(10).Visible = True
            End If
            If IsNull(!APELLIDO_NOMBRE) Then
                grdReferencias.Columns.Item(11).Visible = False
            Else
                grdReferencias.Columns.Item(11).Visible = True
            End If
        End If
  End With
  grdReferencias.Refresh
End Sub

Public Function Orden_Referencia(Cliente As Integer, Optional Orden As String) As String
    Dim rsOrden As ADODB.Recordset
    Set rsOrden = New ADODB.Recordset
    Dim Indice As String
    Dim Sql As String

If cltIndice1.Item_Selecionado = "AIZ" Then
    Orden_Referencia = "ORDER BY FECHA_DESDE , NRO_DESDE"
Else
    
  Indice = cltIndice1.Item_Selecionado
  
    
    Orden_Referencia = ""
     Sql = " SELECT SQL_ORDEN,FECHA, NUMERO, LETRA  From INDICES WHERE (COD_CLIENTE =" & Cliente & ") AND (INDICE = '" & Indice & "')"
     
     
    rsOrden.Open Sql, ConActiva, 0, 1
    If Not rsOrden.EOF Then
        If IsNull(rsOrden!SQL_ORDEN) Then
            Orden_Referencia = ""
            If Not IsNull(rsOrden!NUMERO) Then
                   Orden_Referencia = " ORDER BY  NRO_DESDE ASC"
            End If
            If Not IsNull(rsOrden!fecha) Then
                   Orden_Referencia = " ORDER BY  FECHA_DESDE ASC"
            End If
            If Not IsNull(rsOrden!lETRA) Then
                   Orden_Referencia = " ORDER BY LETRA_DESDE ASC"
            End If
            
        Else
            Orden_Referencia = rsOrden!SQL_ORDEN
        End If
    End If
End If
End Function


Public Sub ExportarExcelReferencia(Filtro As Boolean, Indice As Boolean, DocSolo As Boolean)
   Dim Sql As String
   Dim DATO As String
   Dim rsbasa As New ADODB.Recordset
   Dim oConn As ADODB.Connection
   Dim oRS As ADODB.Recordset
   Dim i As Integer
   Dim TITULOHERANT As String
   Dim PONERTITULOHER As String
   
   On Error GoTo er
            MousePointer = 11
   
            
            
    Rem        Sql = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE,DESCRIPCION,CANTIDAD_CAJAS_ACUMULADO , CANTIDAD_CAJAS_SOLO ,    TIPO_INDICE"
     Rem        Sql = Sql & "  From INDICES"
      Rem       Sql = Sql & "  Where (TIPO_INDICE = 'Sector') and  COD_CLIENTE = " & ctlCliente.Valor
            
            
            
        Sql = " SELECT COD_CLIENTE,TIPO_INDICE, ID_CODIGO_DOCUMENTO, INDICE,DESCRIPCION"
        Sql = Sql & "  From INDICES"
        Sql = Sql & "  Where COD_CLIENTE = " & ctlCliente.Valor
        If Filtro = True And Filtro_Indice_Reporte <> "" Then
            Sql = Sql & Filtro_Indice
        End If
        Sql = Sql & "  ORDER BY INDICE"
       rsbasa.Open Sql, ConActiva, 0, 1
       Set oConn = New ADODB.Connection
       If Dir("C:\Referencia.xls") <> "" Then
           Kill "C:\Referencia.xls"
       End If
       FileCopy strPasoPlanillas & "Referencia.xls", "C:\Referencia.xls"
       oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source=C:\Referencia.xls;" & _
                  "Extended Properties=""Excel 8.0;HDR=NO;"""
       
       '------------------ INICIO INDICE ----------------------------------
       
       Set oRS = New ADODB.Recordset
       oRS.Open "Select * from Indice", oConn, adOpenKeyset, adLockOptimistic
       oRS.MoveFirst
       oRS.Fields(2).value = "" & UCase(ctlCliente.Descripcion)
       oRS.Update
       oRS.MoveNext
       Dim Contenido As String
       Do While Not rsbasa.EOF
            If Indice = True Then
                Contenido = "DOC:" & Format(rsbasa!ID_CODIGO_DOCUMENTO, "000") & "  ID:" & rsbasa!Indice & " " & rsbasa!Descripcion
                Contenido = "Indice:" & rsbasa!Indice & "  DOC:" & Format(rsbasa!ID_CODIGO_DOCUMENTO, "000") & " " & rsbasa!Descripcion
            Else
              If DocSolo = True Then
                If Trim(rsbasa!Tipo_Indice) = "Documento" Then
                    Contenido = Format(rsbasa!ID_CODIGO_DOCUMENTO, "000") & " - " & rsbasa!Descripcion
                Else
                    
                   Contenido = Format(rsbasa!ID_CODIGO_DOCUMENTO, "000") & " - " & rsbasa!Descripcion
                    Rem Contenido = rsbasa!DESCRIPCION
                End If
              Else
                Contenido = "DOC:" & Format(rsbasa!ID_CODIGO_DOCUMENTO, "000") & " " & rsbasa!Descripcion
              End If
                Rem  Contenido = Contenido & " // CANT. DE CAJAS ACUMULADO : " & rsbasa!CANTIDAD_CAJAS_ACUMULADO & " // CANT. CAJAS INDICE SOLO :" & rsbasa!CANTIDAD_CAJAS_SOLO
               Rem   Contenido = Contenido & " // CANT. DE CAJAS ACUMULADO : " & rsbasa!CANTIDAD_CAJAS_ACUMULADO & " // CANT. CAJAS INDICE SOLO :" & rsbasa!CANTIDAD_CAJAS_SOLO
            End If
            Select Case Len(rsbasa!Indice) / 3
            Case 1
                oRS.Fields(0).value = Contenido
            Case 2
                oRS.Fields(0).value = "- - - - - - - - "
                oRS.Fields(1).value = Contenido
            Case 3
                oRS.Fields(0).value = "- - - - - - - - "
                oRS.Fields(1).value = "- - - - - - - - "
                oRS.Fields(2).value = Contenido
            Case 4
                oRS.Fields(0).value = "- - - - - - - - "
                oRS.Fields(1).value = "- - - - - - - - "
                oRS.Fields(2).value = "- - - - - - - - "
                oRS.Fields(3).value = Contenido
            Case 5
                oRS.Fields(0).value = "- - - - - - - - "
                oRS.Fields(1).value = "- - - - - - - - "
                oRS.Fields(2).value = "- - - - - - - - "
                oRS.Fields(3).value = "- - - - - - - - "
                oRS.Fields(4).value = Contenido
            Case 6
                oRS.Fields(0).value = "- - - - - - - - "
                oRS.Fields(1).value = "- - - - - - - - "
                oRS.Fields(2).value = "- - - - - - - - "
                oRS.Fields(3).value = "- - - - - - - - "
                oRS.Fields(4).value = "- - - - - - - - "
                oRS.Fields(5).value = Contenido
            Case 7
                oRS.Fields(0).value = "- - - - - - - - "
                oRS.Fields(1).value = "- - - - - - - - "
                oRS.Fields(2).value = "- - - - - - - - "
                oRS.Fields(3).value = "- - - - - - - - "
                oRS.Fields(4).value = "- - - - - - - - "
                oRS.Fields(5).value = "- - - - - - - - "
                oRS.Fields(6).value = Contenido
            Case 8
                oRS.Fields(0).value = "- - - - - - - - "
                oRS.Fields(1).value = "- - - - - - - - "
                oRS.Fields(2).value = "- - - - - - - - "
                oRS.Fields(3).value = "- - - - - - - - "
                oRS.Fields(4).value = "- - - - - - - - "
                oRS.Fields(5).value = "- - - - - - - - "
                oRS.Fields(6).value = "- - - - - - - - "
                oRS.Fields(7).value = Contenido
            Case 9
                oRS.Fields(0).value = "- - - - - - - - "
                oRS.Fields(1).value = "- - - - - - - - "
                oRS.Fields(2).value = "- - - - - - - - "
                oRS.Fields(3).value = "- - - - - - - - "
                oRS.Fields(4).value = "- - - - - - - - "
                oRS.Fields(5).value = "- - - - - - - - "
                oRS.Fields(6).value = "- - - - - - - - "
                oRS.Fields(7).value = "- - - - - - - - "
                oRS.Fields(8).value = Contenido
            Case 10
                oRS.Fields(0).value = "- - - - - - - - "
                oRS.Fields(1).value = "- - - - - - - - "
                oRS.Fields(2).value = "- - - - - - - - "
                oRS.Fields(3).value = "- - - - - - - - "
                oRS.Fields(4).value = "- - - - - - - - "
                oRS.Fields(5).value = "- - - - - - - - "
                oRS.Fields(6).value = "- - - - - - - - "
                oRS.Fields(7).value = "- - - - - - - - "
                oRS.Fields(8).value = "- - - - - - - - "
                oRS.Fields(9).value = Contenido
            End Select
            oRS.Update
            oRS.MoveNext
            rsbasa.MoveNext
        Loop
    '------------------------- FIN INCIDE ----------------------------------------------
    
    If DocSolo = False Then
            ' ------------------------INICIO REFERENCIA ---------------------------------------
            Set oRS = New ADODB.Recordset
            oRS.Open "Select * from Referencias", oConn, adOpenKeyset, adLockOptimistic
            Sql = " SELECT  COD_ID_REFERENCIA, INDICES.TITULOHERENCIA,INDICES.DESCRIPCION AS DESCRIPCIONINDICE , REFERENCIAS.NRO_CAJA, REFERENCIAS.ITEM,"
            Sql = Sql & vbCrLf & " REFERENCIAS.INDICE, REFERENCIAS.DESCRIPCION,"
            Sql = Sql & vbCrLf & " REFERENCIAS.COD_CLIENTE,INDICES.ID_CODIGO_DOCUMENTO,"
            Sql = Sql & vbCrLf & " REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA,"
            Sql = Sql & vbCrLf & " REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA,"
            Sql = Sql & vbCrLf & " REFERENCIAS.LETRA_DESDE, REFERENCIAS.LETRA_HASTA,"
            Sql = Sql & vbCrLf & " REFERENCIAS.EXPEDIENTE, REFERENCIAS.APELLIDO_NOMBRE, REFERENCIAS.BORRADO"
            Sql = Sql & vbCrLf & " From REFERENCIAS, INDICES"
            Sql = Sql & vbCrLf & " WHERE REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND"
            Sql = Sql & vbCrLf & " REFERENCIAS.INDICE = INDICES.INDICE "
            If Filtro = True And Filtro_Reporte <> "" Then
                Sql = Sql & vbCrLf & "  " & Filtro_Reporte
            Else
                Sql = Sql & vbCrLf & " AND REFERENCIAS.COD_CLIENTE =" & ctlCliente.Valor
            End If
            Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.INDICE, REFERENCIAS.FECHA_DESDE,REFERENCIAS.NRO_DESDE"
            Set rsbasa = New ADODB.Recordset
            rsbasa.Open Sql, ConActiva, 0, 1
            TITULOHERANT = "NO"
            Do While Not rsbasa.EOF
          
       
                If IsNull(rsbasa!TituloHerencia) Then
                    PONERTITULOHER = "DOC:" & Format(rsbasa!ID_CODIGO_DOCUMENTO, "000") & "_________________________"
                Else
                    If Indice = True Then
                        Rem PONERTITULOHER = "DOC:" & Format(rsbasa!ID_CODIGO_DOCUMENTO, "000") & "  ID:" & rsbasa!Indice & "  //" & Trim(rsbasa!TituloHerencia)
                        PONERTITULOHER = "Indice: " & rsbasa!Indice & "  DOC:" & Format(rsbasa!ID_CODIGO_DOCUMENTO, "000") & "  //" & Trim(rsbasa!TituloHerencia)
                    Else
                        PONERTITULOHER = "DOC:" & Format(rsbasa!ID_CODIGO_DOCUMENTO, "000") & "  //" & Trim(rsbasa!TituloHerencia)
                    End If
                End If
                If TITULOHERANT = "NO" Then
                    TITULOHERANT = PONERTITULOHER
                    oRS.Fields(0).value = PONERTITULOHER
                    oRS.Update
                    oRS.MoveNext
                Else
                    If TITULOHERANT = PONERTITULOHER Then
                    Else
                        TITULOHERANT = PONERTITULOHER
                        oRS.Fields(0).value = PONERTITULOHER
                        oRS.Update
                        oRS.MoveNext
                    End If
                End If
                oRS.Fields(0).value = PONERTITULOHER
                oRS.Fields(1).value = rsbasa!NRO_CAJA
                oRS.Fields(3).value = rsbasa!FECHA_DESDE
                oRS.Fields(4).value = rsbasa!FECHA_HASTA
                If Not IsNull(rsbasa!NRO_DESDE) Then
                    oRS.Fields(5).value = rsbasa!NRO_DESDE
                End If
                If Not IsNull(rsbasa!NRO_HASTA) Then
                    oRS.Fields(6).value = rsbasa!NRO_HASTA
                End If
                oRS.Fields(7).value = rsbasa!LETRA_DESDE
                oRS.Fields(8).value = rsbasa!LETRA_HASTA
                oRS.Fields(9).value = rsbasa!APELLIDO_NOMBRE
                oRS.Fields(10).value = rsbasa!EXPEDIENTE
                If Len(Trim(rsbasa!Descripcion)) > 250 Then
                    oRS.Fields(2).value = UCase(Mid(Trim(rsbasa!Descripcion), 1, 250))
                    oRS.Update
                    oRS.MoveNext
                  Rem  oRS.Fields(2).Value = UCase(Mid(Trim(rsbasa!DESCRIPCION), 251, 500))
                Else
                    oRS.Fields(2).value = UCase(Trim(rsbasa!Descripcion))
                End If
                oRS.Fields(11).value = rsbasa!COD_ID_REFERENCIA
                oRS.Update
                oRS.MoveNext
                rsbasa.MoveNext
            Loop
            '----------------------------------- FIN REFERENCIA ---------------------
    End If
    
    rsbasa.Close
    oConn.Close
    MousePointer = 0
    MsgBox "La exportacion a finalizado", vbInformation
    Exit Sub
er:
    MousePointer = 0
    MsgBox Err.Description

End Sub



Public Sub ExportarExcelporCaja(Filtro As Boolean, ID_referencia As Boolean)
   Dim Sql As String
   Dim rsbasa As New ADODB.Recordset
   
   
   
   Dim TITULOHERANT As String
   Dim PONERTITULOHER As String
   
   
   
   On Error GoTo er
   MousePointer = 11
      If Dir("C:\Referencia por caja.xls") <> "" Then
           Kill "C:\Referencia por caja.xls"
       End If
       FileCopy strPasoPlanillas & "Referencia por caja.xls", "C:\Referencia por caja.xls"
              
       Dim xlApp As Excel.Application
       Dim xlBook As Excel.Workbook
       Dim xlSheet As Excel.Worksheet
       Set xlApp = New Excel.Application
       Set xlBook = xlApp.Workbooks.Open("C:\Referencia por caja.xls")
       Set xlSheet = xlBook.Worksheets.Item(1)

      
                ' ------------------------INICIO REFERENCIA ---------------------------------------
                 Sql = " SELECT COD_ID_REFERENCIA, INDICES.TITULOHERENCIA,INDICES.DESCRIPCION AS DESCRIPCIONINDICE , REFERENCIAS.NRO_CAJA, REFERENCIAS.ITEM,"
                 Sql = Sql & vbCrLf & " REFERENCIAS.INDICE, REFERENCIAS.DESCRIPCION,"
                 Sql = Sql & vbCrLf & " REFERENCIAS.COD_CLIENTE,INDICES.ID_CODIGO_DOCUMENTO,"
                 Sql = Sql & vbCrLf & " REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA,"
                 Sql = Sql & vbCrLf & " REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA,"
                 Sql = Sql & vbCrLf & " REFERENCIAS.LETRA_DESDE, REFERENCIAS.LETRA_HASTA,"
                 Sql = Sql & vbCrLf & " REFERENCIAS.EXPEDIENTE, REFERENCIAS.APELLIDO_NOMBRE, REFERENCIAS.BORRADO"
                 Sql = Sql & vbCrLf & " From REFERENCIAS, INDICES"
                 Sql = Sql & vbCrLf & " WHERE REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND"
                 Sql = Sql & vbCrLf & " REFERENCIAS.INDICE = INDICES.INDICE AND REFERENCIAS.BORRADO <> '1'"
                 If Filtro = True And Filtro_Reporte <> "" Then
                     Sql = Sql & vbCrLf & "  " & Filtro_Reporte
                 Else
                     Sql = Sql & vbCrLf & " AND REFERENCIAS.COD_CLIENTE =" & ctlCliente.Valor
                 End If
                 Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.NRO_CAJA , REFERENCIAS.INDICE "
                 Set rsbasa = New ADODB.Recordset
                 rsbasa.Open Sql, ConActiva, 0, 1
                 TITULOHERANT = "NO"
                 Dim Descripcion As String
                 Dim Caja As Long
                 Dim Titulo As String
                 Dim R As Integer
                     xlSheet.Cells.Item(2, 3) = "REFERENCIAS POR CAJAS " & UCase(ctlCliente.Descripcion)
                     R = 3
                 Do While Not rsbasa.EOF
                     Descripcion = ""
                     If IsNull(rsbasa!TituloHerencia) Then
                         PONERTITULOHER = "Doc: " & rsbasa!ID_CODIGO_DOCUMENTO & "- - - - - - - - "
                     Else
                         PONERTITULOHER = "Doc: " & rsbasa!ID_CODIGO_DOCUMENTO & "  " & Trim(rsbasa!TituloHerencia)
                     End If
                     If Caja = 0 Then
                             Caja = rsbasa!NRO_CAJA
                             R = R + 1
                             xlSheet.Cells.Item(R, 1) = "" & rsbasa!NRO_CAJA
                             Titulo = PONERTITULOHER
                             xlSheet.Cells.Item(R, 2) = Titulo
                     End If
                     If Caja <> rsbasa!NRO_CAJA Then
                            Caja = rsbasa!NRO_CAJA
                             R = R + 1
                             xlSheet.Cells.Item(R, 1) = "" & rsbasa!NRO_CAJA
                             Titulo = PONERTITULOHER
                             xlSheet.Cells.Item(R, 2) = Titulo
                     Else ' CAJA IGUAL ANTERIOR
                             If Titulo <> PONERTITULOHER Then 'SI EL TITULO ES DISTINTO
                                 R = R + 1
                                 Titulo = PONERTITULOHER
                                 xlSheet.Cells.Item(R, 2) = Titulo
                             End If
                     End If
                     R = R + 1
                     
                     
                     If Not IsNull(rsbasa!Descripcion) Then
                         Descripcion = Trim(rsbasa!Descripcion)
                     End If
                     If Not IsNull(rsbasa!FECHA_DESDE) Then
                         Descripcion = Descripcion & " Fecha Desde:" & rsbasa!FECHA_DESDE
                     End If
                     If Not IsNull(rsbasa!FECHA_HASTA) Then
                         Descripcion = Descripcion & " Fecha Hasta:" & rsbasa!FECHA_HASTA
                     End If
                     
                     If Not IsNull(rsbasa!NRO_DESDE) Then
                         Descripcion = Descripcion & " Nro. desde:" & rsbasa!NRO_DESDE
                     End If
                     If Not IsNull(rsbasa!NRO_HASTA) Then
                        Descripcion = Descripcion & " Nro. hasta:" & rsbasa!NRO_HASTA
                     End If
                     
                     
                     
                     If Not IsNull(rsbasa!LETRA_DESDE) Then
                        Descripcion = Descripcion & "Letra Desde:" & rsbasa!LETRA_DESDE
                     End If
                     If Not IsNull(rsbasa!LETRA_HASTA) Then
                        Descripcion = Descripcion & "Letra Hasta:" & rsbasa!LETRA_HASTA
                     End If
                     
                     If Not IsNull(rsbasa!APELLIDO_NOMBRE) Then
                        Descripcion = Descripcion & "Nombre:" & rsbasa!APELLIDO_NOMBRE
                     End If
                     If Not IsNull(rsbasa!EXPEDIENTE) Then
                        Descripcion = Descripcion & "Exp.:" & rsbasa!EXPEDIENTE
                     End If
                    
                    If ID_referencia = True Then
                         xlSheet.Cells.Item(R, 4) = "ID: " & rsbasa!COD_ID_REFERENCIA
                    End If
                    
                    
                        Rem  xlSheet.Cells.Item(R, 3).WrapText = True
                         xlSheet.Cells.Item(R, 3) = Trim(Descripcion)
                    
                    rsbasa.MoveNext
                    
                 Loop
                 '----------------------------------- FIN REFERENCIA ---------------------
    rsbasa.Close
    
    xlBook.Save
    xlBook.Close
    xlApp.Workbooks.Close
    Set xlApp = Nothing
    
    MousePointer = 0
    MsgBox "La exportación de datos a terminado", vbInformation
    Exit Sub
er:
    MousePointer = 0
    MsgBox Err.Description

End Sub
Public Sub ExportarExcelporCaja2(Filtro As Boolean)
   Dim Sql As String
   Dim DATO As String
   Dim rsbasa As New ADODB.Recordset
   Dim oConn As ADODB.Connection
   Dim oRS As ADODB.Recordset
   Dim i As Integer
   Dim TITULOHERANT As String
   Dim PONERTITULOHER As String
   On Error GoTo er
   MousePointer = 11
      Set oConn = New ADODB.Connection
      If Dir("C:\Referencia por caja.xls") <> "" Then
           Kill "C:\Referencia por caja.xls"
       End If
       FileCopy "strPasoPlanillasReferencia por caja.xls", "C:\Referencia por caja.xls"
       Set oConn = New ADODB.Connection
       oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=C:\Referencia por caja.xls;" & _
                   "Extended Properties=""Excel 8.0;HDR=NO;"""
              
    ' ------------------------INICIO REFERENCIA ---------------------------------------
    Set oRS = New ADODB.Recordset
    oRS.Open "Select * from Cajas", oConn, adOpenKeyset, adLockOptimistic
    Sql = " SELECT COD_ID_REFERENCIA, INDICES.TITULOHERENCIA,INDICES.DESCRIPCION AS DESCRIPCIONINDICE , REFERENCIAS.NRO_CAJA, REFERENCIAS.ITEM,"
    Sql = Sql & vbCrLf & " REFERENCIAS.INDICE, REFERENCIAS.DESCRIPCION,"
    Sql = Sql & vbCrLf & " REFERENCIAS.COD_CLIENTE,INDICES.ID_CODIGO_DOCUMENTO,"
    Sql = Sql & vbCrLf & " REFERENCIAS.FECHA_DESDE, REFERENCIAS.FECHA_HASTA,"
    Sql = Sql & vbCrLf & " REFERENCIAS.NRO_DESDE, REFERENCIAS.NRO_HASTA,"
    Sql = Sql & vbCrLf & " REFERENCIAS.LETRA_DESDE, REFERENCIAS.LETRA_HASTA,"
    Sql = Sql & vbCrLf & " REFERENCIAS.EXPEDIENTE, REFERENCIAS.APELLIDO_NOMBRE, REFERENCIAS.BORRADO"
    Sql = Sql & vbCrLf & " From REFERENCIAS, INDICES"
    Sql = Sql & vbCrLf & " WHERE REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND"
    Sql = Sql & vbCrLf & " REFERENCIAS.INDICE = INDICES.INDICE "
    If Filtro = True And Filtro_Reporte <> "" Then
        Sql = Sql & vbCrLf & "  " & Filtro_Reporte
    Else
        Sql = Sql & vbCrLf & " AND REFERENCIAS.COD_CLIENTE =" & ctlCliente.Valor
    End If
    Sql = Sql & vbCrLf & " ORDER BY REFERENCIAS.NRO_CAJA , REFERENCIAS.INDICE "
    Set rsbasa = New ADODB.Recordset
    rsbasa.Open Sql, ConActiva, 0, 1
    TITULOHERANT = "NO"
    Dim Descripcion As String
    Dim Caja As Long
    Dim Titulo As String
        oRS.MoveNext
    Do While Not rsbasa.EOF
               oRS.MoveNext
        If IsNull(rsbasa!TituloHerencia) Then
            PONERTITULOHER = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & "- - - - - - - - "
        Else
            PONERTITULOHER = "DOC:" & rsbasa!ID_CODIGO_DOCUMENTO & "//" & Trim(rsbasa!TituloHerencia)
        End If
               
        If Caja = 0 Then
                Caja = rsbasa!NRO_CAJA
                oRS.Fields(0).value = rsbasa!NRO_CAJA
                Titulo = PONERTITULOHER
                oRS.Fields(1).value = Titulo
                oRS.Update
                oRS.MoveNext
        End If
        If Caja <> rsbasa!NRO_CAJA Then
                Caja = rsbasa!NRO_CAJA
                oRS.Fields(0).value = rsbasa!NRO_CAJA
                Titulo = PONERTITULOHER
                oRS.Fields(1).value = Titulo
                oRS.Update
                oRS.MoveNext
        Else ' CAJA IGUAL ANTERIOR
                If Titulo <> PONERTITULOHER Then 'SI EL TITULO ES DISTINTO
                    Titulo = PONERTITULOHER
                    oRS.Fields(1).value = Titulo
                    oRS.Update
                    oRS.MoveNext
                End If
        End If
        
        oRS.Fields(1).value = rsbasa!COD_ID_REFERENCIA
        If Not IsNull(rsbasa!Descripcion) Then
            Descripcion = Trim(rsbasa!Descripcion)
        End If
        If Not IsNull(rsbasa!FECHA_DESDE) Then
            Descripcion = Descripcion & " Fecha Desde:" & rsbasa!FECHA_DESDE
        End If
        If Not IsNull(rsbasa!FECHA_HASTA) Then
            Descripcion = Descripcion & " Fecha Hasta:" & rsbasa!FECHA_HASTA
        End If
        
        If Not IsNull(rsbasa!NRO_DESDE) Then
            Descripcion = Descripcion & " Nro. desde:" & rsbasa!NRO_DESDE
        End If
        If Not IsNull(rsbasa!NRO_HASTA) Then
           Descripcion = Descripcion & " Nro. hasta:" & rsbasa!NRO_HASTA
        End If
        
        
        
        If Not IsNull(rsbasa!LETRA_DESDE) Then
           Descripcion = Descripcion & "Letra Desde:" & rsbasa!LETRA_DESDE
        End If
        If Not IsNull(rsbasa!LETRA_HASTA) Then
           Descripcion = Descripcion & "Letra Hasta:" & rsbasa!LETRA_HASTA
        End If
        
        If Not IsNull(rsbasa!APELLIDO_NOMBRE) Then
           Descripcion = Descripcion & "Nombre:" & rsbasa!APELLIDO_NOMBRE
        End If
        If Not IsNull(rsbasa!EXPEDIENTE) Then
           Descripcion = Descripcion & "Exp.:" & rsbasa!EXPEDIENTE
        End If
        
       If Len(Descripcion) > 250 Then
            oRS.Fields(2).value = Mid(Descripcion, 1, 250)
        Else
            oRS.Fields(2).value = Mid(Descripcion, 1, 250)
            oRS.Fields(3).value = Mid(Descripcion, 251, 500)
        End If
        oRS.Update
        rsbasa.MoveNext
    Loop
    '----------------------------------- FIN REFERENCIA ---------------------
    rsbasa.Close
    oConn.Close
    MousePointer = 0
    Exit Sub
er:
    MousePointer = 0
    MsgBox Err.Description

End Sub



Public Sub BuscarIndice(DATO As String, Optional EXPANDER As Boolean)
'    Dim I  As Integer
'    Dim A As Integer
'    Dim B As Integer
'    Dim Indice As String
'
'
'        For I = 1 To cltIndice1.Nodes.Count
'            trvIndices.Nodes.Item(I).BackColor = &H80000005
'            If dato = "" Or dato = " " Then
'            Else
'                B = InStr(UCase(trvIndices.Nodes.Item(I).Text), "-")
'                If UCase(trvIndices.Nodes.Item(I).Text) <> "TODAS LAS CATEGORIAS" Then
'                   If Mid(dato, 1, 1) = "0" Then
'                       ' BUSCAR INDICE
'                       Indice = Mid(trvIndices.Nodes.Item(I).Text, 1, B - 2)
'                       If Indice = UCase(dato) Then
'                            A = 1
'                       Else
'                            A = 0
'                       End If
'                   Else
'                        ' BUSCAR NOMBRE
'                        A = InStr(UCase(trvIndices.Nodes.Item(I).Text), UCase(dato))
'                    End If
'                    If A = 0 Then
'                      If EXPANDER = True Then
'                        trvIndices.Nodes.Item(I).Expanded = False
'                      End If
'                    Else
'                        trvIndices.Nodes.Item(I).Expanded = True
'                        trvIndices.Nodes.Item(I).Selected = True
'                        trvIndices.Nodes.Item(I).BackColor = &HFFFF00
'                    End If
'                End If
'            End If
'        Next
End Sub

Public Sub ResizePic()
' altura
Exit Sub
picBuscar.Top = cltIndice1.Top
picCargar.Top = cltIndice1.Top
Dim AlturaDisponible As Integer
    AlturaDisponible = MDIfrmInicio.Height - cltIndice1.Top - StatusBar1.Height - 900
    picBuscar.Height = AlturaDisponible
    picCargar.Height = AlturaDisponible
    cltIndice1.Height = AlturaDisponible
    ' Grillas
    grdModificacion.Height = picCargar.Height - grdModificacion.Top - 100
    grdReferencias.Height = picBuscar.Height - grdReferencias.Top - 100

Dim MargenAncho As Integer
   Rem MargenAncho = MDIfrmInicio.Width - picCargar.Width - trvIndices.Width
    Rem picBuscar.left = MargenAncho
    Rem picCargar.left = MargenAncho
    Rem trvIndices.Width = picCargar.left - 100
    
    MargenAncho = MDIfrmInicio.Width - picCargar.Width - cltIndice1.Width
    picBuscar.Left = cltIndice1.Width + 100
    picCargar.Left = cltIndice1.Width + 100
   
    
End Sub

Public Sub ImprimirControlCarga()
    Dim Sql As String
    MousePointer = 11
    Sql = "  SELECT * "
    Sql = Sql & " From REFERENCIASLARGO"
    Sql = Sql & "  Where " & FechaServerTipo("REFERENCIASLARGO. FECHA_MODIFICACION") & " ='" & Format(Now, "DD/MM/YYYY")
    Sql = Sql & "' AND USUARIO_MODIFICACION  = '" & ctlPersonal.Valor & "'"
    If IsNull(ctlPersonal.Valor) Then
        MsgBox "Ingrese el Personal", vbCritical
        MousePointer = 0
        Exit Sub
    End If
    frmReportes.ImprimirReporte PasoReportes + "rptControlReferencias.rpt", Sql, True
    MousePointer = 0
End Sub
Public Sub ImprimirControlReferencia()
        MousePointer = 11
        Dim Sql As String
        
        
        If IsNull(ctlPersonal.Valor) Then
            MsgBox "Ingrese el Personal", vbCritical
            MousePointer = 0
            Exit Sub
        End If
'          sql = sql & "  TO_DATE('" & txtFechaInicio.Text & "', 'DD/MM/YYYY HH24:MI:SS')"
'        sql = sql & "  AND TO_DATE('" & txtFechaFin.Text & "','DD/MM/YYYY HH24:MI:SS')"
'
'        sql = "   SELECT *"
'        sql = sql & "  From V_CONTROL_REFERENCIAS_DATOS"
'        sql = sql & "  WHERE  "
'        sql = sql & "  USUARIO_MODIFICACION = '" & ctlPersonal.Valor & "'"
'        sql = sql & "  AND   FECHA_MODIFICACION >  "
'        sql = sql & "  '" & Format(txtFechaInicio.Text, "DD/MM/YYYY") & "'"
'
'        frmReportes.ImprimirReporte PasoReportes + "rptControlReferenciasDatos.rpt", sql, True
'        MousePointer = 0
End Sub

Public Sub ImprimirControlCajasSinReferencias()
        MousePointer = 11
        
        Dim rs As New ADODB.Recordset
        Dim FiltroCaja As String
        Dim Sql As String
        
'        SELECT     CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, REFERENCIAS.NRO_CAJA AS EXPR1
'FROM         CONTENEDOR LEFT OUTER JOIN
'                      REFERENCIAS ON CONTENEDOR.NRO_CAJA = REFERENCIAS.NRO_CAJA AND CONTENEDOR.COD_CLIENTE = REFERENCIAS.COD_CLIENTE
'WHERE     (CONTENEDOR.ESTADO IN (2, 3)) AND (CONTENEDOR.COD_CLIENTE = 4) AND (REFERENCIAS.NRO_CAJA IS NULL)
'ORDER BY CONTENEDOR.NRO_CAJA
        
'        Sql = " SELECT CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA"
'        Sql = Sql & " From CONTENEDOR, REFERENCIAS"
'        Sql = Sql & "  WHERE CONTENEDOR.NRO_CAJA = REFERENCIAS.NRO_CAJA (+) AND"
'        Sql = Sql & " CONTENEDOR.COD_CLIENTE = REFERENCIAS.COD_CLIENTE (+)"
'        Sql = Sql & " AND CONTENEDOR.ESTADO IN (2, 3) "
'        Sql = Sql & " AND CONTENEDOR.COD_CLIENTE = " & ctlCliente.Valor
'        Sql = Sql & " AND REFERENCIAS.NRO_CAJA IS NULL "
'        Sql = Sql & "  ORDER BY CONTENEDOR.NRO_CAJA"
        
        Sql = " SELECT     CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA"
Sql = Sql & vbCrLf & " FROM         REFERENCIAS RIGHT OUTER JOIN"
  Sql = Sql & "                    CONTENEDOR ON REFERENCIAS.NRO_CAJA = CONTENEDOR.NRO_CAJA AND REFERENCIAS.COD_CLIENTE = CONTENEDOR.COD_CLIENTE"

Sql = Sql & " WHERE     "
Rem Sql = Sql & ""(CONTENEDOR.ESTADO IN (2, 3))AND "
Sql = Sql & "  CONTENEDOR.COD_CLIENTE = " & ctlCliente.Valor
Sql = Sql & " AND (REFERENCIAS.NRO_CAJA IS NULL)"




Sql = " SELECT  CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA "
Sql = Sql & vbCrLf & " FROM         REFERENCIAS RIGHT OUTER JOIN"
Sql = Sql & vbCrLf & " CONTENEDOR ON REFERENCIAS.NRO_CAJA = CONTENEDOR.NRO_CAJA AND REFERENCIAS.COD_CLIENTE = CONTENEDOR.COD_CLIENTE"
Sql = Sql & vbCrLf & " Where CONTENEDOR.COD_CLIENTE = " & ctlCliente.Valor
Sql = Sql & vbCrLf & " And REFERENCIAS.NRO_CAJA Is Null "


        
        FiltroCaja = ""
        
        rs.Open Sql, ConActiva, 0, 1
        Do While Not rs.EOF
            FiltroCaja = FiltroCaja & rs!NRO_CAJA & ","
            rs.MoveNext
        Loop
        
FiltroCaja = Mid(FiltroCaja, 1, Len(FiltroCaja) - 1)
        Set rs = New ADODB.Recordset
        rs.CursorLocation = adUseClient
        
        Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO AS R_INTERNO, REMITOS_CUERPO.NRO_REM_PROV AS REMITO_CLIENTE, REMITOS_CUERPO.FECHA,"
        Sql = Sql & vbCrLf & " REMITOS_CUERPO.ID_CLIENTE, REMITOS_DETALLE.DESDE AS NRO_CAJA, CLIENTEUSUARIO.APELLIDO_NOMBRE"
        Sql = Sql & vbCrLf & "  FROM         REMITOS_CUERPO INNER JOIN"
        Sql = Sql & vbCrLf & "  REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & "  CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
        Sql = Sql & vbCrLf & " WHERE  REMITOS_CUERPO.ID_CLIENTE = " & ctlCliente.Valor
        Sql = Sql & vbCrLf & " AND (REMITOS_CUERPO.TIPO = 0)"
        Sql = Sql & vbCrLf & " AND (REMITOS_DETALLE.DESDE IN (" & FiltroCaja & "))"
        Sql = Sql & vbCrLf & " ORDER BY REMITOS_CUERPO.NRO_REMITO DESC"
        
        
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        frmInforme.CargarInforme "Informe Cajas sin Referencias", rs
        frmInforme.Show
        MousePointer = 0
        
End Sub

Private Sub txtReferenciaLote_Change()
  Rem  txtPasoImagenes.Text = "\\Serverbackup_1\E\Usuarios\Basa\Operaciones\Referencias\" & txtReferenciaLote.Text
End Sub

Public Sub ActualizarGrillaCarga()
Dim sSQL As String
Dim rsReferencias As New ADODB.Recordset
rsReferencias.CursorLocation = adUseClient
        sSQL = "  SELECT NRO_CAJA AS CAJA, INDICE, DESCRIPCION, FECHA_DESDE AS DESDE,"
        sSQL = sSQL & vbCrLf & " FECHA_HASTA AS HASTA, NRO_DESDE, NRO_HASTA, LETRA_DESDE,"
        sSQL = sSQL & vbCrLf & " LETRA_HASTA , COD_ID_REFERENCIA"
        sSQL = sSQL & vbCrLf & " From REFERENCIAS"
        sSQL = sSQL & vbCrLf & " Where COD_CLIENTE = " & ctlCliente.Valor
        sSQL = sSQL & vbCrLf & " And NRO_CAJA = " & txtNro_Caja.Text
        sSQL = sSQL & vbCrLf & " ORDER BY COD_ID_REFERENCIA DESC"
        rsReferencias.Open sSQL, ConActiva, 0, 1
        
       
       Set grdModificacion.DataSource = rsReferencias.DataSource
       grdModificacion.DataMember = rsReferencias.DataMember
       grdModificacion.Columns.Item(0).Width = 800
       grdModificacion.Columns.Item(1).Width = 1000
       grdModificacion.Columns.Item(2).Width = 4000
       grdModificacion.Columns.Item(3).Width = 1000
       grdModificacion.Columns.Item(4).Width = 1000
       grdModificacion.Columns.Item(5).Width = 1000
       grdModificacion.Columns.Item(6).Width = 1000
       grdModificacion.Columns.Item(7).Width = 1000
       grdModificacion.Columns.Item(8).Width = 1000
       grdModificacion.Columns.Item(9).Width = 1000
End Sub

Public Sub Buscar_Inidice_Por_caja()
    Dim rsIndice As New ADODB.Recordset
  
  Dim sSQL As String
  Dim Item As Integer
        If txtNro_Caja = "" Then
           cltIndice1.Actualizar ctlCliente.Valor, Nulo, 0
            Exit Sub
        End If
        
        sSQL = " SELECT CLIENTEUSUARIO.COD_INDICE"
        sSQL = sSQL & vbCrLf & " FROM REMITOS_CUERPO, REMITOS_DETALLE, CLIENTEUSUARIO"
        sSQL = sSQL & vbCrLf & " Where REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO  AND"
        sSQL = sSQL & vbCrLf & " REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
        sSQL = sSQL & vbCrLf & " AND REMITOS_CUERPO.ID_CLIENTE = " & ctlCliente.Valor
        sSQL = sSQL & vbCrLf & " AND REMITOS_DETALLE.DESDE =" & txtNro_Caja.Text
        sSQL = sSQL & vbCrLf & " AND REMITOS_CUERPO.TIPO = 0"

        rsIndice.Open sSQL, ConActiva, 0, 1
        If Not rsIndice.EOF Then
             cltIndice1.Actualizar ctlCliente.Valor, Nulo, 0, rsIndice!Cod_Indice
         Else
                MsgBox "No se encontro el remito"
                cltIndice1.Actualizar ctlCliente.Valor, Nulo, 0
        End If
End Sub

Public Sub BorrarCaja(Cliente As Integer, CAJAS As String)

Dim Sql As String

    Sql = " INSERT INTO REFERENCIAS_HISTORICOS"
Sql = Sql & " SELECT     REFERENCIAS.*"
Sql = Sql & " From REFERENCIAS"
Sql = Sql & " WHERE COD_CLIENTE = " & Cliente
Sql = Sql & " AND NRO_CAJA IN (" & CAJAS & "  )"



End Sub

Public Sub SUMAR_DESCRIPCION_DOCUMENTO(IDREFERENCIA As Long)
 Dim rs As New ADODB.Recordset
 Dim Sql As String
 Dim DESCR As String
 
 Sql = " SELECT     REFERENCIAS.COD_ID_REFERENCIA, REFERENCIAS.DESCRIPCION, INDICES.DESCRIPCION AS DESCRIPCION_INDICE"
Sql = Sql & " FROM         REFERENCIAS INNER JOIN"
Sql = Sql & " INDICES ON REFERENCIAS.INDICE = INDICES.INDICE AND REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE"
Sql = Sql & "  Where REFERENCIAS.COD_ID_REFERENCIA = " & IDREFERENCIA

If Not rs.EOF Then
    DESCR = Replace(Trim(rs!Descripcion), vbCrLf, "") & " " & Trim(rs!DESCRIPCION_INDICE)
End If

 
ExecutarSql " UPDATE    REFERENCIAS SET DESCRIPCION ='" & DESCR & "'Where COD_ID_REFERENCIA = " & IDREFERENCIA

End Sub
