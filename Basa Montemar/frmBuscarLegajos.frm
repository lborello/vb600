VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmBuscarLegajos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BUSCAR LEGAJOS"
   ClientHeight    =   9750
   ClientLeft      =   585
   ClientTop       =   1035
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   13845
   Begin TabDlg.SSTab sstLegajos 
      Height          =   9555
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   16854
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483626
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Buscar Legajos Montemar"
      TabPicture(0)   =   "frmBuscarLegajos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ctlIndiceLegajo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Rearchivo"
      TabPicture(1)   =   "frmBuscarLegajos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraOrdenLegajos"
      Tab(1).Control(1)=   "grdRearchivo"
      Tab(1).Control(2)=   "cmdCopiarExcel"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Rearchivo Digital"
      TabPicture(2)   =   "frmBuscarLegajos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sstRearchivoDigital"
      Tab(2).Control(1)=   "txtDescripcion"
      Tab(2).Control(2)=   "cmdBuscarRearchivoDigital"
      Tab(2).Control(3)=   "txtNroLegajo"
      Tab(2).Control(4)=   "Label9"
      Tab(2).Control(5)=   "Label7"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmBuscarLegajos.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "ctlClienteBuscarLegajo"
      Tab(3).Control(1)=   "cmdAceptarBuscarlegajo"
      Tab(3).Control(2)=   "txtLegajoBuscar"
      Tab(3).Control(3)=   "grdCargarBuscarLegajos"
      Tab(3).Control(4)=   "DataGrid1"
      Tab(3).Control(5)=   "Label12"
      Tab(3).Control(6)=   "Label11"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Asignacion de carga "
      TabPicture(4)   =   "frmBuscarLegajos.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "MSFlexGrid1"
      Tab(4).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6375
         Left            =   -74100
         TabIndex        =   69
         Top             =   2160
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   11245
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Height          =   1515
         Left            =   180
         TabIndex        =   23
         Top             =   1020
         Width           =   2715
         Begin VB.CheckBox chkReferencias 
            Caption         =   "Referencias"
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
            Left            =   180
            TabIndex        =   45
            Top             =   180
            Width           =   1335
         End
         Begin VB.CheckBox chkLegajos 
            Caption         =   "Legajos"
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
            Left            =   180
            TabIndex        =   44
            Top             =   480
            Width           =   1095
         End
         Begin VB.CheckBox ChkRearchivoDigital 
            Caption         =   "Rearchivo Digital"
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
            Left            =   180
            TabIndex        =   25
            Top             =   1080
            Width           =   1995
         End
         Begin VB.CheckBox chkRearchivoLote 
            Caption         =   "Rearchivo Físico"
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
            Left            =   180
            TabIndex        =   24
            Top             =   780
            Width           =   1875
         End
      End
      Begin Controles.cltGenerico ctlClienteBuscarLegajo 
         Height          =   375
         Left            =   -73440
         TabIndex        =   43
         Top             =   1020
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   661
      End
      Begin VB.CommandButton cmdAceptarBuscarlegajo 
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
         Left            =   -68280
         TabIndex        =   42
         Top             =   6000
         Width           =   1215
      End
      Begin VB.TextBox txtLegajoBuscar 
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
         Left            =   -73620
         TabIndex        =   40
         Top             =   1500
         Width           =   2055
      End
      Begin MSFlexGridLib.MSFlexGrid grdCargarBuscarLegajos 
         Height          =   3915
         Left            =   -74880
         TabIndex        =   39
         Top             =   1980
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   6906
         _Version        =   393216
         Cols            =   3
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   38
         Top             =   6540
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4683
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
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
      Begin VB.Frame Frame3 
         Height          =   1515
         Left            =   3000
         TabIndex        =   26
         Top             =   1020
         Width           =   9615
         Begin VB.CheckBox chkSolicitarAño 
            Caption         =   "Solicitar Año"
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
            Left            =   7560
            TabIndex        =   46
            Top             =   1020
            Width           =   1815
         End
         Begin VB.TextBox txtLecturaLegajos 
            BackColor       =   &H00FFC0C0&
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
            Left            =   1620
            TabIndex        =   30
            Top             =   900
            Width           =   1215
         End
         Begin VB.CommandButton cmdBuscar 
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
            Height          =   450
            Left            =   6720
            TabIndex        =   29
            Top             =   840
            Width           =   405
         End
         Begin VB.TextBox txtBuscarLegajo 
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
            Left            =   3600
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   900
            Width           =   3015
         End
         Begin VB.ComboBox CboCampo 
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
            ItemData        =   "frmBuscarLegajos.frx":008C
            Left            =   5460
            List            =   "frmBuscarLegajos.frx":009F
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   300
            Width           =   3795
         End
         Begin Controles.cltGenerico ctlCliente 
            Height          =   360
            Left            =   780
            TabIndex        =   31
            Top             =   300
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   635
         End
         Begin VB.Label lblIndice 
            Caption         =   "Label10"
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
            Left            =   8160
            TabIndex        =   70
            Top             =   1140
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Legajo:"
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
            Left            =   2940
            TabIndex        =   35
            Top             =   960
            Width           =   675
         End
         Begin VB.Label Label5 
            Caption         =   "Codigo de barra:"
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
            TabIndex        =   34
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Criterio:"
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
            Left            =   4740
            TabIndex        =   33
            Top             =   360
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
            TabIndex        =   32
            Top             =   360
            Width           =   555
         End
      End
      Begin TabDlg.SSTab sstRearchivoDigital 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   20
         Top             =   1620
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   9975
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Resultado de Busqueda"
         TabPicture(0)   =   "frmBuscarLegajos.frx":00F8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdRearchivoDigital"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Ver Imagen"
         TabPicture(1)   =   "frmBuscarLegajos.frx":0114
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ctlVerImagenes1"
         Tab(1).ControlCount=   1
         Begin Controles.ctlVerImagenes ctlVerImagenes1 
            Height          =   4815
            Left            =   -74700
            TabIndex        =   66
            Top             =   480
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   8493
         End
         Begin MSDataGridLib.DataGrid grdRearchivoDigital 
            Height          =   4815
            Left            =   60
            TabIndex        =   21
            Top             =   600
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   8493
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   18
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
         Left            =   -70560
         TabIndex        =   18
         Top             =   1080
         Width           =   3615
      End
      Begin VB.CommandButton cmdBuscarRearchivoDigital 
         Caption         =   "Buscar"
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
         Left            =   -66840
         TabIndex        =   17
         Top             =   1080
         Width           =   1275
      End
      Begin VB.TextBox txtNroLegajo 
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
         Left            =   -73620
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton cmdCopiarExcel 
         Caption         =   "Copiar Datos Excel"
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
         Left            =   -67740
         TabIndex        =   14
         Top             =   8400
         Width           =   2760
      End
      Begin MSDataGridLib.DataGrid grdRearchivo 
         Height          =   3615
         Left            =   -73920
         TabIndex        =   13
         Top             =   3840
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   6376
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   17
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
      Begin VB.Frame Frame1 
         Height          =   6855
         Left            =   3000
         TabIndex        =   12
         Top             =   2580
         Width           =   9795
         Begin TabDlg.SSTab SSTab1 
            Height          =   4455
            Left            =   120
            TabIndex        =   47
            Top             =   2100
            Width           =   9555
            _ExtentX        =   16854
            _ExtentY        =   7858
            _Version        =   393216
            Tabs            =   2
            TabHeight       =   520
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
            TabPicture(0)   =   "frmBuscarLegajos.frx":0130
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grdSeleccionLegajos"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdLecturaLegajo"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "cmdEntrada"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "cmdBorrar"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "cmdBuscarRearchivo"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "cmdRotulos"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "cmdBuscarLegajos"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "cmdLimpiarLegajos"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "cmdPasarTodosLegajos"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).ControlCount=   9
            TabCaption(1)   =   "Varios"
            TabPicture(1)   =   "frmBuscarLegajos.frx":014C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "cmdLimpiar"
            Tab(1).Control(1)=   "cmdReporteBusqueda"
            Tab(1).Control(2)=   "cmdInsertarBusqueda"
            Tab(1).Control(3)=   "grdVarios"
            Tab(1).ControlCount=   4
            Begin VB.CommandButton cmdLimpiar 
               Caption         =   "Limpiar"
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
               Left            =   -71520
               TabIndex        =   63
               Top             =   3540
               Width           =   1575
            End
            Begin VB.CommandButton cmdReporteBusqueda 
               Caption         =   "Reporte"
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
               Left            =   -73200
               TabIndex        =   59
               Top             =   3540
               Width           =   1575
            End
            Begin VB.CommandButton cmdInsertarBusqueda 
               Caption         =   "Busqueda"
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
               Left            =   -74820
               TabIndex        =   58
               Top             =   3540
               Width           =   1515
            End
            Begin VB.CommandButton cmdPasarTodosLegajos 
               Caption         =   "Pasar"
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
               Left            =   1500
               TabIndex        =   55
               Top             =   4020
               Width           =   1215
            End
            Begin VB.CommandButton cmdLimpiarLegajos 
               Caption         =   "Limpiar"
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
               Left            =   2760
               TabIndex        =   54
               Top             =   4020
               Width           =   1215
            End
            Begin VB.CommandButton cmdBuscarLegajos 
               Caption         =   "Buscar"
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
               Left            =   4020
               TabIndex        =   53
               Top             =   4020
               Width           =   1215
            End
            Begin VB.CommandButton cmdRotulos 
               Caption         =   "Rotulos"
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
               Left            =   240
               TabIndex        =   52
               Top             =   4020
               Width           =   1215
            End
            Begin VB.CommandButton cmdBuscarRearchivo 
               BackColor       =   &H80000010&
               Caption         =   "Buscar Rear."
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
               Left            =   4020
               MaskColor       =   &H00808080&
               TabIndex        =   51
               Top             =   3540
               Width           =   1215
            End
            Begin VB.CommandButton cmdBorrar 
               Caption         =   "Borrar"
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
               Left            =   2760
               TabIndex        =   50
               Top             =   3540
               Width           =   1215
            End
            Begin VB.CommandButton cmdEntrada 
               BackColor       =   &H80000010&
               Caption         =   "Entrada"
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
               Left            =   1500
               MaskColor       =   &H00808080&
               TabIndex        =   49
               Top             =   3540
               Width           =   1215
            End
            Begin VB.CommandButton cmdLecturaLegajo 
               Caption         =   "Lectura "
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
               Left            =   240
               TabIndex        =   48
               Top             =   3540
               Width           =   1215
            End
            Begin MSFlexGridLib.MSFlexGrid grdSeleccionLegajos 
               Height          =   2955
               Left            =   240
               TabIndex        =   56
               Top             =   480
               Width           =   9075
               _ExtentX        =   16007
               _ExtentY        =   5212
               _Version        =   393216
               AllowUserResizing=   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin MSFlexGridLib.MSFlexGrid grdVarios 
               Height          =   2895
               Left            =   -74940
               TabIndex        =   57
               Top             =   480
               Width           =   9195
               _ExtentX        =   16219
               _ExtentY        =   5106
               _Version        =   393216
               Cols            =   8
               AllowUserResizing=   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin MSFlexGridLib.MSFlexGrid grdResultadoBusqueda 
            Height          =   1695
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   9435
            _ExtentX        =   16642
            _ExtentY        =   2990
            _Version        =   393216
            Cols            =   8
            BackColorSel    =   16744576
            SelectionMode   =   1
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
      End
      Begin VB.Frame fraOrdenLegajos 
         Height          =   2415
         Left            =   -73920
         TabIndex        =   1
         Top             =   1080
         Width           =   10575
         Begin VB.TextBox txtDigito 
            Height          =   375
            Left            =   3600
            TabIndex        =   71
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Informe por ordenes"
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
            Left            =   7080
            TabIndex        =   68
            Top             =   1080
            Width           =   2415
         End
         Begin VB.CommandButton cmdImformeCajaRearchivo 
            Caption         =   "Informe Caja Rearchivo"
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
            Left            =   4440
            TabIndex        =   67
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox txtCajaRearchivo 
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
            Left            =   1800
            TabIndex        =   64
            Top             =   720
            Width           =   1695
         End
         Begin Controles.cltGenerico ctlPersonal 
            Height          =   375
            Left            =   1800
            TabIndex        =   60
            Top             =   1500
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   661
         End
         Begin VB.CommandButton cmdReImprimirOrden 
            Caption         =   "Imprimir Orden"
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
            Left            =   7080
            TabIndex        =   9
            Top             =   240
            Width           =   2400
         End
         Begin VB.CommandButton cmdOrdenCompleto 
            Caption         =   "Orden Completo"
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
            Left            =   4440
            TabIndex        =   8
            Top             =   240
            Width           =   2400
         End
         Begin VB.TextBox txtOrden 
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
            Left            =   1800
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
         Begin VB.CommandButton cmdOrdenLegajos 
            Caption         =   " Generar Orden"
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
            Left            =   4440
            TabIndex        =   6
            Top             =   720
            Width           =   2400
         End
         Begin VB.TextBox txtFechaOrden 
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
            Left            =   1800
            TabIndex        =   5
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton cmdOrdenesPendientes 
            Caption         =   "Ordenes Pendientes de ordenamiento"
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
            Left            =   1620
            TabIndex        =   4
            Top             =   1920
            Width           =   3720
         End
         Begin VB.CommandButton cmdOrdenesPorcaja 
            Caption         =   "Ordenes Pendientes de control"
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
            Left            =   5460
            TabIndex        =   3
            Top             =   1920
            Width           =   3240
         End
         Begin VB.CommandButton cmdControlOrden 
            Caption         =   "Control Orden"
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
            Left            =   7080
            TabIndex        =   2
            Top             =   660
            Width           =   2400
         End
         Begin VB.Label Label13 
            Caption         =   "Caja Rearch:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   780
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Personal"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   1560
            Width           =   1155
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha de Orden:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   61
            Top             =   -60
            Width           =   1755
         End
         Begin VB.Label Label2 
            Caption         =   "Nº de Orden:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   420
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha de Orden:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1140
            Width           =   1395
         End
      End
      Begin Controles.cltIndice ctlIndiceLegajo 
         Height          =   6795
         Left            =   120
         TabIndex        =   36
         Top             =   2640
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   11986
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
      Begin VB.Label Label12 
         Caption         =   "Legajo:"
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
         Left            =   -74820
         TabIndex        =   41
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Cliente legajo:"
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
         Left            =   -74820
         TabIndex        =   37
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Descripcion"
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
         Left            =   -71640
         TabIndex        =   19
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Nº de Legajo"
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
         Left            =   -74760
         TabIndex        =   15
         Top             =   1140
         Width           =   1095
      End
   End
   Begin VB.Menu mnuImagen 
      Caption         =   "Imgen"
      Visible         =   0   'False
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuVerPDF 
         Caption         =   "Ver PDF"
      End
      Begin VB.Menu mnuVerImagen 
         Caption         =   "Ver Imagen"
      End
   End
   Begin VB.Menu mnuArbol 
      Caption         =   "Arbol"
      Visible         =   0   'False
      Begin VB.Menu mnuBuscarLegajo 
         Caption         =   "Buscar Legajo"
      End
   End
   Begin VB.Menu mnugrdBuscar 
      Caption         =   "grd Buscar"
      Begin VB.Menu mnugrdVerImagen 
         Caption         =   "Ver imagen"
      End
      Begin VB.Menu mnuCopiarPaso 
         Caption         =   "Copiar Paso"
      End
   End
End
Attribute VB_Name = "frmBuscarLegajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim rslegajos As ADODB.Recordset


Public Sub BuscarLegajosRearchivo2(COD_CLIENTE As Integer, ID_CLIENTE_LEGAJO As Long)

End Sub


Private Sub cboBuscar_Change()
   
End Sub

Private Sub cmdBorrar_Click()
    Dim SqlIns As String
    Dim SqlUpd As String
    Dim SqlFiltro As String
    
If MsgBox("Esta usded segura de Borrar los legajos", vbYesNo) = vbYes Then
        ConBasa.BeginTrans
    On Error GoTo ErrorCap
        SqlIns = " INSERT INTO HISTORICO_LEGAJOS "
        SqlIns = SqlIns & vbCrLf & " (ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE,"
        SqlIns = SqlIns & vbCrLf & " CLIENTE_LEGAJO, DESCRIPCION, NRO_CAJA, COD_CLIENTE,"
        SqlIns = SqlIns & vbCrLf & " COD_UBICACION, COD_ESTADO, NOMBRE, COD_REMITO,"
        SqlIns = SqlIns & vbCrLf & " FECHA, ID_PERSONAL, FECHA_ACTUALIZACION,"
        SqlIns = SqlIns & vbCrLf & " ID_LEGAJO_ECOGAS, ID_CLIENTE_BASE, ERRORTIPEO,"
        SqlIns = SqlIns & vbCrLf & " CARGAPAGADA, NRO_REM_PROV, ORDEN,"
        SqlIns = SqlIns & vbCrLf & " NUMERO_LEGAJO_CLIENTE, FECHAPAGO, PEGADOETIQUETA,"
        SqlIns = SqlIns & vbCrLf & " CANTIDAD_CARACTERES,FECHA_HISTORICO)"
        SqlIns = SqlIns & vbCrLf & " SELECT ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE,"
        SqlIns = SqlIns & vbCrLf & " CLIENTE_LEGAJO, DESCRIPCION, NRO_CAJA, COD_CLIENTE,"
        SqlIns = SqlIns & vbCrLf & " COD_UBICACION, COD_ESTADO, NOMBRE, COD_REMITO,"
        SqlIns = SqlIns & vbCrLf & " FECHA, ID_PERSONAL, FECHA_ACTUALIZACION,"
        SqlIns = SqlIns & vbCrLf & " ID_LEGAJO_ECOGAS, ID_CLIENTE_BASE, ERRORTIPEO,"
        SqlIns = SqlIns & vbCrLf & " CARGAPAGADA, NRO_REM_PROV, ORDEN,"
        SqlIns = SqlIns & vbCrLf & " NUMERO_LEGAJO_CLIENTE, FECHAPAGO, PEGADOETIQUETA,"
        SqlIns = SqlIns & vbCrLf & " CANTIDAD_CARACTERES, " & SysDateMinutoSegundo
        SqlIns = SqlIns & vbCrLf & " FROM LEGAJOS "
        
        SqlUpd = " Update LEGAJOS "
        SqlUpd = SqlUpd & vbCrLf & " SET COD_INDICE = NULL, CLIENTE_LEGAJO = NULL,"
        SqlUpd = SqlUpd & vbCrLf & " DESCRIPCION = NULL, NRO_CAJA = NULL,"
        SqlUpd = SqlUpd & vbCrLf & " COD_ESTADO = NULL, NOMBRE = NULL,"
        SqlUpd = SqlUpd & vbCrLf & " COD_REMITO = NULL, FECHA = NULL, ID_PERSONAL = NULL,"
        SqlUpd = SqlUpd & vbCrLf & " FECHA_ACTUALIZACION = NULL, NRO_REM_PROV = NULL,"
        SqlUpd = SqlUpd & vbCrLf & " ORDEN = NULL, NUMERO_LEGAJO_CLIENTE = NULL,"
        SqlUpd = SqlUpd & vbCrLf & " ID_LEGAJO_ECOGAS = Null"
        
        
        Dim i As Integer
        For i = 1 To grdSeleccionLegajos.Rows - 1
            If i = 1 Then
                SqlFiltro = " WHERE ( COD_CLIENTE =" & grdSeleccionLegajos.TextMatrix(i, 1) & " AND ID_CLIENTE_LEGAJO =" & grdSeleccionLegajos.TextMatrix(i, 2) & ")"
            Else
                SqlFiltro = SqlFiltro & vbCrLf & " OR (COD_CLIENTE =" & grdSeleccionLegajos.TextMatrix(i, 1) & " AND ID_CLIENTE_LEGAJO =" & grdSeleccionLegajos.TextMatrix(i, 2) & ")"
            End If
        Next
        SqlIns = SqlIns & SqlFiltro
        SqlUpd = SqlUpd & SqlFiltro
        ExecutarSql SqlIns
        ExecutarSql SqlUpd
        ConBasa.CommitTrans
        MsgBox "La actualización fue exitosa", vbInformation
        TitulosSeleccionLegajos
        Else
        
        End If
        
        Exit Sub
ErrorCap:
         ConBasa.RollbackTrans
         MsgBox "La Actualizacion NO fue realizada", vbCritical
        
End Sub

Private Sub cmdBuscar_Click()


BuscarLegajosRearchivo ctlCliente.Valor, txtBuscarLegajo.Text


    
'On Error GoTo salir
'    Set rslegajos = New ADODB.Recordset
'    rslegajos.CursorLocation = adUseClient
'    Dim Sql As String
'    Dim Filtro As String
'    Dim detalle As String
'    Dim Año As String
'
'    If CboCampo.Text = "" Then
'        MsgBox "Ingrese el Campo ", vbInformation
'        Exit Sub
'    End If
'
'    If IsNull(ctlCliente.Valor) Then
'        MsgBox "Ingrese el Cliente ", vbInformation
'        Exit Sub
'    End If
'    txtBuscarLegajo = Replace(txtBuscarLegajo, vbCrLf, "")
'
'    If Mid(txtBuscarLegajo.Text, Len(txtBuscarLegajo.Text)) = "," Then
'        txtBuscarLegajo.Text = Mid(txtBuscarLegajo.Text, 1, Len(txtBuscarLegajo.Text) - 1)
'    End If
'
'    Sql = " SELECT  INDICES.DESCRIPCION, LEGAJOS.ID_CLIENTE_LEGAJO, LEGAJOS.CLIENTE_LEGAJO , DESCRIPCION_REMITO , LEGAJOS.NRO_CAJA, LEGAJOS.COD_ESTADO ,NOMBRE  "
'    Sql = Sql & vbCrLf & " FROM LEGAJOS LEFT OUTER JOIN"
'    Sql = Sql & vbCrLf & " INDICES ON LEGAJOS.COD_CLIENTE = INDICES.COD_CLIENTE AND LEGAJOS.COD_INDICE = INDICES.INDICE"
'    Sql = Sql & vbCrLf & " where LEGAJOS.COD_CLIENTE = " & ctlCliente.Valor & " And "
'    If lblIndice.Caption <> "" Then
'         Sql = Sql & vbCrLf & " COD_INDICE like '" & lblIndice.Caption & "%' AND "
'    End If
'
'    Select Case CboCampo.Text
'        Case "ID_CLIENTE_LEGAJO"
'            Filtro = " ID_CLIENTE_LEGAJO IN (" & txtBuscarLegajo & ")"
'        Case "CLIENTE_LEGAJO_LETRA"
'            Filtro = " CLIENTE_LEGAJO like '%" & txtBuscarLegajo & "%'"
'        Case "CLIENTE_LEGAJO_NUMERO"
'             Filtro = " NUMERO_LEGAJO_CLIENTE IN (" & txtBuscarLegajo & ")"
'        Case "NOMBRE"
'             Filtro = " NOMBRE like '%" & txtBuscarLegajo & "%'"
'        Case "DESCRIPCION"
'            Filtro = " DESCRIPCION like '%" & txtBuscarLegajo & "%'"
'    End Select
'
'   Rem TitulosBuscar
'    Sql = Sql & Filtro
'        rslegajos.Open Sql, ConActiva, 0, 1
'
'If (rslegajos.EOF) Then
' MsgBox "No exsite el legajo"
'End If
'
'    Do While Not rslegajos.EOF
'        grdResultadoBusqueda.AddItem grdResultadoBusqueda.Rows & vbTab & "Legajos" & vbTab & rslegajos!Descripcion & vbTab & rslegajos!ID_CLIENTE_LEGAJO & vbTab & Trim(rslegajos!CLIENTE_LEGAJO & "  " & Replace(rslegajos!DESCRIPCION_REMITO, Chr(9), "")) & vbTab & rslegajos!Cod_Estado & vbTab & rslegajos!NRO_CAJA & vbTab & rslegajos!Nombre
'
'        rslegajos.MoveNext
'    Loop
'
'    If CboCampo.Text = "ID_CLIENTE_LEGAJO" Then
'        Exit Sub
'    End If
'
'    ' Referencia
'
'If chkReferencias.value = 1 Then
'    If lblIndice.Caption <> "" Then
'
'
'
'        Sql = " SELECT COD_CLIENTE, NRO_CAJA, NRO_DESDE, NRO_HASTA,FECHA_DESDE, INDICE "
'        Sql = Sql & vbCrLf & " From REFERENCIAS "
'        Sql = Sql & vbCrLf & " WHERE  COD_CLIENTE= " & ctlCliente.Valor & " And "
'
'        If lblIndice.Caption <> "" Then
'             Sql = Sql & vbCrLf & " INDICE like '" & lblIndice.Caption & "%' AND "
'        End If
'        If IsNumeric(txtBuscarLegajo.Text) Then
'            Sql = Sql & vbCrLf & txtBuscarLegajo & "  BETWEEN NRO_DESDE AND NRO_HASTA "
'
'               If chkSolicitarAño.value = 1 Then
'                   Año = InputBox("Ingrese el año de 4 cifras")
'                    Sql = Sql & vbCrLf & " AND  Year(FECHA_DESDE) = " & Año
'               End If
'
'
'                Set rslegajos = New ADODB.Recordset
'                rslegajos.Open Sql, ConActiva, 0, 1
'                Do While Not rslegajos.EOF
'
'                If IsNull(rslegajos!FECHA_DESDE) Then
'                    detalle = " Nro_desde: " & rslegajos!NRO_DESDE & "   Nro_hasta:" & rslegajos!NRO_HASTA
'                Else
'                    detalle = " Nro_desde: " & rslegajos!NRO_DESDE & "   Nro_hasta:" & rslegajos!NRO_HASTA & "  AÑO:" & Format(rslegajos!FECHA_DESDE, "YY")
'                End If
'
'                grdResultadoBusqueda.AddItem vbTab & "Referencias" & vbTab & detalle & vbTab & rslegajos!NRO_CAJA & vbTab & txtBuscarLegajo & vbTab & "rslegajos!Cod_Estado" & vbTab & rslegajos!NRO_CAJA
'                rslegajos.MoveNext
'                Loop
'         Else
'            MsgBox "No se realizo la busqueda en referencia puesto que no numerico "
'         End If
'          Else
'            MsgBox "No se realizo la busqueda en referencia puesto que se asigno un incice "
'        End If
'
'
'End If
'
'
'
'
'   Rem  ----------- orde de documentacion ----------
'
'If chkRearchivoLote.value = 1 Then
'
'
'
'                If CboCampo.Text = "CLIENTE_LEGAJO_LETRA" Then
'
'
'        Sql = " SELECT COD_DOCUMENTACION, ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO, ORDENAR_DOCUMENTACION_DETALLE.COD_INDICE,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE, ORDENAR_DOCUMENTACION_DETALLE.COD_ESTADO,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.CONTENEDOR_PROV, INDICES.DESCRIPCION,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Nro_Caja, Orden"
'        Sql = Sql & " FROM  ORDENAR_DOCUMENTACION_DETALLE LEFT OUTER JOIN"
'        Sql = Sql & " INDICES ON ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = INDICES.COD_CLIENTE AND"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Indice = INDICES.Indice"
'        Sql = Sql & " WHERE  ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO like '%" & txtBuscarLegajo & "%'"
'        Sql = Sql & " AND ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE =" & ctlCliente.Valor
'
'
'      Else
'         Sql = " SELECT COD_DOCUMENTACION, ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO, ORDENAR_DOCUMENTACION_DETALLE.COD_INDICE,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE, ORDENAR_DOCUMENTACION_DETALLE.COD_ESTADO,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.CONTENEDOR_PROV, INDICES.DESCRIPCION,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Nro_Caja , ORDEN"
'        Sql = Sql & " FROM  ORDENAR_DOCUMENTACION_DETALLE LEFT OUTER JOIN"
'        Sql = Sql & " INDICES ON ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = INDICES.COD_CLIENTE AND"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Indice = INDICES.Indice"
'        Sql = Sql & " WHERE  ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO_NUMERO in ( " & Trim(txtBuscarLegajo.Text) & ")"
'        Sql = Sql & " AND ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE =" & ctlCliente.Valor
'
'
'      End If
'
'
'        Set rslegajos = New ADODB.Recordset
'        rslegajos.Open Sql, ConActiva, 0, 1
'        Do While Not rslegajos.EOF
'            grdResultadoBusqueda.AddItem vbTab & "Orden Documentacion" & vbTab & rslegajos!Descripcion & vbTab & rslegajos!COD_DOCUMENTACION & vbTab & rslegajos!Elemento & vbTab & rslegajos!Cod_Estado & vbTab & rslegajos!Cod_Nro_Caja & vbTab & vbTab & "Orden: " & rslegajos!COD_DOCUMENTACION & "  Pos:" & rslegajos!Orden & "   Prov: " & rslegajos!Contenedor_Prov
'            rslegajos.MoveNext
'        Loop
'End If
'
'
''    ------------- rearchiv digital -------------
'
'            If ChkRearchivoDigital.value = 1 Then
'
'                If CboCampo.Text = "CLIENTE_LEGAJO_LETRA" Then
'                    Sql = "  SELECT     DOCUMENTOS_DIGITALES.id ,DOCUMENTOS_DIGITALES.COD_CLIENTE, DOCUMENTOS_DIGITALES.NRO_CAJA, DOCUMENTOS_DIGITALES.COD_ESTADO,LOTE,"
'                    Sql = Sql & " DOCUMENTOS_DIGITALES.LETRA_DESDE , INDICES.DESCRIPCION , IMPRESO,  LOTE, IMAGEN_ORIGEN,NOMBRE"
'                    Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
'                    Sql = Sql & "  DOCUMENTOS_DIGITALES_LOTE ON"
'                    Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE LEFT OUTER"
'                    Sql = Sql & " Join  INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
'                    Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = " & ctlCliente.Valor
'                    Sql = Sql & " AND DOCUMENTOS_DIGITALES.LETRA_DESDE LIKE '%" & Trim(txtBuscarLegajo) & "%'"
'
'                Else
'
'                        Sql = "  SELECT     DOCUMENTOS_DIGITALES.id ,DOCUMENTOS_DIGITALES.COD_CLIENTE, DOCUMENTOS_DIGITALES.NRO_CAJA, DOCUMENTOS_DIGITALES.COD_ESTADO,LOTE,"
'                    Sql = Sql & " DOCUMENTOS_DIGITALES.LETRA_DESDE , INDICES.DESCRIPCION , IMPRESO,  LOTE, IMAGEN_ORIGEN,NOMBRE"
'                    Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
'                    Sql = Sql & "  DOCUMENTOS_DIGITALES_LOTE ON"
'                    Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE LEFT OUTER"
'                    Sql = Sql & " Join  INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
'                    Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = " & ctlCliente.Valor
'                    Sql = Sql & " AND DOCUMENTOS_DIGITALES.NRO_DESDE IN ( " & txtBuscarLegajo & ")"
'                   End If
'
'
'
'
'                Set rslegajos = New ADODB.Recordset
'                rslegajos.Open Sql, ConActiva, 0, 1
'                Do While Not rslegajos.EOF
'                    grdResultadoBusqueda.AddItem vbTab & "Rearchivo Digital" & vbTab & rslegajos!Descripcion & vbTab & rslegajos!ID & vbTab & rslegajos!LETRA_DESDE & vbTab & "IMPRESO:" & rslegajos!Impreso & vbTab & rslegajos!NRO_CAJA & vbTab & Trim(rslegajos!Nombre) & vbTab & Trim(rslegajos!lote) & " Pos: " & rslegajos!IMAGEN_ORIGEN
'                    rslegajos.MoveNext
'                Loop
'            End If
'
'            txtBuscarLegajo.Text = ""
'    Exit Sub
'salir:
'    MsgBox "Verifique los dato", vbInformation
'    txtBuscarLegajo.Text = ""

'
'On Error GoTo salir
'    Set rslegajos = New ADODB.Recordset
'    rslegajos.CursorLocation = adUseClient
'    Dim Sql As String
'    Dim Filtro As String
'    Dim detalle As String
'    Dim Año As String
'
'    If cboCampo.Text = "" Then
'        MsgBox "Ingrese el Campo ", vbInformation
'        Exit Sub
'    End If
'
'    If IsNull(ctlCliente.Valor) Then
'        MsgBox "Ingrese el Cliente ", vbInformation
'        Exit Sub
'    End If
'    txtBuscarLegajo = Replace(txtBuscarLegajo, vbCrLf, "")
'
'    If Mid(txtBuscarLegajo.Text, Len(txtBuscarLegajo.Text)) = "," Then
'        txtBuscarLegajo.Text = Mid(txtBuscarLegajo.Text, 1, Len(txtBuscarLegajo.Text) - 1)
'    End If
'
'    Sql = " SELECT  INDICES.DESCRIPCION, LEGAJOS.ID_CLIENTE_LEGAJO, LEGAJOS.CLIENTE_LEGAJO , DESCRIPCION_REMITO , LEGAJOS.NRO_CAJA, LEGAJOS.COD_ESTADO ,NOMBRE  "
'    Sql = Sql & vbCrLf & " FROM LEGAJOS LEFT OUTER JOIN"
'    Sql = Sql & vbCrLf & " INDICES ON LEGAJOS.COD_CLIENTE = INDICES.COD_CLIENTE AND LEGAJOS.COD_INDICE = INDICES.INDICE"
'    Sql = Sql & vbCrLf & " where LEGAJOS.COD_CLIENTE = " & ctlCliente.Valor & " And "
'    If lblIndice.Caption <> "" Then
'         Sql = Sql & vbCrLf & " COD_INDICE like '" & lblIndice.Caption & "%' AND "
'    End If
'
'    Select Case cboCampo.Text
'        Case "ID_CLIENTE_LEGAJO"
'            Filtro = " ID_CLIENTE_LEGAJO IN (" & txtBuscarLegajo & ")"
'        Case "CLIENTE_LEGAJO_LETRA"
'            Filtro = " CLIENTE_LEGAJO like '%" & txtBuscarLegajo & "%'"
'        Case "CLIENTE_LEGAJO_NUMERO"
'             Filtro = " NUMERO_LEGAJO_CLIENTE IN (" & txtBuscarLegajo & ")"
'        Case "NOMBRE"
'             Filtro = " NOMBRE like '%" & txtBuscarLegajo & "%'"
'        Case "DESCRIPCION"
'            Filtro = " DESCRIPCION like '%" & txtBuscarLegajo & "%'"
'    End Select
'
'   Rem TitulosBuscar
'    Sql = Sql & Filtro
'        rslegajos.Open Sql, ConActiva, 0, 1
'
'If (rslegajos.EOF) Then
' MsgBox "No exsite el legajo"
'End If
'
'    Do While Not rslegajos.EOF
'        grdResultadoBusqueda.AddItem grdResultadoBusqueda.Rows & vbTab & "Legajos" & vbTab & rslegajos!Descripcion & vbTab & rslegajos!ID_CLIENTE_LEGAJO & vbTab & Trim(rslegajos!CLIENTE_LEGAJO & "  " & Replace(rslegajos!DESCRIPCION_REMITO, Chr(9), "")) & vbTab & rslegajos!Cod_Estado & vbTab & rslegajos!NRO_CAJA & vbTab & rslegajos!Nombre
'
'        rslegajos.MoveNext
'    Loop
'
'    If cboCampo.Text = "ID_CLIENTE_LEGAJO" Then
'        Exit Sub
'    End If
'
'    ' Referencia
'
'If chkReferencias.value = 1 Then
'    If lblIndice.Caption <> "" Then
'
'
'
'        Sql = " SELECT COD_CLIENTE, NRO_CAJA, NRO_DESDE, NRO_HASTA,FECHA_DESDE, INDICE "
'        Sql = Sql & vbCrLf & " From REFERENCIAS "
'        Sql = Sql & vbCrLf & " WHERE  COD_CLIENTE= " & ctlCliente.Valor & " And "
'
'        If lblIndice.Caption <> "" Then
'             Sql = Sql & vbCrLf & " INDICE like '" & lblIndice.Caption & "%' AND "
'        End If
'        If IsNumeric(txtBuscarLegajo.Text) Then
'            Sql = Sql & vbCrLf & txtBuscarLegajo & "  BETWEEN NRO_DESDE AND NRO_HASTA "
'
'               If chkSolicitarAño.value = 1 Then
'                   Año = InputBox("Ingrese el año de 4 cifras")
'                    Sql = Sql & vbCrLf & " AND  Year(FECHA_DESDE) = " & Año
'               End If
'
'
'                Set rslegajos = New ADODB.Recordset
'                rslegajos.Open Sql, ConActiva, 0, 1
'                Do While Not rslegajos.EOF
'
'                If IsNull(rslegajos!FECHA_DESDE) Then
'                    detalle = " Nro_desde: " & rslegajos!NRO_DESDE & "   Nro_hasta:" & rslegajos!NRO_HASTA
'                Else
'                    detalle = " Nro_desde: " & rslegajos!NRO_DESDE & "   Nro_hasta:" & rslegajos!NRO_HASTA & "  AÑO:" & Format(rslegajos!FECHA_DESDE, "YY")
'                End If
'
'                grdResultadoBusqueda.AddItem vbTab & "Referencias" & vbTab & detalle & vbTab & rslegajos!NRO_CAJA & vbTab & txtBuscarLegajo & vbTab & "rslegajos!Cod_Estado" & vbTab & rslegajos!NRO_CAJA
'                rslegajos.MoveNext
'                Loop
'         Else
'            MsgBox "No se realizo la busqueda en referencia puesto que no numerico "
'         End If
'          Else
'            MsgBox "No se realizo la busqueda en referencia puesto que se asigno un incice "
'        End If
'
'
'End If
'
'
'
'
'   Rem  ----------- orde de documentacion ----------
'
'If chkRearchivoLote.value = 1 Then
'
'
'
'                If cboCampo.Text = "CLIENTE_LEGAJO_LETRA" Then
'
'
'        Sql = " SELECT COD_DOCUMENTACION, ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO, ORDENAR_DOCUMENTACION_DETALLE.COD_INDICE,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE, ORDENAR_DOCUMENTACION_DETALLE.COD_ESTADO,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.CONTENEDOR_PROV, INDICES.DESCRIPCION,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Nro_Caja, Orden"
'        Sql = Sql & " FROM  ORDENAR_DOCUMENTACION_DETALLE LEFT OUTER JOIN"
'        Sql = Sql & " INDICES ON ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = INDICES.COD_CLIENTE AND"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Indice = INDICES.Indice"
'        Sql = Sql & " WHERE  ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO like '%" & txtBuscarLegajo & "%'"
'        Sql = Sql & " AND ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE =" & ctlCliente.Valor
'
'
'      Else
'         Sql = " SELECT COD_DOCUMENTACION, ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO, ORDENAR_DOCUMENTACION_DETALLE.COD_INDICE,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE, ORDENAR_DOCUMENTACION_DETALLE.COD_ESTADO,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.CONTENEDOR_PROV, INDICES.DESCRIPCION,"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Nro_Caja , ORDEN"
'        Sql = Sql & " FROM  ORDENAR_DOCUMENTACION_DETALLE LEFT OUTER JOIN"
'        Sql = Sql & " INDICES ON ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE = INDICES.COD_CLIENTE AND"
'        Sql = Sql & " ORDENAR_DOCUMENTACION_DETALLE.Cod_Indice = INDICES.Indice"
'        Sql = Sql & " WHERE  ORDENAR_DOCUMENTACION_DETALLE.ELEMENTO_NUMERO in ( " & Trim(txtBuscarLegajo.Text) & ")"
'        Sql = Sql & " AND ORDENAR_DOCUMENTACION_DETALLE.COD_CLIENTE =" & ctlCliente.Valor
'
'
'      End If
'
'
'        Set rslegajos = New ADODB.Recordset
'        rslegajos.Open Sql, ConActiva, 0, 1
'        Do While Not rslegajos.EOF
'            grdResultadoBusqueda.AddItem vbTab & "Orden Documentacion" & vbTab & rslegajos!Descripcion & vbTab & rslegajos!COD_DOCUMENTACION & vbTab & rslegajos!Elemento & vbTab & rslegajos!Cod_Estado & vbTab & rslegajos!Cod_Nro_Caja & vbTab & vbTab & "Orden: " & rslegajos!COD_DOCUMENTACION & "  Pos:" & rslegajos!Orden & "   Prov: " & rslegajos!Contenedor_Prov
'            rslegajos.MoveNext
'        Loop
'End If
'
'
''    ------------- rearchiv digital -------------
'
'            If ChkRearchivoDigital.value = 1 Then
'
'                If cboCampo.Text = "CLIENTE_LEGAJO_LETRA" Then
'                    Sql = "  SELECT     DOCUMENTOS_DIGITALES.id ,DOCUMENTOS_DIGITALES.COD_CLIENTE, DOCUMENTOS_DIGITALES.NRO_CAJA, DOCUMENTOS_DIGITALES.COD_ESTADO,LOTE,"
'                    Sql = Sql & " DOCUMENTOS_DIGITALES.LETRA_DESDE , INDICES.DESCRIPCION , IMPRESO,  LOTE, IMAGEN_ORIGEN,NOMBRE"
'                    Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
'                    Sql = Sql & "  DOCUMENTOS_DIGITALES_LOTE ON"
'                    Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE LEFT OUTER"
'                    Sql = Sql & " Join  INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
'                    Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = " & ctlCliente.Valor
'                    Sql = Sql & " AND DOCUMENTOS_DIGITALES.LETRA_DESDE LIKE '%" & Trim(txtBuscarLegajo) & "%'"
'
'                Else
'
'                        Sql = "  SELECT     DOCUMENTOS_DIGITALES.id ,DOCUMENTOS_DIGITALES.COD_CLIENTE, DOCUMENTOS_DIGITALES.NRO_CAJA, DOCUMENTOS_DIGITALES.COD_ESTADO,LOTE,"
'                    Sql = Sql & " DOCUMENTOS_DIGITALES.LETRA_DESDE , INDICES.DESCRIPCION , IMPRESO,  LOTE, IMAGEN_ORIGEN,NOMBRE"
'                    Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
'                    Sql = Sql & "  DOCUMENTOS_DIGITALES_LOTE ON"
'                    Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE LEFT OUTER"
'                    Sql = Sql & " Join  INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
'                    Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = " & ctlCliente.Valor
'                    Sql = Sql & " AND DOCUMENTOS_DIGITALES.NRO_DESDE IN ( " & txtBuscarLegajo & ")"
'                   End If
'
'
'
'
'                Set rslegajos = New ADODB.Recordset
'                rslegajos.Open Sql, ConActiva, 0, 1
'                Do While Not rslegajos.EOF
'                    grdResultadoBusqueda.AddItem vbTab & "Rearchivo Digital" & vbTab & rslegajos!Descripcion & vbTab & rslegajos!ID & vbTab & rslegajos!LETRA_DESDE & vbTab & "IMPRESO:" & rslegajos!Impreso & vbTab & rslegajos!NRO_CAJA & vbTab & Trim(rslegajos!Nombre) & vbTab & Trim(rslegajos!lote) & " Pos: " & rslegajos!IMAGEN_ORIGEN
'                    rslegajos.MoveNext
'                Loop
'            End If
'
'            txtBuscarLegajo.Text = ""
'    Exit Sub
'salir:
'    MsgBox "Verifique los dato", vbInformation
'    txtBuscarLegajo.Text = ""
'
'
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


Private Sub cmdBuscarRearchivo_Click()
    On Error GoTo salir:
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim i As Integer
        rs.CursorLocation = adUseClient
        If grdSeleccionLegajos.Rows < 1 Then
        
            MsgBox "Ingrese legajos a buscar"
            Exit Sub
        End If
        MousePointer = 11
        Sql = "  SELECT ID_ORDEN_LEGAJO AS Orden, COD_CLIENTE AS Cliente,"
        Sql = Sql & vbCrLf & " COD_ID_CLIENTE_LEGAJO AS Etiqueta,"
        Sql = Sql & vbCrLf & " ELEMENTO AS Legajo, FECHA AS Fecha,COD_ESTADO As Estado , RESPONSABLE_CARGA as Resp_carga"
        Sql = Sql & vbCrLf & " From ORDEN_LEGAJOS_DETALLE, ORDEN_LEGAJOS"
        Sql = Sql & vbCrLf & " Where ORDEN_LEGAJOS_DETALLE.COD_ORDEN_LEGAJO = ORDEN_LEGAJOS.ID_ORDEN_LEGAJO"
        Sql = Sql & vbCrLf & " AND  ( ( COD_CLIENTE =" & grdSeleccionLegajos.TextMatrix(1, 1) & " AND ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO =" & grdSeleccionLegajos.TextMatrix(1, 2) & ")"
        
        
        For i = 2 To grdSeleccionLegajos.Rows - 1
              Sql = Sql & vbCrLf & "  OR ( COD_CLIENTE =" & grdSeleccionLegajos.TextMatrix(i, 1) & " AND ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO =" & grdSeleccionLegajos.TextMatrix(i, 2) & "  )  "
        Next
        Sql = Sql & vbCrLf & " )  ORDER BY COD_CLIENTE,COD_ID_CLIENTE_LEGAJO  , ORDEN_LEGAJOS.ID_ORDEN_LEGAJO"
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        DATOSGRILLA grdRearchivo, rs
        MousePointer = 0
        MsgBox "Operación terminada", vbInformation
        
        Exit Sub
        
salir:
    MsgBox Err.Description
End Sub

Private Sub cmdBuscarRearchivoDigital_Click()
Dim RsDigital As ADODB.Recordset
 Dim Sql As String

If txtNroLegajo.Text <> "" Then
Sql = " SELECT INDICES.DESCRIPCION AS DESC_INDICE,"
Sql = Sql & vbCrLf & " IMPRESO, REARCHIVO_DIGITAL_DETALLE.ELEMENTO,"
Sql = Sql & vbCrLf & " REARCHIVO_DIGITAL_DETALLE.ID,"
Sql = Sql & vbCrLf & " REARCHIVO_DIGITAL_DETALLE.COD_REARCHIVO_DIGITAL,"
Sql = Sql & vbCrLf & "  REARCHIVO_DIGITAL_DETALLE.PASO_INTERNO,"
Sql = Sql & vbCrLf & " REARCHIVO_DIGITAL_DETALLE.DESCRIPCION"
Sql = Sql & vbCrLf & "  From REARCHIVO_DIGITAL_DETALLE, INDICES"
Sql = Sql & vbCrLf & "  Where REARCHIVO_DIGITAL_DETALLE.COD_DOCUMENTO = INDICES.ID_CODIGO_DOCUMENTO"
Sql = Sql & vbCrLf & "      AND (INDICES.COD_CLIENTE = 40) AND"
Sql = Sql & vbCrLf & "    (REARCHIVO_DIGITAL_DETALLE.ELEMENTO LIKE '%" & txtNroLegajo & "%')"
   Else
   Sql = " SELECT INDICES.DESCRIPCION AS DESC_INDICE,"
Sql = Sql & vbCrLf & " IMPRESO, REARCHIVO_DIGITAL_DETALLE.ELEMENTO,"
Sql = Sql & vbCrLf & " REARCHIVO_DIGITAL_DETALLE.ID,"
Sql = Sql & vbCrLf & " REARCHIVO_DIGITAL_DETALLE.COD_REARCHIVO_DIGITAL,"
Sql = Sql & vbCrLf & "  REARCHIVO_DIGITAL_DETALLE.PASO_INTERNO,"
Sql = Sql & vbCrLf & " REARCHIVO_DIGITAL_DETALLE.DESCRIPCION"
Sql = Sql & vbCrLf & "  From REARCHIVO_DIGITAL_DETALLE, INDICES"
Sql = Sql & vbCrLf & "  Where REARCHIVO_DIGITAL_DETALLE.COD_DOCUMENTO = INDICES.ID_CODIGO_DOCUMENTO"
Sql = Sql & vbCrLf & "      AND (INDICES.COD_CLIENTE = 40) AND"
Sql = Sql & vbCrLf & "    (REARCHIVO_DIGITAL_DETALLE.DESCRIPCION LIKE '%" & UCase(txtDescripcion.Text) & "%')"
   
   End If
   
    
       Set RsDigital = New ADODB.Recordset
    RsDigital.CursorLocation = adUseClient
  RsDigital.Open Sql, ConActiva, 0, 1
    Set grdRearchivoDigital.DataSource = RsDigital.DataSource
    grdRearchivoDigital.DataMember = RsDigital.DataMember
    grdRearchivoDigital.ReBind
    grdRearchivoDigital.Refresh
    sstRearchivoDigital.Tab = 0
    
End Sub

Private Sub cmdControlOrden_Click()
Dim Sql As String
Dim Afect As Integer
    If TxtOrden.Text = "" Then
        MsgBox "Ingrese la orden"
        Exit Sub

    End If
    Sql = " Update ORDEN_LEGAJOS SET COD_ESTADO = 4 "
    Sql = Sql & " Where COD_ESTADO = 2 and (ID_ORDEN_LEGAJO = " & TxtOrden.Text & " )"
     Afect = ExecutarSql(Sql)
    If Afect > 0 Then
        MsgBox "La acutalizacion se realizo con exito", vbInformation
     Else
        MsgBox "La acutalizacion no se realizo", vbCritical
    End If
End Sub

Private Sub cmdCopiarExcel_Click()
    CopiarDatosGrilla grdRearchivo
End Sub

Private Sub cmdEntrada_Click()
Dim FechaEntrada As String
Dim lote As Long
Dim Sql As String
Dim rs As ADODB.Recordset
Dim Valor As String


If IsNull(ctlCliente.Valor) Then
        MsgBox "Ingrese el codigo del cliente"
    Exit Sub
End If

FechaEntrada = InputBox("Ingrese la fecha de entrada", "Fecha", Format(Now, "DD/MM/YYYY"))
lote = InputBox("Ingrese el N de lote el 0 toma solo la fecha", "Lote", "0")
Sql = " SELECT     ELEMENTO, TIPO, COD_CLIENTE"
Sql = Sql & " From ENTRADA"
Sql = Sql & "  Where TIPO = 3"
Sql = Sql & "  AND COD_CLIENTE =" & ctlCliente.Valor
Sql = Sql & "  AND FECHA =" & FechaFormato(FechaEntrada)
If lote <> 0 Then
    Sql = Sql & "  AND lote = " & lote
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

If Valor <> "" Then
    txtBuscarLegajo.Text = Mid(Valor, 1, Len(Valor) - 1)
End If

End Sub

Private Sub cmdImformeCajaRearchivo_Click()
Dim Sql As String




If txtCajaRearchivo.Text <> "" Then
Sql = " SELECT  * "
Sql = Sql & vbCrLf & "  From V_ORDEN_LEGAJOS"
Sql = Sql & vbCrLf & " where REARCHIVO_CAJA = " & txtCajaRearchivo.Text
 Sql = Sql & vbCrLf & "  ORDER BY ID_ORDEN_LEGAJO"


        
        frmReportes.ImprimirReporte PasoReportes + "rptLegajosReachivoCaja.rpt", Sql, True
Else
MsgBox "iNGRESE LA CAJA DE REARCHIVO"
End If


End Sub

Private Sub cmdInsertarBusqueda_Click()
        Dim Sql  As String
        Dim i As Integer
        Dim ID_LOTE_BUSQUEDA, TIPO, fecha, COD_CLIENTE, NRO_CAJA, Legajo, Descripcion As String
        Dim MaxLote As Integer
        Dim rs As New ADODB.Recordset
        
        
        rs.Open " SELECT MAX(ID_LOTE_BUSQUEDA) AS maxLote From TEM_BUSQUEDA", ConActiva, 0, 1
        ID_LOTE_BUSQUEDA = rs!MaxLote + 1
        
        For i = 1 To grdVarios.Rows - 1
            
            TIPO = "'" & grdVarios.TextMatrix(i, 1) & "'"
            fecha = "'" & SysDate & "'"
            COD_CLIENTE = grdVarios.TextMatrix(i, 2)
            NRO_CAJA = grdVarios.TextMatrix(i, 3)
            Legajo = "'" & grdVarios.TextMatrix(i, 4) & "'"
            Descripcion = "'" & grdVarios.TextMatrix(i, 5) & "'"
            
            Sql = " INSERT INTO TEM_BUSQUEDA "
            Sql = Sql & " (ID_LOTE_BUSQUEDA,TIPO "
            Sql = Sql & "  , FECHA, COD_CLIENTE "
            Sql = Sql & "  , NRO_CAJA, LEGAJO "
            Sql = Sql & "  , DESCRIPCION ) "
            Sql = Sql & "  VALUES "
            Sql = Sql & " (" & ID_LOTE_BUSQUEDA & "," & TIPO
            Sql = Sql & " ," & fecha & "," & COD_CLIENTE
            Sql = Sql & " ," & NRO_CAJA & "," & Legajo
            Sql = Sql & "," & Descripcion & ")"
            ExecutarSql Sql
             
        Next
        
        MsgBox "Lote de busqueda " & ID_LOTE_BUSQUEDA
        TitulosVarios
        ReporteBusqueda CInt(ID_LOTE_BUSQUEDA)
End Sub

Private Sub cmdLecturaLegajo_Click()
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
 
MsgBox "Lectura es :" & MaxLectura
End Sub

Private Sub cmdLimpiar_Click()
TitulosVarios
End Sub

Private Sub cmdLimpiarLegajos_Click()
TitulosBuscar
    TitulosSeleccionLegajos
End Sub

Private Sub cmdOrdenCompleto_Click()
Dim Sql As String
Dim Afect As Integer
'    If IsNull(ctlPersonal.Valor) Then
'        MsgBox "Ingrese el personal"
'        Exit Sub
'
'    End If
    If Not IsDate(txtFechaOrden.Text) Then
        MsgBox "Ingrese la fecha dl orden"
        Exit Sub
    End If

Sql = " Update ORDEN_LEGAJOS SET COD_ESTADO = 2, RESPONSABLE_ORDEN = " & ctlPersonal.Valor & ",fecha_orden = " & FechaFormato(txtFechaOrden.Text)
Sql = Sql & " Where (ID_ORDEN_LEGAJO = " & TxtOrden.Text & " ) And (Cod_Estado = 0)"
ExecutarSql Sql
    Sql = " Update ORDEN_LEGAJOS_DETALLE Set COD_ESTADO_DETALLE = 2 Where COD_ESTADO_DETALLE = 0 AND COD_ORDEN_LEGAJO = " & TxtOrden.Text
     Afect = ExecutarSql(Sql)
        If Afect > 0 Then
            Dim rs As ADODB.Recordset
            Dim cantidad As Integer
            Dim Cliente As Integer
            Set rs = New ADODB.Recordset
            rs.Open "SELECT COUNT(*) AS CANTIDAD FROM ORDEN_LEGAJOS_DETALLE WHERE COD_ORDEN_LEGAJO = " & TxtOrden.Text, ConActiva, 0, 1
            cantidad = rs!cantidad
            Set rs = New ADODB.Recordset
            rs.Open "SELECT COD_CLIENTE FROM ORDEN_LEGAJOS WHERE ID_ORDEN_LEGAJO = " & TxtOrden.Text, ConActiva, 0, 1

            Cliente = rs!COD_CLIENTE

            Cliente = rs!COD_CLIENTE

        If Not rs.EOF Then
            InsertarProducion ctlPersonal.Valor, 11, "Orden legajo:" & TxtOrden.Text, cantidad, Cliente
            TxtOrden.Text = ""
        Else
            MsgBox "No se relizo la actualizacion", vbInformation
        End If
Else
 MsgBox "No se relizo la actualizacion", vbInformation
 
End If
End Sub

Private Sub cmdOrdenesPendientes_Click()
 Dim Sql As String
        Sql = " SELECT CANTIDAD, ID_ORDEN_LEGAJO, COD_CLIENTE,"
        Sql = Sql & vbCrLf & "  COD_ESTADO, FECHA, RESPONSABLE_CARGA,"
        Sql = Sql & vbCrLf & "  RESPONSABLE_ORDEN, RESPONSABLE_CONTROL,"
        Sql = Sql & vbCrLf & "  FECHA_ORDEN"
        Sql = Sql & vbCrLf & " From ORDEN_LEGAJOS"
        Sql = Sql & vbCrLf & " Where (Cod_Estado = 0)"
        Sql = Sql & vbCrLf & " order by ID_ORDEN_LEGAJO "
        frmReportes.ImprimirReporte PasoReportes + "rptOrdenUbicacionLegajosControl.rpt", Sql, True
End Sub

Private Sub cmdOrdenesPendientesControl_Click()
 Dim Sql As String
        Sql = " SELECT CANTIDAD, ID_ORDEN_LEGAJO, COD_CLIENTE,"
        Sql = Sql & vbCrLf & "  COD_ESTADO, FECHA, RESPONSABLE_CARGA,"
        Sql = Sql & vbCrLf & "  RESPONSABLE_ORDEN, RESPONSABLE_CONTROL,"
        Sql = Sql & vbCrLf & "  FECHA_ORDEN"
        Sql = Sql & vbCrLf & " From ORDEN_LEGAJOS"
        Sql = Sql & vbCrLf & " Where (Cod_Estado = 2)"
        Sql = Sql & vbCrLf & " order by ID_ORDEN_LEGAJO "
        frmReportes.ImprimirReporte PasoReportes + "rptOrdenUbicacionLegajosControl.rpt", Sql, True
End Sub

Private Sub cmdOrdenLegajos_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String
    Dim MaxOrden As Integer
    Dim i As Integer
    On Error GoTo salir
    
    Dim CONlEGAJOS As New ADODB.Connection
    CONlEGAJOS.Open strConBasa
       
    If IsNull(ctlPersonal.Valor) Then
        MsgBox "Ingrese Personal", vbInformation
        Exit Sub
    End If
    If grdSeleccionLegajos.Rows = 1 Then
        MsgBox "No a selecionado los registos", vbInformation
        Exit Sub
    End If
    
    If Trim(txtCajaRearchivo.Text) = "" Then
        MsgBox "No ingreso una caja", vbInformation
        Exit Sub
    Else
        If Not IsNumeric(txtCajaRearchivo.Text) Then
            MsgBox "La Caja no es un numero", vbInformation
            Exit Sub
        End If
    End If

If ControlCajasRearchivo(txtCajaRearchivo.Text, txtDigito.Text) = False Then
    Exit Sub

End If

           
           
           MousePointer = 11
           rs.Open "Select max(ID_ORDEN_LEGAJO) as maxid from orden_legajos ", ConActiva, 0, 1
           MaxOrden = rs!MaxID + 1
           
           
           With grdSeleccionLegajos
               For i = 1 To .Rows - 1
                   Sql = " INSERT INTO ORDEN_LEGAJOS_DETALLE "
                   Sql = Sql & "(COD_ORDEN_LEGAJO, COD_ID_CLIENTE_LEGAJO, ORDEN,"
                   Sql = Sql & "  ELEMENTO, NRO_CAJA,COD_ESTADO_DETALLE)VALUES ("
                   Sql = Sql & MaxOrden & "," & .TextMatrix(i, 2) & ",0,"
                   Sql = Sql & "'" & .TextMatrix(i, 3) & "'," & .TextMatrix(i, 4) & ",0)"
                   ExecutarSql Sql
               
                Sql = "  Update LEGAJOS "
                Sql = Sql & " SET REARCHIVO_CAJA =" & txtCajaRearchivo.Text
                Sql = Sql & "  Where ID_CLIENTE_LEGAJO = " & .TextMatrix(i, 2)
                Sql = Sql & "  And COD_CLIENTE = " & ctlCliente.Valor
               ExecutarSql Sql
               
               
               Next
                Sql = " INSERT INTO ORDEN_LEGAJOS"
                Sql = Sql & "(ID_ORDEN_LEGAJO, COD_CLIENTE, COD_ESTADO, FECHA,RESPONSABLE_CARGA,CANTIDAD , REARCHIVO_CAJA ) VALUES ("
                Sql = Sql & MaxOrden & "," & ctlCliente.Valor & ", 0 ," & SysDate & "," & ctlPersonal.Valor & "," & i & "," & txtCajaRearchivo.Text & "  )"
                ExecutarSql Sql
                InsertarProducion ctlPersonal.Valor, 10, "Carga Orden Legajos:" & MaxOrden, .Rows - 1, ctlCliente.Valor
           End With

           MsgBox "La orden se realizo con exito", vbInformation
           Sql = "  SELECT * From V_ORDEN_LEGAJOS "
           Sql = Sql & vbCrLf & " Where ID_ORDEN_LEGAJO =" & MaxOrden
           Sql = Sql & vbCrLf & " Order By  ESTANTERIA,NRO_CAJA,COD_ID_CLIENTE_LEGAJO"
           frmReportes.ImprimirReporte PasoReportes + "rptOrdenUbicacionLegajos.rpt", Sql, True
           MousePointer = 0
           
           Exit Sub
salir:
          
           MsgBox "No se realizo la actualización"
           
End Sub



Private Sub cmdPasarTodosLegajos_Click()
Dim i As Integer
On Error GoTo salir
  With grdResultadoBusqueda
For i = 1 To .Rows - 1
     If .TextMatrix(i, 1) = "Legajos" Then
  
      grdSeleccionLegajos.AddItem i & vbTab & ctlCliente.Valor & vbTab & .TextMatrix(i, 3) & vbTab & .TextMatrix(i, 4) & vbTab & .TextMatrix(i, 6)
    End If
    
 Next
   End With
salir:

End Sub




Private Sub Command3_Click()

End Sub

Private Sub cmdReImprimirOrden_Click()
    Dim Sql As String
    Dim Orden As Integer
    Orden = InputBox("Ingrese la Orden")
            Sql = "  SELECT * From V_ORDEN_LEGAJOS "
           Sql = Sql & vbCrLf & " Where ID_ORDEN_LEGAJO =" & Orden
           Sql = Sql & vbCrLf & " Order By  ESTANTERIA,NRO_CAJA,COD_ID_CLIENTE_LEGAJO"
           frmReportes.ImprimirReporte PasoReportes + "rptOrdenUbicacionLegajos.rpt", Sql, True
           MousePointer = 0
End Sub

Private Sub ReporteBusqueda(Lote_Busqueda As Integer)
  Dim sSQL As String
    
    
        sSQL = " SELECT  ID_LOTE_BUSQUEDA, FECHA, TIPO, COD_CLIENTE, NRO_CAJA, LEGAJO, DESCRIPCION, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
        sSQL = sSQL & " UB_PROVISORIA "
        sSQL = sSQL & " FROM  V_TEM_BUSQUEDA "
        sSQL = sSQL & " where ID_LOTE_BUSQUEDA = " & Lote_Busqueda
      Rem   sSQL = sSQL & " ORDER BY ESTANTERIA, HORIZONTAL, VERTICAL "
       frmReportes.ImprimirReporte PasoReportes + "BusquedaGeneral.rpt", sSQL, True


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

Private Sub Command1_Click()
Dim RsDigital As New ADODB.Recordset
Dim Sql As String


   
 
    
   If Trim(txtCajaRearchivo.Text) <> "" Then
   
  
Sql = "SELECT     ID_ORDEN_LEGAJO, REARCHIVO_CAJA, COD_CLIENTE, COD_ESTADO, FECHA, RESPONSABLE_CARGA, RESPONSABLE_ORDEN, RESPONSABLE_CONTROL,"
Sql = Sql & " FECHA_ORDEN , cantidad"
Sql = Sql & " From basasql.dbo.ORDEN_LEGAJOS"
Sql = Sql & " Where REARCHIVO_CAJA = " & txtCajaRearchivo.Text
Sql = Sql & " ORDER BY ID_ORDEN_LEGAJO"
   
     RsDigital.Open Sql, ConActiva, 0, 1
    Set grdRearchivo.DataSource = RsDigital.DataSource
    grdRearchivo.DataMember = RsDigital.DataMember
    grdRearchivo.ReBind
    grdRearchivo.Refresh
  
    Else
    
    MsgBox "Ingrese la caja de rearchivo"
    End If

End Sub

Private Sub ctlCliente_Click()
ctlIndiceLegajo.Actualizar ctlCliente.Valor, Nulo, 0
lblIndice.Caption = ""
End Sub

Private Sub ctlIndiceLegajo_DblClick()
    lblIndice.Caption = ctlIndiceLegajo.Item_Selecionado
End Sub


Private Sub ctlIndiceLegajo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuArbol
End If
End Sub

Private Sub Form_Load()
    TitulosSeleccionLegajos
    TitulosBuscar
    ctlCliente.TipoControl = Cliente
    ctlPersonal.TipoControl = Personal
    ctlClienteBuscarLegajo.TipoControl = Cliente
    TitulosVarios
End Sub

Private Sub Form_Resize()
'    sstLegajos.Width = frmBuscarLegajos.Width - 400
'    sstLegajos.Height = frmBuscarLegajos.Height - 1000
'    sstRearchivoDigital.Width = sstLegajos.Width - 400
'    sstRearchivoDigital.Height = sstLegajos.Height - 1500
'    vieRearchivoDigital.Width = sstRearchivoDigital.Width - 200
'    vieRearchivoDigital.Height = sstRearchivoDigital.Height - 900
'    grdRearchivoDigital.Width = vieRearchivoDigital.Width
'    grdRearchivoDigital.Height = vieRearchivoDigital.Height
End Sub


Public Sub TitulosSeleccionLegajos()
    grdSeleccionLegajos.Cols = 5
    grdSeleccionLegajos.Rows = 1
    grdSeleccionLegajos.ColAlignment(1) = 4
    grdSeleccionLegajos.ColAlignment(2) = 4
    grdSeleccionLegajos.ColAlignment(3) = 4
    grdSeleccionLegajos.ColWidth(0) = 500
    grdSeleccionLegajos.ColWidth(1) = 2000
    grdSeleccionLegajos.ColWidth(2) = 2000
    grdSeleccionLegajos.ColWidth(3) = 2000
    grdSeleccionLegajos.ColWidth(4) = 2000
    grdSeleccionLegajos.TextMatrix(0, 1) = "Cliente"
    grdSeleccionLegajos.TextMatrix(0, 2) = "Etiqueta"
    grdSeleccionLegajos.TextMatrix(0, 3) = "Legajo Cliente"
    grdSeleccionLegajos.TextMatrix(0, 4) = "Caja"

End Sub



Public Sub TitulosBuscar()
With grdResultadoBusqueda
    .Clear
    .Cols = 9
    .Rows = 1
    .ColAlignment(1) = 0
    .ColAlignment(2) = 0
    .ColAlignment(3) = 0
    .ColAlignment(4) = 0
    .ColAlignment(5) = 0
    .ColAlignment(6) = 0
    .ColAlignment(7) = 0
    .ColAlignment(8) = 0
 
    .ColWidth(0) = 400
    .ColWidth(1) = 2000
    .ColWidth(2) = 3500
    .ColWidth(3) = 1000
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1000
    .ColWidth(7) = 3000
    .ColWidth(8) = 2000
    .TextMatrix(0, 1) = "Proceso"
    .TextMatrix(0, 2) = "Tipo Doc"
    .TextMatrix(0, 3) = "Etiqueta"
    .TextMatrix(0, 4) = "Legajo"
    .TextMatrix(0, 5) = "Estado"
    .TextMatrix(0, 6) = "Caja"
    .TextMatrix(0, 7) = "Nombre"
     .TextMatrix(0, 8) = "Lote"
     
    End With

End Sub

Public Sub TitulosVarios()
With grdVarios
    .Clear
    .Cols = 6
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
    End With

End Sub




Private Sub grdRearchivoDigital_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then

 PopupMenu mnuImagen

 
 
End If


End Sub

Private Sub grdResultadoBusqueda_DblClick()
Dim TIPO As String
Dim Cliente As String
Dim Caja As String
Dim Elemento As String
Dim Descripcion As String


    TIPO = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 1)
    Cliente = ctlCliente.Valor
    Caja = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 6)
    Elemento = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 4)
    
    Select Case grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 1)
    Case "Legajos"
       Descripcion = "Etiqueta :" & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 3)
    Case "Referencias"
         Descripcion = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 2)
    Case "Rearchivo Digital"
         Descripcion = Trim(grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 8))
    Case "Orden Documentacion"
          Descripcion = Trim(grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 8))
    
    End Select

grdVarios.AddItem "" & vbTab & TIPO & vbTab & Cliente & vbTab & Caja & vbTab & Elemento & vbTab & Descripcion


If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 1) = "Legajos" Then
    
    
    If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 5) <> 3 Then

        grdSeleccionLegajos.AddItem "" & vbTab & ctlCliente.Valor & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 3) & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 4) & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 6)
    Else
    
         If MsgBox("El legajo esta en consulta Usted queiere ingresarlo Igual", vbYesNo) = vbYes Then
            grdSeleccionLegajos.AddItem "" & vbTab & ctlCliente.Valor & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 3) & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 4) & vbTab & grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.Row, 6)
         End If
         
    
    End If
    
 
End If



End Sub

Private Sub grdResultadoBusqueda_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
 
    If grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 1) = "Rearchivo Digital" Then
        PopupMenu mnugrdBuscar
    
        
    End If
    
 End If
 
End Sub

Private Sub mnuBuscarLegajo_Click()
 ctlIndiceLegajo.BuscarTipoIndice "Legajo", True
End Sub

Private Sub mnuCopiarPaso_Click()
Clipboard.Clear
Dim ID As Long
ID = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 3)

Clipboard.SetText PasoImagenes & BuscarDirectorioPaso(ID) & "\" & CStr(ID) & ".TIF"
MsgBox "Informacion copiada", vbInformation
End Sub

Private Sub mnugrdVerImagen_Click()
Dim ID As Long
ID = grdResultadoBusqueda.TextMatrix(grdResultadoBusqueda.RowSel, 3)
  
  ctlVerImagenes1.PonerImagen PasoImagenes & BuscarDirectorioPaso(ID) & "\" & CStr(ID) & ".TIF"
  sstLegajos.Tab = 2
  sstRearchivoDigital.Tab = 1
End Sub

Private Sub mnuImprimir_Click()
'Dim i As Integer
'Dim ID As String
'Dim PASO_INTERNO As String
'Dim sql As String
'For i = 1 To grdRearchivoDigital.Columns.Count - 1
'
'  Rem MsgBox grdRearchivoDigital.Columns.Item(i).Caption
'
'  If grdRearchivoDigital.Columns.Item(i).Caption = "ID" Then
'       ID = grdRearchivoDigital.Columns.Item(i).Text
'  End If
'
'  If grdRearchivoDigital.Columns.Item(i).Caption = "PASO_INTERNO" Then
'       PASO_INTERNO = grdRearchivoDigital.Columns.Item(i).Text
'  End If
'
'
'Next
'Dim docOrigen As MODI.Document
' Set docOrigen = New MODI.Document
'                docOrigen.Create PasoImagenesMontemar & PASO_INTERNO & "\" & ID & ".tif"
'
'
'                docOrigen.PrintOut , , , , , False, miPRINT_PAGE
'
'            If MsgBox("Usted desea marcar el archivo como Impreso", vbYesNo) = vbYes Then
'               sql = " Update REARCHIVO_DIGITAL_DETALLE"
'               sql = sql & " SET IMPRESO ='SI'"
'               sql = sql & " Where ID = " & ID
'               ExecutarSql sql
'                cmdBuscarRearchivoDigital_Click
'
'            Else
'
'             sql = " Update REARCHIVO_DIGITAL_DETALLE"
'               sql = sql & " SET IMPRESO ='NO'"
'               sql = sql & " Where ID = " & ID
'               ExecutarSql sql
'
'            End If
'
'

End Sub

Private Sub mnuVerImagen_Click()


'Dim i As Integer
'Dim ID As String
'Dim PASO_INTERNO As String
'For i = 1 To grdRearchivoDigital.Columns.Count - 1
'
'  Rem MsgBox grdRearchivoDigital.Columns.Item(i).Caption
'
'  If grdRearchivoDigital.Columns.Item(i).Caption = "ID" Then
'       ID = grdRearchivoDigital.Columns.Item(i).Text
'  End If
'
'  If grdRearchivoDigital.Columns.Item(i).Caption = "PASO_INTERNO" Then
'       PASO_INTERNO = grdRearchivoDigital.Columns.Item(i).Text
'  End If
'
'
'Next
'vieRearchivoDigital.MostrarImagen PasoImagenesMontemar & PASO_INTERNO & "\" & ID & ".tif"
'
'sstRearchivoDigital.Tab = 1

End Sub

Private Sub mnuVerPDF_Click()
'Dim i As Integer
'Dim ID As String
'Dim PASO_INTERNO As String
'For i = 1 To grdRearchivoDigital.Columns.Count - 1
'
'  Rem MsgBox grdRearchivoDigital.Columns.Item(i).Caption
'
'  If grdRearchivoDigital.Columns.Item(i).Caption = "ID" Then
'       ID = grdRearchivoDigital.Columns.Item(i).Text
'  End If
'
'  If grdRearchivoDigital.Columns.Item(i).Caption = "PASO_INTERNO" Then
'       PASO_INTERNO = grdRearchivoDigital.Columns.Item(i).Text
'  End If
'
'
'Next
'Dim docOrigen As MODI.Document
' Set docOrigen = New MODI.Document
'                docOrigen.Create PasoImagenesMontemar & PASO_INTERNO & "\" & ID & ".tif"
'
'              Rem  docOrigen.PrintOut , , , "Acrobat PDFWriter", , False, miPRINT_PAGE
'              docOrigen.PrintOut , , , "Adobe PDF", , False, miPRINT_PAGE

End Sub

Private Sub txtBuscarLegajo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar_Click
End If
End Sub

Private Sub txtDigito_LostFocus()
If ControlCajasRearchivo(txtCajaRearchivo.Text, txtDigito.Text) Then
Else

End If

End Sub

Private Sub txtLecturaLegajos_KeyPress(KeyAscii As Integer)
        Dim Etiqueta As Long
        Dim Cliente As Integer
    On Error GoTo salir

If KeyAscii = 13 Then
                If UCase(Mid(txtLecturaLegajos.Text, 1, 2)) = "L1" Then
                    Cliente = Mid(txtLecturaLegajos, 3, 3)
                    
                    txtBuscarLegajo = txtBuscarLegajo & CLng(Mid(txtLecturaLegajos, 6, 6)) & ","
                    BuscarLegajosRearchivo ctlCliente.Valor, txtBuscarLegajo.Text
                    txtBuscarLegajo.Text = ""
                    KeyAscii = 0
                End If
                If UCase(Mid(txtLecturaLegajos.Text, 1, 2)) = "L2" Then
                    Etiqueta = CLng(Mid(txtLecturaLegajos, 3))
                    Cliente = ctlCliente.Valor
                    
                    txtBuscarLegajo = txtBuscarLegajo & CLng(Mid(txtLecturaLegajos, 3)) & ","
                    BuscarLegajosRearchivo ctlCliente.Valor, txtBuscarLegajo.Text
                    txtBuscarLegajo.Text = ""
                    KeyAscii = 0
                End If
                If UCase(Mid(txtLecturaLegajos.Text, 1, 2)) = "12" And Len(txtLecturaLegajos.Text) = 12 Then
                    Etiqueta = CLng(Mid(txtLecturaLegajos, 3))
                    Cliente = ctlCliente.Valor
                    
                    txtBuscarLegajo = txtBuscarLegajo & CLng(Mid(txtLecturaLegajos, 3)) & ","
                    BuscarLegajosRearchivo ctlCliente.Valor, txtBuscarLegajo.Text
                    txtBuscarLegajo.Text = ""
                    KeyAscii = 0
                End If
                
                If UCase(Mid(txtLecturaLegajos.Text, 1, 2)) = "12" And Len(txtLecturaLegajos.Text) = 13 Then
                    Etiqueta = CLng(Mid(txtLecturaLegajos, 3, 10))
                    Cliente = ctlCliente.Valor
                    
                    txtBuscarLegajo = txtBuscarLegajo & Etiqueta & ","
                    BuscarLegajosRearchivo ctlCliente.Valor, txtBuscarLegajo.Text
                    txtBuscarLegajo.Text = ""
                    KeyAscii = 0
                End If
                
                
                txtLecturaLegajos = ""
            End If
            Exit Sub
salir:
    MsgBox "Verifique los dato", vbInformation
    txtBuscarLegajo.Text = ""

End Sub

Private Sub txtLegajo_Change()

End Sub

Private Sub txtLegajo_KeyDown(KeyCode As Integer, Shift As Integer)

End Sub

Private Sub txtLegajo_KeyPress(KeyAscii As Integer)
   

End Sub

Private Sub txtLegajoBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        grdCargarBuscarLegajos.AddItem "" & vbTab & ctlClienteBuscarLegajo.Valor & vbTab & txtLegajoBuscar.Text
        txtLegajoBuscar.Text = ""
    End If
End Sub


Public Function ControlCajasRearchivo(Caja As Long, DIGITO As Integer) As Boolean
    ControlCajasRearchivo = True
    Dim rs As New ADODB.Recordset
    Dim Sql As String

        
        Sql = " SELECT     CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, CAJAS.DIGITO_VERIFICADOR"
        Sql = Sql & " FROM         CONTENEDOR INNER JOIN"
        Sql = Sql & " CAJAS ON CONTENEDOR.COD_CLIENTE = CAJAS.FK_CLIENTE AND CONTENEDOR.NRO_CAJA = CAJAS.NRO_CAJA"
        Sql = Sql & " Where CONTENEDOR.NRO_CAJA = " & Caja
        Sql = Sql & " and COD_CLIENTE =291"
        
        rs.Open Sql, strConBasa

        If Not rs.EOF Then
            If rs!estado <> 2 Then
             MsgBox "Estado incorrecto"
             ControlCajasRearchivo = False
             Else
             If rs!Digito_Verificador <> DIGITO Then
                     MsgBox "Digito Incorrecto"
                    ControlCajasRearchivo = False
             Else
             ControlCajasRearchivo = True
             End If
             
            End If
            
        
        
        Else
            MsgBox "No existe la caja"
            ControlCajasRearchivo = False
        End If

End Function




Public Sub BuscarLegajosRearchivo(COD_CLIENTE As Integer, ID_CLIENTE_LEGAJO As Long)

On Error GoTo salir
    Set rslegajos = New ADODB.Recordset
    rslegajos.CursorLocation = adUseClient
    Dim Sql As String

    If Mid(txtBuscarLegajo.Text, Len(txtBuscarLegajo.Text)) = "," Then
        txtBuscarLegajo.Text = Mid(txtBuscarLegajo.Text, 1, Len(txtBuscarLegajo.Text) - 1)
    End If
    Sql = " SELECT  INDICES.DESCRIPCION,nro_desde, letra_desde , LEGAJOS.ID_CLIENTE_LEGAJO, LEGAJOS.CLIENTE_LEGAJO , DESCRIPCION_REMITO , LEGAJOS.NRO_CAJA, LEGAJOS.COD_ESTADO ,NOMBRE  "
    Sql = Sql & vbCrLf & " FROM LEGAJOS LEFT OUTER JOIN"
    Sql = Sql & vbCrLf & " INDICES ON LEGAJOS.COD_CLIENTE = INDICES.COD_CLIENTE AND LEGAJOS.COD_INDICE = INDICES.INDICE"
    Sql = Sql & vbCrLf & " where LEGAJOS.COD_CLIENTE = " & COD_CLIENTE & " And "
    Sql = Sql & vbCrLf & " ID_CLIENTE_LEGAJO =  (" & ID_CLIENTE_LEGAJO & ")"
   rslegajos.Open Sql, ConActiva, 0, 1
    If (rslegajos.EOF) Then
        MsgBox "No exsite el legajo"
        Sql = " Insert Into MONTEMAR_CONTROL_2020(COD_CLIENTE, ID_CLIENTE_LEGAJO)"
        Sql = Sql & vbCrLf & "  VALUES (" & COD_CLIENTE & "," & ID_CLIENTE_LEGAJO & ")"
        ExecutarSql Sql
         txtBuscarLegajo.Text = ""
    End If

    Do While Not rslegajos.EOF

        grdResultadoBusqueda.AddItem grdResultadoBusqueda.Rows & vbTab & "Legajos" & vbTab & rslegajos!Descripcion & vbTab & rslegajos!ID_CLIENTE_LEGAJO & vbTab & Trim(rslegajos!NRO_DESDE & "  " & rslegajos!LETRA_DESDE) & vbTab & rslegajos!Cod_Estado & vbTab & rslegajos!NRO_CAJA & vbTab & rslegajos!Nombre
        rslegajos.MoveNext
    Loop
    Exit Sub
salir:
    MsgBox "Verifique los dato", vbInformation
    txtBuscarLegajo.Text = ""
End Sub
