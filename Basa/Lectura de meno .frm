VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLecturaMemo 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura del Colector"
   ClientHeight    =   8145
   ClientLeft      =   3210
   ClientTop       =   1905
   ClientWidth     =   14280
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8145
   ScaleWidth      =   14280
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Subir Lecturas"
      TabPicture(0)   =   "Lectura de meno .frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblOrden"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Cajaº"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPaso"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "grdLectura"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "grdPasar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAceptar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDigiverificador"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtClienteLectura"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCajaLectura"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtDescripcion"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboTareas"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lstPersonal"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdLimpiarLista"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdMarcarTodos"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Toolbar1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdContarCliente"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "BuscarCajas"
      TabPicture(1)   =   "Lectura de meno .frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdFechaMayor"
      Tab(1).Control(1)=   "txtFechaMayor"
      Tab(1).Control(2)=   "cboTipo"
      Tab(1).Control(3)=   "cmdBuscar"
      Tab(1).Control(4)=   "txtCaja"
      Tab(1).Control(5)=   "cmdCopiarExcel"
      Tab(1).Control(6)=   "txtLectura"
      Tab(1).Control(7)=   "cmdBuscarLectura"
      Tab(1).Control(8)=   "grdBuscar"
      Tab(1).Control(9)=   "Label9"
      Tab(1).Control(10)=   "Label4"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Buscar Lecturas"
      TabPicture(2)   =   "Lectura de meno .frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(1)=   "txtNumeroLecturaFiltro"
      Tab(2).Control(2)=   "grdLecturasCuerpo"
      Tab(2).Control(3)=   "cmdExportarExcel"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Cajas con  Error"
      TabPicture(3)   =   "Lectura de meno .frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdAnalisisBasa"
      Tab(3).Control(1)=   "grdControl"
      Tab(3).Control(2)=   "cmdAnalisisCustodia"
      Tab(3).Control(3)=   "TxtOrdenControl"
      Tab(3).Control(4)=   "GUARDAR"
      Tab(3).Control(5)=   "cmdSubirArchivo"
      Tab(3).Control(6)=   "grdControlError"
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Marcar Cajas"
      TabPicture(4)   =   "Lectura de meno .frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdTipoReferencia"
      Tab(4).Control(1)=   "cmdControlReferencias"
      Tab(4).Control(2)=   "txtLectutaTipoReferencia"
      Tab(4).Control(3)=   "cmdMarcarCajas"
      Tab(4).Control(4)=   "cboTipoReferencia"
      Tab(4).Control(5)=   "Label7"
      Tab(4).Control(6)=   "Label3"
      Tab(4).ControlCount=   7
      Begin VB.CommandButton cmdFechaMayor 
         Caption         =   "..."
         Height          =   315
         Left            =   -65280
         TabIndex        =   53
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtFechaMayor 
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
         Left            =   -66960
         TabIndex        =   52
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdExportarExcel 
         Caption         =   "Excel"
         Height          =   375
         Left            =   -68760
         TabIndex        =   51
         Top             =   840
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid grdLecturasCuerpo 
         Height          =   5895
         Left            =   -74760
         TabIndex        =   50
         Top             =   1680
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   10398
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin MSDataGridLib.DataGrid grdTipoReferencia 
         Height          =   4335
         Left            =   -74400
         TabIndex        =   49
         Top             =   2580
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   7646
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
      Begin VB.CommandButton cmdControlReferencias 
         Caption         =   "Control"
         Height          =   375
         Left            =   -64860
         TabIndex        =   48
         Top             =   1740
         Width           =   1935
      End
      Begin VB.TextBox txtLectutaTipoReferencia 
         Height          =   375
         Left            =   -72420
         TabIndex        =   46
         Text            =   "0"
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton cmdMarcarCajas 
         Caption         =   "Marcar Cajas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -64920
         TabIndex        =   45
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cboTipoReferencia 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "Lectura de meno .frx":008C
         Left            =   -72360
         List            =   "Lectura de meno .frx":008E
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   1080
         Width           =   7095
      End
      Begin VB.CommandButton cmdContarCliente 
         Caption         =   "C.."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   42
         Top             =   2040
         Width           =   495
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Lectura de meno .frx":0090
         Left            =   -74760
         List            =   "Lectura de meno .frx":009D
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   840
         Width           =   1575
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   240
         TabIndex        =   23
         Top             =   780
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1(0)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Aceptar"
               Object.ToolTipText     =   "Aceptar"
               ImageIndex      =   49
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cancelar"
               Object.ToolTipText     =   "Cancelar"
               ImageIndex      =   50
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Borrar"
               Object.ToolTipText     =   "Borrar"
               ImageIndex      =   51
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Conectar"
               Object.ToolTipText     =   "Conectar"
               ImageIndex      =   52
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Desconextar"
               Object.ToolTipText     =   "Desconextar"
               ImageKey        =   "Desconextar"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Refrescar"
               Object.ToolTipText     =   "Refrescar"
               ImageIndex      =   54
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Entrada"
               Object.ToolTipText     =   "Entrada de cajas"
               ImageIndex      =   55
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageIndex      =   57
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Pedro"
               ImageIndex      =   59
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Simular"
               Object.ToolTipText     =   "Importar Excel"
               ImageIndex      =   56
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "AROG"
               ImageIndex      =   58
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "plaReferencia"
               Object.ToolTipText     =   "Planilla de Referencia"
               ImageIndex      =   59
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "CopiarMemo"
               Object.ToolTipText     =   "Copia los archivos al server"
               ImageIndex      =   15
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdAnalisisBasa 
         Caption         =   "Analisis Basa"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -66360
         TabIndex        =   39
         Top             =   840
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid grdControl 
         Height          =   6015
         Left            =   -74640
         TabIndex        =   38
         Top             =   1560
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   10610
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   18
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
      Begin VB.CommandButton cmdAnalisisCustodia 
         Caption         =   "Analisis Custodia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68280
         TabIndex        =   37
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox TxtOrdenControl 
         Height          =   375
         Left            =   -70680
         TabIndex        =   36
         Top             =   840
         Width           =   2175
      End
      Begin VB.CommandButton GUARDAR 
         Caption         =   "cmdGuardar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72540
         TabIndex        =   35
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdSubirArchivo 
         Caption         =   "Subir Archivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74400
         TabIndex        =   34
         Top             =   840
         Width           =   1455
      End
      Begin MSFlexGridLib.MSFlexGrid grdControlError 
         Height          =   5775
         Left            =   -74640
         TabIndex        =   33
         Top             =   1560
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   10186
         _Version        =   393216
         Cols            =   6
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
      Begin VB.CommandButton cmdMarcarTodos 
         Caption         =   "Marcar Todos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9660
         TabIndex        =   32
         Top             =   7260
         Width           =   1575
      End
      Begin VB.CommandButton cmdLimpiarLista 
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11460
         TabIndex        =   31
         Top             =   7260
         Width           =   1275
      End
      Begin VB.ListBox lstPersonal 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6180
         ItemData        =   "Lectura de meno .frx":00C3
         Left            =   9660
         List            =   "Lectura de meno .frx":00CA
         Style           =   1  'Checkbox
         TabIndex        =   30
         Top             =   900
         Width           =   3675
      End
      Begin VB.ComboBox cboTareas 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Lectura de meno .frx":00DB
         Left            =   1440
         List            =   "Lectura de meno .frx":0103
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1620
         Width           =   6975
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4980
         TabIndex        =   24
         Top             =   660
         Visible         =   0   'False
         Width           =   3375
         Begin VB.OptionButton optColector 
            Caption         =   "Colector"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton LecturaLlavero 
            Height          =   255
            Left            =   1320
            TabIndex        =   25
            Top             =   180
            Width           =   255
         End
         Begin VB.Image imgEstado 
            Height          =   360
            Left            =   3540
            Picture         =   "Lectura de meno .frx":0283
            Top             =   60
            Width           =   360
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            Caption         =   "Estado:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   27
            Top             =   180
            Width           =   765
         End
      End
      Begin VB.TextBox txtDescripcion 
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
         Left            =   1440
         TabIndex        =   15
         Top             =   1260
         Width           =   6975
      End
      Begin VB.TextBox txtCajaLectura 
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
         Left            =   2580
         TabIndex        =   14
         Top             =   2100
         Width           =   1035
      End
      Begin VB.TextBox txtClienteLectura 
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
         Left            =   4440
         TabIndex        =   13
         Top             =   2100
         Width           =   795
      End
      Begin VB.TextBox txtDigiverificador 
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
         Left            =   7500
         TabIndex        =   12
         Top             =   2100
         Width           =   315
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "..."
         Height          =   315
         Left            =   7860
         TabIndex        =   11
         Top             =   2100
         Width           =   375
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   315
         Left            =   -71700
         TabIndex        =   8
         Top             =   840
         Width           =   315
      End
      Begin VB.TextBox txtCaja 
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
         Left            =   -72960
         TabIndex        =   6
         Top             =   840
         Width           =   1155
      End
      Begin VB.CommandButton cmdCopiarExcel 
         Caption         =   "Excel"
         Height          =   315
         Left            =   -64800
         TabIndex        =   5
         Top             =   840
         Width           =   915
      End
      Begin VB.TextBox txtLectura 
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
         Left            =   -70320
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscarLectura 
         Caption         =   "..."
         Height          =   315
         Left            =   -69000
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtNumeroLecturaFiltro 
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
         Left            =   -73620
         TabIndex        =   1
         Top             =   900
         Width           =   3255
      End
      Begin MSDataGridLib.DataGrid grdBuscar 
         Height          =   5475
         Left            =   -74760
         TabIndex        =   9
         Top             =   2160
         Width           =   11715
         _ExtentX        =   20664
         _ExtentY        =   9657
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
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   1
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
      Begin MSFlexGridLib.MSFlexGrid grdPasar 
         Height          =   4395
         Left            =   300
         TabIndex        =   10
         Top             =   2940
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7752
         _Version        =   393216
         Cols            =   7
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
      Begin MSFlexGridLib.MSFlexGrid grdLectura 
         Height          =   4455
         Left            =   240
         TabIndex        =   16
         Top             =   2940
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7858
         _Version        =   393216
         Cols            =   6
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
      Begin VB.Label Label9 
         Caption         =   "Fecha Mayor a "
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
         Left            =   -68280
         TabIndex        =   54
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label7 
         Caption         =   "Lectura"
         Height          =   375
         Left            =   -74400
         TabIndex        =   47
         Top             =   1740
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Referencias"
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
         Left            =   -74400
         TabIndex        =   44
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Label lblPaso 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   2520
         Width           =   9015
      End
      Begin VB.Label Label2 
         Caption         =   "Tareas"
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
         Left            =   300
         TabIndex        =   29
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
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
         Index           =   1
         Left            =   300
         TabIndex        =   22
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Cajaº 
         Caption         =   "Caja"
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
         Left            =   2040
         TabIndex        =   21
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Cliente"
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
         Left            =   3720
         TabIndex        =   20
         Top             =   2160
         Width           =   675
      End
      Begin VB.Label Label6 
         Caption         =   "D Verific."
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
         Left            =   6660
         TabIndex        =   19
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Orden"
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
         Index           =   2
         Left            =   300
         TabIndex        =   18
         Top             =   2100
         Width           =   675
      End
      Begin VB.Label lblOrden 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1440
         TabIndex        =   17
         Top             =   2100
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Lectura"
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
         Left            =   -71100
         TabIndex        =   7
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label8 
         Caption         =   "Filtro"
         Height          =   315
         Left            =   -74340
         TabIndex        =   2
         Top             =   900
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4560
      Top             =   4080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   5940
      Top             =   1500
   End
   Begin MSCommLib.MSComm Comm1 
      Left            =   2640
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      BaudRate        =   19200
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4320
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":096D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":0AC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":0C21
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1073
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":14C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1917
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1D69
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":21BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":260D
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":2A5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":2EB1
            Key             =   "Aceptar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":31CB
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":34E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":37FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":3C51
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":40A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":47F5
            Key             =   "VerMas1"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":4BEF
            Key             =   "VerMenos1"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":4FED
            Key             =   "VerMas3"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":56E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":5DE1
            Key             =   "VerMas4"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":64DB
            Key             =   "VerMas6"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":6BD5
            Key             =   "VerMenos"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":716F
            Key             =   "VerMas"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":7709
            Key             =   "Conextar"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":7E03
            Key             =   "Desconextar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   0
      Left            =   1860
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   60
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":81F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":8469
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":8827
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":8C61
            Key             =   "Borrar1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":9064
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":948B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":9832
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":9BF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":9FB6
            Key             =   "Salvar2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":A232
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":A4B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":A87D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":AC3A
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":AEBB
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":B12C
            Key             =   "Casa"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":B4FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":B8D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":BB54
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":BD25
            Key             =   "Atras2"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":C0DF
            Key             =   "Inicio"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":C1C4
            Key             =   "Fin"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":C2A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":C674
            Key             =   "Adelante2"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":CA28
            Key             =   "Correo2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":CE1B
            Key             =   "Bandera"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":D07E
            Key             =   "trvt2"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":D43F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":DD19
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":DFB9
            Key             =   "Cancelar1"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":E2D3
            Key             =   "Aceptar1"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":E5ED
            Key             =   "trvt"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":E6C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":E799
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":EBC0
            Key             =   "Atras3"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":F2BA
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":F9B4
            Key             =   "Adelante"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":100AE
            Key             =   "Correo3"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":107A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":10EA2
            Key             =   "Correo4"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1123C
            Key             =   "Correo"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":11F16
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":127F0
            Key             =   "Punto"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":130CA
            Key             =   "Cancelar2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":13498
            Key             =   "Aceptar2"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":13851
            Key             =   "Aceptar"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1412B
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":14A05
            Key             =   "Conextar"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":150FF
            Key             =   "Desconextar"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":157F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1620B
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":16C1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":16E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1711C
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1739A
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":17604
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":18016
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":18A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1943A
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":19E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1A0E1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Index           =   1
      Left            =   540
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   50
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1A4FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1A776
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1AB34
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1AF6E
            Key             =   "Borrar1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1B371
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1B798
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1BB3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1BEFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1C2C3
            Key             =   "Salvar2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1C53F
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1C7C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1CB8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1CF47
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1D1C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1D439
            Key             =   "Casa"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1D80C
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1DBE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1DE61
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1E032
            Key             =   "Atras2"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1E3EC
            Key             =   "Inicio"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1E4D1
            Key             =   "Fin"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1E5B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1E981
            Key             =   "Adelante2"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1ED35
            Key             =   "Correo2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1F128
            Key             =   "Bandera"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1F38B
            Key             =   "trvt2"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":1F74C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":20026
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":202C6
            Key             =   "Cancelar1"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":205E0
            Key             =   "Aceptar1"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":208FA
            Key             =   "trvt"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":209D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":20AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":20ECD
            Key             =   "Atras3"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":215C7
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":21CC1
            Key             =   "Adelante"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":223BB
            Key             =   "Correo3"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":22AB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":231AF
            Key             =   "Correo4"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":23549
            Key             =   "Correo"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":24223
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":24AFD
            Key             =   "Punto"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":253D7
            Key             =   "Cancelar2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":257A5
            Key             =   "Aceptar2"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":25B5E
            Key             =   "Aceptar"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":26438
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":26D12
            Key             =   "Conextar"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":2740C
            Key             =   "Desconextar"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":27B06
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lectura de meno .frx":28518
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu crdComms 
      Caption         =   "Comms"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
         Shortcut        =   ^C
      End
   End
End
Attribute VB_Name = "frmLecturaMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CommOpen As Integer
Dim DATO As String
Dim Cliente As Integer
Dim bComConnected As Boolean
Dim Pos As Integer
Private Sub Grabar()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim i As Integer
    Dim NUMERO_LECTURA As Long
    Dim ConLectura As New ADODB.Connection
    Dim cant  As Integer
    Dim CanPersonas As Integer
    Dim Unidades As Double
    Dim L As Integer
    Dim ConPers  As Boolean
    
    ConPers = False
    
    On Error GoTo salir:
    
    ConPers = False
    
     For L = 0 To lstPersonal.ListCount - 1
           If lstPersonal.Selected(L) = True Then
                ConPers = True
            End If
      Next
    
    
    If ConPers = False Then
        MsgBox "Falta el personal Asignado"
        Exit Sub
    End If
    
    
        
       ConLectura.Open strConBasa
       ConLectura.BeginTrans
        
    Sql = " SELECT MAX(NUMERO_LECTURA) AS MAX_NUMERO_LECTURA"
    Sql = Sql & "  From LECTURA_COLECTOR_CUERPO "
        rs.Open Sql, ConActiva, 0, 1
    If Not rs.EOF Then
        NUMERO_LECTURA = CLng(rs!MAX_NUMERO_LECTURA) + 1
        
    End If
    
    Sql = "  INSERT INTO LECTURA_COLECTOR_CUERPO"
    Sql = Sql & "(NUMERO_LECTURA, DESCRIPCION, USUARIO_CREACION,"
    Sql = Sql & " FECHA_CREACION)"
    Sql = Sql & " VALUES (" & NUMERO_LECTURA & ",'" & MDIfrmInicio.StaInicio.Panels(2).Text & " - " & Trim(UCase(txtDescripcion)) & "','" & Usuario & "'," & SysDate & ")"
    ConLectura.Execute Sql
    For i = 1 To grdLectura.Rows - 1
        If grdLectura.TextMatrix(i, 2) <> "" Then
            Sql = " INSERT INTO LECTURACOLECTOR (NUMERO_LECTURA, CAJA, CLIENTE, ORDEN,  TIPO, TIPOREFERENCIA)"
            Sql = vbCrLf & Sql & "  VALUES (" & NUMERO_LECTURA & "," & grdLectura.TextMatrix(i, 2) & ", " & grdLectura.TextMatrix(i, 3) & "," & grdLectura.TextMatrix(i, 1) & "," & Mid(grdLectura.TextMatrix(i, 4), 1, 2) & ",'" & Trim(grdLectura.TextMatrix(i, 6)) & "' )"
            ConLectura.Execute Sql
             If Mid(grdLectura.TextMatrix(i, 4), 1, 2) = "00" Then
                Sql = " Update basasql.dbo.CONTENEDOR"
                Sql = vbCrLf & Sql & " SET UB_PROVISORIA ='" & Trim(UCase(txtDescripcion)) & "'"
                Sql = vbCrLf & Sql & " Where COD_CLIENTE = " & grdLectura.TextMatrix(i, 3)
                Sql = vbCrLf & Sql & " And NRO_CAJA = " & grdLectura.TextMatrix(i, 2)
                ExecutarSql Sql
            End If
            
            
        End If
    Next
    cant = grdLectura.Rows - 1
    For L = 0 To lstPersonal.ListCount - 1
        If lstPersonal.Selected(L) = True Then
            CanPersonas = CanPersonas + 1
        End If
    Next
    
    Unidades = cant / CanPersonas
    Unidades = Format(Unidades, "000,00")
    
    
    
    For L = 0 To lstPersonal.ListCount - 1
    
    If lstPersonal.Selected(L) = True Then
    
    
    Sql = " INSERT INTO LECTURAS_TAREAS"
    Sql = Sql & "(FK_LECTURA, DESCRIPCION, FK_PERSONAL, CANTIDAD)"
    Sql = Sql & " VALUES (" & NUMERO_LECTURA & " ,'" & Trim(cboTareas.Text) & "'," & Mid(lstPersonal.List(L), 1, 3) & ",'" & Unidades & "')"
      ConLectura.Execute Sql
      
    End If
    
        
    
    Next
    
      ConLectura.CommitTrans
      Clipboard.Clear
       Clipboard.SetText NUMERO_LECTURA
      MsgBox "EL NUMERO DE LECTURA ES " & NUMERO_LECTURA & "ESTA COPIADO EN MEMORIA"
    If lblPaso.Caption <> "" Then
    FileSystem.FileCopy lblPaso, Mid(lblPaso.Caption, 1, Len(lblPaso.Caption) - 4) & " personal  " & MDIfrmInicio.StaInicio.Panels(2).Text & "   lectura _ " & NUMERO_LECTURA & ".txt"
    
    Kill lblPaso
    End If
    
    lblPaso.Caption = ""
    CargarCuerpoLectura 0
    TituloGrilla
    cmdLimpiarLista_Click
  
    Exit Sub
salir:
    ConLectura.RollbackTrans
    MsgBox Err.Description
    Err.Clear
    
    
       TituloGrilla
    
    
End Sub

Private Sub Cargar_Grilla()
    Dim rsCliente As ADODB.Recordset
    On Error GoTo salir
    Dim caja1 As String
        TituloGrilla
        If Len(DATO) < 4 Then
        DATO = ""
        Exit Sub
        End If
        DATO = Replace(DATO, vbCrLf, "")
        DATO = Replace(DATO, "", "")
        DATO = Replace(DATO, Chr(0), "")
        DATO = Replace(DATO, "&FIN", "")
        DATO = Replace(DATO, "&00000000", "")
        For i = 1 To Len(DATO)
            caja1 = Mid(DATO, i + 4, 8)
            Cliente = Mid(DATO, i, 4)
            grdLectura.AddItem CStr("Posicion")
            grdLectura.TextMatrix(grdLectura.Rows - 1, 1) = grdLectura.Rows - 1
            If Trim(caja1) = "" Then
            Exit For
            End If
            
            
            If Cliente = 9999 Then
                    grdLectura.TextMatrix(grdLectura.Rows - 1, 2) = CStr(caja1)
                    grdLectura.TextMatrix(grdLectura.Rows - 1, 3) = CStr(Cliente)
                    i = i + 12
            Else
                                   
                        If CLng(caja1) < 100000 Then
                            
                                
                                grdLectura.TextMatrix(grdLectura.Rows - 1, 2) = CStr(caja1)
                                grdLectura.TextMatrix(grdLectura.Rows - 1, 3) = CStr(Cliente)
                                i = i + 12
                         
                        Else
                                    If CLng(caja1) > 730000 Then
                                            If Digito_Verificador(CLng(caja1)) = CLng(Cliente) Then
                                                grdLectura.TextMatrix(grdLectura.Rows - 1, 2) = CStr(caja1)
                                                grdLectura.TextMatrix(grdLectura.Rows - 1, 3) = BuscarCliente(CLng(caja1))
                                                i = i + 12
                                            Else
                                                MsgBox "Error en lectura descartela caja ORDEN:  " & grdLectura.Rows - 1 & " CAJA " & caja1, vbCritical
                                            End If
                                        Else
                                        MsgBox "ERROR EN CAJA"
                                    End If
                              End If
            End If
                Next
        DATO = ""
        
        Exit Sub
salir:
        MsgBox "ERROR EN PROCESO"
End Sub

Private Sub cmdopen_Click()
    mnuOpen_Click
End Sub

Private Sub cmdAceptar_Click()
If txtCajaLectura.Text > 400000 Then
    If txtDigiverificador.Text = "" Then
        MsgBox "Ingrese le digito verificador"
        Exit Sub
    Else
        If Digito_Verificador(txtCajaLectura.Text) = txtDigiverificador.Text And Len(txtCajaLectura.Text) = 6 Then
            Dim rs As New ADODB.Recordset
            Dim Sql  As String
            Sql = " SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA"
            Sql = Sql & " From dbo.Cajas"
            Sql = Sql & "  Where ID_CAJA = " & txtCajaLectura.Text
            rs.Open Sql, ConActiva, 0, 1
            If Not rs.EOF Then
                grdLectura.TextMatrix(lblOrden.Caption, 2) = rs!NRO_CAJA
                 grdLectura.TextMatrix(lblOrden.Caption, 3) = rs!FK_CLIENTE
            Else
            MsgBox "Cajas sin cliente asignado"
               grdLectura.TextMatrix(lblOrden.Caption, 2) = txtCajaLectura.Text
                grdLectura.TextMatrix(lblOrden.Caption, 3) = 0
            
            End If
            
            
            
        Else
            MsgBox "Ingrese Error en caja"
            Exit Sub
        End If
        
            
    End If
Else
 grdLectura.TextMatrix(lblOrden.Caption, 2) = txtCajaLectura.Text
 grdLectura.TextMatrix(lblOrden.Caption, 3) = txtClienteLectura.Text

End If

End Sub

Private Sub cmdAnalisisBasa_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
grdControlError.Visible = False
grdControl.Visible = True

Sql = " SELECT     CONTROLCAJASMIGUEL.NUMEROCONTROL, CONTROLCAJASMIGUEL.ORDEN, CONTROLCAJASMIGUEL.IDCAJA, CONTROLCAJASMIGUEL.CLIENTE,"
 Sql = Sql & "                     CONTROLCAJASMIGUEL.EMPRESA, CAJAS.NRO_CAJA AS CAJA_BASA, CAJAS.FK_CLIENTE AS CLIENTE_BASA, CONTENEDOR.ESTADO, CONTENEDOR.ESTANTERIA,"
  Sql = Sql & " CONTENEDOR.Horizontal , CONTENEDOR.Vertical"
Sql = Sql & " FROM         CONTROLCAJASMIGUEL LEFT OUTER JOIN"
 Sql = Sql & "  CONTENEDOR ON CONTROLCAJASMIGUEL.CLIENTE = CONTENEDOR.COD_CLIENTE AND"
 Sql = Sql & "  CONTROLCAJASMIGUEL.IDCAJA = CONTENEDOR.NRO_CAJA LEFT OUTER JOIN"
   Sql = Sql & "  CAJAS ON CONTROLCAJASMIGUEL.CLIENTE = CAJAS.FK_CLIENTE AND CONTROLCAJASMIGUEL.IDCAJA = CAJAS.NRO_CAJA"
Sql = Sql & " WHERE     (CONTROLCAJASMIGUEL.EMPRESA LIKE N'%BAS%') AND (CONTROLCAJASMIGUEL.NUMEROCONTROL = 17072013111734)"
Sql = Sql & " ORDER BY CAJA_BASA"


rs.Open Sql, strConBasa, 1, 2
    
    
   Set grdControl.DataSource = rs.DataSource
   grdControl.DataMember = rs.DataMember

   grdControl.Refresh
    CopiarDatosGrilla grdControl
End Sub

Private Sub cmdAnalisisCustodia_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
grdControlError.Visible = False
grdControl.Visible = True



Sql = "  SELECT  CONTROLCAJASMIGUEL.NUMEROCONTROL, CONTROLCAJASMIGUEL.ORDEN, CONTROLCAJASMIGUEL.IDCAJA, CONTROLCAJASMIGUEL.CLIENTE,"
Sql = Sql & vbCrLf & " SUBSTRING(CAJAS_DATAS_1707.Ubicacion, 10, 1) AS VER_CUSTO, CONTROLCAJASMIGUEL.EMPRESA, CAJAS.NRO_CAJA AS CAJA_BASA,"
Sql = Sql & vbCrLf & " CAJAS.FK_CLIENTE AS CLIENTE_BASA, CAJAS_DATAS_1707.IDCaja AS CAJA_CUSTODIA, CAJAS_DATAS_1707.IDCliente AS CLIENTE_CUSTODIA,"
Sql = Sql & vbCrLf & " CAJAS_DATAS_1707.Estado AS ESTADO_CUSTODIA, CAJAS_BAJAS_DATA_1707.IDCliente AS CLIENTE_BAJA, CAJAS_BAJAS_DATA_1707.FechaBaja"
Sql = Sql & vbCrLf & " FROM  CONTROLCAJASMIGUEL INNER JOIN"
Sql = Sql & vbCrLf & " CAJAS_BAJAS_DATA_1707 ON CONTROLCAJASMIGUEL.IDCAJA = CAJAS_BAJAS_DATA_1707.IDCaja LEFT OUTER JOIN"
Sql = Sql & vbCrLf & " CAJAS_DATAS_1707 ON CONTROLCAJASMIGUEL.IDCAJA = CAJAS_DATAS_1707.IDCaja LEFT OUTER JOIN"
Sql = Sql & vbCrLf & " CAJAS ON CONTROLCAJASMIGUEL.IDCAJA = CAJAS.ID_CAJA"
Sql = Sql & vbCrLf & " WHERE     (CONTROLCAJASMIGUEL.EMPRESA LIKE N'%CUS%') "
Sql = Sql & vbCrLf & " AND (CONTROLCAJASMIGUEL.NUMEROCONTROL = " & txtOrdenControl.Text & ")"
Sql = Sql & vbCrLf & " ORDER BY CAJA_CUSTODIA, CAJA_BASA"
rs.Open Sql, strConBasa, 1, 2
    
    
   Set grdControl.DataSource = rs.DataSource
   grdControl.DataMember = rs.DataMember

   grdControl.Refresh
    CopiarDatosGrilla grdControl
   



End Sub

Private Sub cmdBuscar_Click()
On Error GoTo salir:
 Dim rs As New ADODB.Recordset
        Dim Sql As String
        rs.CursorLocation = adUseClient
        MousePointer = 11
       
        
        
        If Not IsNumeric(txtCaja.Text) Then
            MsgBox "Ingrese la caja ", vbCritical
            Exit Sub
        End If
        
        Sql = "  SELECT LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA AS LECTURA,"
        Sql = Sql & vbCrLf & "  LECTURA_COLECTOR_CUERPO.FECHA_CREACION AS FECHA,"
        Sql = Sql & vbCrLf & "  LECTURACOLECTOR.CAJA AS CAJA,"
        Sql = Sql & vbCrLf & "  LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN,"
        Sql = Sql & vbCrLf & "  LECTURA_COLECTOR_CUERPO.Descripcion ,LECTURACOLECTOR.TIPO, TIPOREFERENCIA "
        Sql = Sql & vbCrLf & "  FROM LECTURACOLECTOR, LECTURA_COLECTOR_CUERPO"
        Sql = Sql & vbCrLf & "  Where LECTURACOLECTOR.NUMERO_LECTURA = LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA"
        If txtCaja.Text < 70000 Then
            Sql = Sql & vbCrLf & "  AND LECTURACOLECTOR.CLIENTE = " & InputBox("Ingrese el cliente")
        End If
        Sql = Sql & vbCrLf & "  AND LECTURACOLECTOR.CAJA = " & txtCaja.Text
        Sql = Sql & vbCrLf & "  ORDER BY LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA "
        
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        DATOSGRILLA grdBuscar, rs
        MousePointer = 0
        Exit Sub
salir:
        MsgBox "ERROR"
End Sub

Private Sub cmdBuscarLectura_Click()
 Dim rs As New ADODB.Recordset
        Dim Sql As String
        rs.CursorLocation = adUseClient
        MousePointer = 11
        
        If Not IsNumeric(txtLectura.Text) Then
            MsgBox "Ingrese la LECTURA ", vbCritical
            Exit Sub
        End If
       
        
        Sql = "  SELECT NUMERO_LECTURA, CAJA, CLIENTE, ORDEN , LECTURACOLECTOR.TIPO, TIPOREFERENCIA  From LECTURACOLECTOR"
        Sql = Sql & vbCrLf & " Where NUMERO_LECTURA = " & txtLectura.Text
        Sql = Sql & vbCrLf & " ORDER BY ORDEN"
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        DATOSGRILLA grdBuscar, rs
        MousePointer = 0
End Sub


Private Sub cmdCajasSinMarcar_Click()
 Dim rsLectura  As New ADODB.Recordset
    
    On Error GoTo salir:
    Dim Sql As String
        Sql = " SELECT     CAJA, CLIENTE, NUMERO_LECTURA, ORDEN "
        Sql = Sql & " FROM  dbo.LECTURACOLECTOR "
        Sql = Sql & " Where NUMERO_LECTURA = " & InputBox("Ingrese el Numero de Lectura", "", 0)
        Sql = Sql & " ORDER BY ORDEN "
        rsLectura.Open Sql, strConBasa, 0, 1
        Do While Not rsLectura.EOF
            Sql = " Update dbo.Cajas "
            Sql = Sql & " SET FK_TIPO_REFERENCIA = 1090 "
            Sql = Sql & " , FK_LECTURA = " & rsLectura!NUMERO_LECTURA
            Sql = Sql & " , FK_PERSONAL_ASIGNACION_TIPO =" & MDIfrmInicio.StaInicio.Panels(2).Text
            Sql = Sql & " Where FK_CLIENTE = " & rsLectura!Cliente
            Sql = Sql & " AND NRO_CAJA = " & rsLectura!Caja
            Sql = Sql & " AND FK_TIPO_REFERENCIA IS NULL "
            ExecutarSql Sql
            rsLectura.MoveNext
         Loop
         
         MsgBox "La actualizacion se realizo con exito", vbInformation
         Exit Sub
salir:
         MsgBox Err.Description
End Sub

Private Sub cmdContarCliente_Click()
Dim i As Integer
Dim cantidad As Integer


For i = 1 To grdLectura.Rows - 1
   If CLng(txtClienteLectura.Text) = CLng(grdLectura.TextMatrix(i, 3)) Then
           cantidad = cantidad + 1
   
   
   
   End If
   
Next
MsgBox " La cantidad de elementos es " & cantidad
End Sub

Private Sub cmdControlReferencias_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    
    
    

    

Sql = " SELECT     LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.TIPOREFERENCIA, COUNT(*) AS Cantidad"
Sql = Sql & " FROM  LECTURA_COLECTOR_CUERPO INNER JOIN"
Sql = Sql & " LECTURACOLECTOR ON LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = LECTURACOLECTOR.NUMERO_LECTURA"
Sql = Sql & " Where LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = " & txtLectutaTipoReferencia.Text
Sql = Sql & " GROUP BY LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.TIPOREFERENCIA"
Sql = Sql & " Having (LECTURACOLECTOR.Cliente < 9000)"
Sql = Sql & " ORDER BY LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.TIPOREFERENCIA"


rs.Open Sql, ConActiva
        DATOSGRILLA grdTipoReferencia, rs
        MousePointer = 0
End Sub

Private Sub cmdCopiarExcel_Click()
    CopiarDatosGrilla grdBuscar
End Sub

Private Sub cmdFiltro_Click()

End Sub

Private Sub cmdExportarExcel_Click()
CopiarDatosGrilla grdLecturasCuerpo
End Sub

Private Sub cmdFechaMayor_Click()
 Dim rs As New ADODB.Recordset
        Dim Sql As String
        rs.CursorLocation = adUseClient
        MousePointer = 11
        
        If Trim(txtFechaMayor.Text) = "" Then
            MsgBox "Ingrese la LECTURA ", vbCritical
            Exit Sub
        End If
        Sql = " SELECT LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA, LECTURA_COLECTOR_CUERPO.USUARIO_CREACION,"
        Sql = Sql & vbCrLf & " LECTURA_COLECTOR_CUERPO.FECHA_CREACION, LECTURA_COLECTOR_CUERPO.DESCRIPCION, LECTURACOLECTOR.NUMERO_LECTURA AS Expr1,"
        Sql = Sql & vbCrLf & " LECTURACOLECTOR.Caja , LECTURACOLECTOR.Cliente, LECTURACOLECTOR.Orden, LECTURACOLECTOR.TIPO, LECTURACOLECTOR.TipoReferencia"
        Sql = Sql & vbCrLf & " FROM LECTURA_COLECTOR_CUERPO INNER JOIN"
        Sql = Sql & vbCrLf & " LECTURACOLECTOR ON LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = LECTURACOLECTOR.NUMERO_LECTURA"
        Sql = Sql & vbCrLf & " WHERE FECHA_CREACION > '" & txtFechaMayor.Text & "'"
        Sql = Sql & vbCrLf & " ORDER BY LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA, LECTURACOLECTOR.ORDEN"
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        DATOSGRILLA grdBuscar, rs
        MousePointer = 0


End Sub

Private Sub cmdLimpiarLista_Click()
For i = 0 To lstPersonal.ListCount - 1
lstPersonal.Selected(i) = False

Next



End Sub

Private Sub cmdMarcaDigitalizar_Click()
 Dim rsLectura  As New ADODB.Recordset
 
 
On Error GoTo salir:
    Dim Sql As String
        Sql = " SELECT     CAJA, CLIENTE, NUMERO_LECTURA, ORDEN "
        Sql = Sql & " FROM  dbo.LECTURACOLECTOR "
        Sql = Sql & " Where NUMERO_LECTURA = " & InputBox("Ingrese el Numero de Lectura", "", 0)
        Sql = Sql & " ORDER BY ORDEN "
        rsLectura.Open Sql, strConBasa, 0, 1
        Do While Not rsLectura.EOF
            Sql = " Update dbo.Cajas "
            Sql = Sql & " SET FK_TIPO_REFERENCIA = 1060 "
            Sql = Sql & " , FK_LECTURA = " & rsLectura!NUMERO_LECTURA
            Sql = Sql & " , FK_PERSONAL_ASIGNACION_TIPO =" & MDIfrmInicio.StaInicio.Panels(2).Text
            Sql = Sql & " Where FK_CLIENTE = " & rsLectura!Cliente
            Sql = Sql & " AND NRO_CAJA = " & rsLectura!Caja
            Sql = Sql & " AND  FK_TIPO_REFERENCIA <> 1070  "
            
            ExecutarSql Sql
            rsLectura.MoveNext
         Loop
         
         MsgBox "La actualizacion se realizo con exito", vbInformation
         
          Exit Sub
 
salir:
 MsgBox Err.Description
End Sub

Private Sub cmdMarcarCajasReferencias_Click()
Dim rsLectura  As New ADODB.Recordset

On Error GoTo salir:

    Dim Sql As String
        Sql = " SELECT     CAJA, CLIENTE, NUMERO_LECTURA, ORDEN "
        Sql = Sql & " FROM  dbo.LECTURACOLECTOR "
        Sql = Sql & " Where NUMERO_LECTURA = " & InputBox("Ingrese el Numero de Lectura", "", 0)
        Sql = Sql & " ORDER BY ORDEN "
        rsLectura.Open Sql, strConBasa, 0, 1
        Do While Not rsLectura.EOF
            Sql = " Update dbo.Cajas "
            Sql = Sql & " SET FK_TIPO_REFERENCIA = 1000 "
            Sql = Sql & " , FK_LECTURA = " & rsLectura!NUMERO_LECTURA
            Sql = Sql & " , FK_PERSONAL_ASIGNACION_TIPO =" & MDIfrmInicio.StaInicio.Panels(2).Text
            Sql = Sql & " Where FK_CLIENTE = " & rsLectura!Cliente
            Sql = Sql & " And  NRO_CAJA = " & rsLectura!Caja
            ExecutarSql Sql
            rsLectura.MoveNext
         Loop
 MsgBox "La actualizacion se realizo con exito", vbInformation
 Exit Sub
 
salir:
 MsgBox Err.Description
 
 
End Sub

Private Sub cmdMarcaRearchivo_Click()
 Dim rsLectura  As New ADODB.Recordset
 
On Error GoTo salir:
 
    Dim Sql As String
        Sql = " SELECT     CAJA, CLIENTE, NUMERO_LECTURA, ORDEN "
        Sql = Sql & " FROM  dbo.LECTURACOLECTOR "
        Sql = Sql & " Where NUMERO_LECTURA = " & InputBox("Ingrese el Numero de Lectura", "", 0)
        Sql = Sql & " ORDER BY ORDEN "
        rsLectura.Open Sql, strConBasa, 0, 1
        Do While Not rsLectura.EOF
            Sql = " Update dbo.Cajas "
            Sql = Sql & " SET FK_TIPO_REFERENCIA = 1040 "
            Sql = Sql & " , FK_LECTURA = " & rsLectura!NUMERO_LECTURA
            Sql = Sql & " , FK_PERSONAL_ASIGNACION_TIPO =" & MDIfrmInicio.StaInicio.Panels(2).Text
            Sql = Sql & " Where FK_CLIENTE = " & rsLectura!Cliente
            Sql = Sql & " AND NRO_CAJA = " & rsLectura!Caja
            Sql = Sql & " AND  FK_TIPO_REFERENCIA <> 1050  "
            ExecutarSql Sql
            rsLectura.MoveNext
         Loop
         
         MsgBox "La actualizacion se realizo con exito", vbInformation
          Exit Sub
 
salir:
 MsgBox Err.Description
End Sub

Private Sub cmdMarcarLegajos_Click()
    Dim rsLectura  As New ADODB.Recordset
    
    On Error GoTo salir:
    Dim Sql As String
        Sql = " SELECT     CAJA, CLIENTE, NUMERO_LECTURA, ORDEN "
        Sql = Sql & " FROM  dbo.LECTURACOLECTOR "
        Sql = Sql & " Where NUMERO_LECTURA = " & InputBox("Ingrese el Numero de Lectura", "", 0)
        Sql = Sql & " ORDER BY ORDEN "
        rsLectura.Open Sql, strConBasa, 0, 1
        Do While Not rsLectura.EOF
            Sql = " Update dbo.Cajas "
            Sql = Sql & " SET FK_TIPO_REFERENCIA = 1010 "
            Sql = Sql & " , FK_LECTURA = " & rsLectura!NUMERO_LECTURA
            Sql = Sql & " , FK_PERSONAL_ASIGNACION_TIPO =" & MDIfrmInicio.StaInicio.Panels(2).Text
            Sql = Sql & " Where FK_CLIENTE = " & rsLectura!Cliente
            Sql = Sql & " AND NRO_CAJA = " & rsLectura!Caja
            Sql = Sql & " AND (FK_TIPO_REFERENCIA <> 1020 OR FK_TIPO_REFERENCIA IS NULL) "
            ExecutarSql Sql
            rsLectura.MoveNext
         Loop
         
         MsgBox "La actualizacion se realizo con exito", vbInformation
         Exit Sub
salir:
         MsgBox Err.Description
End Sub

Private Sub MemoWinCe()
TituloGrilla
grdPasar.Clear
grdPasar.Rows = 1
Dim cont As Integer
Dim VarTexto As String
cont = 0
Dim Cliente As String
Dim Elemento As String
Dim antElemento As String
Dim Control As String
Dim TIPO As String
Dim TipoReferencia As String
Dim Orden As Integer
On Error GoTo salir
Dim Sql As String
Dim Paso As String
Dim CountError As Integer
CountError = 1
Paso = InputBox("Paso archivo", Paso, "\\222.15.19.251\planta2\Lectura\Lecturas Diarias\")
CommonDialog1.FileName = Paso & "*.txt"
CommonDialog1.ShowOpen
Dim rs As New ADODB.Recordset
Dim FK_CLIENTE  As Integer

Dim ERROR As String

Dim Pos As Integer
inicio:
On Error GoTo salir:
If "*.txt" = CommonDialog1.FileName Then
    Exit Sub
End If
lblPaso.Caption = CommonDialog1.FileName
Open CommonDialog1.FileName For Input As #1
Dim P As Integer
Do Until EOF(1)
    Line Input #1, VarTexto
    If VarTexto <> "" Then
        grdPasar.AddItem VarTexto
        Cliente = grdPasar.TextMatrix(grdPasar.Rows - 1, 3)
        Elemento = grdPasar.TextMatrix(grdPasar.Rows - 1, 2)
        Control = grdPasar.TextMatrix(grdPasar.Rows - 1, 5)
        If grdPasar.TextMatrix(grdPasar.Rows - 1, 1) <> "" Then
            Pos = grdPasar.TextMatrix(grdPasar.Rows - 1, 1)
        Else
            Pos = 0
        End If
        TipoReferencia = grdPasar.TextMatrix(grdPasar.Rows - 1, 6)
        If Trim(Cliente) = "" Then
            MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
            ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
            Elemento = ""
            Cliente = ""
        End If
        TIPO = grdPasar.TextMatrix(grdPasar.Rows - 1, 4)
        If Mid(grdPasar.TextMatrix(grdPasar.Rows - 1, 4), 1, 2) = "C5" Then
            TIPO = "IDDD"
        End If
       Orden = grdLectura.Rows
'       If Elemento = 2174 Then
'            MsgBox "eeee"
'       End If

       
       Select Case TIPO
       Case "ESTA"
            TIPO = "90-Estanteria"

       Case "PERS"
            TIPO = "91-Pesonal"
            For i = 0 To lstPersonal.ListCount - 1
                If Format(Elemento, "000") = Mid(lstPersonal.List(i), 1, 3) Then
                    lstPersonal.Selected(i) = True
                End If
             Next
       
       Case "IDDD"
                Sql = "    SELECT     NRO_CAJA, FK_CLIENTE, ID_CAJA, DIGITO_VERIFICADOR"
                Sql = Sql & " From Cajas "
                Sql = Sql & " Where  ID_CAJA = " & Elemento
                Set rs = New ADODB.Recordset
                rs.Open Sql, ConActiva, 0, 1
                If rs.EOF Then
                   MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
                       ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                    Elemento = ""
                    Cliente = ""
                Else
                    If IsNull(rs!FK_CLIENTE) Then
                        MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
                        
                      ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                       Elemento = ""
                        Cliente = ""
                    Else
                        Cliente = rs!FK_CLIENTE
                        Elemento = rs!NRO_CAJA
                        TIPO = "00-Caja"
                    End If
                End If

        Case "CUST", "BASA"
            Set rs = New ADODB.Recordset
            Sql = "    SELECT     NRO_CAJA, FK_CLIENTE, ID_CAJA, DIGITO_VERIFICADOR"
            Sql = Sql & " From Cajas "
            
           If CLng(Elemento) > 100000 Then
                Sql = Sql & " Where NRO_CAJA = " & Elemento
            Else
             If Mid(Control, 1, 2) = "C6" Then
                Sql = Sql & " Where ID_CAJA = " & Elemento
             Else
                
                Sql = Sql & " Where FK_CLIENTE = " & Cliente
                Sql = Sql & " and NRO_CAJA = " & Elemento
              End If
              
            End If
            
            rs.Open Sql, strConBasa
            If rs.EOF Then
                 If CInt(Cliente) = 39 Then
                            If MsgBox("Quiere crear la caja " & Elemento & " para el cliente 39", vbYesNo) = vbYes Then
                                 CrearCajas CLng(Elemento), CInt(Cliente)
'                            Else
'                             ERROR = ERROR & vbCrLf & " CAJA " & Caja & " CLIENTE: " & Cliente
                            End If


                       Else
                
                        MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
                       ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                        Elemento = ""
                        Cliente = ""
                End If
            Else
                If IsNull(rs!FK_CLIENTE) Then
                    If IsNull(rs!NRO_CAJA) Then
                        MsgBox "El cliente y el elemento es nulo"
                    Else
                        MsgBox "El cliente es nulo para el elemento : " & rs!NRO_CAJA
                        Elemento = ""
                        Cliente = ""
                    End If
                Else
                    Cliente = rs!FK_CLIENTE
                    Elemento = rs!NRO_CAJA
                    TIPO = "00-Caja"
                End If
                
            End If
            
            
            

        Case "VBAS"
            Set rs = New ADODB.Recordset
            Sql = "    SELECT     NRO_CAJA, FK_CLIENTE, ID_CAJA, DIGITO_VERIFICADOR"
            Sql = Sql & " From Cajas "
            
            If CLng(Elemento) > 100000 Then
                Sql = Sql & " Where NRO_CAJA = " & Elemento
            Else
                Sql = Sql & " Where FK_CLIENTE = " & Cliente
                Sql = Sql & " and NRO_CAJA = " & Elemento
            End If
            
            rs.Open Sql, strConBasa
            If rs.EOF Then
                
                MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
               ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                Elemento = ""
                Cliente = ""
            Else
                If CLng(Cliente) = rs!FK_CLIENTE Or CLng(Cliente) = rs!Digito_Verificador Then
                   If IsNull(rs!FK_CLIENTE) Then
                        MsgBox "El cliente no existe para la  caja " & rs!NRO_CAJA
                        ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                        Elemento = ""
                        Cliente = ""
                   Else
                        Cliente = rs!FK_CLIENTE
                        Elemento = rs!NRO_CAJA
                        TIPO = "00-Caja"
                    End If
                Else
                    MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
                    ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                    Elemento = ""
                    Cliente = ""
                End If
            End If
            
       Case "VCUS"
            Set rs = New ADODB.Recordset
            Sql = "    SELECT     NRO_CAJA, FK_CLIENTE, ID_CAJA, DIGITO_VERIFICADOR"
            Sql = Sql & " From Cajas "
            Sql = Sql & " Where ID_CAJA = " & Elemento
            rs.Open Sql, strConBasa
            If rs.EOF Then
                MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
                ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                Elemento = ""
                Cliente = ""
            Else
                If Not IsNull(rs!FK_CLIENTE) Then
                    If CLng(Cliente) = rs!FK_CLIENTE Or CLng(Cliente) = rs!Digito_Verificador Then
                        Cliente = rs!FK_CLIENTE
                        Elemento = rs!NRO_CAJA
                        TIPO = "00-Caja"
                    Else
                        MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
                        ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                        Elemento = ""
                        Cliente = ""
                    End If
                Else
                    MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
                    ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                    Elemento = ""
                    Cliente = ""
                End If
                    
                End If
            
            
    Case "LEG2", "VLEG"
            Set rs = New ADODB.Recordset
            
            
            Sql = "    SELECT     ID_CLIENTE_LEGAJO, ID_LEGAJO, COD_CLIENTE"
            Sql = Sql & " From LEGAJOS"

            If CLng(Elemento) > 217267 Then
                Sql = Sql & " Where ID_LEGAJO  = " & Elemento
            Else
                Sql = Sql & " Where COD_CLIENTE = " & Cliente
                Sql = Sql & " and ID_CLIENTE_LEGAJO = " & Elemento
            End If
            
            
            rs.Open Sql, strConBasa
            If rs.EOF Then
                
                MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
               ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
                Elemento = ""
                Cliente = ""
            Else
            If IsNull(rs!COD_CLIENTE) Then
                MsgBox "El elemento no existe " & Elemento
                 GoTo salir
            End If
                Cliente = rs!COD_CLIENTE
                Elemento = rs!ID_CLIENTE_LEGAJO
                TIPO = "03-Legajo"
            End If
         
         Case "VLIB", "LIBR"
            Set rs = New ADODB.Recordset
            
            Sql = "   SELECT     NRO_LIBRO_INTERNO, COD_CLIENTE"
            Sql = Sql & " From basasql.dbo.LIBROS"
            Sql = Sql & " Where COD_CLIENTE = " & Cliente
            Sql = Sql & " And NRO_LIBRO_INTERNO = " & Elemento
            
            
            
            
            rs.Open Sql, strConBasa
            If rs.EOF Then
                
                MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " Caja " & Elemento
               ERROR = ERROR & vbTab & CLng(Elemento) & vbTab & Pos & vbCrLf
               
                Elemento = ""
                Cliente = ""
            Else
                Cliente = rs!COD_CLIENTE
                Elemento = rs!NRO_LIBRO_INTERNO
                TIPO = "01-Libro"
            End If
            
        
        End Select
    End If
final:



Control = grdPasar.TextMatrix(grdPasar.Rows - 1, 5)

    If Elemento <> "" And antElemento <> Elemento Then
        grdLectura.AddItem vbTab & Orden & vbTab & Elemento & vbTab & Cliente & vbTab & TIPO & vbTab & Control & vbTab & TipoReferencia
        antElemento = Elemento
    End If
ErrorProximo:
    
'    If Elemento <> "" Then
'        grdLectura.AddItem vbTab & Orden & vbTab & Elemento & vbTab & Cliente & vbTab & TIPO & vbTab & Control & vbTab & TipoReferencia
'        antElemento = Elemento
'    End If
Loop
Close #1
If ERROR <> "" Then
Clipboard.Clear
Clipboard.SetText ERROR
MsgBox "Los errores fueron copiados"


End If



Exit Sub
salir:
Close #1
  If Err.Number = 55 Then
      CountError = CountError + 1
      If CountError > 5 Then
          MsgBox Err.Description
          Else
      GoTo inicio
      End If
  End If
MsgBox Err.Description
  
  
  
'        If Cliente < 8999 Then
'                If Caja > 100000 Then
'                            Set rs = New ADODB.Recordset
'                        Sql = "    SELECT     NRO_CAJA, FK_CLIENTE, ID_CAJA, DIGITO_VERIFICADOR"
'                        Sql = Sql & " From Cajas "
'                        Sql = Sql & " Where ID_CAJA = " & Caja
'                        rs.Open Sql, strConBasa
'                        If Not rs.EOF Then
'                            If IsNull(rs!FK_CLIENTE) Then
'                                FK_CLIENTE = 0
'                            Else
'                                FK_CLIENTE = rs!FK_CLIENTE
'                            End If
'
'                            If (rs!Digito_Verificador = Cliente) Or (BuscarDigitoVerificadorCajas(CStr(Caja)) = Cliente) Then
'                                grdLectura.AddItem vbTab & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & vbTab & Caja & vbTab & FK_CLIENTE & vbTab & vbTab & "OK"
'                            Else
'                                grdLectura.AddItem vbTab & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & vbTab & Caja & vbTab & FK_CLIENTE & vbTab & vbTab & "No Verificado"
'                            End If
'                        Else
'                                MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " CAJA " & Caja
'                        End If
'                Else
'
'                        Set rs = New ADODB.Recordset
'                        Sql = "    SELECT     NRO_CAJA, FK_CLIENTE, ID_CAJA, DIGITO_VERIFICADOR"
'                        Sql = Sql & " From Cajas "
'                        If Mid(VarTexto, 20, 4) = "VBAS" Or Mid(VarTexto, 20, 4) = "BASA" Then
'                            Sql = Sql & " Where FK_CLIENTE = " & Cliente
'                            Sql = Sql & " AND NRO_CAJA = " & Caja
'                        Else
'                            If Cliente > 7 Then
'                                If Mid(VarTexto, 20, 4) = "VCUS" Or Mid(VarTexto, 20, 4) = "CUST" Then
'                                    Sql = Sql & " Where FK_CLIENTE = " & Cliente
'                                    Sql = Sql & " AND NRO_CAJA = " & Caja
'                                 End If
'                            Else
'                                Sql = Sql & " Where DIGITO_VERIFICADOR = " & Cliente
'                                Sql = Sql & " AND NRO_CAJA = " & Caja
'                            End If
'                        End If
'                        rs.Open Sql, strConBasa
'                         If Not rs.EOF Then
'                         If IsNull(rs!FK_CLIENTE) Then
'                             MsgBox "Error pos  : " & grdPasar.TextMatrix(grdPasar.Rows - 1, 1)
'                              grdLectura.AddItem vbTab & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & vbTab & Caja & vbTab & rs!FK_CLIENTE & vbTab & vbTab & "NO"
'                         Else
'
'                             grdLectura.AddItem vbTab & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & vbTab & Caja & vbTab & rs!FK_CLIENTE & vbTab & vbTab & "OK"
'                             End If
'
'                        Else
'
'
'                       If CInt(Cliente) = 39 Then
'                            If MsgBox("Quiere crear la caja " & Caja & " para el cliente 39", vbYesNo) = vbYes Then
'                                 CrearCajas CLng(Caja), CInt(Cliente)
'                            Else
'                             ERROR = ERROR & vbCrLf & " CAJA " & Caja & " CLIENTE: " & Cliente
'                            End If
'
'
'                       Else
'
'                       MsgBox "Error en lectura  pos:" & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & " CAJA " & Caja & " CLIENTE: " & Cliente & " EMPRESA:" & Mid(VarTexto, 20, 4)
'                    End If
'
'                        End If
'                End If
'    Else
'            If Cliente = 9999 Then
'             grdLectura.AddItem vbTab & grdPasar.TextMatrix(grdPasar.Rows - 1, 1) & vbTab & Caja & vbTab & Cliente & vbTab & vbTab & "Estanteria"
'            End If
'
'
'
'            If Cliente = 9000 Then
'             For i = 0 To lstPersonal.ListCount - 1
'
'                If Format(Caja, "000") = Mid(lstPersonal.List(i), 1, 3) Then
'
'                    lstPersonal.Selected(i) = True
'                End If
'
'
'             Next
'
'
'            End If
'
'
'        End If
'    Else
'
'    End If
'
'
'Final:
'Loop
'
'  lblPaso.Caption = CommonDialog1.FileName
'If ERROR <> "" Then
'    MsgBox "lOS ERRORES SERAN COPIADOS"
'    Clipboard.Clear
'    Clipboard.SetText ERROR
'    End If
'
'Close #1
'
'Exit Sub
'salir:
'Close #1
'    If Err.Number = 55 Then
'        CountError = CountError + 1
'        If CountError > 5 Then
'            MsgBox Err.Description
'            Else
'        GoTo inicio
'        End If
'    End If
'MsgBox Err.Description
End Sub

Private Sub cmdMarcarCajas_Click()
    Dim Lectura As Long
        Lectura = InputBox("Ingreso la Lectura")
        Dim Sql As String
        Dim rs As New ADODB.Recordset
        
        Sql = " SELECT LECTURACOLECTOR.ID, LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, "
        Sql = Sql & vbCrLf & " LECTURACOLECTOR.Orden , LECTURACOLECTOR.TIPO, LECTURACOLECTOR.TipoReferencia, CAJAS.ID_CAJA "
        Sql = Sql & vbCrLf & " FROM LECTURACOLECTOR INNER JOIN "
        Sql = Sql & vbCrLf & " CAJAS ON LECTURACOLECTOR.CLIENTE = CAJAS.FK_CLIENTE AND LECTURACOLECTOR.CAJA = CAJAS.NRO_CAJA "
        Sql = Sql & vbCrLf & " Where LECTURACOLECTOR.NUMERO_LECTURA = " & Lectura
        Sql = Sql & vbCrLf & " ORDER BY LECTURACOLECTOR.ID "
        rs.Open Sql, strConBasa
        
        If Mid(cboTipoReferencia.Text, 1, 4) = 1060 Then
            MarcarCajasDigital (Lectura)
         End If
        
        Do While Not rs.EOF
             Sql = " Update basasql.dbo.CAJAS "
            Sql = Sql & vbCrLf & " SET FK_TIPO_REFERENCIA = " & Mid(cboTipoReferencia.Text, 1, 4)
            Sql = Sql & vbCrLf & ", FK_TIPO_REFERENCIA_PERSONAL = " & MDIfrmInicio.StaInicio.Panels(2).Text
            Sql = Sql & vbCrLf & ", TIPO_REFERENCIA_FECHA = " & SysDate
            Sql = Sql & vbCrLf & " Where ID_CAJA = " & rs!ID_CAJA
            ExecutarSql Sql
            rs.MoveNext
        Loop
        
        MsgBox "Terminado"
 End Sub

Private Sub cmdMarcarTodos_Click()
Dim i As Integer

For i = 0 To lstPersonal.ListCount - 1
    lstPersonal.Selected(i) = True
Next


End Sub

Private Sub cmdSubirArchivo_Click()
grdControlError.Visible = True
grdControl.Visible = False
grdPasar.Clear
grdPasar.Rows = 1
Dim cont As Integer
Dim VarTexto As String
cont = 0
Dim Cliente As Long
Dim Caja As Long
On Error GoTo salir
Dim Sql As String
Dim Paso As String
Dim CountError As Integer
CountError = 1
Paso = InputBox("Paso archivo", Paso, "Z:\Planta\LECTURAS\LECTURAS PENDIENTES\")
CommonDialog1.FileName = Paso & "*.txt"
CommonDialog1.ShowOpen
Dim rs As New ADODB.Recordset
Dim FK_CLIENTE  As Integer
Dim ERROR As String
inicio:
Open CommonDialog1.FileName For Input As #1
Dim P As Integer
Do Until EOF(1)
    Line Input #1, VarTexto
    If VarTexto <> "" Then
        grdControlError.AddItem VarTexto
       
    End If
final:
Loop
If ERROR <> "" Then
    MsgBox "lOS ERRORES SERAN COPIADOS"
    Clipboard.Clear
    Clipboard.SetText ERROR
    End If

Close #1

Exit Sub
salir:
Close #1
    If Err.Number = 55 Then
        CountError = CountError + 1
        If CountError > 5 Then
            MsgBox Err.Description
            Else
        GoTo inicio
        End If
    End If
MsgBox Err.Description



End Sub

Private Sub Comm1_OnComm()
    Select Case Comm1.CommEvent
      Case comBreak
            Debug.Print "Se ha recibido una interrupción"
      Case comEventFrame
            Debug.Print "Error de trama"
      Case comEventOverrun
            Debug.Print "Datos perdidos"
      Case comEventRxOver
            Debug.Print "Desbordamiento del búfer de recepción."
      Case comEventRxParity
            Debug.Print "Error de paridad."
      Case comEventTxFull
            Debug.Print "Búfer de transmisión lleno."
      Case comEventDCB
            Debug.Print "Error inesperado al recuperar DCB"
      Case comEvCD
            Debug.Print "Cambio en la línea CD"
      Case comEvCTS
            Debug.Print "Cambio en la línea CTS."
            If DATO <> "" Then
                Cargar_Grilla
            End If
      Case comEvDSR
            Debug.Print "Cambio en la línea DSR."
      Case comEvRing
            Debug.Print "Cambio en el indicador de llamadas."
      Case comEvReceive
            DATO = DATO & Comm1.Input
            Debug.Print "comEvReceive"
      Case comEvSend
            Debug.Print "Hay un SThreshold caracteres en el búfer de transmisión."
      Case comEvEOF
            Debug.Print "Se ha encontrado un carácter EOF en la entrada."
   End Select

End Sub

Private Sub Command1_Click()
Dim rsLectura  As New ADODB.Recordset

On Error GoTo salir:

    Dim Sql As String
        Sql = " SELECT     CAJA, CLIENTE, NUMERO_LECTURA, ORDEN "
        Sql = Sql & " FROM  dbo.LECTURACOLECTOR "
        Sql = Sql & " Where NUMERO_LECTURA = " & InputBox("Ingrese el Numero de Lectura", "", 0)
        Sql = Sql & " ORDER BY ORDEN "
        rsLectura.Open Sql, strConBasa, 0, 1
        Do While Not rsLectura.EOF
            Sql = " Update dbo.Cajas "
            Sql = Sql & " SET FK_TIPO_REFERENCIA = NULL "
            Sql = Sql & " , FK_LECTURA = " & rsLectura!NUMERO_LECTURA
            Sql = Sql & " Where FK_CLIENTE = " & rsLectura!Cliente
            Sql = Sql & " And  NRO_CAJA = " & rsLectura!Caja
            ExecutarSql Sql
            rsLectura.MoveNext
         Loop
 MsgBox "La actualizacion se realizo con exito", vbInformation
 Exit Sub
 
salir:
 MsgBox Err.Description
End Sub

Private Sub Command2_Click()
On Error GoTo SettingsFail
    Comm1.settings = Text2.Text
    Exit Sub
SettingsFail:
    MsgBox ERROR$
    Exit Sub
End Sub



Private Sub Command4_Click()
    Label4 = "EV_RECEIVE"
    Comm1.InputMode = comInputModeText
    Text3 = Comm1.Input
End Sub

Private Sub Command6_Click()
    With Combo
    .AddItem "flexSortNone" ' 0
    .AddItem "flexSortGenericAscending" '1
    .AddItem "flexSortGenericDescending" '2
    .AddItem "flexSortNumericAscending" '3
    .AddItem "flexSortNumericDescending" '4
    .AddItem "flexSortStringNoCaseAsending" '5
    .AddItem "flexSortNoCaseDescending" '6
    .AddItem "flexSortStringAscending" '7
    .AddItem "flexSortStringDescending" '8
    .ListIndex = 0
End With


End Sub

Private Sub Command5_Click()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Sql As String
Dim i As Integer
con.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Z:\Tareas\Migracion  de P&L\Administracion\Proyecto\MIGRA.mdb"

rs.Open " SELECT Empresa_1Id , CajaId FROM CAJASCLIENTES", con

Do While Not rs.EOF
   MsgBox rs!Empresa_1Id
    rs.MoveNext
Loop



End Sub


Public Sub ReadBarCode_Click()
    Dim nRC As Long
    Dim nNumOfBarcodes As Long
    Dim i As Long
    Dim j As Long
        
    Dim arrbyteBarcode(99) As Byte '100 elements
    Dim nBytesRead As Long
    Dim bstrBarcode As String
    Dim bstrTmp As String * 50
    
    On Error GoTo salir
    
    
    If Toolbar1.Buttons(5).Enabled = True Then
    MsgBox "El puesto no esta Abierto"
    Exit Sub
    End If
                   
    
    Pos = 0
    'Determine if we can read the data
    nRC = csp2ReadData
        
    If nRC > 0 Then
        'cs1504 has barcodes!
        nNumOfBarcodes = nRC
        
       
            
'        WriteToLog "Reading " & CStr(nRC) & _
'            " Barcodes at " & Time
            
        'Check to see that we are in ascii mode...
        If csp2GetASCIIMode = PARAM_ON Then
        
'            DisplayInBCWindow "ASCII Mode ON"
            
            For i = 0 To (nNumOfBarcodes - 1)
                nBytesRead = csp2GetPacket(arrbyteBarcode(0), i, 100)
                
                If nBytesRead > 0 Then
                    'bstrBarcode = "Rcvd: "
                             
                    'Display the Barcode type
                    nRC = csp2GetCodeType(arrbyteBarcode(1), bstrTmp, Len(bstrTmp))
                    
'                    DisplayInBCWindow bstrTmp
                    bstrBarcode = " "
                                                       
                    ' display the barcode is ascii
                    ' skip the length, type, .... timestamp
                    For j = 2 To (nBytesRead - 5)
                        bstrBarcode = bstrBarcode & Chr(arrbyteBarcode(j))
                        'DisplayInBCWindow Chr(arrbyteBarcode(j))
                    Next j
                    
                    'Display the timestamp
                    nRC = csp2TimeStamp2Str(arrbyteBarcode(nBytesRead - 4), bstrTmp, Len(bstrTmp))
                    LectutaLlavero (Trim(bstrBarcode))
                End If
            Next i
        Else
            'Add binary mode packets handling here..
'            DisplayInBCWindow "Binary Mode ON"
        End If
        
    Else
'        DisplayInBCWindow "No Barcodes to Read."
    End If
    Exit Sub
salir:
   MsgBox Err.Description & " " & bstrBarcode
    
End Sub

Private Sub Form_Load()
        
        
        
        CommOpen = False
        imgEstado.Picture = ImageList2.ListImages.Item("Desconextar").Picture
        TituloGrilla
        CargarCuerpoLectura 0
       
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(6).Enabled = False
        
        
        
        
       
        
        cboTipo.ListIndex = 0
        
        
        
         
          
          
         
         
         lstPersonal.Clear
         
         If MDIfrmInicio.StaInicio.Panels(2) = 37 Or MDIfrmInicio.StaInicio.Panels(2) = 12 Or MDIfrmInicio.StaInicio.Panels(2) = 31 Or MDIfrmInicio.StaInicio.Panels(2) = 48 Or MDIfrmInicio.StaInicio.Panels(2) = 47 Or MDIfrmInicio.StaInicio.Panels(2) = 19 Or MDIfrmInicio.StaInicio.Panels(2) = 17 Or MDIfrmInicio.StaInicio.Panels(2) = 38 Then
            lstPersonal.Enabled = True
          End If
          
         
         Dim rs As New ADODB.Recordset
         Dim Sql As String
         
         lstPersonal.Enabled = False
         
         Sql = " SELECT NAVES, IDPERSONAL, NOMBRE, APELLIDO, ACTIVO "
         Sql = Sql & " From PERSONAL "
         Sql = Sql & " Where (NAVES = 1)"
         Sql = Sql & " ORDER BY APELLIDO, NOMBRE"
         
         If MDIfrmInicio.StaInicio.Panels(2) = 19 Then
            lstPersonal.Enabled = True
         End If
         
         rs.Open Sql, strConBasa
         Do While Not rs.EOF
            lstPersonal.AddItem Format(rs!idPersonal, "000") & " " & Trim(rs!Apellido) & " " & Trim(rs!Nombre)
            rs.MoveNext
         Loop
         
         If MDIfrmInicio.StaInicio.Panels(2) = 37 Or MDIfrmInicio.StaInicio.Panels(2) = 12 Or MDIfrmInicio.StaInicio.Panels(2) = 31 Or MDIfrmInicio.StaInicio.Panels(2) = 48 Or MDIfrmInicio.StaInicio.Panels(2) = 47 Or MDIfrmInicio.StaInicio.Panels(2) = 19 Or MDIfrmInicio.StaInicio.Panels(2) = 17 Or MDIfrmInicio.StaInicio.Panels(2) = 38 Then
            lstPersonal.Enabled = True
          End If
         
         CargarTipoReferencias
         
'          Select Case ButtonMenu.Text
'   Case "Envio_Alsina_D"
'   Case "Envio_Basa_D"
'   Case "Ingreso_Alsina_D"
'   Case "Ingreso_Alsina_D"
'   End Select
         
        

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Toolbar1.Buttons(5).Enabled = False Then
Cancel = 1
MsgBox "Cerrar el puerto"
End If
                   
End Sub

Private Sub grdBuscar_DblClick()
txtNumeroLecturaFiltro = grdBuscar.Text
txtNumeroLecturaFiltro_KeyPress 13
End Sub

Private Sub grdLectura_Click()
lblOrden.Caption = grdLectura.Row
txtCajaLectura.Text = grdLectura.TextMatrix(grdLectura.Row, 2)
txtClienteLectura.Text = grdLectura.TextMatrix(grdLectura.Row, 3)
End Sub

Private Sub grdLectura_DblClick()
'    Select Case grdLectura.TextMatrix(0, grdLectura.Col)
'    Case "Orden"
'        grdLectura.Sort = 3
'    Case "Caja"
'        grdLectura.Sort = 3
'    Case "Cliente"
'        grdLectura.Sort = 3
'    Case "Razon Social"
'        grdLectura.Sort = 7
'    End Select
    
End Sub

Private Sub GUARDAR_Click()
Dim i As Integer
Dim Sql As String
Dim NUMEROCONTROL As Double

NUMEROCONTROL = Replace(date, "/", "") & Replace(Time, ":", "")
For i = 0 To grdControlError.Rows - 1


If grdControlError.TextMatrix(i, 3) <> "" Then
   If grdControlError.TextMatrix(i, 3) < 9000 Then
   
  Sql = "  INSERT INTO CONTROLCAJASMIGUEL"
  Sql = Sql & "                     (NUMEROCONTROL, ORDEN, IDCAJA, CLIENTE, EMPRESA, CODIGO)"
Sql = Sql & " VALUES     (      " & NUMEROCONTROL
        Sql = Sql & " , " & grdControlError.TextMatrix(i, 1)
        Sql = Sql & " , " & grdControlError.TextMatrix(i, 2)
        Sql = Sql & " , " & grdControlError.TextMatrix(i, 3)
        Sql = Sql & " , '" & grdControlError.TextMatrix(i, 4) & "'"
        Sql = Sql & " , '" & grdControlError.TextMatrix(i, 5) & "'"
         Sql = Sql & ")"
  
        ExecutarSql Sql
     End If
     End If
Next
txtOrdenControl.Text = NUMEROCONTROL
MsgBox "terminado orden" & NUMEROCONTROL
 End Sub

Private Sub LecturaLlavero_Click()
ReadBarCode_Click
End Sub

Private Sub mnuClose_Click()
On Error GoTo CloseFail
    Comm1.PortOpen = False
    Text2.Enabled = True
    CommOpen = False
    Label7.Caption = "Closed"
    Exit Sub
CloseFail:
    MsgBox ERROR$
    Exit Sub
End Sub

Private Sub mnuOpen_Click()
On Error GoTo OpenFail
    Comm1.CommPort = Combo1.Text
    Comm1.PortOpen = True
    Label7.Caption = "Open"
    CommOpen = True
    
    Exit Sub
OpenFail:
    MsgBox ERROR$
    Exit Sub
End Sub

Private Sub Timer1_Timer()
Dim Valor As String
Timer1.Enabled = False
Label4 = "EV_RECEIVE"
Valor = Comm1.Input
Valor = Mid(Valor, 1, Len(Valor) - 1)
Debug.Print Valor
If Mid(Valor, 1, 3) <> "FIN" Then
    DATO = DATO & Valor
Else
    If MsgBox("TERMINO LA TRANSMICION" & vbCrLf & "Usted desea Insertar los datos", vbYesNo) = vbYes Then
        Cargar_Grilla
    End If


End If

 Debug.Print DATO
  Debug.Print " DATO  "
   Debug.Print " DATO  "
End Sub

Public Sub TituloGrilla()
    grdLectura.Clear
    grdLectura.Cols = 7
    grdLectura.Rows = 1
    grdLectura.ColWidth(0) = 100
    grdLectura.ColWidth(1) = 800
    grdLectura.ColWidth(2) = 1500
    grdLectura.ColWidth(3) = 1000
    grdLectura.ColWidth(4) = 1500
    grdLectura.ColWidth(5) = 2000
    grdLectura.ColWidth(6) = 2500
    
    grdLectura.ColAlignment(0) = 1
    grdLectura.ColAlignment(1) = 1
    grdLectura.ColAlignment(2) = 1
    grdLectura.ColAlignment(3) = 1
    grdLectura.ColAlignment(4) = 1
    grdLectura.ColAlignment(5) = 1
    grdLectura.ColAlignment(6) = 1

    grdLectura.TextMatrix(0, 1) = "Orden"
    grdLectura.TextMatrix(0, 2) = "Elemento"
    grdLectura.TextMatrix(0, 3) = "Cliente"
    grdLectura.TextMatrix(0, 4) = "Tipo"
    grdLectura.TextMatrix(0, 5) = "Control"
    grdLectura.TextMatrix(0, 6) = "Tipo Referencia"
    
    lblOrden.Caption = ""
    txtCajaLectura.Text = ""
    txtClienteLectura.Text = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Dim rs As New ADODB.Recordset
 Dim Deposito As String
 
 Select Case Button.Key
 Case "plaReferencia"
  PlanillaReferncia
 Case "CopiarMemo"
 CopiarServer

 
 Case "Entrada"
    If MDIfrmInicio.StaInicio.Panels(2) = 19 Or MDIfrmInicio.StaInicio.Panels(2) = 48 Then
           EntradaElementos
     Else
           MsgBox "Solo el Usuario 19 puede realizar entradas ", vbCritical
     End If
  
 Case "AROG"
    MemoWinCe
 Case "Borrar"
 
    grdLectura.RemoveItem (grdLectura.Row)
 Case "Cancelar"
    TituloGrilla
 Case "Conextar"
            On Error GoTo CloseFail1
            If optColector.value = True Then
                Comm1.PortOpen = True
                CommOpen = True
                Toolbar1.Buttons(5).Enabled = False
                 Toolbar1.Buttons(6).Enabled = True
                imgEstado.Picture = ImageList2.ListImages.Item("Conextar").Picture
            End If
            If optLlavero.value = True Then
               Dim dwRC As Long
                dwRC = csp2Init(0)
                If dwRC <> STATUS_OK Then
                   imgEstado.Picture = ImageList2.ListImages.Item("Desconextar").Picture
                   Toolbar1.Buttons(5).Enabled = True
                Else
                   
                   Toolbar1.Buttons(5).Enabled = False
                   Toolbar1.Buttons(6).Enabled = True
                   imgEstado.Picture = ImageList2.ListImages.Item("Conextar").Picture
                    Rem SelectCOMcb.Enabled = False
                    bComConnected = True
                End If
'
'                'Start the CTS State Timer
'                Rem CTSStateTimer.Enabled = True
'
'               Rem  UpdateMainFormCtrls
'
''                     GetTime.Enabled = True
''        SetTime.Enabled = True
''        SetDefaults.Enabled = True
''        PowerDown.Enabled = True
''        WakeUp.Enabled = True
'
'        ReadBarCode.Enabled = True
'        AutoDownloadEnabled.Enabled = True
'        AutoDownloadApply.Enabled = True
'
'        tbDeviceState.Text = "Connected"
'
            
            
            End If
            Exit Sub
CloseFail1:
            MsgBox ERROR$
            Exit Sub
 Case "Desconextar"
        On Error GoTo CloseFail
        
            If optColector.value = True Then
                    Comm1.PortOpen = False
                    CommOpen = False
                   Toolbar1.Buttons(5).Enabled = True
                   Toolbar1.Buttons(6).Enabled = False
                    imgEstado.Picture = ImageList2.ListImages.Item("Desconextar").Picture
            End If
                    
            If optLlavero.value = True Then
                'Close down the com ports...
                                   Toolbar1.Buttons(5).Enabled = True
                   Toolbar1.Buttons(6).Enabled = False
    csp2Restore
    imgEstado.Picture = ImageList2.ListImages.Item("Desconextar").Picture
   
    Rem SelectCOMcb.Enabled = True
    Rem UpdateMainFormCtrls
    
    'Stop the CTS State Timer
    Rem CTSStateTimer.Enabled = False
    
    Rem WriteToLog "COM Port Closed"
            
            End If
            
        
        Exit Sub
CloseFail:
        MsgBox ERROR$
        Exit Sub
Case "Aceptar"
        Grabar
Case "Buscar"
    Dim rsLectura As New ADODB.Recordset
    rsLectura.CursorLocation = adUseClient
    Dim DATO As String
        DATO = InputBox("Ingrese la cadena de busqueda", "", "")
   
    rsLectura.Open ("SELECT * FROM LECTURA_COLECTOR_CUERPO  where DESCRIPCION like '%" & DATO & "%'   ORDER BY NUMERO_LECTURA DESC"), ConActiva, 0, 1
    
    Set grdLecturasCuerpo.DataSource = rsLectura.DataSource
        grdLecturasCuerpo.DataMember = rsLectura.DataMember
Case "Pedro"
InsertarLecturaPedro
Case "Simular"
    LeerDatos
 End Select
 MousePointer = 0
 
End Sub

Private Sub LeerDatos()

Dim L As String
Dim i As Long
Dim DATO As String
Dim datoInicio As String
Dim espacio As Integer
Dim comienzo As Integer
Dim Cliente As Integer
On Error GoTo salir

L = Clipboard.GetText
L = Trim(L)
comienzo = 1
espacio = 1
L = Replace(L, vbCrLf, "&")
Dim inicio As Long
Dim Fin As Integer
Dim postab As Integer
Dim Caja As Long
inicio = 1
Fin = 1
Dim C As Integer
Dim Orden As Integer
Dim TIPO As String
Dim TipoDato As String

TIPO = InputBox("Ingrese el tipo 0 - cajas , 1 - libros , 3 - legajos ")

For i = 1 To Len(L)
    If Mid(L, i, 1) = "&" Then
        DATO = Mid(L, inicio, i - inicio)
      
       postab = InStr(1, DATO, vbTab)
        
        Cliente = Mid(DATO, 1, postab - 1)
        Caja = Mid(DATO, postab)
        
        Select Case TIPO
        Case 0
            TipoDato = "00-Caja"
        Case 1
            TipoDato = "01-Libro"
        Case 3
            TipoDato = "03-Legajo"
        End Select
        
        inicio = i + 1
        Orden = Orden + 1
        grdLectura.AddItem vbTab & Orden & vbTab & Caja & vbTab & Cliente & vbTab & TipoDato & vbTab & "Virtual"
   End If
    
    
'
'    datoInicio = Mid(L, comienzo)
'    espacio = InStr(datoInicio, "&")
'    dato = Mid(datoInicio, 1, espacio - 1)
'    comienzo = espacio + 1
    Rem CargarGrilla CStr(dato)
Next
DATO = Mid(L, inicio, i - inicio)


Exit Sub

salir:
MsgBox Err.Description
   
End Sub


Public Sub CargarCuerpoLectura(Filtro As String)
On Error GoTo salir:

    Dim rsLectura As New ADODB.Recordset
    rsLectura.CursorLocation = adUseClient
    Dim Sql As String
    
        Sql = " SELECT        TOP (100) LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA, LECTURA_COLECTOR_CUERPO.USUARIO_CREACION,"
        Sql = Sql & " LECTURA_COLECTOR_CUERPO.FECHA_CREACION, LECTURA_COLECTOR_CUERPO.DESCRIPCION, COUNT(*) AS CANTIDAD"
        Sql = Sql & "  FROM            LECTURA_COLECTOR_CUERPO INNER JOIN"
        Sql = Sql & "          LECTURACOLECTOR ON LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = LECTURACOLECTOR.NUMERO_LECTURA"
        Sql = Sql & "   GROUP BY LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA, LECTURA_COLECTOR_CUERPO.USUARIO_CREACION, LECTURA_COLECTOR_CUERPO.FECHA_CREACION,"
        Sql = Sql & "           LECTURA_COLECTOR_CUERPO.Descripcion , LECTURACOLECTOR.NUMERO_LECTURA"

    
    If IsNumeric(Filtro) Then
          If Filtro = 0 Then
            rsLectura.Open (Sql & " ORDER BY FECHA_CREACION DESC, USUARIO_CREACION"), ConActiva, 0, 1
          Else
            rsLectura.Open (Sql & "  WHERE USUARIO_CREACION = '" & Filtro & "' ORDER BY FECHA_CREACION DESC, USUARIO_CREACION"), ConActiva, 0, 1
        End If
    Else
        If IsDate(Filtro) Then
            rsLectura.Open (Sql & " HAVING FECHA_CREACION ='" & Filtro & "'  ORDER BY FECHA_CREACION DESC, USUARIO_CREACION"), ConActiva, 0, 1
        Else
            rsLectura.Open Sql & " HAVING LECTURA_COLECTOR_CUERPO.DESCRIPCION LIKE '%" & Filtro & "%'  ORDER BY FECHA_CREACION DESC, USUARIO_CREACION", ConActiva, 0, 1
        End If
        
    End If
    
    Set grdLecturasCuerpo.DataSource = rsLectura.DataSource
        grdLecturasCuerpo.DataMember = rsLectura.DataMember
Exit Sub
salir:
MsgBox Err.Description
End Sub

Public Sub CargarLibros()
    Dim rsCliente As ADODB.Recordset
    Dim caja1 As String
        Debug.Print DATO
        Exit Sub
        TituloGrilla
        If Len(DATO) < 4 Then
            DATO = ""
            Exit Sub
        End If
            DATO = Replace(DATO, vbCrLf, "")
            Cliente = Mid(DATO, 1, 2)
            caja1 = Mid(DATO, 3, 5)
            On Error GoTo MalTomado
            
        For i = 9 To Len(DATO)
                    If Mid(DATO, i, 1) = "@" Then
                        'Registro = Mid(Dato, Comenzar, 12)
                        'Posicion = Mid(Dato, 1, 3)
                        caja1 = Mid(DATO, i + 4, 5)
                        If (Mid(DATO, i + 1, 3)) <> "" Then
                            Cliente = Mid(DATO, i + 1, 3)
                        Else
                            Cliente = 0
                        End If
                        If CStr(Cliente) <> 0 Then
                            grdLectura.AddItem CStr(Posicion)
                            grdLectura.TextMatrix(grdLectura.Rows - 1, 1) = grdLectura.Rows - 1
                            grdLectura.TextMatrix(grdLectura.Rows - 1, 2) = CStr(caja1)
                            grdLectura.TextMatrix(grdLectura.Rows - 1, 3) = CStr(Cliente)
                            Set rsCliente = New ADODB.Recordset
                            rsCliente.Open ("Select * from clientes where id_cliente = " & CStr(Cliente)), ConActiva, 0, 1
                            If Not rsCliente.EOF Then
                              grdLectura.TextMatrix(grdLectura.Rows - 1, 4) = Trim(UCase(rsCliente.Fields("Razon_Social").value))
                            End If
                        End If
                        Comenzar = i + 1
                    End If
                Next
        Set rsCliente = Nothing
        DATO = ""
        Exit Sub
MalTomado:
        MsgBox "Realice la operacion Nuevamente"
        DATO = ""
End Sub

Public Sub LectutaLlavero(DATO As String)
Dim rs As New ADODB.Recordset
Dim Sql As String
grdLectura.AddItem CStr(Pos)
Pos = Pos + 1
Select Case UCase(Mid(DATO, 1, 2))
Case "C1"

            grdLectura.TextMatrix(grdLectura.Rows - 1, 1) = Pos
            grdLectura.TextMatrix(grdLectura.Rows - 1, 2) = Mid(DATO, 6)
            grdLectura.TextMatrix(grdLectura.Rows - 1, 3) = Mid(DATO, 3, 3)
Case "C5"
Sql = " SELECT     ID_CAJA, DIGITO_VERIFICADOR, FK_CLIENTE"
Sql = Sql & "  From dbo.Cajas"
Sql = Sql & "  Where ID_CAJA =" & Mid(DATO, 3, 7)

rs.Open Sql, ConActiva, 0, 1

If rs.EOF Then
    MsgBox "ERROR LECTURA"
Else
    If CLng(rs!Digito_Verificador) = CLng(Mid(DATO, 10)) Then
        grdLectura.TextMatrix(grdLectura.Rows - 1, 1) = Pos
            grdLectura.TextMatrix(grdLectura.Rows - 1, 2) = rs!ID_CAJA
            If Not IsNull(rs!FK_CLIENTE) Then
                grdLectura.TextMatrix(grdLectura.Rows - 1, 3) = rs!FK_CLIENTE
            Else
                 grdLectura.TextMatrix(grdLectura.Rows - 1, 3) = 0
            End If
            
        
    End If
End If

Case "ER"
            grdLectura.TextMatrix(grdLectura.Rows - 1, 1) = Pos
            grdLectura.TextMatrix(grdLectura.Rows - 1, 2) = Mid(DATO, 4)
            grdLectura.TextMatrix(grdLectura.Rows - 1, 3) = 0

End Select

End Sub

Public Function BuscarCliente(FK_CAJAS As Long) As Integer
    Dim rs As New ADODB.Recordset
    Dim Sql As String

Sql = " SELECT  FK_CLIENTE From dbo.Cajas Where ID_CAJA = " & FK_CAJAS

rs.Open Sql, ConActiva, 0, 1
    If rs.EOF Then
        BuscarCliente = 0
    Else
        If Not IsNull(rs!FK_CLIENTE) Then
            BuscarCliente = rs!FK_CLIENTE
        Else
            BuscarCliente = 0
        End If
        
    End If
    

End Function

Public Sub InsertarLecturaPedro()
Dim Sql As String
 Dim Sql2 As String
 Dim RsPedro As New ADODB.Recordset
 Dim MaxLectBasa As New ADODB.Recordset
 Dim ConPedro As New ADODB.Connection
 Dim ConBasa As New ADODB.Connection
 
 Dim MAXLECBASA As Double
 
    MaxLectBasa.Open "SELECT MAX(NUMERO_LECTURA) AS MaxLectura FROM dbo.LECTURA_COLECTOR_CUERPO ", ConActiva, 0, 1
    MAXLECBASA = MaxLectBasa!MaxLectura
    ConPedro.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Serverbasa1\SistemasBasa\ClientesBases\cambio.mdb;Persist Security Info=False"

Sql = " SELECT LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA, LECTURA_COLECTOR_CUERPO.DESCRIPCION, LECTURA_COLECTOR_CUERPO.FECHA_CREACION, LECTURA_COLECTOR_CUERPO.USUARIO_CREACION, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN"
Sql = Sql & "  FROM LECTURA_COLECTOR_CUERPO INNER JOIN LECTURACOLECTOR ON LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = LECTURACOLECTOR.NUMERO_LECTURA"
Sql = Sql & "  Where LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA > " & MAXLECBASA
Sql = Sql & "  ORDER BY LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA, LECTURACOLECTOR.ORDEN;"

RsPedro.Open Sql, ConPedro
 ConBasa.Open strConBasa
 Dim Lectura As Double
 
 Do While Not RsPedro.EOF
        If Lectura <> RsPedro!NUMERO_LECTURA Then
            Sql = "INSERT INTO dbo.LECTURA_COLECTOR_CUERPO"
            Sql = Sql & " (NUMERO_LECTURA, USUARIO_CREACION, FECHA_CREACION, DESCRIPCION)"
            Sql = Sql & "  VALUES (" & RsPedro!NUMERO_LECTURA & ",99,'" & RsPedro!FECHA_CREACION & "','" & Trim(RsPedro!Descripcion) & "')"
            ExecutarSql Sql
            Lectura = RsPedro!NUMERO_LECTURA
        End If
        
        
        Sql = " INSERT INTO dbo.LECTURACOLECTOR"
        Sql = Sql & " (NUMERO_LECTURA, CAJA, CLIENTE, ORDEN)"
        Sql = Sql & "  VALUES (" & RsPedro!NUMERO_LECTURA & "," & RsPedro!Caja & "," & RsPedro!Cliente & "," & RsPedro!Orden & ")"
        ExecutarSql Sql
     RsPedro.MoveNext
 Loop
 
 MsgBox "Exportacion Terminada"
End Sub

Private Sub txtNumeroLecturaFiltro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtNumeroLecturaFiltro.Text) = "" Then
              txtNumeroLecturaFiltro.Text = 0
        End If
    
        CargarCuerpoLectura txtNumeroLecturaFiltro.Text
        txtNumeroLecturaFiltro.Text = 0
    End If
End Sub

Public Sub PlanillaReferncia()
    Dim Usuario As String
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim CAJA_1, CAJA_2, CAJA_3, CAJA_4, CAJA_5, CAJA_6, CAJA_7 As String
    Dim DIG_1, DIG_2 As String
    Dim PER_1, PER_2, PER_3 As String
    Dim Personal As String
    Dim CLI_1, CLI_2, CLI_3 As String
  MousePointer = 11
    On Error GoTo salir
    
    
'    Sql = " SELECT     LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN,"
'    Sql = Sql & " Cajas.Digito_Verificador"
'    Sql = Sql & "  FROM         LECTURACOLECTOR INNER JOIN"
'    Sql = Sql & "  CAJAS ON LECTURACOLECTOR.CAJA = CAJAS.NRO_CAJA AND LECTURACOLECTOR.CLIENTE = CAJAS.FK_CLIENTE"
'    Sql = Sql & "  Where LECTURACOLECTOR.NUMERO_LECTURA = " & InputBox("Ingrese el  numero de lectura", , 0)
'    Sql = Sql & "  ORDER BY LECTURACOLECTOR.ORDEN DESC"


Sql = "  SELECT     NUMERO_LECTURA, CAJA, CLIENTE, ORDEN"
Sql = Sql & "  From LECTURACOLECTOR"
Sql = Sql & "  Where NUMERO_LECTURA = " & InputBox("Ingrese el  numero de lectura", , 0)
Sql = Sql & "  ORDER BY ORDEN"

    rs.Open Sql, strConBasa


ExecutarSql "DELETE FROM IMPRESION_REFERENCIA"
Personal = InputBox("INGRESE EL PERSONAL")

Do While Not rs.EOF
CAJA_1 = "'" & Mid(rs!Caja, Len(rs!Caja), 1) & "'"
If Len(rs!Caja) > 1 Then
    CAJA_2 = "'" & Mid(rs!Caja, Len(rs!Caja) - 1, 1) & "'"
Else
     CAJA_2 = "'0'"
End If

If Len(rs!Caja) > 2 Then
    CAJA_3 = "'" & Mid(rs!Caja, Len(rs!Caja) - 2, 1) & "'"
Else
    CAJA_3 = "'0'"
End If


If Len(rs!Caja) > 3 Then
    CAJA_4 = "'" & Mid(rs!Caja, Len(rs!Caja) - 3, 1) & "'"
Else
     CAJA_4 = "'0'"
End If


If Len(rs!Caja) > 4 Then
    CAJA_5 = Mid(rs!Caja, Len(rs!Caja) - 4, 1)
Else
     CAJA_5 = "'0'"
End If


If Len(rs!Caja) > 5 Then
    CAJA_6 = Mid(rs!Caja, Len(rs!Caja) - 5, 1)
    Else
    CAJA_6 = "'0'"
End If
If Len(rs!Caja) > 6 Then

    CAJA_7 = Mid(rs!Caja, Len(rs!Caja) - 6, 1)
Else
    CAJA_7 = "'0'"
End If

Rem DIG_1 = "'" & Mid(rs!Digito_Verificador, Len(rs!Digito_Verificador), 1) & "'"
Rem DIG_2 = "'" & Mid(rs!Digito_Verificador, Len(rs!Digito_Verificador) - 1, 1) & "'"

DIG_1 = "'0'"
DIG_2 = "'0'"



PER_1 = "'" & Mid(Personal, Len(Personal), 1) & "'"

If Len(Personal) > 1 Then
    PER_2 = "'" & Mid(Personal, Len(Personal) - 1, 1) & "'"
Else
    PER_2 = "'0'"
End If

If Len(Personal) > 2 Then
    PER_3 = "'" & Mid(Personal, Len(Personal) - 2, 1) & "'"
Else
    PER_3 = "'0'"
End If




CLI_1 = "'" & Mid(rs!Cliente, Len(rs!Cliente), 1) & "'"

If Len(rs!Cliente) > 1 Then
    CLI_2 = "'" & Mid(rs!Cliente, Len(rs!Cliente) - 1, 1) & "'"
Else
    CLI_2 = "'0'"
End If

If Len(rs!Cliente) > 2 Then
    CLI_3 = "'" & Mid(rs!Cliente, Len(rs!Cliente) - 2, 1) & "'"
Else
    CLI_3 = "'0'"
End If




Sql = " INSERT INTO IMPRESION_REFERENCIA"
Sql = Sql & " (CAJA_1, CAJA_2, CAJA_3, CAJA_4, CAJA_5, CAJA_6, CAJA_7, DIG_1, DIG_2, PER_1, PER_2, PER_3, CLI_1, CLI_2, CLI_3, ORDEN)"
Sql = Sql & "  VALUES  "
Sql = Sql & "(" & CAJA_1 & "," & CAJA_2 & "," & CAJA_3 & "," & CAJA_4 & "," & CAJA_5 & "," & CAJA_6 & "," & CAJA_7 & "," & DIG_1 & "," & DIG_2 & "," & PER_1 & "," & PER_2 & "," & PER_3 & "," & CLI_1 & "," & CLI_2 & "," & CLI_3 & "," & rs!Orden & ")"
ExecutarSql Sql
rs.MoveNext
Loop



  Sql = " SELECT * "
 Sql = Sql & "  FROM  IMPRESION_REFERENCIA"
 Sql = Sql & " order by ORDEN  "




frmReportes.ImprimirReporte PasoReportes & "ReferenciaManual.rpt", Sql, True
MousePointer = 0


salir:

End Sub

Public Sub CopiarServer()


Dim sArchivo As String
Dim Paso As String
Paso = "\\PCTELEMEMO\Memo\"

sArchivo = FileSystem.Dir(Paso & "*.txt")
Do While sArchivo <> ""

Select Case UCase(Mid(sArchivo, 1, 2))
Case "00"
    FileSystem.FileCopy Paso & sArchivo, "Z:\Planta\LECTURAS\REMITO ENTRADA\" & sArchivo
Case "CP"
    FileSystem.FileCopy Paso & sArchivo, "Z:\Planta\LECTURAS\LECTURAS PENDIENTES\Cambio de posición\" & sArchivo
Case "RO"
    FileSystem.FileCopy Paso & sArchivo, "Z:\Planta\LECTURAS\LECTURAS PENDIENTES\Rotulos - RO\" & sArchivo
Case "EP"
     FileSystem.FileCopy Paso & sArchivo, "Z:\Planta\LECTURAS\LECTURAS PENDIENTES\Entrada a Planta - EP\" & sArchivo
Case "EA"
    FileSystem.FileCopy Paso & sArchivo, "Z:\Planta\LECTURAS\LECTURAS PENDIENTES\Entrada a Planta - EP\" & sArchivo
    
Case Else
    FileSystem.FileCopy Paso & sArchivo, "Z:\Planta\LECTURAS\LECTURAS PENDIENTES\" & sArchivo & ".txt"
End Select
    Kill Paso & sArchivo
    sArchivo = Dir()
Loop

MsgBox "Terminado"


End Sub

Public Sub ActualizarTipoReferencia(Cliente As Long, Caja As Long, TipoReferencia As String)
Dim rs As New ADODB.Recordset
Dim Sql As String

TipoReferencia = Mid(TipoReferencia, 1, 4)

Select Case Mid(TipoReferencia, 1, 4)
Case 1001, 1002, 1003, 1004, 1010, 1040, 1060, 1090
    Sql = " SELECT ESTADO "
    Sql = Sql & " From basasql.dbo.CONTENEDOR "
    Sql = Sql & " Where COD_CLIENTE = " & Cliente
    Sql = Sql & " AND NRO_CAJA = " & Caja
    Sql = Sql & " AND ESTADO = 5 "
    rs.Open Sql, strConBasa
    If Not rs.EOF Then
        Sql = " Update basasql.dbo.CAJAS"
        Sql = Sql & " SET FK_TIPO_REFERENCIA =" & TipoReferencia
        Sql = Sql & " Where FK_CLIENTE = " & Cliente
        Sql = Sql & " And NRO_CAJA = " & Caja
        ExecutarSql Sql
    
    End If
    
    
End Select

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

Public Sub EntradaElementos()
    Dim NumeroLectura As String
    Dim Sql As String
    Dim rsLectura As New ADODB.Recordset
    Dim FechaEntrada As String
    Dim Personal As String
    Dim TipoElemento As Integer
    Dim TipoReferencias As Integer
    NumeroLectura = InputBox("Ingrese el numero de lectura", "Entrada")
    On Error GoTo salir
     
        Sql = " SELECT ID, NUMERO_LECTURA, CAJA, CLIENTE, ORDEN, TIPO, TIPOREFERENCIA"
        Sql = Sql & " From basasql.dbo.LECTURACOLECTOR"
        Sql = Sql & " Where (Cliente < 8999)"
        Sql = Sql & " AND NUMERO_LECTURA = " & NumeroLectura
        Sql = Sql & " ORDER BY ID DESC"
           
           
        Sql = "  SELECT LECTURACOLECTOR.ID, LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN, LECTURACOLECTOR.TIPO,"
        Sql = Sql & " LECTURACOLECTOR.TipoReferencia , LECTURA_COLECTOR_CUERPO.UTILIZADA"
        Sql = Sql & " FROM LECTURA_COLECTOR_CUERPO INNER JOIN"
        Sql = Sql & " LECTURACOLECTOR ON LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = LECTURACOLECTOR.NUMERO_LECTURA"
        Sql = Sql & " Where LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = " & NumeroLectura
        Sql = Sql & " And (LECTURACOLECTOR.Cliente < 8999) "
        Sql = Sql & " And (LECTURA_COLECTOR_CUERPO.UTILIZADA Is Null)"
        Sql = Sql & " ORDER BY LECTURACOLECTOR.ID DESC"
         
         rsLectura.Open Sql, strConBasa
         FechaEntrada = SysDate
         Personal = MDIfrmInicio.StaInicio.Panels.Item(2).Text
         If rsLectura.EOF Then
            MsgBox "Lectura utilizada"
            Exit Sub
         End If
         
         
         Do While Not rsLectura.EOF
            TipoElemento = CInt(Mid(rsLectura!TIPO, 1, 2))
             
            If rsLectura!TipoReferencia = "" Then
                TipoReferencias = "0000"
            Else
                TipoReferencias = CInt(Mid(rsLectura!TipoReferencia, 1, 4))
            End If
            
            Sql = " INSERT INTO ENTRADA("
            Sql = Sql & " ELEMENTO, COD_CLIENTE "
            Sql = Sql & " , TIPO, FECHA"
            Sql = Sql & " , COD_PERSONAL, COD_ESTADO)"
            Sql = Sql & " Values ("
            Sql = Sql & rsLectura!Caja & "," & rsLectura!Cliente
            Sql = Sql & " ," & TipoElemento & "," & FechaEntrada
            Sql = Sql & " ," & Personal & ",0)"
            ExecutarSql Sql
            If TipoElemento = 0 Then
                 MarcarTipoReferencia rsLectura!Cliente, rsLectura!Caja, TipoReferencias, True
            End If
            rsLectura.MoveNext
          Loop

            Sql = " Update LECTURA_COLECTOR_CUERPO"
            Sql = Sql & " SET UTILIZADA = '" & Now & "'"
            Sql = Sql & "  Where NUMERO_LECTURA = " & NumeroLectura
            Sql = Sql & "  And (UTILIZADA Is Null)"
            ExecutarSql Sql
            MsgBox "Terminado"
salir:
 
End Sub

Public Function MarcarCajasDigital(Lectura As Long)
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
'    Sql = " SELECT LECTURA_COLECTOR_CUERPO.FECHA_CREACION, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN,"
'    Sql = Sql & " LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA"
'    Sql = Sql & " FROM LECTURACOLECTOR INNER JOIN "
'    Sql = Sql & " LECTURA_COLECTOR_CUERPO ON LECTURACOLECTOR.NUMERO_LECTURA = LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA"
'    Sql = Sql & " Where LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = " & Lectura
'    Sql = Sql & " And LECTURACOLECTOR.Cliente <> 9999 "
'    Sql = Sql & " ORDER BY LECTURACOLECTOR.ORDEN"
    
     Sql = " SELECT LECTURA_COLECTOR_CUERPO.FECHA_CREACION, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN,"
    Sql = Sql & " LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA"
    Sql = Sql & " FROM LECTURACOLECTOR INNER JOIN "
    Sql = Sql & " LECTURA_COLECTOR_CUERPO ON LECTURACOLECTOR.NUMERO_LECTURA = LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA"
    Sql = Sql & " Where LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA  IN (" & Lectura & ")"
    Sql = Sql & " And LECTURACOLECTOR.Cliente <> 9999 "
    Sql = Sql & " ORDER BY LECTURACOLECTOR.ORDEN"
    

    rs.Open Sql, strConBasa

    Do While Not rs.EOF
        Sql = " SELECT   FK_CLIENTES, FK_CAJAS, REMITO, FECHA_INGRESO "
        Sql = Sql & " From DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & "  Where FK_CLIENTES = " & rs!Cliente
        Sql = Sql & "  And FK_CAJAS = " & rs!Caja
        Set rs2 = New ADODB.Recordset
        rs2.Open Sql, strConBasa
        If Not rs2.EOF Then
            Sql = " UPDATE  DOCUMENTOS_DIGITALES_LOTE"
            Sql = Sql & "   SET  FECHA_INGRESO ='" & rs!FECHA_CREACION & "'"
            Sql = Sql & "   Where FK_CLIENTES = " & rs!Cliente
            Sql = Sql & "   And  FK_CAJAS = " & rs!Caja
            ExecutarSql Sql
        
        Else
        
            Sql = "  Insert INTO DOCUMENTOS_DIGITALES_LOTE ("
            Sql = Sql & vbCrLf & " SUB_LOTE"
            Sql = Sql & vbCrLf & " , Descripcion"
            Sql = Sql & vbCrLf & " , FK_CLIENTES"
            Sql = Sql & vbCrLf & " , FK_INDICES"
            Sql = Sql & vbCrLf & " , FK_CAJAS"
            Sql = Sql & vbCrLf & " , FK_ESTADO"
            Sql = Sql & vbCrLf & " , FK_PERSONAL_PREPARACION"
            Sql = Sql & vbCrLf & " , FK_PERSONAL_SCANNER"
            Sql = Sql & vbCrLf & " , FK_PERSONAL_INDEXACION"
            Sql = Sql & vbCrLf & " , FK_PERSONAL_REORDENAR"
            Sql = Sql & vbCrLf & " , REMITO"
            Sql = Sql & vbCrLf & " , CANTIDAD_IMAGENES"
            Sql = Sql & vbCrLf & " , CANTIDAD_ARCHIVOS"
            Sql = Sql & vbCrLf & " , FECHA_PREPARACION"
            Sql = Sql & vbCrLf & " , FECHA_SCANNER"
            Sql = Sql & vbCrLf & " , FECHA_INDEXACION"
            Sql = Sql & vbCrLf & " , FECHA_REORDENAR"
            Sql = Sql & vbCrLf & " , LOTE_ESTADO"
            Sql = Sql & vbCrLf & " , FECHA_INGRESO"
            Sql = Sql & vbCrLf & " )"
            Sql = Sql & vbCrLf & "  VALUES ( 1 "
            Sql = Sql & vbCrLf & " , 'INGRESO CAJA'"
            Sql = Sql & vbCrLf & " ," & rs!Cliente
            Sql = Sql & vbCrLf & " , 10833"
            Sql = Sql & vbCrLf & " , " & rs!Caja
            Sql = Sql & vbCrLf & " , 0"
            Sql = Sql & vbCrLf & " , 99"
            Sql = Sql & vbCrLf & " , 99"
            Sql = Sql & vbCrLf & " , 99"
            Sql = Sql & vbCrLf & " , 99"
            Sql = Sql & vbCrLf & " , '0001-0000000'"
            Sql = Sql & vbCrLf & " , 0"
            Sql = Sql & vbCrLf & " , 0"
            Sql = Sql & vbCrLf & " ,'" & rs!FECHA_CREACION & "'"
            Sql = Sql & vbCrLf & " ,'" & rs!FECHA_CREACION & "'"
            Sql = Sql & vbCrLf & " ,'" & rs!FECHA_CREACION & "'"
            Sql = Sql & vbCrLf & " ,'" & rs!FECHA_CREACION & "'"
            Sql = Sql & vbCrLf & " ,'CREADO'"
            Sql = Sql & vbCrLf & " ,'" & rs!FECHA_CREACION & "'"
            Sql = Sql & vbCrLf & " )"
            ExecutarSql Sql
        
        
        End If
        
       
       
       
    
       rs.MoveNext
    Loop
    


End Function
