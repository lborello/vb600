VERSION 5.00
Object = "{E1A6B8A3-3603-101C-AC6E-040224009C02}#1.0#0"; "IMGTHUMB.OCX"
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#1.0#0"; "IMGEDIT.OCX"
Object = "{009541A3-3B81-101C-92F3-040224009C02}#1.0#0"; "IMGADMIN.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmPlanillaC 
   Caption         =   "Image Edit"
   ClientHeight    =   14310
   ClientLeft      =   1560
   ClientTop       =   1770
   ClientWidth     =   18435
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   14310
   ScaleWidth      =   18435
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   12375
      Left            =   11280
      TabIndex        =   9
      Top             =   120
      Width           =   6975
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3600
         TabIndex        =   95
         Top             =   2400
         Width           =   615
         Begin VB.OptionButton optBorradoNO 
            Caption         =   "NO"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optBorradosi 
            Caption         =   "SI"
            Height          =   255
            Left            =   120
            TabIndex        =   96
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.CheckBox NoestaDB 
         Caption         =   "No esta BD"
         Height          =   255
         Left            =   2280
         TabIndex        =   94
         Top             =   840
         Width           =   1215
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   2280
         TabIndex        =   91
         Top             =   1560
         Width           =   855
         Begin VB.OptionButton LetraCaja 
            Caption         =   "SI"
            Height          =   375
            Left            =   120
            TabIndex        =   93
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton NO 
            Caption         =   "NO"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdRegistro 
         Caption         =   "Registro"
         Height          =   315
         Left            =   5880
         TabIndex        =   90
         Top             =   10560
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Height          =   2355
         Left            =   300
         TabIndex        =   74
         Top             =   7740
         Width           =   6435
         Begin VB.TextBox txtApellidoNombre 
            Height          =   375
            Index           =   2
            Left            =   180
            TabIndex        =   85
            Top             =   480
            Width           =   6135
         End
         Begin VB.CheckBox chkPrenda 
            Caption         =   "Prenda"
            Height          =   275
            Index           =   2
            Left            =   420
            TabIndex        =   84
            Top             =   1320
            Width           =   1035
         End
         Begin VB.CheckBox chkHipoteca 
            Caption         =   "Hipoteca"
            Height          =   275
            Index           =   2
            Left            =   420
            TabIndex        =   83
            Top             =   1620
            Width           =   1035
         End
         Begin VB.CheckBox chkFianza 
            Caption         =   "Fianza"
            Height          =   275
            Index           =   2
            Left            =   420
            TabIndex        =   82
            Top             =   1920
            Width           =   1035
         End
         Begin VB.CheckBox chkAval 
            Caption         =   "Aval"
            Height          =   275
            Index           =   2
            Left            =   1440
            TabIndex        =   81
            Top             =   1320
            Width           =   1035
         End
         Begin VB.CheckBox chkOtra 
            Caption         =   "Otra"
            Height          =   275
            Index           =   2
            Left            =   1440
            TabIndex        =   80
            Top             =   1620
            Width           =   1035
         End
         Begin VB.TextBox txtNumero 
            Height          =   315
            Index           =   2
            Left            =   2340
            TabIndex        =   79
            Top             =   1380
            Width           =   2775
         End
         Begin VB.OptionButton optDNI 
            Caption         =   "DNI"
            Height          =   275
            Index           =   2
            Left            =   5340
            TabIndex        =   78
            Top             =   1320
            Width           =   735
         End
         Begin VB.OptionButton OPTCI 
            Caption         =   "C.I."
            Height          =   275
            Index           =   2
            Left            =   5340
            TabIndex        =   77
            Top             =   1560
            Width           =   735
         End
         Begin VB.OptionButton OPTLC 
            Caption         =   "L.C."
            Height          =   275
            Index           =   2
            Left            =   5340
            TabIndex        =   76
            Top             =   1800
            Width           =   735
         End
         Begin VB.OptionButton OPTLE 
            Caption         =   "L.E."
            Height          =   275
            Index           =   2
            Left            =   5340
            TabIndex        =   75
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo"
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
            Index           =   2
            Left            =   300
            TabIndex        =   89
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Numero"
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
            Index           =   2
            Left            =   2340
            TabIndex        =   88
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo"
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
            Index           =   2
            Left            =   5160
            TabIndex        =   87
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Apellido y Nombre"
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
            Index           =   2
            Left            =   240
            TabIndex        =   86
            Top             =   180
            Width           =   6135
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2355
         Left            =   300
         TabIndex        =   58
         Top             =   5280
         Width           =   6435
         Begin VB.TextBox txtApellidoNombre 
            Height          =   375
            Index           =   1
            Left            =   180
            TabIndex        =   69
            Top             =   480
            Width           =   6135
         End
         Begin VB.CheckBox chkPrenda 
            Caption         =   "Prenda"
            Height          =   275
            Index           =   1
            Left            =   420
            TabIndex        =   68
            Top             =   1320
            Width           =   1035
         End
         Begin VB.CheckBox chkHipoteca 
            Caption         =   "Hipoteca"
            Height          =   275
            Index           =   1
            Left            =   420
            TabIndex        =   67
            Top             =   1620
            Width           =   1035
         End
         Begin VB.CheckBox chkFianza 
            Caption         =   "Fianza"
            Height          =   275
            Index           =   1
            Left            =   420
            TabIndex        =   66
            Top             =   1920
            Width           =   1035
         End
         Begin VB.CheckBox chkAval 
            Caption         =   "Aval"
            Height          =   275
            Index           =   1
            Left            =   1440
            TabIndex        =   65
            Top             =   1320
            Width           =   1035
         End
         Begin VB.CheckBox chkOtra 
            Caption         =   "Otra"
            Height          =   275
            Index           =   1
            Left            =   1440
            TabIndex        =   64
            Top             =   1620
            Width           =   1035
         End
         Begin VB.TextBox txtNumero 
            Height          =   315
            Index           =   1
            Left            =   2340
            TabIndex        =   63
            Top             =   1380
            Width           =   2655
         End
         Begin VB.OptionButton optDNI 
            Caption         =   "DNI"
            Height          =   275
            Index           =   1
            Left            =   5340
            TabIndex        =   62
            Top             =   1320
            Width           =   735
         End
         Begin VB.OptionButton OPTCI 
            Caption         =   "C.I."
            Height          =   275
            Index           =   1
            Left            =   5340
            TabIndex        =   61
            Top             =   1560
            Width           =   735
         End
         Begin VB.OptionButton OPTLC 
            Caption         =   "L.C."
            Height          =   275
            Index           =   1
            Left            =   5340
            TabIndex        =   60
            Top             =   1800
            Width           =   735
         End
         Begin VB.OptionButton OPTLE 
            Caption         =   "L.E."
            Height          =   275
            Index           =   1
            Left            =   5340
            TabIndex        =   59
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo"
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
            Index           =   1
            Left            =   300
            TabIndex        =   73
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Numero"
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
            Index           =   1
            Left            =   2340
            TabIndex        =   72
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo"
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
            Index           =   1
            Left            =   5160
            TabIndex        =   71
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Apellido y Nombre"
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
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Top             =   180
            Width           =   6135
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2355
         Left            =   240
         TabIndex        =   42
         Top             =   2820
         Width           =   6435
         Begin VB.OptionButton OPTLE 
            Caption         =   "L.E."
            Height          =   275
            Index           =   0
            Left            =   5340
            TabIndex        =   56
            Top             =   2040
            Width           =   735
         End
         Begin VB.OptionButton OPTLC 
            Caption         =   "L.C."
            Height          =   275
            Index           =   0
            Left            =   5340
            TabIndex        =   55
            Top             =   1800
            Width           =   735
         End
         Begin VB.OptionButton OPTCI 
            Caption         =   "C.I."
            Height          =   275
            Index           =   0
            Left            =   5340
            TabIndex        =   54
            Top             =   1560
            Width           =   735
         End
         Begin VB.OptionButton optDNI 
            Caption         =   "DNI"
            Height          =   275
            Index           =   0
            Left            =   5340
            TabIndex        =   53
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtNumero 
            Height          =   315
            Index           =   0
            Left            =   2340
            TabIndex        =   50
            Top             =   1380
            Width           =   2775
         End
         Begin VB.CheckBox chkOtra 
            Caption         =   "Otra"
            Height          =   275
            Index           =   0
            Left            =   1440
            TabIndex        =   48
            Top             =   1620
            Width           =   1035
         End
         Begin VB.CheckBox chkAval 
            Caption         =   "Aval"
            Height          =   275
            Index           =   0
            Left            =   1440
            TabIndex        =   47
            Top             =   1320
            Value           =   1  'Checked
            Width           =   1035
         End
         Begin VB.CheckBox chkFianza 
            Caption         =   "Fianza"
            Height          =   275
            Index           =   0
            Left            =   420
            TabIndex        =   46
            Top             =   1920
            Width           =   1035
         End
         Begin VB.CheckBox chkHipoteca 
            Caption         =   "Hipoteca"
            Height          =   275
            Index           =   0
            Left            =   420
            TabIndex        =   45
            Top             =   1620
            Width           =   1035
         End
         Begin VB.CheckBox chkPrenda 
            Caption         =   "Prenda"
            Height          =   275
            Index           =   0
            Left            =   420
            TabIndex        =   44
            Top             =   1320
            Width           =   1035
         End
         Begin VB.TextBox txtApellidoNombre 
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   43
            Top             =   480
            Width           =   6135
         End
         Begin VB.Label Label10 
            Caption         =   "Apellido y Nombre"
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
            Index           =   0
            Left            =   240
            TabIndex        =   57
            Top             =   180
            Width           =   6135
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo"
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
            Index           =   0
            Left            =   5160
            TabIndex        =   52
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Numero"
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
            Index           =   0
            Left            =   2340
            TabIndex        =   51
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo"
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
            Index           =   0
            Left            =   300
            TabIndex        =   49
            Top             =   960
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   3720
         TabIndex        =   40
         Top             =   11400
         Width           =   1275
      End
      Begin VB.CheckBox chkProcesarTodas 
         Caption         =   "Procesar Todas"
         Height          =   315
         Left            =   420
         TabIndex        =   39
         Top             =   11520
         Width           =   1995
      End
      Begin VB.ComboBox cboBuscaPor 
         Height          =   315
         ItemData        =   "planillac.frx":0000
         Left            =   360
         List            =   "planillac.frx":000D
         TabIndex        =   38
         Top             =   11040
         Width           =   1755
      End
      Begin VB.TextBox txtBuscar 
         Height          =   315
         Left            =   2340
         TabIndex        =   36
         Top             =   11040
         Width           =   3075
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   315
         Left            =   5880
         TabIndex        =   35
         Top             =   11760
         Width           =   915
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
         Height          =   315
         Left            =   2400
         TabIndex        =   34
         Top             =   11400
         Width           =   1155
      End
      Begin VB.CommandButton cmdBuscarOperacion 
         Caption         =   "Operación"
         Height          =   315
         Left            =   4800
         TabIndex        =   32
         Top             =   10560
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscarDocumento 
         Caption         =   "Documento"
         Height          =   315
         Left            =   3720
         TabIndex        =   31
         Top             =   10560
         Width           =   1035
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   240
         Top             =   11940
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   5580
         TabIndex        =   30
         Top             =   11040
         Width           =   1035
      End
      Begin VB.Frame Frame9 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   5820
         TabIndex        =   26
         Top             =   1200
         Width           =   915
         Begin VB.OptionButton Mza 
            Caption         =   "Mza."
            Height          =   315
            Left            =   0
            TabIndex        =   28
            Top             =   360
            Width           =   795
         End
         Begin VB.OptionButton Prev 
            Caption         =   "Prev."
            Height          =   315
            Left            =   0
            TabIndex        =   27
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.OptionButton OTRO 
         Caption         =   "OTRO"
         Height          =   315
         Left            =   5640
         TabIndex        =   25
         Top             =   2040
         Width           =   915
      End
      Begin VB.OptionButton LE 
         Caption         =   "L.E."
         Height          =   315
         Left            =   4860
         TabIndex        =   24
         Top             =   2040
         Width           =   915
      End
      Begin VB.OptionButton LC 
         Caption         =   "L.C."
         Height          =   315
         Left            =   4140
         TabIndex        =   23
         Top             =   2040
         Width           =   915
      End
      Begin VB.OptionButton DNI 
         Caption         =   "D.N.I."
         Height          =   315
         Left            =   3300
         TabIndex        =   22
         Top             =   2040
         Width           =   915
      End
      Begin VB.TextBox Numero_de_Registro 
         Height          =   330
         Left            =   4380
         TabIndex        =   20
         Top             =   780
         Width           =   1170
      End
      Begin VB.TextBox Nro_Operacion 
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   2400
         Width           =   1410
      End
      Begin VB.TextBox CajaNumero_Numero 
         Height          =   330
         Left            =   1920
         TabIndex        =   6
         Top             =   1800
         Width           =   330
      End
      Begin VB.TextBox CajaNumero_Letra 
         Height          =   330
         Left            =   1440
         TabIndex        =   5
         Top             =   1800
         Width           =   330
      End
      Begin VB.TextBox Tipo_Operacion 
         Height          =   330
         Left            =   4380
         TabIndex        =   4
         Top             =   1260
         Width           =   810
      End
      Begin VB.TextBox Nro_Cliente_Documento 
         Height          =   330
         Left            =   5040
         TabIndex        =   3
         Top             =   2400
         Width           =   1530
      End
      Begin VB.TextBox DATA 
         Height          =   330
         Left            =   6300
         TabIndex        =   1
         Top             =   780
         Width           =   330
      End
      Begin VB.TextBox Sucursal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   1320
         TabIndex        =   0
         Top             =   1080
         Width           =   630
      End
      Begin VB.CommandButton cmdAtras 
         Caption         =   "<<"
         Height          =   315
         Left            =   300
         TabIndex        =   13
         Top             =   10560
         Width           =   495
      End
      Begin VB.CommandButton cmdAdelante 
         Caption         =   ">>"
         Height          =   315
         Left            =   1140
         TabIndex        =   12
         Top             =   10560
         Width           =   495
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   10560
         Width           =   1035
      End
      Begin VB.CommandButton cdmBorrar 
         Caption         =   "Borrar"
         Height          =   315
         Left            =   2880
         TabIndex        =   10
         Top             =   10560
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Borrado"
         Height          =   375
         Index           =   3
         Left            =   3000
         TabIndex        =   98
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblbachk 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   480
         TabIndex        =   41
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblCantidad 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   780
         TabIndex        =   37
         Top             =   10560
         Width           =   375
      End
      Begin VB.Label lblError 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   480
         TabIndex        =   29
         Top             =   300
         Width           =   6375
      End
      Begin VB.Label Registro 
         Caption         =   "Registro"
         Height          =   315
         Left            =   3600
         TabIndex        =   21
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "Nº de Oper."
         Height          =   315
         Left            =   300
         TabIndex        =   19
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Caja Numero"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo Operacion"
         Height          =   315
         Left            =   3180
         TabIndex        =   17
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Doc"
         Height          =   255
         Left            =   4560
         TabIndex        =   16
         Top             =   2460
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Data :"
         Height          =   315
         Left            =   5820
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Suc."
         Height          =   255
         Left            =   660
         TabIndex        =   14
         Top             =   1140
         Width           =   435
      End
   End
   Begin ImgeditLibCtl.ImgEdit oleImgEdit1 
      Height          =   12495
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   10935
      _Version        =   65536
      _ExtentX        =   19288
      _ExtentY        =   22040
      _StockProps     =   96
      BorderStyle     =   1
      Image           =   "D:\INTERCAMBIO\Imagenes\0rs80000.tif"
      ImageControl    =   "ImgEdit1"
      Zoom            =   75
      AutoRefresh     =   -1  'True
   End
   Begin ThumbnailLibCtl.ImgThumbnail oleImgThumbnail1 
      Height          =   330
      Left            =   660
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   2955
      _Version        =   65536
      _ExtentX        =   5212
      _ExtentY        =   582
      _StockProps     =   97
      BackColor       =   -2147483638
      BorderStyle     =   1
      BackColor       =   -2147483638
      ThumbHeight     =   120
      BeginProperty ThumbCaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "planillac.frx":004B
      Height          =   1575
      Left            =   300
      TabIndex        =   33
      Top             =   12660
      Width           =   18615
      _ExtentX        =   32835
      _ExtentY        =   2778
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
   Begin AdminLibCtl.ImgAdmin oleImgAdmin1 
      Left            =   -60
      Top             =   960
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   397
      _StockProps     =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSend 
         Caption         =   "&Send..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuCopyPage 
         Caption         =   "Cop&y Page"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDeletePage 
         Caption         =   "&Delete Page"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "&Select"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuDrag 
         Caption         =   "&Drag"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Enabled         =   0   'False
      Begin VB.Menu mnuScaleToGray 
         Caption         =   "Scale to &Gray"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOnePage 
         Caption         =   "&One Page"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuThumbnail 
         Caption         =   "&Thumbnail"
      End
      Begin VB.Menu mnuPageThumbnail 
         Caption         =   "&Page and Thumbnail"
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullScreen 
         Caption         =   "&Full Screen"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolbar 
         Caption         =   "&Toolbar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuPage 
      Caption         =   "&Page"
      Enabled         =   0   'False
      Begin VB.Menu mnuNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "&Previous"
      End
      Begin VB.Menu mnuFirst 
         Caption         =   "&First"
      End
      Begin VB.Menu mnuLast 
         Caption         =   "&Last"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGoTo 
         Caption         =   "&Go To..."
      End
      Begin VB.Menu mnuBack 
         Caption         =   "Go &Back"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintPage 
         Caption         =   "Prin&t Page"
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLeft 
         Caption         =   "Rotate &Left"
      End
      Begin VB.Menu mnuRight 
         Caption         =   "Rotate &Right"
      End
      Begin VB.Menu mnuFlip 
         Caption         =   "&Flip"
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsert 
         Caption         =   "&Insert..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAppend 
         Caption         =   "&Append..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "&Convert..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRescan 
         Caption         =   "&Rescan"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuZoom 
      Caption         =   "&Zoom"
      Enabled         =   0   'False
      Begin VB.Menu mnuZoomIn 
         Caption         =   "Zoom &In"
      End
      Begin VB.Menu mnuZoomOut 
         Caption         =   "Zoom &Out"
      End
      Begin VB.Menu mnuZoomToSelection 
         Caption         =   "Zoom to &Selection"
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFitHeight 
         Caption         =   "Fit to &Height"
      End
      Begin VB.Menu mnuFitWidth 
         Caption         =   "Fit to &Width"
      End
      Begin VB.Menu mnuBestFit 
         Caption         =   "&Best Fit"
      End
      Begin VB.Menu mnuActual 
         Caption         =   "Act&ual Size"
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnu25 
         Caption         =   "&25%"
      End
      Begin VB.Menu mnu50 
         Caption         =   "&50%"
      End
      Begin VB.Menu mnu75 
         Caption         =   "&75%"
      End
      Begin VB.Menu mnu100 
         Caption         =   "&100%"
      End
      Begin VB.Menu mnu200 
         Caption         =   "2&00%"
      End
      Begin VB.Menu mnu400 
         Caption         =   "&400%"
      End
      Begin VB.Menu mnuCustom 
         Caption         =   "&Custom..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAnnotation 
      Caption         =   "&Annotation"
      Enabled         =   0   'False
      Begin VB.Menu mnuHideAnnotation 
         Caption         =   "&Hide Annotation"
      End
      Begin VB.Menu mnuBurnIn 
         Caption         =   "B&urn in Annotation"
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNoTool 
         Caption         =   "&No Tool"
      End
      Begin VB.Menu mnuSelectPointer 
         Caption         =   "Selection &Pointer"
      End
      Begin VB.Menu mnuFreeHand 
         Caption         =   "&Freehand Line"
      End
      Begin VB.Menu mnuHiLight 
         Caption         =   "H&ighlight Line"
      End
      Begin VB.Menu mnuStraightLine 
         Caption         =   "Straight &Line"
      End
      Begin VB.Menu mnuHollowRect 
         Caption         =   "Hollow &Rectangle"
      End
      Begin VB.Menu mnuFillRect 
         Caption         =   "Filled Rectan&gle"
      End
      Begin VB.Menu mnuTypedText 
         Caption         =   "Typed Text"
      End
      Begin VB.Menu mnuAttachNote 
         Caption         =   "Atta&ch-a-note"
      End
      Begin VB.Menu mnuTextFromFile 
         Caption         =   "Te&xt from File"
      End
      Begin VB.Menu mnuStamp 
         Caption         =   "Ru&bber Stamps"
      End
      Begin VB.Menu mnuSep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowTools 
         Caption         =   "Show Toolbox"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Enabled         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "&About..."
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPlanillaC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'               Copyright (C) 1995 Wang
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Wang has no warranty,
' obligations or liability for any Sample Application Files.
'
' This application is intended as an example of how to use the Wang
' Imaging OLE Controls.  As such, we have kept refinements such as
' disabling and enabling menu items, elaborate error handling, etc. to
' a minimum so as not to obscure the code that actually deals with the
' Wang Imaging OLE controls.  There are items on the menus that have
' not been implemented.  These items, in general, would involve creating
' dialog boxes and other UI that are best left to the user.  Once the user
' has an understanding of how to use the Wang Imaging Controls, these
' items should be fairly simple to implement.
' ------------------------------------------------------------------------

Dim Selection As Boolean 'Selection = True, selection rect drawn.
Dim Annot8Visible As Boolean 'Annot8Visible = True, annotation toolbox is
                            'visible
Dim CurrentPage As Integer 'CurPage = currently displayed image page
Dim LastPage As Integer 'LastPage = last page viewed before current page
Dim TotalPages As Integer 'TotalPages = image document page count
Dim numbits As Integer 'number of bits per pixel supported by this device

'Const defines
Const NoTool = 0
Const AnnoSelection = 1
Const AnnoFreehand = 2
Const AnnoHiLight = 3
Const AnnoStraightLine = 4
Const AnnoHollowRect = 5
Const AnnoFilledRect = 6
Const AnnoText = 7
Const AnnoAttachNote = 8
Const AnnoTextFromFile = 9
Const AnnoRubberStamp = 10
Const BestFit = 0
Const FitWidth = 1
Const FitHeight = 2
Const InchToInch = 3
Const ErrCancel = 32755
Const ZoomMax = 6554
Const ZoomMin = 2
Const TiffImage = 1
Const AwdImage = 2
Const BmpImage = 3
Const ImageChanged = "Image has changed.  Do you want to save changes?"

'Win API to determine display capabilities
Dim cnn1 As ADODB.Connection
Dim rsente As ADODB.Recordset
Dim cantidadRegistro  As Integer

Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Sub cdmBorrar_Click()
  If MsgBox("Usted esta segura de Borra el registro ", vbQuestion + vbYesNo) = vbYes Then
    oleImgEdit1.PrintImage 1, 1
    rsente!borrado = True
    rsente.Update
  End If
End Sub

Private Sub Check2_Click()

End Sub

Private Sub cmdActualizar_Click()
  If Actualizar = True Then
  End If
End Sub

Private Sub cmdAdelante_Click()
  Rem  On Error GoTo DatoError2
    
  MoverAdelante
  Rem    Exit Sub
  Rem DatoError2:

End Sub

Private Sub cmdAtras_Click()
  MoverAtrar

End Sub

Private Sub cmdBuscar_Click()
  









  Set rsente = New ADODB.Recordset
  Dim sql As String
  If cboBuscaPor.List(cboBuscaPor.ListIndex) <> "" And txtBuscar <> "" Then
        sql = "SELECT * from planilla_c  WHERE  "
        sql = sql & cboBuscaPor.List(cboBuscaPor.ListIndex) & " = " & txtBuscar
         Set rsente = Nothing
            Set rsente = New ADODB.Recordset
    
        rsente.Open sql, cnn1, adOpenKeyset, adLockOptimistic
  Else
    rsente.Open "SELECT * from planilla_c where ERROR = TRUE and borrado <> true  ", cnn1, adOpenKeyset, adLockOptimistic
  End If

  If Not rsente.EOF Then
    rsente.MoveFirst
    cantidadRegistro = 1
    lblCantidad = cantidadRegistro
    ColocarDatos
  End If


  MsgBox rsente.RecordCount

End Sub

Private Sub cmdBuscarDocumento_Click()


  If Nro_Cliente_Documento <> "" Then
    Adodc1.RecordSource = "SELECT BDENTE.MZ2_NOMBRE, BDENTE.MZ2_NRO_CL, BDENTE.MZ2_NRO_PR, BDENTE.MZ2_SUC, BDENTE.MZ2_TARJET," & _
        " BDENTE.MZ2_CODORI, BDENTE.MZ2_ORIGEN, BDENTE.MZ2_DEUDA, BDENTE.IDTABLE, BDENTE.MZ2_BANCO From BDENTE where MZ2_NRO_CL = " & Nro_Cliente_Documento
     
    Adodc1.Refresh
     
  Else
    MsgBox "Nro_Cliente_Documento ES VACIO"
  End If


End Sub

Private Sub cmdBuscarOperacion_Click()

  If Trim(Nro_Operacion) <> "" Then
    Adodc1.RecordSource = "SELECT BDENTE.MZ2_NOMBRE, BDENTE.MZ2_NRO_CL, BDENTE.MZ2_NRO_PR, BDENTE.MZ2_SUC, BDENTE.MZ2_TARJET," & _
        " BDENTE.MZ2_CODORI, BDENTE.MZ2_ORIGEN, BDENTE.MZ2_DEUDA, BDENTE.IDTABLE, BDENTE.MZ2_BANCO From BDENTE where BDENTE.MZ2_NRO_PR = " & Nro_Operacion
    Adodc1.Refresh
  Else
    MsgBox "Operacion Vacia"
 
  End If
 
End Sub

Private Sub cmdLimpiar_Click()
  Acuerdo_Fecha_Dia = ""
  Acuerdo_Fecha_Mes = ""
  Acuedo_Fecha_Año = ""
  Acuerdo_Firma.Valor = ""
  Acuerdo_Original_Copia.Valor = ""
        
        
  Certificado_Contador.Valor = ""
  Certificado_Extendido.Valor = ""
  Certificado_Extendido_Fecha_D = ""
  Certificado_Extendido_Fecha_M = ""
  Certificado_Extendido_Fecha_A = ""
  Certificado_Gerente.Valor = ""
        
  Documento.Valor = ""
  Documento_Fecha_Origen_Dia = ""
  Documento_Fecha_Origen_Mes = ""
  Documento_Fecha_Origen_Año = ""
  Documento_Fecha_Vencimiento_D = ""
  Documento_Fecha_Vencimiento_M = ""
  Documento_Fecha_Documento_A = ""
  Documento_Firma.Valor = ""
  Documento_Original_Copia.Valor = ""
        
  Liquidacion_Firma.Valor = ""
  Liquidacion_Original_Copia.Valor = ""
  Liquidacion_Prestamo.Valor = ""
  Liquidacion_Sello_Caja.Valor = ""
        
        
        
  Resumenes.Valor = ""
  Resumenes_Fecha_Dia = ""
  Resumenes_Fecha_Año = ""
  Resumenes_Fecha_Mes = ""
        
  Solicitud.Valor = ""
  Solicitud_Fecha_Dia = ""
  Solicitud_Fecha_Mes = ""
  Solicitud_Fecha_Año = ""
  Solicitud_Firma.Valor = ""
  Solicitud_Original_Copia.Valor = ""
  Solicitud_Refinanciacion.Valor = ""
  
  Acuerdo.Value = False
  Convenio.Value = False
  Escritura.Value = False
End Sub

Private Sub cmdProcesar_Click()

    
    
  Dim rsente As ADODB.Recordset
  Dim rsBDENTE As ADODB.Recordset
  Dim rsResidual As ADODB.Recordset
  Dim sql As String
  Dim strCnn  As String
  Dim i As Integer
  Dim ErrorDato As String
    
    
  'Connetion
    
  MousePointer = 11
 Set cnn1 = New ADODB.Connection
      Rem strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\banc.mdb"
    strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=\\Sistemas10\c\TELEform\exp\banc.mdb"
    cnn1.Open strCnn
    
    Set Cfondo = New ADODB.Connection
     strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\residual.mdb"
    Rem strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=\\Server1basa\carpeta segura\residual.mdb"
    Cfondo.Open strCnn '
    Set rsente = New ADODB.Recordset
 
 If chkProcesarTodas.Value Then
    rsente.Open "select * from planilla_c  ", cnn1, adOpenKeyset, adLockOptimistic
    
  Else
    
    BatchNo = InputBox("POR FAVOR INGRESE EL NUMERO  DE Batch", "PROCESO", 0)
    rsente.Open "select * from planilla_c  where BatchNo= " & BatchNo, cnn1, adOpenKeyset, adLockOptimistic
    
  End If
    
   
  Do While Not rsente.EOF

    If Not IsNull(rsente!Numero_de_Registro) And Not IsEmpty(rsente!Numero_de_Registro) Then
      Set rsBDENTE = New ADODB.Recordset
      rsBDENTE.Open "Select * from BDENTE where idtable = " & rsente!Numero_de_Registro, Cfondo, adOpenKeyset
       
    Else

      If rsente!Tipo_Operacion = "17" Or rsente!Tipo_Operacion = "18" Then
        If Not IsNull(rsente!Nro_Operacion) And Not IsNull(rsente!Nro_Cliente_Documento) Then
          If (rsente!Nro_Operacion) = "" And (rsente!Nro_Cliente_Documento) = "" Then
            ErrorDato = ErrorDato & "   " & "PUEDE SER UNA TARJETA"
          Else
            ErrorDato = ErrorDato & "   " & "PUEDE SER UNA TARJETA"
          End If
        End If
      Else
        If Not IsNull(rsente!Nro_Operacion) And Not IsNull(rsente!Nro_Cliente_Documento) Then
          If Trim(rsente!Nro_Operacion) = "" Or Trim(rsente!Nro_Cliente_Documento) = "" Then
            ErrorDato = ErrorDato & "   " & "Numero de Operacion esta en null"
          Else
            Set rsBDENTE = New ADODB.Recordset
            rsBDENTE.Open "Select * from BDENTE where MZ2_NRO_PR =  " & rsente!Nro_Operacion & " AND MZ2_NRO_CL = " & rsente!Nro_Cliente_Documento, Cfondo, adOpenKeyset
            If rsBDENTE.EOF And rsente!Nro_Operacion <> 0 Then
              Set rsBDENTE = New ADODB.Recordset
              rsBDENTE.Open "Select * from BDENTE where MZ2_NRO_PR =  " & rsente!Nro_Operacion, Cfondo, adOpenKeyset
              If Not rsBDENTE.EOF Then
                If rsBDENTE.RecordCount = 1 Then
                  rsente!Numero_de_Registro = rsBDENTE!IDTABLE
                  rsente.Update
                Else
                  Set rsBDENTE = Nothing
                  ErrorDato = ErrorDato & "//" & "NO existe en la tabla"
                End If
              End If
            Else
              If rsBDENTE.EOF Then
                ErrorDato = ErrorDato & "//" & "NO existe en la tabla"
              Else
                                
              End If
                        
            End If
          End If
        Else
          ErrorDato = ErrorDato & "//" & "NO existe en la tabla Doc. o Ope. is null "
        End If
      End If
    End If
     
     
    If Not (rsBDENTE Is Nothing) Then
     
      If Not rsBDENTE.EOF And ErrorDato = "" Then
        If rsBDENTE.RecordCount = 1 Then
          rsente!RegistroEncontrado = rsBDENTE!IDTABLE
        Else
          ErrorDato = ErrorDato & "//" & "NO existen en tabla tabla mas de dos registros"
        End If
      Else
        ErrorDato = ErrorDato & "//" & "NO existen en tabla"
      End If
    Else
      ErrorDato = ErrorDato & "//" & "NO existen en tabla"
    End If
     
   

            
                  
   
            
    If Trim(rsente!DATA) = "" Or IsNull(rsente!DATA) Then
      ErrorDato = ErrorDato & "//" & " Data ** "
    Else
    rsente!DATA = UCase(Trim(rsente!DATA))
    End If
     Select Case Trim(rsente!DATA)
      Case "l", "L", "i"
        rsente!DATA = 1
      Case "u", "U"
        rsente!DATA = "V"
      Case "B"
        rsente!DATA = "8"
      Case "J"
        rsente!DATA = "1"
      Case "x"
        rsente!DATA = "X"
      Case "o", "O"
        rsente!DATA = "0"
      Case "z", "Z"
        rsente!DATA = "2"
      Case "S"
       rsente!DATA = "5"
      Case "A", "2", "3", "5", "4", "6", "8", "28"
        rsente!Mayores_Menores = "A"
      Case "0", "1", "7", "9", "C", "V", "X", "N", "P", "13", "12", "15", "14", "11", "26", "28", "25", "22", "23", "36", "24", "22", "16", "17", "45", "44", "47", "48", "49"
        rsente!Mayores_Menores = "E"
      Case Else
        ErrorDato = ErrorDato & "//" & " Data "
    End Select
            
           
            
   
    If Trim(rsente!Sucursal) = "999" Then
      ErrorDato = ErrorDato & "//" & " LA FICHA ESTA CORRIDA ** "
    End If
        
    
    
           
            
    If Trim(rsente!Tipo_Operacion) = "" Or IsNull(rsente!Tipo_Operacion) Then
      ErrorDato = ErrorDato & "//" & "Tipo_Operacion  **"
      If Not (Trim(rsente!Tipo_Operacion) > 1 And Trim(rsente!Tipo_Operacion) < 44) Then
        ErrorDato = ErrorDato & "//" & "Tipo_Operacion incorrecta  **"
      End If
    End If
    rsente!Apellido_Garantia_1 = UCase(rsente!Apellido_Garantia_1)
    rsente!Apellido_Garantia_2 = UCase(rsente!Apellido_Garantia_2)
    rsente!Apellido_Garantia_3 = UCase(rsente!Apellido_Garantia_3)
            
    If IsNull(rsente!Tipo_Garantia_1) Then
           
    End If
            
    If IsNull(rsente!Tipo_Garantia_2) Then
            
    End If
           
    If IsNull(rsente!Tipo_Garantia_3) Then
            
    End If
           
   On Error GoTo LUIS
    If Trim(ErrorDato) <> "" Then
      rsente!descripcionerror = Mid(ErrorDato, 1, 240)
      rsente!Error = True
      rsente!insertar = False
      rsente.Update
    Else
      rsente!descripcionerror = "NO"
      rsente!Error = False
      rsente!insertar = True
      rsente.Update
    End If
           
    rsente.Update
LUIS:
    ErrorDato = ""
    rsente.MoveNext
  Loop

  MousePointer = 0
End Sub

Private Sub cmdRegistro_Click()
 If Trim(Numero_de_Registro) <> "" Then
    Adodc1.RecordSource = "SELECT BDENTE.MZ2_NOMBRE, BDENTE.MZ2_NRO_CL, BDENTE.MZ2_NRO_PR, BDENTE.MZ2_SUC, BDENTE.MZ2_TARJET," & _
        " BDENTE.MZ2_CODORI, BDENTE.MZ2_ORIGEN, BDENTE.MZ2_DEUDA, BDENTE.IDTABLE, BDENTE.MZ2_BANCO From BDENTE where BDENTE.IDTABLE = " & Numero_de_Registro
    Adodc1.Refresh
  Else
    MsgBox "Numero_Registro Vacio"
 
  End If
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Usted esta segura de Borra el registro ", vbQuestion + vbYesNo) = vbYes Then
        rsente.Delete
        rsente.Update
    End If
End Sub

Private Sub Command2_Click()

  Dim cnn1 As ADODB.Connection
  Dim Cfondo As ADODB.Connection
  Dim rsente As ADODB.Recordset
  Dim rsBDENTE As ADODB.Recordset
  Dim rsResidual As ADODB.Recordset
  Dim sql As String
  Dim strCnn  As String
  Dim i As Integer
    
  'Connetion
  Set cnn1 = New ADODB.Connection
  strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\intercambio\banc.mdb"
  cnn1.Open strCnn
    
  Set Cfondo = New ADODB.Connection
  strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\Fondoresidual.mdb"
  Cfondo.Open strCnn
    
    
  Set rsente = New ADODB.Recordset
  rsente.Open "select * from banc1", cnn1, adOpenKeyset, adLockOptimistic
  Set rsResidual = New ADODB.Recordset
  rsResidual.Open "select * from Residual", Cfondo, adOpenKeyset, adLockOptimistic
        
  Dim ErrorDato As String



  Do While Not rsente.EOF

    If Not IsNull(rsente!Numero_de_Registro) And Not IsEmpty(rsente!Numero_de_Registro) Then
      Set rsBDENTE = New ADODB.Recordset
      rsBDENTE.Open "Select * from BDENTE where idtable = " & rsente!Numero_de_Registro, Cfondo
    Else

      If rsente!Tipo_Operacion = "17" Or rsente!Tipo_Operacion = "18" Then
        If Not IsNull(rsente!Nro_Operacion) And Not IsNull(rsente!Nro_Cliente_Documento) Then
          If (rsente!Nro_Operacion) = "" And (rsente!Nro_Cliente_Documento) = "" Then
            ErrorDato = ErrorDato & "   " & "Solicitud"
          Else

            ErrorDato = ErrorDato & "   " & "PUEDE SER UNA TARJETA"


            '                     Set rsBDENTE = New ADODB.Recordset
            '                     rsBDENTE.Open "Select * from BDENTE where MZ2_TARJET =  " & rsEnte!Nro_Operacion & " AND MZ2_NRO_CL = " & rsEnte!Nro_Cliente_Documento, Cfondo
            '                     If rsBDENTE.EOF Then
            '                        Set rsBDENTE = New ADODB.Recordset
            '                        rsBDENTE.Open "Select * from BDENTE where MZ2_TARJET =  " & rsEnte!Nro_Operacion & " AND MZ2_NRO_CL = " & rsEnte!Nro_Cliente_Documento, Cfondo
            '                    End If
          End If
        End If
      Else
        If Not IsNull(rsente!Nro_Operacion) And Not IsNull(rsente!Nro_Cliente_Documento) Then
          If (rsente!Nro_Operacion) = "" And (rsente!Nro_Cliente_Documento) = "" Then
            ErrorDato = ErrorDato & "   " & "Numero de Operacion esta en null"
          Else
            Set rsBDENTE = New ADODB.Recordset
            rsBDENTE.Open "Select * from BDENTE where MZ2_NRO_PR =  " & rsente!Nro_Operacion & " AND MZ2_NRO_CL = " & rsente!Nro_Cliente_Documento, Cfondo
            If rsBDENTE.EOF And rsente!Nro_Operacion <> 0 Then
              Set rsBDENTE = New ADODB.Recordset
              rsBDENTE.Open "Select * from BDENTE where MZ2_NRO_PR =  " & rsente!Nro_Operacion, Cfondo, adOpenKeyset
              If Not rsBDENTE.EOF Then
                If rsBDENTE.RecordCount = 1 Then
                  rsente!Numero_de_Registro = rsBDENTE!IDTABLE
                  rsente.Update
                Else
                  ErrorDato = ErrorDato & "//" & "NO existe en la tabla"
                End If
              End If
            End If
          End If
        End If
      End If
    End If
     
    ' Controlo que el registro exista en la tabla que me dieron de residual
    
       
    If rsente!Solicitud <> "" Or rsente!Solicitud_Firma <> "" Or rsente!Solicitud_Original_Copia <> "" Or rsente!Solicitud_Refinanciacion <> "" Then
      If Trim(rsente!Solicitud) = "" Then
        ErrorDato = ErrorDato & "//" & "Solicitud ** "
      End If
      If Trim(rsente!Solicitud_Firma) = "" Or IsNull(rsente!Solicitud_Firma) Then
        ErrorDato = ErrorDato & "   " & "Solicitud_Firma ** "
      End If
      If Trim(rsente!Solicitud_Original_Copia) = "" Or IsNull(rsente!Solicitud_Original_Copia) Then
        ErrorDato = ErrorDato & "   " & "Solicitud_Original_Copia ** "
      End If
      If Trim(rsente!Solicitud_Refinanciacion) = "" Or IsNull(rsente!Solicitud_Refinanciacion) Then
        ErrorDato = ErrorDato & "   " & "Solicitud_Refinanciacion"
      End If
    End If
        
        
    If rsente!Acuerdo_Firma <> "" Or rsente!Acuerdo_Original_Copia <> "" Then
      If rsente!Acuerdo_Firma = "" Then
        ErrorDato = ErrorDato & " // " & "Acuerdo_Firma ** "
      End If
      If rsente!Acuerdo_Original_Copia = "" Or IsNull(rsente!Acuerdo_Original_Copia) Then
        ErrorDato = ErrorDato & "   " & "Acuerdo_Original_Copia ** "
      End If
             
    End If
    
        
    If rsente!Liquidacion_Firma <> "" Or rsente!Liquidacion_Original_Copia <> "" Or rsente!Liquidacion_Prestamo <> "" Or rsente!Liquidacion_Sello_Caja <> "" Then
      If rsente!Liquidacion_Firma = "" Or IsNull(rsente!Liquidacion_Firma) Then
        ErrorDato = ErrorDato & " //  " & "Liquidacion_Firma ** "
      End If
            
      If rsente!Liquidacion_Original_Copia = "" Or IsNull(rsente!Liquidacion_Original_Copia) Then
        ErrorDato = ErrorDato & "   " & "Liquidacion_Original_Copia ** "
      End If
            
      If rsente!Liquidacion_Prestamo = "" Or IsNull(rsente!Liquidacion_Prestamo) Then
        ErrorDato = ErrorDato & "   " & "Liquidacion_Prestamo ** "
      End If
            
      If rsente!Liquidacion_Sello_Caja = "" Or IsNull(rsente!Liquidacion_Sello_Caja) Then
        ErrorDato = ErrorDato & "   " & "Liquidacion_Sello_Caja ** "
      End If
    End If
        
        
    If rsente!Documento <> "" Or rsente!Documento_Firma <> "" Or rsente!Documento_Original_Copia <> "" Then
      If rsente!Documento = "" Or IsNull(rsente!Documento) Then
        ErrorDato = ErrorDato & " //" & "Documento ** "
      End If
            
      If rsente!Documento_Firma = "" Or IsNull(rsente!Documento_Firma) Then
        ErrorDato = ErrorDato & " //" & "Documento_Firma ** "
      End If
            
            
      If rsente!Documento_Original_Copia = "" Or IsNull(rsente!Documento_Original_Copia) Then
        ErrorDato = ErrorDato & " //" & "Documento_Original_Copia ** "
      End If
 
    End If
         
        
    If rsente!Certificado_Contador <> "" Or rsente!Certificado_Extendido <> "" Or rsente!Certificado_Gerente <> "" Then
      If rsente!Certificado_Contador = "" Or IsNull(rsente!Certificado_Contador) Then
        ErrorDato = ErrorDato & " //" & "Certificado_Contador ** "
      End If
      If rsente!Certificado_Extendido = "" Or IsNull(rsente!Certificado_Extendido) Then
        ErrorDato = ErrorDato & " " & "Certificado_Extendido ** "
      End If
      If rsente!Certificado_Gerente = "" Or IsNull(rsente!Certificado_Gerente) Then
        ErrorDato = ErrorDato & " " & "Certificado_Gerente ** "
      End If
    End If
        
        
        
        
        
    ErrorDato = ErrorDato & " " & D(rsente!Acuerdo_Fecha_Dia, "Acuerdo_Fecha_Dia")
    ErrorDato = ErrorDato & " " & M(rsente!Acuerdo_Fecha_Mes, "Acuerdo_Fecha_Mes")
    ErrorDato = ErrorDato & " " & A(rsente!Acuedo_Fecha_Año, "Acuedo_Fecha_Año")

    ErrorDato = ErrorDato & " " & D(rsente!Certificado_Extendido_Fecha_D, "Certificado_Extendido_Fecha_D")
    ErrorDato = ErrorDato & " " & M(rsente!Certificado_Extendido_Fecha_M, "Certificado_Extendido_Fecha_M")
    ErrorDato = ErrorDato & " " & A(rsente!Certificado_Extendido_Fecha_A, "Acuedo_Fecha_Año")

    ErrorDato = ErrorDato & " " & D(rsente!Documento_Fecha_Origen_Dia, "Documento_Fecha_Origen_Dia")
    ErrorDato = ErrorDato & " " & M(rsente!Documento_Fecha_Origen_Mes, "Documento_Fecha_Origen_Mes")
    ErrorDato = ErrorDato & " " & A(rsente!Documento_Fecha_Origen_Año, "Documento_Fecha_Origen_Año")
            
    ErrorDato = ErrorDato & " " & D(rsente!Documento_Fecha_Vencimiento_D, "Documento_Fecha_Vencimiento_D")
    ErrorDato = ErrorDato & " " & M(rsente!Documento_Fecha_Vencimiento_M, "Documento_Fecha_Vencimiento_M")
    ErrorDato = ErrorDato & " " & A(rsente!Documento_Fecha_Vencimiento_A, "Documento_Fecha_Vencimiento_A")

    ErrorDato = ErrorDato & " " & D(rsente!Resumenes_Fecha_Dia, "Resumenes_Fecha_Dia")
    ErrorDato = ErrorDato & " " & M(rsente!Resumenes_Fecha_Mes, "Resumenes_Fecha_Mes")
    ErrorDato = ErrorDato & " " & A(rsente!Resumenes_Fecha_Año, "Resumenes_Fecha_Año")
            
    ErrorDato = ErrorDato & " " & D(rsente!Solicitud_Fecha_Dia, "Solicitud_Fecha_Dia")
    ErrorDato = ErrorDato & " " & M(rsente!Solicitud_Fecha_Mes, "Solicitud_Fecha_Mes")
    ErrorDato = ErrorDato & " " & A(rsente!Solicitud_Fecha_Año, "Solicitud_Fecha_Año")
        
        
        
    If Trim(ErrorDato) <> "" Then
      rsente!Error = Mid(ErrorDato, 1, 255)
      rsente.Update
    Else
      rsente!Error = "NO"
      rsente.Update
    End If
        
    ErrorDato = ""
      
      

    rsente.MoveNext
  Loop


End Sub

Private Sub Command3_Click()
  MoverAdelante
  oleImgEdit1.PrintImage 1, 1
 

End Sub

Private Sub cmdSalir_Click()

End Sub

Private Sub DataGrid1_DblClick()
  Numero_de_Registro = DataGrid1.Columns(8).Text

End Sub

Public Function F(D As Variant, M As Variant, A As Variant) As Date

 If IsNull(D) Or IsNull(M) Or IsNull(A) Then
    F = 0
 Else
 
 F = CDate(D & "/" & M & "/" & A)
 
 End If




End Function



Public Function M(mes As Variant, DATO As String) As String
  M = ""
    
  If Trim(mes) <> "" And Not IsNull(mes) Then
    If mes > 12 Or mes < 1 Then
      M = DATO
    End If
  End If

End Function

Public Function A(ano As Variant, DATO As String)
  A = ""
    
  If Trim(ano) <> "" And Not IsNull(ano) Then
    If ano > 99 Or ano < 60 Then
      A = DATO
    End If
  End If
End Function

Public Function D(dia As Variant, DATO As String) As String
  D = ""
  
  If Trim(dia) <> "" And Not IsNull(dia) Then
    If dia > 32 Or dia < 1 Then
      D = DATO
    End If
  End If
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    KeyAscii = vbKeyTab
  End If


End Sub

Private Sub Form_Load()
    
  Dim sql As String
  Dim strCnn  As String
  Dim i As Integer
    
  'Connetion
  Set cnn1 = New ADODB.Connection
    Rem  strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Sistema\Sistema  Residual\banc.mdb"
    strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=\\Sistemas10\c\TELEform\exp\banc.mdb"
    cnn1.Open strCnn
    
    Set Cfondo = New ADODB.Connection
     strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\residual.mdb"
    Rem strCnn = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=\\Server1basa\carpeta segura\residual.mdb"
    Cfondo.Open strCnn
    
 

  'initialize the variables
  Dim dc As Long
  Dim index As Long








  Rem Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=\\Server1basa\carpeta segura\residual.mdb"
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\residual.mdb"


  Selection = False
  Annot8Visible = False
  CurrentPage = 1
  LastPage = 1
  TotalPages = 1

  dc = hdc
  index = 12  '12 = BITSPERPIXEL
  numbits = GetDeviceCaps(dc, index) 'finds out how many colors video driver supports

End Sub




Private Sub Form_Unload(Cancel As Integer)
  'if image has changed, give the user a chance to
  'save it before closing
  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
    End If
  End If

End Sub






Private Sub mnu100_Click()
  'Set zoom to 100% and redisplay image.
  'Zoom value is a float
  oleImgEdit1.Zoom = 100!
  oleImgEdit1.Refresh

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = True
  mnu200.Checked = False
  mnu400.Checked = False
End Sub


Private Sub mnu200_Click()
  'Set zoom to 200% and redisplay image.
  'Zoom value is a float
  oleImgEdit1.Zoom = 200!
  oleImgEdit1.Refresh

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = True
  mnu400.Checked = False
End Sub


Private Sub mnu25_Click()
  'Set zoom to 25% and redisplay image.
  'Zoom value is a float
  oleImgEdit1.Zoom = 25!
  oleImgEdit1.Refresh

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = False
  mnu25.Checked = True
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub


Private Sub mnu400_Click()
  'Set zoom to 400% and redisplay image.
  'Zoom value is a float
  oleImgEdit1.Zoom = 400!
  oleImgEdit1.Refresh

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = True
End Sub


Private Sub mnu50_Click()
  'Set zoom to 50% and redisplay image.
  'Zoom value is a float
  oleImgEdit1.Zoom = 50!
  oleImgEdit1.Refresh

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = False
  mnu25.Checked = False
  mnu50.Checked = True
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub


Private Sub mnu75_Click()
  'Set zoom to 75% and redisplay image.
  'Zoom value is a float
  oleImgEdit1.Zoom = 75!
  oleImgEdit1.Refresh

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = True
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub


Private Sub mnuAbout_Click()
  'Add your code here
  MsgBox "Function to be implemented by user."
End Sub


Private Sub mnuActual_Click()
  'Set fit to inch to inch and redisplay image.
  oleImgEdit1.FitTo (InchToInch)
  oleImgEdit1.Refresh

  'check the current zoom menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = True
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub


Private Sub mnuAppend_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub


Private Sub mnuAttachNote_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool AnnoAttachNote


  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = True
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = False
  mnuStamp.Checked = False
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = False

End Sub


Private Sub mnuBack_Click()
  'Save current page if modified, then return to the
  'previously displayed page.
  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
    End If
  End If
  oleImgEdit1.page = LastPage
  oleImgEdit1.Display

  'Update the selected page thumbnail
  oleImgThumbnail1.DeselectAllThumbs
  oleImgThumbnail1.ThumbSelected(LastPage) = True

End Sub


Private Sub mnuBestFit_Click()
  'zoom the image so that the entire image
  'fits in the display window
  oleImgEdit1.FitTo (BestFit)

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = True
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuZoomToSelection.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub


Private Sub mnuBurnIn_Click()
  'burn in all visible annotations,(1,) and preserve
  'colors, (2).  See documentation for other valid arguments.
  ' ret = oleImgEdit1.BurnInAnnotations(1, 2)

End Sub


Private Sub mnuConvert_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."
End Sub

Private Sub mnuCopy_Click()
  'Copy the selected area to the clipboard.
  If Selection = True Then
    oleImgEdit1.ClipboardCopy
  End If
End Sub


Private Sub mnuCopyPage_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub


Private Sub mnuCustom_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub


Private Sub mnuCut_Click()
  'Cut the selected area to the clipboard.
  If Selection = True Then
    oleImgEdit1.ClipboardCut
  End If

End Sub


Private Sub mnuDeletePage_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub

Private Sub mnuDrag_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub


Private Sub mnuExit_Click()
  'Close the app

  Unload frmSample

End Sub

Private Sub mnuFillRect_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool AnnoFilledRect


  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = True
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = False
  mnuStamp.Checked = False
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = False

End Sub

Private Sub mnuFirst_Click()
  'Save current page if modified, then store the current
  'page number and display the first page
  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
    End If
  End If
  LastPage = oleImgEdit1.page
  oleImgEdit1.page = 1
  oleImgEdit1.Display

  'Update the selected page thumbnail
  oleImgThumbnail1.DeselectAllThumbs
  oleImgThumbnail1.ThumbSelected(1) = True
End Sub

Private Sub mnuFitHeight_Click()
  'Zoom the image so that its vertical
  'dimension fits within the display window
  oleImgEdit1.FitTo (FitHeight)

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = True
  mnuFitWidth.Checked = False
  mnuZoomToSelection.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub

Private Sub mnuFitWidth_Click()
  'Zoom the image so that its horizontal
  'dimension fits within the display window
  oleImgEdit1.FitTo (FitWidth)

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = True
  mnuZoomToSelection.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub

Private Sub mnuFlip_Click()
  'Rotate the image 180 degrees.
  oleImgEdit1.Flip
End Sub

Private Sub mnuFreeHand_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool AnnoFreehand


  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = True
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = False
  mnuStamp.Checked = False
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = False

End Sub

Private Sub mnuFullScreen_Click()
  'resize the Image Edit window to maximize the display

  If mnuFullScreen.Checked Then
    frmSample.WindowState = 0
    mnuFullScreen.Checked = False
  Else
    frmSample.WindowState = 2
    mnuFullScreen.Checked = True
  End If
End Sub

Private Sub mnuGoTo_Click()
  'Save current page if modified, then store the current
  'page number and display the GoTo Page dialog box
  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
    End If
  End If
  LastPage = oleImgEdit1.page
  frmGotoDlg.Show
End Sub

Private Sub mnuHelp_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub


Private Sub mnuHideAnnotation_Click()
  'Toggle the display of annotations
  If mnuHideAnnotation.Checked = True Then
    'show all hidden annotations
    oleImgEdit1.ShowAnnotationGroup
    oleImgEdit1.Refresh
    mnuHideAnnotation.Checked = False
  Else
    'hide all displayed annotations
    oleImgEdit1.HideAnnotationGroup
    oleImgEdit1.Refresh
    mnuHideAnnotation.Checked = True
  End If

End Sub

Private Sub mnuHiLight_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool AnnoHiLight


  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = True
  mnuHollowRect.Checked = False
  mnuStamp.Checked = False
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = False

End Sub

Private Sub mnuHollowRect_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool AnnoHollowRect


  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = True
  mnuStamp.Checked = False
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = False

End Sub

Private Sub mnuInsert_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub

Private Sub mnuLast_Click()
  'Save current page if modified, then store the current
  'page number and display the last page
  Dim page As Long 'number of last page

  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
    End If
  End If
  LastPage = oleImgEdit1.page
  page = oleImgEdit1.PageCount
  oleImgEdit1.page = page
  oleImgEdit1.Display

  'Update the selected page thumbnail
  oleImgThumbnail1.DeselectAllThumbs
  oleImgThumbnail1.ThumbSelected(oleImgEdit1.page) = True

End Sub

Private Sub mnuLeft_Click()
  'Rotate image 90 degrees to the left
  oleImgEdit1.RotateLeft

End Sub

Private Sub mnuNew_Click()
  'if the current image was modified, give the user
  'a chance to save it, then open a new blank image
  'of the same size.
  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
    End If
  End If
  'Use generic display values
  oleImgEdit1.DisplayBlankImage 500, 400, 200, 200, 1
  oleImgEdit1.Image = ""
  oleImgThumbnail1.Image = oleImgEdit1.Image


  'Now that we have an image, enable the needed menus.
  mnuSaveAs.Enabled = True
  mnuSave.Enabled = True
  mnuPrint.Enabled = True
  mnuEdit.Enabled = True
  mnuView.Enabled = True
  mnuPage.Enabled = True
  mnuZoom.Enabled = True
  mnuAnnotation.Enabled = True
  'This is a 1 page image, so disable the page
  'change menu items
  mnuBack.Enabled = False
  mnuFirst.Enabled = False
  mnuGoTo.Enabled = False
  mnuLast.Enabled = False
  mnuNext.Enabled = False
  mnuPrevious.Enabled = False



End Sub

Private Sub mnuNext_Click()
  'Save current page if modified, then store the current
  'page number and display the next page
  Dim page As Long 'Page place holder

  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
    End If
  End If
  LastPage = oleImgEdit1.page
  page = oleImgEdit1.page
  If page = TotalPages Then
    MsgBox "Last Page"
    Exit Sub
  End If
  page = page + 1
  oleImgEdit1.page = page
  oleImgEdit1.Display

  'Update the selected page thumbnail
  oleImgThumbnail1.DeselectAllThumbs
  oleImgThumbnail1.ThumbSelected(oleImgEdit1.page) = True

End Sub

Private Sub mnuNoTool_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool NoTool


  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = True
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = False
  mnuStamp.Checked = False
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = False

End Sub

Private Sub mnuOnePage_Click()
  'hide any thumbnails and expand the Image Edit
  'window to fit in the form

  oleImgThumbnail1.Visible = False
  oleImgEdit1.Visible = True
  oleImgEdit1.Left = frmSample.ScaleLeft
  oleImgEdit1.Top = frmSample.ScaleTop
  oleImgEdit1.Width = frmSample.ScaleWidth
  oleImgEdit1.Height = frmSample.ScaleHeight

  mnuThumbnail.Checked = False
  mnuOnePage.Checked = True
  mnuPageThumbnail.Checked = False

End Sub


Private Sub mnuOpen_Click()
  'open an image doc. If the current doc is modified,
  'try to save it. ShowFileDialog(0) shows Open File
  'dialog. ShowFileDialog(1) shows SaveAs File dialog.


  Dim temp As String 'image name and path

  On Error Resume Next 'handle errors ourselves incase of cancel
  oleImgAdmin1.Flags = 0 'clear Flags
  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
      If Err = ErrCancel Then '32755 = Cancel pressed
        Exit Sub
      End If
    End If
  End If
  oleImgAdmin1.Filter = "All Image Files|*.tif;*.bmp;*.jpg;*.pcx;*.dcx|TIFF files (*.tif)|*.tif|BMP files (*.bmp)|*.bmp|PCX/DCX Document (*.pcx, *.dcx)|*.pcx;*.dcx|JPG File (*.jpg)|*.jpg|All Files (*.*)|*.*|"
  oleImgAdmin1.ShowFileDialog 0, frmSample.hWnd
  If Err = ErrCancel Then '32755 = Cancel pressed
    Exit Sub
  End If
  If oleImgAdmin1.StatusCode <> 0 Then
    MsgBox Err.Description + " Code = " + Hex(oleImgAdmin1.StatusCode), 16
    Exit Sub
  End If
  temp = oleImgAdmin1.Image
  oleImgEdit1.Image = temp
  oleImgThumbnail1.Image = oleImgEdit1.Image
  If numbits > 8 Then 'video driver supports hicolor or truecolor
    oleImgEdit1.ImagePalette = 3 'Set for 24 bit RGB.
  End If
  oleImgEdit1.page = 1
  oleImgEdit1.Display
  TotalPages = oleImgEdit1.PageCount
  oleImgThumbnail1.ThumbSelected(1) = True

  'Now that we have an image, enable the needed menus.
  mnuSaveAs.Enabled = True
  mnuSave.Enabled = True
  mnuPrint.Enabled = True
  mnuEdit.Enabled = True
  mnuView.Enabled = True
  mnuPage.Enabled = True
  mnuZoom.Enabled = True
  mnuAnnotation.Enabled = True
  If oleImgEdit1.PageCount > 1 Then
    mnuBack.Enabled = True
    mnuFirst.Enabled = True
    mnuGoTo.Enabled = True
    mnuLast.Enabled = True
    mnuNext.Enabled = True
    mnuPrevious.Enabled = True
  Else
    mnuBack.Enabled = False
    mnuFirst.Enabled = False
    mnuGoTo.Enabled = False
    mnuLast.Enabled = False
    mnuNext.Enabled = False
    mnuPrevious.Enabled = False
  End If

End Sub

Private Sub mnuOptions_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub


Private Sub mnuPageThumbnail_Click()
  'Show the thumbnails accross the top third of the
  'app window, and the current image in the bottom two
  'thirds of the window.

  oleImgEdit1.Visible = True
  oleImgThumbnail1.Visible = True

  oleImgEdit1.Left = frmSample.ScaleLeft
  oleImgEdit1.Top = frmSample.ScaleHeight / 3
  oleImgEdit1.Width = frmSample.ScaleWidth
  oleImgEdit1.Height = (frmSample.ScaleHeight * 2 / 3)

  oleImgThumbnail1.Left = frmSample.ScaleLeft
  oleImgThumbnail1.Top = frmSample.ScaleTop
  oleImgThumbnail1.Width = frmSample.ScaleWidth
  oleImgThumbnail1.Height = frmSample.ScaleHeight / 3

  mnuThumbnail.Checked = False
  mnuOnePage.Checked = False
  mnuPageThumbnail.Checked = True

End Sub

Private Sub mnuPaste_Click()
  'Paste from the clipboard
  If oleImgEdit1.IsClipboardDataAvailable = True Then
    oleImgEdit1.ClipboardPaste
    Selection = False
  End If

End Sub

Private Sub mnuPrevious_Click()
  'Save current page if modified, then store the current
  'page number and display the previous page
  Dim page As Long 'Page number place holder

  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
    End If
  End If
  LastPage = oleImgEdit1.page
  page = oleImgEdit1.page
  If page = 1 Then
    MsgBox "First Page"
    Exit Sub
  End If
  page = page - 1
  oleImgEdit1.page = page
  oleImgEdit1.Display

  'Update the selected page thumbnail
  oleImgThumbnail1.DeselectAllThumbs
  oleImgThumbnail1.ThumbSelected(oleImgEdit1.page) = True

End Sub

Private Sub mnuPrint_Click()
  'Open ImgAdmin's Print dialog and call ImgEdit's
  'Print function with the user selected options.
  Dim format As Integer
  Dim Annotations As Boolean

  On Error Resume Next 'handle errors ourselves in case of cancel
  If oleImgEdit1.ImageModified = True Then
    If MsgBox("The Image must be saved first if changes are to be printed.  Do you want to save the image?", vbYesNo) = vbYes Then
      mnuSave_Click
    End If
  End If
  oleImgAdmin1.Flags = 0 'clear Flags so print dialog box will display
  oleImgAdmin1.ShowPrintDialog frmSample.hWnd
  If oleImgAdmin1.StatusCode = 0 Then 'OK button selected
    format = oleImgAdmin1.PrintOutputFormat
    Annotations = oleImgAdmin1.PrintAnnotations
    Rem  X = oleImgEdit1.PrintImage(1, 1, format, Annotations)
    oleImgEdit1.PrintImage 1, 1, format, Annotations

  Else
    If Err = ErrCancel Then '32755 = Cancel pressed
      Exit Sub
    Else
      MsgBox Err.Description + " Code = " + Hex(oleImgAdmin1.StatusCode), 16
    End If
  End If
  If oleImgEdit1.StatusCode <> 0 Then
    MsgBox Err.Description + " Code = " + Hex(oleImgEdit1.StatusCode), 16
  End If

End Sub

Private Sub mnuPrintPage_Click()
  'Print the current page.

  On Error Resume Next 'handle errors ourselves
  oleImgEdit1.Zoom = 130
  oleImgEdit1.PrintImage oleImgEdit1.page, oleImgEdit1.page
  If oleImgEdit1.StatusCode <> 0 Then
    MsgBox Err.Description + " Code = " + Hex(oleImgEdit1.StatusCode), 16
  End If

End Sub


Private Sub mnuRescan_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub


Private Sub mnuRight_Click()
  'Rotate image 90 degrees to the right
  oleImgEdit1.RotateRight

End Sub

Private Sub mnuSave_Click()
  'Save the current document
  On Error Resume Next 'handle errors ourselves
  If oleImgEdit1.Image = "" Then
    mnuSaveAs_Click
  Else
    oleImgEdit1.Save (False)
    If oleImgEdit1.StatusCode <> 0 Then
      MsgBox Err.Description + " Code = " + Hex(oleImgEdit1.StatusCode), 16
    End If
  End If

End Sub

Private Sub mnuSaveAs_Click()
  'Open ImgAdmin's SaveAs dialog
  Dim FileType As Integer

  On Error Resume Next 'handle errors ourselves

  'we can write tiff, bmp, and awd files, so set the admin file filter
  'to show only these types.
  oleImgAdmin1.Filter = "TIFF files (*.tif)|*.tif|BMP files (*.bmp)|*.bmp|"
  oleImgAdmin1.ShowFileDialog 1, frmSample.hWnd
  If Err = ErrCancel Then '32755 = Cancel pressed
    Exit Sub
  End If

  If oleImgAdmin1.Image = oleImgEdit1.Image Then 'Save as current name
    oleImgEdit1.Save False
  Else 'Save as newly selected name and change image name to selected name
    'determine from the filter index which file type was selected
    If oleImgAdmin1.FilterIndex = 1 Then
      FileType = TiffImage
    ElseIf oleImgAdmin1.FilterIndex = 2 Then
      FileType = BmpImage
    End If
    oleImgEdit1.SaveAs oleImgAdmin1.Image, FileType
    oleImgEdit1.Image = oleImgAdmin1.Image
    oleImgAdmin1.Image = oleImgEdit1.Image 'this forces a refresh of the properties in the Admin control
    
  End If
  oleImgAdmin1.FilterIndex = 0
  oleImgAdmin1.Filter = ""
  If oleImgEdit1.StatusCode <> 0 Then
    MsgBox Err.Description + " Code = " + Hex(oleImgEdit1.StatusCode), 16
    Exit Sub
  End If

End Sub

Private Sub mnuScaleToGray_Click()
  'toggle image in 4 bit grayscale
  If mnuScaleToGray.Checked = True Then
    oleImgEdit1.DisplayScaleAlgorithm = 0
    oleImgEdit1.Refresh
    mnuScaleToGray.Checked = False
  Else
    oleImgEdit1.DisplayScaleAlgorithm = 2
    oleImgEdit1.Refresh
    mnuScaleToGray.Checked = True
  End If
End Sub



Private Sub mnuSelect_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub


Private Sub mnuSelectPointer_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool AnnoSelection


  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = True
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = False
  mnuStamp.Checked = False
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = False

End Sub


Private Sub mnuSend_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub



Private Sub mnuShowTools_Click()
  'If the tool palette is visible, close it.  If it's
  'not visible, open it.
  If Annot8Visible = True Then
    oleImgEdit1.HideAnnotationToolPalette
    Annot8Visible = False
    mnuShowTools.Checked = False
  Else
    oleImgEdit1.ShowAnnotationToolPalette
    Annot8Visible = True
    mnuShowTools.Checked = True
  End If

End Sub


Private Sub mnuStamp_Click()
  'Bring up the Rubber Stamp Properties dialog to choose the stamp you want.
  oleImgEdit1.ShowRubberStampDialog



  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = False
  mnuStamp.Checked = True
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = False

End Sub


Private Sub mnuStraightLine_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool AnnoStraightLine


  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = False
  mnuStamp.Checked = False
  mnuStraightLine.Checked = True
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = False

End Sub

Private Sub mnuTextFromFile_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool AnnoTextFromFile



  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = False
  mnuStamp.Checked = False
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = True
  mnuTypedText.Checked = False

End Sub

Private Sub mnuThumbnail_Click()
  'Size the thumbnail window to the app window and
  'display it. Hide the image window.
  oleImgThumbnail1.Left = frmSample.ScaleLeft
  oleImgThumbnail1.Top = frmSample.ScaleTop
  oleImgThumbnail1.Width = frmSample.ScaleWidth
  oleImgThumbnail1.Height = frmSample.ScaleHeight
  oleImgThumbnail1.Visible = True
  oleImgEdit1.Visible = False
  'oleImgThumbnail1.Image = oleImgEdit1.Image
  mnuThumbnail.Checked = True
  mnuOnePage.Checked = False
  mnuPageThumbnail.Checked = False

End Sub

Private Sub mnuToolbar_Click()
  'Add your code here.
  MsgBox "Function to be implemented by user."

End Sub


Private Sub mnuTypedText_Click()
  'see documentation for the list of annotation types
  oleImgEdit1.SelectTool AnnoText


  'Check the current annotation tool and uncheck all
  'the others
  mnuNoTool.Checked = False
  mnuSelectPointer.Checked = False
  mnuAttachNote.Checked = False
  mnuFillRect.Checked = False
  mnuFreeHand.Checked = False
  mnuHiLight.Checked = False
  mnuHollowRect.Checked = False
  mnuStamp.Checked = False
  mnuStraightLine.Checked = False
  mnuTextFromFile.Checked = False
  mnuTypedText.Checked = True

End Sub

Private Sub mnuZoomIn_Click()
  'Double the size of the image view
  Dim zoomval As Single 'zoom value

  zoomval = oleImgEdit1.Zoom
  zoomval = zoomval * 2
  If zoomval < ZoomMax Then
    oleImgEdit1.Zoom = zoomval
    oleImgEdit1.Refresh
  Else
    MsgBox "At maximum zoom"
  End If

  'uncheck the zoom menu picks.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub

Private Sub mnuZoomOut_Click()
  'Reduce the image by half
  Dim zoomval As Single 'zoom value

  zoomval = oleImgEdit1.Zoom
  zoomval = zoomval / 2
  If zoomval >= ZoomMin Then
    oleImgEdit1.Zoom = zoomval
    oleImgEdit1.Refresh
  Else
    MsgBox "At minimum zoom"
  End If


  'uncheck the zoom menu picks.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub

Private Sub mnuZoomToSelection_Click()
  'Zoom the part of the image in selection
  'rect to the size of the image window
  If Selection = True Then
    oleImgEdit1.ZoomToSelection
  End If

  'check the current menu pick and uncheck the others.
  mnuBestFit.Checked = False
  mnuFitHeight.Checked = False
  mnuFitWidth.Checked = False
  mnuActual.Checked = False
  mnu25.Checked = False
  mnu50.Checked = False
  mnu75.Checked = False
  mnu100.Checked = False
  mnu200.Checked = False
  mnu400.Checked = False
End Sub

Private Sub oleImgEdit1_SelectionRectDrawn(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
  'Determine if a selection rect has been drawn
  If Width = 0 And Height = 0 Then
    Selection = False
  Else
    Selection = True
  End If
End Sub

Private Sub oleImgEdit1_ToolPaletteHidden(ByVal Left As Long, ByVal Top As Long)
  'The tool palette has been hidden.  Uncheck its menu item.
  Annot8Visible = False
  mnuShowTools.Checked = False

End Sub


Private Sub oleImgThumbnail1_Click(ByVal ThumbNumber As Long)
  'Change the displayed page to the one represented by
  'the thumbnail that the user clicked on
  If ThumbNumber > 0 Then
    frmSample.oleImgEdit1.page = ThumbNumber
    frmSample.oleImgEdit1.Display
    frmSample.oleImgThumbnail1.DeselectAllThumbs
    frmSample.oleImgThumbnail1.ThumbSelected(ThumbNumber) = True
  End If
End Sub






Public Function MoverImagen(Imagen As String) As Boolean
  MoverImagen = True
  Dim temp As String 'image name and path
        
  On Error GoTo LUIS 'handle errors ourselves incase of cancel
  oleImgAdmin1.Flags = 0 'clear Flags
  If oleImgEdit1.ImageModified = True Then
    If MsgBox(ImageChanged, vbYesNo) = vbYes Then
      mnuSave_Click
      If Err = ErrCancel Then '32755 = Cancel pressed
        Exit Function
      End If
    End If
  End If
  oleImgAdmin1.Filter = "All Image Files|*.tif;*.bmp;*.jpg;*.pcx;*.dcx|TIFF files (*.tif)|*.tif|BMP files (*.bmp)|*.bmp|PCX/DCX Document (*.pcx, *.dcx)|*.pcx;*.dcx|JPG File (*.jpg)|*.jpg|All Files (*.*)|*.*|"
  Rem oleImgAdmin1.ShowFileDialog 0, frmSample.hWnd
  If Err = ErrCancel Then '32755 = Cancel pressed
    Exit Function
  End If
  If oleImgAdmin1.StatusCode <> 0 Then
    MsgBox Err.Description + " Code = " + Hex(oleImgAdmin1.StatusCode), 16
    Exit Function
  End If
 Rem temp = "\\Server1basa\carpeta segura\BACKUPLUIS" & Mid(Imagen, 3, 30)
Rem  temp = "D\BACKUPLUIS" & Mid(Imagen, 3, 30)
  
    temp = "\\Sistemas10\F" & Mid(Imagen, 3, 30)
  oleImgEdit1.Zoom = 38
Noimgen:
  oleImgEdit1.Image = temp
  oleImgThumbnail1.Image = oleImgEdit1.Image
  If numbits > 8 Then 'video driver supports hicolor or truecolor
    oleImgEdit1.ImagePalette = 3 'Set for 24 bit RGB.
  End If
  oleImgEdit1.page = 1
  oleImgEdit1.Display
  TotalPages = oleImgEdit1.PageCount
  oleImgThumbnail1.ThumbSelected(1) = True
        
  'Now that we have an image, enable the needed menus.
  mnuSaveAs.Enabled = True
  mnuSave.Enabled = True
  mnuPrint.Enabled = True
  mnuEdit.Enabled = True
  mnuView.Enabled = True
  mnuPage.Enabled = True
  mnuZoom.Enabled = True
  mnuAnnotation.Enabled = True
  If oleImgEdit1.PageCount > 1 Then
    mnuBack.Enabled = True
    mnuFirst.Enabled = True
    mnuGoTo.Enabled = True
    mnuLast.Enabled = True
    mnuNext.Enabled = True
    mnuPrevious.Enabled = True
  Else
    mnuBack.Enabled = False
    mnuFirst.Enabled = False
    mnuGoTo.Enabled = False
    mnuLast.Enabled = False
    mnuNext.Enabled = False
    mnuPrevious.Enabled = False
  End If
  Exit Function
LUIS:
  If Err.Number = 53 Then
    
    temp = "C:\noimagen.tif"
    GoTo Noimgen
    MoverImagen = False
  End If

End Function

Public Sub MoverAdelante()
  If rsente Is Nothing Then
    MsgBox "ULTIMO REGISTRO"
    Exit Sub
  End If
  rsente.Update
  rsente.MoveNext
  If rsente.EOF Then
    MsgBox "ULTIMO REGISTRO"
    rsente.MovePrevious
    Exit Sub
  End If
        
  cantidadRegistro = cantidadRegistro + 1
  lblCantidad = cantidadRegistro
  ColocarDatos
       
    
End Sub


Public Function N(DATO) As String
  If IsNull(DATO) Then
    N = ""
  Else
    N = DATO
  End If

End Function

Public Sub MoverAtrar()
        
  If rsente Is Nothing Then
    MsgBox "ULTIMO REGISTRO"
    Exit Sub
  End If
  rsente.Update
  rsente.MovePrevious
  If rsente.BOF Then
    MsgBox "ULTIMO REGISTRO"
    rsente.MoveNext
    Exit Sub
  End If
  cantidadRegistro = cantidadRegistro - 1
  lblCantidad = cantidadRegistro
  ColocarDatos
End Sub


Public Function MoverIm()

End Function

Public Sub ColocarDatos()
  Dim P1 As Integer
  Dim P2 As Integer
  Dim P3 As Integer

  If MoverImagen(rsente!Suspense_File) Then
             
  Else
    MsgBox "No se encontro la imagen"
    Exit Sub
  End If
  Rem lblCantidad =
        
  Nro_Cliente_Documento = N(rsente!Nro_Cliente_Documento)
  Nro_Operacion = N(rsente!Nro_Operacion)
  Numero_de_Registro = N(rsente!Numero_de_Registro)
  Prev.Value = False
  Mza.Value = False
        
  If Trim(rsente!Banco) = "Mendoza" Then
    Mza.Value = True
  End If
            
  If Trim(rsente!Banco) = "Prevision" Then
    Prev.Value = True
  End If
  If IsNull(rsente!DATA) Then
          
    DATA = ""
  Else
    DATA = rsente!DATA
  End If
  CajaNumero_Letra = N(rsente!CajaNumero_Letra)
  CajaNumero_Numero = N(rsente!CajaNumero_Numero)
  Sucursal = IIf(IsNull(rsente!Sucursal), "", rsente!Sucursal)
        
    If rsente!borrado = True Then
        optBorradosi.Value = True
    Else
        optBorradoNO.Value = True
    End If
    If rsente!NoestaDB Then
        NoestaDB.Value = 1
    Else
        NoestaDB.Value = 0
    End If
        
  DNI.Value = False
  LC.Value = False
  LE.Value = False
  OTRO.Value = False
  If rsente!DOC_de_Identidad = "D.N.I." Then
    DNI.Value = True
  End If
  If rsente!DOC_de_Identidad = "L.C." Then
    LC.Value = True
  End If
  If rsente!DOC_de_Identidad = "L.E.  " Then
    LE.Value = True
  End If
  If rsente!DOC_de_Identidad = "OTRO  " Then
    OTRO.Value = True
  End If
  lblError = rsente!descripcionerror
            
  If IsNull(rsente!Tipo_Operacion) Then
            
    Tipo_Operacion = ""
  Else
    Tipo_Operacion = rsente!Tipo_Operacion
  End If
        
        
  lblbachk = rsente!CSID
  
   If rsente!LetraCaja = True Then
        LetraCaja.Value = True
    Else
        NO.Value = True
    End If
  
        
  txtApellidoNombre(0).Text = ""
  txtApellidoNombre(1).Text = ""
  txtApellidoNombre(2).Text = ""
        
  txtApellidoNombre(0).Text = IIf(IsNull(rsente!Apellido_Garantia_1), "", rsente!Apellido_Garantia_1)
  txtApellidoNombre(1).Text = IIf(IsNull(rsente!Apellido_Garantia_2), "", rsente!Apellido_Garantia_2)
  txtApellidoNombre(2).Text = IIf(IsNull(rsente!Apellido_Garantia_3), "", rsente!Apellido_Garantia_3)
        
  VALORCHK 0, rsente!Tipo_Garantia_1
  VALORCHK 1, rsente!Tipo_Garantia_2
  VALORCHK 2, rsente!Tipo_Garantia_3
            
  txtNumero(0) = ""
  txtNumero(1) = ""
  txtNumero(2) = ""
            
  txtNumero(0) = IIf(IsNull(rsente!Documento_1), "", rsente!Documento_1)
  txtNumero(1) = IIf(IsNull(rsente!Documento_2), "", rsente!Documento_2)
  txtNumero(2) = IIf(IsNull(rsente!Documento_3), "", rsente!Documento_3)
       
  ValorTipoDocumento 0, rsente!TIPO_Doc__1
  ValorTipoDocumento 1, rsente!TIPO_Doc_2
  ValorTipoDocumento 2, rsente!TIPO_Doc__3




End Sub

Public Sub VALORCHK(indice As Integer, DATO As Variant)
  Dim P1 As Integer
  Dim P2 As Integer
  Dim P3 As Integer
  chkAval(indice).Value = False
  chkFianza(indice).Value = False
  chkHipoteca(indice).Value = False
  chkOtra(indice).Value = False
  chkPrenda(indice).Value = False
        
  If IsNull(DATO) Then
          
    Exit Sub
  End If
        
  chkAval(indice).Value = False
  chkFianza(indice).Value = False
  chkHipoteca(indice).Value = False
  chkOtra(indice).Value = False
        
        
  P1 = 0
  P2 = 0
  P3 = 0
  P1 = Len(DATO)
            
  For i = 1 To Len(DATO)
    If Asc(Mid(DATO, i, 1)) = 32 Then
      If P1 <> 0 Then
        P1 = i
        P2 = Len(DATO) - (P1 - 1)
      Else
                   
        If P2 <> 0 Then
          P2 = i
        Else
          P3 = i
        End If
      End If
    End If
  Next
        
        
        
        
  If P1 <> 0 Then
    Select Case Trim(Mid(DATO, 1, P1))
      Case chkAval(indice).Caption
        chkAval(indice).Value = 1
      Case chkFianza(indice).Caption
        chkFianza(indice).Value = 1
      Case chkHipoteca(indice).Caption
        chkHipoteca(indice).Value = 1
      Case chkOtra(indice).Caption
        chkOtra(indice).Value = 1
      Case chkPrenda(indice).Caption
        chkPrenda(indice).Value = 1
    End Select
            
  End If
        
  If P2 <> 0 Then
    Select Case Trim(Mid(DATO, P1, P2))
      Case chkAval(indice).Caption
        chkAval(indice).Value = 1
      Case chkFianza(indice).Caption
        chkFianza(indice).Value = 1
      Case chkHipoteca(indice).Caption
        chkHipoteca(indice).Value = 1
      Case chkOtra(indice).Caption
        chkOtra(indice).Value = 1
      Case chkPrenda(indice).Caption
        chkPrenda(indice).Value = 1
    End Select
        
        
  End If
        
  If P3 <> 0 Then
    Select Case Trim(Mid(DATO, P2, P3))
      Case chkAval(indice).Caption
        chkAval(indice).Value = 1
      Case chkFianza(indice).Caption
        chkFianza(indice).Value = 1
      Case chkHipoteca(indice).Caption
        chkHipoteca(indice).Value = 1
      Case chkOtra(indice).Caption
        chkOtra(indice).Value = 1
      Case chkPrenda(indice).Caption
        chkPrenda(indice).Value = 1
    End Select
  End If
        
        
        
        
End Sub

Public Sub ValorTipoDocumento(indice As Integer, DATO As Variant)
  OPTCI(indice).Value = False
  optDNI(indice).Value = False
  OPTLC(indice).Value = False
  OPTLE(indice).Value = False
  If IsNull(DATO) Then
    Exit Sub
  End If
 
 
  Select Case Trim(DATO)
    Case "D.N.I."
      optDNI(indice).Value = True
    Case "C.I."
      OPTCI(indice).Value = True
    Case "L.C."
      OPTLC(indice).Value = True
    Case "L.E."
      OPTLE(indice).Value = True
  End Select
  
 
 
 
 



End Sub


Public Function Actualizar() As Boolean
  MousePointer = 11
  Dim TipoGarantia As String
        
        
  If Nro_Cliente_Documento <> "" Then
    rsente!Nro_Cliente_Documento = Nro_Cliente_Documento
  End If
  If Nro_Operacion <> "" Then
    rsente!Nro_Operacion = Nro_Operacion
  End If
  If Numero_de_Registro <> "" Then
    rsente!Numero_de_Registro = Numero_de_Registro
  End If

  If Prev.Value Then
    rsente!Banco = "Prevision"
  End If
  If Mza.Value Then
    rsente!Banco = "Mendoza "
  End If

  If Trim(DATA) <> "" Then
    rsente!DATA = DATA
  End If
       
  rsente!CajaNumero_Letra = CajaNumero_Letra
  If CajaNumero_Numero <> "" Then
    rsente!CajaNumero_Numero = CajaNumero_Numero
  End If
  If Sucursal <> "" Then
       
    rsente!Sucursal = Sucursal
  Else
    MousePointer = 0
    MsgBox " SUCURSAL"
  End If
rsente!LetraCaja = LetraCaja.Value
rsente!borrado = optBorradosi.Value
    rsente!NoestaDB = NoestaDB.Value
  
If DNI Then
    rsente!DOC_de_Identidad = "D.N.I."
End If

If LC Then
     rsente!DOC_de_Identidad = "L.C."
End If

If LE Then
    rsente!DOC_de_Identidad = "L.E.  "
End If

If OTRO Then
    rsente!DOC_de_Identidad = "OTRO  "
End If


  If optDNI(0).Value = True Then
    rsente!TIPO_Doc__1 = "D.N.I."
  End If

  If OPTLC(0).Value = True Then
    rsente!TIPO_Doc__1 = "L.C."
  End If
  If OPTLE(0).Value = True Then
    rsente!TIPO_Doc__1 = "L.E."
  End If
  If OPTCI(0).Value = True Then
    rsente!TIPO_Doc__1 = "C.I."
  End If
          
          
  ' ---------------------
       
       
  If optDNI(1).Value = True Then
    rsente!TIPO_Doc_2 = "D.N.I."
  End If

  If OPTLC(1).Value = True Then
    rsente!TIPO_Doc_2 = "L.C."
  End If
  If OPTLE(1).Value = True Then
    rsente!TIPO_Doc_2 = "L.E."
  End If
  If OPTCI(1).Value = True Then
    rsente!TIPO_Doc_2 = "C.I."
  End If
        
  ' ------------------------
        
  If optDNI(2).Value = True Then
    rsente!TIPO_Doc__3 = "D.N.I."
  End If

  If OPTLC(2).Value = True Then
    rsente!TIPO_Doc__3 = "L.C."
  End If
  If OPTLE(2).Value = True Then
    rsente!TIPO_Doc__3 = "L.E."
  End If
  If OPTCI(2).Value = True Then
    rsente!TIPO_Doc__3 = "C.I."
  End If
          
          
  rsente!Tipo_Operacion = Trim(Tipo_Operacion)

            

  rsente!Apellido_Garantia_1 = Trim(txtApellidoNombre(0).Text)
  rsente!Apellido_Garantia_2 = Trim(txtApellidoNombre(1).Text)
  rsente!Apellido_Garantia_3 = Trim(txtApellidoNombre(2).Text)
        
  TipoGarantia = ""
        
  If chkAval(0).Value = 1 Then
    TipoGarantia = "Aval"
  End If
        
  If chkFianza(0).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Fianza"
  End If
        
  If chkHipoteca(0).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Hipoteca"
  End If
        
  If chkOtra(0).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Otra"
  End If
  If chkPrenda(0).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Prenda"
  End If
        
  rsente!Tipo_Garantia_1 = Trim(TipoGarantia)
        
  Rem --------------------
  TipoGarantia = ""
        
  If chkAval(1).Value = 1 Then
    TipoGarantia = "Aval"
  End If
        
  If chkFianza(1).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Fianza"
  End If
        
  If chkHipoteca(1).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Hipoteca"
  End If
        
  If chkOtra(1).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Otra"
  End If
  If chkPrenda(1).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Prenda"
  End If
        
  rsente!Tipo_Garantia_2 = Trim(TipoGarantia)
        
  Rem ------------------------
  TipoGarantia = ""
        
  If chkAval(2).Value = 1 Then
    TipoGarantia = "Aval"
  End If
        
  If chkFianza(2).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Fianza"
  End If
        
  If chkHipoteca(2).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Hipoteca"
  End If
        
  If chkOtra(2).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Otra"
  End If
  If chkPrenda(2).Value = 1 Then
    TipoGarantia = TipoGarantia & " " & "Prenda"
  End If
        
  rsente!Tipo_Garantia_3 = Trim(TipoGarantia)
        
        
        
        
        
        
            
  rsente!Documento_1 = txtNumero(0)
  rsente!Documento_2 = txtNumero(1)
  rsente!Documento_3 = txtNumero(2)
        
  rsente.Update
  MousePointer = 0
End Function



Public Function ll()

End Function

