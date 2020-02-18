VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndices 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indices"
   ClientHeight    =   10140
   ClientLeft      =   720
   ClientTop       =   1050
   ClientWidth     =   12960
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   12960
   Begin TabDlg.SSTab SSTab1 
      Height          =   5115
      Left            =   120
      TabIndex        =   16
      Top             =   4140
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9022
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
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
      TabCaption(0)   =   "Campos"
      TabPicture(0)   =   "frmIndiceClientes.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Costos"
      TabPicture(1)   =   "frmIndiceClientes.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblPrecioPreparacion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label8"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtCostoPreparacion"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtCostoDigitalizacion"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtCostoIndexacion"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtCostoArmado"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtCostoCargaLegajo"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Digitalizacion"
      TabPicture(2)   =   "frmIndiceClientes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame6"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmIndiceClientes.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label21"
      Tab(3).Control(1)=   "Frame7"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame6 
         Caption         =   "Archivo Salida"
         Height          =   3795
         Left            =   -66300
         TabIndex        =   53
         Top             =   960
         Width           =   2775
         Begin VB.ComboBox cboTipoArchivoExtencion 
            Height          =   345
            ItemData        =   "frmIndiceClientes.frx":0070
            Left            =   1620
            List            =   "frmIndiceClientes.frx":007A
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   3180
            Width           =   795
         End
         Begin VB.TextBox txtOrdenArchivoSalida 
            Height          =   330
            Index           =   6
            Left            =   1620
            TabIndex        =   66
            Tag             =   "Etiqueta"
            Text            =   "0"
            Top             =   2760
            Width           =   735
         End
         Begin VB.TextBox txtOrdenArchivoSalida 
            Height          =   330
            Index           =   5
            Left            =   1620
            TabIndex        =   65
            Tag             =   "Letra_Hasta"
            Text            =   "0"
            Top             =   2340
            Width           =   735
         End
         Begin VB.TextBox txtOrdenArchivoSalida 
            Height          =   330
            Index           =   4
            Left            =   1620
            TabIndex        =   63
            Tag             =   "Letra_desde"
            Text            =   "0"
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtOrdenArchivoSalida 
            Height          =   330
            Index           =   3
            Left            =   1620
            TabIndex        =   61
            Tag             =   "Nro_Hasta"
            Text            =   "0"
            Top             =   1500
            Width           =   735
         End
         Begin VB.TextBox txtOrdenArchivoSalida 
            Height          =   330
            Index           =   2
            Left            =   1620
            TabIndex        =   59
            Tag             =   "Nro_Desde"
            Text            =   "0"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtOrdenArchivoSalida 
            Height          =   330
            Index           =   1
            Left            =   1620
            TabIndex        =   57
            Tag             =   "Nro_desde"
            Text            =   "0"
            Top             =   660
            Width           =   735
         End
         Begin VB.TextBox txtOrdenArchivoSalida 
            Height          =   330
            Index           =   0
            Left            =   1620
            TabIndex        =   55
            Tag             =   "ID_Imagen"
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Tipo Archivo"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   3240
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Etiqueta"
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   67
            Top             =   2820
            Width           =   915
         End
         Begin VB.Label Label18 
            Caption         =   "Nro_Hasta"
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   64
            Top             =   2340
            Width           =   915
         End
         Begin VB.Label Label18 
            Caption         =   "Letra _Desde"
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   62
            Top             =   1920
            Width           =   1035
         End
         Begin VB.Label Label18 
            Caption         =   "Nro_Hasta"
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   60
            Top             =   1500
            Width           =   915
         End
         Begin VB.Label Label18 
            Caption         =   "Nro_Desde"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   58
            Top             =   1080
            Width           =   915
         End
         Begin VB.Label Label18 
            Caption         =   "Nro_Hasta"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   660
            Width           =   915
         End
         Begin VB.Label Label18 
            Caption         =   "ID Imagen"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Archivo Entrada"
         Height          =   1455
         Left            =   -72600
         TabIndex        =   50
         Top             =   960
         Width           =   5115
         Begin VB.ComboBox Combo1 
            Height          =   345
            Left            =   2700
            TabIndex        =   71
            Text            =   "Combo1"
            Top             =   900
            Width           =   2175
         End
         Begin VB.TextBox txtCantidad_Imagenes_Max 
            Height          =   315
            Left            =   2700
            TabIndex        =   52
            Text            =   "1"
            Top             =   480
            Width           =   915
         End
         Begin VB.Label Label20 
            Caption         =   "Nombre de perfil digitalizar "
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   900
            Width           =   2295
         End
         Begin VB.Label Label12 
            Caption         =   "Cant Imagenes Max :"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   420
            Width           =   1755
         End
      End
      Begin VB.TextBox txtCostoCargaLegajo 
         Height          =   375
         Left            =   5880
         TabIndex        =   37
         Text            =   "0"
         Top             =   1020
         Width           =   1215
      End
      Begin VB.TextBox txtCostoArmado 
         Height          =   375
         Left            =   2280
         TabIndex        =   35
         Text            =   "0"
         Top             =   2460
         Width           =   1215
      End
      Begin VB.TextBox txtCostoIndexacion 
         Height          =   375
         Left            =   2280
         TabIndex        =   33
         Text            =   "0"
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtCostoDigitalizacion 
         Height          =   375
         Left            =   2280
         TabIndex        =   31
         Text            =   "0"
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox txtCostoPreparacion 
         Height          =   375
         Left            =   2220
         TabIndex        =   29
         Text            =   "0"
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4395
         Left            =   -74700
         TabIndex        =   17
         Top             =   600
         Width           =   12015
         Begin VB.Frame Frame10 
            Caption         =   "Titulo"
            Height          =   4155
            Left            =   2580
            TabIndex        =   95
            Top             =   240
            Width           =   3135
            Begin VB.TextBox txt_Titulo_Etiqueta_Legajo 
               Height          =   330
               Left            =   180
               TabIndex        =   103
               Top             =   3720
               Width           =   2775
            End
            Begin VB.TextBox txt_Titulo_Descripcion 
               Height          =   330
               Left            =   180
               TabIndex        =   102
               Top             =   3300
               Width           =   2775
            End
            Begin VB.TextBox txt_Titulo_Letra_Hasta 
               Height          =   330
               Left            =   180
               TabIndex        =   101
               Top             =   2880
               Width           =   2775
            End
            Begin VB.TextBox txt_Titulo_Letra_Desde 
               Height          =   330
               Left            =   180
               TabIndex        =   100
               Top             =   2460
               Width           =   2775
            End
            Begin VB.TextBox txt_Titulo_Nro_Hasta 
               Height          =   330
               Left            =   180
               TabIndex        =   99
               Top             =   2040
               Width           =   2775
            End
            Begin VB.TextBox txt_Titulo_Nro_Desde 
               Height          =   330
               Left            =   180
               TabIndex        =   98
               Top             =   1620
               Width           =   2775
            End
            Begin VB.TextBox txt_Titulo_Fecha_Hasta 
               Height          =   330
               Left            =   180
               TabIndex        =   97
               Top             =   1200
               Width           =   2775
            End
            Begin VB.TextBox txt_Titulo_Fecha_Desde 
               Height          =   330
               Left            =   180
               TabIndex        =   96
               Top             =   780
               Width           =   2775
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Barra"
            Height          =   4095
            Left            =   6000
            TabIndex        =   88
            Top             =   180
            Width           =   1755
            Begin VB.OptionButton optBarra_Ninguno 
               Caption         =   "Ninguno"
               Height          =   375
               Left            =   60
               TabIndex        =   94
               Top             =   1080
               Width           =   1455
            End
            Begin VB.OptionButton optBarra_NRO_DESDE 
               Caption         =   "NroDesde"
               Height          =   375
               Left            =   60
               TabIndex        =   93
               Top             =   1560
               Width           =   1455
            End
            Begin VB.OptionButton optBarra_NRO_HASTA 
               Caption         =   "NroHasta"
               Height          =   375
               Left            =   60
               TabIndex        =   92
               Top             =   1980
               Width           =   1455
            End
            Begin VB.OptionButton optBarra_Letra_Desde 
               Caption         =   "Letra desde"
               Height          =   375
               Left            =   60
               TabIndex        =   91
               Top             =   2400
               Width           =   1455
            End
            Begin VB.OptionButton optBarra_Letra_Hasta 
               Caption         =   "Letra desde"
               Height          =   375
               Left            =   60
               TabIndex        =   90
               Top             =   2820
               Width           =   1455
            End
            Begin VB.OptionButton optEtiqueta_Legajo 
               Caption         =   "Etiqueta Legajo"
               Height          =   225
               Left            =   60
               TabIndex        =   89
               Top             =   3720
               Width           =   1575
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Frame9"
            Height          =   3675
            Left            =   1860
            TabIndex        =   79
            Top             =   360
            Width           =   315
            Begin VB.CheckBox chkTodosRequerir 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   87
               Top             =   420
               Width           =   195
            End
            Begin VB.CheckBox chk_Requerir_Descripcion 
               Caption         =   "Check1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   86
               Top             =   3300
               Width           =   195
            End
            Begin VB.CheckBox chk_Requerir_Letra_Hasta 
               Caption         =   "Check1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   85
               Top             =   2880
               Width           =   195
            End
            Begin VB.CheckBox chk_Requerir_Letra_Desde 
               Caption         =   "Check1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   84
               Top             =   2460
               Width           =   195
            End
            Begin VB.CheckBox chk_Requerir_Nro_Hasta 
               Caption         =   "Check1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   83
               Top             =   2040
               Width           =   195
            End
            Begin VB.CheckBox chk_Requerir_Nro_Desde 
               Caption         =   "Check1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   82
               Top             =   1620
               Width           =   195
            End
            Begin VB.CheckBox chk_Requerir_Fecha_Hasta 
               Caption         =   "Check1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   81
               Top             =   1200
               Width           =   195
            End
            Begin VB.CheckBox chk_Requerir_Fecha_Desde 
               Caption         =   "Check1"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   80
               Top             =   780
               Width           =   195
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Copiar"
            Height          =   3015
            Left            =   2220
            TabIndex        =   75
            Top             =   420
            Width           =   315
            Begin VB.CheckBox chk_Copiar_Letra 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   78
               Top             =   2580
               Width           =   195
            End
            Begin VB.CheckBox chk_Copiar_Nro 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   77
               Top             =   1800
               Width           =   195
            End
            Begin VB.CheckBox chk_Copiar_Fecha 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   60
               TabIndex        =   76
               Top             =   900
               Width           =   195
            End
         End
         Begin VB.CheckBox chkEtiquetaLegajo 
            Alignment       =   1  'Right Justify
            Caption         =   "Etiqueta Legajo"
            Height          =   375
            Left            =   60
            TabIndex        =   72
            Top             =   3960
            Width           =   1695
         End
         Begin VB.Frame Frame2 
            Caption         =   "Controles Logicos"
            Height          =   4035
            Left            =   7980
            TabIndex        =   38
            Top             =   360
            Width           =   1695
            Begin VB.Frame Frame3 
               Caption         =   "Largo"
               Height          =   3255
               Left            =   120
               TabIndex        =   39
               Top             =   300
               Width           =   1455
               Begin VB.TextBox TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA 
                  Height          =   315
                  Left            =   780
                  TabIndex        =   49
                  Text            =   "0"
                  Top             =   2340
                  Width           =   555
               End
               Begin VB.TextBox TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   48
                  Text            =   "0"
                  Top             =   2340
                  Width           =   555
               End
               Begin VB.TextBox TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA 
                  Height          =   315
                  Left            =   780
                  TabIndex        =   47
                  Text            =   "0"
                  Top             =   1980
                  Width           =   555
               End
               Begin VB.TextBox TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   46
                  Text            =   "0"
                  Top             =   1980
                  Width           =   555
               End
               Begin VB.TextBox TXTCONTROL_LOGICO_LARGO_NRO_HASTA_HASTA 
                  Height          =   315
                  Left            =   780
                  TabIndex        =   45
                  Text            =   "0"
                  Top             =   1560
                  Width           =   555
               End
               Begin VB.TextBox TXTCONTROL_LOGICO_LARGO_NRO_HASTA_INICIO 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   44
                  Text            =   "0"
                  Top             =   1560
                  Width           =   555
               End
               Begin VB.TextBox txtCONTROL_LOGICO_LARGO_NRO_DESDE_HASTA 
                  Height          =   315
                  Left            =   780
                  TabIndex        =   41
                  Text            =   "0"
                  Top             =   1140
                  Width           =   555
               End
               Begin VB.TextBox txtCONTROL_LOGICO_LARGO_NRO_DESDE_INICIO 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   40
                  Text            =   "0"
                  Top             =   1140
                  Width           =   555
               End
               Begin VB.Label Label17 
                  Caption         =   "Hasta"
                  Height          =   375
                  Left            =   780
                  TabIndex        =   43
                  Top             =   480
                  Width           =   555
               End
               Begin VB.Label Label16 
                  Caption         =   "Desde"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   42
                  Top             =   480
                  Width           =   615
               End
            End
         End
         Begin VB.CheckBox chkLetra_Hasta 
            Alignment       =   1  'Right Justify
            Caption         =   "Letra Hasta"
            Height          =   315
            Left            =   60
            TabIndex        =   25
            Top             =   3180
            Width           =   1695
         End
         Begin VB.CheckBox chkLetra_Desde 
            Alignment       =   1  'Right Justify
            Caption         =   "Letra Desde"
            Height          =   315
            Left            =   60
            TabIndex        =   24
            Top             =   2760
            Width           =   1695
         End
         Begin VB.CheckBox chkNro_Hasta 
            Alignment       =   1  'Right Justify
            Caption         =   "Nº Hasta"
            Height          =   315
            Left            =   60
            TabIndex        =   23
            Top             =   2340
            Width           =   1695
         End
         Begin VB.CheckBox chkNro_Desde 
            Alignment       =   1  'Right Justify
            Caption         =   "Nº Desde"
            Height          =   315
            Left            =   60
            TabIndex        =   22
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CheckBox chkFecha_Hasta 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha Hasta"
            Height          =   315
            Left            =   60
            TabIndex        =   21
            Top             =   1500
            Width           =   1695
         End
         Begin VB.CheckBox chkFecha_Desde 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha Desde"
            Height          =   315
            Left            =   60
            TabIndex        =   20
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CheckBox chkHabilitar_descripcion 
            Alignment       =   1  'Right Justify
            Caption         =   "Descripción"
            Height          =   435
            Left            =   60
            TabIndex        =   19
            Top             =   3540
            Width           =   1695
         End
         Begin VB.CheckBox chkTodos 
            Alignment       =   1  'Right Justify
            Caption         =   "Todos"
            Height          =   255
            Left            =   60
            TabIndex        =   18
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "Campo"
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
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Ver"
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
            Left            =   1500
            TabIndex        =   26
            Top             =   360
            Width           =   435
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Referencias"
         Height          =   3615
         Left            =   -73320
         TabIndex        =   74
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label Label21 
         Caption         =   "Letra_Desde:"
         Height          =   315
         Left            =   -74700
         TabIndex        =   73
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Costo Carga de Legajo"
         Height          =   375
         Left            =   3960
         TabIndex        =   36
         Top             =   1020
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Costo Rearmado"
         Height          =   375
         Left            =   360
         TabIndex        =   34
         Top             =   2460
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Costo Indexsacion"
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   1980
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Costo Digitalizacion"
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   1500
         Width           =   1935
      End
      Begin VB.Label lblPrecioPreparacion 
         Caption         =   "Costo Preparacion"
         Height          =   375
         Left            =   360
         TabIndex        =   28
         Top             =   1020
         Width           =   1575
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
      Height          =   315
      Left            =   10560
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtNro_Documento 
      Enabled         =   0   'False
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
      Left            =   9300
      TabIndex        =   13
      Top             =   1500
      Width           =   795
   End
   Begin VB.TextBox txt_ID 
      Enabled         =   0   'False
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
      Left            =   7260
      TabIndex        =   11
      Top             =   1500
      Width           =   675
   End
   Begin VB.ComboBox cboTipo 
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
      ItemData        =   "frmIndiceClientes.frx":0088
      Left            =   1020
      List            =   "frmIndiceClientes.frx":0098
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1500
      Width           =   1635
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   8
      Top             =   9825
      Width           =   12960
      _ExtentX        =   22860
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
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
      Height          =   375
      Left            =   1020
      TabIndex        =   6
      Top             =   1080
      Width           =   9075
   End
   Begin VB.TextBox txtIndice 
      Enabled         =   0   'False
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
      Left            =   3660
      TabIndex        =   4
      Top             =   1500
      Width           =   2295
   End
   Begin Controles.cltIndice ctlIndiceCliente 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   3836
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   -60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   54
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":00C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":033C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":06FA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":0B34
            Key             =   "Borrar1"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":0F37
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":135E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":1705
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":1AC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":1E89
            Key             =   "Salvar2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":2105
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":238B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":2750
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":2B0D
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":2D8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":2FFF
            Key             =   "Casa"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":33D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":37AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":3A27
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":3BF8
            Key             =   "Atras2"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":3FB2
            Key             =   "Inicio"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":4097
            Key             =   "Fin"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":4179
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":4547
            Key             =   "Adelante2"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":48FB
            Key             =   "Correo2"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":4CEE
            Key             =   "Bandera"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":4F51
            Key             =   "trvt2"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":5312
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":5BEC
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":5E8C
            Key             =   "Cancelar1"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":61A6
            Key             =   "Aceptar1"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":64C0
            Key             =   "trvt"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":6596
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":666C
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":6A93
            Key             =   "Atras3"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":718D
            Key             =   "Atras"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":7887
            Key             =   "Adelante"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":7F81
            Key             =   "Correo3"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":867B
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":8D75
            Key             =   "Correo4"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":910F
            Key             =   "Correo"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":9DE9
            Key             =   "Borrar"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":A6C3
            Key             =   "Punto"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":AF9D
            Key             =   "Cancelar2"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":B36B
            Key             =   "Aceptar2"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":B724
            Key             =   "Aceptar"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":BFFE
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":C8D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":D2EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":DCFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":E70E
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":F120
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":FB32
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":10544
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndiceClientes.frx":10F56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   630
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   1111
      ButtonWidth     =   1429
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            Key             =   "Aceptar"
            ImageIndex      =   47
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "Cancelar"
            ImageIndex      =   48
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actual"
            Key             =   "Actual"
            ImageIndex      =   50
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            Key             =   "Print"
            ImageIndex      =   49
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Titulos"
            Key             =   "Titulos"
            ImageIndex      =   53
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cliente T"
            Key             =   "ClienteT"
            ImageIndex      =   51
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "C. Sector"
            Key             =   "CantSector"
            ImageIndex      =   54
         EndProperty
      EndProperty
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   315
      Left            =   1020
      TabIndex        =   2
      Top             =   660
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   556
   End
   Begin VB.Label Label3 
      Caption         =   "Nº DOC:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8280
      TabIndex        =   14
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "ID_INDICE"
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
      Left            =   6120
      TabIndex        =   12
      Top             =   1560
      Width           =   1035
   End
   Begin VB.Label Label13 
      Caption         =   "TIPO"
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
      Left            =   180
      TabIndex        =   9
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "NOMBRE:"
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
      Left            =   180
      TabIndex        =   7
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "INDICE"
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
      Left            =   2820
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "CLIENTE:"
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
      Left            =   180
      TabIndex        =   3
      Top             =   720
      Width           =   795
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuNuevo 
         Caption         =   "Nuevo"
         Begin VB.Menu mnuCrearIgualNivel 
            Caption         =   "Crear Igual Nivel"
         End
         Begin VB.Menu mnuCrearHijo 
            Caption         =   "Crear Hijo"
         End
         Begin VB.Menu mnuIndiceInicial 
            Caption         =   "Indice Inicial"
         End
      End
      Begin VB.Menu mnuModificar 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnuAsignarNumeroDocumento 
         Caption         =   "Asignar Numero de documento"
      End
      Begin VB.Menu mnuBorrarIndice 
         Caption         =   "Borrar"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir Todo"
      End
      Begin VB.Menu mnuImprimirSeleccion 
         Caption         =   "Imprimir Seleccion"
      End
      Begin VB.Menu mnuExpander 
         Caption         =   "Expander"
      End
      Begin VB.Menu mnuSector 
         Caption         =   "Sector"
      End
      Begin VB.Menu mnuDocumentos 
         Caption         =   "Documentos"
      End
      Begin VB.Menu mnuDocumento 
         Caption         =   "Documento"
      End
      Begin VB.Menu mnuLegajo 
         Caption         =   "Legajo"
      End
      Begin VB.Menu mnuBuscarIndice 
         Caption         =   "Buscar Indice"
      End
      Begin VB.Menu mnuBuscarLegajo 
         Caption         =   "Buscar Legajos"
      End
      Begin VB.Menu mnuCopiarIndiceOrigen 
         Caption         =   "Copiar Indice Origen"
      End
      Begin VB.Menu mnuCopiarIndiceDestino 
         Caption         =   "Copiar Indice Destino"
      End
   End
End
Attribute VB_Name = "frmIndices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ExpanderIndex As Integer


Private Sub ctlIndice_PopupMenuAction()
    PopupMenu mnuMenu
End Sub


'Private Sub cmdAceptarCambioIndices_Click()
'
'Dim FK_Indice As Long
'Dim RS As New ADODB.Recordset
'Dim SQL As String
'Dim Descripcion_Anterior As String
'
'FK_Indice = Buscar_ID_Indice_Por_indice(lblDocumentoFinal, ctlCliente.Valor)
'
'
'
'
'SQL = " SELECT   ID_LEGAJO, COD_INDICE, FK_INDICES, COD_CLIENTE , INDICE_ANTERIOR"
'SQL = SQL & " From LEGAJOS"
'SQL = SQL & "  WHERE COD_INDICE = '" & lblDocumentoOrigen & "'"
'SQL = SQL & "  AND COD_CLIENTE = " & ctlCliente.Valor
'
'
'Set RS = New ADODB.Recordset
'RS.CursorLocation = adUseClient
'RS.Open SQL, ConActiva, 3, 2
'
'Do While RS.EOF
'    RS!Indice_Anterior = Trim(RS!Cod_Indice)
'    RS!Cod_Indice = Trim(lblDocumentoFinal.Caption)
'    RS!FK_INDICES = FK_Indice
'    Descripcion_Anterior = RS!Descripcion
'    If chkCopiarDescripcion.value = 1 Then
'        RS!Descripcion = lblDocumentoFinalDescripcion & "// " & Descripcion_Anterior
'    End If
'
'
'
'
'    RS.MoveNext
'Loop
'
'
'
'
''SELECT     COD_ID_REFERENCIA, COD_CLIENTE, INDICE, FK_INDICES, DESCRIPCION, INDICE_ANTERIOR
''From REFERENCIAS
''WHERE     (COD_CLIENTE = 20) AND (INDICE = '001')
''
''
''SELECT     COD_CLIENTE, INDICE, DESCRIPCION, ID
''From DOCUMENTOS_DIGITALES
''WHERE     (COD_CLIENTE = 15) AND (INDICE = N'001')
''
''SELECT     COD_CLIENTE, COD_INDICE, ID, DESCRIPCION
''From ORDENAR_DOCUMENTACION_DETALLE
''WHERE     (COD_CLIENTE = 15) AND (COD_INDICE = '001')
'
'
'
''Update CLIENTEUSUARIO
''SET              COD_INDICE =, COD_CLIENTE =
''WHERE     (COD_INDICE = '001') AND (COD_CLIENTE = 20)
'
'End Sub

'Private Sub Command1_Click()
'INSE
'End Sub
'
Private Sub Command10_Click()

End Sub

Private Sub chkTodos_Click()
chkHabilitar_descripcion.value = chkTodos.value
chkLetra_Desde.value = chkTodos.value
chkLetra_Hasta.value = chkTodos.value
chkNro_Desde.value = chkTodos.value
chkNro_Hasta.value = chkTodos.value
chkFecha_Desde.value = chkTodos.value
chkFecha_Hasta.value = chkTodos.value
End Sub

Private Sub chkTodosRequerir_Click()
chk_Requerir_Descripcion = chkTodosRequerir.value
chk_Requerir_Fecha_Desde = chkTodosRequerir.value
chk_Requerir_Fecha_Hasta = chkTodosRequerir.value
chk_Requerir_Letra_Desde = chkTodosRequerir.value
chk_Requerir_Letra_Hasta = chkTodosRequerir.value
chk_Requerir_Nro_Desde = chkTodosRequerir.value
chk_Requerir_Nro_Hasta = chkTodosRequerir.value
End Sub


Private Sub Command2_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String

Sql = "SELECT     ID, COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, DESCRIPCION, LEN(INDICE) AS Expr1"
Sql = Sql & " From basasql.dbo.INDICES"
Sql = Sql & " WHERE     (COD_CLIENTE = 1197) AND (LEN(INDICE) = 9) and INDICE like '001%'"


'sql = "SELECT     ID, COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, DESCRIPCION, LEN(INDICE) AS Expr1"
'sql = sql & "  From basasql.dbo.INDICES"
'sql = sql & "  WHERE     (COD_CLIENTE = 1197) AND (LEN(INDICE) = 9) and INDICE like '002%'"


rs.Open Sql, strConBasa

 Do While Not rs.EOF
Sql = "  INSERT INTO basasql.dbo.INDICES"
Sql = Sql & "                      (COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, DESCRIPCION, HABILITAR_FECHA_DESDE, HABILITAR_FECHA_HASTA, HABILITAR_LETRA_DESDE,"
Sql = Sql & "                      HABILITAR_LETRA_HASTA, HABILITAR_NRO_DESDE, HABILITAR_NRO_HASTA, HABILITAR_DESCRIPCION, REQUERIR_FECHA_HASTA, REQUERIR_FECHA_DESDE,"
 Sql = Sql & "                     TIPO_INDICE)"
Sql = Sql & " VALUES     (1197,"
Sql = Sql & "6" & Format(rs!ID_CODIGO_DOCUMENTO, "0000")
Sql = Sql & ", '" & Trim(rs!Indice) & "006'"
Sql = Sql & ",'GASTOS - SM " & CInt(rs!ID_CODIGO_DOCUMENTO) & "'"
Sql = Sql & ", 1, 1, 1, 1, 1, 1, 1, 1, 1, 'Documento')"

ExecutarSql Sql



    rs.MoveNext

Loop

MsgBox "terminado"

End Sub

Private Sub ctlCliente_Click()
ActualizarIndice
End Sub

Private Sub ctlIndiceCliente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuMenu
    End If

End Sub

Private Sub Form_Load()
    ctlCliente.TipoControl = Cliente
    ExpanderIndex = 0
End Sub

Private Sub mnuAsignarNumeroDocumento_Click()
    Dim rs As ADODB.Recordset
    Dim Sql As String
    
    Set rs = New ADODB.Recordset
    Dim Doc As Long
    
    
    rs.Open "SELECT COD_CLIENTE, INDICE From REFERENCIAS WHERE  COD_CLIENTE = " & ctlCliente.Valor & "  AND INDICE = '" & ctlIndiceCliente.Item_Selecionado & "%'", ConActiva, 0, 1
    If rs.EOF Then
        Doc = InputBox("Ingrese el numero de documento")
        Set rs = New ADODB.Recordset
        Sql = " SELECT     COD_CLIENTE, ID_CODIGO_DOCUMENTO"
        Sql = Sql & " From INDICES Where COD_CLIENTE =" & ctlCliente.Valor
        Sql = Sql & " And ID_CODIGO_DOCUMENTO = " & Doc
        rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
        
           Sql = "  Update INDICES"
           Sql = Sql & " SET    ID_CODIGO_DOCUMENTO =" & Doc
           Sql = Sql & " Where COD_CLIENTE =" & ctlCliente.Valor
           Sql = Sql & " AND INDICE = '" & ctlIndiceCliente.Item_Selecionado & "'"
            
           ExecutarSql Sql

             ctlIndiceCliente.Actualizar ctlCliente.Valor, Nulo, ExpanderIndex
            
        
        Else
            MsgBox "El documento ya existe"
        End If
        
        
        
    
    Else
    
    End If
    
    
 
End Sub

Private Sub mnuBorrarIndice_Click()
        Dim rsBorrar As New ADODB.Recordset
        Dim RScontrolLegajos As New ADODB.Recordset
        
        Dim Sql As String
       
        Dim ItemSelec As String
            
            ItemSelec = ctlIndiceCliente.Item_Selecionado
            Sql = " SELECT COD_CLIENTE, INDICE "
            Sql = Sql & " From REFERENCIAS "
            Sql = Sql & "  Where COD_CLIENTE = " & ctlCliente.Valor
            Sql = Sql & "  AND INDICE like '" & ItemSelec & "%'"
            
            
            
            rsBorrar.Open Sql, ConActiva, 0, 1
            
            
            
            
            If rsBorrar.EOF Then
            
                
                
                Sql = " SELECT     COD_INDICE, COD_CLIENTE"
               Sql = Sql & " From LEGAJOS"
               Sql = Sql & " WHERE     "
                Sql = Sql & " COD_INDICE  like '" & ItemSelec & "%'"
                Sql = Sql & " AND COD_CLIENTE  =  " & ctlCliente.Valor
                
                RScontrolLegajos.Open Sql, ConActiva, 0, 1
                
                If RScontrolLegajos.EOF Then
            
                Sql = " DELETE FROM INDICES"
                Sql = Sql & "  WHERE COD_CLIENTE = " & ctlCliente.Valor
                Sql = Sql & "  AND INDICE like '" & ItemSelec & "%'"
                ExecutarSql Sql
                MsgBox "Los Indices fueron borrados", vbInformation
            Else
                 MsgBox "Para estos Indices existen cajas activas", vbCritical
                 
            End If
            End If
            




            
            
           ActualizarIndice

End Sub

Private Sub mnuBuscarIndice_Click()
Dim a As String
a = InputBox("Ingrese la frase a Buscar")
ctlIndiceCliente.BuscarIndice a, True
End Sub

Private Sub mnuBuscarLegajo_Click()
ctlIndiceCliente.BuscarTipoIndice "Legajo", True
End Sub

'Private Sub mnuCopiarIndiceDestino_Click()
'    lblDocumentoFinal.Caption = ctlIndiceCliente.Item_Selecionado
'    lblDocumentoFinalDescripcion.Caption = ctlIndiceCliente.Descripcion
'End Sub
'
'Private Sub mnuCopiarIndiceOrigen_Click()
'    lblDocumentoOrigen.Caption = ctlIndiceCliente.Item_Selecionado
'    lblDocumentoOrigenDescripcion.Caption = ctlIndiceCliente.Descripcion
'End Sub
'
Private Sub mnuCrearHijo_Click()
Dim rs As New ADODB.Recordset
 Dim Sql As String
 Dim NUMERO As String
 Dim Contador As String
 Dim dif As String
 Dim ItemSele As String
 Dim Filtro  As String
 
 
 
 ItemSele = ctlIndiceCliente.Item_Selecionado
 If Len(ItemSele) = 3 Then
    Filtro = "  LIKE '" & ItemSele & "%'"
 Else
    Filtro = "  LIKE '" & Mid(ItemSele, 1, Len(ItemSele)) & "%'"
 End If
 
 
 Sql = "  SELECT COD_CLIENTE, INDICE From INDICES"
 Sql = Sql & vbCrLf & " Where COD_CLIENTE = " & ctlCliente.Valor
 Sql = Sql & vbCrLf & " AND INDICE " & Filtro
 Sql = Sql & vbCrLf & " AND LEN(INDICE) = " & Len(ItemSele) + 3
 Sql = Sql & vbCrLf & " ORDER BY INDICE"
 
 rs.Open Sql, ConActiva, 0, 1
    Do While Not rs.EOF
        
       Rem  MsgBox rs!Indice
        NUMERO = Trim(rs!Indice)
        rs.MoveNext
    Loop
    If Len(NUMERO) = 3 Then
        TxtIndice.Text = Format(CInt(NUMERO) + 1, "000")
    Else
        If NUMERO = "" Then
            TxtIndice.Text = ItemSele & "001"
        Else
            Contador = Mid(NUMERO, Len(NUMERO) - 2)
            dif = Mid(NUMERO, 1, Len(NUMERO) - 3)
            TxtIndice.Text = dif & Format(CInt(Contador) + 1, "000")
        End If
        
      End If

End Sub

Private Sub mnuCrearIgualNivel_Click()
 Dim rs As New ADODB.Recordset
 Dim Sql As String
 Dim NUMERO As String
 Dim Contador As String
 Dim dif As String
  If ctlIndiceCliente.Item_Selecionado = "AIZ" Then
    TxtIndice.Text = "001"
    Exit Sub
  End If
  
  
  Sql = "  SELECT COD_CLIENTE, INDICE From INDICES"
 Sql = Sql & " Where COD_CLIENTE = " & ctlCliente.Valor
 Sql = Sql & " AND INDICE LIKE '" & Mid(ctlIndiceCliente.Item_Selecionado, 1, Len(ctlIndiceCliente.Item_Selecionado) - 3) & "%'"
 Sql = Sql & " AND LEN(INDICE) = " & Len(ctlIndiceCliente.Item_Selecionado)
 Sql = Sql & " ORDER BY INDICE"
 
 rs.Open Sql, ConActiva, 0, 1
    Do While Not rs.EOF
        
       Rem  MsgBox rs!Indice
        NUMERO = Trim(rs!Indice)
        rs.MoveNext
    Loop
    If Len(NUMERO) = 3 Then
        TxtIndice.Text = Format(CInt(NUMERO) + 1, "000")
    Else
        Contador = Mid(NUMERO, Len(NUMERO) - 2)
        dif = Mid(NUMERO, 1, Len(NUMERO) - 3)
        TxtIndice.Text = dif & Format(CInt(Contador) + 1, "000")
      End If


End Sub

Private Sub mnuDocumento_Click()
    Update_Tipo_Indice ("Documento")
End Sub

Private Sub mnuDocumentos_Click()
    Update_Tipo_Indice ("Documentos")
End Sub

Private Sub mnuExpander_Click()
    ExpanderIndex = ctlIndiceCliente.Index_Selecionado
    ctlIndiceCliente.EXPANDER ExpanderIndex
End Sub


Private Sub mnuImprimir_Click()
        ImprimirIndice "0", ctlCliente.Valor
End Sub

Private Sub mnuImprimirSeleccion_Click()
        ImprimirIndice ctlIndiceCliente.Item_Selecionado, ctlCliente.Valor
End Sub

Private Sub mnuIndiceInicial_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
  
  
 Sql = "  SELECT COD_CLIENTE, INDICE From INDICES"
 Sql = Sql & " Where COD_CLIENTE = " & ctlCliente.Valor
 rs.Open Sql, ConActiva, 0, 1
  If rs.EOF Then
    TxtIndice.Text = "001"
  Else
    MsgBox "Ya existen indices", vbCritical
  End If
End Sub

Private Sub mnuLegajo_Click()
 Update_Tipo_Indice ("Legajo")
End Sub

Private Sub mnuModificar_Click()
    TxtIndice.Enabled = False
    TxtIndice.Text = ctlIndiceCliente.Item_Selecionado
    RecuperarIndice ctlCliente.Valor, TxtIndice.Text
   LLENAR_CAMPOS_MODIFICAR ctlCliente.Valor, ctlIndiceCliente.Item_Selecionado
   StatusBar1.Panels.Item(1).Text = "Modificar"
End Sub

Private Sub mnuNuevo_Click()
  Rem  txtIndice.Enabled = True
    TxtIndice.Text = ctlIndiceCliente.Item_Selecionado
    
    LimpiarCampos
   StatusBar1.Panels.Item(1).Text = "Nuevo"
End Sub

Private Sub mnuSector_Click()
Update_Tipo_Indice ("Sector")
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
TxtIndice.Enabled = False
Select Case Button.Caption
Case "Print"
    ImprimirIndice "0", ctlCliente.Valor
Case "Actual."
   ActualizarIndice
Case "Aceptar"
     
        Actualizar StatusBar1.Panels.Item(1).Text
     
     LimpiarCampos
     TxtIndice.Text = ""
     
     ActualizarIndice
Case "Cancelar"
Case "C. Sector"
    CantidadCajasSector
Case "Titulos"
    TituloHerencia 0
Case "Cliente T"
    If ctlCliente.Valor <> 0 Then
        TituloHerencia ctlCliente.Valor
    Else
        MsgBox "Ingrese el cliente", vbCritical
    End If
End Select

End Sub


Public Sub InsertarIndice()
    Dim RsControl As ADODB.Recordset
    Set RsControl = New ADODB.Recordset
    Dim RsMAXDOC As ADODB.Recordset
    Set RsControl = New ADODB.Recordset
    Dim Maxdoc As Integer
    
    Dim Sql As String
    Sql = "SELECT MAX(ID_CODIGO_DOCUMENTO) AS MAXDOC From INDICES Where COD_CLIENTE = " & ctlCliente.Valor
    Set RsMAXDOC = New ADODB.Recordset
    RsMAXDOC.Open Sql, ConActiva, 0, 1
    
    If IsNull(RsMAXDOC!Maxdoc) Then
        Maxdoc = 1
    Else
        Maxdoc = RsMAXDOC!Maxdoc + 1
    End If
    If cboTipo.ListIndex = -1 Then
        MsgBox "Ingrese el tipo de documento", vbInformation
        Exit Sub
    End If
    RsControl.Open "Select * from INDICES WHERE COD_CLIENTE = " & ctlCliente.Valor & " AND INDICE = '" & Trim(TxtIndice.Text) & "'", ConActiva, 0, 1
    If Not RsControl.EOF Then
        MsgBox "ATENCION LA INSERCION NO SE PUEDE REALIZAR", vbCritical
        Exit Sub
    End If
    Dim fecha, NUMERO, lETRA, EXPEDIENTE, APELLIDO_NOMBRE, Descripcion As String
    Dim TOOLTIPFECHA, TOOLTIPNUMERO, TOOLTIPLETRA, TOOLTIPEXPEDIENTE, TOOLTIPAPELLIDO_NOMBRE, FECHA_MODIFICACION, TOOLTIPDESCRIPCION As String
    Dim MASK_EXPEDIENTE, MASK_LETRA, COD_CLIENTE, Indice As String
    COD_CLIENTE = ctlCliente.Valor
    Indice = "'" & Trim(TxtIndice.Text) & "'"
    If Trim(txtDescripcion.Text) <> "" Then
        Descripcion = "'" & Trim(txtDescripcion.Text) & "'"
    Else
        MsgBox "Ingrese descripcion"
        Exit Sub
    End If
    
    
    
'    'Fecha
'    If chkFecha.Value = 1 Then
'       Fecha = "'1'"
'    Else
'       Fecha = "Null"
'    End If
'    If Trim(txtAyudaFecha) <> "" Then
'       TOOLTIPFECHA = "'" & Trim(txtAyudaFecha) & "'"
'    Else
'       TOOLTIPFECHA = "Null"
'    End If
'
'    'Numero
'    If chkNumero.Value Then
'        NUMERO = "'1'"
'    Else
'        NUMERO = "NULL"
'    End If
'    If Trim(txtAyudaNumero.Text) <> "" Then
'        TOOLTIPNUMERO = "'" & Trim(txtAyudaNumero.Text) & "'"
'    Else
'        TOOLTIPNUMERO = "Null"
'    End If
'
'    'Letra
'    If chkLetra.Value = 1 Then
'        LETRA = "'1'"
'    Else
'        LETRA = "Null"
'    End If
'    If Trim(txtAyudaLetra.Text) <> "" Then
'        TOOLTIPLETRA = "'" & Trim(txtAyudaLetra.Text) & "'"
'    Else
'        TOOLTIPLETRA = "Null"
'    End If
'    If Trim(txtFormatoLetra.Text) <> "" Then
'        MASK_LETRA = "'" & Trim(txtFormatoLetra.Text) & "'"
'    Else
'        MASK_LETRA = "Null"
'    End If
'
'    If chkExpediente.Value = 1 Then
'        EXPEDIENTE = "'1'"
'    Else
'         EXPEDIENTE = "Null"
'    End If
'    If Trim(txtFormatoExpediente.Text) <> "" Then
'        MASK_EXPEDIENTE = "'" & Trim(txtFormatoExpediente.Text) & "'"
'    Else
'        MASK_EXPEDIENTE = "Null"
'    End If
'    If Trim(txtAyudaExpediente.Text) <> "" Then
'        TOOLTIPEXPEDIENTE = "'" & Trim(txtAyudaExpediente.Text) & "'"
'    Else
'        TOOLTIPEXPEDIENTE = "Null"
'    End If
'
'    If chkNombre.Value = 1 Then
'        APELLIDO_NOMBRE = "'1'"
'    Else
'        APELLIDO_NOMBRE = "Null"
'    End If
'    If Trim(txtAyudaApellidoNombre.Text) <> "" Then
'        TOOLTIPAPELLIDO_NOMBRE = "'" & Trim(txtAyudaApellidoNombre.Text) & "'"
'    Else
'        TOOLTIPAPELLIDO_NOMBRE = "Null"
'    End If
'    If Trim(txtAyudaDescripcion) <> "" Then
'        TOOLTIPDESCRIPCION = "'" & Trim(txtAyudaDescripcion) & "'"
'    Else
'        TOOLTIPDESCRIPCION = "Null"
'    End If
'
'        Sql = " INSERT INTO INDICES (COD_CLIENTE,ID_CODIGO_DOCUMENTO, INDICE, DESCRIPCION ,FECHA, NUMERO,"
'        Sql = Sql & vbCrLf & " LETRA, EXPEDIENTE, APELLIDO_NOMBRE,"
'        Sql = Sql & vbCrLf & " MASK_EXPEDIENTE, MASK_LETRA, TOOLTIPFECHA,"
'        Sql = Sql & vbCrLf & " TOOLTIPNUMERO, TOOLTIPLETRA, TOOLTIPEXPEDIENTE,"
'        Sql = Sql & vbCrLf & " TOOLTIPAPELLIDO_NOMBRE, FECHA_MODIFICACION,"
'        Sql = Sql & vbCrLf & " TOOLTIPDESCRIPCION,TIPO_INDICE)"
'        Sql = Sql & vbCrLf & " VALUES (" & COD_CLIENTE & "," & Maxdoc & "," & Indice & "," & DESCRIPCION & "," & Fecha & "," & NUMERO & ","
'        Sql = Sql & vbCrLf & LETRA & "," & EXPEDIENTE & "," & APELLIDO_NOMBRE & ","
'        Sql = Sql & vbCrLf & MASK_EXPEDIENTE & "," & MASK_LETRA & "," & TOOLTIPFECHA & ","
'        Sql = Sql & vbCrLf & TOOLTIPNUMERO & "," & TOOLTIPLETRA & "," & TOOLTIPEXPEDIENTE & ","
'        Sql = Sql & vbCrLf & TOOLTIPAPELLIDO_NOMBRE & "," & SysDate & ","
'        Sql = Sql & vbCrLf & TOOLTIPDESCRIPCION & ",'" & Trim(cboTipo.Text) & "')"
'        ExecutarSql Sql
End Sub

Public Sub ActualizarIndices()
'    Dim Fecha, NUMERO, LETRA, EXPEDIENTE, APELLIDO_NOMBRE, DESCRIPCION As String
'    Dim TOOLTIPFECHA, TOOLTIPNUMERO, TOOLTIPLETRA, TOOLTIPEXPEDIENTE, TOOLTIPAPELLIDO_NOMBRE, FECHA_MODIFICACION, TOOLTIPDESCRIPCION As String
'    Dim MASK_EXPEDIENTE, MASK_LETRA, COD_CLIENTE, Indice As String
'    Dim Sql As String
'
'
'    COD_CLIENTE = ctlCliente.Valor
'    Indice = "'" & Trim(TxtIndice.Text) & "'"
'
'    If Trim(txtDescripcion.Text) <> "" Then
'        DESCRIPCION = "'" & Trim(txtDescripcion.Text) & "'"
'    Else
'        MsgBox "Ingrese descripcion"
'        Exit Sub
'    End If
'
'    'Fecha
'    If chkFecha.Value = 1 Then
'       Fecha = "'1'"
'    Else
'       Fecha = "Null"
'    End If
'    If Trim(txtAyudaFecha) <> "" Then
'       TOOLTIPFECHA = "'" & Trim(txtAyudaFecha) & "'"
'    Else
'       TOOLTIPFECHA = "Null"
'    End If
'
'    'Numero
'    If chkNumero.Value Then
'        NUMERO = "'1'"
'    Else
'        NUMERO = "NULL"
'    End If
'    If Trim(txtAyudaNumero.Text) <> "" Then
'        TOOLTIPNUMERO = "'" & Trim(txtAyudaNumero.Text) & "'"
'    Else
'        TOOLTIPNUMERO = "Null"
'    End If
'
'    'Letra
'    If chkLetra.Value = 1 Then
'        LETRA = "'1'"
'    Else
'        LETRA = "Null"
'    End If
'    If Trim(txtAyudaLetra.Text) <> "" Then
'        TOOLTIPLETRA = "'" & Trim(txtAyudaLetra.Text) & "'"
'    Else
'        TOOLTIPLETRA = "Null"
'    End If
'    If Trim(txtFormatoLetra.Text) <> "" Then
'        MASK_LETRA = "'" & Trim(txtFormatoLetra.Text) & "'"
'    Else
'        MASK_LETRA = "Null"
'    End If
'
'    If chkExpediente.Value = 1 Then
'        EXPEDIENTE = "'1'"
'    Else
'         EXPEDIENTE = "Null"
'    End If
'    If Trim(txtFormatoExpediente.Text) <> "" Then
'        MASK_EXPEDIENTE = "'" & Trim(txtFormatoExpediente.Text) & "'"
'    Else
'        MASK_EXPEDIENTE = "Null"
'    End If
'    If Trim(txtAyudaExpediente.Text) <> "" Then
'        TOOLTIPEXPEDIENTE = "'" & Trim(txtAyudaExpediente.Text) & "'"
'    Else
'        TOOLTIPEXPEDIENTE = "Null"
'    End If
'
'    If chkNombre.Value = 1 Then
'        APELLIDO_NOMBRE = "'1'"
'    Else
'        APELLIDO_NOMBRE = "Null"
'    End If
'    If Trim(txtAyudaApellidoNombre.Text) <> "" Then
'        TOOLTIPAPELLIDO_NOMBRE = "'" & Trim(txtAyudaApellidoNombre.Text) & "'"
'    Else
'        TOOLTIPAPELLIDO_NOMBRE = "Null"
'    End If
'    If Trim(txtAyudaDescripcion) <> "" Then
'        TOOLTIPDESCRIPCION = "'" & Trim(txtAyudaDescripcion) & "'"
'    Else
'        TOOLTIPDESCRIPCION = "Null"
'    End If
'
'    Sql = " Update INDICES"
'    Sql = Sql & vbCrLf & " SET DESCRIPCION =" & DESCRIPCION & ",  FECHA =" & Fecha & ","
'    Sql = Sql & vbCrLf & " NUMERO =" & NUMERO & ", LETRA =" & LETRA & ", EXPEDIENTE =" & EXPEDIENTE & ", APELLIDO_NOMBRE =" & APELLIDO_NOMBRE & ","
'    Sql = Sql & vbCrLf & " MASK_EXPEDIENTE =" & MASK_EXPEDIENTE & " , MASK_LETRA =" & MASK_LETRA & " , TOOLTIPFECHA =" & TOOLTIPFECHA & ","
'    Sql = Sql & vbCrLf & " TOOLTIPNUMERO =" & TOOLTIPNUMERO & ", TOOLTIPLETRA =" & TOOLTIPLETRA & ","
'    Sql = Sql & vbCrLf & " TOOLTIPEXPEDIENTE =" & TOOLTIPEXPEDIENTE & ", TOOLTIPAPELLIDO_NOMBRE =" & TOOLTIPAPELLIDO_NOMBRE & ","
'    Sql = Sql & vbCrLf & " FECHA_MODIFICACION =" & SysDate & ", TOOLTIPDESCRIPCION =" & TOOLTIPDESCRIPCION
'    Sql = Sql & vbCrLf & " WHERE (COD_CLIENTE = " & COD_CLIENTE & ") AND "
'    Sql = Sql & vbCrLf & " (INDICE =" & Indice & ")"
'    ExecutarSql Sql

End Sub

Public Sub RecuperarIndice(COD_CLIENTE As Integer, Indice As String)
    Dim rsIndice As ADODB.Recordset
    Set rsIndice = New ADODB.Recordset
    
    rsIndice.Open "Select * From indiceS where cod_cliente =" & COD_CLIENTE & " and Indice='" & Indice & "'", ConActiva, 0, 1
    Dim fecha, NUMERO, lETRA, EXPEDIENTE, APELLIDO_NOMBRE, Descripcion As String
    Dim TOOLTIPFECHA, TOOLTIPNUMERO, TOOLTIPLETRA, TOOLTIPEXPEDIENTE, TOOLTIPAPELLIDO_NOMBRE, FECHA_MODIFICACION, TOOLTIPDESCRIPCION As String
    Dim MASK_EXPEDIENTE, MASK_LETRA As String
    Dim Sql As String
    
'    With rsIndice
'        If Not rsIndice.EOF Then
'            COD_CLIENTE = ctlCliente.Valor
'            Indice = "'" & Trim(TxtIndice.Text) & "'"
'            If IsNull(!DESCRIPCION) Then
'                txtDescripcion.Text = ""
'            Else
'                txtDescripcion.Text = Trim(!DESCRIPCION)
'            End If
'            'Fecha
'            If IsNull(!Fecha) Then
'               chkFecha.Value = 0
'            Else
'               chkFecha.Value = 1
'            End If
'
'            If IsNull(!TOOLTIPFECHA) Then
'               txtAyudaFecha = ""
'            Else
'               txtAyudaFecha = Trim(!TOOLTIPFECHA)
'            End If
'            'Numero
'            If IsNull(!NUMERO) Then
'                chkNumero.Value = 0
'            Else
'                chkNumero.Value = 1
'            End If
'            If IsNull(!TOOLTIPNUMERO) Then
'                txtAyudaNumero.Text = ""
'            Else
'                txtAyudaNumero.Text = Trim(!TOOLTIPNUMERO)
'            End If
'            'Letra
'            If IsNull(!LETRA) Then
'               chkLetra.Value = 0
'            Else
'               chkLetra.Value = 1
'            End If
'            If IsNull(!TOOLTIPLETRA) Then
'               txtAyudaLetra.Text = ""
'            Else
'                txtAyudaLetra.Text = Trim(!TOOLTIPLETRA)
'            End If
'            If IsNull(!MASK_LETRA) Then
'                txtFormatoLetra.Text = ""
'            Else
'                txtFormatoLetra.Text = Trim(!MASK_LETRA)
'            End If
'            If IsNull(!EXPEDIENTE) Then
'                chkExpediente.Value = 0
'            Else
'                chkExpediente.Value = 1
'            End If
'            If IsNull(!MASK_EXPEDIENTE) Then
'              txtFormatoExpediente.Text = ""
'            Else
'              txtFormatoExpediente.Text = Trim(!MASK_EXPEDIENTE)
'            End If
'            If IsNull(!TOOLTIPEXPEDIENTE) Then
'                txtAyudaExpediente.Text = ""
'            Else
'                txtAyudaExpediente.Text = Trim(!TOOLTIPEXPEDIENTE)
'            End If
'            If IsNull(!APELLIDO_NOMBRE) Then
'               chkNombre.Value = 0
'            Else
'               chkNombre.Value = 1
'            End If
'            If IsNull(!TOOLTIPAPELLIDO_NOMBRE) Then
'                txtAyudaApellidoNombre.Text = ""
'            Else
'                txtAyudaApellidoNombre.Text = Trim(!TOOLTIPAPELLIDO_NOMBRE)
'            End If
'            If IsNull(!TOOLTIPDESCRIPCION) Then
'                txtAyudaDescripcion = ""
'            Else
'                txtAyudaDescripcion.Text = Trim(!TOOLTIPDESCRIPCION)
'            End If
'        End If
'    End With

End Sub

Public Sub LimpiarCampos()
txtDescripcion.Text = ""
cboTipo.ListIndex = -1
TxtIndice.Text = ""
txt_ID.Text = ""
txtNro_Documento.Text = ""
chk_Requerir_Descripcion.value = 0
chk_Requerir_Fecha_Desde.value = 0
chk_Requerir_Fecha_Hasta.value = 0
chk_Requerir_Letra_Desde.value = 0
chk_Requerir_Letra_Hasta.value = 0
chk_Requerir_Nro_Desde.value = 0
chk_Requerir_Nro_Hasta.value = 0
chkFecha_Desde.value = 0
chkFecha_Hasta.value = 0
chkHabilitar_descripcion.value = 0
chkLetra_Desde.value = 0
chkLetra_Hasta.value = 0
chkNro_Desde.value = 0
chkNro_Hasta.value = 0
chkTodos.value = 0
chkTodosRequerir.value = 0
chk_Copiar_Fecha.value = 0
chk_Copiar_Letra.value = 0
chk_Copiar_Nro.value = 0
txt_Titulo_Descripcion.Text = ""
txt_Titulo_Fecha_Desde.Text = ""
txt_Titulo_Fecha_Hasta.Text = ""
txt_Titulo_Letra_Desde.Text = ""
txt_Titulo_Letra_Hasta.Text = ""
txt_Titulo_Nro_Desde.Text = ""
txt_Titulo_Nro_Hasta.Text = ""
txtDescripcion.Text = ""
txtCONTROL_LOGICO_LARGO_NRO_DESDE_INICIO = ""
txtCONTROL_LOGICO_LARGO_NRO_DESDE_HASTA = ""
TXTCONTROL_LOGICO_LARGO_NRO_HASTA_INICIO = ""
TXTCONTROL_LOGICO_LARGO_NRO_HASTA_HASTA = ""
TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO = ""
TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA = ""
TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO = ""
TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA = ""
txtCostoPreparacion = ""
txtCostoDigitalizacion = ""
txtCostoIndexacion = ""
txtCostoArmado = ""
txtCostoCargaLegajo = ""
cboTipoArchivoExtencion.ListIndex = -1
optBarra_Ninguno.value = True
chkEtiquetaLegajo.value = 0




'            txtDescripcion.Text = ""
'            chkFecha.Value = 0
'            txtAyudaFecha = ""
'            chkNumero.Value = 0
'            txtAyudaNumero.Text = ""
'            chkLetra.Value = 0
'            txtAyudaLetra.Text = ""
'            txtFormatoLetra.Text = ""
'            chkExpediente.Value = 0
'            txtFormatoExpediente.Text = ""
'            txtAyudaExpediente.Text = ""
'            chkNombre.Value = 0
'            txtAyudaApellidoNombre.Text = ""
'            txtAyudaDescripcion = ""
End Sub





Public Sub Update_Tipo_Indice(DATO As String)

Dim Sql As String
Sql = " Update INDICES SET TIPO_INDICE = '" & DATO & "'"
Sql = Sql & " WHERE COD_CLIENTE = " & ctlCliente.Valor
Sql = Sql & " AND INDICE = '" & ctlIndiceCliente.Item_Selecionado & "'"
ExecutarSql Sql
ActualizarIndice

End Sub

Public Function Selecion_Imagen(DATO) As Integer
 If IsNull(DATO) Then
 Selecion_Imagen = 1
 Else
    Select Case Trim(DATO)
    Case "Sector"
        Selecion_Imagen = 12
    Case "Documentos"
        Selecion_Imagen = 2
    Case "Documento"
        Selecion_Imagen = 10
    End Select
 End If
End Function

Public Sub ActualizarIndice()
    ctlIndiceCliente.Actualizar ctlCliente.Valor, Nulo, ExpanderIndex
    End Sub

Public Sub ImprimirIndice(Indice As String, Cliente As Integer)
    Dim Sql As String
    MousePointer = 11
    Sql = " SELECT *"
    Sql = Sql & "  From V_INDICES"
    Sql = Sql & " Where COD_CLIENTE = " & Cliente
    Sql = Sql & " AND INDICE like '" & Indice & "%'"
    Sql = Sql & " ORDER BY INDICE"
    frmReportes.ImprimirReporte PasoReportes + "rptindices.rpt", Sql, True
    MousePointer = 0
End Sub

Public Sub CantidadCajasSector()
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim RSCONT As ADODB.Recordset
    Dim Indice As String


         Sql = " SELECT INDICE"
        Sql = Sql & " From INDICES"
        Sql = Sql & " WHERE COD_CLIENTE = " & ctlCliente.Valor
        Sql = Sql & " AND TIPO_INDICE = 'Sector'"
        Sql = Sql & " ORDER BY INDICE"
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        Do While Not rs.EOF
            Indice = Trim(rs!Indice)
            Sql = " SELECT COUNT(COUNT(*)) AS cantidad"
            Sql = Sql & "  From REFERENCIAS "
            Sql = Sql & "  WHERE COD_CLIENTE = " & ctlCliente.Valor
            Sql = Sql & "  AND (INDICE LIKE '" & Indice & "%')"
            Sql = Sql & "  GROUP BY NRO_CAJA"
            Set RSCONT = New ADODB.Recordset
            RSCONT.Open Sql, ConActiva, 0, 1
            If Not RSCONT.EOF Then
                Sql = " Update INDICES"
                Sql = Sql & "  Set CANTIDAD_CAJAS_ACUMULADO  = " & RSCONT!cantidad
                Sql = Sql & "  WHERE COD_CLIENTE = " & ctlCliente.Valor
               Sql = Sql & "   AND (INDICE = '" & Indice & "')"
                ExecutarSql (Sql)
            End If
            
            Sql = " SELECT COUNT(COUNT(*)) AS cantidad"
            Sql = Sql & " From REFERENCIAS"
            Sql = Sql & "  WHERE COD_CLIENTE = " & ctlCliente.Valor
            Sql = Sql & "  AND (INDICE LIKE '" & Indice & "')"
            Sql = Sql & "  GROUP BY NRO_CAJA"
            Set RSCONT = New ADODB.Recordset
            RSCONT.Open Sql, ConActiva, 0, 1
            If Not RSCONT.EOF Then
                Sql = " Update INDICES"
                Sql = Sql & "  Set CANTIDAD_CAJAS_SOLO  = " & RSCONT!cantidad
                Sql = Sql & "  WHERE COD_CLIENTE =  " & ctlCliente.Valor
                Sql = Sql & "  AND (INDICE = '" & Indice & "')"
                ExecutarSql (Sql)
            End If
            rs.MoveNext
        Loop

End Sub

Public Sub LLENAR_CAMPOS_MODIFICAR(Cliente As Integer, Indice As String)

    Dim rsIndices As ADODB.Recordset
    Dim Sql As String
    Dim I As Integer
    
    
        Sql = "  SELECT        ID, COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, TIPO_INDICE, DESCRIPCION, TITULO_FECHA_DESDE, TITULO_FECHA_HASTA, TITULO_LETRA_DESDE,"
        Sql = Sql & " TITULO_LETRA_HASTA, TITULO_NRO_DESDE, TITULO_NRO_HASTA, TITULO_DESCRIPCION, HABILITAR_FECHA_DESDE, HABILITAR_FECHA_HASTA,"
        Sql = Sql & " HABILITAR_LETRA_DESDE, HABILITAR_LETRA_HASTA, HABILITAR_NRO_DESDE, HABILITAR_NRO_HASTA, HABILITAR_DESCRIPCION, REQUERIR_FECHA_DESDE,"
        Sql = Sql & " REQUERIR_FECHA_HASTA, REQUERIR_LETRA_DESDE, REQUERIR_LETRA_HASTA, REQUERIR_NRO_DESDE, REQUERIR_NRO_HASTA, REQUERIR_DESCRIPCION,"
        Sql = Sql & " COPIAR_FECHA, COPIAR_LETRA, COPIAR_NRO, HABILITAR_ETIQUETA_LEGAJO, CONTROL_LOGICO_LARGO_NRO_DESDE_INICIO,"
        Sql = Sql & " CONTROL_LOGICO_LARGO_NRO_DESDE_HASTA, CONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO, CONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA,"
        Sql = Sql & " COSTOPREPARACION , COSTODIGITALIZACION, COSTOINDEXACION, COSTOARMADO, COSTOCARGALEGAJO , BARRA , "
        Sql = Sql & " CONTROL_LOGICO_LARGO_NRO_HASTA_INICIO, CONTROL_LOGICO_LARGO_NRO_HASTA_HASTA, CONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO,"
        Sql = Sql & " CONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA , BARRA , HABILITAR_ETIQUETA_LEGAJO , TIPO_ARCHIVO "
        Sql = Sql & " From INDICES"
        Sql = Sql & "  Where COD_CLIENTE = " & Cliente
        Sql = Sql & " And INDICE = '" & Trim(Indice) & "'"
    
Set rsIndices = New ADODB.Recordset
rsIndices.Open Sql, ConActiva, 0, 1

With rsIndices
    
    If Not .EOF Then
    
    
    txt_ID.Text = rsIndices!ID
    
    For I = 0 To cboTipo.ListCount - 1
            cboTipo.ListIndex = I
        If cboTipo.Text = Trim(rsIndices!Tipo_Indice) Then
            Exit For
        End If
        
        
    Next
    
    txtNro_Documento.Text = rsIndices!ID_CODIGO_DOCUMENTO
    TxtIndice.Text = Trim(rsIndices!Indice)
    If Not IsNull(!Descripcion) Then
        txtDescripcion.Text = Trim(!Descripcion)
    Else
        txtDescripcion.Text = ""
    End If
    
    If Not IsNull(!TITULO_FECHA_DESDE) Then
        txt_Titulo_Fecha_Desde.Text = Trim(!TITULO_FECHA_DESDE)
    Else
        txt_Titulo_Fecha_Desde.Text = ""
    End If
    
    If Not IsNull(!TITULO_FECHA_HASTA) Then
        txt_Titulo_Fecha_Hasta.Text = Trim(!TITULO_FECHA_HASTA)
    Else
        txt_Titulo_Fecha_Hasta.Text = ""
    End If
    
    If Not IsNull(!TITULO_LETRA_DESDE) Then
        txt_Titulo_Letra_Desde.Text = Trim(!TITULO_LETRA_DESDE)
    Else
        txt_Titulo_Letra_Desde.Text = ""
    End If
    
    If Not IsNull(!TITULO_LETRA_HASTA) Then
        txt_Titulo_Letra_Hasta.Text = Trim(!TITULO_LETRA_HASTA)
    Else
        txt_Titulo_Letra_Hasta.Text = ""
    End If
    
    
    If Not IsNull(!TITULO_NRO_DESDE) Then
        txt_Titulo_Nro_Desde.Text = Trim(!TITULO_NRO_DESDE)
     Else
        txt_Titulo_Nro_Desde.Text = ""
    End If
    
    If Not IsNull(!TITULO_NRO_HASTA) Then
        txt_Titulo_Nro_Hasta.Text = Trim(!TITULO_NRO_HASTA)
     Else
        txt_Titulo_Nro_Hasta.Text = ""
    End If
    
    
    If Not IsNull(!TITULO_DESCRIPCION) Then
        txt_Titulo_Descripcion.Text = Trim(!TITULO_DESCRIPCION)
    Else
        txt_Titulo_Descripcion.Text = ""
    End If
    
    If !HABILITAR_FECHA_DESDE = True Then
        chkFecha_Desde.value = 1
    Else
        chkFecha_Desde.value = 0
    End If
    
        
    If !HABILITAR_FECHA_HASTA = True Then
       chkFecha_Hasta.value = 1
    Else
        chkFecha_Hasta.value = 0
    End If
    
    
    If !HABILITAR_LETRA_DESDE = True Then
        chkLetra_Desde.value = 1
    Else
        chkLetra_Desde.value = 0
    End If
    
    If !HABILITAR_LETRA_HASTA = True Then
        chkLetra_Hasta.value = 1
    Else
        chkLetra_Hasta.value = 0
    End If
     
     If !HABILITAR_NRO_DESDE = True Then
        chkNro_Desde.value = 1
     Else
        chkNro_Desde.value = 0
     End If
     
     If !HABILITAR_NRO_HASTA = True Then
        chkNro_Hasta.value = 1
     Else
        chkNro_Hasta.value = 0
     End If
      
      If !HABILITAR_DESCRIPCION = True Then
           chkHabilitar_descripcion = 1
      Else
            chkHabilitar_descripcion.value = 0
      End If
      
      If !REQUERIR_FECHA_DESDE = True Then
        chk_Requerir_Fecha_Desde.value = 1
      Else
        chk_Requerir_Fecha_Desde.value = 0
      End If
      
      If !REQUERIR_FECHA_HASTA = True Then
          chk_Requerir_Fecha_Hasta.value = 1
      Else
          chk_Requerir_Fecha_Hasta.value = 0
      End If
      
      If !REQUERIR_LETRA_DESDE = True Then
        chk_Requerir_Letra_Desde.value = 1
      Else
        chk_Requerir_Letra_Desde.value = 0
      End If
      
      If !REQUERIR_LETRA_HASTA = True Then
        chk_Requerir_Letra_Hasta.value = 1
      Else
        chk_Requerir_Letra_Hasta.value = 0
      End If
      
      If !REQUERIR_NRO_DESDE = True Then
        chk_Requerir_Nro_Desde.value = 1
      Else
        chk_Requerir_Nro_Desde.value = 0
      End If
      
      If !REQUERIR_NRO_HASTA = True Then
        chk_Requerir_Nro_Hasta.value = 1
      Else
        chk_Requerir_Nro_Hasta.value = 0
      End If
      
      If !REQUERIR_DESCRIPCION = True Then
        chk_Requerir_Descripcion.value = 1
      Else
        chk_Requerir_Descripcion.value = 0
      End If
    
        If !COPIAR_FECHA = True Then
            chk_Copiar_Fecha.value = 1
        Else
            chk_Copiar_Fecha.value = 0
        End If
        
        If !COPIAR_LETRA = True Then
            chk_Copiar_Letra.value = 1
        Else
            chk_Copiar_Fecha.value = 0
        End If
        
        If !COPIAR_NRO = True Then
            chk_Copiar_Nro.value = 1
        Else
            chk_Copiar_Nro.value = 0
        End If
        
            
        If Not IsNull(!CONTROL_LOGICO_LARGO_NRO_DESDE_INICIO) Then
        
            txtCONTROL_LOGICO_LARGO_NRO_DESDE_INICIO.Text = !CONTROL_LOGICO_LARGO_NRO_DESDE_INICIO
        Else
            txtCONTROL_LOGICO_LARGO_NRO_DESDE_INICIO.Text = ""
        End If
           
        If Not IsNull(!CONTROL_LOGICO_LARGO_NRO_DESDE_HASTA) Then
            txtCONTROL_LOGICO_LARGO_NRO_DESDE_HASTA.Text = !CONTROL_LOGICO_LARGO_NRO_DESDE_HASTA
        Else
            txtCONTROL_LOGICO_LARGO_NRO_DESDE_HASTA.Text = ""
        End If
        
        
        
        If Not IsNull(!CONTROL_LOGICO_LARGO_NRO_HASTA_INICIO) Then
            TXTCONTROL_LOGICO_LARGO_NRO_HASTA_INICIO.Text = !CONTROL_LOGICO_LARGO_NRO_HASTA_INICIO
        Else
            TXTCONTROL_LOGICO_LARGO_NRO_HASTA_INICIO.Text = ""
        End If
           
        If Not IsNull(!CONTROL_LOGICO_LARGO_NRO_HASTA_HASTA) Then
            TXTCONTROL_LOGICO_LARGO_NRO_HASTA_HASTA.Text = !CONTROL_LOGICO_LARGO_NRO_HASTA_HASTA
        Else
            TXTCONTROL_LOGICO_LARGO_NRO_HASTA_HASTA.Text = ""
        End If
        
        
        
        If Not IsNull(!CONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO) Then
            TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO.Text = !CONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO
        Else
            TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO.Text = ""
        End If
        
        If Not IsNull(!CONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA) Then
            TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA.Text = !CONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA
        Else
            TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA.Text = ""
        End If
        
        
        
        If Not IsNull(!CONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO) Then
            TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO.Text = !CONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO
        Else
            TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO.Text = ""
        End If
        
        If Not IsNull(!CONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA) Then
            TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA.Text = !CONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA
        Else
            TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA.Text = ""
        End If
        
        
        If Not IsNull(!COSTOPREPARACION) Then
            txtCostoPreparacion.Text = !COSTOPREPARACION
        Else
            txtCostoPreparacion.Text = ""
        End If
        
        
        If Not IsNull(!COSTODIGITALIZACION) Then
            txtCostoDigitalizacion.Text = !COSTODIGITALIZACION
        Else
            txtCostoDigitalizacion.Text = ""
        End If
        
        If Not IsNull(!COSTOINDEXACION) Then
            txtCostoIndexacion.Text = !COSTOINDEXACION
        Else
            txtCostoIndexacion.Text = ""
        End If
 
        If Not IsNull(!COSTOARMADO) Then
            txtCostoArmado.Text = !COSTOARMADO
        Else
            txtCostoArmado.Text = ""
        End If
        
        
        If Not IsNull(!COSTOCARGALEGAJO) Then
            txtCostoCargaLegajo.Text = !COSTOCARGALEGAJO
        Else
            txtCostoCargaLegajo.Text = ""
        End If
        
        If (!HABILITAR_ETIQUETA_LEGAJO) = True Then
            chkEtiquetaLegajo.value = 1
        Else
            chkEtiquetaLegajo.value = 0
        End If
        
        
    
       If IsNull(!BARRA) Then
            optBarra_Ninguno.value = True
        
       Else
        
        Select Case UCase(Trim(!BARRA))
         Case "NRO_DESDE"
            optBarra_NRO_DESDE.value = True
         Case "NRO_HASTA"
            optBarra_NRO_HASTA.value = True
         Case "LETRA_DESDE"
            optBarra_Letra_Desde = True
         Case "LETRA_HASTA"
            optBarra_Letra_Hasta = True
         Case "ETIQUETA_LEGAJO"
            optEtiqueta_Legajo.value = True
        End Select
       End If
       
    If IsNull(!TIPO_ARCHIVO) Then
        cboTipoArchivoExtencion.ListIndex = -1
    Else
        cboTipoArchivoExtencion.Text = !TIPO_ARCHIVO
    End If
    
    
    
    
    End If
    End With




End Sub



Public Sub Actualizar(Operacion As String)

On Error GoTo salir:
    
    Dim Descripcion, TITULO_FECHA_DESDE, TITULO_FECHA_HASTA, TITULO_LETRA_DESDE, TITULO_LETRA_HASTA, TITULO_NRO_DESDE, TITULO_NRO_HASTA, TITULO_DESCRIPCION As String
    Dim HABILITAR_FECHA_DESDE, HABILITAR_FECHA_HASTA, HABILITAR_LETRA_DESDE, HABILITAR_LETRA_HASTA, HABILITAR_NRO_DESDE, HABILITAR_NRO_HASTA, HABILITAR_DESCRIPCION As Integer
    Dim REQUERIR_FECHA_DESDE, REQUERIR_FECHA_HASTA, REQUERIR_LETRA_DESDE, REQUERIR_LETRA_HASTA, REQUERIR_NRO_DESDE, REQUERIR_NRO_HASTA, REQUERIR_DESCRIPCION As Integer
    Dim COPIAR_FECHA, COPIAR_LETRA, COPIAR_NRO As Integer
    Dim CONTROL_LOGICO_LARGO_NRO_DESDE_INICIO  As String
    Dim CONTROL_LOGICO_LARGO_NRO_DESDE_HASTA As String
    Dim CONTROL_LOGICO_LARGO_NRO_HASTA_INICIO As String
    Dim CONTROL_LOGICO_LARGO_NRO_HASTA_HASTA As String
    Dim CONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO As String
    Dim CONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA As String
    Dim CONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO As String
    Dim CONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA As String
    Dim COSTOPREPARACION As String
    Dim COSTODIGITALIZACION As String
    Dim COSTOINDEXACION As String
    Dim COSTOARMADO As String
    Dim COSTOCARGALEGAJO As String
    Dim BARRA As String
    Dim HABILITAR_ETIQUETA_LEGAJO As Integer
    Dim TIPO_ARCHIVO As String
    
    
    
    
         
        
    
    
    
    
    
    
    Dim RsMAXDOC As ADODB.Recordset
    Dim Sql As String
    
    If IsNull(ctlCliente.Valor) Then
         MsgBox "Ingrese le cliente"
           Exit Sub
    End If
    
    If Trim(TxtIndice.Text) = "" Then
        MsgBox "Ingrese el indice"
        Exit Sub
    End If
    
    If cboTipo.Text = "" Then
        MsgBox "Falta el tipo de Documento"
        Exit Sub
    End If
    
    
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "Falta la descripcion"
        Exit Sub
    Else
        Descripcion = "'" & UCase(Trim(txtDescripcion.Text)) & "'"
    End If
    
    If Trim(txt_Titulo_Fecha_Desde.Text) = "" Then
        TITULO_FECHA_DESDE = "Null"
    Else
        TITULO_FECHA_DESDE = "'" & Trim(txt_Titulo_Fecha_Desde.Text) & "'"
    End If
    
    
    
    If Trim(txt_Titulo_Fecha_Hasta.Text) = "" Then
        TITULO_FECHA_HASTA = "Null"
    Else
        TITULO_FECHA_HASTA = "'" & Trim(txt_Titulo_Fecha_Hasta.Text) & "'"
    End If
    
    If Trim(txt_Titulo_Letra_Desde.Text) = "" Then
        TITULO_LETRA_DESDE = "Null"
    Else
        TITULO_LETRA_DESDE = "'" & Trim(txt_Titulo_Letra_Desde.Text) & "'"
    End If
    
    If Trim(txt_Titulo_Letra_Hasta.Text) = "" Then
        TITULO_LETRA_HASTA = "Null"
    Else
        TITULO_LETRA_HASTA = "'" & Trim(txt_Titulo_Letra_Hasta.Text) & "'"
    End If
    
    
    If Trim(txt_Titulo_Nro_Desde.Text) = "" Then
        TITULO_NRO_DESDE = "Null"
     Else
       TITULO_NRO_DESDE = "'" & Trim(txt_Titulo_Nro_Desde.Text) & "'"
    End If
    
    If Trim(txt_Titulo_Nro_Hasta.Text) = "" Then
        TITULO_NRO_HASTA = "Null"
     Else
        TITULO_NRO_HASTA = "'" & Trim(txt_Titulo_Nro_Hasta.Text) & "'"
    End If
    
    
    If Trim(txt_Titulo_Descripcion.Text) = "" Then
       TITULO_DESCRIPCION = "Null"
    Else
       TITULO_DESCRIPCION = "'" & Trim(txt_Titulo_Descripcion.Text) & "'"
    End If
    

    
    
    HABILITAR_FECHA_DESDE = chkFecha_Desde.value
    HABILITAR_FECHA_HASTA = chkFecha_Hasta.value
    HABILITAR_LETRA_DESDE = chkLetra_Desde.value
    HABILITAR_LETRA_HASTA = chkLetra_Hasta.value
    HABILITAR_NRO_DESDE = chkNro_Desde.value
    HABILITAR_NRO_HASTA = chkNro_Hasta.value
    HABILITAR_DESCRIPCION = chkHabilitar_descripcion.value
    REQUERIR_FECHA_DESDE = chk_Requerir_Fecha_Desde.value
    REQUERIR_FECHA_HASTA = chk_Requerir_Fecha_Hasta.value
    REQUERIR_LETRA_DESDE = chk_Requerir_Letra_Desde.value
    REQUERIR_LETRA_HASTA = chk_Requerir_Letra_Hasta.value
    REQUERIR_NRO_DESDE = chk_Requerir_Nro_Desde.value
    REQUERIR_NRO_HASTA = chk_Requerir_Nro_Hasta.value
    REQUERIR_DESCRIPCION = chk_Requerir_Descripcion.value
    COPIAR_FECHA = chk_Copiar_Fecha.value
    COPIAR_NRO = chk_Copiar_Nro.value
    COPIAR_LETRA = chk_Copiar_Letra.value
    
    
    
    
        If Trim(txtCONTROL_LOGICO_LARGO_NRO_DESDE_INICIO.Text) <> "" Then
            CONTROL_LOGICO_LARGO_NRO_DESDE_INICIO = CInt(txtCONTROL_LOGICO_LARGO_NRO_DESDE_INICIO.Text)
        Else
            CONTROL_LOGICO_LARGO_NRO_DESDE_INICIO = "NULL"
        End If
        If Trim(txtCONTROL_LOGICO_LARGO_NRO_DESDE_HASTA.Text) <> "" Then
           CONTROL_LOGICO_LARGO_NRO_DESDE_HASTA = txtCONTROL_LOGICO_LARGO_NRO_DESDE_HASTA.Text
        Else
           CONTROL_LOGICO_LARGO_NRO_DESDE_HASTA = "NULL"
        End If
        
        
        
        If Trim(TXTCONTROL_LOGICO_LARGO_NRO_HASTA_INICIO.Text) <> "" Then
           CONTROL_LOGICO_LARGO_NRO_HASTA_INICIO = CInt(TXTCONTROL_LOGICO_LARGO_NRO_HASTA_INICIO.Text)
        Else
           CONTROL_LOGICO_LARGO_NRO_HASTA_INICIO = "NULL"
        End If
           
        If Trim(TXTCONTROL_LOGICO_LARGO_NRO_HASTA_HASTA.Text) <> "" Then
            CONTROL_LOGICO_LARGO_NRO_HASTA_HASTA = CInt(TXTCONTROL_LOGICO_LARGO_NRO_HASTA_HASTA.Text)
        Else
            CONTROL_LOGICO_LARGO_NRO_HASTA_HASTA = "NULL"
        End If
        
        
        
        If Trim(TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO.Text) <> "" Then
           CONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO = CInt(TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO.Text)
        Else
           CONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO = "NULL"
        End If
        
        
        
        If Trim(TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA.Text) <> "" Then
            CONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA = CInt(TXTCONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA.Text)
        Else
            CONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA = "NULL"
        End If
        
        
        
        If TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO.Text <> "" Then
            CONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO = CInt(TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO.Text)
        Else
            CONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO = "NULL"
        End If
        
        If Trim(TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA.Text) <> "" Then
            CONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA = CInt(TXTCONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA.Text)
        Else
            CONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA = "NULL"
        End If
        
        
        If Trim(txtCostoPreparacion.Text) <> "" Then
            COSTOPREPARACION = "'" & txtCostoPreparacion.Text & "'"
        Else
            COSTOPREPARACION = "NULL"
        End If
        
        
        If Trim(txtCostoDigitalizacion.Text) <> "" Then
           COSTODIGITALIZACION = "'" & txtCostoDigitalizacion.Text & "'"
        Else
           COSTODIGITALIZACION = "NULL"
        End If
        
        
        If Trim(txtCostoIndexacion.Text) <> "" Then
            COSTOINDEXACION = "'" & txtCostoIndexacion.Text & "'"
        Else
            COSTOINDEXACION = "NULL"
        End If
 
        
        If Trim(txtCostoArmado.Text) <> "" Then
            COSTOARMADO = "'" & txtCostoArmado.Text & "'"
        Else
            COSTOARMADO = "NULL"
        End If
        
        
        
        
        
        If Trim(txtCostoCargaLegajo.Text) <> "" Then
            COSTOCARGALEGAJO = "'" & txtCostoCargaLegajo.Text & "'"
        Else
            COSTOCARGALEGAJO = "NULL"
        End If
    
         BARRA = "NULL"
         
         If optBarra_NRO_DESDE.value = True Then
            BARRA = "'NRO_DESDE'"
         End If
         
         If optBarra_NRO_HASTA.value = True Then
            BARRA = "'NRO_HASTA'"
         End If
         
         If optBarra_Letra_Desde = True Then
            BARRA = "'LETRA_DESDE'"
         End If
         
         If optBarra_Letra_Hasta = True Then
            BARRA = "'LETRA_HASTA'"
         End If
    
        If optEtiqueta_Legajo.value = True Then
            BARRA = "'ETIQUETA_LEGAJO'"
        End If
        
        
       If cboTipoArchivoExtencion.Text = "" Then
            TIPO_ARCHIVO = "Null"
       Else
            TIPO_ARCHIVO = "'" & cboTipoArchivoExtencion.Text & "'"
       End If
         
        
            HABILITAR_ETIQUETA_LEGAJO = chkEtiquetaLegajo.value
        
   
    If Operacion = "Modificar" Then
    
    Sql = " Update INDICES"
    Sql = Sql & vbCrLf & " SET "
    Sql = Sql & vbCrLf & " TIPO_INDICE ='" & Trim(cboTipo.Text) & "', INDICE ='" & TxtIndice.Text & "', DESCRIPCION =" & Descripcion
    Sql = Sql & vbCrLf & " , TITULO_FECHA_DESDE = " & TITULO_FECHA_DESDE & ", TITULO_FECHA_HASTA =" & TITULO_FECHA_HASTA
    Sql = Sql & vbCrLf & " , TITULO_LETRA_DESDE = " & TITULO_LETRA_DESDE & ", TITULO_LETRA_HASTA =" & TITULO_LETRA_HASTA
    Sql = Sql & vbCrLf & " , TITULO_NRO_DESDE =" & TITULO_NRO_DESDE & ", TITULO_NRO_HASTA =" & TITULO_NRO_HASTA
    Sql = Sql & vbCrLf & " , TITULO_DESCRIPCION =" & TITULO_DESCRIPCION & ", HABILITAR_FECHA_DESDE =" & HABILITAR_FECHA_DESDE
    Sql = Sql & vbCrLf & " , HABILITAR_FECHA_HASTA =" & HABILITAR_FECHA_HASTA & ", HABILITAR_LETRA_DESDE =" & HABILITAR_LETRA_DESDE
    Sql = Sql & vbCrLf & " , HABILITAR_LETRA_HASTA =" & HABILITAR_LETRA_HASTA & ", HABILITAR_NRO_DESDE =" & HABILITAR_NRO_DESDE
    Sql = Sql & vbCrLf & " , HABILITAR_NRO_HASTA =" & HABILITAR_NRO_HASTA & " , HABILITAR_DESCRIPCION =" & HABILITAR_DESCRIPCION
    Sql = Sql & vbCrLf & " , REQUERIR_FECHA_DESDE =" & REQUERIR_FECHA_DESDE & ",REQUERIR_FECHA_HASTA =" & REQUERIR_FECHA_HASTA
    Sql = Sql & vbCrLf & " , REQUERIR_LETRA_DESDE =" & REQUERIR_LETRA_DESDE & ",REQUERIR_LETRA_HASTA =" & REQUERIR_LETRA_HASTA
    Sql = Sql & vbCrLf & " , REQUERIR_NRO_DESDE =" & REQUERIR_NRO_DESDE & ",REQUERIR_NRO_HASTA =" & REQUERIR_NRO_HASTA
    Sql = Sql & vbCrLf & " , REQUERIR_DESCRIPCION =" & REQUERIR_DESCRIPCION
    Sql = Sql & vbCrLf & " , COPIAR_FECHA = " & COPIAR_FECHA
    Sql = Sql & vbCrLf & " , COPIAR_NRO = " & COPIAR_NRO
    Sql = Sql & vbCrLf & " , COPIAR_LETRA = " & COPIAR_LETRA
    Sql = Sql & vbCrLf & " , CONTROL_LOGICO_LARGO_NRO_DESDE_INICIO = " & CONTROL_LOGICO_LARGO_NRO_DESDE_INICIO
    Sql = Sql & vbCrLf & " , CONTROL_LOGICO_LARGO_NRO_DESDE_HASTA = " & CONTROL_LOGICO_LARGO_NRO_DESDE_HASTA
    Sql = Sql & vbCrLf & " , CONTROL_LOGICO_LARGO_NRO_HASTA_INICIO = " & CONTROL_LOGICO_LARGO_NRO_HASTA_INICIO
    Sql = Sql & vbCrLf & " , CONTROL_LOGICO_LARGO_NRO_HASTA_HASTA = " & CONTROL_LOGICO_LARGO_NRO_HASTA_HASTA
    Sql = Sql & vbCrLf & " , CONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO = " & CONTROL_LOGICO_LARGO_LETRA_DESDE_INICIO
    Sql = Sql & vbCrLf & " , CONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA = " & CONTROL_LOGICO_LARGO_LETRA_DESDE_HASTA
    Sql = Sql & vbCrLf & " , CONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO = " & CONTROL_LOGICO_LARGO_LETRA_HASTA_INICIO
    Sql = Sql & vbCrLf & " , CONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA = " & CONTROL_LOGICO_LARGO_LETRA_HASTA_HASTA
    Sql = Sql & vbCrLf & " , COSTOPREPARACION = " & COSTOPREPARACION
    Sql = Sql & vbCrLf & " , COSTODIGITALIZACION = " & COSTODIGITALIZACION
    Sql = Sql & vbCrLf & " , COSTOINDEXACION = " & COSTOINDEXACION
    Sql = Sql & vbCrLf & " , COSTOARMADO = " & COSTOARMADO
    Sql = Sql & vbCrLf & " , COSTOCARGALEGAJO = " & COSTOCARGALEGAJO
    Sql = Sql & vbCrLf & " , BARRA = " & BARRA
    Sql = Sql & vbCrLf & " , HABILITAR_ETIQUETA_LEGAJO = " & HABILITAR_ETIQUETA_LEGAJO
    Sql = Sql & vbCrLf & " , TIPO_ARCHIVO = " & TIPO_ARCHIVO
    Sql = Sql & vbCrLf & " Where ID =" & txt_ID.Text
    ExecutarSql Sql
End If

If Operacion = "Nuevo" Then


        If Trim(txtNro_Documento.Text) = "" Then
          
            Sql = "SELECT MAX(ID_CODIGO_DOCUMENTO) AS MAXDOC From INDICES Where COD_CLIENTE = " & ctlCliente.Valor
            Set RsMAXDOC = New ADODB.Recordset
            RsMAXDOC.Open Sql, ConActiva, 0, 1
            
            If IsNull(RsMAXDOC!Maxdoc) Then
                txtNro_Documento.Text = 1
            Else
                txtNro_Documento.Text = RsMAXDOC!Maxdoc + 1
            End If
        
        
        End If
        
        
        
           Sql = " INSERT INTO INDICES "
           Sql = Sql & vbCrLf & " ( "
           Sql = Sql & vbCrLf & "  COD_CLIENTE, ID_CODIGO_DOCUMENTO"
           Sql = Sql & vbCrLf & " , INDICE, DESCRIPCION, TIPO_INDICE"
           Sql = Sql & vbCrLf & " , TITULO_FECHA_DESDE, TITULO_FECHA_HASTA"
           Sql = Sql & vbCrLf & " , TITULO_LETRA_DESDE, TITULO_LETRA_HASTA"
           Sql = Sql & vbCrLf & " , TITULO_NRO_DESDE, TITULO_NRO_HASTA"
           Sql = Sql & vbCrLf & " , TITULO_DESCRIPCION, HABILITAR_FECHA_DESDE"
           Sql = Sql & vbCrLf & " , HABILITAR_FECHA_HASTA, HABILITAR_LETRA_DESDE"
           Sql = Sql & vbCrLf & " , HABILITAR_LETRA_HASTA, HABILITAR_NRO_DESDE"
           Sql = Sql & vbCrLf & " , HABILITAR_NRO_HASTA,HABILITAR_DESCRIPCION"
           Sql = Sql & vbCrLf & " , REQUERIR_FECHA_DESDE, REQUERIR_FECHA_HASTA"
           Sql = Sql & vbCrLf & " , REQUERIR_LETRA_DESDE, REQUERIR_LETRA_HASTA"
           Sql = Sql & vbCrLf & " , REQUERIR_NRO_DESDE, REQUERIR_NRO_HASTA"
           Sql = Sql & vbCrLf & " , REQUERIR_DESCRIPCION"
           Sql = Sql & vbCrLf & " , COPIAR_FECHA, COPIAR_LETRA"
           Sql = Sql & vbCrLf & " , COPIAR_NRO "
           Sql = Sql & vbCrLf & " ) "
           Sql = Sql & vbCrLf & " VALUES ("
           Sql = Sql & vbCrLf & ctlCliente.Valor & "," & txtNro_Documento.Text
           Sql = Sql & vbCrLf & " ,'" & Trim(TxtIndice.Text) & "'," & Descripcion & ",'" & cboTipo.Text & "'"
           Sql = Sql & vbCrLf & " ," & TITULO_FECHA_DESDE & "," & TITULO_FECHA_HASTA
           Sql = Sql & vbCrLf & " ," & TITULO_LETRA_DESDE & "," & TITULO_LETRA_HASTA
           Sql = Sql & vbCrLf & " ," & TITULO_NRO_DESDE & "," & TITULO_NRO_HASTA
           Sql = Sql & vbCrLf & " ," & TITULO_DESCRIPCION & " ," & HABILITAR_FECHA_DESDE
           Sql = Sql & vbCrLf & " ," & HABILITAR_FECHA_HASTA & "," & HABILITAR_LETRA_DESDE
           Sql = Sql & vbCrLf & " ," & HABILITAR_LETRA_HASTA & "," & HABILITAR_NRO_DESDE
           Sql = Sql & vbCrLf & " ," & HABILITAR_NRO_HASTA & "," & HABILITAR_DESCRIPCION
           Sql = Sql & vbCrLf & " ," & REQUERIR_FECHA_DESDE & "," & REQUERIR_FECHA_HASTA
           Sql = Sql & vbCrLf & " ," & REQUERIR_LETRA_DESDE & "," & REQUERIR_LETRA_HASTA
           Sql = Sql & vbCrLf & " ," & REQUERIR_NRO_DESDE & "," & REQUERIR_NRO_HASTA
           Sql = Sql & vbCrLf & " ," & REQUERIR_DESCRIPCION
           Sql = Sql & vbCrLf & " ," & COPIAR_FECHA & "," & COPIAR_LETRA
           Sql = Sql & vbCrLf & " ," & COPIAR_NRO
           Sql = Sql & vbCrLf & " )"
           ExecutarSql Sql
        
    End If
Exit Sub

salir:

MsgBox Err.Description



End Sub

Private Sub txtIndice_DblClick()
 Clipboard.SetText TxtIndice.Text
End Sub

Public Sub INSE()
Dim rs As New ADODB.Recordset
Dim Sql As String
Dim C As New ADODB.Connection
C.Open strConBasa
Sql = " SELECT INDICE_SUCURSAL, DOCUMENTO, DESCRIPCION "
Sql = Sql & " From DISCO "
Sql = Sql & " ORDER BY INDICE_SUCURSAL "

rs.Open Sql, ConActiva, 0, 1

Dim ID As Integer
ID = 6206


Do While Not rs.EOF
ID = ID + 1
    Sql = "INSERT INTO INDICES"
    Sql = Sql & " ( COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, DESCRIPCION, TIPO_INDICE)"
Sql = Sql & "  VALUES (281," & rs!Documento & ",'" & Trim(rs!INDICE_SUCURSAL) & "','" & UCase(Trim(rs!Descripcion)) & "','Sector')"

C.Execute Sql

    rs.MoveNext
Loop




    
    
    


End Sub
