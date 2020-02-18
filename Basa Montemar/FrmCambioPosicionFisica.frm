VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCambioPosicionFisica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Posición"
   ClientHeight    =   8775
   ClientLeft      =   435
   ClientTop       =   630
   ClientWidth     =   11940
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
   ScaleHeight     =   8775
   ScaleWidth      =   11940
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   14420
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "FrmCambioPosicionFisica.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblContar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLectura"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCantidadTotal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblFecha"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "pbsCambios"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DataGrid1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "grdCambioPosicion"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdRotuloChico"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdAceptar"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdImprimir"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdCancelar"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdContenedorModulos"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtLectura"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdAnularPociciones"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdAnularPosiciones"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdReporteEstanteria"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdCambioModulo"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtModulo_Vertical"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtModulo_Horizontal"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdContenedor"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdVerPosiciones"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "chkLiberarModulo"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtTomarLectura"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdSql"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Command1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cmdverificar"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FrmCambioPosicionFisica.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FrmCambioPosicionFisica.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   3195
         Left            =   -74700
         TabIndex        =   52
         Top             =   720
         Width           =   9315
         Begin VB.CommandButton cmdAceptar_Cambio 
            Caption         =   "Acepta"
            Height          =   375
            Left            =   7620
            TabIndex        =   55
            Top             =   2640
            Width           =   1395
         End
         Begin VB.Frame Frame4 
            Caption         =   "Caja a posicionar"
            Height          =   2235
            Left            =   4680
            TabIndex        =   54
            Top             =   300
            Width           =   4395
            Begin VB.TextBox txtClienteFinal 
               Height          =   375
               Left            =   1320
               TabIndex        =   64
               Top             =   540
               Width           =   1995
            End
            Begin VB.TextBox txtCajaFinal 
               Height          =   375
               Left            =   1320
               TabIndex        =   62
               Top             =   1140
               Width           =   1995
            End
            Begin VB.Label Label12 
               Caption         =   "Cliente"
               Height          =   315
               Left            =   240
               TabIndex        =   63
               Top             =   600
               Width           =   795
            End
            Begin VB.Label Label11 
               Caption         =   "Caja:"
               Height          =   315
               Left            =   240
               TabIndex        =   61
               Top             =   1200
               Width           =   795
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Caja a liberar"
            Height          =   2235
            Left            =   180
            TabIndex        =   53
            Top             =   360
            Width           =   4215
            Begin VB.TextBox txtClienteLiberar 
               Height          =   375
               Left            =   1440
               TabIndex        =   59
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtCajaLiberar 
               Height          =   375
               Left            =   1440
               TabIndex        =   57
               Top             =   1080
               Width           =   1695
            End
            Begin VB.Label lblID_Posicion 
               Height          =   435
               Left            =   1380
               TabIndex        =   60
               Top             =   1680
               Width           =   1335
            End
            Begin VB.Label Label10 
               Caption         =   "Cliente"
               Height          =   315
               Left            =   360
               TabIndex        =   58
               Top             =   540
               Width           =   795
            End
            Begin VB.Label Label9 
               Caption         =   "Caja:"
               Height          =   315
               Left            =   360
               TabIndex        =   56
               Top             =   1140
               Width           =   795
            End
         End
      End
      Begin VB.CommandButton cmdverificar 
         Caption         =   "Verificar"
         Height          =   375
         Left            =   7320
         TabIndex        =   42
         Top             =   2460
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Colector"
         Height          =   375
         Left            =   8760
         TabIndex        =   41
         Top             =   2460
         Width           =   1200
      End
      Begin VB.CommandButton cmdSql 
         Caption         =   "Sql"
         Height          =   375
         Left            =   10200
         TabIndex        =   40
         Top             =   2460
         Width           =   1200
      End
      Begin VB.TextBox txtTomarLectura 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   8880
         PasswordChar    =   "*"
         TabIndex        =   39
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CheckBox chkLiberarModulo 
         Caption         =   "Liberar Modulo"
         Height          =   255
         Left            =   7320
         TabIndex        =   38
         Top             =   540
         Width           =   1695
      End
      Begin VB.CommandButton cmdVerPosiciones 
         Caption         =   "Ver Posiciones LIBRES"
         Height          =   330
         Left            =   2460
         TabIndex        =   37
         Top             =   3360
         Width           =   3855
      End
      Begin VB.CommandButton cmdContenedor 
         Caption         =   "Contenedor"
         Height          =   345
         Left            =   4800
         TabIndex        =   34
         Top             =   2460
         Width           =   1515
      End
      Begin VB.TextBox txtModulo_Horizontal 
         Height          =   330
         Left            =   3600
         TabIndex        =   33
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtModulo_Vertical 
         Height          =   330
         Left            =   2340
         TabIndex        =   32
         Top             =   1980
         Width           =   1095
      End
      Begin VB.CommandButton cmdCambioModulo 
         Caption         =   "Cambio Modulo"
         Enabled         =   0   'False
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
         Left            =   4740
         TabIndex        =   31
         Top             =   1980
         Width           =   1515
      End
      Begin VB.CommandButton cmdReporteEstanteria 
         Caption         =   "Reporte Estanteria"
         Height          =   330
         Left            =   120
         TabIndex        =   30
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton cmdAnularPosiciones 
         Caption         =   "Anular Ubicacion"
         Height          =   330
         Left            =   2460
         TabIndex        =   29
         Top             =   2940
         Width           =   1935
      End
      Begin VB.CommandButton cmdAnularPociciones 
         Caption         =   "Posiciones sin uso"
         Height          =   330
         Left            =   120
         TabIndex        =   28
         Top             =   2940
         Width           =   2175
      End
      Begin VB.TextBox txtLectura 
         Height          =   330
         Left            =   120
         TabIndex        =   27
         Top             =   1980
         Width           =   1875
      End
      Begin VB.CommandButton cmdContenedorModulos 
         Caption         =   "Contenedor Modulos"
         Height          =   375
         Left            =   4440
         TabIndex        =   26
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ubicación"
         Height          =   1515
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   6135
         Begin VB.Frame fraEstanteria 
            Caption         =   "Estanteria"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Left            =   720
            TabIndex        =   21
            Top             =   240
            Width           =   1275
            Begin VB.TextBox txtEstanteria_Desde 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   180
               TabIndex        =   23
               Text            =   "316"
               Top             =   300
               Width           =   900
            End
            Begin VB.TextBox txtEstanteria_Hasta 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   180
               TabIndex        =   22
               Text            =   "318"
               Top             =   660
               Width           =   900
            End
         End
         Begin VB.Frame fraHorizontal 
            Caption         =   "Horizontal"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Left            =   3480
            TabIndex        =   18
            Top             =   240
            Width           =   1275
            Begin VB.TextBox txtHorizontal_Desde 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   180
               TabIndex        =   20
               Text            =   "1"
               Top             =   300
               Width           =   900
            End
            Begin VB.TextBox txtHorizontal_Hasta 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   180
               TabIndex        =   19
               Text            =   "27"
               Top             =   660
               Width           =   900
            End
         End
         Begin VB.Frame fraVertical 
            Caption         =   "Vertical"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1155
            Left            =   2100
            TabIndex        =   15
            Top             =   240
            Width           =   1275
            Begin VB.TextBox txtVertical_Desde 
               Alignment       =   2  'Center
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
               Left            =   180
               TabIndex        =   17
               Text            =   "1"
               Top             =   300
               Width           =   960
            End
            Begin VB.TextBox txtVertical_Hasta 
               Alignment       =   2  'Center
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
               Left            =   180
               TabIndex        =   16
               Text            =   "50"
               Top             =   660
               Width           =   960
            End
         End
         Begin VB.Frame FraPosicion 
            Caption         =   "Posición"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   4800
            TabIndex        =   11
            Top             =   180
            Width           =   1155
            Begin VB.OptionButton optAmbos 
               Caption         =   "Ambos"
               Height          =   330
               Left            =   120
               TabIndex        =   14
               Top             =   780
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton optAtras 
               Caption         =   "Atras"
               Height          =   330
               Left            =   120
               TabIndex        =   13
               Top             =   480
               Width           =   915
            End
            Begin VB.OptionButton optFrente 
               Caption         =   "Frente"
               Height          =   330
               Left            =   120
               TabIndex        =   12
               Top             =   180
               Width           =   855
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Desde"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   420
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "Hasta"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   8940
         TabIndex        =   9
         Top             =   7500
         Width           =   1395
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Etiqueta T"
         Height          =   375
         Left            =   4380
         TabIndex        =   8
         Top             =   7500
         Width           =   1155
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7440
         TabIndex        =   7
         Top             =   7500
         Width           =   1395
      End
      Begin VB.CommandButton cmdRotuloChico 
         Caption         =   "Imprimir Rotulo chico"
         Height          =   375
         Left            =   1860
         TabIndex        =   6
         Top             =   7500
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Imprimir "
         Height          =   375
         Left            =   60
         TabIndex        =   5
         Top             =   7500
         Width           =   1635
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Etiqueta "
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         Top             =   7500
         Width           =   1155
      End
      Begin MSFlexGridLib.MSFlexGrid grdCambioPosicion 
         Height          =   3495
         Left            =   180
         TabIndex        =   2
         Top             =   3840
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   6165
         _Version        =   393216
         Cols            =   6
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3495
         Left            =   180
         TabIndex        =   3
         Top             =   3840
         Visible         =   0   'False
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   6165
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
      Begin MSComctlLib.ProgressBar pbsCambios 
         Height          =   375
         Left            =   2400
         TabIndex        =   35
         Top             =   2460
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         ForeColor       =   &H80000007&
         Height          =   255
         Index           =   0
         Left            =   9360
         TabIndex        =   51
         Top             =   540
         Width           =   855
      End
      Begin VB.Label lblFecha 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10/10/2000"
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   10260
         TabIndex        =   50
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad Total"
         ForeColor       =   &H80000007&
         Height          =   255
         Left            =   7440
         TabIndex        =   49
         Top             =   2100
         Width           =   1215
      End
      Begin VB.Label lblCantidadTotal 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   8880
         TabIndex        =   48
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label lblLectura 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   8880
         TabIndex        =   47
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblContar 
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   8880
         TabIndex        =   46
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Lectura :"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   7440
         TabIndex        =   45
         Top             =   1020
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Proceso : "
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   7440
         TabIndex        =   44
         Top             =   1380
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Lectura Manual:"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   7440
         TabIndex        =   43
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Lectura de Cajas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   2460
         Width           =   1395
      End
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cambio de Ubicación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   10275
   End
End
Attribute VB_Name = "frmCambioPosicionFisica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bloque As Boolean

Private Sub cmdAceptar_Cambio_Click()

lblID_Posicion.Caption = ""
Dim ID_Poscicion_Liberada As Long
Dim SQL As String

            SQL = " INSERT INTO CAMBIOPOSICION"
            SQL = SQL & " (ID_PERSONAL, FECHA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA)"
            SQL = SQL & "  SELECT    " & MDIfrmInicio.StaInicio.Panels(2).Text & ", GETDATE() AS FECHA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO,"
            SQL = SQL & "  COD_CLIENTE , NRO_CAJA"
            SQL = SQL & "  From CONTENEDOR"
            SQL = SQL & "  Where COD_CLIENTE = " & txtClienteLiberar.Text
            SQL = SQL & "  and  NRO_CAJA = " & txtCajaLiberar.Text
            ExecutarSql SQL
                        
            
            SQL = " INSERT INTO CAMBIOPOSICION"
            SQL = SQL & " (ID_PERSONAL, FECHA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA)"
            SQL = SQL & "  SELECT    " & MDIfrmInicio.StaInicio.Panels(2).Text & ", GETDATE() AS FECHA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO,"
            SQL = SQL & "  COD_CLIENTE , NRO_CAJA"
            SQL = SQL & "  From CONTENEDOR"
            SQL = SQL & "  Where COD_CLIENTE = " & txtClienteFinal.Text
            SQL = SQL & "  and  NRO_CAJA = " & txtCajaFinal.Text

            ExecutarSql SQL
            
            Dim RS As New ADODB.Recordset
            
            SQL = " SELECT  ID_CONTENEDOR, COD_CLIENTE, NRO_CAJA"
            SQL = SQL & "  From basasql.dbo.CONTENEDOR"
            SQL = SQL & "  Where COD_CLIENTE = " & txtClienteLiberar.Text
            SQL = SQL & "  and  NRO_CAJA = " & txtCajaLiberar.Text
            Set RS = New ADODB.Recordset
            
            RS.Open SQL, strConBasa
            
            If Not RS.EOF Then
                lblID_Posicion.Caption = RS!ID_CONTENEDOR
            Else
                Exit Sub
            End If
            
            
            SQL = " SELECT     TOP (1) ID_CONTENEDOR, ESTANTERIA, NRO_CAJA"
            SQL = SQL & " From CONTENEDOR"
            SQL = SQL & " Where (Estanteria > 150) And (estado = 1) And (COD_CLIENTE Is Null)"
            SQL = SQL & " ORDER BY ESTANTERIA"
            
            
            
            Set RS = New ADODB.Recordset
            
            RS.Open SQL, strConBasa
            
            
            If Not RS.EOF Then
                ID_Poscicion_Liberada = RS!ID_CONTENEDOR
            
            End If
            
            
            
            
         Set RS = New ADODB.Recordset
            
            RS.Open SQL, strConBasa
            
            If Not RS.EOF Then
                SQL = " UPDATE    TOP (2) CONTENEDOR"
                SQL = SQL & " SET NRO_CAJA = NULL, ESTADO = 1, COD_CLIENTE = NULL"
                SQL = SQL & "  Where "
                SQL = SQL & " COD_CLIENTE = " & txtClienteLiberar.Text
                SQL = SQL & "  And NRO_CAJA = " & txtCajaLiberar.Text
                ExecutarSql SQL
                
                SQL = " UPDATE    TOP (2) CONTENEDOR"
                SQL = SQL & " SET NRO_CAJA = " & txtCajaLiberar.Text
                SQL = SQL & " , ESTADO = 2 "
                SQL = SQL & " , COD_CLIENTE = " & txtClienteLiberar.Text
                SQL = SQL & "  Where "
                SQL = SQL & " ID_CONTENEDOR = " & ID_Poscicion_Liberada
                ExecutarSql SQL
               
                SQL = " UPDATE    TOP (2) CONTENEDOR"
                SQL = SQL & " SET  NRO_CAJA = NULL, ESTADO = 1, COD_CLIENTE = NULL"
                SQL = SQL & "  Where "
                SQL = SQL & " COD_CLIENTE = " & txtClienteFinal.Text
                SQL = SQL & "  And NRO_CAJA = " & txtCajaFinal.Text
                ExecutarSql SQL
               
               SQL = " UPDATE    TOP (2) CONTENEDOR"
                SQL = SQL & " SET NRO_CAJA = " & txtCajaFinal.Text
                SQL = SQL & " , ESTADO = 2 "
                SQL = SQL & " , COD_CLIENTE = " & txtClienteFinal.Text
                SQL = SQL & "  Where "
                SQL = SQL & " ID_CONTENEDOR = " & lblID_Posicion.Caption
                ExecutarSql SQL
               
               
            End If
            
            
            
    MsgBox "TERMINADO"
            
            
            
End Sub

Private Sub cmdAceptar_Click()

Dim I As Integer
    If Validar Then
        MousePointer = 11
        MovimientoCajas1
        cmdAceptar.Enabled = False
        MousePointer = 0
    End If
End Sub

Private Sub cmdAnularPociciones_Click()
 Dim SQL As String
 Dim RS   As New ADODB.Recordset

SQL = " SELECT     ESTANTERIA, MODULO_H, COUNT(*) AS cant, VERTICAL"
SQL = SQL & " From basasql.dbo.CONTENEDOR "
SQL = SQL & " Where (estado = 1)"
SQL = SQL & " GROUP BY ESTANTERIA, MODULO_H, VERTICAL"
SQL = SQL & " HAVING      (ESTANTERIA between  " & txtEstanteria_Desde.Text & " and " & txtEstanteria_Hasta.Text & ")"
SQL = SQL & " AND (COUNT(*) = 5) AND (MODULO_H IN (1, 2, 3))"
SQL = SQL & " ORDER BY ESTANTERIA, MODULO_H, VERTICAL"
RS.Open SQL, strConBasa
Do While Not RS.EOF
        SQL = " UPDATE    CONTENEDOR"
        SQL = SQL & " SET   ESTADO =0 "
        SQL = SQL & " Where Estanteria = " & RS!Estanteria
        SQL = SQL & " And (COD_CLIENTE Is Null)"
        SQL = SQL & " And (estado = 1) "
        SQL = SQL & " And Modulo_H = " & RS!Modulo_H
        SQL = SQL & " And Vertical =" & RS!Vertical
        ExecutarSql SQL
        
    RS.MoveNext
    
 Loop
 MsgBox "Listo"


End Sub

Private Sub cmdAnularPosiciones_Click()
Dim SQL As String
SQL = " Update basasql.dbo.CONTENEDOR"
SQL = SQL & " Set estado = 0 "
SQL = SQL & "  WHERE     ESTANTERIA BETWEEN " & txtEstanteria_Desde.Text & " AND  " & txtEstanteria_Hasta.Text
SQL = SQL & "  AND HORIZONTAL BETWEEN " & txtHorizontal_Desde.Text & "  AND " & txtEstanteria_Hasta.Text
SQL = SQL & "  AND VERTICAL BETWEEN " & txtVertical_Desde.Text & " AND  " & txtVertical_Hasta.Text
SQL = SQL & "  AND (ESTADO = 1)"
ExecutarSql SQL
MsgBox "Terminado"


End Sub

Private Sub cmdCambioModulo_Click()

Dim SQL As String
Dim con As New ADODB.Connection
con.Open strConBasa



SQL = " UPDATE    CONTENEDOR "
SQL = SQL & " SET   MODULO_V = " & txtModulo_Vertical.Text
SQL = SQL & " , MODULO_H = " & txtModulo_Horizontal.Text
SQL = SQL & " WHERE ESTANTERIA =  " & txtEstanteria_Desde.Text
SQL = SQL & " AND VERTICAL BETWEEN " & txtVertical_Desde.Text & "  AND " & txtVertical_Hasta.Text
SQL = SQL & " AND HORIZONTAL BETWEEN " & txtHorizontal_Desde.Text & " AND " & txtHorizontal_Hasta.Text


con.Execute SQL
MsgBox "TERMINADO"


End Sub

Private Sub cmdCancelar_Click()
    grdCambioPosicion.Clear
    TituloGrilla
End Sub

Private Sub CmdConfigurarEstanteria_Click()

'Dim RS As New ADODB.Recordset
'Dim rs2 As New ADODB.Recordset
'Dim sql As String
''SQL = " SELECT     MAX(VERTICAL) AS VMAX, MIN(VERTICAL) AS VMIN, MAX(HORIZONTAL) AS HMAX, MIN(HORIZONTAL) AS HMIN"
''SQL = SQL & " From CONTENEDOR "
''SQL = SQL & "  Where Estanteria = " & txtEstanteriaCustodia.Text
''SQL = SQL & "  And Modulo_V = " & txtV.Text
''SQL = SQL & "  And Modulo_H = " & txtH.Text
'
'
'
''Set RS = New ADODB.Recordset
''
''RS.Open SQL, strConBasa , 0 ,1
''
''
''If Not RS.EOF Then
''    txtDesde.Text = txtEstanteriaCustodia.Text
''    txtHasta.Text = txtEstanteriaCustodia.Text
''    If IsNull(RS!HMin) Then
''        MsgBox "Mal el rango"
''        txtHorizontalDesde.Text = 0
''        txtHorizontalHasta.Text = 0
''        txtVerticalDesde.Text = 0
''        txtVerticalHasta.Text = 0
''        optAmbos.value = True
''        Exit Sub
''    Else
''        txtHorizontalDesde.Text = RS!HMin
''        txtHorizontalHasta.Text = RS!HMAX
''        txtVerticalDesde.Text = RS!VMIN
''        txtVerticalHasta.Text = RS!VMAX
''        optAmbos.value = True
''    End If
''End If
''
''
''
''
''SQL = " SELECT     codCLIENTE, CAJA"
''SQL = SQL & " From contenedor "
''SQL = SQL & " Where Estanteria = " & txtEstanteriaCustodia.Text
''SQL = SQL & " AND  H = " & txtH.Text
''SQL = SQL & " AND  V = " & txtV.Text
''SQL = SQL & " ORDER BY ORDEN "
''
'sql = " SELECT  ESTANTERIA,HORIZONTAL,VERTICAL,COD_CLIENTE,NRO_CAJA "
'
'
'
'
'sql = sql & " From CONTENEDOR"
'sql = sql & " Where Estanteria = " & txtEstanteriaCustodia.Text
'
'If txtH.Text <> "" Then
'    sql = sql & " And Modulo_H = " & txtH.Text
'End If
'
'If txtV.Text <> "" Then
'sql = sql & " And Modulo_V = " & txtV.Text
'End If
'
'sql = sql & " ORDER BY HORIZONTAL, VERTICAL, ADELANTE_ATRAS"
'
'
'frmReportes.ImprimirReporte PasoReportes & "rptRotuloEtiquetaCOrdoba.rpt", sql, True
'
'
''
''  Set rs2 = New ADODB.Recordset
''
''
''     rs2.Open SQL, strConBasa , 0 ,1
''
''Dim I As Integer
''
''
''
''
''    Dim Cliente As Integer
''    Dim Caja As Long
''
''
''I = 1
''TituloGrilla
''
''grdCambioPosicion.ColWidth(0) = 500
''
''    Do While Not rs2.EOF
''
''    If rs2!NRO_CAJA > 700000 Then
''        Cliente = BuscarCliente(rs2!NRO_CAJA)
''    Else
''        Cliente = CInt(rs2!COD_CLIENTE)
''    End If
''
''
''        Caja = CLng(rs2!NRO_CAJA)
''
''        Set RS = New ADODB.Recordset
''         RS.Open "SELECT * FROM CLIENTES WHERE ID_CLIENTE = " & Cliente, strConBasa , 0 ,1
''        If Not RS.EOF Then
''            lblIDCliente = RS!id_cliente
''            lblCliente = Trim(UCase(RS!RAZON_SOCIAL))
''        End If
''            grdCambioPosicion.AddItem I & vbTab & Caja & vbTab & Cliente & vbTab & lblCliente
''            I = I + 1
''            ContarGrilla grdCambioPosicion, lblCantidadTotal
''        rs2.MoveNext
''    Loop
''
''grdCambioPosicion.Visible = True
''
'' DataGrid1.Visible = False
''


End Sub

Private Sub cmdContenedor_Click()
Dim SQL As String
Dim RS As New ADODB.Recordset

DataGrid1.Visible = True
grdCambioPosicion.Visible = False
SQL = " SELECT     ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,  ESTADO, COD_CLIENTE, NRO_CAJA , MODULO_V , MODULO_H  "
SQL = SQL & " From CONTENEDOR "
SQL = SQL & " WHERE  ESTANTERIA BETWEEN " & txtEstanteria_Desde.Text & " AND " & txtEstanteria_Hasta.Text
SQL = SQL & " AND HORIZONTAL BETWEEN  " & txtHorizontal_Desde.Text & " AND  " & txtHorizontal_Hasta.Text
SQL = SQL & " AND VERTICAL BETWEEN " & txtVertical_Desde.Text & "  AND " & txtVertical_Hasta.Text
SQL = SQL & " ORDER BY ESTANTERIA, VERTICAL,  HORIZONTAL "

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open SQL, ConActiva, 2, 3


Set DataGrid1.DataSource = RS
DataGrid1.ReBind
DataGrid1.Refresh
CopiarDatosGrilla DataGrid1

End Sub

Private Sub cmdContenedorModulos_Click()
Dim SQL As String
Dim RS As New ADODB.Recordset

DataGrid1.Visible = True
grdCambioPosicion.Visible = False
SQL = " SELECT     ESTANTERIA, HORIZONTAL, VERTICAL,  ESTADO, COD_CLIENTE, NRO_CAJA , MODULO_V , MODULO_H  "
SQL = SQL & " From CONTENEDOR "
SQL = SQL & " WHERE  ESTANTERIA = " & txtEstanteria_Desde.Text
SQL = SQL & " AND Modulo_V = " & txtModulo_Vertical.Text
SQL = SQL & " AND Modulo_H = " & txtModulo_Horizontal.Text
SQL = SQL & " ORDER BY ESTANTERIA, VERTICAL,  HORIZONTAL "

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open SQL, ConActiva, 2, 3


Set DataGrid1.DataSource = RS
DataGrid1.ReBind
DataGrid1.Refresh
CopiarDatosGrilla DataGrid1


 
End Sub

Private Sub cmdImprimir_Click()
If Trim(lblLectura.Caption) <> "" Then
     ImprimirRotulosLectura lblLectura.Caption, True
End If
End Sub

Private Sub cmdReporteEstanteria_Click()
    Dim SQL As String
        SQL = " SELECT     *"
        SQL = SQL & vbCrLf & " From basasql.dbo.CONTENEDOR"
        SQL = SQL & vbCrLf & " WHERE     (ESTANTERIA >= " & txtEstanteria_Desde.Text & " AND ESTANTERIA <= " & txtEstanteria_Hasta.Text & ")"
        frmReportes.ImprimirReporte PasoReportes & "rptControlEstanteria.rpt", SQL, True
End Sub

Private Sub cmdRotuloChico_Click()
        If Trim(lblLectura.Caption) <> "" Then
             ImprimirRotulosLectura lblLectura.Caption, False
        End If
End Sub

Private Sub cmdSql_Click()
' Dim rs2 As New ADODB.Recordset
'    Dim Sql As String
'    Dim Caja As Long
'    Dim Cliente As Integer
'
'    lblIDCliente = ""
'    Dim E As Integer
'    Sql = "SELECT * From CONTENEDOR WHERE ESTANTERIA = " & InputBox("INGRESE LA ESTANTERIA")
'    Sql = Sql & " and  not nro_caja is null"
'    Sql = Sql & " order by vertical , horizontal "
'     rs2.Open Sql, ConActiva, 0, 1
'    E = 0
'
'   DataGrid1.Visible = False
'    grdCambioPosicion.Visible = True
'   Do While Not rs2.EOF
'        Caja = rs2!NRO_CAJA
'        Cliente = rs2!COD_CLIENTE
'        grdCambioPosicion.AddItem "" & vbTab & Caja & vbTab & Cliente & vbTab & "lblCliente"
'        ContarGrilla grdCambioPosicion, lblCantidadTotal
'        rs2.MoveNext
'    Loop
End Sub

Private Sub cmdverificar_Click()

Verificar

'    Dim rsPosicionLibre As New ADODB.Recordset
'    Dim R As Integer
'    Dim C As Integer
'    Dim Sql  As String
'    Dim Sql2 As String
'
'    Dim RS As ADODB.Recordset
'    Contar = 0
'    If cmdAceptar.Enabled Then
'        MsgBox "Error en Procesos", vbCritical
'        Exit Sub
'    End If
'    DataGrid1.Visible = False
'    grdCambioPosicion.Visible = True
'
'    MousePointer = 11
'      If Trim(txtDesde) <> "" Then
'
'            If optAtras.value = True Then
'                Sql = "Select * from contenedor where estado = 1 and ( cod_cliente=0 or cod_cliente is null) "
'                Sql = Sql & " AND (estanteria Between  " & txtDesde & "  AND " & txtHasta
'                Sql = Sql & ") AND ( HORIZONTAL Between " & txtHorizontalDesde & " AND " & txtHorizontalHasta & ")"
'                Sql = Sql & " AND (VERTICAL Between " & txtVerticaldesde & " AND " & txtVerticalHasta & ") "
'                Sql = Sql & " AND Adelante_Atras = 1 "
'                Sql = Sql & " Order by estanteria,vertical,Horizontal"
'
'                Sql2 = "Select count(*) as contar from contenedor where estado = 1 and ( cod_cliente=0 or cod_cliente is null) "
'                Sql2 = Sql2 & " AND (estanteria Between  " & txtDesde & "  AND " & txtHasta
'                Sql2 = Sql2 & ") AND ( HORIZONTAL Between " & txtHorizontalDesde & " AND " & txtHorizontalHasta & ")"
'                Sql2 = Sql2 & " AND ( VERTICAL Between " & txtVerticaldesde & " AND " & txtVerticalHasta & ") "
'                Sql2 = Sql2 & " AND Adelante_Atras = 1 "
'            Else
'                If optFrente.value = True Then
'                        Sql = "Select * from contenedor where estado = 1 and ( cod_cliente=0 or cod_cliente is null) "
'                        Sql = Sql & " AND ( estanteria Between  " & txtDesde & "  AND " & txtHasta
'                        Sql = Sql & ")AND ( HORIZONTAL Between " & txtHorizontalDesde & " AND " & txtHorizontalHasta & ")"
'                        Sql = Sql & " AND ( VERTICAL Between " & txtVerticaldesde & " AND " & txtVerticalHasta & ") "
'                        Sql = Sql & " AND Adelante_Atras = 2 "
'                        Sql = Sql & " Order by estanteria,vertical,Horizontal"
'
'                        Sql2 = "Select count(*) as contar from contenedor where estado = 1 and ( cod_cliente=0 or cod_cliente is null) "
'                        Sql2 = Sql2 & " AND ( estanteria Between  " & txtDesde & "  AND " & txtHasta
'                        Sql2 = Sql2 & ")AND ( HORIZONTAL Between " & txtHorizontalDesde & " AND " & txtHorizontalHasta & ")"
'                        Sql2 = Sql2 & " AND ( VERTICAL Between " & txtVerticaldesde & " AND " & txtVerticalHasta & ") "
'                        Sql2 = Sql2 & " AND Adelante_Atras = 2 "
'                Else
'                        Sql = "Select * from contenedor where estado = 1 and ( cod_cliente=0 or cod_cliente is null) "
'                        Sql = Sql & " AND ( estanteria Between  " & txtDesde & "  AND " & txtHasta
'                        Sql = Sql & ")AND ( HORIZONTAL Between " & txtHorizontalDesde & " AND " & txtHorizontalHasta & ")"
'                        Sql = Sql & " AND ( VERTICAL Between " & txtVerticaldesde & " AND " & txtVerticalHasta & ") "
'                        Sql = Sql & " Order by estanteria, vertical,adelante_atras DESC,horizontal"
'
'                        Sql2 = "Select count(*) as contar from contenedor where estado = 1 and ( cod_cliente=0 or cod_cliente is null) "
'                        Sql2 = Sql2 & " AND (estanteria Between  " & txtDesde & "  AND " & txtHasta
'                        Sql2 = Sql2 & ") AND ( HORIZONTAL Between " & txtHorizontalDesde & " AND " & txtHorizontalHasta & ")"
'                        Sql2 = Sql2 & " AND ( VERTICAL Between " & txtVerticaldesde & " AND " & txtVerticalHasta & ") "
'                End If
'            End If
'           Else
'           MsgBox "ingrese desde"
'           Exit Sub
'           End If
'         rsPosicionLibre.Open Sql2, ConActiva, 0, 1
'        lblContar.Caption = "CONTROL DE CANTIDAD"
'        If Not rsPosicionLibre.EOF Then
'        If lblCantidadTotal.Caption = "" Then
'        Exit Sub
'
'        End If
'
'            If lblCantidadTotal > CLng(rsPosicionLibre!Contar) Then
'                MsgBox "No Alcanzan las posiciones" & vbCrLf & "Solo quedan para ese rango " & CStr(rsPosicionLibre!Contar) & " posiciones", vbCritical
'                MousePointer = 0
'                Exit Sub
'            End If
'        End If
'
'      Set rsPosicionLibre = New ADODB.Recordset
'      rsPosicionLibre.Open Sql, ConActiva, 0, 1
'        Set Ccajas = New Cajas
'        lblContar.Caption = "CONTROL DE POSICIONES "
'        Do While Not rsPosicionLibre.EOF
'            For R = 1 To grdCambioPosicion.Rows
'                If R >= grdCambioPosicion.Rows Then
'                    Exit Do
'                End If
'                With rsPosicionLibre
'                    Ccajas.Add CStr(grdCambioPosicion.TextMatrix(R, 1)), CInt(!Estanteria), CInt(!Horizontal), CInt(!Vertical), CInt(!Adelante_Atras), CInt(!NRO_ESTANTE), 0, grdCambioPosicion.TextMatrix(R, 2), grdCambioPosicion.TextMatrix(R, 1)
'                    rsPosicionLibre.MoveNext
'                End With
'            Next
'        Loop
'        grdCambioPosicion.Cols = 8
'        grdCambioPosicion.Clear
'        grdCambioPosicion.ColWidth(0) = 200
'        grdCambioPosicion.ColWidth(1) = 800
'        grdCambioPosicion.ColWidth(2) = 3500
'        grdCambioPosicion.ColWidth(3) = 800
'        grdCambioPosicion.ColWidth(4) = 800
'        grdCambioPosicion.ColWidth(5) = 800
'        grdCambioPosicion.ColWidth(6) = 800
'        grdCambioPosicion.ColWidth(7) = 1100
'
'        grdCambioPosicion.Font.Size = 9
'        grdCambioPosicion.TextMatrix(0, 1) = "ID_CLI"
'        grdCambioPosicion.TextMatrix(0, 2) = "RAZON SOCIAL"
'        grdCambioPosicion.TextMatrix(0, 3) = "CAJA"
'        grdCambioPosicion.TextMatrix(0, 4) = "EST."
'        grdCambioPosicion.TextMatrix(0, 5) = "HOR."
'        grdCambioPosicion.TextMatrix(0, 6) = "VER."
'        grdCambioPosicion.TextMatrix(0, 7) = "AD/AT"
'
'        grdCambioPosicion.ColAlignment(1) = 1
'        grdCambioPosicion.ColAlignment(2) = 1
'        grdCambioPosicion.ColAlignment(3) = 1
'        grdCambioPosicion.ColAlignment(4) = 1
'        grdCambioPosicion.ColAlignment(5) = 1
'        grdCambioPosicion.ColAlignment(6) = 1
'        grdCambioPosicion.ColAlignment(7) = 1
'        grdCambioPosicion.Rows = Ccajas.Count + 1
'
'        For i = 1 To grdCambioPosicion.Rows - 1
'            grdCambioPosicion.TextMatrix(i, 1) = Ccajas.Item(i).COD_CLIENTE
''            Set Rs = New ADODB.Recordset
''             Rs.Open "SELECT * FROM CLIENTES WHERE ID_CLIENTE = " & Ccajas.Item(i).Cod_Cliente, strConBasa , 0 ,1
''            If Not Rs.EOF Then
''                 grdCambioPosicion.TextMatrix(i, 2) = Trim(UCase(Rs!Razon_Social))
''            End If
'            grdCambioPosicion.TextMatrix(i, 3) = Ccajas.Item(i).NRO_CAJA
'            grdCambioPosicion.TextMatrix(i, 4) = Ccajas.Item(i).Estanteria
'            grdCambioPosicion.TextMatrix(i, 5) = Ccajas.Item(i).Horizontal
'            grdCambioPosicion.TextMatrix(i, 6) = Ccajas.Item(i).Vertical
'            If Ccajas.Item(i).Adelante_Atras = 1 Then
'                grdCambioPosicion.TextMatrix(i, 7) = "ATRAS"
'            Else
'                grdCambioPosicion.TextMatrix(i, 7) = "ADELANTE"
'            End If
'        Next
'        MousePointer = 0
'        cmdAceptar.Enabled = True
End Sub


Private Sub Verificar()
    If (Trim(txtEstanteria_Desde.Text) = "" And Trim(txtEstanteria_Hasta.Text) = "" _
    And Trim(txtHorizontal_Desde.Text) = "" And Trim(txtHorizontal_Hasta.Text) _
    And Trim(txtVertical_Desde.Text) = "" And Trim(txtVertical_Hasta.Text) = "") Then
        MsgBox "Los datos son incorrectos"
        Exit Sub
    End If
   
    Dim rsPosicionLibre As New ADODB.Recordset
    Dim SQL  As String
    Dim SqlContar As String
    Dim SqlG As String
       
       
                SQL = "Select * from contenedor "
                SqlContar = "Select count(*) as Cantidad from contenedor"
                SqlG = "  where estado = 1 and ( cod_cliente=0 or cod_cliente is null) "
                SqlG = SqlG & " AND (estanteria Between  " & txtEstanteria_Desde & "  AND " & txtEstanteria_Hasta
                SqlG = SqlG & ") AND ( HORIZONTAL Between " & txtHorizontal_Desde & " AND " & txtHorizontal_Hasta & ")"
                SqlG = SqlG & " AND (VERTICAL Between " & txtVertical_Desde & " AND " & txtVertical_Hasta & ") "
                If optAtras.value = True Then
                    SqlG = SqlG & " AND Adelante_Atras = 1 "
                End If
                If optFrente.value = True Then
                    SqlG = SqlG & " AND Adelante_Atras =2 "
                End If
                SqlContar = SqlContar & SqlG
                SQL = SQL & SqlG & " Order by estanteria,vertical,Horizontal"
                
                
                
       
        rsPosicionLibre.Open SqlContar, ConActiva, 0, 1
        
        
        If Not rsPosicionLibre.EOF Then
            If lblCantidadTotal.Caption > CLng(rsPosicionLibre!cantidad) Then
                MsgBox "No Alcanzan las posiciones" & vbCrLf & "Solo quedan para ese rango " & CStr(rsPosicionLibre!cantidad) & " posiciones", vbCritical
                MousePointer = 0
                Exit Sub
            End If
        Else
        
               MsgBox "No Alcanzan las posiciones" & vbCrLf & "Solo quedan para ese rango " & CStr(rsPosicionLibre!Contar) & " posiciones", vbCritical
                MousePointer = 0
                Exit Sub
        End If
        
        
        grdCambioPosicion.Cols = 9
        grdCambioPosicion.ColWidth(0) = 10
        grdCambioPosicion.ColWidth(1) = 800
        grdCambioPosicion.ColWidth(2) = 800
        grdCambioPosicion.ColWidth(3) = 800
        grdCambioPosicion.ColWidth(4) = 800
        grdCambioPosicion.ColWidth(5) = 800
        grdCambioPosicion.ColWidth(6) = 800
        grdCambioPosicion.ColWidth(7) = 800
        

        grdCambioPosicion.Font.Size = 9
        grdCambioPosicion.TextMatrix(0, 1) = "CAJA"
        grdCambioPosicion.TextMatrix(0, 2) = "ID_CLIENTE"
        grdCambioPosicion.TextMatrix(0, 3) = "ESTADO"
        grdCambioPosicion.TextMatrix(0, 4) = "EST."
        grdCambioPosicion.TextMatrix(0, 5) = "HOR."
        grdCambioPosicion.TextMatrix(0, 6) = "VER."
        grdCambioPosicion.TextMatrix(0, 7) = "AD/AT"
        grdCambioPosicion.TextMatrix(0, 8) = "ID_CONFINAL"
        
        

        grdCambioPosicion.ColAlignment(1) = 1
        grdCambioPosicion.ColAlignment(2) = 1
        grdCambioPosicion.ColAlignment(3) = 1
        grdCambioPosicion.ColAlignment(4) = 1
        grdCambioPosicion.ColAlignment(5) = 1
        grdCambioPosicion.ColAlignment(6) = 1
        grdCambioPosicion.ColAlignment(7) = 1
        grdCambioPosicion.ColAlignment(8) = 1
        
Dim R As Integer
      Set rsPosicionLibre = New ADODB.Recordset
      rsPosicionLibre.Open SQL, ConActiva, 0, 1
          Do While Not rsPosicionLibre.EOF
            For R = 1 To grdCambioPosicion.Rows
                If R >= grdCambioPosicion.Rows Then
                    Exit Do
                End If
                With grdCambioPosicion
                    .TextMatrix(R, 4) = rsPosicionLibre!Estanteria
                    .TextMatrix(R, 5) = rsPosicionLibre!Horizontal
                    .TextMatrix(R, 6) = rsPosicionLibre!Vertical
                    .TextMatrix(R, 7) = rsPosicionLibre!Adelante_Atras
                    .TextMatrix(R, 8) = rsPosicionLibre!ID_CONTENEDOR
                End With
                    rsPosicionLibre.MoveNext
            Next
        Loop
        cmdAceptar.Enabled = True
        
End Sub


Private Sub cmdVerPosiciones_Click()
Dim SQL As String
Dim RS As New ADODB.Recordset

DataGrid1.Visible = True
grdCambioPosicion.Visible = False
SQL = " SELECT     ID_CONTENEDOR,  ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, ESTADO, COD_CLIENTE, NRO_CAJA"
SQL = SQL & " From CONTENEDOR "
SQL = SQL & " WHERE  ESTANTERIA BETWEEN " & txtEstanteria_Desde.Text & " AND " & txtEstanteria_Hasta.Text
SQL = SQL & " AND HORIZONTAL BETWEEN  " & txtHorizontal_Desde.Text & " AND  " & txtHorizontal_Hasta.Text
SQL = SQL & " AND VERTICAL BETWEEN " & txtVertical_Desde.Text & "  AND " & txtVertical_Hasta.Text
SQL = SQL & " AND ESTADO IN(1,0,3) "
SQL = SQL & " ORDER BY ESTANTERIA, VERTICAL, ADELANTE_ATRAS, HORIZONTAL "

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open SQL, ConActiva, 2, 3


Set DataGrid1.DataSource = RS
DataGrid1.ReBind
DataGrid1.Refresh
CopiarDatosGrilla DataGrid1


End Sub

Private Sub Command1_Click()
    Dim RS As New ADODB.Recordset
    Dim Lectura As Long
    Dim SQL As String
    Dim Sql_General As String
    Dim ContarSql As String
    Dim Bloque As String
    Dim C As Integer
    
        grdCambioPosicion.Visible = True
        DataGrid1.Visible = False
        grdCambioPosicion.Clear
        TituloGrilla
    
        Lectura = InputBox("Por Favor Ingrese el numero de Lectura ", "Lectura", lblLectura.Caption)
        lblLectura = Lectura
        
        Dim Est As Integer
        Dim M_Vert As Integer
        Dim M_Hor As Integer
        
        
        SQL = " SELECT     NUMERO_LECTURA, CLIENTE, CAJA"
        SQL = SQL & " From LECTURACOLECTOR"
        SQL = SQL & "  Where (Cliente = 9999)"
        SQL = SQL & "  And NUMERO_LECTURA = " & Lectura

 Set RS = New ADODB.Recordset
 txtLectura.Text = ""
 
        RS.Open SQL, strConBasa, 0, 1
        
        If Not RS.EOF Then
            Bloque = RS!Caja
                    Est = Mid(RS!Caja, 1, 4)
                    M_Vert = Mid(RS!Caja, 5, 2)
                    M_Hor = Mid(RS!Caja, 7, 2)
                    SQL = " SELECT     ESTANTERIA, MAX(HORIZONTAL) AS MaxHorizontal, MIN(HORIZONTAL) AS MinHorizontal, MAX(VERTICAL) AS MaxVertical, MIN(VERTICAL) AS MinVertical"
                    SQL = SQL & " From CONTENEDOR"
                    SQL = SQL & " Where Modulo_V =  " & M_Vert
                    SQL = SQL & " And Modulo_H = " & M_Hor
                    SQL = SQL & " GROUP BY ESTANTERIA"
                    SQL = SQL & " Having Estanteria = " & Est
                    txtLectura.Text = RS!Caja
                Set RS = New ADODB.Recordset
                RS.Open SQL, strConBasa, 0, 1
                                
                If Not RS.EOF Then
                    txtEstanteria_Desde.Text = Est
                    txtEstanteria_Hasta.Text = Est
                    txtHorizontal_Desde.Text = RS!MinHorizontal
                    txtHorizontal_Hasta.Text = RS!MaxHorizontal
                    txtVertical_Desde.Text = RS!MinVertical
                    txtVertical_Hasta.Text = RS!MaxVertical
                Else
                    txtEstanteria_Desde.Text = Est
                    txtEstanteria_Hasta.Text = Est
                    txtHorizontal_Desde.Text = 0
                    txtHorizontal_Hasta.Text = 0
                    txtVertical_Desde.Text = 0
                    txtVertical_Hasta.Text = 0
                End If
             
        
        End If
        
       If chkLiberarModulo.value = False Then
                Dim SQLD As String
                ContarSql = " Select  count(*) AS CANTIDAD "
                SQL = " SELECT     CONTENEDOR.ID_CONTENEDOR, CONTENEDOR.NRO_CAJA, CONTENEDOR.COD_CLIENTE, CLIENTES.RAZON_SOCIAL"
                Sql_General = " FROM         LECTURACOLECTOR INNER JOIN"
                Sql_General = Sql_General & " CLIENTES ON LECTURACOLECTOR.CLIENTE = CLIENTES.ID_CLIENTE INNER JOIN"
                Sql_General = Sql_General & "  CONTENEDOR ON LECTURACOLECTOR.CLIENTE = CONTENEDOR.COD_CLIENTE AND LECTURACOLECTOR.CAJA = CONTENEDOR.NRO_CAJA"
                Sql_General = Sql_General & " Where   (NOT (CONTENEDOR.ESTADO IS NULL))   "
                Sql_General = Sql_General & " And LECTURACOLECTOR.NUMERO_LECTURA = " & Lectura
                ContarSql = ContarSql & Sql_General
                SQL = SQL & Sql_General & "  ORDER BY LECTURACOLECTOR.ORDEN  "
                SQLD = " SELECT     LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.ORDEN, CONTENEDOR.ESTADO, CONTENEDOR.ID_CONTENEDOR,"
                SQLD = SQLD & vbCrLf & "  CONTENEDOR.NRO_CAJA , CONTENEDOR.COD_CLIENTE, Clientes.RAZON_SOCIAL"
                SQLD = SQLD & vbCrLf & " FROM         LECTURACOLECTOR LEFT OUTER JOIN"
                SQLD = SQLD & vbCrLf & " CLIENTES ON LECTURACOLECTOR.CLIENTE = CLIENTES.ID_CLIENTE LEFT OUTER JOIN"
                SQLD = SQLD & vbCrLf & "  CONTENEDOR ON LECTURACOLECTOR.CLIENTE = CONTENEDOR.COD_CLIENTE AND LECTURACOLECTOR.CAJA = CONTENEDOR.NRO_CAJA"
                SQLD = SQLD & vbCrLf & " Where LECTURACOLECTOR.CLIENTE < 9000 AND LECTURACOLECTOR.NUMERO_LECTURA = " & Lectura
                SQLD = SQLD & vbCrLf & " ORDER BY LECTURACOLECTOR.ORDEN"
                Dim ERROR As String
                Set RS = New ADODB.Recordset
                RS.Open ContarSql, strConBasa, 0, 1
                    If Not RS.EOF Then
                        lblCantidadTotal.Caption = RS!cantidad
                    Else
                        lblCantidadTotal.Caption = "0"
                    Exit Sub
                    End If
                Set RS = New ADODB.Recordset
                RS.Open SQLD, strConBasa, 0, 1
    Else
        SQL = " SELECT     ORDEN, ID_CONTENEDOR, COD_CLIENTE as Cliente , NRO_CAJA as Caja, ESTADO, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS ,  CLIENTES.RAZON_SOCIAL "
        SQL = SQL & vbCrLf & " FROM CONTENEDOR INNER JOIN "
        SQL = SQL & vbCrLf & " CLIENTES ON CONTENEDOR.COD_CLIENTE = CLIENTES.ID_CLIENTE "
        SQL = SQL & vbCrLf & " WHERE     ESTANTERIA BETWEEN " & txtEstanteria_Desde.Text & " And " & txtEstanteria_Hasta.Text
        SQL = SQL & vbCrLf & " AND HORIZONTAL BETWEEN " & txtHorizontal_Desde.Text & " AND " & txtHorizontal_Hasta.Text
        SQL = SQL & vbCrLf & " AND VERTICAL BETWEEN " & txtVertical_Desde.Text & " AND " & txtVertical_Hasta.Text
       Rem  sql = sql & vbCrLf & " AND ADELANTE_ATRAS = 1"
        SQL = SQL & vbCrLf & " AND ESTADO <> 1 "
        Set RS = New ADODB.Recordset
            RS.Open SQL, ConActiva, 0, 1
            txtEstanteria_Desde.Text = 150
            txtEstanteria_Hasta.Text = 180
            txtHorizontal_Desde.Text = 1
            txtHorizontal_Hasta.Text = 50
            txtVertical_Desde.Text = 1
            txtVertical_Hasta.Text = 200
    End If
    
        
    Do While Not RS.EOF
      If chkLiberarModulo.value = 1 Then
            grdCambioPosicion.AddItem RS!ID_CONTENEDOR & vbTab & RS!Caja & vbTab & RS!Cliente & vbTab & RS!estado & vbTab & RS!RAZON_SOCIAL
      Else
            If RS!estado = 2 Then
                    C = C + 1
                    If RS!Cliente = 16 Then
                        MsgBox "Cliente 16  incorrecto consultar al Jefe de planta Caja : " & RS!Caja
                    Else
                        grdCambioPosicion.AddItem RS!ID_CONTENEDOR & vbTab & RS!Caja & vbTab & RS!Cliente & vbTab & RS!estado & vbTab & RS!RAZON_SOCIAL
                    End If
            Else
                    If RS!estado = 3 Then
                        If chkLiberarModulo.value = True Then
                            SQL = " Update CONTENEDOR Set estado = 2 "
                            SQL = SQL & " Where (estado = 3 ) "
                            SQL = SQL & " And COD_CLIENTE = " & RS!COD_CLIENTE
                            SQL = SQL & " And NRO_CAJA = " & RS!NRO_CAJA
                            ExecutarSql SQL
                            C = C + 1
                            grdCambioPosicion.AddItem RS!ID_CONTENEDOR & vbTab & RS!Caja & vbTab & RS!Cliente & vbTab & RS!estado & vbTab & RS!RAZON_SOCIAL
                        Else
                            ERROR = ERROR & vbCrLf & RS!Cliente & vbTab & RS!Caja & vbTab & RS!Orden
                        End If
                    Else
                        ERROR = ERROR & vbCrLf & RS!Cliente & vbTab & RS!Caja & vbTab & RS!Orden
                    End If
             End If
        End If
            RS.MoveNext
        Loop
        Dim Titul As String
         lblCantidadTotal.Caption = C
        If ERROR <> "" Then
            MsgBox "EXISTEN ERRORES SERAN COPIADOS A MEMORIA "
            Clipboard.Clear
            Clipboard.SetText ERROR
            grdCambioPosicion.Clear
            grdCambioPosicion.Rows = 1
            If Trim(Bloque) <> "" Then
                Titul = "Veridicar el:  " & Bloque
            End If
            ERROR = Titul & vbCrLf & "CLIENTE" & vbTab & "CAJA" & vbTab & "Orden" & vbTab & "LECTURA ERROR:" & Lectura & vbCrLf & ERROR
            Clipboard.SetText ERROR
'                grdCambioPosicion.Clear
'                grdCambioPosicion.Rows = 1
        
        End If
        
            
End Sub


Private Sub Command2_Click()
 Dim SQL As String
        
        SQL = " SELECT  *"
        SQL = SQL & vbCrLf & " From V_LECTURAROTULO"
        SQL = SQL & vbCrLf & " Where NUMERO_LECTURA = " & lblLectura.Caption
        SQL = SQL & vbCrLf & " ORDER BY ORDEN "
        frmReportes.ImprimirReporte PasoReportes & "rptRotuloEtiquetaBarra.rpt", SQL, True
  
End Sub

Private Sub Command3_Click()
If Trim(lblLectura.Caption) <> "" Then
     ImprimirRotulosLectura lblLectura.Caption, False
End If
End Sub

Private Sub Command4_Click()
Dim SQL As String
Dim RS As New ADODB.Recordset

DataGrid1.Visible = True
grdCambioPosicion.Visible = False
SQL = " SELECT     ESTANTERIA, HORIZONTAL, VERTICAL,  ESTADO, COD_CLIENTE, NRO_CAJA , MODULO_V , MODULO_H  "
SQL = SQL & " From CONTENEDOR "
SQL = SQL & " WHERE  ESTANTERIA = " & txtEstanteria_Desde.Text
SQL = SQL & " AND HORIZONTAL BETWEEN  " & txtHorizontal_Desde.Text & " AND  " & txtHorizontal_Hasta.Text
SQL = SQL & " AND VERTICAL BETWEEN " & txtVertical_Desde.Text & "  AND " & txtVertical_Hasta.Text
SQL = SQL & " ORDER BY ESTANTERIA, VERTICAL,  HORIZONTAL "

    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open SQL, ConActiva, 2, 3


Set DataGrid1.DataSource = RS
DataGrid1.ReBind
DataGrid1.Refresh
CopiarDatosGrilla DataGrid1
End Sub

Private Sub Form_Load()
    Rem IdEntrega = 0
    Rem idRecibe = 0
    Rem IdClienteAnterior = 0
    lblFecha = Format(SysDate2, "dd/mm/yyyy")
    TituloGrilla
    Rem ctlResponsable.TipoControl = PERSONAL
    cmdAceptar.Enabled = False
If MDIfrmInicio.StaInicio.Panels(2).Text = 19 Or MDIfrmInicio.StaInicio.Panels(2).Text = 17 Or MDIfrmInicio.StaInicio.Panels(2).Text = 61 Then
    Bloque = True
    cmdCambioModulo.Enabled = True
Else
    Bloque = False
End If

End Sub

Function ProximoRemito() As Long
  Dim SQL As String
  Dim OraMax As ADODB.Recordset
  SQL = "Select Max(Nro_Remito) Maximo From Remitos_Cuerpo"
  Set OraMax = New ADODB.Recordset
  OraMax.Open ConBasa, SQL
  If IsNull(OraMax("Maximo")) Then ProximoRemito = 1: Exit Function
  ProximoRemito = Val(OraMax("Maximo")) + 1
End Function
Public Sub Guardar_Remito()
'    Dim Sql As String
'    Dim r As Integer
'    Dim C As Integer
'    Dim oradyn As New ADODB.Recordset
'    Dim Proximo_Nro_Remito As Long
'    On Error GoTo OraError
'        If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
'                Screen.MousePointer = 11
'               Rem  OraSession.BeginTrans
'                Proximo_Nro_Remito = ProximoRemito
'
'                ' INSERTAR EN REMITO CUERPO
'                Sql = "Insert into Remitos_Cuerpo (Nro_Remito,NRO_REM_PROV, Tipo, Operacion,"
'                Sql = Sql & vbCrLf & " Estado, Fecha, Id_Cliente, Observaciones, Cantidad, "
'                Sql = Sql & vbCrLf & " Audit_Usuario, Audit_Fecha, Fecha_Ingreso,Fecha_Error)"
'                Sql = Sql & vbCrLf & " Values (" & CLng(Proximo_Nro_Remito) & ",'" & lblNumeroRemito & "',"       ' Nro Remito
'                Sql = Sql & 4 & ","                  ' Tipo
'                Sql = Sql & 0 & ","           ' Operacion
'                Sql = Sql & vbCrLf & "  0," ' ESTADO
'                Sql = Sql & SysDate & ","           ' Fecha
'                Sql = Sql & CInt(lblIDCliente.Caption) & ","                 ' Id Cliente
'                Sql = Sql & " '" & "" & "',"                             ' Observaciones
'                Sql = Sql & lblCantidad.Caption & ","                    ' Cantidad
'                Sql = Sql & vbCrLf & "  '" & UCase(UserName$) & "',"                          ' Usuario
'                Sql = Sql & SysDate & ","  ' Fecha y Hora
'                Sql = Sql & SysDate & ","
'                Sql = Sql & 0 & ")"
'                Debug.Print Sql
'                OraDatabase.ExecuteSQL Sql
'                'INSERTAR EN REMITO DELTALLE
'                For r = 1 To grdGuardiayCustodia.Rows - 1
'                   For C = 1 To grdGuardiayCustodia.Cols - 1
'                        If grdGuardiayCustodia.TextMatrix(r, C) <> "" Then
'                            Sql = "Insert into Remitos_Detalle(Nro_Remito, Desde, Hasta,"
'                            Sql = Sql & vbCrLf & " Tipo_Almacenado, Detalle, Audit_Usuario, Audit_Fecha)"
'                            Sql = Sql & vbCrLf & " Values (" + Format(Proximo_Nro_Remito) + ","
'                            Sql = Sql & grdGuardiayCustodia.TextMatrix(r, C) & ","
'                            Sql = Sql & grdGuardiayCustodia.TextMatrix(r, C) & ","
'                            Sql = Sql & vbCrLf & 0 & ","
'                            Sql = Sql & "'',"
'                            Sql = Sql & " '" + UCase(UserName$) + "', "                                ' Usuario
'                            Sql = Sql & SysDate & " )"
'                            'MsgBox Sql
'                            Debug.Print Sql
'                            OraDatabase.ExecuteSQL Sql
'                            GrabarMovHistorico Proximo_Nro_Remito, grdGuardiayCustodia.TextMatrix(r, C), grdGuardiayCustodia.TextMatrix(r, C), lblIDCliente, 0, 4, 0, SysDate
'                        End If
'                     Next
'                 Next
'                'INSERTAR EN MOVIMIENTOS
'                Sql = "Insert into Movimientos(Id_cliente, Fecha, Nro_Remito,"
'                Sql = Sql & vbCrLf & " Tipo_Movim, Oper_Movim, Cantidad, Audit_Usuario,"
'                Sql = Sql & vbCrLf & " Audit_Fecha)"
'                Sql = Sql & vbCrLf & " Values ( " & CInt(lblIDCliente.Caption) & "," ' Id Cliente
'                Sql = Sql & SysDate & "," ' Fecha
'                Sql = Sql & CLng(Proximo_Nro_Remito) & ","  ' nro remito
'                Sql = Sql & vbCrLf & "        " & 4 & "," ' Tipo
'                Sql = Sql & 0 & ","  ' Operacion
'                Sql = Sql & CLng(lblCantidad.Caption) & ","   ' Cantidad
'                Sql = Sql & " '" + UCase(UserName$) + "',"                           ' Usuario
'                Sql = Sql & vbCrLf & "        " & SysDate & ")"  ' Fecha de cargar
'                Debug.Print Sql
'                OraDatabase.ExecuteSQL Sql
'                'MOVIMIENTO EN TABLA CONTENEDO
'                            For r = 1 To grdGuardiayCustodia.Rows - 1
'                               For C = 1 To grdGuardiayCustodia.Cols - 1
'                                   If grdGuardiayCustodia.TextMatrix(r, C) <> "" Then
'                                        Sql = "UPDATE CONTENEDOR SET "
'                                        Sql = Sql & vbCrLf & " ESTADO = 1 "
'                                        Sql = Sql & vbCrLf & " , COD_CLIENTE = 0 "
'                                        Sql = Sql & vbCrLf & " , NRO_CAJA = '' "
'                                        Sql = Sql & ", NRO_REMITO = " & Proximo_Nro_Remito
'                                        Sql = Sql & ", F_MODIFICACION = " & SysDate
'                                        Sql = Sql & vbCrLf & " WHERE "
'                                        Sql = Sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption)
'                                        Sql = Sql & " AND NRO_CAJA = " & CLng(grdGuardiayCustodia.TextMatrix(r, C))
'                                        Sql = Sql & " AND ESTADO = 5 "
'                                        Debug.Print Sql
'                                        OraDatabase.ExecuteSQL Sql
'                                   End If
'                               Next
'                            Next
'                OraSession.CommitTrans
'                MsgBox "El remito fue grabado con exito", vbExclamation, "Remito"
'                MsgBox "NUMERO DE MOVIMIENTO ES " & Proximo_Nro_Remito
'                Screen.MousePointer = 0
'                On Error GoTo ErrorPrn
'                Unload Me
'        End If
'    Exit Sub
'OraError:
'        Screen.MousePointer = 0
'        OraSession.Rollback
'        frmLogOraError.Show MODAL
'        Exit Sub
'
'ErrorPrn:
'        MsgBox Error
'        Exit Sub
'
End Sub
Sub GrabarMovHistorico(mov_nrorem, mov_desde, mov_hasta, _
mov_cliente, mov_elem, mov_tipo, mov_oper, mov_fecha)
Dim R As Single
Dim SQL As String
Dim oradyn As ADODB.Recordset
    
    SQL = "Select * from Mov_Cajas"
    Set oradyn = New ADODB.Recordset
    oradyn.Open SQL, ConActiva, 0, 1
    
    For R = mov_desde To mov_hasta
        oradyn.AddNew
        oradyn!NRO_REMITO = mov_nrorem
        oradyn!NRO_CAJA = R
        oradyn!id_cliente = mov_cliente
        oradyn!Elemento = mov_elem
        oradyn!TIPO = mov_tipo
        oradyn!Operacion = mov_oper
        oradyn!FECHA_MOVIMIENTO = SysDate
        oradyn!ANULADO = 0
        oradyn!AUDIT_USUARIO = UserName
        oradyn!AUDIT_FECHA = SysDate2
        oradyn.Update
    Next
    
End Sub
Public Sub Hablar(Data As String)
'    MMControl5.Command = "close"
'    MMControl5.DeviceType = "WaveAudio"
'    MMControl5.FileName = "D:\numeros\" & Data & ".wav"
'    MMControl5.Command = "open"
'    MMControl5.Command = "Prev"
'    MMControl5.Command = "Play"
End Sub

Public Sub TituloGrilla()
    grdCambioPosicion.ColWidth(0) = 100
    grdCambioPosicion.ColWidth(1) = 1000
    grdCambioPosicion.ColWidth(2) = 1000
    grdCambioPosicion.ColWidth(3) = 1000
    grdCambioPosicion.ColWidth(4) = 6000
    
    grdCambioPosicion.ColAlignment(1) = 4
    grdCambioPosicion.ColAlignment(2) = 4
    grdCambioPosicion.ColAlignment(3) = 2
    grdCambioPosicion.ColAlignment(4) = 2
    
    grdCambioPosicion.TextMatrix(0, 1) = "Caja"
    grdCambioPosicion.TextMatrix(0, 2) = "ID"
    grdCambioPosicion.TextMatrix(0, 3) = "ESTADO"
    grdCambioPosicion.TextMatrix(0, 4) = "RAZON SOCIAL"
    
    grdCambioPosicion.Rows = 1
    grdCambioPosicion.Cols = 5
  
    
    
    
End Sub
Public Sub CargarGrilla(Valor As String)
'    Dim C As Integer
'    Dim r As Integer
'    Dim RsEstadoCaja As ADODB.Recordset
'    Set RsEstadoCaja = New ADODB.Recordset
'    RsEstadoCaja.Open "Select * from contenedor where cod_cliente= " & CInt(lblIDCliente) & " and nro_caja = " & Valor, ConActiva, 0, 1
'    If Not RsEstadoCaja.EOF Then
'      Select Case CInt(RsEstadoCaja!estado)
'      Case 2
'
'            For r = 1 To grdGuardiayCustodia.Rows - 1
'                For C = 1 To grdGuardiayCustodia.Cols - 1
'                    If grdGuardiayCustodia.TextMatrix(r, C) = Valor Then
'                        Hablar "REPETIDA"
'                        Exit Sub
'                    End If
'                    If grdGuardiayCustodia.TextMatrix(r, C) = "" Then
'                        grdGuardiayCustodia.TextMatrix(r, C) = Valor
'                        Hablar "ENTRADA"
'                        ContarGrilla grdCambioPosicion, lblCantidadTotal
'                        Exit Sub
'                    End If
'                Next
'            Next
'            grdGuardiayCustodia.AddItem ""
'            grdGuardiayCustodia.TextMatrix(grdGuardiayCustodia.Rows - 1, 1) = Valor
'            Hablar "ENTRADA"
'            ContarGrilla grdCambioPosicion, lblCantidadTotal
'
'    Case Else
'        Hablar "ESTADO"
'    End Select
'   End If
End Sub
Public Sub ImprimirRemito(NumeroRemito As Long)
'    Dim Sql As String
'    Dim sql1 As String
'    Dim BANDERA As Boolean
'    Dim Responsables As String
'    Dim rs As ADODB.Recordset
'    Dim ANTERIOR As Long
'    If NumeroRemito = 0 Or IsNull(NumeroRemito) Then
'        Exit Sub
'    End If
'
'   On Error GoTo Err
'    Sql = "    SELECT"
'    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.OBSERVACIONES,"
'    Sql = Sql & vbCrLf & "    REMITOS_DETALLE.DESDE,"
'    Sql = Sql & vbCrLf & "    REQUERIMIENTO.SECTOR, REQUERIMIENTO.SOLICITANTE, REQUERIMIENTO.FECHARECEPCION,"
'    Sql = Sql & vbCrLf & "    REMITO_TIPO.DESCRIPCION,"
'    Sql = Sql & vbCrLf & "    REMITO_OPERACION.DESCRIPCION,"
'    Sql = Sql & vbCrLf & "    REMITO_ESTADOS.DESCRIPCION,"
'    Sql = Sql & vbCrLf & "    clientes.id_cliente , clientes.RAZON_SOCIAL, clientes.CALLE, clientes.NUMERO, clientes.LOCALIDAD"
'    Sql = Sql & vbCrLf & "From"
'    Sql = Sql & vbCrLf & "    BASA.REMITOS_CUERPO REMITOS_CUERPO,"
'    Sql = Sql & vbCrLf & "    BASA.REMITOS_DETALLE REMITOS_DETALLE,"
'    Sql = Sql & vbCrLf & "    BASA.REQUERIMIENTO REQUERIMIENTO,"
'    Sql = Sql & vbCrLf & "    BASA.REMITO_TIPO REMITO_TIPO,"
'    Sql = Sql & vbCrLf & "   BASA.REMITO_OPERACION REMITO_OPERACION,"
'    Sql = Sql & vbCrLf & "   BASA.REMITO_ESTADOS REMITO_ESTADOS,"
'    Sql = Sql & vbCrLf & "   BASA.clientes clientes"
'    Sql = Sql & vbCrLf & " Where"
'    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO AND"
'    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.NRO_REMITO = REQUERIMIENTO.IDREMITO AND"
'    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.TIPO = REMITO_TIPO.ID AND"
'    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID AND"
'    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
'    Sql = Sql & vbCrLf & "    REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
'    Sql = Sql & vbCrLf & "    REMITOS_CUERPO.NRO_REMITO =" & NumeroRemito
'
'            sql1 = " SELECT"
'            sql1 = sql1 & vbCrLf & "      H_ESTADO_REQUE.IDREQUERIMIENTO,"
'            sql1 = sql1 & vbCrLf & "    H_ESTADO_REQUE.IDESTADO,"
'            sql1 = sql1 & vbCrLf & "     H_ESTADO_REQUE.CONTADOR,"
'            sql1 = sql1 & vbCrLf & "     PERSONAL.NOMBRE,PERSONAL.APELLIDO"
'            sql1 = sql1 & vbCrLf & "  From"
'            sql1 = sql1 & vbCrLf & "     BASA.H_ESTADO_REQUE , PERSONAL, Requerimiento"
'            sql1 = sql1 & vbCrLf & "  Where"
'            sql1 = sql1 & vbCrLf & "     H_ESTADO_REQUE.idPersonal = PERSONAL.idPersonal"
'            sql1 = sql1 & vbCrLf & "     AND H_ESTADO_REQUE.IDESTADO = REQUERIMIENTO.IDESTADO"
'            sql1 = sql1 & vbCrLf & "     AND H_ESTADO_REQUE.IDREQUERIMIENTO = REQUERIMIENTO.IDREQUERIMIENTO"
'            sql1 = sql1 & vbCrLf & "    AND H_ESTADO_REQUE.IDREQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
'            sql1 = sql1 & vbCrLf & " Order By"
'            sql1 = sql1 & vbCrLf & "     H_ESTADO_REQUE.IDREQUERIMIENTO Asc"
'    Set rs = New ADODB.Recordset
'    rs.Open sql1, ConActiva, 0, 1
'    Do While Not rs.EOF
'        If CLng(rs!IDREQUERIMIENTO) = ANTERIOR Then
'            Responsables = Responsables & " , " & CStr(rs!Nombre) & " " & CStr(rs!Apellido)
'        Else
'           If BANDERA = False Then
'                ANTERIOR = rs!IDREQUERIMIENTO
'                BANDERA = True
'                Responsables = CStr(rs!Nombre) & " " & CStr(rs!Apellido)
'           Else
'                Exit Do
'           End If
'        End If
'        rs.MoveNext
'    Loop
'    DoEvents
'    CryRemito.Connect = "DSN = bpdc;UID = " & UserName & ";PWD = " & Password
'    CryRemito.ReportFileName = "\\Server1basa\Sistemas\Requerimientos\remito.rpt"
'    CryRemito.DiscardSavedData = True
'    CryRemito.Formulas(0) = "f ='" & " : " & Responsables & "'"
'    CryRemito.Formulas(1) = "COPIA ='" & "ORIGINAL" & " '"
'    CryRemito.SQLQuery = Sql
'    CryRemito.Destination = 1
'    CryRemito.Action = 1
'    CryRemito.DiscardSavedData = True
'    CryRemito.Formulas(0) = "f ='" & " : " & Responsables & "'"
'    CryRemito.Formulas(1) = "COPIA ='" & "DUPLICADO" & " '"
'    CryRemito.SQLQuery = Sql
'    CryRemito.Destination = 1
'    CryRemito.Action = 1
'    Exit Sub
'Err:
'    MsgBox "Atencion error al imprimir el remito " & vbCrLf & "Por favor intentolo desde la aplicacion de control de estados", vbInformation, "Error de Impresion"
End Sub

Public Function Validar() As Boolean
    Validar = True

'    If IsNull(ctlResponsable.Valor) Then
'        MsgBox "Falta el responsable"
'        Validar = False
'         Exit Function
'    End If
    If grdCambioPosicion.Rows < 1 Then
        MsgBox "no tiene caja"
        Validar = False
         Exit Function
    End If
'    If grdCambioPosicion.TextMatrix(0, 3) <> "CAJA" Then
'        MsgBox "usted debe derificar las posiocnes"
'        Validar = False
'         Exit Function
'    End If
End Function

Private Sub lblDescripcionRequerimiento_DblClick()
   Rem  txtObservaciones.Text = Trim(UCase(lblDescripcionRequerimiento.Caption))
End Sub

Private Sub lblCantidadDevolucion_Change()
'    If lblCantidadGuardia <> "" Then
'       lblCantidadTotal = lblCantidadGuardia
'    End If
'    If lblCantidadDevolucion <> "" Then
'          lblCantidadTotal = lblCantidadDevolucion
'    End If
'    If lblCantidadGuardia <> "" Then
'        If lblCantidadDevolucion <> "" Then
'            lblCantidadTotal = CInt(lblCantidadGuardia) + CInt(lblCantidadDevolucion)
'        End If
'    End If
End Sub
Private Sub lblCantidadGuardia_Change()
'    If lblCantidadGuardia <> "" Then
'       lblCantidadTotal = lblCantidadGuardia
'    End If
'    If lblCantidadDevolucion <> "" Then
'          lblCantidadTotal = lblCantidadDevolucion
'    End If
'    If lblCantidadGuardia <> "" Then
'        If lblCantidadDevolucion <> "" Then
'            lblCantidadTotal = CInt(lblCantidadGuardia) + CInt(lblCantidadDevolucion)
'        End If
'    End If
End Sub

Private Sub MMControl1_PlayClick(Cancel As Integer)
    Dim j
    j = 0
End Sub

Private Sub txtCaja_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Dim Cliente As Integer
'        Dim Caja As Long
'          If txtCaja <> "" And IsNumeric(txtCaja) Then
'               Caja = CLng(txtCaja)
'               CargarGrilla CLng(Caja)
'          End If
'          txtCaja = ""
'          If chkManual = 1 Then
'            txtTomarLectura.SetFocus
'          End If
'    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtTomarLectura_KeyPress(KeyAscii As Integer)
'        If KeyAscii = 13 Then
'            Dim Cliente As Integer
'            Dim Caja As Long
'            Dim rs As ADODB.Recordset
'            Select Case UCase(Mid(txtTomarLectura.Text, 1, 3))
'            Case "CAN"
'                If lblCantidadTotal <> "" Then
'                    Hablar "C" & lblCantidadTotal.Caption
'                Else
'                    Hablar "C" & "0"
'                End If
'            Case "R02"
'                    Dim rsRemitosFisico As ADODB.Recordset
'                    Dim Sql As String
'                    lblSucursalRemito = CLng(Mid(txtTomarLectura, 4, 4))
'                    lblNumeroRemito = CLng(Mid(txtTomarLectura, 8))
'                    Sql = "select * from REMITOS_FISICOS where"
'                    Sql = Sql & " SUCURSAL  = " & lblSucursalRemito
'                    Sql = Sql & " and NRO_REMITO_FISICO = " & lblNumeroRemito
'                    Sql = Sql & " and estado= 1"
'                    Set rsRemitosFisico = New ADODB.Recordset
'                    rsRemitosFisico.Open Sql, ConActiva, 0, 1
'                    If Not rsRemitosFisico.EOF Then
'                        lblSucursalRemito = ""
'                        lblNumeroRemito = ""
'                        Hablar "YAPROCE"
'                    End If
'            Case "P01"
'                Dim rsPersonal As ADODB.Recordset
'                    lblIDPersonal = CInt(Mid(txtTomarLectura, 4))
'                    Set rsPersonal = NEWADODB.Recordset
'                   rsPersonal.Open "Select * from Personal where idpersonal =" & CInt(lblIDPersonal), ConActiva, 0, 1
'                    If Not rsPersonal.EOF Then
'                        lblEntregaNombre = UCase(CStr(rsPersonal!Apellido) & "  " & CStr(rsPersonal!Nombre))
'                    End If
'            Case "C10"
'                If txtTomarLectura <> "" Then
'                    Caja = Mid(txtTomarLectura.Text, 6)
'                    Cliente = Mid(txtTomarLectura.Text, 3, 3)
'
'                    Set rs = OraDatabase.CreateDynaset("SELECT * FROM CLIENTES WHERE ID_CLIENTE = " & Cliente, ORADYN_READONLY)
'                    If Not rs.EOF Then
'                        lblIDCliente = rs!id_cliente
'                        lblCliente = Trim(UCase(rs!RAZON_SOCIAL))
'                    End If
'                    If Not ControlDuplicado(grdCambioPosicion, Caja) Then
'                        grdCambioPosicion.AddItem "" & vbTab & Caja & vbTab & Cliente & vbTab & lblCliente
'                        ContarGrilla grdCambioPosicion, lblCantidadTotal
'                        Hablar "ENTRADA"
'                    Else
'                        Hablar "REPETIDA"
'                    End If
'                Else
'                    Hablar "CLIENTE"
'                End If
'            Case Else
'
'                If txtTomarLectura <> "" And Len(txtTomarLectura) > 16 Then
'                    Caja = Mid(txtTomarLectura.Text, Len(txtTomarLectura.Text) - 5)
'                    Cliente = Mid(txtTomarLectura.Text, Len(txtTomarLectura.Text) - 8, 3)
'
'                    Set rs = OraDatabase.CreateDynaset("SELECT * FROM CLIENTES WHERE ID_CLIENTE = " & Cliente, ORADYN_READONLY)
'                    If Not rs.EOF Then
'                        lblIDCliente = rs!id_cliente
'                        lblCliente = Trim(UCase(rs!RAZON_SOCIAL))
'                    End If
'                    If Not ControlDuplicado(grdCambioPosicion, Caja) Then
'                        grdCambioPosicion.AddItem "" & vbTab & Caja & vbTab & Cliente & vbTab & lblCliente
'                        ContarGrilla grdCambioPosicion, lblCantidadTotal
'                        Hablar "ENTRADA"
'                    Else
'                        Hablar "REPETIDA"
'                    End If
'                Else
'                    Hablar "CLIENTE"
'                End If
'            End Select
'            txtTomarLectura = ""
'            txtTomarLectura.SetFocus
'        End If
End Sub

Public Function ContarGrilla(Grilla As MSFlexGrid, lblCantidad As Label) As Integer
    Dim I As Integer
    Dim R As Integer
    Dim C As Integer
        With Grilla
            For R = 1 To .Rows - 1
                I = I + 1
            Next
        End With
        lblCantidad = I
End Function

Public Sub limpiar()
'        TituloGrilla
'       Rem  lblIDCliente = ""
'        Rem lblCantidadDevolucion = ""
'        Rem lblCantidadGuardia = ""
'        lblCliente = ""
'        lblCantidadTotal = ""
'        lblEntregaNombre = ""
'        lblIDPersonal = ""
'        lblIDPersonalRecibe = ""
'        lblNumeroRemito = ""
'        lblRecibeNombre = ""
'        txtDesde = ""
'        txtHasta = ""
End Sub

Public Sub MovimientoCajas()
    
    
'    Dim Sql As String
'    Dim Sql2 As String
'    Dim Contar As Integer
'    Dim AntEstanteria, AntHorizontal, AntVertical, _
'    AntAdelante_Atras, AntNro_Estante, AntEstado, AntF_Modificacion, _
'    AntNro_Remito, AntIdRequerimiento As String
'    Dim NueEstanteria, NueHorizontal, NueVertical, NueAdelante_Atras As String
'    Dim NRO_CAJA As Long
'    Dim COD_CLIENTE As Integer
'    Dim ConCambio As New ADODB.Connection
'    ConCambio.Open strConBasa
'    Dim rsContenedor As ADODB.Recordset
'
'    Contar = 1
'
'
'    ConCambio.BeginTrans
'     On Error GoTo ErrorG
'    MsgBox "Inicio de cambio de posiciones", vbInformation
'     pbsCambios.Max = grdCambioPosicion.Rows - 1
'     pbsCambios.value = i
'    For i = 1 To grdCambioPosicion.Rows - 1
'        pbsCambios.value = i
'        pbsCambios.Refresh
'        frmCambioPosicionFisica.Refresh
'        COD_CLIENTE = grdCambioPosicion.TextMatrix(i, 1)
'        NRO_CAJA = grdCambioPosicion.TextMatrix(i, 3)
'        Debug.Print NRO_CAJA
'        Sql = "Select * from contenedor where cod_cliente = " & COD_CLIENTE
'        Sql = Sql & " AND NRO_CAJA =" & NRO_CAJA
'        Set rsContenedor = New ADODB.Recordset
'        rsContenedor.Open Sql, strConBasa, adOpenStatic, adLockReadOnly
'         If Not rsContenedor.EOF Then
'            With rsContenedor
'
'                AntEstanteria = CInt(!Estanteria)
'                NueEstanteria = grdCambioPosicion.TextMatrix(i, 4)
'                AntHorizontal = CInt(!Horizontal)
'                NueHorizontal = grdCambioPosicion.TextMatrix(i, 5)
'                AntVertical = CInt(!Vertical)
'                NueVertical = grdCambioPosicion.TextMatrix(i, 6)
'                AntAdelante_Atras = CInt(!Adelante_Atras)
'                If grdCambioPosicion.TextMatrix(i, 7) = "ATRAS" Then
'                    NueAdelante_Atras = 1
'                Else
'                    NueAdelante_Atras = 2
'                End If
'                If Not IsNull(!NRO_ESTANTE) Then
'                      AntNro_Estante = CInt(!NRO_ESTANTE)
'                Else
'                    AntNro_Estante = 0
'                End If
'                AntEstado = CInt(!estado)
'                If IsNull(!F_MODIFICACION) Then
'                    AntF_Modificacion = "NULL"
'                Else
'                    AntF_Modificacion = "'" & Format(CDate(!F_MODIFICACION), "dd/mm/yyyy") & "'"
'                End If
'                If IsNull(!NRO_REMITO) Then
'                    AntNro_Remito = "Null"
'                Else
'                    AntNro_Remito = CLng(!NRO_REMITO)
'                End If
'                If IsNull(!IDREQUERIMIENTO) Then
'                    AntIdRequerimiento = "Null"
'                Else
'                    AntIdRequerimiento = CLng(!IDREQUERIMIENTO)
'                End If
'                ID_Personal = ctlResponsable.Valor
'            End With
'
'            lblContar = Contar + 1
'            lblContar.Refresh
'            Contar = Contar + 1
'            Sql = "INSERT INTO CAMBIOPOSICION (ESTANTERIA, HORIZONTAL, VERTICAL,"
'            Sql = Sql & vbCrLf & " ADELANTE_ATRAS, NRO_ESTANTE, ESTADO,"
'            Sql = Sql & vbCrLf & " COD_CLIENTE, NRO_CAJA, FECHA,"
'            Sql = Sql & vbCrLf & " ID_PERSONAL )"
'            Sql = Sql & vbCrLf & " VALUES (" & AntEstanteria & "," & AntHorizontal & "," & AntVertical & ","
'            Sql = Sql & vbCrLf & AntAdelante_Atras & "," & AntNro_Estante & " ," & AntEstado & ","
'            Sql = Sql & vbCrLf & COD_CLIENTE & "," & NRO_CAJA & "," & SysDate & ","
'            Sql = Sql & vbCrLf & ID_Personal & " )"
'            Rem Debug.Print Sql
'             ConCambio.Execute Sql
'
'            Sql = " UPDATE CONTENEDOR SET ESTADO=" & 1 & ","
'            Rem Sql = " UPDATE CONTENEDOR SET ESTADO=" & 0 & ","
'            Sql = Sql & vbCrLf & " COD_CLIENTE=NULL, NRO_CAJA=NULL,"
'            Sql = Sql & vbCrLf & " NRO_REMITO=NULL, F_MODIFICACION=NULL,"
'            Sql = Sql & vbCrLf & " IDREQUERIMIENTO=NULL ,UB_PROVISORIA=NULL"
'            Sql = Sql & vbCrLf & " Where COD_CLIENTE = " & COD_CLIENTE
'            Sql = Sql & vbCrLf & " And NRO_CAJA = " & NRO_CAJA
'            Rem Debug.Print Sql
'            ConCambio.Execute Sql
'
'            Sql = " UPDATE CONTENEDOR SET ESTADO=" & AntEstado & ","
'            Sql = Sql & vbCrLf & " COD_CLIENTE=" & COD_CLIENTE & ", NRO_CAJA=" & NRO_CAJA
'            Sql = Sql & vbCrLf & ", NRO_REMITO = " & AntNro_Remito & ", F_MODIFICACION = " & SysDate & " , IDREQUERIMIENTO = " & AntIdRequerimiento
'            Sql = Sql & vbCrLf & " Where Estanteria = " & NueEstanteria & " And Horizontal = " & NueHorizontal & ""
'            Sql = Sql & vbCrLf & " And Vertical = " & NueVertical & " And ADELANTE_ATRAS = " & NueAdelante_Atras
'            Rem Debug.Print Sql
'            ConCambio.Execute Sql
'            If i = 1 Then
'                Sql2 = Sql2 & vbCrLf & "WHERE ( COD_CLIENTE = " & COD_CLIENTE & " AND NRO_CAJA = " & NRO_CAJA & ")"
'            Else
'                Sql2 = Sql2 & vbCrLf & " OR ( COD_CLIENTE = " & COD_CLIENTE & " AND NRO_CAJA = " & NRO_CAJA & ")"
'            End If
'        End If
'
'     Next
'         ConCambio.CommitTrans
'        Sql2 = Sql2 & vbCrLf & " Order by Estanteria , Vertical, Horizontal ,Adelante_Atras"
'        MsgBox "Cambio de posiciones finalizado", vbInformation
'    Exit Sub
'ErrorG:
'        ConCambio.Rollback
'        MsgBox "Los cambios No fueron realizados"
End Sub

Public Sub ImprimirRotulosLectura(Lectura As Long, Grande As Boolean)
   Dim SQL As String
        
        SQL = " SELECT  *"
        SQL = SQL & vbCrLf & " From V_LECTURAROTULO "
        SQL = SQL & vbCrLf & " Where NUMERO_LECTURA = " & Lectura
        SQL = SQL & vbCrLf & " ORDER BY ORDEN "
       
            frmReportes.ImprimirReporte PasoReportes & "rptCambioPosicionEtiqueta.rpt", SQL, True
       
End Sub

Public Sub ImprimirRotulosLecturaBarra(Lectura As Integer, Grande As Boolean)
   Dim SQL As String
        
        SQL = " SELECT  *"
        SQL = SQL & vbCrLf & " From V_LECTURAROTULO "
        SQL = SQL & vbCrLf & " Where NUMERO_LECTURA = " & Lectura
        SQL = SQL & vbCrLf & " ORDER BY ORDEN "
        
      
            frmReportes.ImprimirReporte PasoReportes & "rptRotuloEtiquetaBarra.rpt", SQL, True
        

End Sub



Public Function ControlDuplicado(Grilla As MSFlexGrid, Valor) As Boolean
    Dim I As Integer
    ControlDuplicado = False
    For I = 1 To Grilla.Rows - 1
        If Grilla.TextMatrix(I, 1) = Valor Then
            ControlDuplicado = True
            Exit For
        End If
    Next
End Function

Public Sub InsertarCambioPosicion(COD_CLIENTE As Integer, NRO_CAJA As Long)
'        Dim Estanteria, Horizontal, Vertical, Adelante_Atras, NRO_ESTANTE, _
'        estado As Integer
'        Dim rsContenedor As ADODB.Recordset
'        Sql = "Select * from contenedor where cod_cliente = " & COD_CLIENTE
'        Sql = Sql & " AND NRO_CAJA =" & NRO_CAJA
'        Set rsContenedor = New ADODB.Recordset
'        rsContenedor.Open Sql, ConActiva, 0, 1
'            If Not rsContenedor.EOF Then
'                With rsContenedor
'                    Estanteria = CInt(!Estanteria)
'                    Horizontal = CInt(!Horizontal)
'                    Vertical = CInt(!Vertical)
'                    Adelante_Atras = CInt(!Adelante_Atras)
'                    NRO_ESTANTE = CInt(!NRO_ESTANTE)
'                    estado = CInt(!estado)
'                    COD_CLIENTE = CInt(!COD_CLIENTE)
'                    NRO_CAJA = CLng(!NRO_CAJA)
'                    ID_Personal = CInt(lblIDPersonal)
'                    If IsNull(!F_MODIFICACION) Then
'                        F_MODIFICACION = ""
'                    Else
'                        F_MODIFICACION = CDate(!F_MODIFICACION)
'                    End If
'                    If IsNull(!NRO_REMITO) Then
'                        NRO_REMITO = Null
'                    Else
'                        NRO_REMITO = CLng(!NRO_REMITO)
'                    End If
'                    If IsNull(!IDREQUERIMIENTO) Then
'                        IDREQUERIMIENTO = Null
'                    Else
'                        IDREQUERIMIENTO = CLng(!IDREQUERIMIENTO)
'                    End If
'                    End With
'                    Sql = "INSERT INTO CAMBIOPOSICION (ESTANTERIA, HORIZONTAL, VERTICAL,"
'                    Sql = Sql & vbCrLf & " ADELANTE_ATRAS, NRO_ESTANTE, ESTADO,"
'                    Sql = Sql & vbCrLf & " COD_CLIENTE, NRO_CAJA, FECHA,"
'                    Sql = Sql & vbCrLf & " ID_PERSONAL )"
'                    Sql = Sql & vbCrLf & " VALUES (" & Estanteria & "," & Horizontal & "," & Vertical & ","
'                    Sql = Sql & vbCrLf & Adelante_Atras & "," & NRO_ESTANTE & " ," & estado & ","
'                    Sql = Sql & vbCrLf & COD_CLIENTE & "," & NRO_CAJA & "," & SysDate & ","
'                    Sql = Sql & vbCrLf & ID_Personal & " )"
'                    ExecutarSql (Sql)
'            End If
End Sub

'Public Sub ActualizarPosicionNueva(NRO_CAJA As Integer, Cod_Cliente As Integer, Estanteria As Integer, Horizontal As Integer, Vertical As Integer, Adelante_Atras As Integer)
'        Dim rsContenedor As OraDynaset
'        Dim Estado As Integer
'        Dim NRO_REMITO, F_MODIFICACION, IDRequerimiento As String
'        sql = "Select * from contenedor where cod_cliente = " & Cod_Cliente
'        sql = sql & " AND NRO_CAJA =" & NRO_CAJA
'        Set rsContenedor = OraDatabase.CreateDynaset(sql, ORADYN_READONLY)
'        If Not rsContenedor.EOF Then
'
'
'                    If IsNull(NRO_REMITO) Then
'                        sql = sql & vbCrLf & ", NRO_REMITO= NULL"
'                    Else
'                        sql = sql & vbCrLf & ", NRO_REMITO=" & NRO_REMITO
'                    End If
'                    If (F_MODIFICACION) = "" Then
'                        sql = sql & vbCrLf & ", F_MODIFICACION= NULL"
'                    Else
'                        sql = sql & vbCrLf & ", F_MODIFICACION=" & F_MODIFICACION
'                    End If
'                    If IsNull(IDRequerimiento) Then
'                        sql = sql & vbCrLf & " , IDREQUERIMIENTO=NULL"
'                    Else
'                        sql = sql & vbCrLf & " , IDREQUERIMIENTO=" & IDRequerimiento
'                    End If
'                    sql = " UPDATE CONTENEDOR SET ESTADO=" & Estado & ","
'                    sql = sql & vbCrLf & " COD_CLIENTE=" & Cod_Cliente & ", NRO_CAJA=" & NRO_CAJA
'
'                    sql = sql & vbCrLf & " Where Estanteria = " & Estanteria & " And Horizontal = " & Horizontal & ""
'                    sql = sql & vbCrLf & " And Vertical = " & Vertical & " And ADELANTE_ATRAS = " & Adelante_Atras
'                    OraDatabase.ExecuteSQL (sql)
'
'
'End Sub

Public Function BuscarCliente(ID_CAJA As Long) As Integer
    Dim RS As New ADODB.Recordset
    
    RS.Open "SELECT    FK_CLIENTE From Cajas Where ID_CAJA = " & ID_CAJA, ConActiva, 0, 1
    
    If RS.EOF Then
        BuscarCliente = 0
    Else
        If IsNull(RS!FK_CLIENTE) Then
            BuscarCliente = 0
        Else
            BuscarCliente = RS!FK_CLIENTE
        End If
        
        
    End If
    

End Function

Public Sub MovimientoCajas1()

Dim SQL As String
Dim ConCambio As New ADODB.Connection

ConCambio.Open strConBasa
On Error GoTo salir:
ConCambio.BeginTrans
    Dim R As Integer
       If Bloque = True Then
             If MsgBox("Marcar Bloque ", vbYesNo) = vbYes Then
                Bloque = True
            End If
        End If
         For R = 1 To grdCambioPosicion.Rows - 1
            SQL = " INSERT INTO CAMBIOPOSICION"
            SQL = SQL & " (ID_PERSONAL, FECHA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA)"
            SQL = SQL & "  SELECT    " & MDIfrmInicio.StaInicio.Panels(2).Text & ", GETDATE() AS FECHA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO,"
            SQL = SQL & "  COD_CLIENTE , NRO_CAJA"
            SQL = SQL & "  From CONTENEDOR"
            SQL = SQL & "  Where ID_CONTENEDOR = " & grdCambioPosicion.TextMatrix(R, 0)
            ConCambio.Execute SQL
            
            SQL = " Update CONTENEDOR"
            SQL = SQL & " SET  "
            SQL = SQL & " ESTADO =1"
            SQL = SQL & ", COD_CLIENTE =Null"
            SQL = SQL & " , NRO_CAJA =" & grdCambioPosicion.TextMatrix(R, 0)
            SQL = SQL & " , NRO_REMITO =Null"
            SQL = SQL & " , UB_PROVISORIA =Null"
            SQL = SQL & "  Where ID_CONTENEDOR = " & grdCambioPosicion.TextMatrix(R, 0)
            ConCambio.Execute SQL
            
            
            SQL = " Update CONTENEDOR"
            SQL = SQL & " SET  "
            SQL = SQL & " ESTADO =" & grdCambioPosicion.TextMatrix(R, 3)
            SQL = SQL & ", COD_CLIENTE =" & grdCambioPosicion.TextMatrix(R, 2)
            SQL = SQL & " , NRO_CAJA =" & grdCambioPosicion.TextMatrix(R, 1)
            SQL = SQL & " , FECHAPOSICION = " & SysDate2
            If Bloque = True Then
                
                    SQL = SQL & " ,FECHABLOQUE=" & SysDate
             End If
            
            SQL = SQL & "  Where ID_CONTENEDOR = " & grdCambioPosicion.TextMatrix(R, 8)
            ConCambio.Execute SQL
            lblContar.Caption = R
            lblContar.Refresh
            frmCambioPosicionFisica.Refresh
        Next

ConCambio.CommitTrans
MsgBox "Terminado"
cmdCancelar_Click
Exit Sub

salir:


ConCambio.RollbackTrans

MsgBox Err.Description
End Sub
