VERSION 5.00
Begin VB.Form frmProcesos 
   Caption         =   "Form1"
   ClientHeight    =   11115
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   11115
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command105 
      Caption         =   "Command105"
      Height          =   615
      Left            =   12120
      TabIndex        =   128
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton Command104 
      Caption         =   "Command104"
      Height          =   435
      Left            =   10980
      TabIndex        =   127
      Top             =   60
      Width           =   1515
   End
   Begin VB.CommandButton Command103 
      Caption         =   "Command103"
      Height          =   375
      Left            =   11460
      TabIndex        =   126
      Top             =   1200
      Width           =   1395
   End
   Begin VB.CommandButton Command102 
      Caption         =   "ECOGAS INDICES"
      Height          =   495
      Left            =   8640
      TabIndex        =   125
      Top             =   900
      Width           =   2175
   End
   Begin VB.CommandButton Command101 
      Caption         =   "Command101"
      Height          =   555
      Left            =   11460
      TabIndex        =   124
      Top             =   420
      Width           =   1215
   End
   Begin VB.CommandButton Command100 
      Caption         =   "Command100"
      Height          =   615
      Left            =   8580
      TabIndex        =   123
      Top             =   0
      Width           =   1995
   End
   Begin VB.CommandButton Command99 
      Caption         =   "Command99"
      Height          =   975
      Left            =   8340
      TabIndex        =   122
      Top             =   3060
      Width           =   2055
   End
   Begin VB.CommandButton cmdReferenciasFaltantesDisco 
      Caption         =   "cmdReferenciasFaltantesDisco"
      Height          =   615
      Left            =   5700
      TabIndex        =   121
      Top             =   10380
      Width           =   1995
   End
   Begin VB.CommandButton Command98 
      Caption         =   "Command98"
      Height          =   735
      Left            =   1440
      TabIndex        =   120
      Top             =   10320
      Width           =   1695
   End
   Begin VB.CommandButton Command97 
      Caption         =   "Command97"
      Height          =   615
      Left            =   11160
      TabIndex        =   119
      Top             =   10320
      Width           =   1815
   End
   Begin VB.CommandButton cmdExpresoLujan 
      Caption         =   "Expreso Lujan"
      Height          =   735
      Left            =   5460
      TabIndex        =   118
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton cmdBajasDatas 
      Caption         =   "BajasDatas"
      Height          =   735
      Left            =   3360
      TabIndex        =   116
      Top             =   10200
      Width           =   1995
   End
   Begin VB.CommandButton Command95 
      Caption         =   "FONDO2"
      Height          =   555
      Left            =   8280
      TabIndex        =   115
      Top             =   10200
      Width           =   2355
   End
   Begin VB.CommandButton Command94 
      Caption         =   "Command94"
      Height          =   615
      Left            =   0
      TabIndex        =   114
      Top             =   10200
      Width           =   1335
   End
   Begin VB.CommandButton Command93 
      Caption         =   "GARBARINO"
      Height          =   1095
      Left            =   4380
      TabIndex        =   113
      Top             =   8580
      Width           =   3255
   End
   Begin VB.CommandButton Command92 
      Caption         =   "Command92"
      Height          =   495
      Left            =   0
      TabIndex        =   112
      Top             =   8880
      Width           =   2655
   End
   Begin VB.CommandButton Command91 
      Caption         =   "Command91"
      Height          =   495
      Left            =   720
      TabIndex        =   111
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton Command90 
      Caption         =   "Command90"
      Height          =   375
      Left            =   480
      TabIndex        =   110
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command89 
      Caption         =   "Command89"
      Height          =   615
      Left            =   480
      TabIndex        =   109
      Top             =   9240
      Width           =   1575
   End
   Begin VB.CommandButton Command88 
      Caption         =   "Command88"
      Height          =   615
      Left            =   480
      TabIndex        =   108
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton Command87 
      Caption         =   "Command87"
      Height          =   855
      Left            =   360
      TabIndex        =   107
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton Command86 
      Caption         =   "fondo"
      Height          =   735
      Left            =   360
      TabIndex        =   106
      Top             =   7380
      Width           =   2055
   End
   Begin VB.CommandButton cmdReparacionMiguel 
      Caption         =   "Reparacion Miguel"
      Height          =   375
      Left            =   360
      TabIndex        =   105
      Top             =   5520
      Width           =   2415
   End
   Begin VB.CommandButton Command85 
      Caption         =   "Command85"
      Height          =   1575
      Left            =   1680
      TabIndex        =   104
      Top             =   2760
      Width           =   3735
   End
   Begin VB.CommandButton Command84 
      Caption         =   "Command84"
      Height          =   615
      Left            =   9360
      TabIndex        =   103
      Top             =   9360
      Width           =   2295
   End
   Begin VB.CommandButton cmdLegajosFondo 
      Caption         =   "Legajos Fondo"
      Height          =   495
      Left            =   10320
      TabIndex        =   101
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdRecuperacionCajas 
      Caption         =   "Recuperacion Cajas"
      Height          =   615
      Left            =   8280
      TabIndex        =   100
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command81 
      Caption         =   "Command81"
      Height          =   1335
      Left            =   3720
      TabIndex        =   98
      Top             =   3840
      Width           =   4215
   End
   Begin VB.CommandButton Command80 
      Caption         =   "Command80"
      Height          =   855
      Left            =   8280
      TabIndex        =   97
      Top             =   3240
      Width           =   2655
   End
   Begin VB.CommandButton Command79 
      Cancel          =   -1  'True
      Caption         =   "aLSINA"
      Height          =   495
      Left            =   9960
      TabIndex        =   96
      Top             =   2880
      Width           =   3375
   End
   Begin VB.CommandButton Command78 
      Caption         =   "Command78"
      Height          =   975
      Left            =   3960
      TabIndex        =   95
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton Command77 
      Caption         =   "Command77"
      Height          =   855
      Left            =   2400
      TabIndex        =   94
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command76 
      Caption         =   "Command76"
      Height          =   975
      Left            =   2520
      TabIndex        =   93
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command75 
      Caption         =   "Command75"
      Height          =   495
      Left            =   2040
      TabIndex        =   92
      Top             =   9480
      Width           =   735
   End
   Begin VB.CommandButton Command74 
      Caption         =   "Command74"
      Height          =   735
      Left            =   8280
      TabIndex        =   91
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command73 
      Caption         =   "Command73"
      Height          =   855
      Left            =   3960
      TabIndex        =   90
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command72 
      Caption         =   "Command72"
      Height          =   975
      Left            =   5640
      TabIndex        =   89
      Top             =   4560
      Width           =   2535
   End
   Begin VB.CommandButton Command71 
      Caption         =   "Command71"
      Height          =   855
      Left            =   9480
      TabIndex        =   88
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton Command70 
      Caption         =   "Command70"
      Height          =   735
      Left            =   8520
      TabIndex        =   87
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command68 
      Caption         =   "Command68"
      Height          =   615
      Left            =   9480
      TabIndex        =   85
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command66 
      Caption         =   "Command66"
      Height          =   975
      Left            =   2880
      TabIndex        =   83
      Top             =   9000
      Width           =   615
   End
   Begin VB.CommandButton Command65 
      Caption         =   "Command65"
      Height          =   615
      Left            =   3480
      TabIndex        =   82
      Top             =   4440
      Width           =   1935
   End
   Begin VB.CommandButton Command64 
      Caption         =   "Command64"
      Height          =   495
      Left            =   6240
      TabIndex        =   81
      Top             =   8760
      Width           =   2535
   End
   Begin VB.CommandButton Command63 
      Caption         =   "Command63"
      Height          =   735
      Left            =   6000
      TabIndex        =   80
      Top             =   9240
      Width           =   2415
   End
   Begin VB.CommandButton Command62 
      BackColor       =   &H8000000D&
      Caption         =   "tps"
      Height          =   615
      Left            =   120
      TabIndex        =   79
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton Command61 
      Caption         =   "Command61"
      Height          =   615
      Left            =   3480
      TabIndex        =   78
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton Command59 
      Caption         =   "Command59"
      Height          =   615
      Left            =   11040
      TabIndex        =   76
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton Command56 
      Caption         =   "Command56"
      Height          =   495
      Left            =   3600
      TabIndex        =   73
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command48 
      Caption         =   "Command48"
      Height          =   495
      Left            =   10800
      TabIndex        =   65
      Top             =   4260
      Width           =   1215
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Command45"
      Height          =   615
      Left            =   8460
      TabIndex        =   62
      Top             =   4260
      Width           =   2115
   End
   Begin VB.CommandButton Command44 
      Caption         =   "Command44"
      Height          =   615
      Left            =   2880
      TabIndex        =   61
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4275
      Left            =   2940
      TabIndex        =   56
      Top             =   4800
      Width           =   9615
      Begin VB.CommandButton Command96 
         Caption         =   "Command96"
         Height          =   615
         Left            =   4320
         TabIndex        =   117
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton Command83 
         Caption         =   "Command83"
         Height          =   855
         Left            =   240
         TabIndex        =   102
         Top             =   120
         Width           =   2535
      End
      Begin VB.CommandButton Command82 
         Caption         =   "Command82"
         Height          =   1215
         Left            =   2640
         TabIndex        =   99
         Top             =   1320
         Width           =   3855
      End
      Begin VB.CommandButton Command69 
         Caption         =   "Command69"
         Height          =   735
         Left            =   7440
         TabIndex        =   86
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton Command67 
         Caption         =   "Command67"
         Height          =   615
         Left            =   7560
         TabIndex        =   84
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton Command60 
         Caption         =   "NEXTEL"
         Height          =   255
         Left            =   7800
         TabIndex        =   77
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton Command58 
         Caption         =   "Ubicacion CUstodia"
         Height          =   615
         Left            =   7680
         TabIndex        =   75
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton Command57 
         Caption         =   "PLANILLA"
         Height          =   495
         Left            =   3000
         TabIndex        =   74
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton Command55 
         Caption         =   "Contenedor 25"
         Height          =   615
         Left            =   4200
         TabIndex        =   72
         Top             =   3360
         Width           =   2655
      End
      Begin VB.CommandButton Command54 
         Caption         =   "Contenedor25"
         Height          =   495
         Left            =   2640
         TabIndex        =   71
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command53 
         Caption         =   "DISCO ROLLO"
         Height          =   735
         Left            =   2400
         TabIndex        =   70
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command52 
         Caption         =   "Command52"
         Height          =   615
         Left            =   240
         TabIndex        =   69
         Top             =   3600
         Width           =   2655
      End
      Begin VB.CommandButton Command51 
         Caption         =   "disco tarjetas"
         Height          =   735
         Left            =   480
         TabIndex        =   68
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CommandButton Command50 
         Caption         =   "Command50"
         Height          =   615
         Left            =   4080
         TabIndex        =   67
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CommandButton Command49 
         Caption         =   "Command49"
         Height          =   615
         Left            =   7380
         TabIndex        =   66
         Top             =   900
         Width           =   1935
      End
      Begin VB.CommandButton Command47 
         Caption         =   "Command47"
         Height          =   555
         Left            =   180
         TabIndex        =   64
         Top             =   2040
         Width           =   2235
      End
      Begin VB.CommandButton Command46 
         Caption         =   "Command46"
         Height          =   675
         Left            =   4080
         TabIndex        =   63
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CommandButton Command43 
         Caption         =   "Command43"
         Height          =   555
         Left            =   4140
         TabIndex        =   60
         Top             =   480
         Width           =   2775
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Command42"
         Height          =   435
         Left            =   240
         TabIndex        =   59
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CommandButton Command41 
         Caption         =   "Command41"
         Height          =   375
         Left            =   240
         TabIndex        =   58
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Command40 
         Caption         =   "Command40"
         Height          =   435
         Left            =   240
         TabIndex        =   57
         Top             =   420
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Command39"
      Height          =   315
      Left            =   5580
      TabIndex        =   55
      Top             =   4320
      Width           =   2475
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Command38"
      Height          =   495
      Left            =   7920
      TabIndex        =   54
      Top             =   3240
      Width           =   2775
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Command37"
      Height          =   495
      Left            =   6120
      TabIndex        =   53
      Top             =   3960
      Width           =   2655
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Command36"
      Height          =   375
      Left            =   5580
      TabIndex        =   52
      Top             =   3840
      Width           =   2475
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Command35"
      Height          =   375
      Left            =   5580
      TabIndex        =   51
      Top             =   3360
      Width           =   2475
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Command34"
      Height          =   315
      Left            =   5580
      TabIndex        =   50
      Top             =   2940
      Width           =   2475
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Command33"
      Height          =   315
      Left            =   5580
      TabIndex        =   49
      Top             =   2580
      Width           =   2535
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Command32"
      Height          =   315
      Left            =   2820
      TabIndex        =   48
      Top             =   2400
      Width           =   2535
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Command31"
      Height          =   315
      Left            =   2820
      TabIndex        =   47
      Top             =   1680
      Width           =   2535
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Sistema de  paso ecogas"
      Height          =   315
      Left            =   2820
      TabIndex        =   46
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Command29"
      Height          =   315
      Left            =   2820
      TabIndex        =   45
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   8640
      TabIndex        =   37
      Top             =   1680
      Width           =   4455
      Begin VB.CommandButton cmdCambioUsuario 
         Caption         =   "CambioUsuario"
         Height          =   375
         Left            =   2640
         TabIndex        =   44
         Top             =   1620
         Width           =   1575
      End
      Begin VB.TextBox txtCliente 
         BackColor       =   &H0000C000&
         Height          =   375
         Left            =   1560
         TabIndex        =   43
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox txtUsuarioBorrar 
         Height          =   375
         Left            =   1560
         TabIndex        =   41
         Top             =   780
         Width           =   2655
      End
      Begin VB.TextBox txtUsuarioFinal 
         Height          =   375
         Left            =   1560
         TabIndex        =   39
         Top             =   300
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   180
         TabIndex        =   42
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario a borrar:"
         Height          =   255
         Left            =   180
         TabIndex        =   40
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario definitivo:"
         Height          =   255
         Left            =   180
         TabIndex        =   38
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Command28"
      Height          =   315
      Left            =   5520
      TabIndex        =   36
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Command27"
      Height          =   315
      Left            =   5520
      TabIndex        =   35
      Top             =   1740
      Width           =   2535
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Command26"
      Height          =   315
      Left            =   5520
      TabIndex        =   34
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Command25"
      Height          =   315
      Left            =   5520
      TabIndex        =   33
      Top             =   900
      Width           =   2535
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Command24"
      Height          =   315
      Left            =   5520
      TabIndex        =   32
      Top             =   540
      Width           =   2535
   End
   Begin VB.CommandButton Command23 
      Caption         =   "control de cajas referencia"
      Height          =   315
      Left            =   5520
      TabIndex        =   31
      Top             =   180
      Width           =   2535
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Command22"
      Height          =   315
      Left            =   5400
      TabIndex        =   30
      Top             =   -1200
      Width           =   2535
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Command21"
      Height          =   315
      Left            =   3660
      TabIndex        =   29
      Top             =   4020
      Width           =   2535
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Command20"
      Height          =   315
      Left            =   2820
      TabIndex        =   28
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Command19"
      Height          =   315
      Left            =   3660
      TabIndex        =   27
      Top             =   2220
      Width           =   2535
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Command18"
      Height          =   315
      Left            =   3660
      TabIndex        =   26
      Top             =   3660
      Width           =   2535
   End
   Begin VB.CommandButton Command17 
      Caption         =   "controlEntrada"
      Height          =   315
      Left            =   2820
      TabIndex        =   25
      Top             =   2760
      Width           =   2535
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Command16"
      Height          =   315
      Left            =   3660
      TabIndex        =   24
      Top             =   3300
      Width           =   2535
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Command15"
      Height          =   315
      Left            =   2820
      TabIndex        =   23
      Top             =   900
      Width           =   2535
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   315
      Left            =   3660
      TabIndex        =   22
      Top             =   2940
      Width           =   2535
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Command13"
      Height          =   315
      Left            =   2820
      TabIndex        =   21
      Top             =   480
      Width           =   2535
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   315
      Left            =   3660
      TabIndex        =   20
      Top             =   1860
      Width           =   2535
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   315
      Left            =   3660
      TabIndex        =   19
      Top             =   2580
      Width           =   2535
   End
   Begin VB.CommandButton Command10 
      Caption         =   "fecha custodia"
      Height          =   315
      Left            =   60
      TabIndex        =   18
      Top             =   4380
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   315
      Left            =   2820
      TabIndex        =   17
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   315
      Left            =   60
      TabIndex        =   16
      Top             =   4020
      Width           =   2535
   End
   Begin VB.CommandButton cmdControlPedro 
      Caption         =   "Control_Pedro"
      Height          =   315
      Left            =   60
      TabIndex        =   15
      Top             =   3660
      Width           =   2535
   End
   Begin VB.CommandButton cmdCambionNumeroLectura 
      Caption         =   "Numero de lectura"
      Height          =   315
      Left            =   60
      TabIndex        =   14
      Top             =   3300
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   315
      Left            =   3660
      TabIndex        =   13
      Top             =   1500
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   315
      Left            =   3660
      TabIndex        =   12
      Top             =   1140
      Width           =   2535
   End
   Begin VB.CommandButton cmdPasarCajasOsepCarmen 
      Caption         =   "Pasar Cajas Osep Carmen"
      Height          =   315
      Left            =   3660
      TabIndex        =   11
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Actualizar Documentos Digitales"
      Height          =   315
      Left            =   60
      TabIndex        =   10
      Top             =   4740
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   315
      Left            =   60
      TabIndex        =   9
      Top             =   2940
      Width           =   2535
   End
   Begin VB.CommandButton cmdLeerPlanillaCordoba 
      Caption         =   "Leer Planilla Cordoba"
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   2580
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0FF&
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   5100
      Width           =   2535
   End
   Begin VB.CommandButton cmdCajasCrear 
      Caption         =   "Crear Cajas"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   2220
      Width           =   2535
   End
   Begin VB.CommandButton cmdActualizacionlegajos 
      Caption         =   "actualizacion legajos"
      Height          =   315
      Left            =   60
      TabIndex        =   5
      Top             =   1860
      Width           =   2535
   End
   Begin VB.CommandButton cmdMarcarcajaslegajos 
      Caption         =   "Cajas Con Legajos"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   1500
      Width           =   2535
   End
   Begin VB.CommandButton cmdMarcarcajasdebaja 
      Caption         =   "marcarCajasBajas"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   1140
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "llenar tabla Cajas"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   780
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Montemar y la caja Nro desde"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Peparar Fecha Hasta Legajos"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2535
   End
End
Attribute VB_Name = "frmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim CONCUSTODIA As ADODB.Connection
    Dim ID_indice As Long
Private Sub cmdActualizacionlegajos_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
        Sql = "   SELECT     dbo.LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA, dbo.LECTURA_COLECTOR_CUERPO.FECHA_CREACION,"
        Sql = Sql & " dbo.LECTURA_COLECTOR_CUERPO.DESCRIPCION, dbo.LECTURACOLECTOR.CAJA, dbo.LECTURACOLECTOR.CLIENTE,"
        Sql = Sql & " dbo.CLIENTES.RAZON_SOCIAL, dbo.CONTENEDOR.ESTANTERIA, dbo.CONTENEDOR.HORIZONTAL, dbo.CONTENEDOR.VERTICAL,"
        Sql = Sql & " dbo.CONTENEDOR.Adelante_Atras , dbo.CONTENEDOR.Estado, dbo.CONTENEDOR.UB_PROVISORIA"
        Sql = Sql & " FROM dbo.LECTURA_COLECTOR_CUERPO INNER JOIN"
        Sql = Sql & " dbo.LECTURACOLECTOR ON dbo.LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = dbo.LECTURACOLECTOR.NUMERO_LECTURA INNER JOIN"
        Sql = Sql & " dbo.CLIENTES ON dbo.LECTURACOLECTOR.CLIENTE = dbo.CLIENTES.ID_CLIENTE INNER JOIN"
        Sql = Sql & " dbo.CONTENEDOR ON dbo.LECTURACOLECTOR.CAJA = dbo.CONTENEDOR.NRO_CAJA AND"
        Sql = Sql & " dbo.LECTURACOLECTOR.CLIENTE = dbo.CONTENEDOR.COD_CLIENTE LEFT OUTER JOIN"
        Sql = Sql & " dbo.LEGAJOS ON dbo.LECTURACOLECTOR.CAJA = dbo.LEGAJOS.NRO_CAJA AND"
        Sql = Sql & " dbo.LECTURACOLECTOR.Cliente = dbo.LEGAJOS.COD_CLIENTE"
        Sql = Sql & " WHERE (dbo.LECTURA_COLECTOR_CUERPO.DESCRIPCION LIKE '%LEGAJOS%') AND (dbo.LEGAJOS.ID_LEGAJO IS NULL)"
        Sql = Sql & " ORDER BY dbo.LECTURA_COLECTOR_CUERPO.FECHA_CREACION"
    Dim con As New ADODB.Connection
    con.Open strConBasa
    rs.Open Sql, con
    Do While Not rs.EOF
        Sql = "  Update dbo.Cajas"
        Sql = Sql & " SET  FK_TIPO_REFERENCIA =1010"
        Sql = Sql & "  Where NRO_CAJA = " & rs!Caja
        Sql = Sql & "  And FK_CLIENTE = " & rs!Cliente
        con.Execute Sql
        rs.MoveNext
    Loop
End Sub

Private Sub cmdBajasDatas_Click()


Dim conData As New ADODB.Connection
Dim RsBaja As New ADODB.Recordset
Dim rsTEM_IVA_DATA As New ADODB.Recordset

Dim Sql As String

    
conData.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=BAJASDISCO"



Sql = "SELECT [Cajas]   From [basasql].[dbo].[CAJASBASA2002JOSE]"


RsBaja.Open Sql, strConBasa

    Do While Not RsBaja.EOF
        Sql = " UPDATE UNNAMED SET IDCliente = 1197, Estado = 'OCUPADA'"
        Sql = Sql & vbCrLf & " WHERE IDCaja=" & RsBaja!CAJAS
        conData.Execute Sql
        RsBaja.MoveNext
    Loop





'Sql = "SELECT FacturaABC, NumeroFactura, FechaFacturacion, NombreCliente, "
'Sql = Sql & vbCrLf & "   CUIT,  Subtotal, TotalFacturado"
'Sql = Sql & vbCrLf & " From FACTURA "
' Sql = Sql & vbCrLf & " where FechaFacturacion >= " & FECHADATA_Dias(txtFechaDesde.Text)
' Sql = Sql & vbCrLf & "  And FechaFacturacion <= " & FECHADATA_Dias(txtFechaHasta.Text)
'
' Rem Sql = Sql & vbCrLf & " where FechaFacturacion BETWEEN " & FECHADATA_Dias(txtFechaDesde.Text) & " and " & FECHADATA_Dias(txtFechaHasta.Text)
'  Rem  Sql = Sql & vbCrLf & " order by [FacturaABC]      ,[NumeroFactura] "
'
'  SELECT UNNAMED.[IDCliente], UNNAMED.[IDCaja], UNNAMED.[Estado]
'From UNNAMED
'WHERE (((UNNAMED.[IDCaja])=2234));
'
'
'UPDATE UNNAMED SET UNNAMED.[IDCliente] = 0, UNNAMED.[Estado] = "RESERVADA"
'WHERE (((UNNAMED.[IDCaja])=2234));


End Sub

Private Sub cmdCajasCrear_Click()

On Error GoTo salir:
    Dim rsControlCajas As New ADODB.Recordset
    Dim SqlLectura As String
    Dim Sql As String
    Dim CajasLectura As String
    Dim cantidad As Integer
    Dim Lectura As New ADODB.Recordset
    
    
        SqlLectura = "SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_CONTENEDOR"
        SqlLectura = SqlLectura & "  From dbo.Cajas"
        SqlLectura = SqlLectura & "  WHERE     (ID_CAJA BETWEEN 744282 AND 744345)"
        rsControlCajas.Open SqlLectura, ConActiva, adOpenStatic, adLockReadOnly
        cantidad = rsControlCajas.RecordCount
    
     Dim Caja As Long
     Dim rsContenedor As New ADODB.Recordset
        rsContenedor.CursorLocation = adUseClient
        Sql = " SELECT  TOP " & cantidad & " ID_CONTENEDOR, NRO_CAJA, COD_CLIENTE, ESTADO, F_MODIFICACION , FK_CAJAS "
        Sql = Sql & "  From CONTENEDOR "
        Sql = Sql & "  WHERE ESTADO = 1 AND COD_CLIENTE IS NULL AND"
        Sql = Sql & " NRO_CAJA IS NULL AND ESTANTERIA BETWEEN 150 AND 190 "
        rsContenedor.Open Sql, ConActiva, adOpenKeyset, adLockPessimistic
    
    If rsContenedor.EOF Then
    
        MsgBox "No hay estanterias diponibles"
        Exit Sub
    
    End If
    
    
    
    Do While Not rsControlCajas.EOF
        
        rsContenedor!NRO_CAJA = rsControlCajas!ID_CAJA
        rsContenedor!FK_CAJAS = rsControlCajas!ID_CAJA
        rsContenedor!COD_CLIENTE = rsControlCajas!FK_CLIENTE
        rsContenedor!estado = 5
        
        rsContenedor.Update
        rsContenedor.MoveNext
        rsControlCajas.MoveNext
    Loop
    
        
        
     Exit Sub
salir:
    Rem conVacias.RollbackTrans
    MsgBox "Error en la generacion de cajas"
End Sub

Private Sub cmdCambionNumeroLectura_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Lectura As Integer
Dim cone As New ADODB.Connection

cone.Open strConBasa
cone.BeginTrans

Sql = " SELECT     NUMERO_LECTURA"
Sql = Sql & " From dbo.LECTURA_COLECTOR_CUERPO "
Sql = Sql & " Where (NUMERO_LECTURA > 30099) "

rs.Open Sql, ConActiva, 0, 1
Lectura = 11590

Do While Not rs.EOF
    
   Sql = " Update dbo.LECTURA_COLECTOR_CUERPO"
Sql = Sql & " Set NUMERO_LECTURA = " & Lectura
Sql = Sql & "  Where NUMERO_LECTURA =" & rs!NUMERO_LECTURA
cone.Execute Sql

Sql = " Update dbo.Cajas "
Sql = Sql & " Set FK_LECTURA = " & Lectura
Sql = Sql & " Where FK_LECTURA = " & rs!NUMERO_LECTURA
cone.Execute Sql

Sql = " Update dbo.LECTURACOLECTOR "
Sql = Sql & "  Set NUMERO_LECTURA = " & Lectura
Sql = Sql & "  Where NUMERO_LECTURA = " & rs!NUMERO_LECTURA
cone.Execute Sql
    
rs.MoveNext
        Lectura = Lectura + 1
        
    
Loop
cone.CommitTrans

cone.RollbackTrans

End Sub

Private Sub cmdCambioUsuario_Click()
    Dim Sql As String

    Sql = " Update REMITOS_CUERPO"
    Sql = Sql & " Set COD_USUARIO_CLIENTE = " & txtUsuarioFinal.Text
    Sql = Sql & " WHERE COD_USUARIO_CLIENTE IN (" & txtUsuarioBorrar.Text & ")"
    Sql = Sql & " AND ID_CLIENTE = " & txtCliente.Text
    ExecutarSql Sql
    
    Sql = "  Update REQUERIMIENTO"
    Sql = Sql & "  Set COD_USUARIO_CLIENTE = " & txtUsuarioFinal.Text
    Sql = Sql & " WHERE ID_CLIENTE = " & txtCliente.Text
    Sql = Sql & " AND COD_USUARIO_CLIENTE IN (" & txtUsuarioBorrar.Text & ")"
    ExecutarSql Sql
    
    
    Sql = " Update Cajas "
    Sql = Sql & " Set FK_CLIENTES_USUARIO = " & txtUsuarioFinal.Text
    Sql = Sql & " WHERE  FK_CLIENTES_USUARIO IN (" & txtUsuarioBorrar.Text & ")"
    Sql = Sql & " AND FK_CLIENTE = " & txtCliente.Text
    ExecutarSql Sql
    
    
    Sql = " DELETE FROM CLIENTEUSUARIO "
    Sql = Sql & " WHERE ID_CLIENTEUSUARIO IN (" & txtUsuarioBorrar.Text & ")"
    Sql = Sql & " AND  COD_CLIENTE = " & txtCliente.Text
    ExecutarSql Sql




End Sub

Private Sub cmdControlPedro_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset

rs.CursorLocation = adUseClient

Sql = "SELECT   ID,   CAJA, CLIENTE, ORDEN, ESTANTERIA, VERTICAL, HORIZONTAL"
Sql = Sql & " From CONTROLPEDRO"
Sql = Sql & " ORDER BY id"


rs.Open Sql, ConActiva, adOpenDynamic, adLockPessimistic

Dim cambio As String


Do While Not rs.EOF

    If rs!Cliente = 9999 Then
    
        If cambio <> rs!Caja Then
            cambio = rs!Caja
        End If
        
    Else
        rs!Estanteria = Mid(cambio, 1, 4)
        rs!Vertical = Mid(cambio, 5, 2)
        rs!Horizontal = Mid(cambio, 7, 2)
    End If
        
        
        rs.Update
        
        
    rs.MoveNext
Loop





End Sub

Private Sub cmdExpresoLujan_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset

    Sql = "SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.DESCRIPCION, DOCUMENTOS_DIGITALES.Exportado"
    Sql = Sql & " FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
    Sql = Sql & "                     DOCUMENTOS_DIGITALES ON"
    Sql = Sql & "                     DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " WHERE     (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401) AND (DOCUMENTOS_DIGITALES.DESCRIPCION IS NULL) AND"
    Sql = Sql & "                     (NOT (DOCUMENTOS_DIGITALES.Exportado IS NULL))"
rs.Open Sql, strConBasa

Do While Not rs.EOF
    Sql = " Update DOCUMENTOS_DIGITALES"
    Sql = Sql & "  SET DESCRIPCION ='expoertado:" & Trim(rs!Exportado) & "'"
    Sql = Sql & "  Where ID = " & rs!ID
    ExecutarSql Sql
    rs.MoveNext
Loop


End Sub

Private Sub cmdLeerPlanillaCordoba_Click()
        Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
        Dim R As Integer
        Dim C As Integer
        Dim Sql As String
        Dim KF_CLIENTE As Integer
        Dim NRO_CAJA As Long
        Dim con As New ADODB.Connection
        Dim P As Integer
        Dim Bloque As String
        

        'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open(strPasoPlanillas & "Cordoba_deposito.xls")
        P = libroEx.Worksheets.Count
        
        con.Open strConBasa
       For P = 1 To P
            Set hojaEx = libroEx.Worksheets.Item(P)
            For R = 1 To 100
                For C = 1 To 15
                        If hojaEx.Cells(R, C) <> "" Then
                                 If IsNumeric(hojaEx.Cells(R, C)) Then
                                     NRO_CAJA = hojaEx.Cells(R, C)
                                     KF_CLIENTE = 231
                                Else
                                    If Not IsNumeric(Mid(hojaEx.Cells(R, C), InStr(1, hojaEx.Cells(R, C), "-", vbTextCompare) + 1)) Then
                                        NRO_CAJA = 1
                                        KF_CLIENTE = 1
                                    Else
                                        If Not IsNumeric(Mid(hojaEx.Cells(R, C), 1, InStr(1, hojaEx.Cells(R, C), "-", vbTextCompare) - 1)) Then
                                            NRO_CAJA = 1
                                            KF_CLIENTE = 1
                                        Else
                                            NRO_CAJA = Mid(hojaEx.Cells(R, C), InStr(1, hojaEx.Cells(R, C), "-", vbTextCompare) + 1)
                                            KF_CLIENTE = Mid(hojaEx.Cells(R, C), 1, InStr(1, hojaEx.Cells(R, C), "-", vbTextCompare) - 1)
                                        End If
                                        
                                    End If
                                End If
                         Else
                            NRO_CAJA = 0
                            KF_CLIENTE = 0
                         End If
                         
                         If hojaEx.Cells(R, 16) <> "" Then
                            Bloque = hojaEx.Cells(R, 16)
                         End If
                    Sql = " INSERT INTO dbo.CAJAS_CORDOBA"
                    Sql = Sql & " (KF_CLIENTE, NRO_CAJA, PLANILLA, BLOQUE, COL, FILA, VALOR , PLANILLA_NUMERO )"
                    Sql = Sql & "  VALUES     (" & KF_CLIENTE & "," & NRO_CAJA & ",'" & hojaEx.Name & "','" & Bloque & "'," & C & "," & R & " ,'" & Trim(hojaEx.Cells(R, C)) & "'," & P & ")"
                    con.Execute Sql
               Next
            Next
      Rem      libroEx.Close
      
            Text1.Text = P
            Text1.Refresh
        Next
        
        
        ApExcel.Quit
        Set hojaEx = Nothing
        Set libroEx = Nothing
        Set ApExcel = Nothing
       
End Sub

Private Sub cmdLegajosFondo_Click()

    Dim Sql As String
    Dim rslegajos As New ADODB.Recordset
    Dim RsControl As New ADODB.Recordset
    Sql = "SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE, FK_INDICES, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA,"
    Sql = Sql & vbCrLf & " DESCRIPCION, NRO_CAJA, REARCHIVO_CAJA, COD_CLIENTE, ID_PERSONAL, FK_PERSONAL_CREACION, FECHA_CREACION, FK_PERSONAL_ACTUALIZACION,"
    Sql = Sql & vbCrLf & " FECHA_ACTUALIZACION , Cod_Estado, Fecha"
    Sql = Sql & vbCrLf & " From basasql.dbo.AAAALEG49"
    Sql = Sql & vbCrLf & " ORDER BY ID_LEGAJO"
    
    Dim con As New ADODB.Connection
    
    rslegajos.Open Sql, ConBasa
    con.Open strConBasa
    Dim FECHA_DESDE As String
    Dim FECHA_HASTA As String
      
    Do While Not rslegajos.EOF
    
If (rslegajos!FECHA_DESDE) = "NULL" Then
    FECHA_DESDE = "Null"
Else
    FECHA_DESDE = Mid(rslegajos!FECHA_DESDE, 9, 2) & "/" & Mid(rslegajos!FECHA_DESDE, 6, 2) & "/" & Mid(rslegajos!FECHA_DESDE, 1, 4)
FECHA_DESDE = FechaFormato(FECHA_DESDE)
End If

If (rslegajos!FECHA_HASTA) = "NULL" Then
    FECHA_HASTA = "Null"
Else
    FECHA_HASTA = Mid(rslegajos!FECHA_HASTA, 9, 2) & "/" & Mid(rslegajos!FECHA_HASTA, 6, 2) & "/" & Mid(rslegajos!FECHA_HASTA, 1, 4)
    FECHA_HASTA = FechaFormato(FECHA_HASTA)
End If

     
 Sql = " INSERT INTO LEGAJOS"
 Sql = Sql & vbCrLf & "("
 Sql = Sql & vbCrLf & " ID_LEGAJO"
 Sql = Sql & vbCrLf & ",ID_CLIENTE_LEGAJO"
 Sql = Sql & vbCrLf & ",COD_INDICE"
 Sql = Sql & vbCrLf & ",FK_INDICES"
 Sql = Sql & vbCrLf & ",LETRA_DESDE"
 Sql = Sql & vbCrLf & ",LETRA_HASTA"
 Sql = Sql & vbCrLf & ",NRO_DESDE"
 Sql = Sql & vbCrLf & ",NRO_HASTA"
 Sql = Sql & vbCrLf & ",FECHA_DESDE"
 Sql = Sql & vbCrLf & ",FECHA_HASTA"
 Sql = Sql & vbCrLf & ",DESCRIPCION"
 Sql = Sql & vbCrLf & ",NRO_CAJA"
 Sql = Sql & vbCrLf & ",REARCHIVO_CAJA"
 Sql = Sql & vbCrLf & ",COD_CLIENTE"
 Sql = Sql & vbCrLf & ",COD_ESTADO"
 Sql = Sql & vbCrLf & ",ID_PERSONAL"
 Sql = Sql & vbCrLf & ",FK_PERSONAL_CREACION"
 Rem Sql = Sql & vbCrLf & ",FECHA_CREACION"
 Sql = Sql & vbCrLf & ",FK_PERSONAL_ACTUALIZACION"
  Sql = Sql & vbCrLf & ",FECHA_ACTUALIZACION"
 Sql = Sql & vbCrLf & ")"
Sql = Sql & vbCrLf & " VALUES ("
Sql = Sql & vbCrLf & rslegajos!ID_LEGAJO
 Sql = Sql & vbCrLf & "," & rslegajos!ID_CLIENTE_LEGAJO
 Sql = Sql & vbCrLf & ",'" & Trim(rslegajos!Cod_Indice)
 Sql = Sql & vbCrLf & "'," & rslegajos!FK_INDICES
 Sql = Sql & vbCrLf & ",'" & Trim(rslegajos!LETRA_DESDE)
 Sql = Sql & vbCrLf & "','" & Trim(rslegajos!LETRA_HASTA)
 Sql = Sql & vbCrLf & "'," & rslegajos!NRO_DESDE
 Sql = Sql & vbCrLf & "," & rslegajos!NRO_HASTA
 Sql = Sql & vbCrLf & "," & FECHA_DESDE
 Sql = Sql & vbCrLf & "," & FECHA_HASTA
 Sql = Sql & vbCrLf & ",'" & Trim(rslegajos!Descripcion)
 Sql = Sql & vbCrLf & "'," & rslegajos!NRO_CAJA
 Sql = Sql & vbCrLf & "," & rslegajos!REARCHIVO_CAJA
 Sql = Sql & vbCrLf & "," & rslegajos!COD_CLIENTE
 Sql = Sql & vbCrLf & "," & rslegajos!Cod_Estado
 Sql = Sql & vbCrLf & "," & rslegajos!ID_Personal
 Sql = Sql & vbCrLf & "," & rslegajos!FK_PERSONAL_CREACION
 Rem Sql = Sql & vbCrLf & "," & FechaFormato(Trim(rsLegajos!FECHA_CREACION))
 Sql = Sql & vbCrLf & "," & rslegajos!FK_PERSONAL_ACTUALIZACION
  Sql = Sql & vbCrLf & "," & FechaFormato("26/02/2013")
  Sql = Sql & vbCrLf & ")"
  Set RsControl = New ADODB.Recordset
RsControl.Open " SELECT ID_CLIENTE_LEGAJO From LEGAJOS Where ID_LEGAJO =" & rslegajos!ID_LEGAJO, strConBasa

If RsControl.EOF Then
    ConBasa.Execute Sql
End If

     
     rslegajos.MoveNext
    Loop
    

End Sub

Private Sub cmdMarcarcajasdebaja_Click()
 Dim rs As New ADODB.Recordset
 Dim rs2 As New ADODB.Recordset
 Dim Sql As String
 
 
 Sql = " SELECT     dbo.REMITOS_CUERPO.ID_CLIENTE, dbo.REMITOS_CUERPO.TIPO, dbo.REMITOS_CUERPO.ANULADO, dbo.REMITOS_CUERPO.FECHA,"
 Sql = Sql & "     dbo.REMITOS_CUERPO.ESTADO, dbo.REMITOS_CUERPO.CANTIDAD, dbo.REMITOS_DETALLE.DESDE, dbo.REMITOS_DETALLE.HASTA,"
 Sql = Sql & " dbo.REMITOS_DETALLE.NRO_REMITO"
Sql = Sql & " FROM         dbo.REMITOS_CUERPO INNER JOIN"
Sql = Sql & "  dbo.REMITOS_DETALLE ON dbo.REMITOS_CUERPO.NRO_REMITO = dbo.REMITOS_DETALLE.NRO_REMITO"
Sql = Sql & "  Where (dbo.REMITOS_CUERPO.id_cliente = 29) And (dbo.REMITOS_CUERPO.TIPO = 3)"
Sql = Sql & "  ORDER BY dbo.REMITOS_DETALLE.DESDE"

rs2.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
    
    Sql = " SELECT     COD_CLIENTE, NRO_CAJA, DESCRIPCION"
    Sql = Sql & " From dbo.REFERENCIAS "
    Sql = Sql & " Where (COD_CLIENTE = 20) And (NRO_CAJA = 20)"
    
    
    rs2.Open Sql

    rs.MoveNext
Loop

 
End Sub

Private Sub cmdMarcarcajaslegajos_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    
    Sql = " SELECT     NRO_CAJA, COD_CLIENTE"
    Sql = Sql & " From CAJAS_CON_LEGAJOS "
    Sql = Sql & " ORDER BY COD_CLIENTE, NRO_CAJA"
    
    Dim con As New ADODB.Connection
    
    
    rs.Open Sql, ConActiva, 0, 1
    
    
    con.Open strConBasa
    Do While Not rs.EOF
    
    
    
    Sql = " Update dbo.Cajas"
    Sql = Sql & "  SET  FK_TIPO_REFERENCIA =1025 "
    Sql = Sql & "  Where FK_CLIENTE = " & rs!COD_CLIENTE
    Sql = Sql & " And NRO_CAJA = " & rs!NRO_CAJA
     Sql = Sql & "  AND (FK_TIPO_REFERENCIA IS NULL)"
        con.Execute Sql
        rs.MoveNext
    Loop
    
    
    
    Sql = " SELECT     NRO_CAJA, COD_CLIENTE"
    Sql = Sql & " From CAJAS_CON_LEGAJOS "
    Sql = Sql & " ORDER BY COD_CLIENTE, NRO_CAJA"
    
'    Dim con As New ADODB.Connection
'
'
'    rs.Open sql, strConBasa , 0 ,1
'
'
'    con.Open strConBasa , 0 ,1
'    Do While Not rs.EOF
'
'
'
'    sql = " Update dbo.Cajas"
'    sql = sql & "  SET  FK_TIPO_REFERENCIA =1025 "
'    sql = sql & "  Where FK_CLIENTE = " & rs!COD_CLIENTE
'    sql = sql & " And NRO_CAJA = " & rs!NRO_CAJA
'     sql = sql & "  AND (FK_TIPO_REFERENCIA IS NULL)"
'        con.Execute sql
'        rs.MoveNext
'    Loop
'
    
    
End Sub

Private Sub cmdPasarCajasOsepCarmen_Click()

Dim Sql As String


Dim CAJAS As String

Dim concajas As New ADODB.Connection
concajas.Open ConBasa
CAJAS = " 38896 "

Sql = " Update dbo.CONTENEDOR"
Sql = Sql & " Set COD_CLIENTE = 77"
Sql = Sql & " Where (COD_CLIENTE = 20)"
Sql = Sql & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute Sql



Sql = " Update dbo.cajas "
Sql = Sql & "  Set FK_CLIENTE = 77"
Sql = Sql & "  WHERE     (FK_CLIENTE = 20) "
Sql = Sql & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute Sql


Sql = " Update dbo.REFERENCIAS"
Sql = Sql & " Set COD_CLIENTE = 77"
Sql = Sql & " Where (COD_CLIENTE = 20)"
Sql = Sql & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute Sql

Sql = "  Update dbo.MOV_CAJAS2 "
Sql = Sql & " Set id_cliente = 77 "
Sql = Sql & " Where (Tipo_elemento = 0)"
Sql = Sql & " AND (ELEMENTO IN (" & CAJAS & "))"
Sql = Sql & "AND (ID_CLIENTE = 20)"
concajas.Execute Sql


End Sub

Private Sub cmdRecuperacionCajas_Click()

    Dim rsLectura As New ADODB.Recordset
    Dim Sql As String
    Dim rsContenedor As New ADODB.Recordset
    Dim rsControlContenedor As New ADODB.Recordset
    Dim sqlControlContenedor As String
    Dim rsControlCajas As New ADODB.Recordset
    Dim rsCajas As New ADODB.Recordset
    Dim sqlControlCajas As String
    Dim Sqlc As String
    Dim con As New ADODB.Connection
    
    con.Open strConBasa

        Sql = " SELECT LECTURACOLECTOR.ID, LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE,"
        Sql = Sql & vbCrLf & " LECTURACOLECTOR.ORDEN, CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, CONTENEDOR.ESTADO,"
        Sql = Sql & vbCrLf & " CONTENEDOR.COD_CLIENTE , CONTENEDOR.NRO_CAJA"
        Sql = Sql & vbCrLf & " FROM         LECTURACOLECTOR LEFT OUTER JOIN"
        Sql = Sql & vbCrLf & " CONTENEDOR ON LECTURACOLECTOR.CAJA = CONTENEDOR.NRO_CAJA AND LECTURACOLECTOR.CLIENTE = CONTENEDOR.COD_CLIENTE"
        Sql = Sql & vbCrLf & " Where (LECTURACOLECTOR.NUMERO_LECTURA = " & InputBox("Ingrese el numero de lectura", "Lectura", 0) & " ) "
        Sql = Sql & vbCrLf & " And (LECTURACOLECTOR.Cliente < 9000) "
        Sql = Sql & vbCrLf & "  And (CONTENEDOR.Estanteria Is Null) "
        Sql = Sql & vbCrLf & "  ORDER BY LECTURACOLECTOR.CAJA "
        rsLectura.Open Sql, strConBasa
 
        Sqlc = " SELECT   Top 2000  ID_CONTENEDOR, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA, NRO_REMITO, F_MODIFICACION,"
        Sqlc = Sqlc & vbCrLf & "  IDREQUERIMIENTO, NUEVA, BAJA, UB_PROVISORIA, COD_CAJA, JERAQUIA, COD_INDICE, COD_CLIENTE_USUARIO, COD_RESPONSABLE_POSICION,"
        Sqlc = Sqlc & vbCrLf & " FECHA_CREACION, MODULO_V, MODULO_H, CONTROL, MODULO, COD_REMITO_GUARDA, COD_USUARIO_CLIENTE_GUARDA, COD_INDICE_SECTOR, ORDEN,"
        Sqlc = Sqlc & vbCrLf & " FECHAPOSICION"
        Sqlc = Sqlc & vbCrLf & "  From basasql.dbo.CONTENEDOR"
        Sqlc = Sqlc & vbCrLf & "  WHERE     (ESTANTERIA BETWEEN 150 AND 160)"
        Sqlc = Sqlc & vbCrLf & " AND (ESTADO = 1)"
        Sqlc = Sqlc & vbCrLf & " AND (COD_CLIENTE IS NULL)"
        Sqlc = Sqlc & vbCrLf & "  ORDER BY ESTANTERIA, VERTICAL"
        rsContenedor.Open Sqlc, strConBasa
        
        Sql = "  SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_CONTENEDOR, FK_ESTADO"
        Sql = Sql & vbCrLf & "   From basasql.dbo.Cajas"
        Sql = Sql & vbCrLf & "  WHERE     (ID_CAJA BETWEEN 737161 AND 750924) AND (FK_CLIENTE IS NULL)"
        Sql = Sql & vbCrLf & "   ORDER BY ID_CAJA"
        
        rsCajas.Open Sql, strConBasa
        
        
        Do While Not rsLectura.EOF
        
            sqlControlCajas = " SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_ESTADO"
            sqlControlCajas = sqlControlCajas & " From basasql.dbo.Cajas"
            sqlControlCajas = sqlControlCajas & " Where FK_CLIENTE = " & rsLectura!Cliente
            sqlControlCajas = sqlControlCajas & " And NRO_CAJA = " & rsLectura!Caja
            rsControlCajas.Open sqlControlCajas, strConBasa
            If rsControlCajas.EOF Then
            Sql = " Update basasql.dbo.Cajas"
            Sql = Sql & vbCrLf & " SET  FK_CLIENTE = " & rsLectura!Cliente
            Sql = Sql & vbCrLf & " , NRO_CAJA = " & rsLectura!Caja
            Sql = Sql & vbCrLf & " , FK_ESTADO = 2"
            Sql = Sql & vbCrLf & " Where ID_CAJA = " & rsCajas!ID_CAJA
            con.Execute Sql
            End If
            
        
        
        
            sqlControlContenedor = " SELECT     ID_CONTENEDOR, COD_CLIENTE, NRO_CAJA, ESTADO "
            sqlControlContenedor = sqlControlContenedor & " From basasql.dbo.CONTENEDOR "
            sqlControlContenedor = sqlControlContenedor & " Where COD_CLIENTE = " & rsLectura!Cliente
            sqlControlContenedor = sqlControlContenedor & " And NRO_CAJA = " & rsLectura!Caja
            rsControlContenedor.Open sqlControlContenedor, strConBasa
            If rsControlContenedor.EOF Then
                Sql = " Update basasql.dbo.CONTENEDOR"
                Sql = Sql & vbCrLf & " SET  COD_CLIENTE =" & rsLectura!Cliente
                Sql = Sql & vbCrLf & " , NRO_CAJA =" & rsLectura!Caja
                Sql = Sql & vbCrLf & " , ESTADO =2 "
                Sql = Sql & vbCrLf & " Where ID_CONTENEDOR = " & rsContenedor!ID_CONTENEDOR
                con.Execute Sql
            End If
            rsContenedor.MoveNext
            rsLectura.MoveNext
            rsCajas.MoveNext
        Loop
        
    
    


End Sub

Private Sub cmdReferenciasFaltantesDisco_Click()
Dim con As New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=D:\CAJASCUSTODIA\ReferenciasDisco.mdb"
Dim rs As New ADODB.Recordset
Dim Sql As String

Dim FECHA_DESDE As String
Dim FECHA_HASTA As String
Dim Caja As Long
Dim Descripcion As String

Sql = " SELECT CAJASDISCOBASA.Id, CAJASDISCOBASA.CLIENTE, CAJASDISCOBASA.CAJAS, CAJASDISCOBASA.ESTADO, DCTO0197.IDTIPODOCUMENTO, DCTO0197.NOMBRETIPODOCUMENTO, DCTO0197.DESCRIPCION, DCTO0197.FECHADESDE, DCTO0197.FECHAHASTA"
Sql = Sql & " FROM CAJASDISCOBASA INNER JOIN DCTO0197 ON CAJASDISCOBASA.CAJAS = DCTO0197.IDCAJA"
Sql = Sql & " ORDER BY CAJASDISCOBASA.CAJAS;"

Dim con2 As New ADODB.Connection
con2.Open strConBasa


Set rs = New ADODB.Recordset
rs.Open Sql, con

Do While Not rs.EOF
    FECHA_DESDE = DateAdd("d", rs!fechadesde, "28/12/1800")
    FECHA_HASTA = DateAdd("d", rs!FechaHasta, "28/12/1800")
    Caja = rs!CAJAS
    Descripcion = UCase(Trim(rs!Descripcion)) & "/" & UCase(Trim(Replace(rs!NOMBRETIPODOCUMENTO, "", "")))
     Descripcion = Trim(Descripcion)
    Descripcion = Replace(Descripcion, vbCr, "")
    Descripcion = Replace(Descripcion, "  ", "")
    Descripcion = Mid(Replace(Descripcion, ".", ""), 1, Len(Descripcion) - 1)
    Sql = "  INSERT INTO "
    Sql = Sql & vbCrLf & " REFERENCIAS_DISCO_27102014("
    Sql = Sql & vbCrLf & " COD_CLIENTE"
    Sql = Sql & vbCrLf & " ,NRO_CAJA"
    Sql = Sql & vbCrLf & " ,COD_TIPO_ALMACENAMIENTO"
    Sql = Sql & vbCrLf & " ,INDICE"
    Sql = Sql & vbCrLf & " ,DESCRIPCION"
    Sql = Sql & vbCrLf & " ,FECHA_DESDE"
    Sql = Sql & vbCrLf & " ,FECHA_HASTA"
    Sql = Sql & vbCrLf & " ,FECHA_MODIFICACION"
    Sql = Sql & vbCrLf & " ,FECHA_CREACION"
    Sql = Sql & vbCrLf & " ,USUARIO_MODIFICACION"
    Sql = Sql & vbCrLf & " ,FK_PERSONAL_CREACION"
    Sql = Sql & vbCrLf & " ,FK_PERSONAL_MODIFICACION"
    Sql = Sql & vbCrLf & " ,BORRADO"
    Sql = Sql & vbCrLf & " )"
    Sql = Sql & vbCrLf & " VALUES ("
    Sql = Sql & vbCrLf & " 1197"
    Sql = Sql & vbCrLf & " ," & Caja
    Sql = Sql & vbCrLf & " , 0"
    Sql = Sql & vbCrLf & " ,'001'"
    Sql = Sql & vbCrLf & " ,'" & Descripcion
    Sql = Sql & vbCrLf & " ','" & FECHA_DESDE & "'"
    Sql = Sql & vbCrLf & " ,'" & FECHA_HASTA & "'"
    Sql = Sql & vbCrLf & " ,'27/10/2014'"
    Sql = Sql & vbCrLf & " ,'27/10/2014', '17', 17, 17, '0')"
    con2.Execute Sql
    rs.MoveNext
Loop

End Sub

Private Sub cmdReparacionMiguel_Click()


Dim Sql As String
Dim ConCambio As New ADODB.Connection
ConCambio.Open strConBasa

Dim rs As New ADODB.Recordset
On Error GoTo salir:
ConCambio.BeginTrans
Dim R As Integer



    
   Sql = " SELECT     ALSINAFINAL.ID_CONTENEDOR, ALSINAFINAL.COD_CLIENTE, ALSINAFINAL.NRO_CAJA, ALSINAFINAL.LECTURA, ALSINAFINAL.EMPRESA,"
    Sql = Sql & "                     ALSINAFINAL.BARRAANTERIOR, ALSINAFINAL.ESTANTERIA, ALSINAFINAL.HORIZONTAL, ALSINAFINAL.VERTICAL, ALSINAFINAL.FECHA_CREACION,"
    Sql = Sql & "                  ALSINAFINAL.FECHACONTROL, ALSINAFINAL.ESTADOCONTROL, ALSINAFINAL.ESTADOCAJA, ALSINAFINAL.ESTADOCONTENEDOR, ALSINAFINAL.ESTADOALSINA,"
      Sql = Sql & "                   CONTENEDOR.ESTANTERIA AS Expr1, CONTENEDOR.HORIZONTAL AS Expr2, CONTENEDOR.VERTICAL AS Expr3, CONTENEDOR.ADELANTE_ATRAS,"
       Sql = Sql & "                  CONTENEDOR.ID_CONTENEDOR AS ID_CONTENEDOR_BASA, CONTENEDOR.ESTADO,"
        Sql = Sql & "                 CONTENEDOR_1.ID_CONTENEDOR AS ID_CONTENEDOR_FINAL"
 Sql = Sql & "   FROM         ALSINAFINAL INNER JOIN"
  Sql = Sql & "                       CONTENEDOR ON ALSINAFINAL.COD_CLIENTE = CONTENEDOR.COD_CLIENTE AND ALSINAFINAL.NRO_CAJA = CONTENEDOR.NRO_CAJA INNER JOIN"
    Sql = Sql & "                     CONTENEDOR AS CONTENEDOR_1 ON ALSINAFINAL.ESTANTERIA = CONTENEDOR_1.ESTANTERIA AND"
    Sql = Sql & "                     ALSINAFINAL.Horizontal = CONTENEDOR_1.Horizontal And ALSINAFINAL.Vertical = CONTENEDOR_1.Vertical"
 Sql = Sql & "   WHERE     (CONTENEDOR.ESTANTERIA IN (5079, 5080, 5081)) AND (CONTENEDOR.MODULO_H = 1) AND (CONTENEDOR.FECHAPOSICION IS NULL)"
    
    
    rs.Open Sql, strConBasa
Do While Not rs.EOF





Sql = " INSERT INTO CAMBIOPOSICION"
Sql = Sql & " (ID_PERSONAL, FECHA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA)"
Sql = Sql & "  SELECT      17 , GETDATE() AS FECHA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO,"
Sql = Sql & "  COD_CLIENTE , NRO_CAJA"
Sql = Sql & "  From CONTENEDOR"
Sql = Sql & "  Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR_BASA
ConCambio.Execute Sql

Sql = " Update CONTENEDOR"
Sql = Sql & " SET  "
Sql = Sql & " ESTADO =1"
Sql = Sql & ", COD_CLIENTE =Null"
Sql = Sql & " , NRO_CAJA =" & rs!ID_CONTENEDOR_BASA
Sql = Sql & " , NRO_REMITO =Null"
Sql = Sql & " , UB_PROVISORIA =Null"
Sql = Sql & "  Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR_BASA
ConCambio.Execute Sql


Sql = " Update CONTENEDOR"
Sql = Sql & " SET  "
Sql = Sql & " ESTADO =" & rs!estado
Sql = Sql & ", COD_CLIENTE =" & rs!COD_CLIENTE
Sql = Sql & " , NRO_CAJA =" & rs!NRO_CAJA
Sql = Sql & " , FECHAPOSICION = " & SysDate2
Sql = Sql & "  Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR_FINAL


ConCambio.Execute Sql

rs.MoveNext

Loop
ConCambio.CommitTrans
MsgBox "Terminado"

Exit Sub

salir:


ConCambio.RollbackTrans

MsgBox Err.Description


End Sub

Private Sub Command1_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset



Sql = "  SELECT     ID_LEGAJO, FECHA_HASTA, ID_CLIENTE_LEGAJO, COD_INDICE, LETRA_DESDE, LETRA_HASTA, NRO_HASTA, FECHA_DESDE, DATEPART(MM, FECHA_HASTA)"
Sql = Sql & " AS Expr1, CLIENTE_LEGAJO, COD_CLIENTE, FK_INDICES"
Sql = Sql & " From dbo.LEGAJOS"
Sql = Sql & " WHERE     (DATEPART(MM, FECHA_DESDE) = 1) "
Sql = Sql & " AND (DATEPART(dd, FECHA_DESDE) = 1) "
Sql = Sql & " AND (FECHA_HASTA  IS NULL) "


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

Private Sub Command10_Click()
'Dim Fecha As String
'Fecha CUSTODIA
'Fecha = DateAdd("d", -76298, "20/11/2009")
'28/12/1800
End Sub

Private Sub Command100_Click()
Dim MyName As String
Dim VarTexto As String
Dim Caja As Long
Dim CajaAnterior As Long
Dim Sql As String
Dim IDLegajo As Long

Dim con As New ADODB.Connection
con.Open strConBasa
Dim Paso As String

Paso = "Z:\Para Ver\291_03\"
MyName = Dir(Paso & "*.txt", vbDirectory)
Dim P As Integer
Close #1
Do While MyName <> ""
 Caja = Mid(MyName, 7, 7)
     Sql = "  Insert "
            Sql = Sql & " Into basasql.dbo.CONTROLFONDO3(Caja, Legajo, Archivo)"
            Sql = Sql & " VALUES (" & Caja & "," & 0 & ",'" & MyName & "')"
            con.Execute Sql
 
    Open Paso & MyName For Input As #1

            
    Do Until EOF(1)
    
    CajaAnterior = 0
        Line Input #1, VarTexto
        If VarTexto <> "" Then
        IDLegajo = 0
        
'             If Mid(VarTexto, 20, 4) = "BASA" Or Mid(VarTexto, 20, 4) = "VBAS" Or Mid(VarTexto, 20, 4) = "ESTA" Then
''             Caja = Mid(VarTexto, 5, 9)
''              If CajaAnterior <> Caja Then
''
''                CajaAnterior = Caja
''              End If
''
'             Else
'
             
            IDLegajo = Mid(VarTexto, 5, 9)
            Sql = "  Insert "
            Sql = Sql & " Into basasql.dbo.CONTROLFONDO2(Caja, Legajo, Archivo)"
            Sql = Sql & " VALUES (" & Caja & "," & IDLegajo & ",'" & MyName & "')"
            con.Execute Sql
             
'             End If
            


 
 
        End If
    Loop
    Close #1
    MyName = Dir()
Loop
End Sub

Private Sub Command101_Click()
    Dim Sql As String
    
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Sql = " SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_REMITO_CUSTODIA"
Sql = Sql & " From CAJAS "
Sql = Sql & " Where (FK_CLIENTE < 1000)"
Sql = Sql & " ORDER BY FK_CLIENTE, NRO_CAJA"

rs.Open Sql, strConBasa

Do While Not rs.EOF

    Sql = " SELECT REMITOS_DETALLE.NRO_REMITO "
    Sql = Sql & " FROM REMITOS_CUERPO LEFT OUTER JOIN"
    Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO"
    Sql = Sql & " WHERE (REMITOS_CUERPO.TIPO = 0) "
    Sql = Sql & " AND REMITOS_CUERPO.ID_CLIENTE =" & rs!FK_CLIENTE
    Sql = Sql & " AND ("
    Sql = Sql & rs!NRO_CAJA
    Sql = Sql & " BETWEEN REMITOS_DETALLE.DESDE AND REMITOS_DETALLE.HASTA)"
    Set rs2 = New ADODB.Recordset
    rs2.Open Sql, strConBasa
    If Not rs2.EOF Then
     Sql = " Update CAJAS"
        Sql = Sql & "  SET  FK_REMITO_CUSTODIA =" & rs2!NRO_REMITO
        Sql = Sql & "  Where ID_CAJA =" & rs!ID_CAJA
        ExecutarSql Sql
    
    End If
    
     
rs.MoveNext
Loop





End Sub

Private Sub Command102_Click()
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Sql As String
        


'SQL = " SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA"
'SQL = SQL & " From CAJAS"
'SQL = SQL & "  Where (FK_CLIENTE = 4)"
'SQL = SQL & "  Order By NRO_CAJA"

Sql = " SELECT ID_CAJA, FK_CLIENTE, INDICES, NRO_CAJA, FK_ESTADO"
Sql = Sql & "  From CAJAS"
Sql = Sql & "  Where (FK_CLIENTE = 4) "
Sql = Sql & "  And (INDICES Is Null)"
Sql = Sql & "  Order By NRO_CAJA"

rs.Open Sql, strConBasa


Do While Not rs.EOF
'    SQL = " SELECT     INDICE "
'    SQL = SQL & "  From REFERENCIAS "
'    SQL = SQL & "  Where COD_CLIENTE = " & RS!FK_CLIENTE
'    SQL = SQL & "  And NRO_CAJA = " & RS!NRO_CAJA
    
    Sql = " SELECT     COD_INDICE, NRO_CAJA, COD_CLIENTE"
    Sql = Sql & " From basasql.dbo.LEGAJOS"
    Sql = Sql & "  Where (COD_CLIENTE = 4)"
    Sql = Sql & " And NRO_CAJA =" & rs!NRO_CAJA
    
    
    Set rs2 = New ADODB.Recordset
    rs2.Open Sql, strConBasa
    If Not rs2.EOF Then
        If Not IsNull(rs2!Cod_Indice) Then
            Sql = " Update CAJAS "
            Sql = Sql & "  SET INDICES = '" & Mid(rs2!Cod_Indice, 1, 9) & "'"
            Sql = Sql & "  Where ID_CAJA =" & rs!ID_CAJA
            ExecutarSql Sql
        End If
    End If

    rs.MoveNext

Loop




End Sub

Private Sub Command103_Click()
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Sql As String
        


'SQL = " SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA"
'SQL = SQL & " From CAJAS"
'SQL = SQL & "  Where (FK_CLIENTE = 4)"
'SQL = SQL & "  Order By NRO_CAJA"

Sql = " SELECT ID_CAJA, FK_CLIENTE, INDICES, NRO_CAJA, FK_ESTADO"
Sql = Sql & "  From CAJAS"
Sql = Sql & "  Where (FK_CLIENTE = 4) "
Sql = Sql & "  And not (INDICES Is Null)"
Sql = Sql & "  Order By NRO_CAJA"

rs.Open Sql, strConBasa


Do While Not rs.EOF
            Sql = " Update CAJAS "
            Sql = Sql & "  SET INDICES = '" & Mid(rs!INDICES, 1, 6) & "'"
            Sql = Sql & "  Where ID_CAJA =" & rs!ID_CAJA
ExecutarSql Sql
    rs.MoveNext

Loop

End Sub

Private Sub Command104_Click()

Dim Sql As String
Dim sArchivo As String

Dim con As New ADODB.Connection
con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\factura\Database12.accdb;Persist Security Info=False"


Dim rs As New ADODB.Recordset

Dim tabla As String
Dim Cliente As Integer

  sArchivo = Dir("D:\DISCOSUCURSALES" & "\*.TPS")
    Do While sArchivo <> ""
        tabla = Mid(sArchivo, 1, 8)
        Cliente = Mid(sArchivo, 5, 4)
        
        Sql = " SELECT IDCAJA "
        Sql = Sql & " FROM " & tabla
        Set rs = New ADODB.Recordset
        rs.Open Sql, con
        
        Do While Not rs.EOF
            Sql = " Insert  "
            Sql = Sql & " DISCO_SUSURSALES_CAJAS ("
            Sql = Sql & " CLIENTE_CUSTODIA, CAJA )"
            Sql = Sql & " VALUES  (" & Cliente & "," & rs!IDCaja & ")"
            ExecutarSql Sql
            rs.MoveNext
        Loop
        
        
        sArchivo = Dir
    Loop





End Sub

Private Sub Command105_Click()
    Dim comPlanilla As New ADODB.Connection
    Dim RsPlanilla As New ADODB.Recordset
    Dim Sql As String
    Dim i As Integer
    Dim FECHA_DESDE(4) As String
    Dim FECHA_HASTA(4) As String
    Dim NUMERO_DESDE(4) As String
    Dim NUMERO_HASTA(4) As String
    Dim Indice(4) As String
    Dim Descripcion(4) As String
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet
    Dim C_Error As Integer
    Dim C_Caja As Integer
    Dim C_Indice As Integer
    Dim C_Etiqueta As Integer
    Dim C_Fecha_desde As Integer
    Dim C_Fecha_hasta As Integer
    Dim C_N°_Desde As Integer
    Dim C_N°_Hasta As Integer
    Dim C_Letra_Desde As Integer
    Dim C_Letra_Hasta As Integer
    Dim C_Descripcion As Integer
    Dim R As String
    Dim ErrorGeneral As Boolean
    Dim strError As String
    Dim FechaHora As String
    Dim NombreArchivo As String
    
    FechaHora = Trim(Format(Now, "hhmmss"))
    
 
    C_Error = 1
    C_Caja = 2
    C_Indice = 4
    C_Etiqueta = 3
    C_Fecha_desde = 6
    C_Fecha_hasta = 7
    C_N°_Desde = 8
    C_N°_Hasta = 9
    C_Letra_Desde = 10
    C_Letra_Hasta = 11
    C_Descripcion = 5
    
    
    
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Open("\\Server-basa\Sistemas\Referencias\Planilla Modelo.xls", , True)
    Set hojaEx = libroEx.Worksheets.Item(1)
   
            'SELECT Referencias.[Remote_User], Referencias.[Time_Stamp], Referencias.[Suspense_File], Referencias.[Remote_Uid], Referencias.[Remote_Fax], Referencias.[Remote_Bid], Referencias.[Remote_Cmp], Referencias.[Remote_Phn], Referencias.[CSID], Referencias.[Verify_Wks], Referencias.[Form_Id], Referencias.[BatchNo], Referencias.[BatchDir], Referencias.[BatchPgNo], Referencias.[BatchPgCnt], Referencias.[BatchRDate], Referencias.[BatchScOpr], Referencias.[BatchTrack], Referencias.[Route_To], Referencias.[Image_Seq], Referencias.[BatchPgDta], Referencias.[Form_Notes], Referencias.[CAJA_Nº], Referencias.[INDICE_1], Referencias.[DIA_HASTA_4], Referencias.[MES_HASTA_4], Referencias.[AÑO_HASTA_4], Referencias.[NUMERO_HASTA_4], Referencias.[DIA_DESDE_4], Referencias.[MES_DESDE_4], Referencias.[AÑO_DESDE_4], Referencias.[NUMERO_DESDE_4], Referencias.[DIA_DESDE_3], Referencias.[MES_DEDE_3], Referencias.[AÑO_DESDE_3], Referencias.[NUMERO_DESDE_3], Referencias.[DIA_HASTA_3], Referencias.[MES_HASTA_3], Referencias
            ', Referencias.[NUMERO_HASTA_3], Referencias.[MES_DESDE_2], Referencias.[AÑO_DESDE_2], Referencias.[NUMERO_DESDE_2], Referencias.[DIA_HASTA_2], Referencias.[MES_HASTA_2], Referencias.[AÑO_HASTA_2], Referencias.[NUMERO_HASTA_2], Referencias.[DIA_DESDE_1], Referencias.[MES_DEDE_1], Referencias.[AÑO_DESDE_1], Referencias.[NUMERO_DESDE_1], Referencias.[DIA_HASTA_1], Referencias.[MES_HASTA_1], Referencias.[AÑO_HASTA_1], Referencias.[NUMERO_HASTA_1], Referencias.[INDICE_2], Referencias.[INDICE_3], Referencias.[INDICE_4], Referencias.[PCX_DESCRIPCION_1], Referencias.[DESCRIPCION_1], Referencias.[PCX_DESCRIPCION_2], Referencias.[DESCRIPCION_2], Referencias.[PCX_DESCRIPCION_3], Referencias.[DESCRIPCION_3], Referencias.[PCX_DESCRIPCION_4], Referencias.[DESCRIPCION_4], Referencias.[IDEM_INDICE_1], Referencias.[IDEM_DETALLE_1], Referencias.[IDEM_INDICE_2], Referencias.[IDEM_INDICE_3], Referencias.[IDEM_DETALLE_3], Referencias.[IDEM_DETALLE_2], Referencias.[ENVIO_CAJAS], Referencias.[USUARIO], Referencias.[D
            'FROM Referencias;
            
            Sql = " Select * from Referencias "
            Sql = Sql & " WHERE (((Referencias.[Suspense_File]) Like '%" & InputBox("Ingrese el numero de lote") & "\%'));"
            
            RsPlanilla.Open Sql, comPlanilla
            R = 7
            Do While Not RsPlanilla.EOF
                For i = 1 To 4
                    If Not IsNull(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))) Then
                        Rem MsgBox RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))
                        FECHA_DESDE(i) = Format(Format(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("MES_DESDE_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("AÑO_HASTA_" & CStr(i)), "00"), "DD/MM/YYYY")
                        If Format(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)), "00") <> "00" Then
                            FECHA_HASTA(i) = Format(Format(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("MES_HASTA_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("AÑO_HASTA_" & CStr(i)), "00"), "DD/MM/YYYY")
                        Else
                            FECHA_HASTA(i) = FECHA_DESDE(i)
                        End If
                    End If
                    If Not IsNull(RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))) Then
                        MsgBox RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))
                        NUMERO_DESDE(i) = RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))
                        If Format(RsPlanilla.Fields.Item("NUMERO_HASTA_" & CStr(i)), "") <> "" Then
                            NUMERO_HASTA(i) = RsPlanilla.Fields.Item("NUMERO_HASTA_" & i)
                        Else
                             NUMERO_DESDE(i) = NUMERO_HASTA(i)
                        End If
                    End If
                    If Not IsNull(RsPlanilla.Fields.Item("INDICE_" & CStr(i))) Then
                        Indice(i) = RsPlanilla.Fields.Item("INDICE_" & CStr(i))
                        Else
                            If i <> 1 Then
                                If Not IsNull(RsPlanilla.Fields.Item("IDEM_INDICE_" & CStr(i))) Then
                                    Indice(i) = Indice(CStr(i - 1))
                                End If
                            Else
                                Indice(i) = 0
                            End If
                    End If
                
                    If Not IsNull(RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))) Then
                        Descripcion(i) = RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))
                    Else
                            If i <> 1 Then
                                If Not IsNull(RsPlanilla.Fields.Item("IDEM_DETALLE_" & CStr(i))) Then
                                    Descripcion(i) = Descripcion(CStr(i - 1))
                                End If
                            Else
                                Descripcion(i) = ""
                            End If
                    End If
                
                
                
                
                Next
                
                
                NombreArchivo = RsPlanilla.Fields.Item("CAJA_Nº").value & "_" & FechaHora & ".tif"
                        

              For i = 1 To 4
                
                    If Indice(i) <> "" Then
                        hojaEx.Cells(R, C_Caja) = RsPlanilla.Fields.Item("CAJA_Nº").value
                        hojaEx.Cells(R, C_Caja).Hyperlinks.Add hojaEx.Cells(R, C_Caja), ".\Cajas\" & Trim(RsPlanilla.Fields.Item("CAJA_Nº").value) & "\" & NombreArchivo
                        hojaEx.Cells(R, C_Indice) = Indice(i)
                        hojaEx.Cells(R, C_Descripcion) = Descripcion(i)
                        hojaEx.Cells(R, C_Fecha_desde) = FECHA_DESDE(i)
                        hojaEx.Cells(R, C_Fecha_hasta) = FECHA_HASTA(i)
                        hojaEx.Cells(R, C_N°_Desde) = NUMERO_DESDE(i)
                        hojaEx.Cells(R, C_N°_Hasta) = NUMERO_HASTA(i)
                        R = R + 1
                    End If
                Next
                
                If Dir(RsPlanilla.Fields.Item("Suspense_File").value) <> "" Then
                
                If Dir("\\SERVER-BASA\Sistemas\Referencias\cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value, vbDirectory) = "" Then
                    MkDir "\\SERVER-BASA\Sistemas\Referencias\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value
                    FileCopy RsPlanilla.Fields.Item("Suspense_File").value, "\\SERVER-BASA\Sistemas\Referencias" & "\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value & "\" & NombreArchivo
                 Else
                    FileCopy RsPlanilla.Fields.Item("Suspense_File").value, "\\SERVER-BASA\Sistemas\Referencias" & "\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value & "\" & NombreArchivo
                End If
                
                Else
                    MsgBox "No se encontro La imagen " & RsPlanilla.Fields.Item("Suspense_File").value
                End If
                
                 RsPlanilla.MoveNext
        Loop
                libroEx.SaveAs "\\SERVER-BASA\Sistemas\Referencias\" & InputBox("Ingrese el nombre de la planilla") & Format(Now, "ddmmyyy hhss") & ".xls"
                libroEx.Close
                ApExcel.Quit
                Set hojaEx = Nothing
                Set libroEx = Nothing
                Set ApExcel = Nothing
                
                MsgBox "Terminado"
            

End Sub

Private Sub Command11_Click()

Dim con As New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=C:\datas\DATAS.mdb"
Dim rs As New ADODB.Recordset
Dim DocNumero As Long
Dim Sql As String
Dim i As Integer
Dim C As Integer





Sql = " SELECT DCTO0755.NOMBRETIPODOCUMENTO, DCTO0755.NOMBRESUCURSAL, DCTO0755.DESCRIPCION, DCTO0755.DESDENUMERO, DCTO0755.HASTANUMERO ,DESDENUMERONUEVO,  DCTO0755.IDCAJA"
Sql = Sql & " From DCTO0755"
Sql = Sql & "  Where (((DCTO0755.NOMBRESUCURSAL) = 'ARCHIVO GENERAL'))"
Sql = Sql & "  ORDER BY DCTO0755.IDCAJA;"


Sql = "  SELECT DCT0755FINAL.IDDOCUMENTO, Descripcion, DCT0755FINAL.DESDENUMERONUEVO"
Sql = Sql & " From DCT0755FINAL"
Sql = Sql & " Where DCT0755FINAL.DESDENUMERONUEVO=0 And DCT0755FINAL.NOMBRETIPODOCUMENTO = 'EXPEDIENTES'"
Sql = Sql & " ORDER BY DCT0755FINAL.IDDOCUMENTO "


Sql = "  SELECT DCT0755FINAL.IDDOCUMENTO, DCT0755FINAL.DESCRIPCION, DCT0755FINAL.DESDENUMERO, DCT0755FINAL.HASTANUMERO, DCT0755FINAL.DESDENUMERONUEVO"
Sql = Sql & " From DCT0755FINAL"
Sql = Sql & " Where DCT0755FINAL.NOMBRETIPODOCUMENTO = 'EXPEDIENTES' "
Sql = Sql & " ORDER BY DCT0755FINAL.IDDOCUMENTO, DCT0755FINAL.HASTANUMERO DESC;"






rs.CursorLocation = adUseClient

Set rs = New ADODB.Recordset
rs.Open Sql, con, adOpenKeyset, adLockOptimistic
 
Do While Not rs.EOF

If rs!DESDENUMERO <> 0 Then

    rs!Descripcion = UCase(Trim(rs!Descripcion)) & " DESDE:" & Format(rs!DESDENUMERO, "000000") & " HASTA:" & Format(rs!HASTANUMERO, "000000")

End If
rs!DESDENUMERO = rs!DESDENUMERONUEVO
rs!HASTANUMERO = rs!DESDENUMERONUEVO


'  i = InStr(1, RS!Descripcion, "A - E")
''Rem MsgBox Mid(rs!Descripcion, i + 4)
''   C = InStr(i + 5, RS!Descripcion, " ")
''  If C = 0 Then
''
''  If IsNumeric(Mid(RS!Descripcion, i + 4)) Then
''
''    DocNumero = Mid(RS!Descripcion, i + 4)
''    Else
''    DocNumero = 0
''    End If
''
''
''  Else
''  Rem MsgBox (Mid(rs!Descripcion, i + 5, C - (i + 5)))
''   If IsNumeric(Mid(RS!Descripcion, i + 5)) Then
''    DocNumero = Mid(RS!Descripcion, i + 5)
''   Else
''    DocNumero = 0
''
''   End If
''
''    End If
''
''   RS!DESDENUMERONUEVO = DocNumero
   rs.Update
    rs.MoveNext
    
   Loop



End Sub

Private Sub Command12_Click()

    Dim Sql As String
    Dim rs As New ADODB.Recordset






    Sql = " SELECT     COD_ID_REFERENCIA, NRO_CAJA, FECHA_DESDE, FECHA_HASTA, DATEDIFF(day, FECHA_DESDE, FECHA_HASTA) AS Expr1, COD_CLIENTE"
   Sql = Sql & "  From REFERENCIAS"
   Sql = Sql & "   Where (DateDiff(Day, FECHA_DESDE, FECHA_HASTA) < 0)"
    Sql = Sql & "  ORDER BY DATEDIFF(day, FECHA_DESDE, FECHA_HASTA)"
    
    rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
Dim FECHA_DESDE As String
Dim FECHA_HASTA As String


Do While Not rs.EOF

    FECHA_DESDE = rs!FECHA_DESDE
    FECHA_HASTA = rs!FECHA_HASTA
    rs!FECHA_DESDE = FECHA_HASTA
    rs!FECHA_HASTA = FECHA_DESDE
    rs.Update
    rs.MoveNext
Loop
    
    
    
    
End Sub

Private Sub Command13_Click()
Dim rs As New ADODB.Recordset

Dim Sql As String
Dim ID As Long


Sql = "  SELECT LEGAJOS_CAPITAL.IDDOCUMENTO, LEGAJOS_CAPITAL.Campo4"
Sql = Sql & " , LEGAJOS_CAPITAL.Campo5, LEGAJOS_CAPITAL.Campo6,"
Sql = Sql & " LEGAJOS_CAPITAL.Campo7, LEGAJOS_CAPITAL.Campo8"
Sql = Sql & "  From LEGAJOS_CAPITAL"
Sql = Sql & "  ORDER BY LEGAJOS_CAPITAL.IDDOCUMENTO;"


rs.Open Sql, CONCUSTODIA
ID = 309544
Do While Not rs.EOF

If Not IsNull(rs!Campo5) Then
    ID = ID + 1
    INSERTARCUSTODIA ID, rs!IDDOCUMENTO, rs!Campo5
End If

If Not IsNull(rs!Campo6) Then
ID = ID + 1
    INSERTARCUSTODIA ID, rs!IDDOCUMENTO, rs!Campo6
End If

If Not IsNull(rs!Campo7) Then
ID = ID + 1
    INSERTARCUSTODIA ID, rs!IDDOCUMENTO, rs!Campo7
End If

If Not IsNull(rs!Campo8) Then
ID = ID + 1
    INSERTARCUSTODIA ID, rs!IDDOCUMENTO, rs!Campo8
End If

rs.MoveNext

Loop



End Sub

Private Sub Command14_Click()

Dim concu As New ADODB.Connection

concu.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=tpsescritura"


concu.Execute "DELETE * FROM DCTO0755 "

Dim rs As New ADODB.Recordset


rs.Open " select * from cliente ", concu


Do While Not rs.EOF

 MsgBox rs!Nombre

 rs.MoveNext
Loop





End Sub

Private Sub Command15_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String
Dim concus As New ADODB.Connection
concus.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Serverbasa1\SistemasBasa\ClientesBases\cambio.mdb;Persist Security Info=False"

'Dim Lectura As Long
'Dim EST As Long
'
'sql = " SELECT LECTURA_CAMBIO.NUMERO_LECTURA, LECTURA_CAMBIO.Expr1, LECTURA_COLECTOR_CUERPO.DESCRIPCION, LECTURA_COLECTOR_CUERPO.FECHA_CREACION, LECTURACOLECTOR.ID, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN"
'sql = sql & " FROM (LECTURA_CAMBIO INNER JOIN LECTURA_COLECTOR_CUERPO ON LECTURA_CAMBIO.NUMERO_LECTURA = LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA) INNER JOIN LECTURACOLECTOR ON LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA = LECTURACOLECTOR.NUMERO_LECTURA"
'sql = sql & " Where (((LECTURA_CAMBIO.NUMERO_LECTURA) > 30011) And ((LECTURA_CAMBIO.Expr1) = 1))"
'sql = sql & " ORDER BY LECTURA_CAMBIO.NUMERO_LECTURA, LECTURACOLECTOR.ID;"
'
'
'rs.Open sql, CONCUS
'
'Do While Not rs.EOF
'
'
'
'    If Lectura = rs!NUMERO_LECTURA Then
'        sql = " INSERT INTO  UNIFICADOS "
'        sql = sql & "(LECTURA, ESTANTERIA,  V, H, FECHA, ESTA, CAJA, CLIENTE, ORDEN)"
'        sql = sql & " VALUES (" & rs!NUMERO_LECTURA & "," & Mid(EST, 1, 4) & "," & Mid(EST, 5, 2) & "," & Mid(EST, 7, 2) & ",'" & rs!FECHA_CREACION & "'," & EST & "," & rs!Caja & "," & rs!Cliente & "," & rs!Orden & "  )"
'
'        CONCUS.Execute sql
'    Else
'
'        Lectura = rs!NUMERO_LECTURA
'        EST = rs!Caja
'    End If
'
'
'    rs.MoveNext
'Loop

Dim e As Integer
Dim h As Integer
Dim V As Integer
Dim rsID As ADODB.Recordset
Dim rsUnificado As ADODB.Recordset
Dim ConBa As New ADODB.Connection

ConBa.Open strConBasa

    For e = 2003 To 2029
         For V = 1 To 17
            For h = 1 To 2
            
            Sql = " SELECT ID From LECTURACOLECTOR "
            Sql = Sql & " Where Caja = " & e & Format(V, "00") & Format(h, "00")
            Sql = Sql & " ORDER BY LECTURACOLECTOR.ID DESC;"

            Set rsID = New ADODB.Recordset
                rsID.Open Sql, concus
                
                If Not rsID.EOF Then
                
                Sql = " SELECT TOP 40  LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA, CAJA, CLIENTE, ORDEN, ID, FECHA_CREACION "
                Sql = Sql & " FROM LECTURACOLECTOR INNER JOIN LECTURA_COLECTOR_CUERPO ON LECTURACOLECTOR.NUMERO_LECTURA = LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA"
                Sql = Sql & " Where ID > " & rsID!ID
                Sql = Sql & " ORDER BY ID"
                
                Set rsUnificado = New ADODB.Recordset
                    rsUnificado.Open Sql, concus
                    
                    Do While Not rsUnificado.EOF
        
                        If rsUnificado!Cliente = 9999 Then
                            Exit Do
                        End If
                        
        
                    Sql = " INSERT INTO UNIFICADOSCUSTODIA"
                    Sql = Sql & "   (ESTANTERIA, H, V, CLIENTE, CAJA, ORDEN, FECHA, LECTURA)"
                    Sql = Sql & " VALUES     ("
                    Sql = Sql & e & "," & h & "," & V & "," & rsUnificado!Cliente & "," & rsUnificado!Caja & "," & rsUnificado!Orden & ",'" & rsUnificado!FECHA_CREACION & "'," & rsUnificado!NUMERO_LECTURA & ")"
                    ConBa.Execute Sql
                    rsUnificado.MoveNext
                    Loop
                    
                    
                        
                
                
                
                
                End If
                
            
            
            Next
            
         
         
         Next
   Next

















End Sub

Private Sub Command16_Click()


Dim rs As New ADODB.Recordset
Dim Sql As String
Dim con As New ADODB.Connection
con.Open strConBasa
Set rs = New ADODB.Recordset




Sql = " SELECT     CAJA, ETIQUETA, NUMERO, LETRA, AÑO, DESCRIPCION"
Sql = Sql & " From MUNILUJAN"
Sql = Sql & " ORDER BY ETIQUETA"

rs.Open Sql, ConActiva, 0, 1


Do While Not rs.EOF
Sql = " Update LEGAJOS"
Sql = Sql & " SET  COD_INDICE ='001001', FK_INDICES =2121,COD_ESTADO = 2, LETRA_DESDE ='" & Trim(rs!lETRA) & "'"
Sql = Sql & "  , LETRA_HASTA ='" & Trim(rs!lETRA) & "', NRO_DESDE =" & rs!NUMERO & ", NRO_HASTA =" & rs!NUMERO
Sql = Sql & "  , FECHA_DESDE ='01/01/" & rs!Año & "', FECHA_HASTA ='31/12/" & rs!Año & "', NRO_CAJA =" & rs!Caja
Sql = Sql & "  ,COD_CLIENTE =128, FK_PERSONAL_CREACION =99, FECHA_CREACION ='24/02/2010'"
Sql = Sql & "  ,DESCRIPCION = '" & Trim(UCase(rs!Descripcion)) & "'"
Sql = Sql & " WHERE    ID_LEGAJO = " & rs!Etiqueta & " AND COD_CLIENTE IS NULL  "
con.Execute Sql
    rs.MoveNext
Loop


End Sub

Private Sub Command17_Click()



Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim Sql As String

Sql = " SELECT     ID_UNIFICADO, CLIENTE, CAJA, FECHAENTRADA,fecha ,  ERROR "
Sql = Sql & " From UNIFICADOSCUSTODIA "
Sql = Sql & " ORDER BY ID_UNIFICADO "


Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient

rs.Open Sql, ConActiva, 2, 3


Do While Not rs.EOF

Sql = " SELECT     FECHA"
Sql = Sql & " From MOVIMIENTOS_ELEMENTOS"
Sql = Sql & " WHERE     (COD_TIPO_ALMACENAMIENTO = 0) "
Sql = Sql & " AND (ANULADO IS NULL) AND (COD_OPERACION = 1) "
Sql = Sql & " AND (COD_TIPO = 1) "
Sql = Sql & " AND ELEMENTO = " & rs!Caja
If rs!Cliente <> 0 Then
    Sql = Sql & " AND COD_CLIENTE = " & rs!Cliente
End If
Sql = Sql & " ORDER BY FECHA DESC"

    Set rs2 = New ADODB.Recordset
    
    rs2.Open Sql, ConActiva, 0, 1
    If Not rs2.EOF Then
        If rs2!fecha > rs!fecha Then
        
        rs!ERROR = "CONSULTA"
        End If
        
    
    
    End If
    

    rs.Update
    rs.MoveNext
Loop






End Sub

Private Sub Command18_Click()

Dim rs As ADODB.Recordset
Dim Sql As String


Sql = " SELECT     ID_CONTENEDOR, DIGITO_VERIFICADOR, BARRA"
Sql = Sql & " From CONTENEDOR_CUSTODIA_CONCAJAS  "
Sql = Sql & " ORDER BY ID_CONTENEDOR"

Set rs = New ADODB.Recordset

rs.CursorLocation = adUseClient

rs.Open Sql, ConActiva, 2, 3
Do While Not rs.EOF
    rs!BARRA = Trim("P09" & Format(rs!ID_CONTENEDOR, "0000000") & Format(Digito_Verificador(rs!ID_CONTENEDOR), "00"))
    rs!Digito_Verificador = Digito_Verificador(rs!ID_CONTENEDOR)
    rs.Update
    rs.MoveNext
Loop








End Sub

Private Sub Command19_Click()


Dim rs As New ADODB.Recordset
Dim Sql As String
Dim CAS As Long
Dim i As Integer

Sql = " SELECT     MUNINOCONTROL.* From MUNINOCONTROL "

rs.Open Sql, ConActiva, 0, 1

        Do While Not rs.EOF
            For i = 0 To 10
            
            If Not IsNull(rs.Fields(i).value) Then
                    CAS = rs.Fields(i).value
                    Sql = " INSERT INTO REFERENCIAS"
                    Sql = Sql & " ( COD_CLIENTE, NRO_CAJA "
                    Sql = Sql & " , INDICE, DESCRIPCION "
                    Sql = Sql & " , FECHA_MODIFICACION, FECHA_CREACION"
                    Sql = Sql & " , USUARIO_MODIFICACION, FK_PERSONAL_CREACION"
                    Sql = Sql & " , FK_PERSONAL_MODIFICACION, BORRADO )"
                    Sql = Sql & " VALUES     "
                    Sql = Sql & " ( 128 ," & CAS
                    Sql = Sql & " , '001', 'REFERENCIA ADMINISTRADA POR EL CLIENTE 17/03/2010' "
                    Sql = Sql & " , '17/03/2010', '17/03/2010'"
                    Sql = Sql & " , 99,99 "
                    Sql = Sql & " , 99,0 )"
                    ExecutarSql Sql
                    End If
                    
            Next
            rs.MoveNext
        Loop




End Sub

Private Sub Command2_Click()

Dim Sql As String
Dim COn1 As New ADODB.Connection
Dim rs As New ADODB.Recordset

COn1.Open strConBasa
Sql = " SELECT     dbo.DOCUMENTOS_DIGITALES.NRO_DESDE, dbo.DOCUMENTOS_DIGITALES.NRO_HASTA, dbo.DOCUMENTOS_DIGITALES.ID"
Sql = Sql & " FROM dbo.DOCUMENTOS_DIGITALES INNER JOIN "
Sql = Sql & " dbo.DOCUMENTOS_DIGITALES_LOTE ON "
Sql = Sql & " dbo.DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = dbo.DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE "
Sql = Sql & " Where (dbo.DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES in(40,163, 172)) "
Sql = Sql & " And (dbo.DOCUMENTOS_DIGITALES.NRO_HASTA Is Null) and  not (dbo.DOCUMENTOS_DIGITALES.NRO_desde Is Null) "
Sql = Sql & " ORDER BY dbo.DOCUMENTOS_DIGITALES.ID DESC"


Sql = " SELECT     dbo.DOCUMENTOS_DIGITALES.NRO_DESDE, dbo.DOCUMENTOS_DIGITALES.NRO_HASTA, dbo.DOCUMENTOS_DIGITALES.ID,"
Sql = Sql & vbCrLf & "                      dbo.DOCUMENTOS_DIGITALES.LETRA_DESDE , dbo.DOCUMENTOS_DIGITALES.LETRA_HASTA"
Sql = Sql & vbCrLf & " FROM         dbo.DOCUMENTOS_DIGITALES INNER JOIN"
 Sql = Sql & vbCrLf & "                      dbo.DOCUMENTOS_DIGITALES_LOTE ON"
                      Sql = Sql & vbCrLf & " dbo.DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = dbo.DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
   Sql = Sql & vbCrLf & "  WHERE     (dbo.DOCUMENTOS_DIGITALES.NRO_HASTA IS NULL) AND (NOT (dbo.DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)) OR"
                         Sql = Sql & vbCrLf & "  (dbo.DOCUMENTOS_DIGITALES.LETRA_HASTA IS NULL) AND (NOT (dbo.DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL))"
   Sql = Sql & vbCrLf & "  ORDER BY dbo.DOCUMENTOS_DIGITALES.ID DESC"

rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, adOpenKeyset, adLockPessimistic

Do While Not rs.EOF
 If Not IsNull(rs!NRO_DESDE) Then
    If IsNull(rs!NRO_HASTA) Then
        rs!NRO_HASTA = rs!NRO_DESDE
    End If
 End If
 
If Not IsNull(rs!LETRA_DESDE) Then

    If IsNull(rs!LETRA_HASTA) Then
       rs!LETRA_HASTA = Trim(rs!LETRA_DESDE)
      
    End If
    
 
End If

rs.Update
    
    rs.MoveNext
Loop


End Sub

Private Sub Command20_Click()


 Dim C As String
 Dim e As String
 Dim Sql As String
 Dim RSCON As New ADODB.Recordset
 
 Dim rs As New ADODB.Recordset
 
C = " SELECT     COD_CLIENTE, NRO_CAJA, ID_CONTENEDOR_VIEJO, ID_CONTENEDOR_NUEVO, ESTADO_VIEJO"
C = C & "  From CAJAS_CORDOBA_COMPLETO"
C = C & "  ORDER BY COD_CLIENTE, NRO_CAJA "

rs.Open C, ConActiva, 0, 1


Do While Not rs.EOF
Sql = " Update CONTENEDOR "
Sql = Sql & " SET  ESTADO =1, COD_CLIENTE =NULL, NRO_CAJA =NULL, F_MODIFICACION =NULL, IDREQUERIMIENTO =NULL, NRO_REMITO =NULL"
Sql = Sql & " Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR_VIEJO
ExecutarSql Sql
   
Sql = " Update CONTENEDOR "
Sql = Sql & " SET  ESTADO =" & rs!ESTADO_VIEJO & ", COD_CLIENTE =" & rs!COD_CLIENTE & ", NRO_CAJA =" & rs!NRO_CAJA
Sql = Sql & " Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR_NUEVO
ExecutarSql Sql
   
   
   
    
    rs.MoveNext
Loop



End Sub

Private Sub Command21_Click()
'INSERT INTO REMITOS_CUERPO
'                      (NRO_REMITO, NRO_REM_PROV, TIPO, OPERACION, ESTADO, FECHA, ID_CLIENTE, OBSERVACIONES, CANTIDAD, AUDIT_USUARIO, AUDIT_FECHA,
'                      ANULADO, FECHA_INGRESO, FECHA_ERROR, COD_TIPO_ALMACENAMIENTO, COD_PERSONAL_ENTREGA, COD_PERSONAL_RECIBE,
'                      DESC_CONTROL_REF, COD_FLETE, CONTROL_REFERENCIA, ID_SQL, COD_USUARIO_CLIENTE, FACTURA, IMAGEN, RE_PROV_PURO, BONIFICADA)
'VALUES     (,,,,,,,,,,,,,,,,,,,,,,,,,)



End Sub

Private Sub Command22_Click()
Dim a As String

Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

a = " SELECT     CAJAS, CLIENTE, CONTROL"
a = a & " From CAJASLACAJA2 ORDER BY CAJAS"

rs.CursorLocation = adUseClient
rs.Open a, ConActiva, 2, 3

Do While Not rs.EOF
    Dim s As String
      s = ""
     a = "  SELECT     FK_CLIENTES, FK_CAJAS"
    a = a & "  From DOCUMENTOS_DIGITALES_LOTE"
    a = a & "  Where (FK_CLIENTES = 163)"
    a = a & "  and  FK_CAJAS = " & rs!CAJAS
    
    Set rs2 = New ADODB.Recordset
    rs2.Open a, ConActiva, 0, 1
    If Not rs2.EOF Then
        s = "digital"
       
    End If
    
    
    a = " SELECT     COD_CLIENTE, NRO_CAJA"
    a = a & " From REFERENCIAS"
    a = a & " Where (COD_CLIENTE = 163)"
    a = a & " And NRO_CAJA = " & rs!CAJAS
    Set rs2 = New ADODB.Recordset
    rs2.Open a, ConActiva, 0, 1
    If Not rs2.EOF Then
        s = Trim(s) & "-Referencia"
        
    End If
    
    a = " SELECT     NRO_CAJA, COD_CLIENTE"
    a = a & " From LEGAJOS"
    a = a & " Where (COD_CLIENTE = 163)"
    a = a & " And NRO_CAJA = " & rs!CAJAS
    Set rs2 = New ADODB.Recordset
    rs2.Open a, ConActiva, 0, 1
    If Not rs2.EOF Then
        s = Trim(s) & "-legajos"
    End If
    
    rs!Control = s
        rs.Update
  s = ""
    
    rs.MoveNext
Loop



End Sub

Private Sub Command23_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rsRef As New ADODB.Recordset
Dim Cliente As Integer
Dim Caja As Long



Sql = " SELECT     ESTADO, COD_CLIENTE, NRO_CAJA,"
Sql = Sql & " Into CONTROL_REFERENCIAS"
Sql = Sql & " From CONTENEDOR"
Sql = Sql & " WHERE     (ESTADO IN (2, 3)) "
Sql = Sql & " AND COD_CLIENTE = " & InputBox("INGRESE EL CLIENTE")


ExecutarSql "DELETE  CONTROL_REFERENCIAS "
ExecutarSql Sql


rs.CursorLocation = adUseClient
rs.Open Sql, ConActiva, 3, 2


Do While Not rs.EOF

    Cliente = rs!FK_CLIENTE
    Caja = rs!FK_CAJA

    rs!REFERENCIAS = ""
    Sql = " SELECT     NRO_CAJA, COD_CLIENTE "
    Sql = Sql & " From REFERENCIAS "
    Sql = Sql & " WHERE  NRO_CAJA = " & Caja
    Sql = Sql & " AND COD_CLIENTE = " & Cliente
    
    Set rsRef = New ADODB.Recordset
    rsRef.Open Sql, ConActiva, 0, 1
    If Not rsRef.EOF Then
    
         rs!REFERENCIAS = "rango"
    
    
    End If
    

Sql = "  SELECT     NRO_CAJA, COD_CLIENTE"
Sql = Sql & " From LEGAJOS"
Sql = Sql & " Where NRO_CAJA =" & Caja
Sql = Sql & " And COD_CLIENTE =" & Cliente
    
    Set rsRef = New ADODB.Recordset
    rsRef.Open Sql, ConActiva, 0, 1
    If Not rsRef.EOF Then
    
         rs!REFERENCIAS = Trim(Trim(rs!REFERENCIAS) & " Legajo")
    
    
    End If
    
    
  Sql = " SELECT     FK_CLIENTES, FK_CAJAS"
Sql = Sql & "  From DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  Where FK_CLIENTES =" & Cliente
Sql = Sql & " AND FK_CAJAS = " & Caja
Sql = Sql & " ORDER BY FK_CAJAS DESC"
    Set rsRef = New ADODB.Recordset
    rsRef.Open Sql, ConActiva, 0, 1
    If Not rsRef.EOF Then
         rs!REFERENCIAS = Trim(Trim(rs!REFERENCIAS) & " IMAGEN")
    End If
    
    
    Sql = "  SELECT     COD_CLIENTE, COD_NRO_CAJA"
    Sql = Sql & " From ORDENAR_DOCUMENTACION_DETALLE"
    Sql = Sql & " Where COD_CLIENTE = " & Cliente
    Sql = Sql & " And Cod_Nro_Caja = " & Caja
    
    Set rsRef = New ADODB.Recordset
    rsRef.Open Sql, ConActiva, 0, 1
    If Not rsRef.EOF Then
         rs!REFERENCIAS = Trim(Trim(rs!REFERENCIAS) & " DOCUMENTO")
    End If
    
    
    
    rs.Update
    
    

    rs.MoveNext
Loop


End Sub

Private Sub Command24_Click()
'SELECT     REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.ID_CLIENTE, TIPOREQUERIMIENTO.DESCRIPCION,
'                      REQUERIMIENTO.COD_USUARIO_CLIENTE, CLIENTEUSUARIO.APELLIDO_NOMBRE, REQUERIMIENTO.CANTIDAD_IMAGENES,
'                      REQUERIMIENTO.CANTIDAD, REQUERIMIENTO.IDREMITO, REMITOS_CUERPO.OBSERVACIONES, INDICES_1.DESCRIPCION AS PROVINCIA,
'                      INDICES_2.DESCRIPCION AS SUCURSAL, REQUERIMIENTO.FECHARECEPCION, REQUERIMIENTO.IDTIPOREQUERIMIENTO,
'                      REQUERIMIENTO_ESTADO.DESCRIPCION AS Expr1, REQUERIMIENTO.ANULADO
'FROM         TIPOREQUERIMIENTO RIGHT OUTER JOIN
'                      REMITOS_CUERPO RIGHT OUTER JOIN
'                      REQUERIMIENTO LEFT OUTER JOIN
'                      REQUERIMIENTO_ESTADO ON REQUERIMIENTO.IDESTADO = REQUERIMIENTO_ESTADO.ID_ESTADO ON
'                      REMITOS_CUERPO.NRO_REMITO = REQUERIMIENTO.IDREMITO ON
'                      TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO = REQUERIMIENTO.IDTIPOREQUERIMIENTO LEFT OUTER JOIN
'                      INDICES INDICES_2 RIGHT OUTER JOIN
'                      INDICES INDICES_1 RIGHT OUTER JOIN
'                      CLIENTEUSUARIO ON INDICES_1.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND
'                      INDICES_1.INDICE = SUBSTRING(CLIENTEUSUARIO.COD_INDICE, 1, 3) ON INDICES_2.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND
'                      INDICES_2.INDICE = CLIENTEUSUARIO.COD_INDICE ON REQUERIMIENTO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO
'WHERE     (REQUERIMIENTO.ID_CLIENTE = 231) AND (REQUERIMIENTO.FECHARECEPCION > CONVERT(DATETIME, '2010-04-24 00:00:00', 102))
'ORDER BY INDICES_1.DESCRIPCION, INDICES_2.DESCRIPCION, REQUERIMIENTO.IDTIPOREQUERIMIENTO
'
'
'
'
'FACTURACION SUPERVIELLE






End Sub

Private Sub Command25_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rslegajos As New ADODB.Recordset
Dim rsRequerimiento As New ADODB.Recordset
Dim TIPO As String
Dim flete As String
Dim FECHA_DESDE As String
Dim IDREMITO As String

FECHA_DESDE = "24/04/2010"



Sql = " SELECT NRO_REMITO, REMITOS_CUERPO.NRO_REM_PROV, TIPO_ALMACENAMIENTO.DESCRIPCION AS ELEMENTO,"
Sql = Sql & vbCrLf & " TIPO_REMITO.DESCRIPCION AS TIPO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.OBSERVACIONES, REMITOS_CUERPO.CANTIDAD,"
Sql = Sql & vbCrLf & " CLIENTEUSUARIO.APELLIDO_NOMBRE, INDICES_1.DESCRIPCION AS PROVINCIA, INDICES_2.DESCRIPCION AS SUCURSAL,"
Sql = Sql & vbCrLf & " REMITOS_CUERPO.COBRAR_FLETE"
Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO LEFT OUTER JOIN"
Sql = Sql & vbCrLf & " TIPO_ALMACENAMIENTO ON REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = TIPO_ALMACENAMIENTO.ID RIGHT OUTER JOIN"
Sql = Sql & vbCrLf & " TIPO_REMITO ON REMITOS_CUERPO.TIPO = TIPO_REMITO.ID LEFT OUTER JOIN"
Sql = Sql & vbCrLf & " INDICES INDICES_1 INNER JOIN"
Sql = Sql & vbCrLf & " INDICES INDICES_2 INNER JOIN"
Sql = Sql & vbCrLf & " CLIENTEUSUARIO ON INDICES_2.INDICE = CLIENTEUSUARIO.COD_INDICE AND INDICES_2.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE ON"
Sql = Sql & vbCrLf & " INDICES_1.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND INDICES_1.INDICE = SUBSTRING(CLIENTEUSUARIO.COD_INDICE, 1, 3) ON"
Sql = Sql & vbCrLf & " REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
Sql = Sql & vbCrLf & " WHERE     (REMITOS_CUERPO.ID_CLIENTE = 231) "
Sql = Sql & vbCrLf & " AND (REMITOS_CUERPO.FECHA > CONVERT(DATETIME, '2010-04-24 00:00:00', 102))"
Sql = Sql & vbCrLf & " AND (NOT (REMITOS_CUERPO.TIPO IN (2))) AND (REMITOS_CUERPO.OPERACION = 0)"
                      
 rs.Open Sql, ConActiva, 0, 1
                      
                      ExecutarSql "DELETE FROM TEM_SUPERVIELLE"
                      
    Do While Not rs.EOF
        TIPO = ""
        If Trim(rs!TIPO) = "CONSULTA" Then
            TIPO = "DEVOLU DE CONSULTA " & UCase(rs!Elemento)
        Else
            TIPO = rs!TIPO & " de " & rs!Elemento
        End If
        
        flete = ""
        If IsNull(rs!COBRAR_FLETE) Then
            flete = ""
        Else
            flete = rs!COBRAR_FLETE
        End If
        INSERTAR_FACTURA_SUPER "REMITO", 0, rs!NRO_REMITO, rs!NRO_REM_PROV, TIPO, rs!fecha, rs!OBSERVACIONES, rs!cantidad, 0, rs!APELLIDO_NOMBRE _
        , rs!PROVINCIA, rs!Sucursal, "No tiene", flete, " ", " "
        rs.MoveNext
    Loop
    
    ConBasa.CommandTimeout = 180
    
    
 Sql = " SELECT  INDICES.DESCRIPCION, INDICES.INDICE, COUNT(DISTINCT LEGAJOS.NRO_CAJA) AS CANTIDAD_CAJAS, COUNT(*) AS CANTIDAD_LEGAJOS,"
 Sql = Sql & vbCrLf & " INDICES_1.DESCRIPCION AS PROVINCIA, INDICES_2.DESCRIPCION AS SUCURSAL "
 Sql = Sql & vbCrLf & " FROM LEGAJOS INNER JOIN "
 Sql = Sql & vbCrLf & " INDICES ON LEGAJOS.COD_CLIENTE = INDICES.COD_CLIENTE AND LEGAJOS.COD_INDICE = INDICES.INDICE LEFT OUTER JOIN"
 Sql = Sql & vbCrLf & " INDICES INDICES_2 ON SUBSTRING(LEGAJOS.COD_INDICE, 1, 6) = INDICES_2.INDICE AND"
 Sql = Sql & vbCrLf & " LEGAJOS.COD_CLIENTE = INDICES_2.COD_CLIENTE LEFT OUTER JOIN"
 Sql = Sql & vbCrLf & " INDICES INDICES_1 ON LEGAJOS.COD_CLIENTE = INDICES_1.COD_CLIENTE AND SUBSTRING(LEGAJOS.COD_INDICE, 1, 3) = INDICES_1.INDICE "
 Sql = Sql & vbCrLf & " WHERE   LEGAJOS.COD_CLIENTE = 231 "
 Sql = Sql & vbCrLf & " AND LEGAJOS.FECHA_CREACION >  '" & FECHA_DESDE & "'"
 Sql = Sql & vbCrLf & " GROUP BY INDICES.DESCRIPCION, INDICES.INDICE, INDICES_1.DESCRIPCION, INDICES_2.DESCRIPCION"
 Sql = Sql & vbCrLf & " HAVING      (INDICES.DESCRIPCION LIKE 'CARGA %')"
    
 rslegajos.Open Sql, ConActiva, 0, 1
 
 Do While Not rslegajos.EOF
    
INSERTAR_FACTURA_SUPER "CARGA LEGAJOS", 0, 0, 0, "CARGA DE LEGAJOS", FECHA_DESDE, rslegajos!Descripcion, rslegajos!Cantidad_Cajas, rslegajos!CANTIDAD_LEGAJOS, "", _
rslegajos!PROVINCIA, rslegajos!Sucursal, "", "", "", ""
    
    
    
    rslegajos.MoveNext
    
 Loop
 
 
                      
                      
                      
                      
    Sql = "  SELECT     REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.IDREMITO, REQUERIMIENTO.FECHARECEPCION,"
        Sql = Sql & "      REQUERIMIENTO.DESCRIPCION AS DESCRIPCION_REQUE, INDICES_1.DESCRIPCION AS PROVINCIA, INDICES_2.DESCRIPCION AS SUCURSAL,"
        Sql = Sql & "      TIPOREQUERIMIENTO.DESCRIPCION AS TIPO, CLIENTEUSUARIO.APELLIDO_NOMBRE, REQUERIMIENTO.CANTIDAD_IMAGENES,"
        Sql = Sql & "      REQUERIMIENTO.CANTIDAD, REMITOS_CUERPO.OBSERVACIONES AS REMITO_DESCRIPCION,"
        Sql = Sql & "      REQUERIMIENTO_ESTADO.DESCRIPCION AS ESTADO, REQUERIMIENTO.HORA_ARCHIVISTA, REQUERIMIENTO.FLETE,REQUERIMIENTO.COBRAR"
        Sql = Sql & "  FROM         TIPOREQUERIMIENTO RIGHT OUTER JOIN"
        Sql = Sql & " REMITOS_CUERPO RIGHT OUTER JOIN"
        Sql = Sql & " REQUERIMIENTO LEFT OUTER JOIN"
        Sql = Sql & " REQUERIMIENTO_ESTADO ON REQUERIMIENTO.IDESTADO = REQUERIMIENTO_ESTADO.ID_ESTADO ON"
        Sql = Sql & " REMITOS_CUERPO.NRO_REMITO = REQUERIMIENTO.IDREMITO ON"
        Sql = Sql & " TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO = REQUERIMIENTO.IDTIPOREQUERIMIENTO LEFT OUTER JOIN"
        Sql = Sql & " INDICES INDICES_2 RIGHT OUTER JOIN"
        Sql = Sql & " INDICES INDICES_1 RIGHT OUTER JOIN"
        Sql = Sql & " CLIENTEUSUARIO ON INDICES_1.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND"
        Sql = Sql & " INDICES_1.INDICE = SUBSTRING(CLIENTEUSUARIO.COD_INDICE, 1, 3) ON INDICES_2.COD_CLIENTE = CLIENTEUSUARIO.COD_CLIENTE AND"
        Sql = Sql & " INDICES_2.INDICE = CLIENTEUSUARIO.COD_INDICE ON REQUERIMIENTO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
        Sql = Sql & " WHERE   (REQUERIMIENTO.ID_CLIENTE = 231) "
        Sql = Sql & " AND REQUERIMIENTO.FECHARECEPCION > '" & FECHA_DESDE & "'"
        Sql = Sql & " AND (REQUERIMIENTO.ANULADO IS NULL)"
        Sql = Sql & " ORDER BY INDICES_1.DESCRIPCION, INDICES_2.DESCRIPCION, REQUERIMIENTO.IDTIPOREQUERIMIENTO"
                      
     rsRequerimiento.Open Sql, ConActiva, 0, 1
                      
            Dim HORA_ARCHIVISTA            As String
            Dim COBRAR As String
            Dim DESCRIPCION_REQUE  As String
            Dim Cantidad_Imagenes As Long
          Do While Not rsRequerimiento.EOF
          IDREMITO = ""
          
          If IsNull(rsRequerimiento!IDREMITO) Then
            IDREMITO = 0
          Else
            IDREMITO = rsRequerimiento!IDREMITO
          End If
          
          If IsNull(rsRequerimiento!flete) Then
            flete = ""
          Else
            flete = rsRequerimiento!flete
          End If
                      
          If IsNull(rsRequerimiento!HORA_ARCHIVISTA) Then
            HORA_ARCHIVISTA = ""
          Else
            HORA_ARCHIVISTA = rsRequerimiento!HORA_ARCHIVISTA
          End If
          
          If IsNull(rsRequerimiento!COBRAR) Then
            COBRAR = " "
          Else
            COBRAR = rsRequerimiento!COBRAR
          End If
          
          If IsNull(rsRequerimiento!DESCRIPCION_REQUE) Then
                DESCRIPCION_REQUE = ""
          Else
                DESCRIPCION_REQUE = Mid(rsRequerimiento!DESCRIPCION_REQUE, 1, 40)
          End If
          
          
          If IsNull(rsRequerimiento!Cantidad_Imagenes) Then
            Cantidad_Imagenes = 0
          Else
          Cantidad_Imagenes = rsRequerimiento!Cantidad_Imagenes
          End If
          
            
            
            INSERTAR_FACTURA_SUPER "REQUER", rsRequerimiento!IDREQUERIMIENTO, CLng(IDREMITO), "" _
            , UCase(rsRequerimiento!TIPO), Format(rsRequerimiento!FECHARECEPCION, "DD/MM/YYYY"), DESCRIPCION_REQUE, rsRequerimiento!cantidad _
            , Cantidad_Imagenes, rsRequerimiento!APELLIDO_NOMBRE, rsRequerimiento!PROVINCIA, rsRequerimiento!Sucursal _
            , rsRequerimiento!estado, flete, HORA_ARCHIVISTA, COBRAR
           
          
            rsRequerimiento.MoveNext
          Loop
          
                      
                      
                      
                      
                      
                      
End Sub

Private Sub Command26_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = " SELECT     CONTENEDOR.ID_CONTENEDOR "
Sql = Sql & " FROM CONTENEDOR INNER JOIN "
Sql = Sql & " ESTAMTERIAMAL1 ON CONTENEDOR.ESTANTERIA = ESTAMTERIAMAL1.ESTANTERIA AND"
Sql = Sql & " CONTENEDOR.HORIZONTAL = ESTAMTERIAMAL1.HORIZONTAL AND CONTENEDOR.VERTICAL = ESTAMTERIAMAL1.VERTICAL AND"
Sql = Sql & " CONTENEDOR.Adelante_Atras = ESTAMTERIAMAL1.Adelante_Atras"
Sql = Sql & " Where (CONTENEDOR.COD_CLIENTE Is Null)"

Set rs = New ADODB.Recordset
rs.Open Sql, ConActiva, 0, 1
Do While Not rs.EOF
    
    Sql = " DELETE FROM CONTENEDOR Where ID_CONTENEDOR =  " & rs!ID_CONTENEDOR
    ExecutarSql Sql
    rs.MoveNext

Loop



End Sub

Private Sub Command27_Click()


Dim Sql As String

Dim Sql2 As String
Dim rs As New ADODB.Recordset

Sql = " SELECT REMITOS_CUERPO.NRO_REMITO,  REMITOS_CUERPO.TIPO, REMITOS_CUERPO.OPERACION, REMITOS_CUERPO.ESTADO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.ID_CLIENTE,"
Sql = Sql & " REMITOS_CUERPO.ANULADO, CLIENTEUSUARIO.APELLIDO_NOMBRE, INDICES.DESCRIPCION, REMITOS_DETALLE.DESDE,"
Sql = Sql & " REMITOS_DETALLE.Hasta , CLIENTEUSUARIO.ID_CLIENTEUSUARIO, INDICES.ID"
Sql = Sql & " FROM         REMITOS_CUERPO INNER JOIN"
Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO INNER JOIN"
Sql = Sql & " CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO INNER JOIN"
Sql = Sql & " INDICES ON CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE AND CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE"
Sql = Sql & " Where (REMITOS_CUERPO.TIPO = 0) "
Sql = Sql & " And (REMITOS_CUERPO.id_cliente = 231) "
Sql = Sql & " And (REMITOS_CUERPO.ANULADO Is Null)"
Sql = Sql & " ORDER BY REMITOS_DETALLE.DESDE"



Sql = " SELECT     CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_INDICE, SUEPERVIELLE_GYC_2.Expr1, INDICES.DESCRIPCION, INDICES.ID"
Sql = Sql & " FROM  CAJAS INNER JOIN"
Sql = Sql & " SUEPERVIELLE_GYC_2 ON CAJAS.FK_CLIENTE = SUEPERVIELLE_GYC_2.COD_CLIENTE AND"
Sql = Sql & " CAJAS.NRO_CAJA = SUEPERVIELLE_GYC_2.NRO_CAJA INNER JOIN"
Sql = Sql & " INDICES ON SUEPERVIELLE_GYC_2.Expr1 = INDICES.INDICE AND SUEPERVIELLE_GYC_2.COD_CLIENTE = INDICES.COD_CLIENTE"
Sql = Sql & "  Where (Cajas.FK_Indice Is Null) ORDER BY CAJAS.NRO_CAJA"

Sql = " SELECT     INDICES.Id ,CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_ESTADO, CAJAS.FK_INDICE, REFERENCIAS.INDICE, INDICES.DESCRIPCION"
Sql = Sql & " FROM         CAJAS INNER JOIN"
Sql = Sql & "                      REFERENCIAS ON CAJAS.FK_CLIENTE = REFERENCIAS.COD_CLIENTE AND CAJAS.NRO_CAJA = REFERENCIAS.NRO_CAJA INNER JOIN"
Sql = Sql & "                      INDICES ON REFERENCIAS.INDICE = INDICES.INDICE AND REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE"
Sql = Sql & " Where (Cajas.FK_CLIENTE = 231) And (Cajas.FK_Indice Is Null)"
Sql = Sql & " ORDER BY REFERENCIAS.INDICE DESC "


Sql = " SELECT     CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_ESTADO, CAJAS.FK_INDICE, INDICES.DESCRIPCION, INDICES.ID"
Sql = Sql & "  FROM         CAJAS INNER JOIN"
Sql = Sql & "                       LEGAJOS ON CAJAS.NRO_CAJA = LEGAJOS.NRO_CAJA AND CAJAS.FK_CLIENTE = LEGAJOS.COD_CLIENTE INNER JOIN"
              Sql = Sql & "         INDICES ON SUBSTRING(LEGAJOS.COD_INDICE, 1, 6) = INDICES.INDICE AND LEGAJOS.COD_CLIENTE = INDICES.COD_CLIENTE"
Sql = Sql & "  GROUP BY CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_ESTADO, CAJAS.FK_INDICE, INDICES.DESCRIPCION, INDICES.ID"
Sql = Sql & "  Having (Cajas.FK_CLIENTE = 231) And (Cajas.FK_Indice Is Null)"



Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.TIPO, REMITOS_CUERPO.OPERACION, REMITOS_CUERPO.ESTADO,"
Sql = Sql & "                       REMITOS_CUERPO.FECHA, REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.ANULADO, CLIENTEUSUARIO.APELLIDO_NOMBRE,"
Sql = Sql & "                       INDICES.DESCRIPCION, REMITOS_DETALLE.DESDE, REMITOS_DETALLE.HASTA, CLIENTEUSUARIO.ID_CLIENTEUSUARIO, INDICES.ID,"
              Sql = Sql & "          Cajas.FK_CLIENTE , Cajas.FK_Indice"
Sql = Sql & "  FROM         REMITOS_CUERPO INNER JOIN"
               Sql = Sql & "        REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO INNER JOIN"
             Sql = Sql & "          CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO INNER JOIN"
            Sql = Sql & "           INDICES ON CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE AND CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE INNER JOIN"
            Sql = Sql & "            CAJAS ON REMITOS_DETALLE.DESDE = CAJAS.NRO_CAJA"
Sql = Sql & "  WHERE     (REMITOS_CUERPO.TIPO = 2) AND (REMITOS_CUERPO.ID_CLIENTE = 231) AND (REMITOS_CUERPO.ANULADO IS NULL) AND"
Sql = Sql & "                        (CAJAS.FK_CLIENTE = 231) AND (CAJAS.FK_INDICE IS NULL)"
Sql = Sql & "  ORDER BY REMITOS_DETALLE.DESDE"

rs.CursorLocation = adUseClient

rs.Open Sql, ConActiva, adOpenForwardOnly, adLockReadOnly


Do While Not rs.EOF
    Sql2 = "  Update Cajas "
    Sql2 = Sql2 & " SET FK_INDICE =" & rs!ID
    Sql2 = Sql2 & ", FK_CLIENTES_USUARIO =" & rs!ID_CLIENTEUSUARIO
'    Sql2 = Sql2 & " , FK_REMITO_CUSTODIA =" & rs!NRO_REMITO
    Sql2 = Sql2 & " Where (FK_CLIENTE = 231) "
    Sql2 = Sql2 & " And NRO_CAJA = " & rs!Desde
        
    ExecutarSql Sql2
        
    rs.MoveNext
    
    
Loop






End Sub

Private Sub Command28_Click()


        Dim Sql As String
        Dim rs As New ADODB.Recordset
        
        Dim con As New ADODB.Connection
        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=C:\Cambio de posicion\CAMBIO.mdb"


            Sql = " SELECT CONTENEDOR_CUSTODIA_CONCAJAS.ID_CONTENEDOR, CONTENEDOR_CUSTODIA_CONCAJAS.ESTANTERIA, CONTENEDOR_CUSTODIA_CONCAJAS.CLIENTE, CONTENEDOR_CUSTODIA_CONCAJAS.CAJA, CONTENEDOR_CUSTODIA_CONCAJAS.ORDEN, CONTENEDOR_CUSTODIA_CONCAJAS.VERTICAL, CONTENEDOR_CUSTODIA_CONCAJAS.HORIZONTAL, CONTENEDOR_CUSTODIA_CONCAJAS.MODULO_V, CONTENEDOR_CUSTODIA_CONCAJAS.MODULO_H, CONTENEDOR_CUSTODIA_CONCAJAS.ERROR, CONTENEDOR_CUSTODIA_CONCAJAS.ID_UNIFICADO"
            Sql = Sql & " From CONTENEDOR_CUSTODIA_CONCAJAS"
            Sql = Sql & " Where (((CONTENEDOR_CUSTODIA_CONCAJAS.Cliente) > 0))"
            Sql = Sql & " ORDER BY CONTENEDOR_CUSTODIA_CONCAJAS.ID_CONTENEDOR;"
       
      
      rs.Open Sql, con
      Dim estado As Integer
      
      Do While Not rs.EOF
      
            estado = Esatodo(rs!Caja, rs!Cliente)
            Sql = "  Update CONTENEDOR"
            Sql = Sql & "  SET  ESTADO =1, COD_CLIENTE =null, NRO_CAJA =null"
            Sql = Sql & "  Where COD_CLIENTE = " & rs!Cliente
            Sql = Sql & "  And NRO_CAJA = " & rs!Caja
            ExecutarSql Sql
            
            
            Sql = " Update CONTENEDOR"
            Sql = Sql & " SET      ESTADO = " & estado
            Sql = Sql & "  , COD_CLIENTE = " & rs!Cliente
            Sql = Sql & "  , NRO_CAJA = " & rs!Caja
            Sql = Sql & "  Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR
            
            ExecutarSql Sql
            
            
            rs.MoveNext
      Loop
      
      


End Sub

Private Sub Command29_Click()
Dim Sql As String
Dim rs As ADODB.Recordset

Sql = " SELECT     CAMBPOSI.ESTANTERIA, CAMBPOSI.VERTICAL, CAMBPOSI.HORIZONTAL, CAMBPOSI.CLIENTE, CAMBPOSI.CAJA, CONTENEDOR.ID_CONTENEDOR, "
Sql = Sql & " CONTENEDOR.estado"
Sql = Sql & " FROM CONTENEDOR INNER JOIN"
Sql = Sql & " CAMBPOSI ON CONTENEDOR.ESTANTERIA = CAMBPOSI.ESTANTERIA AND CONTENEDOR.VERTICAL = CAMBPOSI.VERTICAL AND"
Sql = Sql & " CONTENEDOR.Horizontal = CAMBPOSI.Horizontal"
Sql = Sql & " ORDER BY CAMBPOSI.CLIENTE, CAMBPOSI.CAJA"
Set rs = New ADODB.Recordset

rs.Open Sql, ConActiva, 0, 1
Do While Not rs.EOF
        Sql = "Update CONTENEDOR"
        Sql = Sql & "  SET COD_CLIENTE = NULL, NRO_CAJA = NULL, ESTADO = 1"
        Sql = Sql & "  Where COD_CLIENTE = " & rs!Cliente
        Sql = Sql & "  And NRO_CAJA = " & rs!Caja
        ExecutarSql Sql
        
        Sql = " Update CONTENEDOR"
        Sql = Sql & "  SET COD_CLIENTE = " & rs!Cliente
        Sql = Sql & " , NRO_CAJA =" & rs!Caja
        Sql = Sql & " , ESTADO = 2 "
        Sql = Sql & "  Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR
        ExecutarSql Sql
        
        
        
    rs.MoveNext
    
    
Loop

End Sub

Private Sub Command3_Click()
Dim Sql As String
Dim Sqlc As String
Dim rsContenedor As New ADODB.Recordset
Dim rsCajas As New ADODB.Recordset
Dim Legajo As String
Dim RSESTANTERIA  As New ADODB.Recordset



Sqlc = " SELECT ID_CONTENEDOR,COD_CLIENTE, NRO_CAJA, FK_CAJAS "
Sqlc = Sqlc & " From CONTENEDOR "
Sqlc = Sqlc & " Where (FK_CAJAS Is Null) "
Sqlc = Sqlc & " And (Not (COD_CLIENTE Is Null))"
Sqlc = Sqlc & " ORDER BY COD_CLIENTE, NRO_CAJA "


Sql = "  SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_CONTENEDOR"
Sql = Sql & "  From Cajas "
Sql = Sql & "  Where (FK_CLIENTE Is Null)"
Sql = Sql & "  ORDER BY ID_CAJA "

rsContenedor.CursorLocation = adUseClient

rsContenedor.Open Sqlc, ConActiva, adOpenKeyset, adLockPessimistic
rsCajas.Open Sql, ConActiva, adOpenKeyset, adLockPessimistic

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


'
'Dim Sql As String
'Dim Sqlc As String
'Dim rsContenedor As New ADODB.Recordset
'Dim rsCajas As New ADODB.Recordset
'Dim Legajo As String
'Dim RSESTANTERIA  As New ADODB.Recordset
'
'
'Sql = " SELECT dbo.CONTENEDOR.ID_CONTENEDOR, dbo.CONTENEDOR.COD_CLIENTE, dbo.CONTENEDOR.NRO_CAJA, dbo.CONTENEDOR.FK_CAJAS,"
'Sql = Sql & " dbo.Cajas.ID_CAJA , dbo.Cajas.FK_CONTENEDOR"
'Sql = Sql & " FROM dbo.CONTENEDOR INNER JOIN"
'Sql = Sql & " dbo.CAJAS ON dbo.CONTENEDOR.COD_CLIENTE = dbo.CAJAS.FK_CLIENTE AND dbo.CONTENEDOR.NRO_CAJA = dbo.CAJAS.NRO_CAJA"
'Sql = Sql & " Where (dbo.CONTENEDOR.FK_CAJAS Is Null) And (Not (dbo.CONTENEDOR.COD_CLIENTE Is Null))"
'Sql = Sql & " ORDER BY dbo.CONTENEDOR.COD_CLIENTE, dbo.CONTENEDOR.NRO_CAJA"
'
'Set rsContenedor = New ADODB.Recordset
'
'Set ConBasa = New ADODB.Connection
'ConBasa.Open strConBasa , 0 ,1
'
'rsContenedor.Open Sql, strConBasa , 0 ,1
'Do While Not rsContenedor.EOF
'
'
'    Sql = " UPDATE    dbo.CONTENEDOR"
'Sql = Sql & " SET              FK_CAJAS =" & rsContenedor!ID_CAJA
'Sql = Sql & " Where COD_CLIENTE = " & rsContenedor!COD_CLIENTE
'Sql = Sql & " And NRO_CAJA =" & rsContenedor!NRO_CAJA
'ExecutarSql Sql
'
'
'    rsContenedor.MoveNext
'Loop



End Sub

Private Sub Command30_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset



Sql = " SELECT     DOCUMENTO, DESCRIPCION"
Sql = Sql & "  From ECOGAS_DIFERENCIAS ORDER BY DOCUMENTO"

rs.Open Sql, strConBasa
Do While Not rs.EOF
    Sql = " UPDATE INDICES "
    Sql = Sql & " SET DESCRIPCION ='" & Trim(rs!Descripcion) & "'"
    Sql = Sql & " Where COD_CLIENTE =4 "
    Sql = Sql & " And ID_CODIGO_DOCUMENTO = " & rs!Documento
    ExecutarSql Sql
    rs.MoveNext
 Loop


End Sub

Private Sub Command31_Click()
                    
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        
            Sql = " SELECT     LACAJASIMIESTRO.PROCESO, LACAJASIMIESTRO.SUCURSAL, LACAJASIMIESTRO.SINIESTRO, LACAJASIMIESTRO.RAMO, LACAJASIMIESTRO.CAJA,"
            Sql = Sql & " INDICES.COD_CLIENTE , INDICES.ID_CODIGO_DOCUMENTO, INDICES.Indice, INDICES.Descripcion"
            Sql = Sql & " FROM         LACAJASIMIESTRO INNER JOIN"
            Sql = Sql & " INDICES ON LACAJASIMIESTRO.PROCESO = INDICES.ID_CODIGO_DOCUMENTO"
            Sql = Sql & " Where (INDICES.COD_CLIENTE = 163)"
                    
                 rs.Open Sql, strConBasa
                    
            Do While Not rs.EOF
                    Sql = " INSERT INTO REFERENCIAS"
                    Sql = Sql & " ( COD_CLIENTE, NRO_CAJA "
                    Sql = Sql & " , INDICE, COD_DOCUMENTO , DESCRIPCION "
                    Sql = Sql & " , FECHA_MODIFICACION, FECHA_CREACION"
                    Sql = Sql & " , USUARIO_MODIFICACION, FK_PERSONAL_CREACION"
                    Sql = Sql & " , FK_PERSONAL_MODIFICACION, BORRADO,   NRO_DESDE, NRO_HASTA )"
                    Sql = Sql & " VALUES "
                    Sql = Sql & " ( 163 ," & rs!Caja
                    Sql = Sql & " , '" & rs!Indice & "'," & rs!Proceso & ",'CBO Sucursal: " & rs!Sucursal & " Ramo: " & rs!RAMO & "'"
                    Sql = Sql & " , '14/09/2010', '14/09/2010'"
                    Sql = Sql & " , 99,99 "
                    Sql = Sql & " , 99,0," & rs!SINIESTRO & "," & rs!SINIESTRO & ")"
                    ExecutarSql Sql
                  rs.MoveNext
               Loop
               
                    
                    
                    
End Sub

Private Sub Command32_Click()

    Dim rs As New ADODB.Recordset
    Dim ID As Long
    Dim Sql As String
    
    Sql = " SELECT     ID_REFERENCIA, COD_ID_REFERENCIA "
Sql = Sql & "  From REFERENCIAS"
    Sql = Sql & "  Where (Not (NRO_CAJA Is Null)) And (COD_ID_REFERENCIA Is Null) "
    
    rs.CursorLocation = adUseClient
    
    
    rs.Open Sql, ConActiva, 2, 3
    
    ID = 603812
    

    Do While Not rs.EOF
        rs!COD_ID_REFERENCIA = ID
        ID = ID + 1
        rs.Update
        rs.MoveNext

        
    Loop
    
    




End Sub

Private Sub Command33_Click()
    Dim Sql  As String
    Dim Caja As Long
    
    Dim rs As New ADODB.Recordset
    
    Caja = 44405

            Sql = " SELECT     CAJASEXPURGOSUPERVIELLE.COD_ID_REFERENCIA, CAJASEXPURGOSUPERVIELLE.COD_CLIENTE, CAJASEXPURGOSUPERVIELLE.NRO_CAJA,"
            Sql = Sql & " CAJASEXPURGOSUPERVIELLE.INDICE, CAJASEXPURGOSUPERVIELLE.DESCRIPCION, CAJASEXPURGOSUPERVIELLE.FECHA_DESDE,"
            Sql = Sql & " CAJASEXPURGOSUPERVIELLE.FECHA_HASTA, CAJASEXPURGOSUPERVIELLE.NRO_DESDE, CAJASEXPURGOSUPERVIELLE.NRO_HASTA,"
            Sql = Sql & " CAJASEXPURGOSUPERVIELLE.LETRA_DESDE, CAJASEXPURGOSUPERVIELLE.LETRA_HASTA, CAJASEXPURGOSUPERVIELLE.Expr1,"
            Sql = Sql & " CAJASEXPURSUPERCANT.cant"
            Sql = Sql & " FROM CAJASEXPURGOSUPERVIELLE INNER JOIN"
            Sql = Sql & " CAJASEXPURSUPERCANT ON CAJASEXPURGOSUPERVIELLE.NRO_CAJA = CAJASEXPURSUPERCANT.NRO_CAJA"
            Sql = Sql & " GROUP BY CAJASEXPURGOSUPERVIELLE.COD_ID_REFERENCIA, CAJASEXPURGOSUPERVIELLE.COD_CLIENTE, CAJASEXPURGOSUPERVIELLE.NRO_CAJA,"
            Sql = Sql & " CAJASEXPURGOSUPERVIELLE.INDICE, CAJASEXPURGOSUPERVIELLE.DESCRIPCION, CAJASEXPURGOSUPERVIELLE.FECHA_DESDE,"
            Sql = Sql & " CAJASEXPURGOSUPERVIELLE.FECHA_HASTA, CAJASEXPURGOSUPERVIELLE.NRO_DESDE, CAJASEXPURGOSUPERVIELLE.NRO_HASTA,"
            Sql = Sql & " CAJASEXPURGOSUPERVIELLE.LETRA_DESDE, CAJASEXPURGOSUPERVIELLE.LETRA_HASTA, CAJASEXPURGOSUPERVIELLE.Expr1,"
            Sql = Sql & " CAJASEXPURSUPERCANT.cant"
            Sql = Sql & " Having (CAJASEXPURSUPERCANT.cant > 50)"
            Sql = Sql & " ORDER BY CAJASEXPURGOSUPERVIELLE.COD_ID_REFERENCIA"
            
            rs.Open Sql, strConBasa
            
            Do While Not rs.EOF
                rs.MoveNext
                Caja = Caja + 1
                Sql = " Update CAJASEXPURGOSUPERVIELLE SET NRO_CAJA =" & Caja & " Where COD_ID_REFERENCIA = " & rs!COD_ID_REFERENCIA
                ExecutarSql Sql
                rs.MoveNext
          Loop
            




End Sub

Private Sub Command34_Click()
Dim con As New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=E:\datas10112010\Movimiento\Mov10112010.mdb"
Dim rs As New ADODB.Recordset
Dim DocNumero As Long
Dim Sql As String
Dim i As Integer
Dim C As Integer

On Error Resume Next

For C = 1 To 6000
            Set rs = New ADODB.Recordset
            Debug.Print C
            Debug.Print Err.Number
            rs.Open "SELECT NUMEROCAJA FROM Mov" & Format(C, "0000"), con
            If Err.Number = -2147217865 Then
'                MsgBox Err.Description
                Err.Clear
            Else
                Do While Not rs.EOF
                   Sql = "INSERT INTO MOVIMIENTO ( CAJA, CLIENTE )  VALUES  (" & rs!NUMEROCAJA & "," & C & " )"
                   con.Execute Sql
                   rs.MoveNext
                Loop
            End If
             
Next


End Sub

Private Sub Command35_Click()

Dim con As New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=E:\datas10112010\Movimiento\Mov10112010.mdb"
Dim rs As New ADODB.Recordset
Dim DocNumero As Long
Dim Sql As String
Dim i As Integer
Dim C As Integer
Sql = " SELECT MOVIMIENTO.CAJA, Count(*) AS CANTIDAD "
Sql = Sql & " From MOVIMIENTO GROUP BY MOVIMIENTO.CAJA ORDER BY Count(*) DESC; "

rs.Open Sql, con

Do While Not rs.EOF

Sql = "    UPDATE CAJAS SET CAJAS.CONSULTAS =  " & rs!cantidad
Sql = Sql & "  WHERE CAJAS.IDCaja= " & rs!Caja

con.Execute Sql

    rs.MoveNext
Loop



End Sub

Private Sub Command36_Click()
'Dim con As New ADODB.Connection
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=E:\datas10112010\REFERENCIAS\Mov10112010.mdb"
'Dim rs As New ADODB.Recordset
'
'
'DCTO0022
'
'For C = 1 To 5
'
'
'Set rs = New ADODB.Recordset
'rs.Open "SELECT * FROM DCTO0022 ", con
'
'Do While Not rs.EOF
'    rs!FECHA_DESDE = DateAdd("d", rs!FechaDesde, "28/12/1800")
'    rs!FECHA_HASTA = DateAdd("d", rs!FechaHasta, "28/12/1800")
'
'    rs.MoveNext
'Loop
'






End Sub

Private Sub Command37_Click()


Dim rs As New ADODB.Recordset
Dim Sql As String
Dim i As Integer
Dim Campo As String

Sql = "SELECT     CAJA, F3, F4, F5, F6, F7, F8, F9, F10, F11, F12, F13, F14, F15, F16, F17, F18, F19, F20"
Sql = Sql & " From ordenes$ "

rs.Open Sql, strConBasa

    Do While Not rs.EOF
        
        For i = 3 To 20
        Campo = "F" & CStr(i)
            If IsNumeric(rs.Fields(Campo)) = True Then
                If rs.Fields(Campo) <> 0 Then
                Sql = "INSERT INTO COPIARORDENES (CAJA, ORDEN)"
                Sql = Sql & " VALUES (" & rs!Caja & "," & rs.Fields(Campo) & ")"
                ExecutarSql Sql
                End If
            
            
            End If
            
            
        
        Next
        
        
    
    
        rs.MoveNext
    Loop
    


End Sub

Private Sub Command38_Click()


Dim rs As New ADODB.Recordset
Dim Sql As String

Sql = " SELECT     COPIARORDENES.CAJA, LEGAJOS.ID_LEGAJO"
Sql = Sql & " FROM ORDEN_LEGAJOS_DETALLE INNER JOIN"
Sql = Sql & " ORDEN_LEGAJOS ON ORDEN_LEGAJOS_DETALLE.COD_ORDEN_LEGAJO = ORDEN_LEGAJOS.ID_ORDEN_LEGAJO INNER JOIN"
Sql = Sql & " LEGAJOS ON ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO = LEGAJOS.ID_CLIENTE_LEGAJO AND"
Sql = Sql & " ORDEN_LEGAJOS.COD_CLIENTE = LEGAJOS.COD_CLIENTE RIGHT OUTER JOIN"
Sql = Sql & " COPIARORDENES ON ORDEN_LEGAJOS.ID_ORDEN_LEGAJO = COPIARORDENES.ORDEN"
Sql = Sql & " ORDER BY COPIARORDENES.ORDEN"
rs.Open Sql, strConBasa

Do While Not rs.EOF
        Sql = " Update LEGAJOS"
        Sql = Sql & " SET REARCHIVO_CAJA =" & rs!Caja
        Sql = Sql & " Where ID_LEGAJO =" & rs!ID_LEGAJO
        ExecutarSql Sql
        rs.MoveNext
Loop




End Sub

Private Sub Command4_Click()
Dim con As New ADODB.Connection
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=E:\backup base\Comercial\BASA Soluciones Digitales\Basa 2007\Presupuestos y Presentaciones\La Caja Seguro\Archivos\lacaja.mdb"
Dim rs As New ADODB.Recordset
Dim DocNumero As Long
Dim Sql As String
Dim i As Integer
Dim C As Integer


For C = 1 To 5


Set rs = New ADODB.Recordset
rs.Open "SELECT * FROM " & C, con
 
Do While Not rs.EOF
    DocNumero = rs.Fields(5).value
    For i = 6 To 11
     If rs.Fields(i).value = 1 Then
        Sql = "INSERT INTO SEGURO ( DOCUMENTO, SEGURO, PERIODO )"
        Sql = Sql & " values (" & DocNumero & ",'" & rs.Fields(i).Name & "', " & rs.Fields(12).value & " )"
        con.Execute Sql
     End If
     
    
    Next
    
    
    rs.MoveNext
    
   Loop

Next



End Sub

Private Sub Command40_Click()
        Dim con As New ADODB.Connection
        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=E:\datas10112010\Movimiento\Mov10112010.mdb"
        
       Dim CONB As New ADODB.Connection
       CONB.Open strConBasa
        Dim Sql As String
        Dim Sqlc As String
        Dim Sqlu As String
        Dim RSC As ADODB.Recordset
        Dim rs As New ADODB.Recordset
        Dim ID, IDCliente, IDCaja, estado, CONTENEDOR_ESTADO  As String
        Dim DIGITO, Modulo, Ubicacion As String
        Dim MODULO_TEXT, UBICACION_TEXT As String
        
        Sql = "     SELECT ID,IDCliente, IDCaja , ESTADO, DIGITO , MODULO, Ubicacion , "
        Sql = Sql & "   CAJAS.MODULO_TEXT , CAJAS.UBICACION_TEXT , Digito_Verificador "
        Sql = Sql & "   From Cajas "
        Sql = Sql & " where IDCliente in(13,27) "
        Sql = Sql & "   ORDER BY ID "
        
        
        Sqlc = " SELECT     ID_CONTENEDOR, ESTANTERIA, ESTADO "
        Sqlc = Sqlc & " From CONTENEDOR "
        Sqlc = Sqlc & "  Where Estanteria = 150 "
        Sqlc = Sqlc & "  And estado = 1 "
        Sqlc = Sqlc & " ORDER BY ID_CONTENEDOR "
        
        Set RSC = New ADODB.Recordset
        RSC.Open Sqlc, strConBasa

CONB.BeginTrans
        
        rs.Open Sql, con
        
        Do While Not rs.EOF
             ID = rs!ID
             IDCliente = rs!IDCliente + 10000
             IDCaja = rs!IDCaja
             Select Case rs!estado
             Case "EN TRANSITO"
                estado = 1130
                CONTENEDOR_ESTADO = 3
             Case "OCUPADA"
                estado = 1120
                CONTENEDOR_ESTADO = 2
             Case "LIBRE"
                estado = 1100
                                              
                CONTENEDOR_ESTADO = 1
             End Select
             
            Sqlu = " Update CONTENEDOR "
            Sqlu = Sqlu & " SET ESTADO =" & CONTENEDOR_ESTADO
            Sqlu = Sqlu & "  , COD_CLIENTE =" & IDCliente
            Sqlu = Sqlu & "  , NRO_CAJA =" & IDCaja
            Sqlu = Sqlu & "  Where ID_CONTENEDOR = " & RSC!ID_CONTENEDOR
             
             CONB.Execute Sqlu
             RSC.MoveNext
            
            Sqlu = "  INSERT INTO CAJAS"
            Sqlu = Sqlu & " (ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_ESTADO, DIGITO_VERIFICADOR , FK_USUARIO_CREACION_CAJA , FECHA_CREACION_CAJA , FK_MODULO   )"
            Sqlu = Sqlu & " VALUES (" & ID & "," & IDCliente & "," & IDCaja & "," & estado & "," & rs!Digito_Verificador & ",99, '23/12/2010' , " & ID & ")"
            CONB.Execute Sqlu
            rs.MoveNext
        Loop
        
    CONB.RollbackTrans
        
        
End Sub

Private Sub Command41_Click()
 Dim con As New ADODB.Connection
        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=E:\datas10112010\Movimiento\Mov10112010.mdb"
        
       Dim CONB As New ADODB.Connection
       CONB.Open strConBasa
        Dim Sql As String
        Dim Sqlc As String
        Dim Sqlu As String
        Dim RSC As ADODB.Recordset
        Dim rs As New ADODB.Recordset
        Dim ID, IDCliente, IDCaja, estado, CONTENEDOR_ESTADO  As String
        Dim DIGITO, Modulo, Ubicacion As String
        Dim MODULO_TEXT, UBICACION_TEXT As String
        
        Sql = "     SELECT ID,IDCliente, IDCaja , ESTADO, DIGITO , MODULO, Ubicacion , "
        Sql = Sql & "   CAJAS.MODULO_TEXT , CAJAS.UBICACION_TEXT , Digito_Verificador "
        Sql = Sql & "   From Cajas "
        Sql = Sql & " where IDCliente in(13,27) "
        Sql = Sql & "   ORDER BY ID "
        
        
        Sqlc = " SELECT     ID_CONTENEDOR, ESTANTERIA, ESTADO "
        Sqlc = Sqlc & " From CONTENEDOR "
        Sqlc = Sqlc & "  Where Estanteria = 150 "
        Sqlc = Sqlc & "  And estado = 1 "
        Sqlc = Sqlc & " ORDER BY ID_CONTENEDOR "
        
        Set RSC = New ADODB.Recordset
        RSC.Open Sqlc, strConBasa

CONB.BeginTrans
        
        rs.Open Sql, con
        
        Do While Not rs.EOF
             ID = rs!ID
             IDCliente = rs!IDCliente + 10000
             IDCaja = rs!IDCaja
             Select Case rs!estado
             Case "EN TRANSITO"
                estado = 1130
                CONTENEDOR_ESTADO = 3
             Case "OCUPADA"
                estado = 1120
                CONTENEDOR_ESTADO = 2
             Case "LIBRE"
                estado = 1100
                                              
                CONTENEDOR_ESTADO = 1
             End Select
             
            Sqlu = " Update CONTENEDOR "
            Sqlu = Sqlu & " SET ESTADO =" & CONTENEDOR_ESTADO
            Sqlu = Sqlu & "  , COD_CLIENTE =" & IDCliente
            Sqlu = Sqlu & "  , NRO_CAJA =" & IDCaja
            Sqlu = Sqlu & "  Where ID_CONTENEDOR = " & RSC!ID_CONTENEDOR
             
             CONB.Execute Sqlu
             RSC.MoveNext
            
            Sqlu = "  INSERT INTO CAJAS"
            Sqlu = Sqlu & " (ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_ESTADO, DIGITO_VERIFICADOR , FK_USUARIO_CREACION_CAJA , FECHA_CREACION_CAJA , FK_MODULO   )"
            Sqlu = Sqlu & " VALUES (" & ID & "," & IDCliente & "," & IDCaja & "," & estado & "," & rs!Digito_Verificador & ",99, '23/12/2010' , " & ID & ")"
            CONB.Execute Sqlu
            rs.MoveNext
        Loop
        
    CONB.RollbackTrans
        
End Sub


Private Sub Command43_Click()
 Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim Modulo As String
    Dim Sql As String
    
        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=E:\datas10112010\Movimiento\Mov10112010.mdb"
        Sql = "  SELECT CAJAS_ERROR_25.IDCaja, CAJAS_ERROR_25.ID, CAJAS_ERROR_25.DIGITO_TEXT"
        Sql = Sql & "  From CAJAS_ERROR_25 "
        rs.CursorLocation = adUseClient
        rs.Open Sql, con, adOpenKeyset, adLockOptimistic
        Do While Not rs.EOF
            rs!DIGITO_TEXT = Format(Digito_Verificador(rs!IDCaja), "00")
            rs.Update
            rs.MoveNext
         Loop
         MsgBox "Terminado"
End Sub

Private Sub Command44_Click()
    Dim Sql As String
 Dim rs As New ADODB.Recordset
    
    Sql = " SELECT     REFERENCIAS.COD_ID_REFERENCIA"
 Sql = Sql & " FROM         REFERENCIAS LEFT OUTER JOIN"
  Sql = Sql & "  CONTENEDOR ON REFERENCIAS.NRO_CAJA = CONTENEDOR.NRO_CAJA AND REFERENCIAS.COD_CLIENTE = CONTENEDOR.COD_CLIENTE"
 Sql = Sql & " Where (CONTENEDOR.COD_CLIENTE Is Null)"
 Sql = Sql & " ORDER BY REFERENCIAS.COD_CLIENTE"
 
 rs.Open Sql, ConActiva, 0, 1
 
 Do While Not rs.EOF
 
Sql = "  DELETE FROM REFERENCIAS Where COD_ID_REFERENCIA = " & rs!COD_ID_REFERENCIA
ExecutarSql Sql
    rs.MoveNext
 Loop
 

End Sub

Private Sub Command45_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset
Sql = "  SELECT     CLIENTE, CAJA, EMPRESA"
Sql = Sql & "  From CAJASBASAALSINA"
Sql = Sql & "  ORDER BY CAJA DESC"

rs.Open Sql, strConBasa
Do While Not rs.EOF
    If rs!Caja < 100000 Then
        Sql = " Update Cajas SET DEPOSITO ='ALSINA'"
        Sql = Sql & " Where FK_CLIENTE = " & rs!Cliente
        Sql = Sql & " And NRO_CAJA = " & rs!Caja
    Else
        Sql = " Update Cajas SET DEPOSITO ='ALSINA'"
        Sql = Sql & " Where NRO_CAJA = " & rs!Caja
    End If
    ExecutarSql Sql
    rs.MoveNext
Loop




End Sub

Private Sub Command46_Click()
Dim emailOutlookApp As Outlook.Application
Dim emailNameSpace As Outlook.Namespace
Dim emailFolder As Outlook.MAPIFolder
Dim emailItem As Outlook.MailItem
Dim EmailRecipient As Recipient
Dim emailItem2 As Outlook.MailItem
Dim fol As Outlook.Items
Dim F As Long
Dim MaxCorreo As Long
'-----Open Outlook in a background process and the Inbox Folder-----
Set emailOutlookApp = CreateObject("Outlook.Application")
Set emailNameSpace = emailOutlookApp.GetNamespace("MAPI")
 Set emailFolder = emailNameSpace.GetDefaultFolder(olFolderInbox)
Rem Set emailFolder = emailNameSpace.GetDefaultFolder(olFolderSentMail)

MsgBox emailFolder.Folders.Count
Dim i As Integer

For i = 1 To emailFolder.Folders.Count
    MsgBox emailFolder.Folders.Item(i).Name
    
    Set fol = emailFolder.Folders.Item(i)
 

    
    
    For F = 1 To fol.Count - 1
        Set emailItem2 = fol.Item(F)
        MsgBox emailItem2.SenderEmailAddress
        MsgBox emailItem2.SenderName
    
    
    Next
    
    
Next
MsgBox emailFolder.Items.Count

'
'
'
'Dim Sql As String
'Dim Rs As ADODB.Recordset
'Dim ENTRYID As String
'
'
''For i = 1 To emailFolder.Items.Count
''    Set emailItem2 = emailFolder.Items(i)
''
''     emailItem2.FlagRequest = i
''
''Next
'
'    Set emailItem2 = emailFolder.Items(1)
'
''For i = 1 To emailFolder.Items.Count
'' Set emailItem2 = emailFolder.Items(i)
''    emailItem2.Categories = ""
''    emailItem2.Save
'' Next
'
'     Rem If Not IsNumeric(emailItem2.Categories)  Then
'     If Trim(emailItem2.Categories) = "" Then
'
'        Rem ENTRYID = Replace(emailItem2.InternetCodepage & emailItem2.ReceivedByEntryID & emailItem2.Subject, "'", "´" & Mid(emailItem2.Body, 1, 7000))
'        ENTRYID = emailItem2.ENTRYID
'
'           emailItem2.Categories = MaxCorreo
'
'        Sql = "  SELECT     ENTRYID, ID_CORREO"
'        Sql = Sql & " From dbo.CORREOS"
'        Sql = Sql & "  WHERE     (ENTRYID = '" & ENTRYID & " ')"
'        Set Rs = New ADODB.Recordset
'        Rs.Open Sql, strConBasa , 0 ,1
'        If Rs.EOF Then
'           Sql = " INSERT INTO dbo.CORREOS"
'           Sql = Sql & " (ENTRYID, ENVIADO, ASUNTO, CUERPO, KF_USUARIO)"
'           Sql = Sql & " VALUES ('" & ENTRYID & "','" & emailItem2.SenderEmailAddress & "','" & Replace(emailItem2.Subject, "'", "´") & "','" & Mid(Replace(emailItem2.Body, "'", "`"), 1, 7000) & "'," & 99 & ")"
'          Rem  ExecutarSql Sql
'           emailItem2.Categories = MaxCorreo
'           emailItem2.Save
'        Else
'            emailItem2.Categories = Rs!ID_CORREO
'            emailItem2.Save
'
'        End If
'     End If
'
'Next
'
'For i = 1 To emailFolder.Folders.Item("SUPER").Items.Count
'    Set emailItem2 = emailFolder.Folders.Item("SUPER").Items(i)
'   Rem  MsgBox emailItem2.Body
'    emailItem2.Categories = "REQUE:" & i & " id " & emailItem2.ENTRYID
'    Rem emailItem2.FlagStatus = olFlagComplete
'
'     emailItem2.FlagRequest = i
'
''     Sql = " INSERT INTO dbo.CORREOS"
''      Sql = Sql & " (ENTRYID, ENVIADO, ASUNTO, CUERPO, KF_USUARIO)"
''Sql = Sql & " VALUES ('" & emailItem2.InternetCodepage & emailItem2.SenderEmailAddress & emailItem2.Subject & "','" & emailItem2.SenderEmailAddress & "','" & emailItem2.Subject & "','" & Replace(Mid(emailItem2.Body, 1, 2000), "'", "`") & "'," & ctlPersonal.Valor & ")"
''         ExecutarSql Sql
''
'    emailItem2.Save
'Next
'

Set emailNameSpace = Nothing
Set emailFolder = Nothing
Set emailItem = Nothing
Set emailOutlookApp = Nothing

MsgBox "ok"

End Sub

Private Sub Command47_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
    Sql = " SELECT     LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, CONTENEDOR_HISTORICO.ESTANTERIA,"
    Sql = Sql & " CONTENEDOR_HISTORICO.HORIZONTAL, CONTENEDOR_HISTORICO.VERTICAL, CONTENEDOR_HISTORICO.ADELANTE_ATRAS,"
    Sql = Sql & " CONTENEDOR_HISTORICO.NRO_ESTANTE , CONTENEDOR_HISTORICO.estado, CONTENEDOR.NRO_CAJA, CONTENEDOR.COD_CLIENTE"
    Sql = Sql & "  FROM         LECTURACOLECTOR INNER JOIN"
    Sql = Sql & "  CONTENEDOR_HISTORICO ON LECTURACOLECTOR.CAJA = CONTENEDOR_HISTORICO.NRO_CAJA AND"
    Sql = Sql & "  LECTURACOLECTOR.CLIENTE = CONTENEDOR_HISTORICO.COD_CLIENTE INNER JOIN"
    Sql = Sql & "  CONTENEDOR ON CONTENEDOR_HISTORICO.ESTANTERIA = CONTENEDOR.ESTANTERIA AND"
    Sql = Sql & "  CONTENEDOR_HISTORICO.HORIZONTAL = CONTENEDOR.HORIZONTAL AND CONTENEDOR_HISTORICO.VERTICAL = CONTENEDOR.VERTICAL AND"
    Sql = Sql & "  CONTENEDOR_HISTORICO.Adelante_Atras = CONTENEDOR.Adelante_Atras"
    Sql = Sql & "  Where (LECTURACOLECTOR.NUMERO_LECTURA = 15996) And (CONTENEDOR_HISTORICO.Estanteria = 323)"


 Sql = "  SELECT     LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, CAMBIOPOSICION.HORIZONTAL,"
 Sql = Sql & "                       CAMBIOPOSICION.ESTANTERIA, CAMBIOPOSICION.VERTICAL, CAMBIOPOSICION.ADELANTE_ATRAS, CAMBIOPOSICION.NRO_ESTANTE,"
  Sql = Sql & "                      CAMBIOPOSICION.estado , CAMBIOPOSICION.COD_CLIENTE, CAMBIOPOSICION.NRO_CAJA, CAMBIOPOSICION.Fecha"
 Sql = Sql & "  FROM         CAMBIOPOSICION INNER JOIN"
                 Sql = Sql & "       LECTURACOLECTOR ON CAMBIOPOSICION.NRO_CAJA = LECTURACOLECTOR.CAJA AND"
                       Sql = Sql & "  CAMBIOPOSICION.COD_CLIENTE = LECTURACOLECTOR.Cliente"
 Sql = Sql & "  Where (LECTURACOLECTOR.NUMERO_LECTURA = 15996) And (CAMBIOPOSICION.Estanteria = 323)"


Set rs = New ADODB.Recordset




rs.Open Sql, strConBasa





Do While Not rs.EOF
Debug.Print rs!Caja
    Sql = Sql & "  Update CONTENEDOR"
    Sql = Sql & "  SET COD_CLIENTE = NULL, NRO_CAJA = NULL, ESTADO = 1"
    Sql = Sql & "  Where COD_CLIENTE =" & rs!Cliente & "  And NRO_CAJA = " & rs!Caja
    ExecutarSql Sql
    
    
    Sql = " Update CONTENEDOR"
 Sql = Sql & "  SET  NRO_CAJA =" & rs!Caja & ", COD_CLIENTE =" & rs!Cliente & " , ESTADO =" & rs!estado
 Sql = Sql & "  Where Estanteria =  " & rs!Estanteria
   Sql = Sql & "  And Horizontal = " & rs!Horizontal
   Sql = Sql & "  And Vertical = " & rs!Vertical
    Sql = Sql & "  And Adelante_Atras = " & rs!Adelante_Atras
    
    ExecutarSql Sql
    
    rs.MoveNext

Loop



End Sub

Private Sub Command48_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim RSC As New ADODB.Recordset

Sql = " SELECT     CAJA, CLIENTE "
Sql = Sql & " From CAJASSINPOS "

rs.Open Sql, strConBasa

Sql = " SELECT     ID_CONTENEDOR, FK_CAJAS, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA,   NRO_REMITO"
Sql = Sql & "  From CONTENEDOR "
Sql = Sql & "  Where (Estanteria = 150) And (estado = 1) ORDER BY ID_CONTENEDOR"

RSC.Open Sql, strConBasa

Do While Not rs.EOF
    Sql = " Update CONTENEDOR "
Sql = Sql & " SET ESTADO =2"
Sql = Sql & " , COD_CLIENTE =" & rs!Cliente
Sql = Sql & "  , NRO_CAJA =" & rs!Caja
Sql = Sql & "   Where ID_CONTENEDOR = " & RSC!ID_CONTENEDOR
    ExecutarSql Sql

    RSC.MoveNext
    rs.MoveNext
Loop




End Sub

Private Sub Command49_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT     ID_CAJA, ROLLO"
Sql = Sql & "  From Cajas"
Sql = Sql & " Where ROLLO = 440 "
Sql = Sql & " ORDER BY ID_CAJA DESC "
rs.Open Sql, strConBasa
Dim caja1 As Long
Dim caja2 As Long

Do While Not rs.EOF
caja1 = rs!ID_CAJA
rs.MoveNext
caja2 = rs!ID_CAJA
Sql = " Update TEM_CAJA_IMPRE SET    CAJA_1 =" & caja1 & ", CAJA_2 =" & caja2
ExecutarSql Sql
rs.MoveNext

Loop


End Sub

Private Sub Command5_Click()

'Dim Sql As String
'Sql = " Update dbo.DOCUMENTOS_DIGITALES"
'Sql = Sql & "  SELECT     LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, CONTENEDOR_HISTORICO.ESTANTERIA, "
'                      CONTENEDOR_HISTORICO.HORIZONTAL, CONTENEDOR_HISTORICO.VERTICAL, CONTENEDOR_HISTORICO.ADELANTE_ATRAS,
'                      CONTENEDOR_HISTORICO.NRO_ESTANTE , CONTENEDOR_HISTORICO.estado, CONTENEDOR.NRO_CAJA, CONTENEDOR.COD_CLIENTE
'FROM         LECTURACOLECTOR INNER JOIN
'                      CONTENEDOR_HISTORICO ON LECTURACOLECTOR.CAJA = CONTENEDOR_HISTORICO.NRO_CAJA AND
'                      LECTURACOLECTOR.CLIENTE = CONTENEDOR_HISTORICO.COD_CLIENTE INNER JOIN
'                      CONTENEDOR ON CONTENEDOR_HISTORICO.ESTANTERIA = CONTENEDOR.ESTANTERIA AND
'                      CONTENEDOR_HISTORICO.HORIZONTAL = CONTENEDOR.HORIZONTAL AND CONTENEDOR_HISTORICO.VERTICAL = CONTENEDOR.VERTICAL AND
'                      CONTENEDOR_HISTORICO.Adelante_Atras = CONTENEDOR.Adelante_Atras
'WHERE     (LECTURACOLECTOR.NUMERO_LECTURA = 15996) AND (CONTENEDOR_HISTORICO.ESTANTERIA = 323)Set LETRA_HASTA = LETRA_DESDE"
'Sql = Sql & "  Where (Not (LETRA_DESDE Is Null)) And (LETRA_HASTA Is Null)"
'ExecutarSql Sql
'
'Sql = Sql & "  Update dbo.DOCUMENTOS_DIGITALES"
'Sql = Sql & "  Set NRO_HASTA = NRO_DESDE"
'Sql = Sql & "  Where (Not (NRO_DESDE Is Null)) And (NRO_HASTA Is Null) And (NRO_DESDE > 0)"
'ExecutarSql Sql
'
'Sql = Sql & "  Update dbo.DOCUMENTOS_DIGITALES"
'Sql = Sql & " Set FECHA_HASTA = FECHA_DESDE"
'Sql = Sql & "  Where (Not (FECHA_DESDE Is Null)) And (FECHA_HASTA Is Null)"
'ExecutarSql Sql

End Sub

Private Sub Command50_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection


con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\CajasCustodia\Custodia.accdb;Persist Security Info=False"



Sql = " SELECT IDCliente, IDCaja, Ubicacion , estado FROM CajasCustodia  ORDER BY IDCaja;"

 

rs.Open Sql, con

Dim Sqlc As String


Sqlc = " SELECT        ID_CONTENEDOR, ESTANTERIA, ESTADO"
Sqlc = Sqlc & "  From CONTENEDOR"
Sqlc = Sqlc & " WHERE        (ESTANTERIA BETWEEN 1000 AND 2000) "
Sqlc = Sqlc & " ORDER BY ID_CONTENEDOR "


Dim IDCliente As Integer
Dim IDCaja As Long
Dim Ubicacion As Stream
Dim estado As Integer
Dim ConBasa As New ADODB.Connection
Dim RSC As New ADODB.Recordset

  

ConBasa.Open strConBasa

RSC.Open Sqlc, strConBasa
Do While Not rs.EOF

Select Case UCase(Trim(rs!estado))
Case "EN TRANSITO"
estado = 3
Case "LIBRE"
estado = 4
Case "OCUPADA"
estado = 2

Case "RESERVADA"
estado = 4
End Select




IDCliente = rs!IDCliente + 1000
IDCaja = rs!IDCaja
      If IDCaja > 0 Then
'
'       Sql = "  INSERT INTO CAJAS_CUSTODIA_1    (ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_ESTADO, DEPOSITO , DIGITO_VERIFICADOR, FK_USUARIO_CREACION_CAJA , FECHA_CREACION_CAJA)"
'Sql = Sql & " VALUES     (" & IDCaja & " ," & IDCliente & " ," & IDCaja & ",1120, 'ALFONSI' ," & Digito_Verificador(CStr(IDCaja)) & ",99, '19/09/2011' )"
'
'ExecutarSql Sql

Sqlc = " UPDATE       CONTENEDOR SET "
Sqlc = Sqlc & " ESTADO = " & estado
Sqlc = Sqlc & " , FK_CAJAS =" & IDCaja
Sqlc = Sqlc & " , COD_CLIENTE =" & IDCliente
Sqlc = Sqlc & ", NRO_CAJA =" & IDCaja
Sqlc = Sqlc & ", UB_PROVISORIA=" & rs!Ubicacion
Sqlc = Sqlc & " Where ID_CONTENEDOR = " & RSC!ID_CONTENEDOR

    
    ExecutarSql Sqlc

RSC.MoveNext


End If

    rs.MoveNext
Loop





End Sub

Private Sub Command51_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim Codigo_indice As String
Dim ID_indice2 As Integer
Dim Sucursal As Integer
Dim FECHA_DESDE  As String
Dim FECHA_HASTA As String
Dim ConBasa As New ADODB.Connection

ConBasa.Open "Provider=SQLOLEDB.1;Password=21877471; Persist Security Info=False;User ID=sa;Initial Catalog=BASE2;Data Source=SERVER-BASA"

Dim codigodoc As Integer

con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Datos Custodia\DatosCustodia.accdb;Persist Security Info=False"

Rem 2 es tarjetas



Sql = " SELECT * "
Sql = Sql & " From DISCO197TARJETAS "
rs.CursorLocation = adUseClient
rs.Open Sql, con, 2, 3

Dim fecham As String
fecham = " CONVERT(DATETIME, '2011-11-18 12:31:00', 102)"

Do While Not rs.EOF

FECHA_DESDE = DateAdd("d", rs!fechadesde, "28/12/1800")
FECHA_HASTA = DateAdd("d", rs!FechaHasta, "28/12/1800")

    codigodoc = "2" & CStr(Format(rs!DESDENUMERO, "0000"))
    
    
        Codigo_indice = BuscarIndice(1197, CLng(codigodoc))
        If Codigo_indice <> "0" Then
            ID_indice2 = ID_indice
            Sucursal = codigodoc
            rs!Export = 2
            rs.Update
            Sql = " UPDATE DISCO197 SET DISCO197.[EXPORT] = 2"
            Sql = Sql & vbCrLf & " WHERE (((DISCO197.[IDDOCUMENTO])=" & rs!IDDOCUMENTO & "));"
            con.Execute Sql
            
            
            Sql = " INSERT INTO REFERENCIAS"
            Sql = Sql & " (COD_CLIENTE, NRO_CAJA, COD_TIPO_ALMACENAMIENTO"
            Sql = Sql & vbCrLf & ", INDICE, COD_DOCUMENTO, DESCRIPCION "
            Sql = Sql & vbCrLf & " , FECHA_DESDE, FECHA_HASTA "
            Sql = Sql & vbCrLf & " , FK_PERSONAL_CREACION, FK_PERSONAL_MODIFICACION"
            Sql = Sql & vbCrLf & " , BORRADO ,  FECHA_MODIFICACION)"
            Sql = Sql & vbCrLf & " VALUES        ("
            Sql = Sql & vbCrLf & " 1197 ," & rs!IDCaja & ", 0"
            Sql = Sql & vbCrLf & ",'" & Codigo_indice & "'," & codigodoc & ",'" & UCase(Trim(rs!Descripcion)) & "'"
            Sql = Sql & vbCrLf & ",'" & FECHA_DESDE & "','" & FECHA_HASTA & "'"
            Sql = Sql & vbCrLf & ",99, 99 "
            Sql = Sql & vbCrLf & " , '0' ," & fecham & " )"
            
            ExecutarSql Sql
     End If
    
    rs.MoveNext
Loop


End Sub

Private Sub Command52_Click()


Dim Sql As String
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim fecha As String
Dim DATO As String


con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Datos Custodia\DatosCustodia.accdb;Persist Security Info=False"


Sql = " SELECT FacturaABC, NumeroFactura, FechaFacturacion, IDCliente, Nombre, IVA, CUIT, Subtotal, IVAInscripto,IvaNoInscripto, TotalFacturado"
  Sql = Sql & "   From FACTURAOK"
  Sql = Sql & "     ORDER BY FacturaABC, NumeroFactura ;"


rs.Open Sql, con

Clipboard.Clear

Do While Not rs.EOF

fecha = DateAdd("d", rs!FechaFacturacion, "28/12/1800")

DATO = DATO & rs!FacturaABC & vbTab & rs!NumeroFactura & vbTab & fecha & vbTab & rs!IDCliente & vbTab & rs!Nombre & vbTab & rs!IVA & vbTab & rs!Cuit & vbTab & rs!Subtotal & vbTab & rs!IVAInscripto & vbTab & rs!IvaNoInscripto & vbTab & rs!TotalFacturado & vbCrLf

rs.MoveNext

    Loop
    Clipboard.SetText DATO
MsgBox "Los datos fueren copios"

End Sub

Private Sub Command53_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim Codigo_indice As String
Dim ID_indice2 As Integer
Dim Sucursal As Integer
Dim FECHA_DESDE  As String
Dim FECHA_HASTA As String
Dim ConBasa As New ADODB.Connection

ConBasa.Open "Provider=SQLOLEDB.1;Password=21877471; Persist Security Info=False;User ID=sa;Initial Catalog=BASE2;Data Source=SERVER-BASA"

Dim codigodoc As Integer

con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Datos Custodia\DatosCustodia.accdb;Persist Security Info=False"

Rem 2 es tarjetas

rs.CursorLocation = adUseClient

Sql = " SELECT DISCO197ROLLO.IDDOCUMENTO, DISCO197ROLLO.DESCRIPCION, DISCO197ROLLO.DESDENUMERO, DISCO197ROLLO.HASTANUMERO, DISCO197ROLLO.EXPORT"
Sql = Sql & " From DISCO197ROLLO "
Sql = Sql & " WHERE (((DISCO197ROLLO.DESDENUMERO)=0));"
 

rs.Open Sql, con, 2, 3





Do While Not rs.EOF
    rs!DESDENUMERO = Mid(rs!Descripcion, 13)
     rs!HASTANUMERO = Mid(rs!Descripcion, 13)
     rs.Update
    rs.MoveNext
Loop






End Sub

Private Sub Command54_Click()

        Dim ConContenedor25 As New ADODB.Connection
        Dim ConBasa As New ADODB.Connection
        Dim ConCustodiaCajas As New ADODB.Connection

        Dim Sql As String
        Dim SqlBasa As String
        Dim estado As Integer



        Dim rsContenedor25 As New ADODB.Recordset
        Dim rsbasa As New ADODB.Recordset
        Dim RsCustodiaCajas As New ADODB.Recordset

        Dim COD_CLIENTE As Integer
        Dim Cod_Estado As Integer
        Dim cod_ubicacion_prov As String
        Dim cod_contenedor As Long
        Dim cod_caja As Long


        ConContenedor25.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=F:\CAMBIO.mdb"
        ConBasa.Open "Provider=SQLOLEDB.1;Password=21877471; Persist Security Info=False;User ID=sa;Initial Catalog=BASE2;Data Source=SERVER-BASA"
        ConCustodiaCajas.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Datos Custodia\DatosCustodia.accdb;Persist Security Info=False"


Sql = " SELECT ID_CONTENEDOR, ESTANTERIA , HORIZONTAL, VERTICAL "
Sql = Sql & " , COD_CLIENTE ,  NRO_CAJA ,  EMPRESA "
Sql = Sql & " From CONTENEDOR25 "
Sql = Sql & " WHERE ENBASA Is Null AND Not NRO_CAJA Is Null ; "

rsContenedor25.Open Sql, ConContenedor25


Do While Not rsContenedor25.EOF
 On Error GoTo 10
    If rsContenedor25!Empresa = "VCUS" Or rsContenedor25!Empresa = "CUST" Then
     If IsNull(rsContenedor25!NRO_CAJA) Then
        GoTo 10
     End If
     
     
        Sql = " SELECT IDCliente, IDCaja, Ubicacion,Estado "
        Sql = Sql & " FROM CAJAS1  "
        Sql = Sql & " WHERE IDCaja = " & rsContenedor25!NRO_CAJA

         Set RsCustodiaCajas = New ADODB.Recordset
         
         RsCustodiaCajas.Open Sql, ConCustodiaCajas
            
        If Not RsCustodiaCajas.EOF Then
        
            Cod_Estado = 0
        Select Case RsCustodiaCajas!estado
            
        Case "EN TRANSITO"
            Cod_Estado = 3
        Case "LIBRE"
            Cod_Estado = 1
        Case "OCUPADA"
            Cod_Estado = 2
        Case "RESERVADA"
            Cod_Estado = 4
        
        End Select
        
        
        COD_CLIENTE = RsCustodiaCajas!IDCliente + 1000
        cod_ubicacion_prov = RsCustodiaCajas!Ubicacion
        
        
        Set rsbasa = New ADODB.Recordset
        SqlBasa = " SELECT  ID_CONTENEDOR , ESTADO "
        SqlBasa = SqlBasa & " From CONTENEDOR "
        SqlBasa = SqlBasa & " Where COD_CLIENTE = " & COD_CLIENTE
        SqlBasa = SqlBasa & " And NRO_CAJA = " & rsContenedor25!NRO_CAJA
        rsbasa.Open SqlBasa, strConBasa
        
         If Not rsbasa.EOF Then
            cod_contenedor = rsbasa!ID_CONTENEDOR
         End If
         cod_caja = rsContenedor25!NRO_CAJA
    
    Else
      Sql = " UPDATE CONTENEDOR25 SET ENBASA = 2 "
    Sql = Sql & vbCrLf & ", cod_Cliente = 0"
    Sql = Sql & " WHERE ID_CONTENEDOR =" & rsContenedor25!ID_CONTENEDOR
     ConContenedor25.Execute Sql
    
   GoTo 10
    
    End If
    
    
    Else
        
        
        If IsNull(rsContenedor25!NRO_CAJA) Then
        GoTo 10
     End If
        
        Set rsbasa = New ADODB.Recordset
        SqlBasa = " SELECT  ID_CONTENEDOR , ESTADO , COD_CLIENTE "
        SqlBasa = SqlBasa & " From CONTENEDOR "
        SqlBasa = SqlBasa & " Where "
        SqlBasa = SqlBasa & "  NRO_CAJA = " & rsContenedor25!NRO_CAJA
        If rsContenedor25!NRO_CAJA < 100000 Then
            SqlBasa = SqlBasa & " and  cod_Cliente = " & rsContenedor25!COD_CLIENTE
        End If
        rsbasa.Open SqlBasa, strConBasa
        If Not rsbasa.EOF Then
            Cod_Estado = rsbasa!estado
            cod_contenedor = rsbasa!ID_CONTENEDOR
            COD_CLIENTE = rsbasa!COD_CLIENTE
            cod_ubicacion_prov = ""
         End If
         cod_caja = rsContenedor25!NRO_CAJA
        
    

    End If

        
        
       
            Sql = "INSERT INTO CAMBIOPOSICION "
            Sql = Sql & vbCrLf & " (ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA, FECHA, ID_PERSONAL)"
            Sql = Sql & vbCrLf & " SELECT ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA, '10/11/2011' AS FECHA, 99 AS PERSONAL"
            Sql = Sql & vbCrLf & " From CONTENEDOR "
            
            Sql = Sql & vbCrLf & " Where NRO_CAJA = " & cod_caja
            
            If cod_caja < 100000 Then
                 Sql = Sql & vbCrLf & " And  COD_CLIENTE = " & COD_CLIENTE
    End If
           ExecutarSql Sql

            Sql = " UPDATE CONTENEDOR SET ESTADO=" & 1 & ","
            Sql = Sql & vbCrLf & " COD_CLIENTE=NULL, NRO_CAJA=NULL,"
            Sql = Sql & vbCrLf & " NRO_REMITO=NULL, F_MODIFICACION=NULL,"
            Sql = Sql & vbCrLf & " IDREQUERIMIENTO=NULL ,UB_PROVISORIA=NULL"
            Sql = Sql & vbCrLf & " Where NRO_CAJA = " & cod_caja
            If cod_caja < 100000 Then
                Sql = Sql & vbCrLf & " And COD_CLIENTE = " & COD_CLIENTE
            End If
           ExecutarSql Sql

            Sql = " Update Cajas "
            Sql = Sql & vbCrLf & " SET  DEPOSITO = 'ALSINA'"
            Sql = Sql & vbCrLf & " Where NRO_CAJA = " & cod_caja
            If cod_caja < 100000 Then
                Sql = Sql & vbCrLf & " And  FK_CLIENTE = " & COD_CLIENTE
            End If
            ExecutarSql Sql

            Sql = " Update CONTENEDOR "
            Sql = Sql & vbCrLf & " SET ESTADO =" & Cod_Estado
            If cod_ubicacion_prov <> "" Then
                Sql = Sql & vbCrLf & " , UB_PROVISORIA ='" & Trim(cod_ubicacion_prov) & "'"
            End If
            Sql = Sql & vbCrLf & " , COD_CLIENTE =" & COD_CLIENTE
            Sql = Sql & vbCrLf & ", NRO_CAJA =" & cod_caja
            Sql = Sql & vbCrLf & "  Where ID_CONTENEDOR = " & rsContenedor25!ID_CONTENEDOR
           ExecutarSql Sql


    Sql = " UPDATE CONTENEDOR25 SET ENBASA = 1 "
    Sql = Sql & vbCrLf & ", cod_Cliente = " & COD_CLIENTE
    Sql = Sql & " WHERE ID_CONTENEDOR =" & rsContenedor25!ID_CONTENEDOR
     ConContenedor25.Execute Sql
10:
        



    rsContenedor25.MoveNext
Loop



End Sub

Private Sub Command56_Click()
'
'
'
'
'DIA_DESDE_1
'
'SELECT Referencias.[Remote_User], Referencias.[Time_Stamp], Referencias.[Suspense_File], Referencias.[Remote_Uid], Referencias.[Remote_Fax], Referencias.[Remote_Bid], Referencias.[Remote_Cmp], Referencias.[Remote_Phn], Referencias.[CSID], Referencias.[Verify_Wks], Referencias.[Form_Id], Referencias.[BatchNo], Referencias.[BatchDir], Referencias.[BatchPgNo], Referencias.[BatchPgCnt], Referencias.[BatchRDate], Referencias.[BatchScOpr], Referencias.[BatchTrack], Referencias.[Route_To], Referencias.[Image_Seq], Referencias.[BatchPgDta], Referencias.[Form_Notes], Referencias.[CAJA_Nº], Referencias.[INDICE_1], Referencias.[DIA_HASTA_4], Referencias.[MES_HASTA_4], Referencias.[AÑO_HASTA_4], Referencias.[NUMERO_HASTA_4], Referencias.[DIA_DESDE_4], Referencias.[MES_DESDE_4], Referencias.[AÑO_DESDE_4], Referencias.[NUMERO_DESDE_4], Referencias.[DIA_DESDE_3], Referencias.[MES_DEDE_3], Referencias.[AÑO_DESDE_3], Referencias.[NUMERO_DESDE_3], Referencias.[DIA_HASTA_3], Referencias.[MES_HASTA_3], Referencias.[AÑO_HASTA_3
', Referencias.[NUMERO_HASTA_3], Referencias.[MES_DESDE_2], Referencias.[AÑO_DESDE_2], Referencias.[NUMERO_DESDE_2], Referencias.[DIA_HASTA_2], Referencias.[MES_HASTA_2], Referencias.[AÑO_HASTA_2], Referencias.[NUMERO_HASTA_2], Referencias.[DIA_DESDE_1], Referencias.[MES_DEDE_1], Referencias.[AÑO_DESDE_1], Referencias.[NUMERO_DESDE_1], Referencias.[DIA_HASTA_1], Referencias.[MES_HASTA_1], Referencias.[AÑO_HASTA_1], Referencias.[NUMERO_HASTA_1], Referencias.[INDICE_2], Referencias.[INDICE_3], Referencias.[INDICE_4], Referencias.[PCX_DESCRIPCION_1], Referencias.[DESCRIPCION_1], Referencias.[PCX_DESCRIPCION_2], Referencias.[DESCRIPCION_2], Referencias.[PCX_DESCRIPCION_3], Referencias.[DESCRIPCION_3], Referencias.[PCX_DESCRIPCION_4], Referencias.[DESCRIPCION_4], Referencias.[IDEM_INDICE_1], Referencias.[IDEM_DETALLE_1], Referencias.[IDEM_INDICE_2], Referencias.[IDEM_INDICE_3], Referencias.[IDEM_DETALLE_3], Referencias.[IDEM_DETALLE_2], Referencias.[ENVIO_CAJAS], Referencias.[USUARIO], Referencias.[DIA_DESDE_2]
'FROM Referencias;

End Sub


Private Sub Command57_Click()
            
       Dim comPlanilla As New ADODB.Connection
            Dim RsPlanilla As New ADODB.Recordset
            Dim Sql As String
            Dim i As Integer
            Dim FECHA_DESDE(4) As String
            Dim FECHA_HASTA(4) As String
            Dim NUMERO_DESDE(4) As String
            Dim NUMERO_HASTA(4) As String
            Dim Indice(4) As String
            Dim Descripcion(4) As String
            comPlanilla.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=\\server-basa\Sistemas\BaseTeleform\Referencias.mdb"
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet
    Dim C_Error As Integer
    Dim C_Caja As Integer
    Dim C_Indice As Integer
    Dim C_Etiqueta As Integer
    Dim C_Fecha_desde As Integer
    Dim C_Fecha_hasta As Integer
    Dim C_N°_Desde As Integer
    Dim C_N°_Hasta As Integer
    Dim C_Letra_Desde As Integer
    Dim C_Letra_Hasta As Integer
    Dim C_Descripcion As Integer
    Dim R As String
    Dim ErrorGeneral As Boolean
    Dim strError As String
    Dim FechaHora As String
    Dim NombreArchivo As String
    
    FechaHora = Trim(Format(Now, "hhmmss"))
    
 
    C_Error = 1
    C_Caja = 2
    C_Indice = 4
    C_Etiqueta = 3
    C_Fecha_desde = 6
    C_Fecha_hasta = 7
    C_N°_Desde = 8
    C_N°_Hasta = 9
    C_Letra_Desde = 10
    C_Letra_Hasta = 11
    C_Descripcion = 5
    
    
    
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Open("\\Server-basa\Sistemas\Referencias\Planilla Modelo.xls", , True)
    Set hojaEx = libroEx.Worksheets.Item(1)
   
            'SELECT Referencias.[Remote_User], Referencias.[Time_Stamp], Referencias.[Suspense_File], Referencias.[Remote_Uid], Referencias.[Remote_Fax], Referencias.[Remote_Bid], Referencias.[Remote_Cmp], Referencias.[Remote_Phn], Referencias.[CSID], Referencias.[Verify_Wks], Referencias.[Form_Id], Referencias.[BatchNo], Referencias.[BatchDir], Referencias.[BatchPgNo], Referencias.[BatchPgCnt], Referencias.[BatchRDate], Referencias.[BatchScOpr], Referencias.[BatchTrack], Referencias.[Route_To], Referencias.[Image_Seq], Referencias.[BatchPgDta], Referencias.[Form_Notes], Referencias.[CAJA_Nº], Referencias.[INDICE_1], Referencias.[DIA_HASTA_4], Referencias.[MES_HASTA_4], Referencias.[AÑO_HASTA_4], Referencias.[NUMERO_HASTA_4], Referencias.[DIA_DESDE_4], Referencias.[MES_DESDE_4], Referencias.[AÑO_DESDE_4], Referencias.[NUMERO_DESDE_4], Referencias.[DIA_DESDE_3], Referencias.[MES_DEDE_3], Referencias.[AÑO_DESDE_3], Referencias.[NUMERO_DESDE_3], Referencias.[DIA_HASTA_3], Referencias.[MES_HASTA_3], Referencias
            ', Referencias.[NUMERO_HASTA_3], Referencias.[MES_DESDE_2], Referencias.[AÑO_DESDE_2], Referencias.[NUMERO_DESDE_2], Referencias.[DIA_HASTA_2], Referencias.[MES_HASTA_2], Referencias.[AÑO_HASTA_2], Referencias.[NUMERO_HASTA_2], Referencias.[DIA_DESDE_1], Referencias.[MES_DEDE_1], Referencias.[AÑO_DESDE_1], Referencias.[NUMERO_DESDE_1], Referencias.[DIA_HASTA_1], Referencias.[MES_HASTA_1], Referencias.[AÑO_HASTA_1], Referencias.[NUMERO_HASTA_1], Referencias.[INDICE_2], Referencias.[INDICE_3], Referencias.[INDICE_4], Referencias.[PCX_DESCRIPCION_1], Referencias.[DESCRIPCION_1], Referencias.[PCX_DESCRIPCION_2], Referencias.[DESCRIPCION_2], Referencias.[PCX_DESCRIPCION_3], Referencias.[DESCRIPCION_3], Referencias.[PCX_DESCRIPCION_4], Referencias.[DESCRIPCION_4], Referencias.[IDEM_INDICE_1], Referencias.[IDEM_DETALLE_1], Referencias.[IDEM_INDICE_2], Referencias.[IDEM_INDICE_3], Referencias.[IDEM_DETALLE_3], Referencias.[IDEM_DETALLE_2], Referencias.[ENVIO_CAJAS], Referencias.[USUARIO], Referencias.[D
            'FROM Referencias;
            
            Sql = " Select * from Referencias "
            Sql = Sql & " WHERE (((Referencias.[Suspense_File]) Like '%" & InputBox("Ingrese el numero de lote") & "\%'));"
            RsPlanilla.Open Sql, comPlanilla
            R = 7
            Do While Not RsPlanilla.EOF
                For i = 1 To 4
                    If Not IsNull(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))) Then
                        Rem MsgBox RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i))
                        FECHA_DESDE(i) = Format(Format(RsPlanilla.Fields.Item("DIA_DESDE_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("MES_DESDE_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("AÑO_HASTA_" & CStr(i)), "00"), "DD/MM/YYYY")
                        If Format(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)), "00") <> "00" Then
                            FECHA_HASTA(i) = Format(Format(RsPlanilla.Fields.Item("DIA_HASTA_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("MES_HASTA_" & CStr(i)), "00") & "/" & Format(RsPlanilla.Fields.Item("AÑO_HASTA_" & CStr(i)), "00"), "DD/MM/YYYY")
                        Else
                            FECHA_HASTA(i) = FECHA_DESDE(i)
                        End If
                    End If
                    If Not IsNull(RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))) Then
                        MsgBox RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))
                        NUMERO_DESDE(i) = RsPlanilla.Fields.Item("NUMERO_DESDE_" & CStr(i))
                        If Format(RsPlanilla.Fields.Item("NUMERO_HASTA_" & CStr(i)), "") <> "" Then
                            NUMERO_HASTA(i) = RsPlanilla.Fields.Item("NUMERO_HASTA_" & i)
                        Else
                             NUMERO_DESDE(i) = NUMERO_HASTA(i)
                        End If
                    End If
                    If Not IsNull(RsPlanilla.Fields.Item("INDICE_" & CStr(i))) Then
                        Indice(i) = RsPlanilla.Fields.Item("INDICE_" & CStr(i))
                        Else
                            If i <> 1 Then
                                If Not IsNull(RsPlanilla.Fields.Item("IDEM_INDICE_" & CStr(i))) Then
                                    Indice(i) = Indice(CStr(i - 1))
                                End If
                            Else
                                Indice(i) = 0
                            End If
                    End If
                
                    If Not IsNull(RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))) Then
                        Descripcion(i) = RsPlanilla.Fields.Item("DESCRIPCION_" & CStr(i))
                    Else
                            If i <> 1 Then
                                If Not IsNull(RsPlanilla.Fields.Item("IDEM_DETALLE_" & CStr(i))) Then
                                    Descripcion(i) = Descripcion(CStr(i - 1))
                                End If
                            Else
                                Descripcion(i) = ""
                            End If
                    End If
                
                
                
                
                Next
                
                
                NombreArchivo = RsPlanilla.Fields.Item("CAJA_Nº").value & "_" & FechaHora & ".tif"
                        

              For i = 1 To 4
                
                    If Indice(i) <> "" Then
                        hojaEx.Cells(R, C_Caja) = RsPlanilla.Fields.Item("CAJA_Nº").value
                        hojaEx.Cells(R, C_Caja).Hyperlinks.Add hojaEx.Cells(R, C_Caja), ".\Cajas\" & Trim(RsPlanilla.Fields.Item("CAJA_Nº").value) & "\" & NombreArchivo
                        hojaEx.Cells(R, C_Indice) = Indice(i)
                        hojaEx.Cells(R, C_Descripcion) = Descripcion(i)
                        hojaEx.Cells(R, C_Fecha_desde) = FECHA_DESDE(i)
                        hojaEx.Cells(R, C_Fecha_hasta) = FECHA_HASTA(i)
                        hojaEx.Cells(R, C_N°_Desde) = NUMERO_DESDE(i)
                        hojaEx.Cells(R, C_N°_Hasta) = NUMERO_HASTA(i)
                        R = R + 1
                    End If
                Next
                
                If Dir(RsPlanilla.Fields.Item("Suspense_File").value) <> "" Then
                
                If Dir("\\SERVER-BASA\Sistemas\Referencias\cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value, vbDirectory) = "" Then
                    MkDir "\\SERVER-BASA\Sistemas\Referencias\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value
                    FileCopy RsPlanilla.Fields.Item("Suspense_File").value, "\\SERVER-BASA\Sistemas\Referencias" & "\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value & "\" & NombreArchivo
                 Else
                    FileCopy RsPlanilla.Fields.Item("Suspense_File").value, "\\SERVER-BASA\Sistemas\Referencias" & "\Cajas\" & RsPlanilla.Fields.Item("CAJA_Nº").value & "\" & NombreArchivo
                End If
                
                Else
                    MsgBox "No se encontro La imagen " & RsPlanilla.Fields.Item("Suspense_File").value
                End If
                
                 RsPlanilla.MoveNext
        Loop
                libroEx.SaveAs "\\SERVER-BASA\Sistemas\Referencias\" & InputBox("Ingrese el nombre de la planilla") & Format(Now, "ddmmyyy hhss") & ".xls"
                libroEx.Close
                ApExcel.Quit
                Set hojaEx = Nothing
                Set libroEx = Nothing
                Set ApExcel = Nothing
                
                MsgBox "Terminado"
            

End Sub

Private Sub Command58_Click()

Dim rsbasa As New ADODB.Recordset
Dim rscajascus As New ADODB.Recordset
Dim Sql As String

Sql = " SELECT     ID_CONTENEDOR,    NRO_CAJA, COD_CLIENTE, UB_PROVISORIA"
Sql = Sql & " From CONTENEDOR"
Sql = Sql & " WHERE        (NRO_CAJA BETWEEN 100000 AND 400000) AND (COD_CLIENTE > 250) AND (UB_PROVISORIA IS NULL)"


rsbasa.CursorLocation = adUseClient
rsbasa.Open Sql, "Provider=SQLOLEDB.1;Password=21877471; Persist Security Info=False;User ID=sa;Initial Catalog=BASE2;Data Source=SERVER-BASA", 2, 3

Dim concus As New ADODB.Connection

concus.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Datos Custodia\DatosCustodia.accdb;Persist Security Info=False"

 Do While Not rsbasa.EOF
    
    
    
    Set rscajascus = New ADODB.Recordset
    Sql = " SELECT IDCliente, IDCaja, Ubicacion,Estado "
        Sql = Sql & " FROM CAJAS1  "
        Sql = Sql & " WHERE IDCaja = " & rsbasa!NRO_CAJA

rscajascus.Open Sql, concus
    If Not rscajascus.EOF Then
        rsbasa!UB_PROVISORIA = "'" & rscajascus!Ubicacion & "'"
        rsbasa.Update
    End If
    rsbasa.MoveNext
 Loop
 


End Sub

Private Sub Command59_Click()

    Dim Sql As String
    Dim rs  As New ADODB.Recordset
    


    Sql = " SELECT     COD_CLIENTE, INDICE, DESCRIPCION, LEN(INDICE) AS Expr1, ID"
    Sql = Sql & " From INDICES"
    Sql = Sql & " WHERE     (COD_CLIENTE = 1197) "
    Sql = Sql & " AND (INDICE LIKE '002%')"
    Sql = Sql & " AND (LEN(INDICE) = 9)"
    
    rs.Open Sql, strConBasa
     Dim idusu As Integer
     idusu = 4022
    Do While Not rs.EOF
    
    
    Sql = " INSERT INTO CLIENTEUSUARIO"
    Sql = Sql & "                 (ID_CLIENTEUSUARIO, COD_CLIENTE, APELLIDO_NOMBRE, COD_INDICE)"
Sql = Sql & "  VALUES     (" & idusu & ",1197,'" & Trim(Mid(rs!Descripcion, 3, 5)) & "','" & Trim(rs!Indice) & "')"
ExecutarSql Sql
idusu = idusu + 1

        rs.MoveNext
    Loop
    

End Sub

Private Sub Command60_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset


Sql = " SELECT      ID , CAJA From CAJAS99 ORDER BY ID"

rs.Open Sql, ConActiva, 0, 1


Do While Not rs.EOF

Sql = " Update ORDENAR_DOCUMENTACION_DETALLE"
Sql = Sql & " SET  COD_NRO_CAJA =" & rs!Caja
Sql = Sql & " Where COD_CLIENTE = 99 "
Sql = Sql & " And ID = " & rs!ID
 ExecutarSql Sql
 rs.MoveNext
Loop



End Sub

Private Sub Command61_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String


Sql = " SELECT     ID_CLIENTEUSUARIO, COD_CLIENTE, SM, CANTIDAD "
Sql = Sql & " From DISCOCANTIDADES "
Sql = Sql & "  where  (ID_CLIENTEUSUARIO = 4251)"
Sql = Sql & " ORDER BY ID_CLIENTEUSUARIO "

rs.Open Sql, strConBasa


    Dim Operacion As String
    Dim estado As String
    Dim fecha As String
    Dim id_cliente As String
    Dim AUDIT_USUARIO As String
    Dim AUDIT_FECHA As String
    Dim COD_TIPO_ALMACENAMIENTO As String
    Dim COD_PERSONAL_ENTREGA As String
    Dim COD_USUARIO_CLIENTE As Integer
    Dim NRO_REMITO As Long
    Dim cantidad As Long
            
          
    
            



Do While Not rs.EOF

            Operacion = 0
            estado = 0
            fecha = "'01/01/2012'"
            id_cliente = 1197
            AUDIT_USUARIO = 17
            AUDIT_FECHA = "'01/01/2012'"
            COD_TIPO_ALMACENAMIENTO = 0
            COD_PERSONAL_ENTREGA = 17
            COD_USUARIO_CLIENTE = rs!ID_CLIENTEUSUARIO
            NRO_REMITO = ProximoRemito
        cantidad = rs!cantidad



             Sql = " INSERT INTO REMITOS_CUERPO"
            Sql = Sql & vbCrLf & " (NRO_REMITO, NRO_REM_PROV, TIPO, OPERACION, ESTADO,FECHA, ID_CLIENTE, OBSERVACIONES, CANTIDAD,"
            Sql = Sql & vbCrLf & " AUDIT_USUARIO, AUDIT_FECHA,COD_TIPO_ALMACENAMIENTO, COD_PERSONAL_ENTREGA,COD_USUARIO_CLIENTE , COBRAR_FLETE )"
            Sql = Sql & vbCrLf & " VALUES (" & NRO_REMITO & ",'0000000',0," & Operacion & "," & estado & ","
            Sql = Sql & vbCrLf & fecha & "," & id_cliente & ",'Remito de suma de cantidad de cajas de custodia'," & 19401
            Sql = Sql & vbCrLf & ",'" & AUDIT_USUARIO & "'," & AUDIT_FECHA & "," & COD_TIPO_ALMACENAMIENTO & "," & COD_PERSONAL_ENTREGA & "," & COD_USUARIO_CLIENTE & " ,'1' )"
            ExecutarSql Sql
    

    rs.MoveNext
Loop



End Sub

Private Sub Command62_Click()

Dim MyName As String
Dim ID As Long
Dim NRO_CAJA As Long




            MyName = Dir("C:\DatCus\Datas\*.tps", vbDirectory)
            
             
             Do While MyName <> ""
                    If Len(MyName) = 12 Then
                   
                   FileSystem.FileCopy "C:\DatCus\Datas\" & MyName, "C:\DatCus\TPS\" & UCase(MyName)
                   End If
                   MyName = Dir()
            Loop





End Sub

Private Sub Command63_Click()
Dim MyName As String
Dim ID As Long
Dim NRO_CAJA As Long
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim CLIENTE_CUSTODIA, IDSUCURSAL, NOMBRESUCURSAL, IDTIPODOCUMENTO, NOMBRETIPODOCUMENTO, cantidad, CLIENTE_BASA As String

con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\DatCus\referencias.accdb;Persist Security Info=False"


            MyName = Dir("C:\DatCus\TPSTer\DCT*.tps", vbDirectory)
            
             
             Do While MyName <> ""
             
             
             Set rs = New ADODB.Recordset
             
                 Sql = " SELECT IDTIPODOCUMENTO, NOMBRETIPODOCUMENTO, IDSUCURSAL, NOMBRESUCURSAL, Count(*) AS cantidad"
                Sql = Sql & " From " & Mid(MyName, 1, Len(MyName) - 4)
                Sql = Sql & " GROUP BY IDTIPODOCUMENTO, NOMBRETIPODOCUMENTO, IDSUCURSAL, NOMBRESUCURSAL;"
                rs.Open Sql, con
                Do While Not rs.EOF
                
                CLIENTE_CUSTODIA = Mid(MyName, 5, 4)
                If Not IsNull(rs!IDSUCURSAL) Then
                    IDSUCURSAL = rs!IDSUCURSAL
                Else
                    IDSUCURSAL = "NULL"
                End If
                If IsNull(rs!NOMBRESUCURSAL) Then
                    NOMBRESUCURSAL = "Null"
                Else
                    NOMBRESUCURSAL = "'" & rs!NOMBRESUCURSAL & "'"
                End If
                
                
                If IsNull(rs!IDTIPODOCUMENTO) Then
                    IDTIPODOCUMENTO = "Null"
                Else
                    IDTIPODOCUMENTO = rs!IDTIPODOCUMENTO
                End If
                
                If IsNull(rs!NOMBRETIPODOCUMENTO) Then
                    NOMBRETIPODOCUMENTO = "Null"
                Else
                    NOMBRETIPODOCUMENTO = "'" & rs!NOMBRETIPODOCUMENTO & "'"
                End If
                
               If IsNull(rs!cantidad) Then
                    cantidad = "Null"
                Else
                    cantidad = rs!cantidad
                End If
                
                CLIENTE_BASA = CLIENTE_CUSTODIA + 1000
                
                
                Sql = "INSERT INTO INTERCAMBIO"
                Sql = Sql & "(CLIENTE_CUSTODIA, IDSUCURSAL, NOMBRESUCURSAL, IDTIPODOCUMENTO, NOMBRETIPODOCUMENTO, CANTIDAD, CLIENTE_BASA)"
                Sql = Sql & " VALUES     (" & CLIENTE_CUSTODIA & "," & IDSUCURSAL & "," & NOMBRESUCURSAL & "," & IDTIPODOCUMENTO & "," & NOMBRETIPODOCUMENTO & "," & cantidad & "," & CLIENTE_BASA & ")"
                ExecutarSql Sql
                rs.MoveNext
                Loop
                
                
                   MyName = Dir()
            Loop





End Sub

Private Sub Command64_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String
Dim Sql2 As String


Set rs = New ADODB.Recordset

Sql = " SELECT     COD_ID_REFERENCIA, COPIADO"
Sql = Sql & " From ID_referencia"
Sql = Sql & "  Where (COPIADO Is Null)"
Sql = Sql & " ORDER BY COD_ID_REFERENCIA "

rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
Sql2 = "  INSERT INTO REFERENCIAS02042012"
 Sql2 = Sql2 & "                     (COD_CLIENTE, NRO_CAJA, FK_CAJA, COD_TIPO_ALMACENAMIENTO, ITEM, INDICE, COD_DOCUMENTO, DESCRIPCION, FECHA_DESDE, FECHA_HASTA,"
 Sql2 = Sql2 & "                      NRO_DESDE, NRO_HASTA, LETRA_DESDE, LETRA_HASTA, EXPEDIENTE, APELLIDO_NOMBRE, FECHA_MODIFICACION, FECHA_CREACION,"
 Sql2 = Sql2 & "                      USUARIO_MODIFICACION, FK_PERSONAL_CREACION, FK_PERSONAL_MODIFICACION, BORRADO, ID_UNITER, ESTADO, CONTROL, INDICE_ANTERIOR, SECTOR,"
 Sql2 = Sql2 & "                      PASOARCHIVO, ID_IMAGEN, CONTROLEXCEL, COD_ESTADO_DOCUMENTO, PLANILLA)"
Sql2 = Sql2 & "  SELECT     COD_CLIENTE, NRO_CAJA, FK_CAJA, COD_TIPO_ALMACENAMIENTO, ITEM, INDICE, COD_DOCUMENTO, DESCRIPCION, FECHA_DESDE, FECHA_HASTA,"
                      Sql2 = Sql2 & " NRO_DESDE, NRO_HASTA, LETRA_DESDE, LETRA_HASTA, EXPEDIENTE, APELLIDO_NOMBRE, FECHA_MODIFICACION, FECHA_CREACION,"
                      Sql2 = Sql2 & "  USUARIO_MODIFICACION, FK_PERSONAL_CREACION, FK_PERSONAL_MODIFICACION, BORRADO, ID_UNITER, ESTADO, CONTROL, INDICE_ANTERIOR, SECTOR,"
                      Sql2 = Sql2 & " PASOARCHIVO , ID_imagen, ControlExcel, COD_ESTADO_DOCUMENTO, PLANILLA"
Sql2 = Sql2 & " From REFERENCIAS"
Sql2 = Sql2 & "  Where COD_ID_REFERENCIA =" & rs!COD_ID_REFERENCIA
ExecutarSql Sql

Sql = " Update ID_referencia"
Sql = Sql & " Set COPIADO = 1"
Sql = Sql & "  Where COD_ID_REFERENCIA = " & rs!COD_ID_REFERENCIA
ExecutarSql Sql
    rs.MoveNext
Loop



End Sub

Private Sub Command65_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Indice As String


Sql = " SELECT     REFERENCIAS.COD_ID_REFERENCIA, REFERENCIAS.COD_CLIENTE, REFERENCIAS.NRO_CAJA, REFERENCIAS.DESCRIPCION, REFERENCIAS.FK_INDICES,"
                      Sql = Sql & " REFERENCIAS.INDICE, INDICES.DESCRIPCION AS Expr1, INDICES.ID_CODIGO_DOCUMENTO"
Sql = Sql & " FROM         REFERENCIAS INNER JOIN"
              Sql = Sql & "         INDICES ON REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND REFERENCIAS.INDICE = INDICES.INDICE"
Sql = Sql & "  WHERE     (REFERENCIAS.COD_CLIENTE = 1197) AND (REFERENCIAS.DESCRIPCION LIKE '%PLANILLAS DE VENTAS%') AND (INDICES.DESCRIPCION LIKE '%rollo%')"

rs.Open Sql, ConActiva, adOpenStatic, adLockReadOnly


Do While Not rs.EOF

    Indice = "'" & BuscarIndice(1197, CLng("3" & CStr(Mid(rs!ID_CODIGO_DOCUMENTO, 2)))) & "'"


    Sql = " Update REFERENCIAS "
    Sql = Sql & " SET    INDICE =" & Indice
    Sql = Sql & " Where COD_ID_REFERENCIA = " & rs!COD_ID_REFERENCIA
    
    ExecutarSql Sql
        
    rs.MoveNext


Loop







End Sub

Private Sub Command66_Click()

Dim rs As New ADODB.Recordset
Dim RSconrenedor As New ADODB.Recordset
Dim Sql As String
Sql = " SELECT     CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA AS Expr1, CAJAS.FK_ESTADO, CAJAS.FK_REMITO_BAJA"
Sql = Sql & " FROM         CAJAS LEFT OUTER JOIN"
Sql = Sql & "   CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA"
Sql = Sql & " Where (CONTENEDOR.NRO_CAJA Is Null) And (Not (Cajas.FK_CLIENTE Is Null)) And (Cajas.FK_ESTADO <> 1140)"



Sql = " SELECT     CAJAS.FK_CLIENTE, CAJAS.ID_CAJA, CAJAS.NRO_CAJA, CAJAS.FK_ESTADO, CONTENEDOR.COD_CLIENTE"
Sql = Sql & " FROM         CAJAS LEFT OUTER JOIN"
  Sql = Sql & "                     CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA"
Sql = Sql & " Where (Cajas.FK_ESTADO = 1120) And (CONTENEDOR.COD_CLIENTE Is Null)"

rs.Open Sql, ConActiva, adOpenForwardOnly, adLockReadOnly

Sql = "SELECT     TOP 3500 ID_CONTENEDOR, ESTANTERIA, ESTADO, COD_CLIENTE, NRO_CAJA"
Sql = Sql & " From CONTENEDOR "
Sql = Sql & " Where (Estanteria > 149) And (estado = 1)"
Sql = Sql & " ORDER BY ESTANTERIA"
RSconrenedor.CursorLocation = adUseClient

RSconrenedor.Open Sql, ConActiva, 2, adLockBatchOptimistic



Do While Not rs.EOF
    RSconrenedor!estado = 2
    RSconrenedor!COD_CLIENTE = rs!FK_CLIENTE
    RSconrenedor!NRO_CAJA = rs!NRO_CAJA
    RSconrenedor.Update
    
    
    RSconrenedor.MoveNext
    rs.MoveNext
Loop





End Sub

Private Sub Command67_Click()

Dim i As Double

Dim Sql As String
Dim con As New ADODB.Connection
con.Open strConBasa

For i = 1152776 To 4621161
Sql = "INSERT INTO LEGAJOSOK"
       Sql = Sql & "               (ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE, FK_INDICES, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA,"
       Sql = Sql & "                CLIENTE_LEGAJO, DESCRIPCION, NRO_CAJA, REARCHIVO_CAJA, COD_CLIENTE, COD_UBICACION, COD_ESTADO, NOMBRE, COD_REMITO, FECHA, ID_PERSONAL,"
       Sql = Sql & "                FK_PERSONAL_CREACION, FECHA_CREACION, FK_PERSONAL_ACTUALIZACION, FECHA_ACTUALIZACION, ID_LEGAJO_ECOGAS, ID_CLIENTE_BASE, ERRORTIPEO,"
       Sql = Sql & "                CARGAPAGADA, NRO_REM_PROV, ORDEN, NUMERO_LEGAJO_CLIENTE, FECHAPAGO, PEGADOETIQUETA, CANTIDAD_CARACTERES, CONTROL_EXPORT,"
       Sql = Sql & "                DESCRIPCION_REMITO, PASOARCHIVO , DIGITO_VERIFICADOR, REGISTRO_VERIFICADO, CONTROL_PADRON, FK_PERSONAL_ASIGNACION, ROLLO,"
        Sql = Sql & "               INDICE_ANTERIOR, ID_CUSTODIA, COD_CLIENTE_CUSTODIA)"
Sql = Sql & " SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE, FK_INDICES, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA,"
 Sql = Sql & "                      CLIENTE_LEGAJO, DESCRIPCION, NRO_CAJA, REARCHIVO_CAJA, COD_CLIENTE, COD_UBICACION, COD_ESTADO, NOMBRE, COD_REMITO, FECHA, ID_PERSONAL,"
 Sql = Sql & "                      FK_PERSONAL_CREACION, FECHA_CREACION, FK_PERSONAL_ACTUALIZACION, FECHA_ACTUALIZACION, ID_LEGAJO_ECOGAS, ID_CLIENTE_BASE, ERRORTIPEO,"
 Sql = Sql & "                      CARGAPAGADA, NRO_REM_PROV, ORDEN, NUMERO_LEGAJO_CLIENTE, FECHAPAGO, PEGADOETIQUETA, CANTIDAD_CARACTERES, CONTROL_EXPORT,"
 Sql = Sql & "                      DESCRIPCION_REMITO, PASOARCHIVO, DIGITO_VERIFICADOR, REGISTRO_VERIFICADO, CONTROL_PADRON, FK_PERSONAL_ASIGNACION, ROLLO,"
  Sql = Sql & "                     Indice_Anterior , ID_CUSTODIA, COD_CLIENTE_CUSTODIA"
Sql = Sql & "  From LEGAJOS"
Sql = Sql & " Where ID_LEGAJO = " & i

If ExecutarSql(Sql) = 0 Then
    Debug.Print i
End If


Next

End Sub


Private Sub Command68_Click()
Dim Sql As String
Dim rsCajas As New ADODB.Recordset
Dim RsCambioCliente As New ADODB.Recordset


Set rsCajas = New ADODB.Recordset


Sql = " SELECT     CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA AS Expr1, CAJAS.FK_REMITO_BAJA"
Sql = Sql & " FROM         CAJAS LEFT OUTER JOIN"
 Sql = Sql & "                      CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA"
Sql = Sql & " Where (CONTENEDOR.NRO_CAJA Is Null) And (Cajas.FK_CLIENTE > 1000) And (Cajas.FK_REMITO_BAJA Is Null)"

'
'
'SQL = " SELECT     CAJAS_FINAL_REFERENCIAS.COD_CLIENTE, CAJAS_FINAL_REFERENCIAS.NRO_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA AS Expr1"
'SQL = SQL & " FROM         CAJAS_FINAL_REFERENCIAS INNER JOIN"
'SQL = SQL & " CAJAS ON CAJAS_FINAL_REFERENCIAS.NRO_CAJA = CAJAS.ID_CAJA"
'SQL = SQL & " AND CAJAS_FINAL_REFERENCIAS.COD_CLIENTE <> CAJAS.FK_CLIENTE"
'SQL = SQL & " ORDER BY CAJAS_FINAL_REFERENCIAS.NRO_CAJA"
'
'
'SQL = "  SELECT     * "
'SQL = SQL & " FROM  DIFERENCIACUSTODIA "
'
'

Sql = " SELECT     LECTURACOLECTOR.ID, LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN,"
Sql = Sql & "  CONTENEDOR.COD_CLIENTE"
Sql = Sql & "  FROM         LECTURACOLECTOR LEFT OUTER JOIN"
                      Sql = Sql & " CONTENEDOR ON LECTURACOLECTOR.CLIENTE = CONTENEDOR.COD_CLIENTE AND LECTURACOLECTOR.CAJA = CONTENEDOR.NRO_CAJA"
Sql = Sql & "  WHERE     (LECTURACOLECTOR.NUMERO_LECTURA IN (18122)) AND (CONTENEDOR.COD_CLIENTE IS NULL)"


RsCambioCliente.Open Sql, strConBasa
Dim concam As New ADODB.Connection
concam.Open strConBasa

Dim i As Long

Do While Not RsCambioCliente.EOF

'SQL = " Update Cajas "
'SQL = SQL & " Set FK_CLIENTE = " & RsCambioCliente!COD_CLIENTE
'SQL = SQL & " ,  FK_ESTADO =1002"
'SQL = SQL & "  Where ID_CAJA =  " & RsCambioCliente!NRO_CAJA
'concam.Execute SQL

Set rsCajas = New ADODB.Recordset



Sql = "  SELECT     TOP (1) COD_CLIENTE, ESTANTERIA, ID_CONTENEDOR, ESTADO"
Sql = Sql & " From CONTENEDOR"
Sql = Sql & " WHERE     (ESTANTERIA BETWEEN 150 AND 190) AND (ESTADO = 1) AND (COD_CLIENTE IS NULL)"
Sql = Sql & " ORDER BY ESTANTERIA"

rsCajas.Open Sql, strConBasa, 0, adLockReadOnly

Sql = " Update CONTENEDOR "
Sql = Sql & " Set COD_CLIENTE = " & RsCambioCliente!Cliente
Sql = Sql & " , NRO_CAJA = " & RsCambioCliente!Caja
Sql = Sql & " , ESTADO =2 "
Sql = Sql & " Where ID_CONTENEDOR = " & rsCajas!ID_CONTENEDOR
concam.Execute Sql

    

    RsCambioCliente.MoveNext
    
    i = i + 1
Loop

End Sub

Private Sub Command69_Click()



Dim Sql As String
Dim rsCajas As New ADODB.Recordset
Dim RsCambioCliente As New ADODB.Recordset


Set rsCajas = New ADODB.Recordset




Sql = " SELECT     CAJAS_FINAL_REFERENCIAS.COD_CLIENTE, CAJAS_FINAL_REFERENCIAS.NRO_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA AS Expr1"
Sql = Sql & " FROM         CAJAS_FINAL_REFERENCIAS INNER JOIN"
Sql = Sql & " CAJAS ON CAJAS_FINAL_REFERENCIAS.NRO_CAJA = CAJAS.ID_CAJA"
Sql = Sql & " AND CAJAS_FINAL_REFERENCIAS.COD_CLIENTE <> CAJAS.FK_CLIENTE"
Sql = Sql & " ORDER BY CAJAS_FINAL_REFERENCIAS.NRO_CAJA"


Sql = "  SELECT     * "
Sql = Sql & " FROM  DIFERENCIACUSTODIA "

Sql = " SELECT     CONTENEDOR.ESTADO, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_ESTADO, CAJAS.FK_REMITO_BAJA"
Sql = Sql & " FROM         CAJAS LEFT OUTER JOIN"
Sql = Sql & " CONTENEDOR ON CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA AND CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE"
Sql = Sql & " Where(CONTENEDOR.estado Is Null) And (Not (Cajas.FK_CLIENTE Is Null)) And (Cajas.FK_REMITO_BAJA Is Null)"










RsCambioCliente.Open Sql, strConBasa
Dim concam As New ADODB.Connection
concam.Open strConBasa
Rem concam.BeginTrans
Dim i As Long

Do While Not RsCambioCliente.EOF

'SQL = " Update Cajas "
'SQL = SQL & " Set FK_CLIENTE = " & RsCambioCliente!COD_CLIENTE
'SQL = SQL & " ,  FK_ESTADO =1002"
'SQL = SQL & "  Where ID_CAJA =  " & RsCambioCliente!NRO_CAJA
'concam.Execute SQL

Set rsCajas = New ADODB.Recordset



Sql = "  SELECT     TOP (1) COD_CLIENTE, ESTANTERIA, ID_CONTENEDOR, ESTADO"
Sql = Sql & " From CONTENEDOR"
Sql = Sql & " WHERE     (ESTANTERIA BETWEEN 150 AND 190) AND (ESTADO = 1) AND (COD_CLIENTE IS NULL)"
Sql = Sql & " ORDER BY ESTANTERIA"

rsCajas.Open Sql, strConBasa, 0, adLockReadOnly

Sql = " Update CONTENEDOR "
Sql = Sql & " Set COD_CLIENTE = " & RsCambioCliente!FK_CLIENTE
Sql = Sql & " , NRO_CAJA = " & RsCambioCliente!NRO_CAJA
Sql = Sql & " , ESTADO = 2 "
Sql = Sql & " Where ID_CONTENEDOR = " & rsCajas!ID_CONTENEDOR
concam.Execute Sql

    

    RsCambioCliente.MoveNext
    
    i = i + 1
Loop
Rem concam.CommitTrans






Sql = Sql & "  SELECT     CAJAS_FINAL_REFERENCIAS.COD_CLIENTE, CAJAS_FINAL_REFERENCIAS.NRO_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA AS Expr1"
Sql = Sql & " FROM         CAJAS_FINAL_REFERENCIAS INNER JOIN"
Sql = Sql & " CAJAS ON CAJAS_FINAL_REFERENCIAS.NRO_CAJA = CAJAS.ID_CAJA"
Sql = Sql & " Where (Cajas.FK_CLIENTE Is Null)"
Sql = Sql & " ORDER BY CAJAS_FINAL_REFERENCIAS.NRO_CAJA"



End Sub

Private Sub Command7_Click()

    Dim rsContenedor As New ADODB.Recordset
'     Dim rsCajas As New ADODB.Recordset
'   Dim i As Integer
'   Dim sql As String
'   Dim con As New ADODB.Connection
'
'   con.Open strConBasa , 0 ,1
'
'    sql = " SELECT     TOP 620 ESTANTERIA, ID_CONTENEDOR, NRO_CAJA, COD_CLIENTE, ESTADO"
'    sql = sql & " From dbo.CONTENEDOR"
'    sql = sql & " Where (Estanteria > 140) And (Estado = 1)"
'      sql = sql & " ORDER BY ESTANTERIA"
'
'    rsContenedor.Open sql, strConBasa , 0 ,1
'
'
'    sql = "SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_PERSONAL_ENTREGA, FECHA_IMPRESION, DIGITO_VERIFICADOR, FK_USUARIO_CREACION_CAJA"
'sql = sql & " From dbo.cajas"
'sql = sql & "  WHERE     (ID_CAJA BETWEEN 736021 AND 736621)"
'sql = sql & "  ORDER BY ID_CAJA"
'    rsCajas.Open sql, strConBasa , 0 ,1
'
'    For i = 13000 To 13600
'
'
'        sql = " Update dbo.CONTENEDOR"
'        sql = sql & "  SET  NRO_CAJA =" & i
'        sql = sql & "   , COD_CLIENTE =163"
'        sql = sql & "  , ESTADO =2, FK_CAJAS =" & rsCajas!ID_CAJA
'        sql = sql & "  Where ID_CONTENEDOR = " & rsContenedor!ID_CONTENEDOR
'
'    con.Execute sql
'       sql = " Update dbo.Cajas"
'sql = sql & " SET FK_CLIENTE =163"
'sql = sql & ", NRO_CAJA =" & i
'sql = sql & ", FK_ESTADO =1120"
'sql = sql & ", FK_PERSONAL_ENTREGA =38"
'sql = sql & " Where ID_CAJA = " & rsCajas!ID_CAJA
'con.Execute sql
'    rsCajas.MoveNext
'    rsContenedor.MoveNext
'
'    Next
    

End Sub

Private Sub Command70_Click()

'    Dim SQL As String
'    Dim RS As New ADODB.Recordset
'    Dim con As New ADODB.Connection
'    con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\CajasCustodia\Custodia.accdb;Persist Security Info=False"
'    SQL = " SELECT IDCliente, IDCaja, Ubicacion , estado"
'    SQL = SQL & " FROM CajasCustodia "
'    SQL = SQL & " where IDCaja < 300000"
'   SQL = SQL & " ORDER BY IDCaja ;"
'    RS.Open SQL, con
'    Dim Sqlc As String
'    Do While Not RS.EOF
'            SQL = "  Update Cajas"
'            SQL = SQL & " SET DIGITO_VERIFICADOR =" & Mid(RS!Ubicacion)
'            SQL = SQL & " Where ID_CAJA = " & RS!IDCaja
'            ExecutarSql SQL
'
'        RS.MoveNext
'    Loop






End Sub

Private Sub Command71_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String


Sql = "SELECT     MOV_CAJAS2.NRO_REMITO, REMITOS_CUERPO.FECHA"
Sql = Sql & " FROM         MOV_CAJAS2 LEFT OUTER JOIN"
Sql = Sql & "                       REMITOS_CUERPO ON MOV_CAJAS2.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO"
Sql = Sql & "  WHERE     (MOV_CAJAS2.FECHA_MOVIMIENTO < '1990 - 11 - 16 00:00:00.000')"
Sql = Sql & "  GROUP BY MOV_CAJAS2.NRO_REMITO, REMITOS_CUERPO.FECHA"
rs.Open Sql, strConBasa

Do While Not rs.EOF
    
    Sql = " UPDATE    MOV_CAJAS2"
Sql = Sql & " SET  FECHA_MOVIMIENTO = " & FechaFormato(rs!fecha)
Sql = Sql & " Where NRO_REMITO = " & rs!NRO_REMITO
    ExecutarSql Sql
    

    rs.MoveNext
Loop





End Sub

Private Sub Command72_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rsIndice As New ADODB.Recordset

Sql = " SELECT     COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, DESCRIPCION, ID, BORRAR"
Sql = Sql & " From INDICES"
Sql = Sql & " ORDER BY INDICE, COD_CLIENTE"

rsIndice.Open Sql, strConBasa
Dim Control As String

Do While Not rsIndice.EOF

Control = ""

    Sql = " SELECT     COD_CLIENTE, INDICE"
    Sql = Sql & " From REFERENCIAS "
    Sql = Sql & " WHERE     COD_CLIENTE = " & rsIndice!COD_CLIENTE
    Sql = Sql & "  AND INDICE = '" & rsIndice!Indice & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, strConBasa
    If Not rs.EOF Then
        Control = "Refenencia"
    End If
    
    
    
    Sql = " SELECT     COD_CLIENTE, COD_INDICE"
    Sql = Sql & "  From LEGAJOS"
    Sql = Sql & "  WHERE COD_CLIENTE = " & rsIndice!COD_CLIENTE
    Sql = Sql & "  AND COD_INDICE = '" & rsIndice!Indice & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, strConBasa
    If Not rs.EOF Then
        Control = Control & " Legajos"
    End If
    
    
    
    Sql = " SELECT     FK_CLIENTES, FK_INDICES"
Sql = Sql & " From DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where FK_CLIENTES = " & rsIndice!COD_CLIENTE
Sql = Sql & " And FK_INDICES = " & rsIndice!ID

    Set rs = New ADODB.Recordset
    rs.Open Sql, strConBasa
    If Not rs.EOF Then
        Control = Control & " Digital"
    End If


Sql = " SELECT     COD_CLIENTE, COD_INDICE"
Sql = Sql & "  From ORDENAR_DOCUMENTACION_DETALLE"
Sql = Sql & "  Where COD_CLIENTE = " & rsIndice!COD_CLIENTE
Sql = Sql & "  AND COD_INDICE = '" & rsIndice!Indice & "'"

Set rs = New ADODB.Recordset
    rs.Open Sql, strConBasa
    If Not rs.EOF Then
        Control = Control & " rearchivo"
    End If

    
    Sql = " Update INDICES"
Sql = Sql & " SET BORRAR2 ='" & Control & "'"
Sql = Sql & " WHERE     ID = " & rsIndice!ID

ExecutarSql Sql

Control = ""


    rsIndice.MoveNext
Loop






End Sub

Private Sub Command73_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset

Sql = " SELECT [ID]"
Sql = Sql & "  ,[DESCRIPCION]"
Sql = Sql & "   From INDICE_DESCRIP$ "
rs.Open Sql, ConBasa


Do While Not rs.EOF

Sql = " Update INDICES"
Sql = Sql & " SET DESCRIPCION ='" & rs!Descripcion & "'"
Sql = Sql & " WHERE     ID = " & rs!ID
ExecutarSql Sql
rs.MoveNext

Loop



End Sub


Private Sub Command74_Click()
Dim i As Integer
Dim Sql As String





For i = 100 To 9999

    Sql = " INSERT INTO CU (uno, dos)"
    Sql = Sql & " Values(" & i & "," & i + 1 & ")"
    ConBasa.Execute Sql
    i = i + 1
    





Next



End Sub

Private Sub Command75_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset
inicio
Sql = " SELECT     BORRAR_USUARIO, ID_CLIENTEUSUARIO, APELLIDO_NOMBRE, CORREO"
Sql = Sql & " From correo"
Sql = Sql & " ORDER BY CORREO, BORRAR_USUARIO"

rs.Open Sql, ConActiva


Do While Not rs.EOF
   Sql = " Update CLIENTEUSUARIO"
   Sql = Sql & " SET  "
   If Not IsNull(rs!APELLIDO_NOMBRE) Then
        Sql = Sql & "   APELLIDO_NOMBRE ='" & UCase(Trim(rs!APELLIDO_NOMBRE)) & "' , "
        
   End If
   
   
   If Not IsNull(rs!correo) Then
     Sql = Sql & "   CORREO ='" & Trim(LCase(rs!correo)) & "'"
   End If
   
   If rs!BORRAR_USUARIO = 1 Then
    Sql = Sql & "  ,    DESHABILITADO =1"
   End If
    
    If Not IsNull(rs!ID_CLIENTEUSUARIO) Then
    Sql = Sql & "  Where ID_CLIENTEUSUARIO = " & rs!ID_CLIENTEUSUARIO

    ConBasa.Execute Sql
    
    End If
    
    rs.MoveNext
Loop





End Sub

Private Sub Command76_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset

Sql = " SELECT     ID_CLIENTEUSUARIO, Expr1 AS CANTI"
Sql = Sql & " From basasql.dbo.TAZADEUSOUSUARIOS"


rs.Open Sql, ConBasa

Do While Not rs.EOF

    

    rs.MoveNext
Loop




End Sub

Private Sub Command77_Click()
Dim Sql As String

Sql = " SELECT     CAJASDUPLICADAS.COD_CLIENTE, CAJASDUPLICADAS.NRO_CAJA, CONTENEDOR.ESTADO, CONTENEDOR.ID_CONTENEDOR, CONTENEDOR.ESTANTERIA,"
 Sql = Sql & "                     CONTENEDOR.Horizontal , CONTENEDOR.Vertical, CONTENEDOR.Adelante_Atras"
Sql = Sql & " FROM         CAJASDUPLICADAS INNER JOIN"
              Sql = Sql & "         CONTENEDOR ON CAJASDUPLICADAS.COD_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJASDUPLICADAS.NRO_CAJA = CONTENEDOR.NRO_CAJA"
             Sql = Sql & "    Where (CAJASDUPLICADAS.Expr1 = 2)"
Sql = Sql & " ORDER BY CAJASDUPLICADAS.COD_CLIENTE, CAJASDUPLICADAS.NRO_CAJA, CONTENEDOR.ESTANTERIA"

Dim rs As New ADODB.Recordset

rs.Open Sql, ConActiva
Do While Not rs.EOF
        Debug.Print rs!NRO_CAJA
    If rs!Estanteria > 120 And rs!Estanteria < 200 Then
            Sql = " DELETE FROM CONTENEDOR Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR
            ExecutarSql Sql
            rs.MoveNext
    End If
    

Debug.Print rs!NRO_CAJA


    rs.MoveNext
Loop



End Sub

Private Sub Command78_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    
    
    Sql = " SELECT     LEGAJOS.NRO_DESDE, LEGAJOS.NRO_HASTA, LEGAJOS.LETRA_DESDE, LEGAJOS.LETRA_HASTA AS Expr4, DOCUMENTOS_DIGITALES.ID,"
    Sql = Sql & " DOCUMENTOS_DIGITALES.NRO_DESDE AS Expr1, DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE,"
    Sql = Sql & " DOCUMENTOS_DIGITALES.ArchivoNombre"
Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
                      Sql = Sql & "  LEGAJOS ON DOCUMENTOS_DIGITALES.NRO_DESDE = LEGAJOS.ID_LEGAJO"
Sql = Sql & "  WHERE     (DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE IN (20449, 20448, 20447, 20446, 20445, 20444, 20443, 20442, 20441, 20440, 20439, 20438, 20437,"
                      Sql = Sql & "  20436, 20435, 20434, 20433, 20432, 20431, 20430, 20429, 20428))"
Sql = Sql & "  ORDER BY LEGAJOS.NRO_DESDE"



rs.Open Sql, ConActiva

Do While Not rs.EOF
    
Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
Sql = Sql & " SET LETRA_DESDE =" & rs!LETRA_DESDE
Sql = Sql & "  , LETRA_HASTA =" & rs!LETRA_DESDE
Sql = Sql & "  , NRO_DESDE =" & rs!NRO_DESDE
Sql = Sql & "  , NRO_HASTA =" & rs!NRO_HASTA
Sql = Sql & "   Where ID = " & rs!ID

ExecutarSql Sql
    rs.MoveNext

Loop







End Sub

Private Sub Command79_Click()
Dim rs As New ADODB.Recordset
Dim RsControl As New ADODB.Recordset
Dim ControlCaja As String
Dim ControlContenedor As String
Dim ClienteCustodia As Long

Dim con  As New ADODB.Connection


con.Open strConBasa


Dim Sql As String

    Sql = " SELECT   ID_CONTENEDOR, COD_CLIENTE, NRO_CAJA, LECTURA, EMPRESA, BARRAANTERIOR, ESTANTERIA, HORIZONTAL, VERTICAL, FECHA_CREACION, FECHACONTROL,"
    Sql = Sql & " ESTADOCONTROL , ESTADOCAJA, ESTADOCONTENEDOR, ESTADOALSINA"
    Sql = Sql & " From basasql.dbo.ALSINAFINAL"
    Sql = Sql & " ORDER BY ID_CONTENEDOR "
      
      Sql = " SELECT  ID_CONTENEDOR ,    COD_CLIENTE, NRO_CAJA, LECTURA, EMPRESA, BARRAANTERIOR, ESTANTERIA, HORIZONTAL, VERTICAL, ESTADOCAJA, ESTADOCONTENEDOR"
Sql = Sql & " From basasql.dbo.ALSINAFINAL"
Sql = Sql & " WHERE     (ESTADOCAJA = N'CORRECTA') AND (ESTADOCONTENEDOR = N'EST ACT')"
 Sql = Sql & " ORDER BY ID_CONTENEDOR "
    
    rs.Open Sql, strConBasa
 
 Do While Not rs.EOF

 
    Rem caja
    Sql = " SELECT     COUNT(*) as cantidad "
    Sql = Sql & "  From basasql.dbo.Cajas "
         Sql = Sql & "  Where FK_CLIENTE =  " & rs!COD_CLIENTE
            Sql = Sql & "  And NRO_CAJA = " & rs!NRO_CAJA
      
  
        Set RsControl = New ADODB.Recordset
        RsControl.Open Sql, strConBasa
        If RsControl!cantidad = 0 Then
            ControlCaja = "NO EXSISTE"
        End If
        If RsControl!cantidad = 1 Then
            ControlCaja = "CORRECTA"
        End If
        If RsControl!cantidad > 1 Then
            ControlCaja = "DUPLICADA"
        End If
        
        
        Rem contenedor
        

        Sql = " SELECT     COUNT(*) as cantidad "
    Sql = Sql & "  From CONTENEDOR  "
    
            Sql = Sql & "  Where COD_CLIENTE =  " & rs!COD_CLIENTE
            Sql = Sql & "  And NRO_CAJA = " & rs!NRO_CAJA
       
       
        Set RsControl = New ADODB.Recordset
        RsControl.Open Sql, strConBasa
        If RsControl!cantidad = 0 Then
            ControlContenedor = "NO EXSISTE"
        End If
        If RsControl!cantidad = 1 Then
            ControlContenedor = "CORRECTA"
        End If
        If RsControl!cantidad > 1 Then
            ControlContenedor = "DUPLICADA"
        End If
    
     If ControlContenedor = "CORRECTA" Then
        
        
        
        
        
        Sql = " SELECT     COUNT(*) as cantidad  "
        Sql = Sql & "  From basasql.dbo.CONTENEDOR"
        Sql = Sql & "  Where COD_CLIENTE =  " & rs!COD_CLIENTE
        Sql = Sql & " AND   NRO_CAJA = " & rs!NRO_CAJA
        Sql = Sql & "  And Estanteria = " & rs!Estanteria
        Sql = Sql & "  And Horizontal = " & rs!Horizontal
        Sql = Sql & "  And Vertical = " & rs!Vertical
        
        Set RsControl = New ADODB.Recordset
        RsControl.Open Sql, strConBasa
        If RsControl!cantidad = 0 Then
            ControlContenedor = "EST NO ACT"
        End If
        If RsControl!cantidad = 1 Then
            ControlContenedor = "EST ACT"
        End If
     
     End If
     
        Sql = " Update basasql.dbo.ALSINAFINAL"
        Sql = Sql & " SET  ESTADOCAJA ='" & ControlCaja & "',  ESTADOCONTENEDOR = '" & ControlContenedor & "'"
        Sql = Sql & " Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR
        con.Execute Sql
    
    rs.MoveNext
 Loop
 

End Sub

Private Sub Command8_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rsbasa As New ADODB.Recordset

rs.CursorLocation = adUseClient

Sql = "SELECT     ID, CAJA, CLIENTE, LECTURABASA"
Sql = Sql & " From CONTROLPEDRO"
Sql = Sql & " ORDER BY ID"

rs.Open Sql, ConActiva, adOpenDynamic, adLockPessimistic

Dim cambio As String


Do While Not rs.EOF


Sql = " SELECT     LECTURA_COLECTOR_CUERPO.DESCRIPCION, LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA, LECTURACOLECTOR.CAJA,"
Sql = Sql & " LECTURACOLECTOR.Cliente"
Sql = Sql & " FROM LECTURACOLECTOR INNER JOIN LECTURA_COLECTOR_CUERPO ON LECTURACOLECTOR.NUMERO_LECTURA = LECTURA_COLECTOR_CUERPO.NUMERO_LECTURA"
Sql = Sql & " WHERE     (LECTURA_COLECTOR_CUERPO.DESCRIPCION LIKE '%PEDRO%') "
Sql = Sql & " AND LECTURACOLECTOR.CAJA =  " & rs!Caja
If rs!Cliente <> 9999 And rs!Cliente <> 0 Then
    Sql = Sql & " AND LECTURACOLECTOR.CLIENTE =" & rs!Cliente
End If

Set rsbasa = New ADODB.Recordset
rsbasa.Open Sql, ConActiva, 0, 1
     If Not rsbasa.EOF Then
     
        rs!LECTURABASA = rsbasa!NUMERO_LECTURA
     
     End If
        
        rs.Update
        
        
    rs.MoveNext
Loop




End Sub

Private Sub ControlCajas()


Dim ConBasa As New ADODB.Connection
Dim Sql As String

ConBasa.Open strConBasa


Dim rs As New ADODB.Recordset

Dim RsControl As New ADODB.Recordset

rs.CursorLocation = adUseClient




ExecutarSql "DELETE FROM TEM_CONTROL_CAJAS"
Sql = " INSERT INTO TEM_CONTROL_CAJAS"
Sql = Sql & " (FK_LECTURA, FK_CAJA, FK_CLIENTE, ORDEN)"
Sql = Sql & "  SELECT     NUMERO_LECTURA, CAJA, CLIENTE, ORDEN"
Sql = Sql & "  From LECTURACOLECTOR "
Sql = Sql & "  WHERE     (NUMERO_LECTURA IN ( " & InputBox("Ingrese los numeros de lectura separados ,", "", 0) & "))"
Sql = Sql & " ORDER BY NUMERO_LECTURA, ORDEN"

ExecutarSql Sql


Sql = " SELECT FK_LECTURA, FK_CLIENTE, FK_CAJA, ORDEN, REMITO_VACIAS, REMITOS_CUSTODIA, REFERENCIAS_RANGO, REFERENCIA_LEGAJOS,"
Sql = Sql & "  PERSONAL_ASIGNADO From TEM_CONTROL_CAJAS "


rs.Open Sql, ConActiva, adOpenKeyset, adLockOptimistic


Do While Not rs.EOF
    
    Sql = "  SELECT REMITOS_CUERPO.NRO_REMITO , ID_CLIENTE "
    Sql = Sql & " FROM REMITOS_CUERPO INNER JOIN "
    Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO "
    Sql = Sql & " WHERE     (REMITOS_CUERPO.TIPO = 0) AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) "
    If rs!FK_CAJA < 740000 Then
        Sql = Sql & "   AND REMITOS_CUERPO.ID_CLIENTE = " & rs!FK_CLIENTE
    End If
        Sql = Sql & " AND  REMITOS_DETALLE.DESDE = " & rs!FK_CAJA
        
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, adOpenKeyset, adLockPessimistic
       If Not RsControl.EOF Then
            rs!REMITOS_CUSTODIA = RsControl!NRO_REMITO
            rs!FK_CLIENTE = RsControl!id_cliente
       End If
       
       
       
    Sql = "  SELECT REMITOS_CUERPO.NRO_REMITO "
    Sql = Sql & " FROM REMITOS_CUERPO INNER JOIN "
    Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO "
    Sql = Sql & " WHERE     (REMITOS_CUERPO.TIPO = 2) AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) "
    If rs!FK_CAJA < 740000 Then
        Sql = Sql & "   AND REMITOS_CUERPO.ID_CLIENTE = " & rs!FK_CLIENTE
    End If
        Sql = Sql & " AND  REMITOS_DETALLE.DESDE = " & rs!FK_CAJA
        
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, 0, 1
       If Not RsControl.EOF Then
        rs!REMITO_VACIAS = RsControl!NRO_REMITO
       End If
       
       
       
       Sql = " SELECT COUNT(*) AS CANTIDADREFERENCIA "
        Sql = Sql & " From REFERENCIAS "
        
        Sql = Sql & " Where NRO_CAJA = " & rs!FK_CAJA
       
       If rs!FK_CAJA < 740000 Then
         Sql = Sql & " And COD_CLIENTE = " & rs!FK_CLIENTE
    End If
        
       
       
       
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, 0, 1
       If Not RsControl.EOF Then
            rs!REFERENCIAS_RANGO = RsControl!CANTIDADREFERENCIA
       Else
            rs!REFERENCIAS_RANGO = 0
       End If
       
       
        Sql = " SELECT COUNT(*) AS CantidadLegajos "
        Sql = Sql & " From LEGAJOS"
        Sql = Sql & " Where NRO_CAJA = " & rs!FK_CAJA
         If rs!FK_CAJA < 740000 Then
         Sql = Sql & " And COD_CLIENTE = " & rs!FK_CLIENTE
    End If
        
        
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, 0, 1
       If Not RsControl.EOF Then
            rs!REFERENCIA_LEGAJOS = RsControl!CANTIDADLEGAJOS
       Else
            rs!REFERENCIA_LEGAJOS = 0
       End If
       
       
       
       
       
       
      
        Sql = " SELECT     CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, PERSONAL.NOMBRE, PERSONAL.APELLIDO"
        Sql = Sql & " FROM CAJAS INNER JOIN "
        Sql = Sql & " PERSONAL ON CAJAS.FK_PERSONAL_ENTREGA = PERSONAL.IDPERSONAL "
          Sql = Sql & " Where "
        If rs!FK_CAJA < 740000 Then
            Sql = Sql & "   Cajas.FK_CLIENTE = " & rs!FK_CLIENTE & " AND "
         End If
        Sql = Sql & "  Cajas.NRO_CAJA = " & rs!FK_CAJA
       
       Set RsControl = New ADODB.Recordset
       RsControl.Open Sql, ConActiva, 0, 1
       If Not RsControl.EOF Then
            rs!PERSONAL_ASIGNADO = "'" & Trim(RsControl!Apellido) & "'"
       Else
            rs!PERSONAL_ASIGNADO = 0
       End If
       
       
       
    rs.Update
    rs.MoveNext
Loop






End Sub

Private Sub Command80_Click()
Dim rs As New ADODB.Recordset
Dim RsControl As New ADODB.Recordset
Dim ControlCaja As String
Dim ControlContenedor As String
Dim ClienteCustodia As Long




Dim Sql As String



Sql = " SELECT     ALSINAFINAL.ID_CONTENEDOR, ALSINAFINAL.COD_CLIENTE, ALSINAFINAL.NRO_CAJA, ALSINAFINAL.LECTURA, ALSINAFINAL.EMPRESA,"
 Sql = Sql & "                     ALSINAFINAL.BARRAANTERIOR, ALSINAFINAL.ESTANTERIA, ALSINAFINAL.HORIZONTAL, ALSINAFINAL.VERTICAL, ALSINAFINAL.FECHA_CREACION,"
Sql = Sql & "                      ALSINAFINAL.FECHACONTROL, ALSINAFINAL.ESTADOCONTROL, ALSINAFINAL.ESTADOCAJA, ALSINAFINAL.ESTADOCONTENEDOR, ALSINAFINAL.ESTADOALSINA,"
 Sql = Sql & "                      CONTENEDOR.ESTADO, CONTENEDOR.COD_CLIENTE AS Expr1, CONTENEDOR.NRO_CAJA AS Expr2"
Sql = Sql & " FROM         ALSINAFINAL INNER JOIN"
  Sql = Sql & "                    CONTENEDOR ON ALSINAFINAL.ESTANTERIA = CONTENEDOR.ESTANTERIA AND ALSINAFINAL.HORIZONTAL = CONTENEDOR.HORIZONTAL AND"
  Sql = Sql & "                       ALSINAFINAL.Vertical = CONTENEDOR.Vertical"
Sql = Sql & "    WHERE     (ALSINAFINAL.ESTADOCAJA = N'CORRECTA') AND (ALSINAFINAL.ESTADOCONTENEDOR = N'no EXSISTE') AND (CONTENEDOR.ESTADO = 1)"
    
    rs.Open Sql, strConBasa
 
 Do While Not rs.EOF
    Sql = " Update basasql.dbo.CONTENEDOR"
    Sql = Sql & "  SET COD_CLIENTE = " & rs!COD_CLIENTE & ", NRO_CAJA =" & rs!NRO_CAJA & ", ESTADO =2"
    Sql = Sql & " Where (Estanteria =" & rs!Estanteria & " ) And (Horizontal = " & rs!Horizontal & ") And (Vertical = " & rs!Vertical & ")"
    ExecutarSql Sql
    rs.MoveNext
Loop
End Sub

Private Sub Command81_Click()

Dim rsAlsina As New ADODB.Recordset
Dim RsControl As New ADODB.Recordset
Dim rsContenedor As New ADODB.Recordset
Dim ControlCaja As String
Dim ControlContenedor As String
Dim ClienteCustodia As Long




Dim Sql As String


 Sql = " SELECT     ALSINAFINAL.ID_CONTENEDOR AS ALSINAFINAL_ID_CONTENEDOR, ALSINAFINAL.COD_CLIENTE AS ALSINA_COD_CLIENTE,"
  Sql = Sql & "                     ALSINAFINAL.NRO_CAJA AS ALSINA_NRO_CAJA, ALSINAFINAL.ESTANTERIA AS ALSINA_ESTANTERIA, ALSINAFINAL.HORIZONTAL AS ALSINA_HORIZONTAL,"
  Sql = Sql & "                     ALSINAFINAL.VERTICAL AS ALSINA_VERTICAL, CONTENEDOR.ESTANTERIA AS BASA_ESTANTERIA, CONTENEDOR.HORIZONTAL AS BASA_HORIZONTAL,"
  Sql = Sql & "                     CONTENEDOR.VERTICAL AS BASA_VERTICAL, CONTENEDOR.ID_CONTENEDOR AS BASA_ID_CONTENEDOR, CONTENEDOR.ESTADO AS BASA_ESTADO,"
  Sql = Sql & "                     CONTENEDOR.FECHAPOSICION"
Sql = Sql & "  FROM         ALSINAFINAL INNER JOIN"
Sql = Sql & "                       CONTENEDOR ON ALSINAFINAL.COD_CLIENTE = CONTENEDOR.COD_CLIENTE AND ALSINAFINAL.NRO_CAJA = CONTENEDOR.NRO_CAJA AND"
Sql = Sql & "                       ALSINAFINAL.ESTANTERIA > CONTENEDOR.ESTANTERIA"
Sql = Sql & "  WHERE     (ALSINAFINAL.ESTADOCAJA = N'CORRECTA') AND (ALSINAFINAL.ESTADOCONTENEDOR = N'EST NO ACT') AND (CONTENEDOR.FECHAPOSICION IS NULL)"
Sql = Sql & "  ORDER BY ALSINA_COD_CLIENTE, ALSINA_NRO_CAJA, ALSINA_ESTANTERIA"


    rsAlsina.Open Sql, strConBasa
    Dim fecha As String
    
    
    fecha = SysDate2
 
 Do While Not rsAlsina.EOF
 
 





Sql = " INSERT INTO CAMBIOPOSICION"
Sql = Sql & " (ID_PERSONAL, FECHA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA)"
Sql = Sql & "  SELECT    17  , " & fecha & ", ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO,"
Sql = Sql & "  COD_CLIENTE , NRO_CAJA"
Sql = Sql & "  From CONTENEDOR"
Sql = Sql & "  Where ID_CONTENEDOR = " & rsAlsina!BASA_ID_CONTENEDOR
ExecutarSql Sql

Sql = " Update CONTENEDOR"
Sql = Sql & " SET  "
Sql = Sql & " ESTADO =1"
Sql = Sql & ", COD_CLIENTE =Null"
Sql = Sql & " , NRO_CAJA =null"
Sql = Sql & " , NRO_REMITO =Null"
Sql = Sql & " , UB_PROVISORIA =Null"
Sql = Sql & "  Where ID_CONTENEDOR = " & rsAlsina!BASA_ID_CONTENEDOR
ExecutarSql Sql


Sql = " Update CONTENEDOR"
Sql = Sql & " SET  "
Sql = Sql & " ESTADO =" & rsAlsina!BASA_ESTADO
Sql = Sql & ", COD_CLIENTE =" & rsAlsina!ALSINA_COD_CLIENTE
Sql = Sql & " , NRO_CAJA =" & rsAlsina!ALSINA_NRO_CAJA
Sql = Sql & " , FECHAPOSICION = " & fecha
Sql = Sql & "  Where  ESTANTERIA  = " & rsAlsina!ALSINA_ESTANTERIA
Sql = Sql & " and  HORIZONTAL = " & rsAlsina!ALSINA_HORIZONTAL
Sql = Sql & " and  VERTICAL = " & rsAlsina!ALSINA_VERTICAL
ExecutarSql Sql
    rsAlsina.MoveNext
Loop

End Sub

Private Sub Command82_Click()

Dim rsUgarte As New ADODB.Recordset
Dim Sql As String


Dim con As New ADODB.Connection

con.Open strConBasa
    Dim fecha As String
    
    
    fecha = SysDate
 
 
 


Sql = " SELECT     ESTADOCAJA, ID_CONTENEDOR, COD_CLIENTE, NRO_CAJA, ESTANTERIA, HORIZONTAL, VERTICAL"
  Sql = Sql & " From ALSINAFINAL"
  Sql = Sql & " WHERE     (ESTADOCAJA = N'ugarte')"
  Sql = Sql & "  ORDER BY ID_CONTENEDOR "
  
  rsUgarte.Open Sql, strConBasa
  Do While Not rsUgarte.EOF
           Sql = " Update basasql.dbo.Cajas"
        Sql = Sql & "  SET  FK_CLIENTE =1002"
        Sql = Sql & "  , NRO_CAJA =" & Trim(rsUgarte!NRO_CAJA)
        Sql = Sql & "  , FK_ESTADO =1020"
        Sql = Sql & " , FECHA_CREACION_CAJA =" & fecha
        Sql = Sql & " , DEPOSITO ='24012013'"
         Sql = Sql & " Where ID_CAJA = " & rsUgarte!NRO_CAJA
         Sql = Sql & " And (FK_CLIENTE Is Null)"
        con.Execute Sql
   
   
   Sql = " UPDATE  basasql.dbo.CONTENEDOR"
Sql = Sql & "  SET  COD_CLIENTE =1002"
Sql = Sql & "  , NRO_CAJA =" & rsUgarte!NRO_CAJA
Sql = Sql & " , FECHAPOSICION =" & fecha
Sql = Sql & " ,  ESTADO = 2"
Sql = Sql & "  Where Estanteria =  " & rsUgarte!Estanteria
Sql = Sql & "   And Horizontal =  " & rsUgarte!Horizontal
Sql = Sql & "   And VERTICAL = " & rsUgarte!Vertical
Sql = Sql & "   And (COD_CLIENTE Is Null)"
   con.Execute Sql
  
  
        rsUgarte.MoveNext
  Loop
  
  
  
End Sub


Private Sub Command84_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Indice As String
Dim Descripcion As String
Dim con As New ADODB.Connection

Dim cod_doc As Long
Sql = " SELECT     ID, COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, DESCRIPCION, LEN(INDICE) AS Expr1"
Sql = Sql & "  From basasql.dbo.INDICES"
Sql = Sql & "  WHERE     (COD_CLIENTE = 1197) AND (LEN(INDICE) = 9) AND (INDICE LIKE '001%')"
Sql = Sql & " ORDER BY INDICE"

rs.Open Sql, strConBasa
con.Open strConBasa
Do While Not rs.EOF
        Indice = rs!Indice & "004"
        Descripcion = "DEVOLUCIONES - SM " & rs!ID_CODIGO_DOCUMENTO
        cod_doc = rs!ID_CODIGO_DOCUMENTO + 40000
        Sql = " INSERT INTO basasql.dbo.INDICES"
        Sql = Sql & " (COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, DESCRIPCION, HABILITAR_FECHA_DESDE, HABILITAR_FECHA_HASTA)"
        Sql = Sql & "  VALUES     (1197"
        Sql = Sql & "," & cod_doc
        Sql = Sql & ",'" & Trim(Indice) & "'"
        Sql = Sql & ",'" & Trim(Descripcion) & "'"
      Sql = Sql & ",'1'"
      Sql = Sql & ",'1')"
       con.Execute Sql
       
       
        Indice = rs!Indice & "005"
        Descripcion = "NOTAS DE CREDITO - SM " & rs!ID_CODIGO_DOCUMENTO
        cod_doc = rs!ID_CODIGO_DOCUMENTO + 50000
        Sql = " INSERT INTO basasql.dbo.INDICES"
        Sql = Sql & " (COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE, DESCRIPCION, HABILITAR_FECHA_DESDE, HABILITAR_FECHA_HASTA)"
        Sql = Sql & "  VALUES     (1197"
        Sql = Sql & "," & cod_doc
        Sql = Sql & ",'" & Trim(Indice) & "'"
        Sql = Sql & ",'" & Trim(Descripcion) & "'"
      Sql = Sql & ",'1'"
      Sql = Sql & ",'1')"
       con.Execute Sql
    
     

    rs.MoveNext
Loop

End Sub

Private Sub Command85_Click()

Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim Sql As String
con.Open strConBasa

Sql = "SELECT   ESTANTERIA   , VERTICAL, MODULO_H, ESTADO, Expr1"
Sql = Sql & " From basasql.dbo.AAPOSICIONESANULAR"

rs.Open Sql, strConBasa


 Do While Not rs.EOF
Sql = " Update basasql.dbo.CONTENEDOR"
Sql = Sql & "  SET ESTADO =0"
Sql = Sql & " Where Estanteria = " & rs!Estanteria
Sql = Sql & " And Vertical = " & rs!Vertical
Sql = Sql & " And Modulo_H = " & rs!Modulo_H
Sql = Sql & " And estado = 1"


    con.Execute Sql
    rs.MoveNext
    
Loop

End Sub

Private Sub Command86_Click()
Dim MyName As String
Dim VarTexto As String
Dim Caja As Long
Dim CajaAnterior As Long
Dim Sql As String
Dim IDLegajo As Long

Dim con As New ADODB.Connection
con.Open strConBasa
Dim Paso As String

Paso = "Z:\Administracion\Administracion Provincial Del Fondo\"
MyName = Dir(Paso & "*.txt", vbDirectory)
Dim P As Integer
Close #1
Do While MyName <> ""
 Caja = Mid(MyName, 7, 7)
     Sql = "  Insert "
            Sql = Sql & " Into basasql.dbo.CONTROLFONDO(Caja, Legajo, Archivo)"
            Sql = Sql & " VALUES (" & Caja & "," & 0 & ",'" & MyName & "')"
            con.Execute Sql
 
    Open Paso & MyName For Input As #1

            
    Do Until EOF(1)
    
    CajaAnterior = 0
        Line Input #1, VarTexto
        If VarTexto <> "" Then
        IDLegajo = 0
        
'             If Mid(VarTexto, 20, 4) = "BASA" Or Mid(VarTexto, 20, 4) = "VBAS" Or Mid(VarTexto, 20, 4) = "ESTA" Then
''             Caja = Mid(VarTexto, 5, 9)
''              If CajaAnterior <> Caja Then
''
''                CajaAnterior = Caja
''              End If
''
'             Else
'
             
            IDLegajo = Mid(VarTexto, 5, 9)
            Sql = "  Insert "
            Sql = Sql & " Into basasql.dbo.CONTROLFONDO(Caja, Legajo, Archivo)"
            Sql = Sql & " VALUES (" & Caja & "," & IDLegajo & ",'" & MyName & "')"
            con.Execute Sql
             
'             End If
            


 
 
        End If
    Loop
    Close #1
    MyName = Dir()
Loop
End Sub

Private Sub Command87_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim rsCaja As New ADODB.Recordset

Sql = "SELECT     NRO_CAJA, EMPRESA, Expr1, COD_CLIENTE, ID_CONTENEDOR, ESTANTERIA, HORIZONTAL, VERTICAL, FECHA_CREACION, FECHACONTROL, LECTURA, ESTADO,"
Sql = Sql & vbCrLf & "     CONTENEDOR_CLIENTE , CONTENEDOR_CAJA, CONTENEDOR_ID_CONTENEDOR, POSICIONVALIDA"
Sql = Sql & vbCrLf & "   From basasql.dbo.CAJASALSINALECUTAMAYOR"
Sql = Sql & vbCrLf & "   Where     (NRO_CAJA < 100000) AND (EMPRESA LIKE N'%cus%') AND (COD_CLIENTE = 0)"
Sql = Sql & vbCrLf & "   ORDER BY EMPRESA, COD_CLIENTE, NRO_CAJA, LECTURA DESC"


rs.Open Sql, strConBasa

Do While Not rs.EOF
        
     Sql = "SELECT     ID_CAJA, FK_CLIENTE"
Sql = Sql & vbCrLf & " From basasql.dbo.Cajas"
Sql = Sql & vbCrLf & " Where ID_CAJA =" & rs!NRO_CAJA
        
        Set rsCaja = New ADODB.Recordset
        
        rsCaja.Open Sql, strConBasa
        
        If Not rsCaja.EOF Then
            If Not IsNull(rsCaja!FK_CLIENTE) Then
                Sql = " Update basasql.dbo.CAJASALSINALECUTAMAYOR "
                Sql = Sql & " SET  COD_CLIENTE = " & rsCaja!FK_CLIENTE
                Sql = Sql & "  Where NRO_CAJA = " & rsCaja!ID_CAJA
                ExecutarSql Sql
            End If
        End If
        
        

    
     rs.MoveNext
Loop






End Sub

Private Sub Command88_Click()
    Dim Sql As String
     Dim rs As New ADODB.Recordset
        Sql = " SELECT     ID, NRO_CAJA, EMPRESA, LECTURA, COD_CLIENTE"
        Sql = Sql & " From basasql.dbo.CAJASALSINALECUTAMAYOR"
        Sql = Sql & " ORDER BY EMPRESA, COD_CLIENTE, NRO_CAJA, LECTURA DESC"
    Dim anterior As String
    Dim actual As String

anterior = ""

 rs.Open Sql, strConBasa
 
 Do While Not rs.EOF
 
    actual = Trim(CStr(rs!NRO_CAJA)) & UCase(Trim(rs!Empresa)) & Trim(CStr(rs!COD_CLIENTE))
    If anterior <> actual Then
        Sql = " Update basasql.dbo.CAJASALSINALECUTAMAYOR"
        Sql = Sql & " Set ESTADO_ACTUAL = 1"
        Sql = Sql & " Where ID = " & rs!ID
        ExecutarSql Sql
        anterior = actual
     
    End If
    
    
    rs.MoveNext
Loop



End Sub

Private Sub Command89_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
   

Sql = "  SELECT     CAJASALSINALECUTAMAYOR.NRO_CAJA, CAJASALSINALECUTAMAYOR.COD_CLIENTE, CAJASALSINALECUTAMAYOR.ESTANTERIA,"
Sql = Sql & " CAJASALSINALECUTAMAYOR.Horizontal , CAJASALSINALECUTAMAYOR.Vertical, CONTENEDOR_1.estado"
Sql = Sql & " FROM         CAJASALSINALECUTAMAYOR INNER JOIN"
Sql = Sql & "                     CONTENEDOR ON CAJASALSINALECUTAMAYOR.ESTANTERIA = CONTENEDOR.ESTANTERIA AND"
Sql = Sql & "                    CAJASALSINALECUTAMAYOR.HORIZONTAL = CONTENEDOR.HORIZONTAL AND CAJASALSINALECUTAMAYOR.VERTICAL = CONTENEDOR.VERTICAL INNER JOIN"
Sql = Sql & "                   CONTENEDOR AS CONTENEDOR_1 ON CAJASALSINALECUTAMAYOR.NRO_CAJA = CONTENEDOR_1.NRO_CAJA AND"
Sql = Sql & "                 CAJASALSINALECUTAMAYOR.COD_CLIENTE = CONTENEDOR_1.COD_CLIENTE"
Sql = Sql & " Where (CAJASALSINALECUTAMAYOR.ESTADO_ACTUAL = 1) And (CONTENEDOR.COD_CLIENTE Is Null)"
Sql = Sql & " ORDER BY CAJASALSINALECUTAMAYOR.EMPRESA, CAJASALSINALECUTAMAYOR.COD_CLIENTE, CAJASALSINALECUTAMAYOR.NRO_CAJA,"
Sql = Sql & "                     CAJASALSINALECUTAMAYOR.Lectura Desc"

Set rs = New ADODB.Recordset




rs.Open Sql, strConBasa





Do While Not rs.EOF

    Sql = Sql & "  Update CONTENEDOR"
    Sql = Sql & "  SET COD_CLIENTE = NULL, NRO_CAJA = NULL, ESTADO = 1"
    Sql = Sql & "  Where COD_CLIENTE =" & rs!COD_CLIENTE & "  And NRO_CAJA = " & rs!NRO_CAJA
    ExecutarSql Sql
    
    
    Sql = " Update CONTENEDOR"
 Sql = Sql & "  SET  NRO_CAJA =" & rs!NRO_CAJA & ", COD_CLIENTE =" & rs!COD_CLIENTE & " , ESTADO =" & rs!estado
 Sql = Sql & "  Where Estanteria =  " & rs!Estanteria
   Sql = Sql & "  And Horizontal = " & rs!Horizontal
   Sql = Sql & "  And Vertical = " & rs!Vertical
    Rem Sql = Sql & "  And Adelante_Atras = " & rs!Adelante_Atras
    
    ExecutarSql Sql
    
    rs.MoveNext

Loop

End Sub

Private Sub Command90_Click()
 Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
        Dim R As Integer
        Dim C As Integer
        Dim Sql As String
        Dim KF_CLIENTE As Integer
        Dim NRO_CAJA As Long
        Dim con As New ADODB.Connection
        Dim P As Integer
        Dim Bloque As String
      
        

        'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open("c:\291\291.xls")
         Set hojaEx = libroEx.Worksheets.Item(1)
        
        C = 1
        con.Open strConBasa
       For R = 2 To 80
       Rem  MsgBox hojaEx.Cells(R, C)
         For C = 2 To 20
            
             If hojaEx.Cells(R, C) <> "" Then
             
            Sql = "  INSERT INTO basasql.dbo.AAORDEN"
            Sql = Sql & " (CAJA, ORDEN)"
            Sql = Sql & " VALUES ("
            Sql = Sql & hojaEx.Cells(R, 1)
            Sql = Sql & "," & hojaEx.Cells(R, C) & ")"
            ExecutarSql Sql
             End If
             
             
            
            
         Next
       Next
                
            libroEx.Close
End Sub

Private Sub Command91_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Sql As String


Sql = " SELECT     AAORDEN.ID, AAORDEN.CAJA, AAORDEN.ORDEN, ORDEN_LEGAJOS.REARCHIVO_CAJA, ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO,"
Sql = Sql & " ORDEN_LEGAJOS.COD_CLIENTE , ORDEN_LEGAJOS.ID_ORDEN_LEGAJO"
Sql = Sql & " FROM         AAORDEN INNER JOIN"
Sql = Sql & " ORDEN_LEGAJOS ON AAORDEN.ORDEN = ORDEN_LEGAJOS.ID_ORDEN_LEGAJO INNER JOIN"
Sql = Sql & " ORDEN_LEGAJOS_DETALLE ON ORDEN_LEGAJOS.ID_ORDEN_LEGAJO = ORDEN_LEGAJOS_DETALLE.COD_ORDEN_LEGAJO"
Sql = Sql & " ORDER BY AAORDEN.ORDEN"

Sql = "  SELECT     AAORDEN.ID, AAORDEN.CAJA, AAORDEN.ORDEN, ORDEN_LEGAJOS.REARCHIVO_CAJA, ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO,"
Sql = Sql & "                      ORDEN_LEGAJOS.COD_CLIENTE , ORDEN_LEGAJOS.ID_ORDEN_LEGAJO"
Sql = Sql & " FROM         AAORDEN INNER JOIN"
Sql = Sql & "                      ORDEN_LEGAJOS ON AAORDEN.ORDEN = ORDEN_LEGAJOS.ID_ORDEN_LEGAJO INNER JOIN"
 Sql = Sql & "                      ORDEN_LEGAJOS_DETALLE ON ORDEN_LEGAJOS.ID_ORDEN_LEGAJO = ORDEN_LEGAJOS_DETALLE.COD_ORDEN_LEGAJO"
Sql = Sql & " GROUP BY AAORDEN.ID, AAORDEN.CAJA, AAORDEN.ORDEN, ORDEN_LEGAJOS.REARCHIVO_CAJA, ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO,"
  Sql = Sql & "                    ORDEN_LEGAJOS.COD_CLIENTE , ORDEN_LEGAJOS.ID_ORDEN_LEGAJO"
Sql = Sql & " ORDER BY AAORDEN.ORDEN"


rs.Open Sql, strConBasa

Dim i As Integer
Do While Not rs.EOF
        Sql = " Update ORDEN_LEGAJOS "
        Sql = Sql & "  SET REARCHIVO_CAJA =" & rs!Caja
        Sql = Sql & "  Where ID_ORDEN_LEGAJO = " & rs!Orden
        ExecutarSql Sql
    
        Sql = " Update basasql.dbo.LEGAJOS "
        Sql = Sql & " SET REARCHIVO_CAJA =" & rs!Caja
        Sql = Sql & " Where ID_CLIENTE_LEGAJO = " & rs!COD_ID_CLIENTE_LEGAJO
        Sql = Sql & " And COD_CLIENTE = " & rs!COD_CLIENTE
        ExecutarSql Sql
     i = i + 1
     txtCliente.Text = i
     txtCliente.Refresh
    rs.MoveNext
Loop





End Sub

Private Sub Command92_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Etiqueta As String



Sql = " SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_CONTENEDOR, FK_ESTADO, ROLLO, DIGITO_VERIFICADOR"
Sql = Sql & "  From basasql.dbo.Cajas"
Sql = Sql & "  WHERE     (ID_CAJA BETWEEN 916344 AND 916843)"


Sql = " SELECT     CAJASFIDEREAC.COD_CLIENTE, CAJASFIDEREAC.NRO_CAJA"
Sql = Sql & " FROM         CAJASFIDEREAC INNER JOIN"
Sql = Sql & "                      CONTENEDOR ON CAJASFIDEREAC.COD_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJASFIDEREAC.NRO_CAJA = CONTENEDOR.NRO_CAJA LEFT OUTER JOIN"
Sql = Sql & "                      CAJAS ON CAJASFIDEREAC.ID_CAJA = CAJAS.ID_CAJA"
Sql = Sql & " Where (CAJASFIDEREAC.COD_CLIENTE = 1002) And (Cajas.ID_CAJA Is Null)"
Sql = Sql & " ORDER BY CAJASFIDEREAC.NRO_CAJA"



Sql = " SELECT     CAJASFIDEREAC.COD_CLIENTE, CAJASFIDEREAC.NRO_CAJA"
Sql = Sql & " FROM         CAJASFIDEREAC INNER JOIN"
Sql = Sql & "                      CONTENEDOR ON CAJASFIDEREAC.COD_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJASFIDEREAC.NRO_CAJA = CONTENEDOR.NRO_CAJA LEFT OUTER JOIN"
 Sql = Sql & "                     CAJAS ON CAJASFIDEREAC.NRO_CAJA = CAJAS.ID_CAJA"
Sql = Sql & " Where (CAJASFIDEREAC.COD_CLIENTE = 1002) And (Cajas.ID_CAJA Is Null)"
Sql = Sql & " ORDER BY CAJASFIDEREAC.NRO_CAJA"



 Sql = "  SELECT     CAJASFIDEREAC.COD_CLIENTE, CAJASFIDEREAC.NRO_CAJA, COUNT(*) AS Expr1"
 Sql = Sql & "  FROM         CAJASFIDEREAC INNER JOIN"
  Sql = Sql & "                       CONTENEDOR ON CAJASFIDEREAC.COD_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJASFIDEREAC.NRO_CAJA = CONTENEDOR.NRO_CAJA LEFT OUTER JOIN"
   Sql = Sql & "                      CAJAS ON CAJASFIDEREAC.NRO_CAJA = CAJAS.ID_CAJA"
 Sql = Sql & "  Where (Cajas.ID_CAJA Is Null)"
 Sql = Sql & "  GROUP BY CAJASFIDEREAC.COD_CLIENTE, CAJASFIDEREAC.NRO_CAJA"
 Sql = Sql & "  HAVING      (CAJASFIDEREAC.COD_CLIENTE = 1002) AND (COUNT(*) = 1)"
 Sql = Sql & "  ORDER BY CAJASFIDEREAC.NRO_CAJA"
 
 
 
 
Sql = " SELECT      CAMBIOCAJA49.clienteRearchivo, CAMBIOCAJA49.CajaInicial, CAMBIOCAJA49.clientefinal, CAMBIOCAJA49.CajaFinal, ORDEN_LEGAJOS.COD_CLIENTE,"
 Sql = Sql & "                      ORDEN_LEGAJOS.CANTIDAD, ORDEN_LEGAJOS.ID_ORDEN_LEGAJO, ORDEN_LEGAJOS_DETALLE.COD_ORDEN_LEGAJO, ORDEN_LEGAJOS_DETALLE.ORDEN,"
 Sql = Sql & "                      ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO"
 Sql = Sql & " FROM         CAMBIOCAJA49 INNER JOIN"
  Sql = Sql & "                     ORDEN_LEGAJOS ON CAMBIOCAJA49.CajaInicial = ORDEN_LEGAJOS.REARCHIVO_CAJA INNER JOIN"
   Sql = Sql & "                    ORDEN_LEGAJOS_DETALLE ON ORDEN_LEGAJOS.ID_ORDEN_LEGAJO = ORDEN_LEGAJOS_DETALLE.COD_ORDEN_LEGAJO"
 Sql = Sql & " Where (ORDEN_LEGAJOS.COD_CLIENTE = 49)"
 
 
 
 Sql = " SELECT     CAMBIOCAJA49.clienteRearchivo, CAMBIOCAJA49.CajaInicial, CAMBIOCAJA49.clientefinal, CAMBIOCAJA49.CajaFinal, ORDEN_LEGAJOS.COD_CLIENTE,"
  Sql = Sql & "                     ORDEN_LEGAJOS.CANTIDAD, ORDEN_LEGAJOS.ID_ORDEN_LEGAJO, ORDEN_LEGAJOS_DETALLE.COD_ORDEN_LEGAJO, ORDEN_LEGAJOS_DETALLE.ORDEN,"
   Sql = Sql & "                    ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO , LEGAJOS.REARCHIVO_CAJA"
 Sql = Sql & " FROM         CAMBIOCAJA49 INNER JOIN"
   Sql = Sql & "                    ORDEN_LEGAJOS ON CAMBIOCAJA49.CajaInicial = ORDEN_LEGAJOS.REARCHIVO_CAJA INNER JOIN"
     Sql = Sql & "                  ORDEN_LEGAJOS_DETALLE ON ORDEN_LEGAJOS.ID_ORDEN_LEGAJO = ORDEN_LEGAJOS_DETALLE.COD_ORDEN_LEGAJO INNER JOIN"
      Sql = Sql & "                 LEGAJOS ON ORDEN_LEGAJOS_DETALLE.COD_ID_CLIENTE_LEGAJO = LEGAJOS.ID_CLIENTE_LEGAJO AND"
        Sql = Sql & "               ORDEN_LEGAJOS.COD_CLIENTE = LEGAJOS.COD_CLIENTE"
 Sql = Sql & " Where (ORDEN_LEGAJOS.COD_CLIENTE = 49) And ((LEGAJOS.REARCHIVO_CAJA Is Null))"


Sql = " SELECT     clienteRearchivo, CajaInicial, clientefinal, CajaFinal"
 Sql = Sql & " From CAMBIOCAJA49"


rs.Open Sql, strConBasa



                
Do While Not rs.EOF


Sql = " Update ORDEN_LEGAJOS"
Sql = Sql & " SET REARCHIVO_CAJA =" & rs!CajaFinal
Sql = Sql & "  Where (COD_CLIENTE = 49)"
Sql = Sql & "  And REARCHIVO_CAJA = " & rs!CajaInicial

ExecutarSql Sql
    
    rs.MoveNext
Loop



End Sub

Private Sub Command93_Click()
Dim Sql As String
 
 Dim ConEtiquetas As New ADODB.Connection
 ConEtiquetas.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\GARBARINO.mdb;Persist Security Info=False"
Dim i As Long

'GARBARINO2500045000
'COMPUMUNDO6000175000
'DIFITAL7500175500
'TECNOSUR7550176000
'VIAJES7600176500

Dim Caja As String
Dim BARRA As String
Dim Caja_text As String

Rem GARBARINO76001_86000
For i = 86001 To 96000
        Caja = i
        Caja_text = "'" & "01-" & i & "'"
        BARRA = "'" & "01" & Format(i, "0000000000") & "'"
        Sql = " INSERT INTO COMPU86001_96000 ( [CAJA], [CAJA_DIGITO], [BARRA] )"
        Sql = Sql & " values(" & Caja & "," & Caja_text & "," & BARRA & ")"
        ConEtiquetas.Execute Sql
Next




End Sub


Private Sub Command94_Click()
'Insert Top(1000)
'Into basasql.dbo.UGARTEREFERENCIAS(Cliente, Caja, Descripcion, FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA)
'VALUES     (,,,,,,)



End Sub

Private Sub Command95_Click()


Dim rsCajas As New ADODB.Recordset
Dim rsR As New ADODB.Recordset
Dim Sql As String
Dim SqlCajas As String
Dim Descripcion As String
Dim FECHA_DESDE As String
Dim FECHA_HASTA As String
Dim NRO_DESDE As Long
Dim NRO_HASTAE As Long
Dim NRO_HASTA As Long
Dim LETRA_DESDE As String
Dim LETRA_HASTA As String



Sql = " SELECT     CAJA, CLIENTE, CONTROL"
Sql = Sql & " From basasql.dbo.FONDOCAJAS28112013 "
Sql = Sql & "  Where (Control Is Null)"
Sql = Sql & "  ORDER BY CAJA"
SqlCajas = Sql

Set rsCajas = New ADODB.Recordset


rsCajas.Open SqlCajas, strConBasa
Dim C As Integer

'''    Do While Not rsCajas.EOF
'''
'''            Sql = " SELECT CONTROLFONDO.CAJA, CONTROLFONDO.LEGAJO, CONTROLFONDO.ARCHIVO, LEGAJOS.NRO_DESDE, LEGAJOS.NRO_HASTA, LEGAJOS.LETRA_DESDE,"
'''            Sql = Sql & vbCrLf & " LEGAJOS.LETRA_HASTA, CONVERT(char, LEGAJOS.FECHA_DESDE, 103) AS FECHA_DESDE, CONVERT(CHAR, LEGAJOS.FECHA_HASTA, 103) AS FECHA_HASTA,"
'''            Sql = Sql & vbCrLf & " LEGAJOS.DESCRIPCION ,  LEGAJOS.COD_INDICE "
'''            Sql = Sql & vbCrLf & " FROM CONTROLFONDO INNER JOIN"
'''            Sql = Sql & vbCrLf & " LEGAJOS ON CONTROLFONDO.LEGAJO = LEGAJOS.ID_CLIENTE_LEGAJO"
'''            Sql = Sql & vbCrLf & " Where (LEGAJOS.COD_CLIENTE = 49)"
'''            Sql = Sql & vbCrLf & " AND CONTROLFONDO.CAJA = " & rsCajas!Caja
'''            Sql = Sql & vbCrLf & " ORDER BY CONTROLFONDO.CAJA"
'''            Set rsR = New ADODB.Recordset
'''            rsR.Open Sql, strConBasa
'''
'''             Do While Not rsR.EOF
'''
'''             If IsNull(rsR!DESCRIPCION) Then
'''             DESCRIPCION = "NULL"
'''             Else
'''             DESCRIPCION = CStr(rsR!DESCRIPCION)
'''             End If
'''
'''             If IsNull(rsR!FECHA_DESDE) Then
'''             FECHA_DESDE = "Null"
'''             Else
'''             FECHA_DESDE = rsR!FECHA_DESDE
'''             End If
'''
'''             If IsNull(rsR!FECHA_HASTA) Then
'''
'''             FECHA_HASTA = "NULL"
'''             Else
'''             FECHA_HASTA = rsR!FECHA_HASTA
'''             End If
'''
'''
'''             If IsNull(rsR!NRO_DESDE) Then
'''                NRO_DESDE = 0
'''             Else
'''             NRO_DESDE = rsR!NRO_DESDE
'''             End If
'''
'''            If IsNull(rsR!NRO_DESDE) Then
'''            NRO_DESDE = 0
'''            Else
'''
'''            NRO_DESDE = rsR!NRO_DESDE
'''            End If
'''
'''
'''            If IsNull(rsR!LETRA_DESDE) Then
'''                LETRA_DESDE = "NULL"
'''            Else
'''            LETRA_DESDE = rsR!LETRA_DESDE
'''            End If
'''
'''
'''            If IsNull(rsR!LETRA_HASTA) Then
'''                LETRA_HASTA = "NUll"
'''            Else
'''            LETRA_HASTA = rsR!LETRA_HASTA
'''            End If
'''
'''
'''
'''
'''
'''
'''                c = FONDOREFERENCIAS28112013(rsR!Legajo, rsR!COD_INDICE, LETRA_DESDE, LETRA_HASTA _
'''                , NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION _
'''                , rsR!Caja, 49, "Lectura Miguel")
'''
'''                FONDOCAJAS28112013 rsR!Caja, "Lectura Miguel"
'''
'''                rsR.MoveNext
'''             Loop
'''
'''
'''
'''        rsCajas.MoveNext
'''    Loop
    

'''Rem legajos
'''Do While Not rsCajas.EOF
'''
'''
'''
'''
'''            Sql = " SELECT     ID_CLIENTE_LEGAJO, COD_INDICE, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION, NRO_CAJA,"
'''            Sql = Sql & vbCrLf & "          COD_CLIENTE"
'''Sql = Sql & vbCrLf & " From basasql.dbo.LEGAJOS"
'''Sql = Sql & vbCrLf & " Where (COD_CLIENTE = 49)"
'''Sql = Sql & vbCrLf & " And NRO_CAJA = " & rsCajas!Caja
'''
'''
'''            Set rsR = New ADODB.Recordset
'''            rsR.Open Sql, strConBasa
'''
'''             Do While Not rsR.EOF
'''
'''             If IsNull(rsR!DESCRIPCION) Then
'''             DESCRIPCION = "NULL"
'''             Else
'''             DESCRIPCION = CStr(rsR!DESCRIPCION)
'''             End If
'''
'''             If IsNull(rsR!FECHA_DESDE) Then
'''             FECHA_DESDE = "Null"
'''             Else
'''             FECHA_DESDE = rsR!FECHA_DESDE
'''             End If
'''
'''             If IsNull(rsR!FECHA_HASTA) Then
'''
'''             FECHA_HASTA = "NULL"
'''             Else
'''             FECHA_HASTA = rsR!FECHA_HASTA
'''             End If
'''
'''
'''             If IsNull(rsR!NRO_DESDE) Then
'''                NRO_DESDE = 0
'''             Else
'''             NRO_DESDE = rsR!NRO_DESDE
'''             End If
'''
'''            If IsNull(rsR!NRO_DESDE) Then
'''            NRO_DESDE = 0
'''            Else
'''
'''            NRO_DESDE = rsR!NRO_DESDE
'''            End If
'''
'''
'''            If IsNull(rsR!LETRA_DESDE) Then
'''                LETRA_DESDE = "NULL"
'''            Else
'''            LETRA_DESDE = rsR!LETRA_DESDE
'''            End If
'''
'''
'''            If IsNull(rsR!LETRA_HASTA) Then
'''                LETRA_HASTA = "NUll"
'''            Else
'''                LETRA_HASTA = rsR!LETRA_HASTA
'''            End If
'''
'''
'''
'''
'''
'''
'''                c = FONDOREFERENCIAS28112013(rsR!ID_CLIENTE_LEGAJO, rsR!COD_INDICE, LETRA_DESDE, LETRA_HASTA _
'''                , NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION _
'''                , rsR!NRO_CAJA, 49, "Legajos")
'''
'''                FONDOCAJAS28112013 rsR!NRO_CAJA, "Legajos"
'''
'''                rsR.MoveNext
'''             Loop


Do While Not rsCajas.EOF

          


Sql = "  SELECT     LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION, NRO_CAJA, COD_CLIENTE, INDICE"
  Sql = Sql & vbCrLf & " From basasql.dbo.REFERENCIAS"
  Sql = Sql & vbCrLf & " Where (COD_CLIENTE = 49)"
  Sql = Sql & vbCrLf & " And NRO_CAJA = " & rsCajas!Caja


            Set rsR = New ADODB.Recordset
            rsR.Open Sql, strConBasa

             Do While Not rsR.EOF

             If IsNull(rsR!Descripcion) Then
             Descripcion = "NULL"
             Else
             Descripcion = Mid(CStr(rsR!Descripcion), 1, 200)
             End If

             If IsNull(rsR!FECHA_DESDE) Then
             FECHA_DESDE = "Null"
             Else
             FECHA_DESDE = rsR!FECHA_DESDE
             End If

             If IsNull(rsR!FECHA_HASTA) Then

             FECHA_HASTA = "NULL"
             Else
             FECHA_HASTA = rsR!FECHA_HASTA
             End If


             If IsNull(rsR!NRO_DESDE) Then
                NRO_DESDE = 0
             Else
             NRO_DESDE = rsR!NRO_DESDE
             End If

            If IsNull(rsR!NRO_DESDE) Then
            NRO_DESDE = 0
            Else

            NRO_DESDE = rsR!NRO_DESDE
            End If


            If IsNull(rsR!LETRA_DESDE) Then
                LETRA_DESDE = "NULL"
            Else
                LETRA_DESDE = Trim(Mid(rsR!LETRA_DESDE, 1, 199))
            End If


            If IsNull(rsR!LETRA_HASTA) Then
                LETRA_HASTA = "NUll"
            Else
                LETRA_HASTA = Trim(Mid(rsR!LETRA_HASTA, 1, 199))
            End If






                C = FONDOREFERENCIAS28112013(0, rsR!Indice, LETRA_DESDE, LETRA_HASTA _
                , NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, Descripcion _
                , rsR!NRO_CAJA, 49, "Referencias")

                FONDOCAJAS28112013 rsR!NRO_CAJA, "Referencias"

                rsR.MoveNext
             Loop




        rsCajas.MoveNext
    Loop











End Sub

Private Sub Command96_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String

Sql = " SELECT      CAJASCONLEGAJOSENCONSULTA02012014.COD_CLIENTE, CAJASCONLEGAJOSENCONSULTA02012014.NRO_CAJA,"
 Sql = Sql & "                     CAJASCONLEGAJOSENCONSULTA02012014.ESTADO, CAJASCONLEGAJOSENCONSULTA02012014.ID_CONTENEDOR, LEGAJOS.COD_ESTADO,"
 Sql = Sql & "                      LEGAJOS.ID_LEGAJO"
 Sql = Sql & " FROM         CAJASCONLEGAJOSENCONSULTA02012014 INNER JOIN"
  Sql = Sql & "                      LEGAJOS ON CAJASCONLEGAJOSENCONSULTA02012014.COD_CLIENTE = LEGAJOS.COD_CLIENTE AND"
    Sql = Sql & "                    CAJASCONLEGAJOSENCONSULTA02012014.NRO_CAJA = LEGAJOS.NRO_CAJA"
 Sql = Sql & " Where (LEGAJOS.Cod_Estado = 2)"

rs.Open Sql, strConBasa

Do While Not rs.EOF
    
    
    Sql = " Update LEGAJOS"
    Sql = Sql & " Set Cod_Estado = 9"
    Sql = Sql & " Where (Cod_Estado = 2) "
    Sql = Sql & " And ID_LEGAJO =" & rs!ID_LEGAJO
    ExecutarSql Sql

    rs.MoveNext
Loop





End Sub

Private Sub Command97_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
        Sql = "SELECT CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_INDICE, CAJAS.ID_CAJA, SUBSTRING(REFERENCIAS.INDICE, 1, 6) AS Expr1, INDICES.DESCRIPCION,"
        Sql = Sql & " INDICES.ID as IDINDICES"
        Sql = Sql & " FROM CAJAS INNER JOIN"
        Sql = Sql & " REFERENCIAS ON CAJAS.FK_CLIENTE = REFERENCIAS.COD_CLIENTE AND CAJAS.NRO_CAJA = REFERENCIAS.NRO_CAJA INNER JOIN"
        Sql = Sql & " INDICES ON SUBSTRING(REFERENCIAS.INDICE, 1, 3) = INDICES.INDICE AND REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE"
        Sql = Sql & " GROUP BY CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_INDICE, CAJAS.ID_CAJA, SUBSTRING(REFERENCIAS.INDICE, 1, 6), INDICES.DESCRIPCION, INDICES.ID"
        Sql = Sql & " HAVING (CAJAS.FK_CLIENTE = 4) AND (SUBSTRING(REFERENCIAS.INDICE, 1, 6) LIKE '004%')"
        Sql = Sql & " ORDER BY CAJAS.NRO_CAJA"
 
 
'Sql = " SELECT     CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_INDICE, CAJAS.FK_ESTADO, LEGAJOS.NRO_CAJA AS Expr1, LEGAJOS.COD_INDICE,"
'Sql = Sql & "                      INDICES.Descripcion , INDICES.ID as IDINDICES "
'Sql = Sql & " FROM         CAJAS INNER JOIN"
'Sql = Sql & "                      LEGAJOS ON CAJAS.NRO_CAJA = LEGAJOS.NRO_CAJA AND CAJAS.FK_CLIENTE = LEGAJOS.COD_CLIENTE INNER JOIN"
'Sql = Sql & "                      INDICES ON SUBSTRING( LEGAJOS.COD_INDICE,1,6) = INDICES.INDICE AND LEGAJOS.COD_CLIENTE = INDICES.COD_CLIENTE"
'Sql = Sql & " GROUP BY CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA, CAJAS.FK_INDICE, CAJAS.FK_ESTADO, LEGAJOS.NRO_CAJA, LEGAJOS.COD_INDICE,"
'Sql = Sql & "                      INDICES.Descripcion , INDICES.ID"
'Sql = Sql & " Having (CAJAS.FK_CLIENTE = 4) And (CAJAS.FK_Indice Is Null)"
'
 rs.Open Sql, strConBasa
 
 Do While Not rs.EOF
    Sql = " Update basasql.dbo.CAJAS"
    Sql = Sql & " SET FK_INDICE =" & rs!IDINDICES
    Sql = Sql & " Where ID_CAJA = " & rs!ID_CAJA
    ExecutarSql Sql
    rs.MoveNext
 Loop
 
 
 
 
End Sub

Private Sub Command98_Click()
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim Sql As String
con.Open strConBasa


Sql = " SELECT     ESTANTERIA, MODULO_V, MODULO_H, ESTADO, COUNT(*) AS Expr1"
Sql = Sql & " From CONTENEDOR"
Sql = Sql & " GROUP BY ESTANTERIA, MODULO_V, MODULO_H, ESTADO"
Sql = Sql & " HAVING      (ESTANTERIA > 5000) AND (ESTADO = 1) AND (COUNT(*) IN (5, 10, 15))"
Sql = Sql & " ORDER BY ESTANTERIA, MODULO_V, MODULO_H"

rs.Open Sql, strConBasa


 Do While Not rs.EOF
Sql = " Update basasql.dbo.CONTENEDOR"
Sql = Sql & "  SET ESTADO =0"
Sql = Sql & " Where Estanteria = " & rs!Estanteria
Sql = Sql & " And MODULO_V = " & rs!Modulo_V
Sql = Sql & " And Modulo_H = " & rs!Modulo_H
Sql = Sql & " And estado = 1"


    con.Execute Sql
    rs.MoveNext
    
Loop

End Sub

Private Sub Command99_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
Dim ConDisc As New ADODB.Connection
ConDisc.Open strConBasa

Sql = " SELECT   ID, NRO_CAJA, COD_CLIENTE, FECHA, ID_CODIGO_DOCUMENTO"
Sql = Sql & vbCrLf & " From basasql.dbo.EXPURGO_DISCO"
Sql = Sql & vbCrLf & "  ORDER BY NRO_CAJA"
rs.Open Sql, strConBasa

Do While Not rs.EOF
        Sql = " Update basasql.dbo.REFERENCIAS"
        Sql = Sql & vbCrLf & " Set COD_CLIENTE = 11972002"
        Sql = Sql & vbCrLf & " Where (COD_CLIENTE = 1197)"
        Sql = Sql & vbCrLf & " And NRO_CAJA =" & rs!NRO_CAJA
    
    ConDisc.Execute Sql
    rs.MoveNext
Loop


End Sub

Private Sub Form_Load()
'Set CONCUSTODIA = New ADODB.Connection
inicio
'
'CONCUSTODIA.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=C:\datas\DATAS.mdb"
End Sub

Public Sub INSERTARCUSTODIA(ID As Long, IDDOCUMENTO As Long, DESDENUMERONUEVO As Long)

Dim Sql As String
Sql = " INSERT INTO DCT0755FINAL ( IDDOCUMENTO, IDCAJA, UBICACION, IDTIPODOCUMENTO, NOMBRETIPODOCUMENTO, IDSUCURSAL, NOMBRESUCURSAL, DESCRIPCION, FECHADESDE, FECHAHASTA, FECHAVENCIMIENTO, DESDENUMERO, HASTANUMERO, IDPLANILLA, FECHAENTREGA, FECHADEVOLUCION, CONTROLSUPERVISOR, IDPERSONALPLANILLERO, IDPERSONALLLENADOR, IDUSUARIOCARGA, DESDENUMERONUEVO )"
Sql = Sql & " SELECT " & ID & " AS ID, DCT0755FINAL.IDCAJA, DCT0755FINAL.UBICACION, DCT0755FINAL.IDTIPODOCUMENTO, DCT0755FINAL.NOMBRETIPODOCUMENTO, DCT0755FINAL.IDSUCURSAL, DCT0755FINAL.NOMBRESUCURSAL, DCT0755FINAL.DESCRIPCION, DCT0755FINAL.FECHADESDE, DCT0755FINAL.FECHAHASTA, DCT0755FINAL.FECHAVENCIMIENTO, DCT0755FINAL.DESDENUMERO, DCT0755FINAL.HASTANUMERO, DCT0755FINAL.IDPLANILLA, DCT0755FINAL.FECHAENTREGA, DCT0755FINAL.FECHADEVOLUCION, DCT0755FINAL.CONTROLSUPERVISOR, DCT0755FINAL.IDPERSONALPLANILLERO, DCT0755FINAL.IDPERSONALLLENADOR, DCT0755FINAL.IDUSUARIOCARGA, " & DESDENUMERONUEVO & "  AS NUEVO "
Sql = Sql & "  From DCT0755FINAL"
Sql = Sql & "  WHERE DCT0755FINAL.IDDOCUMENTO= " & IDDOCUMENTO


CONCUSTODIA.Execute Sql

End Sub


Function ProximoRemito() As Long
  Dim Sql As String
  Dim OraMax As ADODB.Recordset
  Sql = "Select Max(Nro_Remito) Maximo From Remitos_Cuerpo"
  Set OraMax = New ADODB.Recordset
  OraMax.Open Sql, ConActiva, 0, 1
  If IsNull(OraMax("Maximo")) Then ProximoRemito = 1: Exit Function
  ProximoRemito = Val(OraMax("Maximo")) + 1
End Function

'Public Function InsertarRemitoCuerpo(TIPO As Integer, NRO_REM_PROV As String, OBSERVACIONES As String, cantidad As String, COBRAR_FLETE As String) As Long
'    Dim SQL As String
'
'    Dim Operacion As String
'    Dim estado As String
'    Dim Fecha As String
'    Dim id_cliente As String
'    Dim AUDIT_USUARIO As String
'    Dim AUDIT_FECHA As String
'    Dim COD_TIPO_ALMACENAMIENTO As String
'    Dim COD_PERSONAL_ENTREGA As String
'    Dim COD_USUARIO_CLIENTE As Integer
'    Dim NRO_REMITO As Long
'
'
'            Operacion = ctlRemito_Operacion.Valor
'            estado = ctlRemtito_Estado.Valor
'            Fecha = "'" & mskFechaRemito.Text & "'"
'            id_cliente = ctlCliente.Valor
'            AUDIT_USUARIO = ctlResponsable.Valor
'            AUDIT_FECHA = SysDate
'            COD_TIPO_ALMACENAMIENTO = ctlTipo_Elemento.Valor
'            COD_PERSONAL_ENTREGA = ctlResponsable.Valor
'            COD_USUARIO_CLIENTE = ctlClienteFirma.Valor
'            NRO_REMITO = ProximoRemito
'
'            SQL = " INSERT INTO REMITOS_CUERPO"
'            SQL = SQL & vbCrLf & " (NRO_REMITO, NRO_REM_PROV, TIPO, OPERACION, ESTADO,FECHA, ID_CLIENTE, OBSERVACIONES, CANTIDAD,"
'            SQL = SQL & vbCrLf & " AUDIT_USUARIO, AUDIT_FECHA,COD_TIPO_ALMACENAMIENTO, COD_PERSONAL_ENTREGA,COD_USUARIO_CLIENTE , COBRAR_FLETE )"
'            SQL = SQL & vbCrLf & " VALUES (" & NRO_REMITO & ",'" & NRO_REM_PROV & "'," & TIPO & "," & Operacion & "," & estado & ","
'            SQL = SQL & vbCrLf & Fecha & "," & id_cliente & ",'" & UCase(Trim(OBSERVACIONES)) & "'," & cantidad
'            SQL = SQL & vbCrLf & ",'" & AUDIT_USUARIO & "'," & AUDIT_FECHA & "," & COD_TIPO_ALMACENAMIENTO & "," & COD_PERSONAL_ENTREGA & "," & COD_USUARIO_CLIENTE & " ,'" & COBRAR_FLETE & "' )"
'            ExecutarSql SQL
'
'            InsertarRemitoCuerpo = NRO_REMITO
'
'
'
'End Function
'
Public Sub INSERTAR_FACTURA_SUPER(FORMA As String, REQUERIMIENTO As Long, NRO_REMITO As Long, NRO_REM_PROV As String _
  , TIPO As String, fecha As String, OBSERVACIONES As String, cantidad As Long, CANT_IMAGENES As Long, _
  APELLIDO_NOMBRE As String, PROVINCIA As String, Sucursal As String, estado As String, flete As String, HORA_ARCHIVISTA As String, COBRAR As String)
Dim Sql As String

Sql = " INSERT INTO TEM_SUPERVIELLE"
Sql = Sql & vbCrLf & "  (FORMA , REQUERIMIENTO"
Sql = Sql & vbCrLf & " , NRO_REMITO"
Sql = Sql & vbCrLf & " , NRO_REM_PROV"
Sql = Sql & vbCrLf & " , TIPO"
Sql = Sql & vbCrLf & " , FECHA"
Sql = Sql & vbCrLf & " , OBSERVACIONES"
Sql = Sql & vbCrLf & " , CANTIDAD "
Sql = Sql & vbCrLf & " , CANT_IMAGENES"
Sql = Sql & vbCrLf & " , APELLIDO_NOMBRE"
Sql = Sql & vbCrLf & " , PROVINCIA"
Sql = Sql & vbCrLf & " , SUCURSAL"
Sql = Sql & vbCrLf & " , ESTADO"
Sql = Sql & vbCrLf & " , FLETE"
Sql = Sql & vbCrLf & " , HORA_ARCHIVISTA"
Sql = Sql & vbCrLf & " , COBRAR )"
Sql = Sql & vbCrLf & " VALUES "
Sql = Sql & vbCrLf & "  ('" & Trim(FORMA) & "'"
Sql = Sql & vbCrLf & " , " & REQUERIMIENTO
Sql = Sql & vbCrLf & " ," & NRO_REMITO
Sql = Sql & vbCrLf & " ,'" & NRO_REM_PROV & "'"
Sql = Sql & vbCrLf & " ,'" & TIPO & "'"
Sql = Sql & vbCrLf & " ,'" & fecha & "'"
Sql = Sql & vbCrLf & " ,'" & OBSERVACIONES & "'"
Sql = Sql & vbCrLf & " ,'" & cantidad & "'"
Sql = Sql & vbCrLf & " ,'" & CANT_IMAGENES & "'"
Sql = Sql & vbCrLf & " ,'" & APELLIDO_NOMBRE & "'"
Sql = Sql & vbCrLf & " ,'" & PROVINCIA & "'"
Sql = Sql & vbCrLf & " ,'" & Sucursal & "'"
Sql = Sql & vbCrLf & " ,'" & estado & "'"
Sql = Sql & vbCrLf & " ,'" & flete & "'"
Sql = Sql & vbCrLf & " ,'" & HORA_ARCHIVISTA & "'"
Sql = Sql & vbCrLf & " ,'" & COBRAR & "')"

ExecutarSql Sql
End Sub

Public Function Esatodo(Caja As Long, Cliente As Integer) As Boolean
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    
    Sql = " SELECT     ESTADO From CONTENEDOR "
    Sql = Sql & " Where COD_CLIENTE = " & Cliente
    Sql = Sql & " And NRO_CAJA = " & Caja
    rs.Open Sql, strConBasa
    
    If Not rs.EOF Then
  Esatodo = rs!estado
  Else
  Esatodo = 0
  End If
  
    

End Function






'Dim Sql As String
'
'Sql = " INSERT INTO TEM_SUPERVIELLE"
'Sql = Sql & vbCrLf & "  (FORMA , REQUERIMIENTO"
'Sql = Sql & vbCrLf & " , NRO_REMITO"
'Sql = Sql & vbCrLf & " , NRO_REM_PROV"
'Sql = Sql & vbCrLf & " , TIPO"
'Sql = Sql & vbCrLf & " , FECHA"
'Sql = Sql & vbCrLf & " , OBSERVACIONES"
'Sql = Sql & vbCrLf & " , CANTIDAD "
'Sql = Sql & vbCrLf & " , CANT_IMAGENES"
'Sql = Sql & vbCrLf & " , APELLIDO_NOMBRE"
'Sql = Sql & vbCrLf & " , PROVINCIA"
'Sql = Sql & vbCrLf & " , SUCURSAL"
'Sql = Sql & vbCrLf & " , ESTADO"
'Sql = Sql & vbCrLf & " , FLETE"
'Sql = Sql & vbCrLf & " , HORA_ARCHIVISTA"
'Sql = Sql & vbCrLf & " , COBRAR )"
'Sql = Sql & vbCrLf & " VALUES "
'Sql = Sql & vbCrLf & "  ('" & Trim(FORMA) & "'"
'Sql = Sql & vbCrLf & " , " & REQUERIMIENTO
'Sql = Sql & vbCrLf & " ," & NRO_REMITO
'Sql = Sql & vbCrLf & " ,'" & NRO_REM_PROV & "'"
'Sql = Sql & vbCrLf & " ,'" & TIPO & "'"
'Sql = Sql & vbCrLf & " ,'" & Fecha & "'"
'Sql = Sql & vbCrLf & " ,'" & OBSERVACIONES & "'"
'Sql = Sql & vbCrLf & " ,'" & CANTIDAD & "'"
'Sql = Sql & vbCrLf & " ,'" & CANT_IMAGENES & "'"
'Sql = Sql & vbCrLf & " ,'" & APELLIDO_NOMBRE & "'"
'Sql = Sql & vbCrLf & " ,'" & PROVINCIA & "'"
'Sql = Sql & vbCrLf & " ,'" & Sucursal & "'"
'Sql = Sql & vbCrLf & " ,'" & estado & "'"
'Sql = Sql & vbCrLf & " ,'" & Flete & "'"
'Sql = Sql & vbCrLf & " ,'" & HORA_ARCHIVISTA & "'"
'Sql = Sql & vbCrLf & " ,'" & COBRAR & "')"
'
'ExecutarSql Sql
'End Sub


Public Function FONDOREFERENCIAS28112013(ID_CLIENTE_LEGAJO As Long, _
Cod_Indice As String, _
LETRA_DESDE As String, _
LETRA_HASTA As String, _
NRO_DESDE As Long, _
NRO_HASTA As Long, _
FECHA_DESDE As String, _
FECHA_HASTA As String, _
Descripcion As String, _
NRO_CAJA As String, _
COD_CLIENTE As Integer, _
ORIGEN As String) As Integer




Dim Sql As String

Sql = " INSERT INTO FONDOREFERENCIAS28112013"
Sql = Sql & vbCrLf & "(ID_CLIENTE_LEGAJO,"
Sql = Sql & vbCrLf & " COD_INDICE,"
Sql = Sql & vbCrLf & " LETRA_DESDE,"
Sql = Sql & vbCrLf & " LETRA_HASTA,"
Sql = Sql & vbCrLf & " NRO_DESDE,"
Sql = Sql & vbCrLf & " NRO_HASTA, "
Sql = Sql & vbCrLf & " FECHA_DESDE,"
Sql = Sql & vbCrLf & " FECHA_HASTA,"
Sql = Sql & vbCrLf & " DESCRIPCION,"
Sql = Sql & vbCrLf & " NRO_CAJA,"
Sql = Sql & vbCrLf & " COD_CLIENTE,"
Sql = Sql & vbCrLf & " ORIGEN)"
Sql = Sql & vbCrLf & " VALUES  "
Sql = Sql & vbCrLf & "(" & ID_CLIENTE_LEGAJO & ","
Sql = Sql & vbCrLf & "'" & Trim(Cod_Indice) & "',"
Sql = Sql & vbCrLf & "'" & Trim(LETRA_DESDE) & "',"
Sql = Sql & vbCrLf & "'" & Trim(LETRA_HASTA) & "',"
Sql = Sql & vbCrLf & NRO_DESDE & ","
Sql = Sql & vbCrLf & NRO_HASTA & ", "
Sql = Sql & vbCrLf & FechaFormato(Trim(FECHA_DESDE)) & ","
Sql = Sql & vbCrLf & FechaFormato(Trim(FECHA_HASTA)) & ","
Sql = Sql & vbCrLf & "'" & Trim(Descripcion) & "',"
Sql = Sql & vbCrLf & NRO_CAJA & ","
Sql = Sql & vbCrLf & COD_CLIENTE & ","
Sql = Sql & vbCrLf & "'" & Trim(ORIGEN) & "')"

FONDOREFERENCIAS28112013 = ExecutarSql(Sql)


End Function

Public Sub FONDOCAJAS28112013(CAJAS As Long, Control As String)
            Dim Sql As String
            Sql = " Update basasql.dbo.FONDOCAJAS28112013 "
            Sql = Sql & " SET CONTROL ='" & Control & "'"
            Sql = Sql & "  Where Caja = " & CAJAS
            ExecutarSql Sql

End Sub
