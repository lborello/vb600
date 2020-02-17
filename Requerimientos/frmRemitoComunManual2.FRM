VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmRemitoEntrada 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remito"
   ClientHeight    =   8745
   ClientLeft      =   105
   ClientTop       =   210
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   11595
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   51
      Top             =   2280
      Width           =   11175
      Begin VB.TextBox txtHorasArchivista 
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
         Left            =   9240
         TabIndex        =   56
         Text            =   "0"
         Top             =   180
         Width           =   555
      End
      Begin VB.CheckBox chkFleteConsulta 
         Caption         =   "Flete"
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
         Left            =   10080
         TabIndex        =   53
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkFLeteGyC 
         Caption         =   "Flete"
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
         Left            =   2820
         TabIndex        =   52
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "Horas de Archivista"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7380
         TabIndex        =   57
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5700
         TabIndex        =   55
         Top             =   180
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Guarda y Custodia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   240
         TabIndex        =   54
         Top             =   180
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdLimpiarEntradaEnPlanta 
      Caption         =   "Limpiar Entrada Planta"
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
      Left            =   5700
      TabIndex        =   50
      Top             =   8100
      Width           =   2400
   End
   Begin VB.CheckBox chkEstado 
      Caption         =   "No Estado"
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
      Left            =   6240
      TabIndex        =   49
      Top             =   480
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid grdPasar 
      Height          =   975
      Left            =   480
      TabIndex        =   48
      Top             =   3420
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1720
      _Version        =   393216
      Cols            =   10
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
   Begin VB.CommandButton cmd 
      Caption         =   "Remito"
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
      Left            =   6720
      TabIndex        =   47
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton cmdCorreo 
      Caption         =   "Correo"
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
      Left            =   4440
      TabIndex        =   46
      Top             =   8100
      Width           =   1095
   End
   Begin VB.CommandButton cmdLeer 
      Caption         =   "Leer"
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
      Left            =   5820
      TabIndex        =   45
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "..."
      Height          =   315
      Left            =   10800
      TabIndex        =   44
      Top             =   8100
      Width           =   375
   End
   Begin VB.CheckBox ChkNocontrolEntrada 
      Caption         =   "No Control Entrada"
      Enabled         =   0   'False
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
      Left            =   4260
      TabIndex        =   43
      Top             =   480
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.CommandButton cmdColectorLegajos 
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
      Left            =   4920
      TabIndex        =   42
      Top             =   1260
      Width           =   795
   End
   Begin VB.CheckBox chkEnvioPorCorreo 
      Caption         =   "Envio Por correo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2460
      TabIndex        =   41
      Top             =   480
      Width           =   1695
   End
   Begin Controles.cltGenerico ctlRemito_Tipo 
      Height          =   315
      Left            =   9000
      TabIndex        =   37
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
   End
   Begin Controles.cltGenerico ctlResponsable 
      Height          =   315
      Left            =   1260
      TabIndex        =   36
      Top             =   840
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   556
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   315
      Left            =   1260
      TabIndex        =   35
      Top             =   60
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   556
   End
   Begin Controles.ctlClienteUsuario ctlClienteFirma 
      Height          =   315
      Left            =   1200
      TabIndex        =   34
      Top             =   7560
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   556
   End
   Begin VB.CommandButton cmdLimpiarEntrada 
      Caption         =   "Borrar todos los elementos"
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
      Left            =   8220
      TabIndex        =   1
      Top             =   8100
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.TextBox txtElementoHasta 
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
      Left            =   2940
      TabIndex        =   0
      Top             =   1260
      Width           =   915
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Anular"
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
      Left            =   6180
      TabIndex        =   28
      Top             =   7560
      Width           =   1080
   End
   Begin MSMask.MaskEdBox mskFechaRemito 
      Height          =   375
      Left            =   1260
      TabIndex        =   2
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdColector 
      Caption         =   "Cajas"
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
      Left            =   4080
      TabIndex        =   27
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   4440
      TabIndex        =   26
      Top             =   7560
      Width           =   1080
   End
   Begin VB.CommandButton cmdControl 
      Caption         =   "Control "
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
      TabIndex        =   25
      Top             =   7560
      Width           =   1080
   End
   Begin VB.TextBox txtDescripcionConsulta 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5940
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   6420
      Width           =   5475
   End
   Begin VB.TextBox txtDescripcionCustodia 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   6420
      Width           =   5535
   End
   Begin VB.TextBox txtElemento 
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
      Left            =   1260
      TabIndex        =   4
      Top             =   1260
      Width           =   915
   End
   Begin MSMask.MaskEdBox mskRemitoCustodia 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Top             =   1860
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648384
      Enabled         =   0   'False
      MaxLength       =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "0001-000#####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   8580
      TabIndex        =   10
      Top             =   7560
      Width           =   1080
   End
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
      Left            =   9780
      TabIndex        =   7
      Top             =   7560
      Width           =   1080
   End
   Begin MSFlexGridLib.MSFlexGrid grdGuardiayCustodia 
      Height          =   3375
      Left            =   240
      TabIndex        =   9
      Top             =   2940
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   6
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid grdDevolucionConsulta 
      Height          =   3375
      Left            =   5940
      TabIndex        =   12
      Top             =   2940
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   6
      ForeColor       =   33023
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox mskRemitoConsulta 
      Height          =   315
      Left            =   7740
      TabIndex        =   8
      Top             =   1860
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12640511
      MaxLength       =   13
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "0001-000#####"
      PromptChar      =   "_"
   End
   Begin Controles.cltGenerico ctlTipo_Elemento 
      Height          =   315
      Left            =   9000
      TabIndex        =   38
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
   End
   Begin Controles.cltGenerico ctlRemito_Operacion 
      Height          =   315
      Left            =   9000
      TabIndex        =   39
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
   End
   Begin Controles.cltGenerico ctlRemtito_Estado 
      Height          =   315
      Left            =   9000
      TabIndex        =   40
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
   End
   Begin VB.Label lblSector 
      BackColor       =   &H00FFC0C0&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1200
      TabIndex        =   29
      Top             =   8040
      Width           =   3015
   End
   Begin VB.Label Label17 
      Caption         =   "Sector:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   30
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   31
      Top             =   1320
      Width           =   555
   End
   Begin VB.Label Label15 
      Caption         =   "Cliente firma:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   33
      Top             =   7620
      Width           =   675
   End
   Begin VB.Label Label14 
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7920
      TabIndex        =   3
      Top             =   1260
      Width           =   915
   End
   Begin VB.Label Label12 
      Caption         =   "Transporta:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "Remito:"
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
      Left            =   7920
      TabIndex        =   32
      Top             =   180
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Cliente :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   300
      TabIndex        =   22
      Top             =   120
      Width           =   675
   End
   Begin VB.Line Line1 
      X1              =   180
      X2              =   11340
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label Label10 
      Caption         =   "Elemento :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Elemento:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7920
      TabIndex        =   20
      Top             =   540
      Width           =   915
   End
   Begin VB.Label Label9 
      Caption         =   "Operación:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7920
      TabIndex        =   19
      Top             =   900
      Width           =   915
   End
   Begin VB.Label Label11 
      Caption         =   "Remito:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6660
      TabIndex        =   18
      Top             =   1860
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Remito :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   300
      TabIndex        =   17
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblCantidadDevolucion 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   10500
      TabIndex        =   16
      Top             =   1860
      Width           =   855
   End
   Begin VB.Label lblCantidadGuardia 
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   3780
      TabIndex        =   15
      Top             =   1860
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Cantidad :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9540
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Cantidad :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   2820
      TabIndex        =   13
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   300
      TabIndex        =   11
      Top             =   540
      Width           =   555
   End
End
Attribute VB_Name = "frmRemitoEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim RequerimientoRemito As Long
    Dim cajasRepetidas As String
    Dim ElementosEstado As String
    Dim SrtValor As String
    Dim CorreoEnviado As Boolean
    Dim PasoRemito As String
    Dim Error As String
    Dim RAPC As Boolean
    Dim AbortarEntrada As Boolean
    

Private Sub chkEstado_Click()
    If chkEstado.Value = 0 Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
End Sub

Private Sub cmd_Click()
 
  Dim cont As Integer
    Dim VarTexto As String
    cont = 0
    Dim Cliente As Integer
    Dim Caja As Long
    Rem On Error GoTo salir
    Dim sql As String
    grdPasar.Clear
    grdPasar.Rows = 1
    Dim rs As New ADODB.Recordset
    AbortarEntrada = False
    
   On Error GoTo SALIR2:
   
  
    
    If Dir("\\222.15.19.251\basa\Planta\LECTURAS\REMITO ENTRADA\" & Trim(Format(CLng(ctlCliente.Valor), "00000")) & "-" & Trim(Format(CLng(Mid(mskRemitoConsulta.Text, 6)), "0000000")) & ".txt") = "" Then
     If Dir("\\222.15.19.251\basa\Planta\LECTURAS\REMITO ENTRADA\00000" & "-" & Trim(Format(CLng(Mid(mskRemitoConsulta.Text, 6)), "0000000")) & ".txt") = "" Then
        MsgBox "NO existe archivo"
        Exit Sub
        Else
            PasoRemito = "\\222.15.19.251\basa\Planta\LECTURAS\REMITO ENTRADA\00000" & "-" & Trim(Format(CLng(Mid(mskRemitoConsulta.Text, 6)), "0000000")) & ".txt"
      End If
    
    Else
        PasoRemito = "\\222.15.19.251\basa\Planta\LECTURAS\REMITO ENTRADA\" & Trim(Format(CLng(ctlCliente.Valor), "00000")) & "-" & Trim(Format(CLng(Mid(mskRemitoConsulta.Text, 6)), "0000000")) & ".txt"
    End If
    
    
  
     On Error GoTo salir:

    Open PasoRemito For Input As #1
    
salir:
    If Err.Number = 55 Then
      Close #1
    Open PasoRemito For Input As #1
    End If
    
    

    
    Do Until EOF(1)
        Line Input #1, VarTexto
        If VarTexto <> "" Then
          grdPasar.AddItem VarTexto
        End If
    Loop
    Close #1
    
    Error = ""
    
    Dim R As Integer
       If MsgBox("Usted QuiereMarcar el remito en la entrada", vbYesNo) = vbNo Then

        If ctlTipo_Elemento.DESCRIPCION = "LEGAJOS" Then
                For R = 0 To grdPasar.Rows - 1
                    If Mid(grdPasar.TextMatrix(R, 4), 1, 3) = "LEG" Or Mid(grdPasar.TextMatrix(R, 4), 1, 3) = "VLE" Then
                        If AbortarEntrada = False Then
                          CargarGrilla (CLng(grdPasar.TextMatrix(R, 2)))
                        End If
                    End If
                Next
        End If
        
  
        
        If ctlTipo_Elemento.DESCRIPCION = "CAJAS" Then
            For R = 0 To grdPasar.Rows - 1
            
            If Trim(grdPasar.TextMatrix(R, 4)) = "IDDD" Then
            
            sql = "    SELECT     NRO_CAJA, FK_CLIENTE, ID_CAJA, DIGITO_VERIFICADOR"
            sql = sql & " From Cajas"
            sql = sql & " Where ID_CAJA = " & CLng(grdPasar.TextMatrix(R, 2))
            Set rs = New ADODB.Recordset
            rs.Open sql, ConActiva, 0, 1
            CargarGrilla (rs!NRO_CAJA)
            
            End If
            
            
                If Trim(grdPasar.TextMatrix(R, 4)) = "BASA" Or Trim(grdPasar.TextMatrix(R, 4)) = "VBAS" Or Trim(grdPasar.TextMatrix(R, 4)) = "CUST" Or Trim(grdPasar.TextMatrix(R, 4)) = "VCUS" Then
                    If CargarGrilla(CLng(grdPasar.TextMatrix(R, 2))) = False Then
                    Exit Sub
                    End If
                End If
            Next
        End If
  Else
        Clipboard.Clear
        Dim Marca As String
        Dim dato As String
        
        Marca = InputBox("Ingrese el valor a marcar", "Valor", 1)
        
        If Not IsNumeric(Marca) Then
            Exit Sub
        End If
        
        
        Marca = Marca * -1
         If ctlTipo_Elemento.DESCRIPCION = "LEGAJOS" Then
                For R = 0 To grdPasar.Rows - 1
                    If Mid(grdPasar.TextMatrix(R, 4), 1, 3) = "LEG" Then
                        
                        sql = " Update ENTRADA SET COD_ESTADO =" & Marca
                        sql = sql & " Where COD_ESTADO < 1 "
                        sql = sql & " And TIPO = 3"
                        sql = sql & " And Cod_cliente = " & ctlCliente.Valor
                        sql = sql & " And Elemento = " & CLng(grdPasar.TextMatrix(R, 2))
                        dato = dato & CLng(grdPasar.TextMatrix(R, 2)) & vbCrLf
                        ExecutarSql sql
                    End If
                Next
        End If
        If ctlTipo_Elemento.DESCRIPCION = "CAJAS" Then
            For R = 0 To grdPasar.Rows - 1
                  If Trim(grdPasar.TextMatrix(R, 4)) = "BASA" Or Trim(grdPasar.TextMatrix(R, 4)) = "VBAS" Then
                     sql = " Update ENTRADA SET COD_ESTADO =" & Marca
                        sql = sql & " Where COD_ESTADO < 1 "
                        sql = sql & " And TIPO = 0"
                        sql = sql & " And Cod_cliente = " & ctlCliente.Valor
                        sql = sql & " And Elemento = " & CLng(grdPasar.TextMatrix(R, 2))
                        ExecutarSql sql
                         dato = dato & CLng(grdPasar.TextMatrix(R, 2)) & vbCrLf
                End If
            Next
        End If
        Clipboard.SetText dato
  End If
  
  
   If Error <> "" Then
        MsgBox "Errores copiados a memoria"
        Clipboard.Clear
        Clipboard.SetText Error
        End If
  
  Exit Sub
  
SALIR2:
MsgBox Err.Description

End Sub

Private Sub cmdAceptar_Click()
AbortarEntrada = False
ElementosEstado = ""
RAPC = False
Rem On Error GoTo salir:
If lblCantidadGuardia.Caption <> "" Then
         If MsgBox("Referencia administrada por el cliente", vbYesNo) = vbYes Then
         RAPC = True
         Else
         RAPC = False
         End If
         End If
If CorreoEnviado = False And ctlRemito_Operacion.Valor = 0 Then
    If MsgBox("Usted no envio el correo quiere hacerlo ?", vbYesNo + vbDefaultButton2) = vbYes Then
        cmdCorreo_Click
      Else
       MsgBox "No se envio el correo", vbCritical
    End If
    If mskRemitoConsulta.Text = "0001-000_____" Then
        MsgBox "Usted Debe ingresar el Remito de Consulta ", vbCritical
        Exit Sub
    End If
 End If
 
    If chkFleteConsulta.Value = 1 And chkFLeteGyC.Value = 1 Then
        MsgBox "No se puede asignar dos flete para un remito"
        Exit Sub
    End If
        If ctlCliente.Valor = 231 Or ctlCliente.Valor = 163 Or ctlCliente.Valor = 172 Then
            If ctlRemito_Operacion.Valor = 1 And ctlTipo_Elemento.Valor = 3 Then
                If txtHorasArchivista.Text = 0 Then
                    If MsgBox("Ingrese las Horas de archivista de Busqueda de legajos " & vbCrLf & " Usted quiere contunuar ?", vbCritical + vbYesNo) = vbNo Then
                        Exit Sub
                    End If
                End If
            End If
        End If
 
     If Validar_Datos Then
        If lblCantidadGuardia.Caption <> "" Then
         
            GrabarRemitoCustodia
        End If
        If lblCantidadDevolucion.Caption <> "" Then
            
            Select Case ctlRemito_Tipo.Valor
            Case 1, 0
                GrabarRemitoDevolucion
                
            Case 2
                GrabarRemitoCajasVacias
            Case 3
                If "michu" = InputBox("Ingresa la clave de Bajas y controle el remito") Then
                    GrabarRemitoBajas
                Else
                    MsgBox "La baja de cajas no fue realizada"
                    Exit Sub
                End If
            
            Case 4
                Rem GrabarRemitoDevolucionVacias
            End Select
            
            Dim sql As String
                If chkEnvioPorCorreo.Value = 1 Then
                    sql = " Update Requerimiento  "
                    sql = sql & " SET   ENVIOPORCORREO = 1"
                    sql = sql & " Where IDREQUERIMIENTO = " & RequerimientoRemito
                    ExecutarSql sql
                End If
                
            
            
        End If
        
        If CInt(ctlRemito_Operacion.Valor) = 0 Then
             If MsgBox("Marcar Remito", vbYesNo + vbDefaultButton2) = vbYes Then
                      If PasoRemito <> "" Then
                        FileSystem.FileCopy PasoRemito, Mid(PasoRemito, 1, Len(PasoRemito) - 4) & Format(SysDate_DD_MM_YYYY_mm_ss, " ddmmyyyy hhmm") & " Procesado.txt"
                        FileSystem.Kill PasoRemito
                      End If
                 End If
              End If
        End If
        Limpiar
        Rem  txtElementoHasta.Enabled = False
        If ctlRemito_Operacion.Valor = 1 Then
            frmControlEstados.CargarTree
            Unload Me
  
      End If
      
 Exit Sub
 
Rem salir:
 Rem MsgBox Err.Description
 
End Sub

Private Sub cmdActualizar_Click()
'    Dim rs As New ADODB.Recordset
'    Dim sql As String
'
'    sql = " SELECT     REMITOS_DETALLE.NRO_REMITO, REMITOS_DETALLE.HASTA, REMITOS_CUERPO.ID_CLIENTE"
' sql = sql & " FROM         REMITOS_DETALLE INNER JOIN "
' sql = sql & " REMITOS_CUERPO ON REMITOS_DETALLE.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO"
'
'sql = sql & " Where REMITOS_DETALLE.NRO_REMITO = " & InputBox("Ingrese el remito")
'
'
'rs.Open sql, ConActiva, 0, 1
'
'Do While Not rs.EOF
'        sql = " Update Cajas "
'        sql = sql & " SET TIPO_REFERENCIA = '" & Trim(cboReferencia.Text) & "'"
'        sql = sql & " Where FK_CLIENTE = " & rs!ID_CLIENTE
'        sql = sql & " And NRO_CAJA = " & rs!HASTA
'        ExecutarSql sql
'        rs.MoveNext
'Loop
'
'MsgBox "Terminado "

End Sub

Private Sub cmdCancelar_Click()
   Limpiar
End Sub

Private Sub cmdColector_Click()
    Dim rs2 As New ADODB.Recordset
    Dim sql As String
    Dim Lectura As Long
    
    On Error GoTo salir:
    cajasRepetidas = ""
    
    Lectura = InputBox("Por Favor Ingrese el numero de Lectura ", "Lectura", 0)
    
    
    If IsNumeric(Lectura) Then
        
        sql = " SELECT NUMERO_LECTURA, CAJA, CLIENTE, ORDEN From LECTURACOLECTOR "
        sql = sql & "Where CLIENTE < 9000  and NUMERO_LECTURA = " & Lectura
        sql = sql & " ORDER BY ORDEN "
        rs2.Open sql, ConActiva, 0, 1
        Do While Not rs2.EOF
           If CargarGrilla(CLng(rs2!Caja)) = True Then
                rs2.MoveNext
           Else
                Exit Sub
           End If
           
        Loop
    End If
    
    If cajasRepetidas <> "" Then
        Clipboard.Clear
        MsgBox "Existe Elementos Reperidos " & cajasRepetidas & "Los elementos fueron copiados a Excel "
        Clipboard.SetText cajasRepetidas
    End If
    
    
Exit Sub

salir:


    MsgBox Err.Description
    
        
End Sub


Private Sub cmdColectorLegajos_Click()
  Dim rs2 As New ADODB.Recordset
    Dim sql As String
    Dim Lectura As Integer
    Lectura = InputBox("Por Favor Ingrese el numero de Lectura ", "Lectura")
        
        sql = "  SELECT     ID_Lectura_Legajo, Cod_Legajo_cliente, Cliente"
sql = sql & " From LECTURA_LEGAJO"
sql = sql & " Where ID_Lectura_Legajo = " & Lectura
        
        rs2.Open sql, ConActiva, 0, 1
        Do While Not rs2.EOF
            CargarGrilla CLng(rs2!Cod_Legajo_cliente)
            rs2.MoveNext
        Loop
End Sub

Private Sub cmdControl_Click()
    Dim sql As String
    MousePointer = 11
    
    
    
    Select Case InputBox("INGRESE LA PROVINCIA " & "0-TODOS // 1- MENDOZA //  2-CORDOBA // 3-SAN LUIS // 10-DEPOSITOS")
    Case 0
        sql = "SELECT  *  From ENTRADA Where COD_ESTADO < 1    Order By  FECHA DESC , COD_CLIENTE,TIPO,ELEMENTO"
    Case 1
        sql = "SELECT  *  From ENTRADA Where COD_ESTADO < 1  and COD_PERSONAL IN (19, 49, 21,70, 31, 47, 12,82,17,66,26 , 64,69 )  AND   DESCRIPCION IS NULL   Order By COD_CLIENTE,TIPO,ELEMENTO "
    Case 2
        sql = "SELECT  *  From ENTRADA Where   COD_ESTADO < 1  and COD_PERSONAL IN (38) Order By COD_CLIENTE,TIPO,ELEMENTO"
    Case 3
        sql = "SELECT  *  From ENTRADA Where COD_ESTADO < 1  and COD_PERSONAL IN ( 37) Order By COD_CLIENTE,TIPO,ELEMENTO"
   Case 10
        sql = "SELECT  *  From ENTRADA Where COD_ESTADO < -9    Order By  descripcion,  FECHA DESC , COD_CLIENTE,TIPO,ELEMENTO"
    End Select
    
    
    frmReportes.ImprimirReporte PasoReportes & "ControlElementos.rpt", sql, True
    MousePointer = 0
End Sub

Private Sub cmdCorreo_Click()

Dim correo As String
Dim R As Integer
Dim c As Integer
Dim sql As String
Dim n As Integer
Dim rs As New ADODB.Recordset
Dim Mail As String
Dim Nombre As String

10:

Clipboard.Clear

    correo = Format(ctlCliente.Valor, "00000") & " " & UCase(Trim(ctlCliente.DESCRIPCION)) & vbCr
    correo = vbCrLf & correo & vbCrLf & "    Estimado  " & ctlClienteFirma.DESCRIPCION

        n = 1
        
      If grdDevolucionConsulta.Rows > 1 Then
      If grdDevolucionConsulta.TextMatrix(1, 1) <> "" Then
      correo = correo & vbCrLf & "     Por intermedio de la presente le informamos el movimiento de " & ctlTipo_Elemento.DESCRIPCION
     If ctlRemito_Tipo.Valor = 3 Then
        correo = correo & "  en concepto de " & ctlRemito_Tipo.DESCRIPCION & "  en nuestro sistema "
     Else
       correo = correo & "  en concepto de entrada en nuestro sistema "
     End If
      
        correo = correo & vbCrLf & " Con fecha " & mskFechaRemito.Text & "  el número de remito es " & mskRemitoConsulta.Text
        correo = correo & vbCrLf & ctlTipo_Elemento.DESCRIPCION & " son: " & vbCrLf
      
        With grdDevolucionConsulta
                        For R = 1 To .Rows - 1
                           For c = 1 To .Cols - 1
                             If .TextMatrix(R, c) <> "" Then
                                  If ctlTipo_Elemento.DESCRIPCION = "LEGAJOS" Then
                                     correo = correo & Str(n) & "-" & descrpLegajos(ctlCliente.Valor, .TextMatrix(R, c)) & vbCrLf
                                     n = n + 1
                                    Else
                                        correo = correo & Str(n) & "-" & .TextMatrix(R, c) & vbCrLf
                                        n = n + 1
                                    End If
                                    
                                End If
                             Next
                         Next
                   End With
                   
        End If
        
        End If
        
        n = 1
          
          With grdGuardiayCustodia
        
        
        If .Rows > 1 Then
      
       If .TextMatrix(1, 1) <> "" Then
        correo = correo & vbCrLf & "     Por intermedio de la presente le informamos el ingreso de " & ctlTipo_Elemento.DESCRIPCION & "  por GUARDA Y CUSTODIA a nuestro sistema "
        correo = correo & vbCrLf & "Con fecha " & mskFechaRemito.Text & "  el número de remito es " & mskRemitoCustodia.Text
        correo = correo & vbCrLf & ctlTipo_Elemento.DESCRIPCION & " ingresadas son " & vbCrLf
      
      
                        For R = 1 To .Rows - 1
                           For c = 1 To .Cols - 1
                           
                                          
                           
                                If .TextMatrix(R, c) <> "" Then
                                    
                                   correo = correo & vbCrLf & Str(n) & "-" & .TextMatrix(R, c)
                                    n = n + 1
                                End If
                             Next
                             Rem  correo = correo & vbCrLf
                         
                         Next
                  
            End If
                
        End If
        
                   
                   
              End With
                   
     correo = correo & vbCrLf & txtDescripcionConsulta.Text
     If RAPC = True Then
         correo = correo & vbCrLf & " LAS CAJAS INGRESADAS SOLO PODRAN SER SOLICITADAS POR NUMERO DE CAJA"
     End If
     
    correo = correo & vbCrLf & "Muchas Gracias por su colaboración."
    Clipboard.SetText correo
    
MsgBox "Los datos fueron copiados"

If ctlClienteFirma.Valor Then

sql = " SELECT        ID_CLIENTEUSUARIO, APELLIDO_NOMBRE, CORREO"
sql = sql & " FROM CLIENTEUSUARIO"
sql = sql & " Where ID_CLIENTEUSUARIO = " & ctlClienteFirma.Valor

Set rs = New ADODB.Recordset

rs.Open sql, ConActiva, 0, 1
If Not rs.EOF Then
 If Trim(rs!correo) = "" Or IsNull(rs!correo) Then
 
   If MsgBox("No tiene correo" & vbCrLf & "Usted quiere cargarlo", vbYesNo) = vbYes Then
   sql = " Update CLIENTEUSUARIO"
sql = sql & " SET        CORREO ='" & Trim(InputBox("Ingrese el correo")) & "'"
sql = sql & " Where ID_CLIENTEUSUARIO = " & ctlClienteFirma.Valor
ExecutarSql sql
GoTo 10:
   
   End If
   
    Exit Sub
    
Else
    Mail = Trim(rs!correo)
    Nombre = Trim(rs!APELLIDO_NOMBRE)
 End If
 
 
Else
 MsgBox "NO existe"
 Exit Sub
End If



End If
    
    
    
    
    
If MsgBox("QUIERE ENVIAR EL CORREO", vbYesNo) = vbYes Then

    SendMail Mail, "Remitos de archivos ", correo
CorreoEnviado = True
End If


 
 
End Sub

Private Sub cmdImprimir_Click()
    Dim i As Long
    On Error GoTo salir
    i = InputBox("Ingrese el Nº de remito", "Remito")
   If InputBox("Ingrese el tipo 1- caja , 3 - Libros") = 1 Then
   
    ImprimirRemitoCaja i
    Else
    ImprimirRemitoLibros 1
    
    End If
    
salir:
End Sub

Private Sub Command1_Click()




Dim rs As New ADODB.Recordset
Dim sql As String
Dim R As Integer




End Sub



Private Sub ReconstrucionMovimiento(cod_remito_1 As Long)
 Dim sql As String
 Dim rsRemito As New ADODB.Recordset
 Dim Inicio, Final As Long
 Dim Cod_REMITO, COD_TIPO, COD_OPERACION As String
 Dim COD_TIPO_ALMACENAMIENTO, Elemento, Cod_cliente As String
 Dim Fecha, ANULADO As String
 
    sql = "     SELECT REMITOS_CUERPO.NRO_REMITO,"
    sql = sql & "   REMITOS_CUERPO.TIPO, REMITOS_CUERPO.OPERACION,"
    sql = sql & "   REMITOS_CUERPO.ID_CLIENTE,"
    sql = sql & "   REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO,"
    sql = sql & "   REMITOS_DETALLE.DESDE, REMITOS_DETALLE.HASTA,"
    sql = sql & "   REMITOS_DETALLE.TIPO_ALMACENADO, REMITOS_CUERPO.ANULADO,REMITOS_CUERPO.FECHA"
    sql = sql & "   From REMITOS_DETALLE, REMITOS_CUERPO"
    sql = sql & "   Where REMITOS_DETALLE.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO"
    sql = sql & "   and REMITOS_CUERPO.NRO_REMITO > " & cod_remito_1
    sql = sql & "   ORDER BY REMITOS_CUERPO.NRO_REMITO"
    rsRemito.Open sql, ConActiva, 0, 1
Do While Not rsRemito.EOF

If Not IsNull(rsRemito!DESDE) Then

    If rsRemito!DESDE = rsRemito!HASTA Then
        Cod_REMITO = rsRemito!NRO_REMITO
        COD_TIPO = rsRemito!TIPO
        COD_OPERACION = rsRemito!OPERACION
        COD_TIPO_ALMACENAMIENTO = rsRemito!TIPO_ALMACENADO
        Elemento = rsRemito!DESDE
        Cod_cliente = rsRemito!ID_CLIENTE
        Fecha = "'" & Format(rsRemito!Fecha, "dd/mm/yyyy") & "'"
        If IsNull(rsRemito!ANULADO) Then
            ANULADO = "Null"
        Else
            ANULADO = rsRemito!ANULADO
        End If
        sql = " INSERT INTO MOVIMIENTOS_ELEMENTOS_PRU"
        sql = sql & " (COD_REMITO, COD_TIPO, COD_OPERACION,"
        sql = sql & " COD_TIPO_ALMACENAMIENTO, ELEMENTO, COD_CLIENTE,"
        sql = sql & " FECHA,ANULADO)"
        sql = sql & " VALUES ("
        sql = sql & Cod_REMITO & "," & COD_TIPO & "," & COD_OPERACION & ","
        sql = sql & COD_TIPO_ALMACENAMIENTO & "," & Elemento & "," & Cod_cliente & ","
        sql = sql & Fecha & "," & ANULADO & "  )"
        ExecutarSql sql
     Else
        
        For Inicio = rsRemito!DESDE To rsRemito!HASTA
                Cod_REMITO = rsRemito!NRO_REMITO
                COD_TIPO = rsRemito!TIPO
                COD_OPERACION = rsRemito!OPERACION
                COD_TIPO_ALMACENAMIENTO = rsRemito!TIPO_ALMACENADO
                Elemento = Inicio
                Cod_cliente = rsRemito!ID_CLIENTE
                Fecha = "'" & rsRemito!Fecha & "'"
                If IsNull(rsRemito!ANULADO) Then
                    ANULADO = "Null"
                Else
                    ANULADO = rsRemito!ANULADO
                End If
                sql = " INSERT INTO MOVIMIENTOS_ELEMENTOS_PRU"
                sql = sql & "      (COD_REMITO, COD_TIPO, COD_OPERACION,"
                sql = sql & "    COD_TIPO_ALMACENAMIENTO, ELEMENTO, COD_CLIENTE,"
                sql = sql & "    FECHA,ANULADO)"
                sql = sql & "   VALUES ("
                sql = sql & Cod_REMITO & "," & COD_TIPO & "," & COD_OPERACION & ","
                sql = sql & COD_TIPO_ALMACENAMIENTO & "," & Elemento & "," & Cod_cliente & ","
                sql = sql & Fecha & "," & ANULADO & "  )"
                ExecutarSql sql
       Next
     
     End If
End If

    rsRemito.MoveNext
Loop

End Sub

Private Sub Command2_Click()
'Dim i As Long
'For i = 340 To 709
'    InsertarRemitoDetalle CLng(46426), "BASA", CStr(SysDate), CLng(i), CLng(i), CInt(1)
'Next
End Sub

Private Sub Command3_Click()
    ReconstrucionMovimiento 46707
End Sub


Private Sub cmdLeer_Click()
Dim L As String
Dim i As Integer
Dim dato As String
Dim datoInicio As String
Dim espacio As Integer
Dim comienzo As Integer

On Error GoTo salir

L = Clipboard.GetText
L = Trim(L)
comienzo = 1
espacio = 1
L = Replace(L, vbCrLf, "&")
Dim Inicio As Integer
Dim Fin As Integer

Inicio = 1
Fin = 1

For i = 1 To Len(L)
    
    If Mid(L, i, 1) = "&" Then
        dato = Mid(L, Inicio, i - Inicio)
        Inicio = i + 1
         If CargarGrilla(CLng(dato)) = False Then
            Exit Sub
         End If
         
   End If
    
    
'
'    datoInicio = Mid(L, comienzo)
'    espacio = InStr(datoInicio, "&")
'    dato = Mid(datoInicio, 1, espacio - 1)
'    comienzo = espacio + 1
    Rem CargarGrilla CStr(dato)
Next
dato = Mid(L, Inicio, i - Inicio)

If IsNumeric(dato) Then

CargarGrilla CLng(dato)

End If

Exit Sub

salir:
MsgBox Err.Description
End Sub

Private Sub cmdLimpiarEntrada_Click()
Dim sql As String
If MsgBox("Esta usted seguro de eliminar los elementos pedientes de entrada", vbYesNo) = vbYes Then
    MsgBox MDIfrmInicio.StaInicio.Panels(2).Text
     If MDIfrmInicio.StaInicio.Panels(2).Text = 31 Or MDIfrmInicio.StaInicio.Panels(2).Text = 70 Then
        sql = " UPDATE ENTRADA SET COD_ESTADO = 3, FECHA_LIMPIADO = " & SysDate & " Where (COD_ESTADO < 1) and COD_PERSONAL IN (19, 70, 31,12, 47)"
        ExecutarSql sql
      End If
      
      If MDIfrmInicio.StaInicio.Panels(2).Text = 37 Then
        sql = " UPDATE ENTRADA SET COD_ESTADO = 3, FECHA_LIMPIADO = " & SysDate & " Where (COD_ESTADO < 1) and COD_PERSONAL IN (37)"
        ExecutarSql sql
      End If
      
      
      If MDIfrmInicio.StaInicio.Panels(2).Text = 38 Then
        sql = " UPDATE ENTRADA SET COD_ESTADO = 3, FECHA_LIMPIADO = " & SysDate & "  Where (COD_ESTADO < 1) and COD_PERSONAL IN (38)"
        ExecutarSql sql
      End If
    
End If

End Sub

Private Sub cmdLimpiarEntradaEnPlanta_Click()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim ACTUALIZACIONfECHA As String


sql = " SELECT     ENTRADA.ID_ENTRADA, ENTRADA.FECHA_LIMPIADO, ENTRADA.COD_ESTADO"
sql = sql & " FROM         ENTRADA INNER JOIN"
sql = sql & " CONTENEDOR ON ENTRADA.COD_CLIENTE = CONTENEDOR.COD_CLIENTE AND ENTRADA.ELEMENTO = CONTENEDOR.NRO_CAJA"
sql = sql & " Where (ENTRADA.COD_ESTADO  = 0) And (ENTRADA.TIPO = 0) And (CONTENEDOR.ESTADO = 2)"
sql = sql & " ORDER BY ENTRADA.ID_ENTRADA DESC"
ACTUALIZACIONfECHA = SysDate_mm_ss
 rs.Open sql, ConActiva, 0, 1
Do While Not rs.EOF

sql = " Update ENTRADA"
sql = sql & "  SET   FECHA_LIMPIADO =" & ACTUALIZACIONfECHA & ", COD_ESTADO =3"
sql = sql & " Where ID_ENTRADA = " & rs!ID_ENTRADA
 ExecutarSql sql
    
    rs.MoveNext
Loop





sql = "  SELECT        ENTRADA.ID_ENTRADA, ENTRADA.FECHA_LIMPIADO, ENTRADA.COD_ESTADO, CONTENEDOR.ESTADO, ENTRADA.COD_CLIENTE, ENTRADA.ELEMENTO"
sql = sql & " FROM            ENTRADA LEFT OUTER JOIN"
 sql = sql & "                        CONTENEDOR ON ENTRADA.COD_CLIENTE = CONTENEDOR.COD_CLIENTE AND ENTRADA.ELEMENTO = CONTENEDOR.NRO_CAJA"
sql = sql & " Where (ENTRADA.COD_ESTADO = 0) And (ENTRADA.TIPO = 0) And (CONTENEDOR.ESTADO Is Null)"
sql = sql & " ORDER BY ENTRADA.ID_ENTRADA DESC"


ACTUALIZACIONfECHA = SysDate_mm_ss
Set rs = New ADODB.Recordset
 rs.Open sql, ConActiva, 0, 1
Do While Not rs.EOF

sql = " Update ENTRADA"
sql = sql & "  SET   FECHA_LIMPIADO =" & ACTUALIZACIONfECHA & ", COD_ESTADO =3"
sql = sql & " Where ID_ENTRADA = " & rs!ID_ENTRADA
 ExecutarSql sql
    
    rs.MoveNext
Loop





sql = "  SELECT     ENTRADA.ID_ENTRADA, ENTRADA.FECHA_LIMPIADO, ENTRADA.COD_ESTADO, LEGAJOS.COD_ESTADO AS Expr1"
sql = sql & "  FROM         ENTRADA INNER JOIN"
sql = sql & "  LEGAJOS ON ENTRADA.ELEMENTO = LEGAJOS.ID_CLIENTE_LEGAJO AND ENTRADA.COD_CLIENTE = LEGAJOS.COD_CLIENTE"
sql = sql & "  Where (ENTRADA.COD_ESTADO < 1 ) And (ENTRADA.TIPO = 3) And (LEGAJOS.COD_ESTADO = 2)"
sql = sql & "  ORDER BY ENTRADA.ID_ENTRADA DESC"


Set rs = New ADODB.Recordset
rs.Open sql, ConActiva, 0, 1
Do While Not rs.EOF

sql = " Update ENTRADA"
sql = sql & "  SET   FECHA_LIMPIADO =" & ACTUALIZACIONfECHA & ", COD_ESTADO =3"
sql = sql & " Where ID_ENTRADA = " & rs!ID_ENTRADA
 ExecutarSql sql
    
    rs.MoveNext
Loop

MsgBox "TERMINADO"
End Sub

Private Sub Command4_Click()
    AnularRemito InputBox("Numero de remito para anular")
End Sub

Private Sub Command5_Click()

Dim sql As String
Dim rs As New ADODB.Recordset
sql = " SELECT REMITOELEMENTO.COD_REMITO, REMITOELEMENTO.COD_TIPO_ALMACENAMIENTO"
sql = sql & " From REMITOELEMENTO, REMITOS_CUERPO"
sql = sql & " Where REMITOELEMENTO.Cod_REMITO = REMITOS_CUERPO.NRO_REMITO"
sql = sql & " AND    (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO IS NULL) "
sql = sql & " AND    (REMITOELEMENTO.COD_TIPO_ALMACENAMIENTO <> 0) "
sql = sql & " ORDER BY REMITOELEMENTO.COD_REMITO "
rs.Open sql, ConActiva, 0, 1
Do While Not rs.EOF
    sql = " Update REMITOS_CUERPO SET COD_TIPO_ALMACENAMIENTO =" & rs!COD_TIPO_ALMACENAMIENTO
    sql = sql & " WHERE (COD_TIPO_ALMACENAMIENTO IS NULL) AND NRO_REMITO = " & rs!Cod_REMITO
    ExecutarSql sql
    rs.MoveNext
Loop



End Sub

Private Sub ctlCliente_Click()
If ctlRemito_Operacion.Valor = 0 And Not IsNull(ctlCliente.Valor) Then
   ctlClienteFirma.Clear
   ctlClienteFirma.LlenarConCliente CInt(ctlCliente.Valor)
   ElementosEstado = ""
End If
ElementosEstado = ""
CorreoEnviado = False
End Sub


Private Sub ctlCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtElemento.SetFocus
End If

End Sub

Private Sub ctlClienteFirma_LostFocus()
     lblsector.Caption = ctlClienteFirma.Sector
End Sub

Private Sub ctlTipo_elemento_Click()
 TituloGrilla (ctlTipo_Elemento.DESCRIPCION)
End Sub

Private Sub Form_Load()
On Error GoTo salir

   RequerimientoRemito = 0
   TituloGrilla "CAJAS"
   ctlCliente.TipoControl = Cliente
   ctlResponsable.TipoControl = PERSONAL
   ctlRemito_Tipo.TipoControl = TIPO_REMITO
   ctlTipo_Elemento.TipoControl = Tipo_Remito_almacenamiento
   ctlRemito_Operacion.TipoControl = Tipo_Remito_Operacion
   ctlRemtito_Estado.TipoControl = Tipo_Remito_Estados
   mskFechaRemito.Text = Format(SysDate_DD_MM_YYYY, "dd/mm/yyyy")
If MDIfrmInicio.StaInicio.Panels(2).Text = 17 Or MDIfrmInicio.StaInicio.Panels(2).Text = 47 Or MDIfrmInicio.StaInicio.Panels(2).Text = 46 Or MDIfrmInicio.StaInicio.Panels(2).Text = 48 Then
    ChkNocontrolEntrada.Enabled = True
End If

    

    Exit Sub
salir:
    MsgBox Err.Description
    
End Sub
Function ProximoRemito() As Long
  Dim sql As String
  Dim OraMax As New ADODB.Recordset
  sql = "Select Max(Nro_Remito) Maximo From Remitos_Cuerpo"
  OraMax.Open sql, ConActiva, 0, 1
  If IsNull(OraMax!Maximo) Then
    ProximoRemito = 1:
    Exit Function
  Else
    ProximoRemito = Val(OraMax("Maximo")) + 1
  End If
End Function
Sub GrabarMovHistorico(NRO_REMITO As Long, ID_CLIENTE As Integer, Elemento As Long, TIPO, _
    OPERACION As Long, FECHA_MOVIMIENTO As String, TIPO_ELEMENTO As Integer, conRemito As ADODB.Connection)
    Dim sql As String
    sql = " INSERT INTO MOV_CAJAS2"
    sql = sql & vbCrLf & "   (NRO_REMITO, NRO_CAJA, ID_CLIENTE, ELEMENTO, TIPO,"
    sql = sql & vbCrLf & "   OPERACION, FECHA_MOVIMIENTO, TIPO_ELEMENTO ,"
    sql = sql & vbCrLf & "   AUDIT_USUARIO, AUDIT_FECHA)"
    sql = sql & vbCrLf & " VALUES ("
    sql = sql & vbCrLf & NRO_REMITO & "," & Elemento & "," & ID_CLIENTE & "," & Elemento & "," & TIPO & ","
    sql = sql & vbCrLf & OPERACION & "," & FechaFormato(FECHA_MOVIMIENTO) & "," & TIPO_ELEMENTO
    sql = sql & vbCrLf & ",'Basa'," & FechaFormato(FECHA_MOVIMIENTO) & ")"
   conRemito.Execute sql
    
   
End Sub

Public Sub TituloGrilla(Titulo)
    
    grdGuardiayCustodia.ColWidth(0) = 100
    grdGuardiayCustodia.ColWidth(1) = (grdGuardiayCustodia.Width - 210) / 5
    grdGuardiayCustodia.ColWidth(2) = (grdGuardiayCustodia.Width - 210) / 5
    grdGuardiayCustodia.ColWidth(3) = (grdGuardiayCustodia.Width - 210) / 5
    grdGuardiayCustodia.ColWidth(4) = (grdGuardiayCustodia.Width - 210) / 5
    grdGuardiayCustodia.ColWidth(5) = (grdGuardiayCustodia.Width - 210) / 5
    
    grdGuardiayCustodia.ColAlignment(1) = 4
    grdGuardiayCustodia.ColAlignment(2) = 4
    grdGuardiayCustodia.ColAlignment(3) = 4
    grdGuardiayCustodia.ColAlignment(4) = 4
    grdGuardiayCustodia.ColAlignment(5) = 4
        
    grdGuardiayCustodia.TextMatrix(0, 1) = Titulo
    grdGuardiayCustodia.TextMatrix(0, 2) = Titulo
    grdGuardiayCustodia.TextMatrix(0, 3) = Titulo
    grdGuardiayCustodia.TextMatrix(0, 4) = Titulo
    grdGuardiayCustodia.TextMatrix(0, 5) = Titulo
    
    grdDevolucionConsulta.ColWidth(0) = 100
    grdDevolucionConsulta.ColWidth(1) = (grdDevolucionConsulta.Width - 210) / 5
    grdDevolucionConsulta.ColWidth(2) = (grdDevolucionConsulta.Width - 210) / 5
    grdDevolucionConsulta.ColWidth(3) = (grdDevolucionConsulta.Width - 210) / 5
    grdDevolucionConsulta.ColWidth(4) = (grdDevolucionConsulta.Width - 210) / 5
    grdDevolucionConsulta.ColWidth(5) = (grdDevolucionConsulta.Width - 210) / 5
    
    grdDevolucionConsulta.ColAlignment(1) = 4
    grdDevolucionConsulta.ColAlignment(2) = 4
    grdDevolucionConsulta.ColAlignment(3) = 4
    grdDevolucionConsulta.ColAlignment(4) = 4
    grdDevolucionConsulta.ColAlignment(5) = 4
        
    grdDevolucionConsulta.TextMatrix(0, 1) = Titulo
    grdDevolucionConsulta.TextMatrix(0, 2) = Titulo
    grdDevolucionConsulta.TextMatrix(0, 3) = Titulo
    grdDevolucionConsulta.TextMatrix(0, 4) = Titulo
    grdDevolucionConsulta.TextMatrix(0, 5) = Titulo

End Sub
Public Function CargarGrilla(Valor As String) As Boolean
    Dim c As Integer
    Dim R As Integer
    Dim RsEstado As ADODB.Recordset
    Dim RsEntrada As ADODB.Recordset
    Dim sql As String
    Dim ESTADO As Integer
    Dim PERSONAL  As Integer
    Dim RsBuscarCliente As ADODB.Recordset
  
   
       CargarGrilla = True
     On Error GoTo salir:
Control:
     

       If IsNull(ctlRemito_Tipo.Valor) Or IsNull(ctlTipo_Elemento.Valor) Then
            MsgBox "Ingrese el Tipo Remito /Elemnto", vbCritical
            Exit Function
       End If

        If ctlRemito_Operacion.Valor = 0 Then
            If ChkNocontrolEntrada.Value = 0 Then
 
 If Valor > 100000 Then
  sql = "SELECT * From ENTRADA WHERE  (ELEMENTO = " & Valor & ") AND (TIPO =" & ctlTipo_Elemento.Valor & " ) AND (COD_ESTADO < 1)"
  Else
   sql = "SELECT * From ENTRADA WHERE (COD_CLIENTE = " & ctlCliente.Valor & " ) AND (ELEMENTO = " & Valor & ") AND (TIPO =" & ctlTipo_Elemento.Valor & " ) AND (COD_ESTADO < 1)"
 End If
            
            Set RsEntrada = New ADODB.Recordset
            RsEntrada.Open sql, ConActiva, 0, 1
                If RsEntrada.EOF Then
                Clipboard.Clear
                Clipboard.SetText Valor
                
                      If MsgBox("El elemento " & Valor & vbCrLf & " NO fue tomado por el sistema de entrada " & vbCrLf & "Usted quiere continuar?", vbYesNo) = vbNo Then
                            AbortarEntrada = True
                            CargarGrilla = False
                            Exit Function
                      End If
                      
                     
                Exit Function
            End If
End If

        End If
        
        Dim Msg As String
        
        If chkEstado.Value = 0 Then
Dim rsCaja As New ADODB.Recordset
        Select Case ctlTipo_Elemento.Valor
        Case 0 ' Cajas
         If Valor > 100000 Then
            sql = "Select * from Contenedor where  (NOT (COD_CLIENTE IS NULL))" & " And NRO_CAJA = " & Valor
         Else
            sql = "Select * from Contenedor where cod_cliente= " & CInt(ctlCliente.Valor) & " and nro_caja = " & Valor
         End If
         
        Case 1 ' Libros
            sql = "SELECT ESTADO  , COD_CLIENTE From LIBROS WHERE COD_CLIENTE = " & CInt(ctlCliente.Valor) & " AND NRO_LIBRO_INTERNO = " & Valor
        Case 3 'Legajos
            sql = "SELECT cod_ESTADO AS ESTADO , COD_CLIENTE  From Legajos WHERE COD_CLIENTE = " & CInt(ctlCliente.Valor) & " AND ID_CLIENTE_LEGAJO = " & Valor
        End Select
           Set RsEstado = New ADODB.Recordset
           RsEstado.Open sql, ConActiva, 0, 1
           If RsEstado.EOF Then
             PERSONAL = MDIfrmInicio.StaInicio.Panels.Item(2).Text
            If ctlTipo_Elemento.Valor = 0 And (PERSONAL = 47 Or PERSONAL = 48 Or PERSONAL = 17) Then
                If MsgBox("EL cliente no tiene asignada la caja " & Valor & "  quiere hacerlo", vbYesNo) = vbYes Then
                   If AsingacionCajas(ctlCliente.Valor, CLng(Valor)) Then
                        GoTo Control:
                    End If
                End If
            Else
                
                

                
                If MsgBox("El cliente no tiene ese elemento se registrara en al memoria " & Valor & vbCrLf & "Usted quiere continuar?", vbYesNo) = vbYes Then
                 ElementosEstado = ElementosEstado & vbCrLf & vbTab & Valor & "no tiene el elemento"
                 Clipboard.Clear
                 Clipboard.SetText ElementosEstado
               Else
                Clipboard.Clear
                 Clipboard.SetText ElementosEstado
                 CargarGrilla = False
                 Exit Function
              End If
             End If
              
            Exit Function
            Else
            
            If RsEstado!Cod_cliente <> CInt(ctlCliente.Valor) Then
                    MsgBox "La caja " & Valor & " esta asignada a otro cliente : " & RsEstado!Cod_cliente
                    Exit Function
                End If
                
            
           End If
           
            If Not IsNull(RsEstado!ESTADO) Then
                    ESTADO = CInt(RsEstado!ESTADO)
                Else
                    MsgBox "El elemento es nulo " & Valor
                    Exit Function
                End If
            Else
            
                 ESTADO = 3
            End If
            Select Case ctlRemito_Tipo.Valor
            Case 0, 1 ' Guarda y custodia consulta
                If ctlRemito_Operacion.Valor = 1 Then
                    ' Salida
                    If ESTADO = 2 Then
                        IngresarValorGrilla Valor, grdDevolucionConsulta
                    Else
                        If MsgBox("El estado no es valido " & Valor & vbCrLf & "Usted quiere continuar?", vbYesNo) = vbNo Then
                        CargarGrilla = False
                        Else
                            ElementosEstado = ElementosEstado & vbCrLf & vbTab & Valor & "no tiene el Estado correcto"
                            Clipboard.Clear
                            Clipboard.SetText ElementosEstado
                        End If
                        
                        Exit Function
                    End If
                 Else
                    'Entrada
                    If ESTADO = 2 Then
                    Error = Error & vbCrLf & "Estado en planta" & vbTab & Valor
                        If MsgBox("El estado no es valido EN PLANTA " & Valor & vbCrLf & "Usted quiere continuar?", vbYesNo) = vbNo Then
                                CargarGrilla = False
                                Exit Function
                        End If
                    End If
                    If ESTADO = 3 Then
                        IngresarValorGrilla Valor, grdDevolucionConsulta
                    End If
                    If ESTADO = 4 Then
                        Error = Error & vbCrLf & "Estado en reserva" & vbTab & Valor
                        If MsgBox("El estado no es valido " & Valor & vbCrLf & "Usted quiere continuar?", vbYesNo) = vbNo Then
                        CargarGrilla = False
                        End If
                        
                        Exit Function
                    End If
                    
                    If ESTADO = 5 Then
                        IngresarValorGrilla Valor, grdGuardiayCustodia
                    End If
                    If ESTADO = 7 Then
                        MsgBox "La posición de la caja NO es Valida" & vbCrLf & "Debido a que permaneció mucho tiempo fuera de la planta" & vbCrLf & "Por favor comuníquese con el Jefe de planta", vbInformation
                        UpdateContedor CLng(Valor), ctlCliente.Valor, 7, 3, 0, ConActiva
                        IngresarValorGrilla Valor, grdDevolucionConsulta
                     Exit Function
                    End If
                 
                 End If
         Case 2 ' Vacias
             If ESTADO = 4 Then
                 IngresarValorGrilla Valor, grdDevolucionConsulta
              Else
                Error = Error & vbCrLf & "Estado en reserva " & vbTab & Valor
                    MsgBox "El estado no es correcto " & Valor
                    
                    SrtValor = SrtValor & vbCrLf & "  " & Valor
                    Exit Function
              End If
         Case 3 ' Bajas
                  IngresarValorGrilla Valor, grdDevolucionConsulta
         
        Case 4 'Devolucion de vacias
              If ESTADO = 5 Then
                 IngresarValorGrilla Valor, grdDevolucionConsulta
              Else
                MsgBox "El estado no es correcto"
                Exit Function
              End If
        End Select
        
        
       
        
        
        Exit Function
salir:        MsgBox Err.Description

        CargarGrilla = False
                       
End Function

Public Function Validar_Guardia_Custadia() As Boolean
    Dim rs As New ADODB.Recordset
    Dim R As Integer
    Dim c As Integer
    Dim Filtro As String
    Dim sql As String
    Dim Bandera As Boolean
    Dim i As Integer
        Bandera = False
        Validar_Guardia_Custadia = True
         For R = 1 To grdGuardiayCustodia.Rows - 1
             For c = 1 To grdGuardiayCustodia.Cols - 1
                 If grdGuardiayCustodia.TextMatrix(R, c) <> "" Then
                     Filtro = Filtro & grdGuardiayCustodia.TextMatrix(R, c) & ","
                 End If
             Next
         Next
        Filtro = Mid(Filtro, 1, Len(Filtro) - 1)
        sql = " SELECT ESTADO FROM CONTENEDOR WHERE "
        sql = sql & "COD_CLIENTE = " & CInt(ctlCliente.Valor) 'CAJA
        sql = sql & " AND NRO_CAJA IN  (" & Filtro & ")"
         rs.Open sql, ConActiva, 0, 1
        Do While Not rs.EOF
            If CInt(rs!ESTADO) <> 5 And CInt(rs!ESTADO) <> 4 Then
                MsgBox "La Caja No tiene el estado Correcto"
                Validar_Guardia_Custadia = False
            End If
            rs.MoveNext
        Loop
End Function

Public Function Validar_Consulta() As Boolean
    Dim rs As New ADODB.Recordset
    Dim R As Integer
    Dim c As Integer
    Dim Filtro As String
    Dim sql As String
    Dim Bandera As Boolean
    Dim i As Integer
        Bandera = False
        Validar_Consulta = True
         For R = 1 To grdDevolucionConsulta.Rows - 1
             For c = 1 To grdDevolucionConsulta.Cols - 1
                 If grdDevolucionConsulta.TextMatrix(R, c) <> "" Then
                     Filtro = Filtro & grdDevolucionConsulta.TextMatrix(R, c) & ","
                 End If
             Next
         Next
        Filtro = Mid(Filtro, 1, Len(Filtro) - 1)
        sql = " SELECT ESTADO FROM CONTENEDOR WHERE "
        sql = sql & "COD_CLIENTE = " & CInt(ctlCliente.Valor) 'CAJA
        sql = sql & " AND NRO_CAJA IN  (" & Filtro & ")"
         rs.Open sql, ConActiva, 0, 1
        Do While Not rs.EOF
            If CInt(rs!ESTADO) <> 3 Then
                MsgBox "La Caja No tiene el estado Correcto"
                Validar_Consulta = False
            End If
            rs.MoveNext
        Loop
End Function


Private Sub mskRemitoConsulta_LostFocus()
mskRemitoCustodia.Text = mskRemitoConsulta.Text
End Sub


Private Sub txtElemento_KeyPress(KeyAscii As Integer)
On Error GoTo salir:
        If KeyAscii = 13 Then
            If UCase(Mid(txtElemento.Text, 1, 2)) = "L1" Then
                If ctlCliente.Valor = CInt(Mid(txtElemento, 3, 3)) Then
                        CargarGrilla CLng(Mid(txtElemento, 6, 6))
                 End If
            Else
                If IsNumeric(txtElemento.Text) Then
                    CargarGrilla CLng(txtElemento.Text)
                Else
                    MsgBox "Error en el tipo de elemento", vbInformation
                End If
            End If
            txtElemento.Text = ""
        End If
Exit Sub
CorreoEnviado = False
salir:
MsgBox "Error en numero"
End Sub

Public Function ContarGrilla(Grilla As MSFlexGrid) As Integer
    Dim i As Integer
    Dim R As Integer
    Dim c As Integer
    With Grilla
         For R = 1 To .Rows - 1
             For c = 1 To .Cols - 1
                 If .TextMatrix(R, c) <> "" Then
                     i = i + 1
                 End If
             Next
         Next
     End With
     If Grilla.Name = "grdDevolucionConsulta" Then
        lblCantidadDevolucion = i
     Else
        lblCantidadGuardia = i
     End If
End Function

Public Sub Limpiar()
    grdDevolucionConsulta.Clear
    grdDevolucionConsulta.Rows = 2
    grdGuardiayCustodia.Clear
    grdGuardiayCustodia.Rows = 2
    TituloGrilla "CAJAS"
    lblsector.Caption = ""
    ctlCliente.Valor = ""
    lblCantidadDevolucion = ""
    lblCantidadGuardia = ""
    ctlCliente.Valor = Null
    ctlClienteFirma.Valor = Null
    Rem ctlRemito_Tipo.Valor = Null
    Rem ctlResponsable.Valor = Null
   Rem  ctlTipo_elemento.Valor = Null
    mskRemitoConsulta.Mask = ""
    mskRemitoConsulta.Text = ""
    mskRemitoConsulta.Mask = "0001-000#####"
    mskRemitoCustodia.Mask = ""
    mskRemitoCustodia.Text = ""
    mskRemitoCustodia.Mask = "0001-000#####"
        
        
End Sub

Public Function InsertarRemitoCuerpo(TIPO As Integer, NRO_REM_PROV As String, OBSERVACIONES As String, CANTIDAD As String, COBRAR_FLETE As String, conexion As ADODB.Connection, HORASARCHIVISTA As Integer) As Long
    Dim sql As String
    
    Dim OPERACION As String
    Dim ESTADO As String
    Dim Fecha As String
    Dim ID_CLIENTE As String
    Dim AUDIT_USUARIO As String
    Dim AUDIT_FECHA As String
    Dim COD_TIPO_ALMACENAMIENTO As String
    Dim COD_PERSONAL_ENTREGA As String
    Dim COD_USUARIO_CLIENTE As Integer
    Dim NRO_REMITO As Long
    
            
            
            OPERACION = ctlRemito_Operacion.Valor
            ESTADO = ctlRemtito_Estado.Valor
            Fecha = FechaFormato(mskFechaRemito.Text)
            ID_CLIENTE = ctlCliente.Valor
            AUDIT_USUARIO = MDIfrmInicio.StaInicio.Panels(2).Text
            AUDIT_FECHA = SysDate
            COD_TIPO_ALMACENAMIENTO = ctlTipo_Elemento.Valor
            COD_PERSONAL_ENTREGA = ctlResponsable.Valor
            COD_USUARIO_CLIENTE = ctlClienteFirma.Valor
            NRO_REMITO = ProximoRemito
    
            sql = " INSERT INTO REMITOS_CUERPO"
            sql = sql & vbCrLf & " (NRO_REMITO, NRO_REM_PROV, TIPO, OPERACION, ESTADO,FECHA, ID_CLIENTE, OBSERVACIONES, CANTIDAD,"
            sql = sql & vbCrLf & " AUDIT_USUARIO, AUDIT_FECHA,COD_TIPO_ALMACENAMIENTO, COD_PERSONAL_ENTREGA,COD_USUARIO_CLIENTE , COBRAR_FLETE , HORASARCHIVISTA)"
            sql = sql & vbCrLf & " VALUES (" & NRO_REMITO & ",'" & NRO_REM_PROV & "'," & TIPO & "," & OPERACION & "," & ESTADO & ","
            sql = sql & vbCrLf & Fecha & "," & ID_CLIENTE & ",'" & UCase(Trim(OBSERVACIONES)) & "'," & CANTIDAD
            sql = sql & vbCrLf & ",'" & AUDIT_USUARIO & "'," & AUDIT_FECHA & "," & COD_TIPO_ALMACENAMIENTO & "," & COD_PERSONAL_ENTREGA & "," & COD_USUARIO_CLIENTE & " ,'" & COBRAR_FLETE & "'," & HORASARCHIVISTA & ")"
            conexion.Execute sql
            InsertarRemitoCuerpo = NRO_REMITO
End Function

Public Sub InsertarRemitoDetalle(NRO_REMITO As Long, AUDIT_USUARIO As String, AUDIT_FECHA As String, DESDE As Long, HASTA As String, TIPO_ALMACENADO As Integer, Conecion As ADODB.Connection)
    Dim sql As String
    sql = "INSERT INTO REMITOS_DETALLE"
    sql = sql & vbCrLf & " (NRO_REMITO, AUDIT_USUARIO, AUDIT_FECHA,DESDE, HASTA, TIPO_ALMACENADO)"
    sql = sql & vbCrLf & "  VALUES (" & NRO_REMITO & ",'" & AUDIT_USUARIO & "'," & AUDIT_FECHA & "," & DESDE & "," & HASTA & "," & TIPO_ALMACENADO & ")"
    ExecutarSql sql
    If TIPO_ALMACENADO = 0 Then
        sql = " UPDATE    dbo.CAJAS SET "
        sql = sql & " FK_REMITO_VACIAS =" & NRO_REMITO
        sql = sql & " , FK_ESTADO =1110"
        sql = sql & " Where ID_CAJA = " & DESDE
        Conecion.Execute sql
    End If
End Sub

Public Sub InsertarMovimientosElementos(Cod_REMITO As Long, TIPO_REMITO As Integer, OPERACION As Integer, TIPO_ALMACENAMIENTO As String, Elemento As Long, Cod_cliente As Integer, Fecha As String, conRemito As ADODB.Connection)
    Dim sql As String
    sql = " INSERT INTO MOVIMIENTOS_ELEMENTOS"
    sql = sql & vbCrLf & "  (COD_REMITO, COD_TIPO, COD_OPERACION,COD_TIPO_ALMACENAMIENTO, ELEMENTO, COD_CLIENTE,FECHA)"
    sql = sql & vbCrLf & "  VALUES (" & Cod_REMITO & "," & TIPO_REMITO & "," & OPERACION & "," & TIPO_ALMACENAMIENTO & "," & Elemento & "," & Cod_cliente & "," & FechaFormato(Fecha) & ")"
    conRemito.Execute sql
End Sub

Public Function UpdateContedor(NRO_CAJA As Long, Cod_cliente As Integer, ESTADO_ACTUAL As Integer, ESTADO_FINAL As Integer, Remito As Long, conRemito As ADODB.Connection) As Boolean
    Dim sql As String
    Dim R As Integer
    UpdateContedor = True
    sql = "UPDATE CONTENEDOR SET "
    sql = sql & vbCrLf & " ESTADO = " & ESTADO_FINAL
    sql = sql & ", NRO_REMITO = " & Remito
    sql = sql & ", F_MODIFICACION = " & SysDate
    sql = sql & vbCrLf & " WHERE "
    sql = sql & " COD_CLIENTE = " & Cod_cliente
    sql = sql & " AND NRO_CAJA = " & NRO_CAJA
    sql = sql & " AND ESTADO = " & ESTADO_ACTUAL
     conRemito.Execute sql, R
    
    If ESTADO_FINAL = 2 Then
        If ESTADO_ACTUAL = 5 Then
            sql = " Update Cajas "
            sql = sql & vbCrLf & " SET FK_ESTADO =1120"
            sql = sql & vbCrLf & " , FECHA_MODIFICACION =" & SysDate
            sql = sql & vbCrLf & " , FK_REMITO_CUSTODIA =" & Remito
'             sql = sql & vbCrLf & " , TIPO_REFERENCIA = '" & cboReferencia.Text & "'"
             sql = sql & vbCrLf & " Where FK_CLIENTE = " & Cod_cliente
             sql = sql & vbCrLf & " And NRO_CAJA = " & NRO_CAJA
            conRemito.Execute sql
        Else
            sql = " Update Cajas "
            sql = sql & vbCrLf & " SET FK_ESTADO =1120"
            sql = sql & vbCrLf & " , FECHA_MODIFICACION =" & SysDate
            sql = sql & vbCrLf & " , FK_REMITO_CONSULTA =" & Remito
            sql = sql & vbCrLf & " Where FK_CLIENTE = " & Cod_cliente
            sql = sql & vbCrLf & " And NRO_CAJA = " & NRO_CAJA
            conRemito.Execute sql
        End If
    End If
    
    
    If ESTADO_FINAL = 3 Then
    sql = " Update Cajas "
    sql = sql & vbCrLf & " SET FK_ESTADO =1130"
    sql = sql & vbCrLf & " , FECHA_MODIFICACION =" & SysDate
    sql = sql & vbCrLf & " , FK_REMITO_CONSULTA =" & Remito
    sql = sql & vbCrLf & " Where FK_CLIENTE = " & Cod_cliente
    sql = sql & vbCrLf & " And NRO_CAJA = " & NRO_CAJA
    conRemito.Execute sql
    End If
    
    
    
    If R <> 1 Then
        UpdateContedor = False
        MsgBox "El elemento " & NRO_CAJA & " No fue Actualizado"
    End If
End Function
Public Function BajasContedor(NRO_CAJA As Long, Cod_cliente As Integer, Remito As Long, conRemito As ADODB.Connection) As Boolean
    Dim sql As String
    Dim Afec As Integer
    Dim rs As New ADODB.Recordset
    
    BajasContedor = True
    ' el numero de remito debe ser con el que seva actualizar
    
    sql = " INSERT INTO CONTENEDOR_BAJAS"
    sql = sql & vbCrLf & " (ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
    sql = sql & vbCrLf & " ESTADO, COD_CLIENTE, NRO_CAJA,NRO_REMITO,"
    sql = sql & vbCrLf & " F_MODIFICACION, IDREQUERIMIENTO)"
    sql = sql & vbCrLf & " SELECT ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
    sql = sql & vbCrLf & " ESTADO, COD_CLIENTE, NRO_CAJA," & Remito & ","
    sql = sql & vbCrLf & " F_MODIFICACION , IDREQUERIMIENTO"
    sql = sql & vbCrLf & " From CONTENEDOR"
    sql = sql & vbCrLf & " Where COD_CLIENTE = " & Cod_cliente
    sql = sql & vbCrLf & " And NRO_CAJA = " & NRO_CAJA
    conRemito.Execute sql, Afec
    If Afec <> 1 Then
        BajasContedor = False
    End If
    
    
    sql = " SELECT     ID_CONTENEDOR, NRO_CAJA, COD_CLIENTE"
 sql = sql & vbCrLf & " From CONTENEDOR"
 sql = sql & vbCrLf & " Where Cod_cliente = " & Cod_cliente
  sql = sql & vbCrLf & " And NRO_CAJA = " & NRO_CAJA
    
    rs.Open sql, strConBasa, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Then
    
    
    sql = "UPDATE CONTENEDOR SET "
    sql = sql & vbCrLf & " ESTADO = 1 "
    sql = sql & ", NRO_REMITO = " & Remito
    sql = sql & ", F_MODIFICACION = " & SysDate
    sql = sql & ", COD_CLIENTE = NULL "
    sql = sql & ",  NRO_CAJA = " & rs!ID_CONTENEDOR
    sql = sql & vbCrLf & " WHERE "
    sql = sql & " ID_CONTENEDOR = " & rs!ID_CONTENEDOR
    conRemito.Execute sql
    
    Else
        MsgBox "Error "
    
    End If
        sql = " Update Cajas "
        sql = sql & vbCrLf & " SET FK_ESTADO =1140"
        sql = sql & vbCrLf & " , FK_REMITO_BAJA =" & Remito
        sql = sql & vbCrLf & ", FECHA_MODIFICACION =" & SysDate
        sql = sql & vbCrLf & "  Where FK_CLIENTE = " & Cod_cliente
        sql = sql & vbCrLf & "  And NRO_CAJA = " & NRO_CAJA
     conRemito.Execute sql
    
    
    If Afec <> 1 Then
        BajasContedor = False
    Else
    BajasContedor = True
    End If
End Function
Public Sub BajasContedorVacias(NRO_CAJA As Long, Cod_cliente As Integer, Remito As Long)
    Dim sql As String
    
        sql = " INSERT INTO CONTENEDOR_BAJAS"
        sql = sql & vbCrLf & " (ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
        sql = sql & vbCrLf & " ESTADO, COD_CLIENTE, NRO_CAJA, NRO_REMITO,"
        sql = sql & vbCrLf & " F_MODIFICACION, IDREQUERIMIENTO)"
        sql = sql & vbCrLf & " SELECT ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
        sql = sql & vbCrLf & " ESTADO, COD_CLIENTE, NRO_CAJA, NRO_REMITO,"
        sql = sql & vbCrLf & " F_MODIFICACION , IDREQUERIMIENTO"
        sql = sql & vbCrLf & " From CONTENEDOR"
        sql = sql & vbCrLf & " Where COD_CLIENTE = " & Cod_cliente
        sql = sql & vbCrLf & " And NRO_CAJA = " & NRO_CAJA
        
            ExecutarSql sql
    
    sql = "UPDATE CONTENEDOR SET "
    sql = sql & vbCrLf & " ESTADO = 1 "
    sql = sql & ", NRO_REMITO = " & Remito
    sql = sql & ", F_MODIFICACION = " & SysDate
    sql = sql & ", COD_CLIENTE = NULL "
    sql = sql & ", NRO_CAJA = NULL"
    sql = sql & vbCrLf & " WHERE "
    sql = sql & " COD_CLIENTE = " & Cod_cliente
    sql = sql & " AND NRO_CAJA = " & NRO_CAJA
    sql = sql & " AND ESTADO = 5 "
    
    ExecutarSql sql
End Sub
Public Sub UpdateLegajos(ID_CLIENTE_LEGAJO As Long, Cod_cliente As Integer, COD_ESTADO_ACTUAL As Integer, COD_ESTADO_FINAL As Integer, Cod_REMITO As Long, Fecha As String, conRemito As ADODB.Connection)
    Dim sql As String
    sql = " Update LEGAJOS SET  COD_ESTADO =" & COD_ESTADO_FINAL & ",COD_REMITO =" & Cod_REMITO & ", FECHA =" & Fecha
    sql = sql & vbCrLf & " WHERE ID_CLIENTE_LEGAJO = " & ID_CLIENTE_LEGAJO & " AND COD_CLIENTE = " & Cod_cliente
    Rem & " " AND COD_ESTADO = " & COD_ESTADO_ACTUAL
    conRemito.Execute sql
  
End Sub

Public Sub UpdateLibros(NRO_LIBRO_INTERNO As Long, Cod_cliente As Integer, ESTADO_ACTUAL As Integer, ESTADO_FINAL As Integer, NRO_REMITO As Long, conRemito As ADODB.Connection)
    Dim sql As String
    sql = " Update LIBROS SET  ESTADO = " & ESTADO_FINAL & ", NRO_REMITO =" & NRO_REMITO
    sql = sql & vbCrLf & " Where NRO_LIBRO_INTERNO = " & NRO_LIBRO_INTERNO & " And COD_CLIENTE =" & Cod_cliente & " And ESTADO = " & ESTADO_ACTUAL
    conRemito.Execute sql
End Sub

Public Sub GrabarRemitoCustodia()
    Dim Remito As Long
    Dim R As Integer
    Dim c As Integer
    Dim conRemito As New ADODB.Connection
    
    On Error GoTo salir:
   Rem  conRemito.Mode = adModeShareDenyNone
    
    conRemito.Open strConBasa
    conRemito.BeginTrans
    
    
    If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
           ' strConBasa , 0 ,1.BeginTrans
            Remito = InsertarRemitoCuerpo(0, mskRemitoCustodia.Text, Trim(UCase(txtDescripcionCustodia.Text)), lblCantidadGuardia.Caption, chkFLeteGyC.Value, conRemito, txtHorasArchivista.Text)
            'INSERTAR EN REMITO DELTALLE
            For R = 1 To grdGuardiayCustodia.Rows - 1
               For c = 1 To grdGuardiayCustodia.Cols - 1
                    If grdGuardiayCustodia.TextMatrix(R, c) <> "" Then
                    
                    If AbortarEntrada = True Then
                      GoTo salir
                    End If
                    
                        If RAPC = True Then
                            ReferenciasApc ctlCliente.Valor, grdGuardiayCustodia.TextMatrix(R, c), ctlClienteFirma.Indice, " Referencia administrada por el cliente  remito : " & mskRemitoCustodia.Text
                        End If
                        
                        InsertarRemitoDetalle Remito, UserName, SysDate, grdGuardiayCustodia.TextMatrix(R, c), grdGuardiayCustodia.TextMatrix(R, c), ctlTipo_Elemento.Valor, conRemito
                        GrabarMovHistorico Remito, ctlCliente.Valor, grdGuardiayCustodia.TextMatrix(R, c), 0, ctlRemito_Operacion.Valor, mskFechaRemito.Text, ctlTipo_Elemento.Valor, conRemito
                        InsertarMovimientosElementos Remito, 0, ctlRemito_Operacion.Valor, ctlTipo_Elemento.Valor, grdGuardiayCustodia.TextMatrix(R, c), ctlCliente.Valor, mskFechaRemito.Text, conRemito
                        UpdateEntrada ctlCliente.Valor, grdGuardiayCustodia.TextMatrix(R, c), ctlTipo_Elemento.Valor, conRemito
                        
                        
                        'MOVIMIENTO EN TABLA CONTENIDO
                        
                        Select Case ctlTipo_Elemento.Valor
                        Case 0
                            UpdateContedor grdGuardiayCustodia.TextMatrix(R, c), ctlCliente.Valor, 5, 2, Remito, conRemito
                           Rem  InsertarSectorCaja ctlCliente.Valor, grdGuardiayCustodia.TextMatrix(R, c), ctlClienteFirma.Indice, FechaFormato(mskFechaRemito.Text), ctlClienteFirma.Valor
                           Rem InsertarTipoReferencia ctlCliente.Valor, grdGuardiayCustodia.TextMatrix(R, c), Trim(cboReferencia.Text)
                        Case 1
                            UpdateLibros grdGuardiayCustodia.TextMatrix(R, c), ctlCliente.Valor, 5, 2, Remito, conRemito
                        Case 2
                            UpdateLegajos grdGuardiayCustodia.TextMatrix(R, c), ctlCliente.Valor, 5, 2, Remito, SysDate, conRemito
                        End Select
                    End If
                 Next
             Next
conRemito.CommitTrans
            MsgBox "Nº de movimiento es para " & vbCrLf & " Guarda y Custodia es : " & Remito
            Screen.MousePointer = 0
            
            On Error GoTo ErrorPrn
    End If

Exit Sub
salir:
    Screen.MousePointer = 0
    conRemito.RollbackTrans
    MsgBox "Error en remito", vbCritical
    Exit Sub
    
ErrorPrn:
    MsgBox Error
    Exit Sub
End Sub

Public Sub GrabarRemitoDevolucion()
    Dim Remito As Long
    Dim R As Integer
    Dim c As Integer
    Dim EstadoActual As Integer
    Dim EstadoFinal As Integer
    Dim conRemito As New ADODB.Connection
    Rem conRemito.Mode = adModeShareDenyNone
    conRemito.Open strConBasa
    On Error GoTo salir:
    
    If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
            conRemito.BeginTrans
            Remito = InsertarRemitoCuerpo(1, mskRemitoConsulta.Text, UCase(txtDescripcionConsulta.Text), lblCantidadDevolucion.Caption, chkFleteConsulta.Value, conRemito, txtHorasArchivista.Text)
            'INSERTAR EN REMITO DELTALLE
            If ctlRemito_Operacion.Valor = 0 Then
                EstadoActual = 3
                EstadoFinal = 2
            Else
                EstadoActual = 2
                EstadoFinal = 3
            End If
            
            
           With grdDevolucionConsulta
                For R = 1 To .Rows - 1
                   For c = 1 To .Cols - 1
                   If AbortarEntrada = True Then
                       GoTo salir
                    End If
                        If .TextMatrix(R, c) <> "" Then
                            InsertarRemitoDetalle Remito, UserName, SysDate, grdDevolucionConsulta.TextMatrix(R, c), grdDevolucionConsulta.TextMatrix(R, c), ctlTipo_Elemento.Valor, conRemito
                            GrabarMovHistorico Remito, ctlCliente.Valor, .TextMatrix(R, c), 1, ctlRemito_Operacion.Valor, mskFechaRemito.Text, ctlTipo_Elemento.Valor, conRemito
                            InsertarMovimientosElementos Remito, ctlRemito_Tipo.Valor, ctlRemito_Operacion.Valor, ctlTipo_Elemento.Valor, .TextMatrix(R, c), ctlCliente.Valor, mskFechaRemito.Text, conRemito
                            If ctlRemito_Operacion.Valor = 0 Then
                                UpdateEntrada ctlCliente.Valor, grdDevolucionConsulta.TextMatrix(R, c), ctlTipo_Elemento.Valor, conRemito
                            End If
                            'MOVIMIENTO EN TABLA CONTENIDO
                            Select Case ctlTipo_Elemento.Valor
                            Case 0
                                UpdateContedor grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, EstadoActual, EstadoFinal, Remito, conRemito
                            Case 1
                                UpdateLibros grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, EstadoActual, EstadoFinal, Remito, conRemito
                            Case 3
                                UpdateLegajos grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, EstadoActual, EstadoFinal, Remito, SysDate, conRemito
                               Rem  LlegarDescripcion ctlCliente.Valor, grdDevolucionConsulta.TextMatrix(R, c)
                                
                            End Select
                        End If
                     Next
                 Next
           End With
           
           If ctlRemito_Operacion.Valor = 1 Then
                UpdateRequerimientoRemito RequerimientoRemito, Remito, conRemito
                CRequerimientos.CambioEstado ctlResponsable.Valor, True, 4, 5, conRemito
                conRemito.CommitTrans
                Select Case ctlTipo_Elemento.Valor
                Case 0  'Cajas
                    ImprimirRemitoCaja Remito
                Case 1 'Libros
                    ImprimirRemitoLibros Remito
                Case 3 'Legajos
                    ImprimirRemitoLegajos Remito
                End Select
            Else
            conRemito.CommitTrans
                 'strConBasa.CommitTrans
                MsgBox "El Nº de Remito " & vbCrLf & " para la devolución es : " & Remito, vbInformation
            End If
            Screen.MousePointer = 0
            On Error GoTo ErrorPrn
    End If
Exit Sub
salir:
    Screen.MousePointer = 0
    conRemito.RollbackTrans
    MsgBox "Error En la grabacion"
    Exit Sub
    
ErrorPrn:
    MsgBox Error
    Exit Sub

End Sub
Public Sub GrabarRemitoCajasVacias()
    Dim Remito As Long
    Dim R As Integer
    Dim c As Integer
    Dim EstadoActual As Integer
    Dim EstadoFinal As Integer
    Dim conRemito As New ADODB.Connection
    
    On Error GoTo salir:
     conRemito.Mode = adModeShareDenyNone
    conRemito.Open strConBasa
    conRemito.BeginTrans
    If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
            'strConBasa.BeginTrans
            Remito = InsertarRemitoCuerpo(2, mskRemitoConsulta.Text, UCase(txtDescripcionConsulta.Text), lblCantidadDevolucion.Caption, chkFleteConsulta.Value, conRemito, txtHorasArchivista.Text)
            'INSERTAR EN REMITO DELTALLE
             EstadoActual = 4
             EstadoFinal = 5
            
            
           With grdDevolucionConsulta
                For R = 1 To .Rows - 1
                   For c = 1 To .Cols - 1
                        If .TextMatrix(R, c) <> "" Then
                            InsertarRemitoDetalle Remito, UserName, SysDate, grdDevolucionConsulta.TextMatrix(R, c), grdDevolucionConsulta.TextMatrix(R, c), ctlTipo_Elemento.Valor, conRemito
                            
                            GrabarMovHistorico Remito, ctlCliente.Valor, .TextMatrix(R, c), 2, ctlRemito_Operacion.Valor, mskFechaRemito.Text, ctlTipo_Elemento.Valor, conRemito
                            InsertarMovimientosElementos Remito, ctlRemito_Tipo.Valor, ctlRemito_Operacion.Valor, ctlTipo_Elemento.Valor, .TextMatrix(R, c), ctlCliente.Valor, mskFechaRemito.Text, conRemito
                            If ctlRemito_Operacion.Valor = 0 Then
                                UpdateEntrada ctlCliente.Valor, grdDevolucionConsulta.TextMatrix(R, c), ctlTipo_Elemento.Valor, conRemito
                            End If
                            'MOVIMIENTO EN TABLA CONTENIDO
                             UpdateContedor grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, EstadoActual, EstadoFinal, Remito, conRemito
                        End If
                     Next
                 Next
           End With
           If ctlRemito_Operacion.Valor = 1 Then
                UpdateRequerimientoRemito RequerimientoRemito, Remito, conRemito
                CRequerimientos.CambioEstado ctlResponsable.Valor, True, 4, 5, conRemito
                conRemito.CommitTrans
                ImprimirRemitoCaja Remito
           Else
                conRemito.CommitTrans
                MsgBox "El Nº de Remito " & vbCrLf & " para la devolución es : " & Remito, vbInformation
            End If
            Screen.MousePointer = 0
            On Error GoTo ErrorPrn
    End If
   
Exit Sub
salir:
    Screen.MousePointer = 0
    conRemito.RollbackTrans
    MsgBox "Error En la grabacion"
    Exit Sub
    
ErrorPrn:
    MsgBox Error
    Exit Sub

End Sub

Public Sub GrabarRemitoBajas()
    Dim Remito As Long
    Dim R As Integer
    Dim c As Integer
    Dim conRemito As New ADODB.Connection
    On Error GoTo salir:
   Rem conRemito.Mode = adModeShareDenyNone
    conRemito.Open strConBasa
    If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
            Rem conRemito.BeginTrans
            Remito = InsertarRemitoCuerpo(3, mskRemitoConsulta, txtDescripcionConsulta.Text, lblCantidadDevolucion.Caption, chkFleteConsulta.Value, conRemito, txtHorasArchivista.Text)
            'INSERTAR EN REMITO DELTALLE
            For R = 1 To grdDevolucionConsulta.Rows - 1
               For c = 1 To grdDevolucionConsulta.Cols - 1
                    If grdDevolucionConsulta.TextMatrix(R, c) <> "" Then
                        InsertarRemitoDetalle Remito, UserName, SysDate, grdDevolucionConsulta.TextMatrix(R, c), grdDevolucionConsulta.TextMatrix(R, c), ctlTipo_Elemento.Valor, conRemito
                        GrabarMovHistorico Remito, ctlCliente.Valor, grdDevolucionConsulta.TextMatrix(R, c), 3, ctlRemito_Operacion.Valor, mskFechaRemito.Text, ctlTipo_Elemento.Valor, conRemito
                        InsertarMovimientosElementos Remito, ctlTipo_Elemento.Valor, 3, ctlTipo_Elemento.Valor, grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, mskFechaRemito.Text, conRemito
                        UpdateEntrada ctlCliente.Valor, grdDevolucionConsulta.TextMatrix(R, c), ctlTipo_Elemento.Valor, conRemito
                        
                        'MOVIMIENTO EN TABLA CONTENIDO
                        Select Case ctlTipo_Elemento.Valor
                        Case 0
                             BorraReferencias ctlCliente.Valor, CLng(grdDevolucionConsulta.TextMatrix(R, c)), conRemito
                            InsertarBajas ctlCliente.Valor, grdDevolucionConsulta.TextMatrix(R, c), 0, SysDate, ctlResponsable.Valor, CLng(Remito), conRemito
                           If BajasContedor(grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, CLng(Remito), conRemito) Then
                           Else
                                'strConBasa.RollbackTrans
                                Exit Sub
                                MsgBox "La acutlizacion No se Realizo", vbCritical
                           End If
                        Case 1
                            UpdateLibros grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, 3, 0, Remito, conRemito
                        Case 3
                            BajasLegajos grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, Remito, ctlClienteFirma.Valor, conRemito
                        End Select
                    End If
                 Next
             Next
            'strConBasa.CommitTrans
            MsgBox "  El numero de Devolución es " & vbCrLf & "    " & Remito
            Screen.MousePointer = 0
            On Error GoTo ErrorPrn
    End If
    
   Rem conRemito.CommitTrans
Exit Sub
salir:
    Screen.MousePointer = 0
   Rem conRemito.RollbackTrans
    MsgBox "Error remito"
    Exit Sub
    
ErrorPrn:
    MsgBox Error
    Exit Sub
End Sub
Public Sub GrabarRemitoDevolucionVacias()
'    Dim Remito As Long
'    Dim R As Integer
'    Dim c As Integer
'
'    If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
'            'strConBasa.BeginTrans
'            Remito = InsertarRemitoCuerpo(4, mskRemitoConsulta, txtDescripcionConsulta.Text, lblCantidadDevolucion.Caption, txtFleteConsulta.Text)
'
'
'            'INSERTAR EN REMITO DELTALLE
'            For R = 1 To grdDevolucionConsulta.Rows - 1
'               For c = 1 To grdDevolucionConsulta.Cols - 1
'                    If grdDevolucionConsulta.TextMatrix(R, c) <> "" Then
'                        InsertarRemitoDetalle Remito, UserName, SysDate, grdDevolucionConsulta.TextMatrix(R, c), grdDevolucionConsulta.TextMatrix(R, c), ctlTipo_Elemento.Valor
'                        GrabarMovHistorico Remito, ctlCliente.Valor, CLng(grdDevolucionConsulta.TextMatrix(R, c)), 4, ctlRemito_Operacion.Valor, FechaFormato(mskFechaRemito.Text), ctlTipo_Elemento.Valor
'                        InsertarMovimientosElementos Remito, ctlRemito_Tipo.Valor, 0, ctlTipo_Elemento.Valor, grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, mskFechaRemito.Text
'                        UpdateEntrada ctlCliente.Valor, grdDevolucionConsulta.TextMatrix(R, c), ctlTipo_Elemento.Valor
'                        'MOVIMIENTO EN TABLA CONTENIDO
'                        Select Case ctlTipo_Elemento.Valor
'                        Case 0
'                           BajasContedorVacias grdDevolucionConsulta.TextMatrix(R, c), ctlCliente.Valor, Remito
'                        End Select
'                    End If
'                 Next
'             Next
'            'strConBasa.CommitTrans
'            MsgBox "  El numero de Devolución es " & vbCrLf & "    " & Remito
'            Screen.MousePointer = 0
'            On Error GoTo ErrorPrn
'    End If
'Exit Sub
'OraError:
'    Screen.MousePointer = 0
'    'strConBasa.RollbackTrans
'
'    Exit Sub
'
'ErrorPrn:
'    MsgBox Error
'    Exit Sub
End Sub



Public Sub UpdateEntrada(Cod_cliente As Integer, Elemento As Long, TIPO As Integer, conRemito As ADODB.Connection)
       Dim sql As String
       
       If Elemento > 100000 Then
       
       sql = " Update ENTRADA Set COD_ESTADO = 1"
        sql = sql & vbCrLf & "  WHERE  (ELEMENTO =" & Elemento & ") AND  (TIPO =" & TIPO & ") AND (COD_ESTADO = 0)"
        conRemito.Execute sql
       
       Else
        sql = " Update ENTRADA Set COD_ESTADO = 1"
        sql = sql & vbCrLf & "  WHERE (COD_CLIENTE =" & Cod_cliente & ") AND (ELEMENTO =" & Elemento & ") AND  (TIPO =" & TIPO & ") AND (COD_ESTADO = 0)"
        conRemito.Execute sql
       End If
       
        
End Sub


Public Sub IngresarValorGrilla(Valor As String, Grilla As MSFlexGrid)
    Dim R As Integer
    Dim c As Integer
            For R = 1 To Grilla.Rows - 1
                For c = 1 To Grilla.Cols - 1
                    If Grilla.TextMatrix(R, c) = Valor Then
                        cajasRepetidas = cajasRepetidas & Valor & vbCrLf
                        Exit Sub
                    End If
                    If Grilla.TextMatrix(R, c) = "" Then
                        Grilla.TextMatrix(R, c) = Valor
                        ContarGrilla Grilla
                        Exit Sub
                    End If
                Next
            Next
            Grilla.AddItem ""
            Grilla.TextMatrix(Grilla.Rows - 1, 1) = Valor
            ContarGrilla Grilla
End Sub

Public Sub InsertarBajas(Cod_cliente As Integer, Elemento As Long, TIPO_ELEMENTO As Integer, Fecha As String, Responsable As Integer, Cod_REMITO As Long, conRemito As ADODB.Connection)
    Dim sql As String
        sql = "  INSERT INTO BAJAS (COD_CLIENTE, ELEMENTO, TIPO_ELEMENTO, FECHA,RESPONSABLE, COD_REMITO)"
        sql = sql & " VALUES (" & Cod_cliente & "," & Elemento & "," & TIPO_ELEMENTO & "," & Fecha & "," & Responsable & "," & Cod_REMITO & ")"
        conRemito.Execute sql
End Sub
Public Sub CargarRemitoEntrada()
                ctlCliente.Enabled = True
                ctlRemito_Tipo.Enabled = True
                ctlRemito_Tipo.Valor = Null
                ctlRemtito_Estado.Enabled = False
                ctlRemtito_Estado.Valor = 0
                ctlRemito_Operacion.Enabled = False
                ctlRemito_Operacion.Valor = 0
             Rem   ctlResponsable.Valor = MDIfrmInicio.StaInicio.Panels(2).Text
                ctlRemito_Tipo.Valor = 1
                ctlTipo_Elemento.Valor = 0
                ctlTipo_Elemento.Enabled = True
                ctlClienteFirma.Enabled = True
                mskRemitoConsulta.Enabled = True
                Rem mskRemitoCustodia.Enabled = True
                txtElemento.Enabled = True
             Rem   cmdColector.Enabled = False
End Sub
Public Sub CargarRequerimientoEnRemito(IDREQUERIMIENTO As Long)
    
    On Error GoTo salir:
    Dim rsRequerimiento As ADODB.Recordset
    Dim sSQL As String
        sSQL = "SELECT REQUERIMIENTO.IDREQUERIMIENTO,REQUERIMIENTO.ID_CLIENTE,"
        sSQL = sSQL & vbCrLf & " REQUERIMIENTO.SOLICITANTE,REQUERIMIENTO.FECHARECEPCION,IDPERSONAL,"
        sSQL = sSQL & vbCrLf & "REQUERIMIENTO.CANTIDAD,REQUELIBOSCAJAS.CAJASLIBROS, REQUELIBOSCAJAS.ESTADO as EstadoDetalle , "
        sSQL = sSQL & vbCrLf & " TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO,TIPOREQUERIMIENTO.TIPO_REMITO,"
        sSQL = sSQL & vbCrLf & " Tiporequerimiento.TIPO_REMITO_ESTADO , Tiporequerimiento.TIPO_ALMACENAMIENTO, "
        sSQL = sSQL & vbCrLf & " Requerimiento.DESCRIPCION,REQUERIMIENTO.COD_USUARIO_CLIENTE, FLETE  "
        sSQL = sSQL & vbCrLf & " From Tiporequerimiento, Requerimiento, REQUELIBOSCAJAS"
        sSQL = sSQL & vbCrLf & " Where Tiporequerimiento.IDTIPOREQUERIMIENTO = Requerimiento.IDTIPOREQUERIMIENTO"
        sSQL = sSQL & vbCrLf & " AND Requerimiento.IDREQUERIMIENTO = REQUELIBOSCAJAS.IDREQUERIMIENTOS"
        sSQL = sSQL & vbCrLf & " AND Requerimiento.IDREQUERIMIENTO =" & IDREQUERIMIENTO
        Set rsRequerimiento = New ADODB.Recordset
        rsRequerimiento.Open sSQL, ConActiva, 0, 1
                ctlCliente.Enabled = False
                ctlRemito_Tipo.Enabled = False
                ctlRemtito_Estado.Enabled = False
                ctlRemito_Operacion.Enabled = False
                ctlTipo_Elemento.Enabled = False
                ctlClienteFirma.Enabled = False
                mskRemitoConsulta.Enabled = False
                mskRemitoCustodia.Enabled = False
                txtElemento.Enabled = False
                cmdColector.Enabled = False
        
        With rsRequerimiento
            If Not (.EOF) Then
                RequerimientoRemito = IDREQUERIMIENTO
                ctlCliente.Valor = !ID_CLIENTE
                ctlRemito_Tipo.Valor = !TIPO_REMITO
                ctlRemtito_Estado.Valor = !TIPO_REMITO_ESTADO
                ctlRemito_Operacion.Valor = 1
                ctlTipo_Elemento.Valor = !TIPO_ALMACENAMIENTO
                Rem ctlResponsable.Valor = !IDPERSONAL
                If Not IsNull(!DESCRIPCION) Then
                    txtDescripcionConsulta.Text = Trim(!DESCRIPCION)
                End If
                If Not IsNull(!COD_USUARIO_CLIENTE) Then
                    ctlClienteFirma.Valor = !COD_USUARIO_CLIENTE
                End If
                If !Flete = 1 Then
                    chkFleteConsulta.Value = 1
                End If
            End If
            Do While Not (.EOF)
            If Mid(!EstadoDetalle, 1, 2) = "NO" Then
                MsgBox "El elemento no encontrado " & !CAJASLIBROS
            Else
                CargarGrilla (!CAJASLIBROS)
            End If
                .MoveNext
            Loop
            Clipboard.Clear
            Clipboard.SetText SrtValor
        End With
        Exit Sub
        
        
salir:
MsgBox Err.Description
End Sub

Public Function Validar_Datos() As Boolean
        Validar_Datos = True
        If IsNull(ctlCliente.Valor) Then
            MsgBox "Usted debe ingresar el cliente", vbInformation
            Validar_Datos = False
        End If
        If IsNull(ctlResponsable.Valor) Then
            MsgBox "Usted debe ingresar el responsable", vbInformation
            Validar_Datos = False
        End If
        If IsNull(ctlRemito_Tipo.Valor) Then
            MsgBox "Usted debe ingresar el movimiento", vbInformation
            Validar_Datos = False
        End If
        If IsNull(ctlTipo_Elemento.Valor) Then
            MsgBox "Usted debe ingresar el Tipo de elemento", vbInformation
            Validar_Datos = False
        End If
        If IsNull(ctlRemito_Operacion.Valor) Then
            MsgBox "Usted debe ingresar el Tipo de Operacion", vbInformation
            Validar_Datos = False
        End If
        If IsNull(ctlRemtito_Estado.Valor) Then
            MsgBox "Usted debe ingresar el Tipo de Estado", vbInformation
            Validar_Datos = False
        End If
        If Not (IsDate(mskFechaRemito.Text)) Then
            MsgBox "La fecha es incorrecta", vbInformation
            Validar_Datos = False
        End If
        If IsNull(ctlClienteFirma.Valor) Then
            MsgBox "Usted debe ingresar quien firmo", vbInformation
            Validar_Datos = False
        End If
        If lblCantidadGuardia.Caption <> "" Then
            If mskRemitoCustodia.Text = "0001-000_____" Then
                MsgBox "Usted Debe ingresar el Remito Custodia", vbCritical
                Validar_Datos = False
            End If
         End If
        If lblCantidadDevolucion.Caption <> "" Then
           If ctlRemito_Operacion.Valor = 0 Then
                If mskRemitoConsulta.Text = "0001-000_____" Then
                    MsgBox "Usted Debe ingresar el Remito de Consulta ", vbCritical
                    Validar_Datos = False
                End If
            End If
        End If
        If Len(txtDescripcionConsulta.Text) > 240 Then
            MsgBox "La descripcion es muy larga Consulta", vbInformation
            Validar_Datos = False
        End If
        If Len(txtDescripcionCustodia.Text) > 240 Then
            MsgBox "La descripcion es muy larga Custodia", vbInformation
            Validar_Datos = False
        End If
        
End Function

Public Sub ImprimirRemitoCaja(NRO_REMITO As Long)
    
    Dim rs As New ADODB.Recordset
    Dim NombreRemito As String
    Dim sql As String
    On Error GoTo salir
      NombreRemito = "RemitoCaja.rpt"
    sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.ID_CLIENTE, CLIENTES.EMPRESA"
    sql = sql & " FROM         REMITOS_CUERPO INNER JOIN"
    sql = sql & " CLIENTES ON REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE"
    sql = sql & " Where REMITOS_CUERPO.NRO_REMITO = " & NRO_REMITO
       rs.Open sql, ConActiva, 0, 1
    
    If Trim(rs!EMPRESA) = "CUSTODIA" Then
        NombreRemito = "RemitoCajaCustodia.rpt"
    End If
    
    If Trim(rs!EMPRESA) = "SECURITY" Then
        NombreRemito = "RemitoCajaSecurity.rpt"
    End If
    
    Dim sSQL As String
    Dim Formulas As CRAXDDRT.FormulaFieldDefinitions
    
    
    
    sSQL = " SELECT * FROM  V_REMITO_CAJA"
    sSQL = sSQL & " WHERE NRO_REMITO = " & NRO_REMITO
    
    frmReportes.ImprimirReporte PasoReportes & NombreRemito, sSQL, False, "ORIGINAL"
    frmReportes.ImprimirReporte PasoReportes & NombreRemito, sSQL, True, "ORIGINAL"
    If MsgBox("Usted quiere imprimir la copia", vbYesNo) = vbYes Then
        frmReportes.ImprimirReporte PasoReportes & NombreRemito, sSQL, False, "DUPLICADO"
    End If
    Exit Sub
salir:
    MsgBox Err.Description
End Sub
Public Sub ImprimirRemitoLibros(NRO_REMITO As Long)
 Dim rs As New ADODB.Recordset
    Dim NombreRemito As String
    Dim sql As String
      NombreRemito = "RemitoLibros.rpt"
    sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.ID_CLIENTE, CLIENTES.EMPRESA"
    sql = sql & " FROM         REMITOS_CUERPO INNER JOIN"
    sql = sql & " CLIENTES ON REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE"
    sql = sql & " Where REMITOS_CUERPO.NRO_REMITO = " & NRO_REMITO
    rs.Open sql, ConActiva, 0, 1
    
    If Trim(rs!EMPRESA) = "CUSTODIA" Then
        NombreRemito = "RemitoLibrosCustodia.rpt"
    End If
        

    
    Dim sSQL As String
    sSQL = " SELECT * FROM  V_REMITO_LIBRO"
    sSQL = sSQL & " WHERE NRO_REMITO = " & NRO_REMITO
    frmReportes.ImprimirReporte PasoReportes & NombreRemito, sSQL, False, "ORIGINAL"
    If MsgBox("Usted quiere imprimir la copia", vbYesNo) = vbYes Then
    frmReportes.ImprimirReporte PasoReportes & NombreRemito, sSQL, False, "DUPLICADO"
    End If
End Sub
Public Sub ImprimirRemitoLegajos(NRO_REMITO As Long)
    
 Dim rs As New ADODB.Recordset
    Dim NombreRemito As String
    Dim sql As String
      NombreRemito = "RemitoLegajos.rpt"
    sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.ID_CLIENTE, CLIENTES.EMPRESA"
    sql = sql & " FROM         REMITOS_CUERPO INNER JOIN"
    sql = sql & " CLIENTES ON REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE"
    sql = sql & " Where REMITOS_CUERPO.NRO_REMITO = " & NRO_REMITO

       rs.Open sql, ConActiva, 0, 1
    
    If Trim(rs!EMPRESA) = "CUSTODIA" Then
        NombreRemito = "RemitoLegajosCustodia.rpt"
    End If
    
    If Trim(rs!EMPRESA) = "SECURITY" Then
        NombreRemito = "RemitoLegajosSecurity.rpt"
    End If
    
    Dim sSQL As String
        sSQL = " SELECT * FROM  V_REMITO_LEGAJOS"
        sSQL = sSQL & " WHERE NRO_REMITO = " & NRO_REMITO
        sSQL = sSQL & " ORDER BY  DESDE "
        frmReportes.ImprimirReporte PasoReportes & NombreRemito, sSQL, False, "ORIGINAL"
         If MsgBox("Usted quiere imprimir la copia", vbYesNo) = vbYes Then
            frmReportes.ImprimirReporte PasoReportes & NombreRemito, sSQL, False, "DUPLICADO"
        End If
End Sub
Public Sub UpdateRequerimientoRemito(Requerimiento As Long, Remito As Long, conexion As ADODB.Connection)
    conexion.Execute ("UPDATE REQUERIMIENTO SET IDREMITO = " & Remito & " Where idRequerimiento =" & Requerimiento)
End Sub

Public Sub InsertarSectorCaja(Cod_cliente As Integer, NRO_CAJA As Long, Indice As String, Fecha As String, _
    COD_USUARIO_CLIENTE As Integer)
'Dim Sql As String
'    Sql = " INSERT INTO SECTORCAJAS"
'    Sql = Sql & " (COD_CLIENTE, NRO_CAJA, INDICE, FECHA,"
'    Sql = Sql & "     COD_USUARIO_CLIENTE)"
'    Sql = Sql & "  VALUES "
'    Sql = Sql & "( " & Cod_Cliente & "," & NRO_CAJA & ",'" & Indice & "','" & Fecha & "',"
'    Sql = Sql & COD_USUARIO_CLIENTE & ")"
'    ExecutarSql Sql

End Sub

Public Sub AnularRemito(NRO_REMITO As Long)
    Dim rsRemito As ADODB.Recordset
    Dim RsMov As ADODB.Recordset
    Dim sSQL, Rsql As String
    Dim NroRemito As Long
        
        On Error GoTo salir:
        
        Dim conRemto As New ADODB.Connection
        
        conRemto.Open strConBasa
        conRemto.BeginTrans
        Rsql = "   SELECT REMITOS_CUERPO.TIPO, REMITOS_CUERPO.OPERACION,"
        Rsql = Rsql & " REMITOS_CUERPO.ESTADO, REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO,"
        Rsql = Rsql & " REMITOS_CUERPO.ID_CLIENTE,REMITOS_CUERPO.CANTIDAD, REMITOS_DETALLE.DESDE,"
        Rsql = Rsql & " REMITOS_CUERPO.NRO_Remito ,REMITOS_CUERPO.FECHA"
        Rsql = Rsql & " From REMITOS_CUERPO, REMITOS_DETALLE"
        Rsql = Rsql & " Where REMITOS_CUERPO.NRO_Remito = REMITOS_DETALLE.NRO_Remito"
        Rsql = Rsql & " AND anulado is null"
        Rsql = Rsql & " AND REMITOS_CUERPO.NRO_REMITO = " & NRO_REMITO
        Set rsRemito = New ADODB.Recordset
        rsRemito.Open Rsql, ConActiva, 0, 1
        Do While Not rsRemito.EOF
        ' falta colocar el tipo de elemento
        sSQL = " SELECT NRO_REMITO From MOV_CAJAS2"
        sSQL = sSQL & " Where ID_CLIENTE =" & rsRemito!ID_CLIENTE
        sSQL = sSQL & " And NRO_CAJA = " & rsRemito!DESDE
        sSQL = sSQL & " ORDER BY NRO_REMITO DESC "
        Set RsMov = New ADODB.Recordset
             RsMov.Open sSQL, ConActiva, 0, 1
             If RsMov!NRO_REMITO <> rsRemito!NRO_REMITO Then
                 MsgBox "El elemento " & rsRemito!DESDE & vbCrLf & "Tiene movimiento posteriores", vbInformation
                 Exit Sub
             End If
            rsRemito.MoveNext
        Loop
        'strConBasa.BeginTrans
        Set rsRemito = New ADODB.Recordset
        rsRemito.Open Rsql, ConActiva, 0, 1
        With rsRemito
        If rsRemito.EOF Then
            MsgBox "El remito No existe", vbCritical
            Exit Sub
        Else
        NroRemito = rsRemito!NRO_REMITO
        End If
        
        
        Select Case !TIPO
        Case 0 'Guarda y custodia
            Do While Not rsRemito.EOF
                UpdateContedor !DESDE, !ID_CLIENTE, 2, 5, RemitoAnterior(!DESDE, !ID_CLIENTE, 0), conRemto
                rsRemito.MoveNext
            Loop
        Case 1 ' Consulta
            Do While Not rsRemito.EOF
                If !OPERACION = 0 Then 'ENTRADA
                    Select Case !COD_TIPO_ALMACENAMIENTO
                    Case 0 'CAJA
                        UpdateContedor !DESDE, !ID_CLIENTE, 2, 3, RemitoAnterior(!DESDE, !ID_CLIENTE, 0), conRemto
                    Case 1 'lIBRO
                        UpdateLibros !DESDE, !ID_CLIENTE, 2, 3, RemitoAnterior(!DESDE, !ID_CLIENTE, 1), conRemto
                    Case 3 'LEGAJO
                        UpdateLegajos !DESDE, !ID_CLIENTE, 2, 3, RemitoAnterior(!DESDE, !ID_CLIENTE, 3), !Fecha, conRemto
                    End Select
                End If
                If !OPERACION = 1 Then 'SALIDA
                    Select Case !COD_TIPO_ALMACENAMIENTO
                    Case 0 'CAJA
                        UpdateContedor !DESDE, !ID_CLIENTE, 3, 2, RemitoAnterior(!DESDE, !ID_CLIENTE, 0), conRemto
                    Case 1 'lIBRO
                        UpdateLibros !DESDE, !ID_CLIENTE, 3, 2, RemitoAnterior(!DESDE, !ID_CLIENTE, 1), conRemto
                    Case 3 'LEGAJO
                        UpdateLegajos !DESDE, !ID_CLIENTE, 3, 2, RemitoAnterior(!DESDE, !ID_CLIENTE, 3), !Fecha, conRemto
                    End Select
                End If
                rsRemito.MoveNext
            Loop
        Case 2 ' Cajas Vacias
            Do While Not rsRemito.EOF
                AnuladoCajasVacias !ID_CLIENTE, !DESDE
                rsRemito.MoveNext
            Loop
        End Select
            UpdateRemitoAnulado NroRemito
            'strConBasa.CommitTrans
  End With
     conRemto.CommitTrans
     MsgBox "El remito fue anulado con exito", vbInformation
     Exit Sub
     
salir:
     conRemto.RollbackTrans
     MsgBox "Error", vbInformation
End Sub

Public Sub UpdateRemitoAnulado(NRO_REMITO As Long, Optional DetalleAnulado As String)
        Dim sql As String
        sql = " Update MOV_CAJAS2 Set ANULADO = 1 "
        sql = sql & " Where NRO_Remito = " & NRO_REMITO
        ExecutarSql sql
        sql = "  Update REMITOS_CUERPO"
        sql = sql & vbCrLf & " SET DETALLE_ANULADO ='Anulado Sistem de Remito',"
        sql = sql & vbCrLf & " ANULADO =1, AUDIT_USUARIO ='" & UserName & "', AUDIT_FECHA =" & SysDate
        sql = sql & vbCrLf & " Where NRO_REMITO =" & NRO_REMITO
        ExecutarSql sql
        sql = " Update MOVIMIENTOS_ELEMENTOS"
        sql = sql & " Set ANULADO = 1"
        sql = sql & " Where COD_REMITO =" & NRO_REMITO
        ExecutarSql sql
End Sub

Public Function RemitoAnterior(Elemento As Long, Cod_cliente As Integer, TipoElemento As Integer) As Long
        Dim RsMov As ADODB.Recordset
        Dim sSQL As String
        sSQL = " SELECT COD_REMITO"
        sSQL = sSQL & " From MOVIMIENTOS_ELEMENTOS"
        sSQL = sSQL & " WHERE COD_TIPO_ALMACENAMIENTO = " & TipoElemento
        sSQL = sSQL & " AND  ELEMENTO = " & Elemento
        sSQL = sSQL & " AND COD_CLIENTE = " & Cod_cliente
        sSQL = sSQL & " ORDER BY COD_REMITO DESC "
        RemitoAnterior = 0
        
        Set RsMov = New ADODB.Recordset
             RsMov.Open sSQL, ConActiva, 0, 1
             If Not RsMov.EOF Then
                RsMov.MoveNext
             End If
             If Not RsMov.EOF Then
                RemitoAnterior = RsMov!Cod_REMITO
             End If
End Function

Public Sub AnuladoCajasVacias(Cod_cliente As Integer, NRO_CAJA As Long)
    Dim sql As String
        sql = " UPDATE CONTENEDOR SET COD_CLIENTE = NULL, NRO_CAJA = ID_CONTENEDOR, Estado = 1"
        sql = sql & " WHERE ESTADO = 5 "
        sql = sql & " AND COD_CLIENTE =" & Cod_cliente
        sql = sql & " AND NRO_CAJA =" & NRO_CAJA
        
        ExecutarSql sql
        
     sql = "   Update basasql.dbo.Cajas"
sql = sql & " SET  FK_CLIENTE =Null, FK_ESTADO =4 "
sql = sql & " Where FK_CLIENTE = " & Cod_cliente
sql = sql & "  And NRO_CAJA = " & NRO_CAJA


        ExecutarSql sql

        
End Sub

Private Sub txtElementoHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim i As Long
    If Not (IsNumeric(txtElemento.Text) And IsNumeric(txtElementoHasta.Text)) Then
        MsgBox "Error en el tipo de elemento", vbInformation
        Exit Sub
    End If
    For i = txtElemento.Text To txtElementoHasta.Text
        CargarGrilla Str(i)
    Next
End If

End Sub
Public Sub BajasLegajos(ID_CLIENTE_LEGAJO As Long, Cod_cliente As Integer, NRO_REMITO As Long, COD_CLIENTE_USUARIO As Integer, conRemito As ADODB.Connection)
   Dim rsBajas As ADODB.Recordset
   Dim sql As String
        sql = " INSERT INTO BAJASLEGAJOS "
        sql = sql & vbCrLf & " (ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE,"
        sql = sql & vbCrLf & " CLIENTE_LEGAJO, DESCRIPCION, NRO_CAJA, COD_CLIENTE,"
        sql = sql & vbCrLf & " COD_UBICACION, COD_ESTADO, NOMBRE, COD_REMITO,"
        sql = sql & vbCrLf & " FECHA )"
        sql = sql & vbCrLf & " SELECT ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE,"
        sql = sql & vbCrLf & " CLIENTE_LEGAJO, DESCRIPCION, NRO_CAJA, COD_CLIENTE,"
        sql = sql & vbCrLf & " COD_UBICACION, COD_ESTADO, NOMBRE, COD_REMITO, Fecha"
        sql = sql & vbCrLf & " From Legajos"
        sql = sql & vbCrLf & " Where ID_CLIENTE_LEGAJO in( " & ID_CLIENTE_LEGAJO & ")"
        sql = sql & vbCrLf & " And COD_CLIENTE = " & Cod_cliente
        conRemito.Execute sql
        sql = "  Update LEGAJOS"
        sql = sql & vbCrLf & " SET COD_ESTADO = 0 "
        sql = sql & vbCrLf & " ,Nombre ='Legajo de Baja'"
        sql = sql & vbCrLf & ", Descripcion='Legajo de Baja'"
        sql = sql & vbCrLf & " ,FECHA_ACTUALIZACION =" & SysDate
        sql = sql & vbCrLf & " , COD_REMITO = " & NRO_REMITO
        sql = sql & vbCrLf & " WHERE ID_CLIENTE_LEGAJO in (" & ID_CLIENTE_LEGAJO & ")"
        sql = sql & vbCrLf & " AND COD_CLIENTE = " & Cod_cliente
        conRemito.Execute sql
        

        
End Sub


Public Sub LlegarDescripcion(Cod_cliente As Integer, ID_CLIENTE_LEGAJO As Long)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim dato As String
    sql = " SELECT     NRO_DESDE, LETRA_DESDE, FECHA_DESDE, CLIENTE_LEGAJO, DESCRIPCION, NOMBRE"
    sql = sql & " From LEGAJOS"
    sql = sql & "  Where ID_CLIENTE_LEGAJO = " & ID_CLIENTE_LEGAJO
    sql = sql & "  And Cod_cliente = " & Cod_cliente


If Not rs.EOF Then

        If Not IsNull(rs!NRO_DESDE) Then
            dato = Trim(rs!NRO_DESDE)
        
        Else
            dato = Trim(rs!CLIENTE_LEGAJO)
        End If
        
        
        If Not IsNull(rs!LETRA_DESDE) Then
            dato = dato & " " & rs!LETRA_DESDE
        Else
             dato = dato & " " & Trim(rs!Nombre)
        End If

       If Not IsNull(rs!FECHA_DESDE) Then
            dato = dato & " " & Format(rs!FECHA_DESDE, "YYYY")
       End If
       
       

    sql = " Update LEGAJOS "
    sql = sql & " SET  DESCRIPCION_REMITO ='" & Trim(dato) & "'"
    sql = sql & "  Where ID_CLIENTE_LEGAJO = " & ID_CLIENTE_LEGAJO
    sql = sql & "  And Cod_cliente = " & Cod_cliente
    ExecutarSql sql
    
   
 End If
 




End Sub

Public Sub InsertarTipoReferencia(Cod_cliente As Integer, Caja As Long, TipoRef As String)

Dim ConReferencia As New ADODB.Connection

ConReferencia.Open strConBasa
Dim sql As String

        sql = " UPDATE CAJAS "
        sql = sql & " SET TIPO_REFERENCIA = '" & TipoRef & "'"
        sql = sql & "  Where FK_CLIENTE = " & Cod_cliente
        sql = sql & "  And NRO_CAJA = " & Caja
        ConReferencia.Execute sql

End Sub

Public Function descrpLegajos(Cliente As Integer, Etiqueta As Long) As String
    Dim sql As String
    
    Dim rs As New ADODB.Recordset
    
    sql = " SELECT        ID_CLIENTE_LEGAJO, COD_CLIENTE, NRO_DESDE, LETRA_DESDE, FECHA_DESDE, DESCRIPCION"
sql = sql & "  From LEGAJOS "
sql = sql & "  Where ID_CLIENTE_LEGAJO =" & Etiqueta
sql = sql & "  And Cod_cliente = " & Cliente
    Dim NRO_DESDE As String
    Dim LETRA_DESDE As String
    Dim FECHA_DESDE As String
    Dim DESCRIPCION As String
    
    rs.Open sql, ConActiva, 0, 1
    
    
    If Not rs.EOF Then
    
    
    If Not IsNull(rs!NRO_DESDE) Then
    
    NRO_DESDE = rs!NRO_DESDE
    End If
    
    
    If Not IsNull(rs!LETRA_DESDE) Then
    
    LETRA_DESDE = rs!LETRA_DESDE
    End If
    
    
    If Not IsNull(rs!FECHA_DESDE) Then
        FECHA_DESDE = Format(rs!FECHA_DESDE, "yyyy")
    End If
    
    If Not IsNull(rs!DESCRIPCION) Then
    
    DESCRIPCION = rs!DESCRIPCION
    End If
    
    descrpLegajos = "Etiqueta: " & Etiqueta & " Descripción: " & NRO_DESDE & " " & LETRA_DESDE & " " & FECHA_DESDE & " " & DESCRIPCION
    
    End If




End Function

Public Sub BorraReferencias(Cliente As Integer, Caja As Long, conRemito As ADODB.Connection)

conRemito.Execute (" DELETE FROM REFERENCIAS Where Cod_cliente =" & Cliente & "  And NRO_CAJA = " & Caja)

conRemito.Execute (" DELETE FROM LEGAJOS  Where Cod_cliente = " & Cliente & " And NRO_CAJA = " & Caja)



End Sub



Public Function AsingacionCajas(Cliente As Long, Caja As Long) As Boolean
AsingacionCajas = False

Dim rs As ADODB.Recordset
Dim rsControl As ADODB.Recordset

Dim sql As String

If Caja > 0 Then
            If Cliente = 39 Then
                sql = " SELECT  top (1)   ID_CAJA, FK_CLIENTE, NRO_CAJA "
                sql = sql & " From Cajas "
                sql = sql & " Where ID_CAJA BETWEEN 384941 and 400000  And (FK_CLIENTE Is Null)"
                sql = sql & "  ORDER BY ID_CAJA"
                Set rs = New ADODB.Recordset
                rs.Open sql, ConActiva
                If Not rs.EOF Then
                    sql = " Update Cajas"
                    sql = sql & "  SET FK_CLIENTE = " & Cliente & ", FK_ESTADO = 1150 , NRO_CAJA = " & Caja
                    sql = sql & " Where ID_CAJA = " & rs!ID_CAJA & " And (FK_CLIENTE Is Null)"
                    ExecutarSql sql
                End If
            Else
                sql = " Update Cajas"
                sql = sql & "  SET FK_CLIENTE = " & Cliente & ", FK_ESTADO = 1150"
                sql = sql & " Where ID_CAJA = " & Caja & " And (FK_CLIENTE Is Null)"
                ExecutarSql sql
            End If
            
            
            If Caja > 300000 Then
            
                    sql = " SELECT     ID_CONTENEDOR, ESTADO, COD_CLIENTE"
                    sql = sql & " From basasql.dbo.CONTENEDOR"
                    sql = sql & " Where NRO_CAJA = " & Caja
                    sql = sql & " And Not (COD_CLIENTE Is Null)"
                    Set rs = New ADODB.Recordset
                    rs.Open sql, strConBasa
                      If Not rs.EOF Then
                            sql = " Update basasql.dbo.CONTENEDOR"
                            sql = sql = " SET   COD_CLIENTE =" & Cliente
                            sql = sql = "  Where NRO_CAJA = " & Caja
                            sql = sql = "  And (Not (COD_CLIENTE Is Null))"
                            ExecutarSql sql
                      Else
                      
                      
                      sql = " SELECT  TOP 1 ID_CONTENEDOR, NRO_CAJA, COD_CLIENTE, ESTADO, FK_CAJAS,F_MODIFICACION"
                    sql = sql & "  From CONTENEDOR"
                    sql = sql & "  WHERE ESTADO = 1 AND COD_CLIENTE IS NULL  "
                    sql = sql & " AND ESTANTERIA BETWEEN 150 AND 190 "
                    Set rs = New ADODB.Recordset
                    rs.Open sql, ConActiva, 0, 1
                    If Not rs.EOF Then
                        sql = " Update CONTENEDOR"
                        sql = sql & " SET ESTADO = 5, COD_CLIENTE =" & Cliente & ", NRO_CAJA =" & Caja
                        sql = sql & " Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR
                        ExecutarSql sql
                        AsingacionCajas = True
                    End If
                      End If
              
              Else
            
                    sql = " SELECT  TOP 1 ID_CONTENEDOR, NRO_CAJA, COD_CLIENTE, ESTADO, FK_CAJAS,F_MODIFICACION"
                    sql = sql & "  From CONTENEDOR"
                    sql = sql & "  WHERE ESTADO = 1 AND COD_CLIENTE IS NULL  "
                    sql = sql & " AND ESTANTERIA BETWEEN 150 AND 190 "
                    Set rs = New ADODB.Recordset
                    rs.Open sql, ConActiva, 0, 1
                    If Not rs.EOF Then
                        sql = " Update CONTENEDOR"
                        sql = sql & " SET ESTADO = 5, COD_CLIENTE =" & Cliente & ", NRO_CAJA =" & Caja
                        sql = sql & " Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR
                        ExecutarSql sql
                        AsingacionCajas = True
                    End If
           End If
  Else
        MsgBox "La caja no puede ser asignada"
        Exit Function
  End If
  
        

End Function

Public Sub ReferenciasApc(Cliente As Integer, Caja As Long, Indice As String, DESCRIPCION As String)
Dim sql As String

sql = sql & " INSERT INTO basasql.dbo.REFERENCIAS"
sql = sql & "(COD_CLIENTE"
sql = sql & " , NRO_CAJA"
sql = sql & " , INDICE"
sql = sql & " , DESCRIPCION"
sql = sql & " , FECHA_CREACION"
sql = sql & " , FECHA_MODIFICACION"
sql = sql & " , BORRADO)"
sql = sql & " VALUES   (  "
sql = sql & Cliente
sql = sql & " ," & Caja
sql = sql & " ,'" & ctlClienteFirma.Indice & "'"
sql = sql & " , '" & DESCRIPCION & "'"
sql = sql & " , " & SysDate
sql = sql & " , " & SysDate
sql = sql & " , 0)"
ExecutarSql sql

End Sub
