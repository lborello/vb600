VERSION 5.00
Begin VB.Form frmActualizacion_legajos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ADMINISTRADOR"
   ClientHeight    =   8745
   ClientLeft      =   3000
   ClientTop       =   2820
   ClientWidth     =   13770
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   13770
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   555
      Left            =   7080
      TabIndex        =   72
      Top             =   7680
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Cambio de orden"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10260
      TabIndex        =   62
      Top             =   7740
      Width           =   1695
   End
   Begin VB.TextBox Text3 
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
      Left            =   10680
      TabIndex        =   61
      Top             =   6300
      Width           =   1575
   End
   Begin VB.TextBox Text2 
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
      Left            =   10680
      TabIndex        =   60
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox Text1 
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
      Left            =   10680
      TabIndex        =   59
      Top             =   7140
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Caption         =   "IMPRESIÓN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1815
      Left            =   3720
      TabIndex        =   56
      Top             =   6120
      Width           =   2775
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir Etiquetas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   300
         TabIndex        =   58
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Imprimir caja con estanteria"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   180
         TabIndex        =   57
         Top             =   1080
         Width           =   2355
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "REQUERIMIENTOS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1815
      Left            =   180
      TabIndex        =   50
      Top             =   6120
      Width           =   3255
      Begin VB.CommandButton Command2 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   53
         Top             =   1320
         Width           =   1515
      End
      Begin VB.TextBox txtRequerimiento 
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
         Left            =   1560
         TabIndex        =   52
         Top             =   420
         Width           =   1515
      End
      Begin VB.TextBox txtRequeEstado 
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
         Left            =   1560
         TabIndex        =   51
         Top             =   840
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "Requerimiento"
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
         Left            =   180
         TabIndex        =   55
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Estado"
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
         Left            =   180
         TabIndex        =   54
         Top             =   900
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdCajasUgarte 
      Caption         =   "Cajas Ugarte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11220
      TabIndex        =   9
      Top             =   4260
      Width           =   1755
   End
   Begin VB.CommandButton cmdCajasConLegajos 
      Caption         =   "Cajas con legajos en referencias"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   8
      Top             =   2520
      Width           =   2475
   End
   Begin VB.CommandButton cmdRecuperarCajas 
      Caption         =   "Recuperar Cajas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11220
      TabIndex        =   7
      Top             =   3780
      Width           =   1755
   End
   Begin VB.Frame Frame2 
      Caption         =   "OSEP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1035
      Left            =   6780
      TabIndex        =   5
      Top             =   6120
      Width           =   2295
      Begin VB.CommandButton cmdCambioOsep 
         Caption         =   "Cambio Osep"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdBorrarLectura 
      Caption         =   "Borrar Lectura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11220
      TabIndex        =   4
      Top             =   4740
      Width           =   1755
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado Cajas Por Lectura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   8400
      TabIndex        =   1
      Top             =   3360
      Width           =   2595
      Begin VB.ComboBox cboestado 
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
         ItemData        =   "frmActualizacion_legajos.frx":0000
         Left            =   120
         List            =   "frmActualizacion_legajos.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdCambioEstadoPorLectura 
         Caption         =   "Estado caja por lectura"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Liberar cajas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11220
      TabIndex        =   0
      Top             =   3300
      Width           =   1755
   End
   Begin VB.Frame Frame3 
      Caption         =   "ETIQUETAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   3015
      Left            =   180
      TabIndex        =   10
      Top             =   180
      Width           =   13035
      Begin VB.TextBox txtEtiquetaBorrarCaja 
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
         Left            =   10980
         TabIndex        =   75
         Text            =   "0"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtEtiquetaBorrarHasta 
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
         Left            =   10980
         TabIndex        =   73
         Text            =   "0"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Borrar Etiqueta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11040
         TabIndex        =   34
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtEtiquetaBorrarDesde 
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
         Left            =   10980
         TabIndex        =   31
         Text            =   "0"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtClienteEtiquetaBorrar 
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
         Left            =   10980
         TabIndex        =   30
         Text            =   "0"
         Top             =   300
         Width           =   1575
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Borrar todos los legajos de una caja"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   7020
         TabIndex        =   27
         Top             =   1500
         Width           =   1935
      End
      Begin VB.TextBox txtBorrarCajasLegajos 
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
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtBorrarclienteLegajos 
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
         Top             =   540
         Width           =   1575
      End
      Begin VB.CommandButton cmdActualizarEstado 
         Caption         =   "Estado Legajo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4740
         TabIndex        =   21
         Top             =   1860
         Width           =   1455
      End
      Begin VB.TextBox txtClienteLegajo 
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
         Left            =   4620
         TabIndex        =   20
         Top             =   540
         Width           =   1575
      End
      Begin VB.TextBox txtEtiqueta 
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
         Left            =   4620
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtEstadoLegajo 
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
         Left            =   4620
         TabIndex        =   18
         Top             =   1380
         Width           =   1575
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   14
         Top             =   1860
         Width           =   1575
      End
      Begin VB.TextBox txtCajaDestino 
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
         Left            =   1680
         TabIndex        =   13
         Top             =   1380
         Width           =   1455
      End
      Begin VB.TextBox txtHasta 
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
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtDesde 
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
         Left            =   1680
         TabIndex        =   11
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label23 
         Caption         =   "Caja"
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
         Left            =   9720
         TabIndex        =   76
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label22 
         Caption         =   "Etiqueta Hasta"
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
         Left            =   9720
         TabIndex        =   74
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Etiqueta Desde "
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
         Left            =   9720
         TabIndex        =   33
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Cliente 
         Caption         =   "Cliente"
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
         Left            =   9720
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Caja"
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
         Left            =   6600
         TabIndex        =   29
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label14 
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
         Height          =   315
         Left            =   6600
         TabIndex        =   28
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label8 
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
         Height          =   315
         Left            =   3780
         TabIndex        =   24
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Etiqueta"
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
         Left            =   3780
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.Label EStadoLegajo 
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
         Height          =   255
         Left            =   3780
         TabIndex        =   22
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Caja destino"
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
         Left            =   180
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Etiqueta Hasta"
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
         Left            =   180
         TabIndex        =   16
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Etiqueta Desde:"
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
         Left            =   180
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame CAJAS 
      Caption         =   "CAJAS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2475
      Left            =   240
      TabIndex        =   35
      Top             =   3480
      Width           =   13035
      Begin VB.Frame Frame6 
         Height          =   1815
         Left            =   5100
         TabIndex        =   66
         Top             =   360
         Width           =   2895
         Begin VB.TextBox txtCajaCustodia 
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
            Left            =   840
            TabIndex        =   71
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtClienteCustodia 
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
            Left            =   840
            TabIndex        =   70
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmsReuperarCajasCustodia 
            Caption         =   "Recuperar Cajas Custodia"
            Height          =   375
            Left            =   120
            TabIndex        =   67
            Top             =   1200
            Width           =   2655
         End
         Begin VB.Label Label21 
            Caption         =   "Caja"
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
            TabIndex        =   69
            Top             =   840
            Width           =   555
         End
         Begin VB.Label Label20 
            Caption         =   "Cliente Inicial"
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
            TabIndex        =   68
            Top             =   360
            Width           =   795
         End
      End
      Begin VB.TextBox txtClienteInicial 
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
         Left            =   3540
         TabIndex        =   46
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox txtClienteFinal 
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
         Left            =   3540
         TabIndex        =   45
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCambioCaja 
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
         Left            =   3540
         TabIndex        =   44
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Cambio de Cliente"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2940
         TabIndex        =   43
         Top             =   1740
         Width           =   1695
      End
      Begin VB.TextBox txtCajaEstado 
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
         Left            =   1080
         TabIndex        =   39
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox txtCajaCaja 
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
         Left            =   1080
         TabIndex        =   38
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtClienteCaja 
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
         Left            =   1080
         TabIndex        =   37
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Estado Caja"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   660
         TabIndex        =   36
         Top             =   1740
         Width           =   1515
      End
      Begin VB.Label Label15 
         Caption         =   "Cliente Inicial"
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
         Left            =   2340
         TabIndex        =   49
         Top             =   540
         Width           =   795
      End
      Begin VB.Label Label16 
         Caption         =   "Cliente Final"
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
         Left            =   2340
         TabIndex        =   48
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label17 
         Caption         =   "Caja:"
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
         Left            =   2340
         TabIndex        =   47
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label Label11 
         Caption         =   "Caja"
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
         Left            =   240
         TabIndex        =   41
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label12 
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
         Height          =   315
         Left            =   240
         TabIndex        =   40
         Top             =   540
         Width           =   735
      End
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000004&
      Caption         =   "Caja:"
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
      Left            =   9720
      TabIndex        =   65
      Top             =   6420
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "Orden."
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
      Left            =   9720
      TabIndex        =   64
      Top             =   6840
      Width           =   795
   End
   Begin VB.Label Label5 
      Caption         =   "Caja Final:"
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
      Left            =   9660
      TabIndex        =   63
      Top             =   7260
      Width           =   1035
   End
End
Attribute VB_Name = "frmActualizacion_legajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdActualizar_Click()
Dim RS As New ADODB.Recordset
Dim SQL As String
SQL = "SELECT COUNT(*) AS cantidad  "
SQL = SQL & "  From LEGAJOS"
SQL = SQL & " WHERE ID_LEGAJO BETWEEN " & txtDesde.Text
SQL = SQL & "  AND  " & txtHasta.Text

RS.Open SQL, strConBasa

 If Not RS.EOF Then
    If MsgBox(" La cantidad es " & RS!cantidad & " quiere continuar ? ", vbYesNo) = vbYes Then
        
        SQL = "  Update LEGAJOS"
        SQL = SQL & " Set NRO_CAJA = " & txtCajaDestino.Text
        SQL = SQL & " WHERE ID_LEGAJO BETWEEN " & txtDesde.Text
        SQL = SQL & "  AND  " & txtHasta.Text
        ExecutarSql SQL
        MsgBox "Terminado", vbInformation
    End If
    
 
 End If
 



End Sub

Private Sub cmdActualizarEstado_Click()

Dim SQL As String

SQL = " Update LEGAJOS "
SQL = SQL & " SET COD_ESTADO =" & txtEstadoLegajo.Text
SQL = SQL & "  Where ID_CLIENTE_LEGAJO = " & txtEtiqueta.Text
SQL = SQL & "  And COD_CLIENTE = " & txtClienteLegajo.Text

ExecutarSql SQL

End Sub

Private Sub cmdBorrarLectura_Click()

Dim SQL As String
Dim Lectura As Long

Lectura = InputBox("INGRESE EL NUMERO DE LECTURA")

SQL = " DELETE FROM LECTURACOLECTOR Where NUMERO_LECTURA = " & Lectura
ExecutarSql SQL

SQL = " SELECT     NUMERO_LECTURA From LECTURA_COLECTOR_CUERPO Where NUMERO_LECTURA = " & Lectura
ExecutarSql SQL


MsgBox "Lectura Borrada"



End Sub

Private Sub cmdCajasConLegajos_Click()
Dim SQL As String
Dim fecha As String
Dim RS As New ADODB.Recordset
SQL = " SELECT     LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA, LEGAJOS.COD_INDICE"
SQL = SQL & " FROM         LEGAJOS LEFT OUTER JOIN"
SQL = SQL & " REFERENCIAS ON LEGAJOS.NRO_CAJA = REFERENCIAS.NRO_CAJA AND LEGAJOS.COD_CLIENTE = REFERENCIAS.COD_CLIENTE"
SQL = SQL & "  Where (REFERENCIAS.NRO_CAJA Is Null) and not (LEGAJOS.NRO_CAJA is null) and not ( LEGAJOS.COD_CLIENTE is null)  "
SQL = SQL & "  GROUP BY LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA, LEGAJOS.COD_INDICE"
SQL = SQL & "  ORDER BY LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_CAJA, LEGAJOS.COD_INDICE"
fecha = SysDateMinutoSegundo
RS.Open SQL, strConBasa

Do While Not RS.EOF
    SQL = " INSERT INTO basasql.dbo.REFERENCIAS "
    SQL = SQL & " (COD_CLIENTE "
    SQL = SQL & "  , NRO_CAJA"
    SQL = SQL & " , INDICE"
    SQL = SQL & " , DESCRIPCION"
    SQL = SQL & " , FECHA_MODIFICACION"
    SQL = SQL & " , FECHA_CREACION"
    SQL = SQL & " , USUARIO_MODIFICACION"
    SQL = SQL & " , FK_PERSONAL_CREACION"
    SQL = SQL & " , FK_PERSONAL_MODIFICACION"
    SQL = SQL & " , BORRADO)"
    SQL = SQL & "  VALUES     ("
    SQL = SQL & RS!COD_CLIENTE
    SQL = SQL & "  , " & RS!NRO_CAJA
    SQL = SQL & " ,'" & RS!Cod_Indice & "'"
    SQL = SQL & " , '" & "LEGAJOS DE SISTEMA" & "'"
    SQL = SQL & " , " & fecha
    SQL = SQL & " , " & fecha
    SQL = SQL & " , 17"
    SQL = SQL & " , 17"
    SQL = SQL & " , 17"
    SQL = SQL & " , 0)"
    ExecutarSql SQL
    
    RS.MoveNext
Loop

End Sub

Private Sub cmdCajasUgarte_Click()
    Dim SQL As String
    Dim estado As String
    Dim RS As New ADODB.Recordset
    Dim Lectura As Long

    
Lectura = InputBox("Ingrese el numero de lectura")

        Rem para crear en cajas
        
        SQL = " SELECT  LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.NUMERO_LECTURA, CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"
        SQL = SQL & "   FROM CAJAS RIGHT OUTER JOIN "
        SQL = SQL & "   LECTURACOLECTOR ON CAJAS.ID_CAJA = LECTURACOLECTOR.CAJA"
        SQL = SQL & "   Where LECTURACOLECTOR.NUMERO_LECTURA = " & Lectura
        SQL = SQL & "   And (LECTURACOLECTOR.Cliente = 1002) "
        SQL = SQL & "   And (Cajas.ID_CAJA Is Null) "
        RS.Open SQL, strConBasa
        Do While Not RS.EOF
        SQL = " Insert Into basasql.dbo.Cajas("
        SQL = SQL & " ID_CAJA "
        SQL = SQL & " , FK_CLIENTE "
        SQL = SQL & " , NRO_CAJA "
        SQL = SQL & " , FK_ESTADO "
        SQL = SQL & " , FECHA_CREACION_CAJA "
        SQL = SQL & " , FK_USUARIO_CREACION_CAJA "
        SQL = SQL & " , DIGITO_VERIFICADOR) "
        SQL = SQL & "   VALUES ( "
        SQL = SQL & RS!Caja
        SQL = SQL & "  ,1002"
        SQL = SQL & "," & RS!Caja
        SQL = SQL & ", 1120"
        SQL = SQL & " ," & SysDate
        SQL = SQL & " , 17 "
        SQL = SQL & " ,0 ) "
        ExecutarSql SQL
        RS.MoveNext
            
            
            
        Loop
        
        
        
        Rem para crear en contenedor  con las null y las otras cambio de estado
        SQL = " SELECT LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.CAJA,"
        SQL = SQL & "   LECTURACOLECTOR.NUMERO_LECTURA , CONTENEDOR.estado,"
        SQL = SQL & "   CONTENEDOR.COD_CLIENTE,CONTENEDOR.NRO_CAJA "
        SQL = SQL & "   FROM    LECTURACOLECTOR LEFT OUTER JOIN "
        SQL = SQL & "   CONTENEDOR ON LECTURACOLECTOR.CAJA = CONTENEDOR.NRO_CAJA "
        SQL = SQL & "   AND LECTURACOLECTOR.CLIENTE = CONTENEDOR.COD_CLIENTE "
        SQL = SQL & "   Where LECTURACOLECTOR.NUMERO_LECTURA = " & Lectura
        SQL = SQL & "   And (LECTURACOLECTOR.Cliente = 1002) "
        RS.Open SQL, strConBasa
        Do While Not RS.EOF
        
            
        
        
        
            RS.MoveNext
        Loop
        



MsgBox "Terminado"


End Sub

Private Sub cmdCambioEstadoPorLectura_Click()
Dim SQL As String
Dim estado As String
Dim RS As New ADODB.Recordset
If Not IsNumeric(Mid(cboestado.Text, 1, 2)) Then
 MsgBox "iNGRESE EL ESTADO"
 Exit Sub
 
End If


SQL = " SELECT     CAJA, CLIENTE, ORDEN, NUMERO_LECTURA"
SQL = SQL & "  From LECTURACOLECTOR"
SQL = SQL & "  Where NUMERO_LECTURA = " & InputBox("Ingrese el numero de lectura")
SQL = SQL & "  ORDER BY ORDEN, CAJA"
RS.Open SQL, ConActiva, 0, 1

Do While Not RS.EOF
    SQL = " UPDATE CONTENEDOR SET "
    SQL = SQL & vbCrLf & " ESTADO = " & Mid(cboestado.Text, 1, 2)
    SQL = SQL & vbCrLf & " WHERE "
    SQL = SQL & vbCrLf & "  (NOT (ESTADO IN (2, 3))) "
    SQL = SQL & " AND COD_CLIENTE = " & RS!Cliente
    SQL = SQL & " AND NRO_CAJA = " & RS!Caja
    ExecutarSql SQL
    RS.MoveNext
Loop


MsgBox "Terminado"
End Sub

Private Sub cmdCambioOsep_Click()
Dim CAJAS As String

Dim concajas As New ADODB.Connection
concajas.Open strConBasa
CAJAS = InputBox("Ingrese las cajas separadas por ,", , 0)
Dim clienteInicial As Integer
Dim clienteFinal As Integer

clienteInicial = 20
clienteFinal = InputBox("Ingrese el cliente destino")

SQL = " Update dbo.CONTENEDOR"
SQL = SQL & " Set COD_CLIENTE = " & clienteFinal
SQL = SQL & " Where COD_CLIENTE = " & clienteInicial
SQL = SQL & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute SQL



SQL = " Update dbo.cajas "
SQL = SQL & "  Set FK_CLIENTE = " & clienteFinal
SQL = SQL & "  WHERE  FK_CLIENTE =  " & clienteInicial
SQL = SQL & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute SQL


SQL = " Update dbo.REFERENCIAS"
SQL = SQL & " Set COD_CLIENTE = " & clienteFinal
SQL = SQL & " Where COD_CLIENTE = " & clienteInicial
SQL = SQL & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute SQL

SQL = "  Update dbo.MOV_CAJAS2 "
SQL = SQL & " Set id_cliente =  " & clienteFinal
SQL = SQL & " Where (Tipo_elemento = 0)"
SQL = SQL & " AND (ELEMENTO IN (" & CAJAS & "))"
SQL = SQL & " AND ID_CLIENTE = " & clienteInicial
concajas.Execute SQL

MsgBox "Terminado"
Exit Sub
salir:

End Sub

Private Sub cmdRecuperarCajas_Click()

'Dim cAJAS As Long
'Dim SQL As String
'
'
' For cAJAS = 400001 To 401000
'            Etiqueta = 110000000000# + cAJAS
'            SQL = " INSERT INTO CAJAS "
'            SQL = SQL & "  (ID_CAJA, NRO_CAJA,  FK_ESTADO "
'            SQL = SQL & "  , FECHA_CREACION_CAJA, FK_USUARIO_CREACION_CAJA, DIGITO_VERIFICADOR , ETIQUETA, ROLLO ) "
'            SQL = SQL & "  VALUES    "
'            SQL = SQL & "( " & cAJAS & "," & cAJAS & ", 4 "
'            SQL = SQL & "," & SysDate & ",99," & DigitoEAN13(Trim(Str(Etiqueta))) & ",'" & Etiqueta & "'," & 0 & ")"
'            ExecutarSql SQL
'            Next


Dim rsLectura As New ADODB.Recordset
    Dim SQL As String
    Dim rsContenedor As New ADODB.Recordset
    Dim rsControlContenedor As New ADODB.Recordset
    Dim sqlControlContenedor As String
    Dim rsControlCajas As New ADODB.Recordset
    Dim rsCajas As New ADODB.Recordset
    Dim sqlControlCajas As String
    Dim Sqlc As String
    Dim con As New ADODB.Connection

    con.Open strConBasa

        SQL = " SELECT LECTURACOLECTOR.ID, LECTURACOLECTOR.NUMERO_LECTURA, LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE,"
        SQL = SQL & vbCrLf & " LECTURACOLECTOR.ORDEN, CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL, CONTENEDOR.VERTICAL, CONTENEDOR.ESTADO,"
        SQL = SQL & vbCrLf & " CONTENEDOR.COD_CLIENTE , CONTENEDOR.NRO_CAJA"
        SQL = SQL & vbCrLf & " FROM         LECTURACOLECTOR LEFT OUTER JOIN"
        SQL = SQL & vbCrLf & " CONTENEDOR ON LECTURACOLECTOR.CAJA = CONTENEDOR.NRO_CAJA AND LECTURACOLECTOR.CLIENTE = CONTENEDOR.COD_CLIENTE"
        SQL = SQL & vbCrLf & " Where (LECTURACOLECTOR.NUMERO_LECTURA = " & InputBox("Ingrese el numero de lectura", "Lectura", 0) & " ) "
        SQL = SQL & vbCrLf & " And (LECTURACOLECTOR.Cliente < 9000) "
        SQL = SQL & vbCrLf & "  And (CONTENEDOR.Estanteria Is Null) "
        SQL = SQL & vbCrLf & "  ORDER BY LECTURACOLECTOR.CAJA "
        rsLectura.Open SQL, strConBasa

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

        SQL = "  SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_CONTENEDOR, FK_ESTADO"
        SQL = SQL & vbCrLf & "   From basasql.dbo.Cajas"
        SQL = SQL & vbCrLf & "  WHERE     (ID_CAJA BETWEEN 737161 AND 750924) AND (FK_CLIENTE IS NULL)"
        SQL = SQL & vbCrLf & "   ORDER BY ID_CAJA"

        rsCajas.Open SQL, strConBasa


        Do While Not rsLectura.EOF

            sqlControlCajas = " SELECT     ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_ESTADO"
            sqlControlCajas = sqlControlCajas & " From basasql.dbo.Cajas"
            sqlControlCajas = sqlControlCajas & " Where FK_CLIENTE = " & rsLectura!Cliente
            sqlControlCajas = sqlControlCajas & " And NRO_CAJA = " & rsLectura!Caja
            Set rsControlCajas = New ADODB.Recordset
            rsControlCajas.Open sqlControlCajas, strConBasa
            If rsControlCajas.EOF Then
                SQL = " Update basasql.dbo.Cajas"
                SQL = SQL & vbCrLf & " SET  FK_CLIENTE = " & rsLectura!Cliente
                SQL = SQL & vbCrLf & " , NRO_CAJA = " & rsLectura!Caja
                SQL = SQL & vbCrLf & " , FK_ESTADO = 2"
                SQL = SQL & vbCrLf & " Where ID_CAJA = " & rsCajas!ID_CAJA
                con.Execute SQL
            End If




            sqlControlContenedor = " SELECT     ID_CONTENEDOR, COD_CLIENTE, NRO_CAJA, ESTADO "
            sqlControlContenedor = sqlControlContenedor & " From basasql.dbo.CONTENEDOR "
            sqlControlContenedor = sqlControlContenedor & " Where COD_CLIENTE = " & rsLectura!Cliente
            sqlControlContenedor = sqlControlContenedor & " And NRO_CAJA = " & rsLectura!Caja
            Set rsControlContenedor = New ADODB.Recordset
            rsControlContenedor.Open sqlControlContenedor, strConBasa
            If rsControlContenedor.EOF Then
                SQL = " Update basasql.dbo.CONTENEDOR"
                SQL = SQL & vbCrLf & " SET  COD_CLIENTE =" & rsLectura!Cliente
                SQL = SQL & vbCrLf & " , NRO_CAJA =" & rsLectura!Caja
                SQL = SQL & vbCrLf & " , ESTADO =2 "
                SQL = SQL & vbCrLf & " Where ID_CONTENEDOR = " & rsContenedor!ID_CONTENEDOR
                con.Execute SQL
            End If
            rsContenedor.MoveNext
            rsLectura.MoveNext
            rsCajas.MoveNext
        Loop

    MsgBox "terminado"
    
End Sub

Private Sub cmsReuperarCajasCustodia_Click()

    Dim SQL As String
    Dim rsCajas As New ADODB.Recordset
    Dim rsContenedor As New ADODB.Recordset
    Dim Etiqueta As String
        
        SQL = "  SELECT ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_CONTENEDOR, FK_ESTADO "
        SQL = SQL & vbCrLf & " From basasql.dbo.Cajas "
        SQL = SQL & vbCrLf & " WHERE ID_CAJA  = " & txtCajaCustodia.Text
                
        rsCajas.Open SQL, strConBasa
        
      If rsCajas.EOF Then
            Etiqueta = 110000000000# + txtCajaCustodia.Text
            SQL = " INSERT INTO CAJAS "
            SQL = SQL & "  (ID_CAJA, NRO_CAJA,  FK_ESTADO "
            SQL = SQL & "  , FECHA_CREACION_CAJA, FK_USUARIO_CREACION_CAJA, DIGITO_VERIFICADOR , ETIQUETA, ROLLO ) "
            SQL = SQL & "  VALUES    "
            SQL = SQL & "( " & txtCajaCustodia.Text & "," & txtCajaCustodia.Text & ", 1020 "
            SQL = SQL & "," & SysDate & ",99," & DigitoEAN13(Trim(Str(Etiqueta))) & ",'" & Etiqueta & "'," & 0 & ")"
            ExecutarSql SQL
      Else
        
        If IsNull(rsCajas!FK_CLIENTE) Then
            SQL = " UPDATE    TOP (1) CAJAS"
            SQL = SQL & vbCrLf & " SET  FK_CLIENTE = " & txtClienteCustodia.Text
            SQL = SQL & vbCrLf & ", NRO_CAJA =" & txtCajaCustodia.Text
            SQL = SQL & vbCrLf & ", FK_ESTADO = 1020"
            SQL = SQL & vbCrLf & " Where ID_CAJA = " & txtCajaCustodia.Text
            ExecutarSql SQL
        
        End If
      End If
        
        SQL = " SELECT      ESTADO, COD_CLIENTE, NRO_CAJA"
        SQL = SQL & vbCrLf & " From basasql.dbo.CONTENEDOR"
        SQL = SQL & vbCrLf & " Where COD_CLIENTE = " & txtClienteCustodia.Text
        SQL = SQL & vbCrLf & " And NRO_CAJA = " & txtCajaCustodia.Text
        rsContenedor.Open SQL, strConBasa
        
       If rsContenedor.EOF Then
            If Not IsNull(rsContenedor!COD_CLIENTE) Then
                    SQL = " SELECT  TOP (1) ID_CONTENEDOR "
                    SQL = SQL & vbCrLf & " From basasql.dbo.CONTENEDOR"
                    SQL = SQL & vbCrLf & " WHERE ESTANTERIA BETWEEN 150 AND 160 "
                    SQL = SQL & vbCrLf & " AND (ESTADO = 1)"
                    SQL = SQL & vbCrLf & " AND (COD_CLIENTE IS NULL)"
                    Set rsContenedor = New ADODB.Recordset
                    rsContenedor.Open SQL, strConBasa
                    SQL = " Update basasql.dbo.CONTENEDOR"
                    SQL = SQL & vbCrLf & " SET COD_CLIENTE =" & txtClienteCustodia.Text
                    SQL = SQL & vbCrLf & ", NRO_CAJA =" & txtCajaCustodia.Text
                    SQL = SQL & vbCrLf & ", ESTADO = 2"
                    SQL = SQL & vbCrLf & " Where ID_CONTENEDOR = " & rsContenedor!ID_CONTENEDOR
                    ExecutarSql SQL
             Else
                MsgBox "La Caja ya esta en contenedor"
                Exit Sub
            End If
        End If
        
        

    MsgBox "terminado"
End Sub


Private Sub Command1_Click()
    Dim ID_LEGAJO As String
    Dim SQL As String
    Dim RS As New ADODB.Recordset
    
    Dim ConLegajo As New ADODB.Connection
    
ConLegajo.Open strConBasa
        SQL = " SELECT  LEGAJOS.ID_LEGAJO, LEGAJOS.ID_CLIENTE_LEGAJO, LEGAJOS.NRO_CAJA, LEGAJOS.COD_CLIENTE, LEGAJOS.COD_ESTADO,"
        SQL = SQL & " MOV_CAJAS2.TIPO_ELEMENTO , REARCHIVO_CAJA , MOV_CAJAS2.NRO_REMITO"
        SQL = SQL & " FROM LEGAJOS LEFT OUTER JOIN MOV_CAJAS2 ON LEGAJOS.ID_CLIENTE_LEGAJO = MOV_CAJAS2.ELEMENTO AND LEGAJOS.COD_CLIENTE = MOV_CAJAS2.ID_CLIENTE"
        SQL = SQL & " WHERE LEGAJOS.ID_CLIENTE_LEGAJO BETWEEN " & txtEtiquetaBorrarDesde.Text & " AND " & txtEtiquetaBorrarHasta.Text
        SQL = SQL & " AND LEGAJOS.NRO_CAJA = " & txtEtiquetaBorrarCaja.Text
        SQL = SQL & " AND COD_CLIENTE = " & txtClienteEtiquetaBorrar.Text
        
        If (txtEtiquetaBorrarHasta.Text - txtEtiquetaBorrarDesde.Text) > 100 And (txtEtiquetaBorrarHasta.Text - txtEtiquetaBorrarDesde.Text) < 0 Then
            MsgBox "No se pueden Borrar mas de 100 etiquetas por ves", vbCritical
            Exit Sub
        End If
        
        
        
        RS.Open SQL, ConBasa
        
         If MsgBox("Esta usted seguro de borrar la cantidad de " & (txtEtiquetaBorrarHasta.Text - txtEtiquetaBorrarDesde.Text) & " registros de legajos", vbCritical + vbYesNo) = vbYes Then
            Do While Not RS.EOF
                If IsNull(RS!REARCHIVO_CAJA) Then
                   If IsNull(RS!NRO_REMITO) Then
                        If Cod_Estado <> 2 Then
                            SQL = " UPDATE    LEGAJOS "
                            SQL = SQL & "  SET LETRA_DESDE = NULL, LETRA_HASTA = NULL, NRO_DESDE = NULL, NRO_HASTA = NULL, FECHA_DESDE = NULL, FECHA_HASTA = NULL,"
                            SQL = SQL & " DESCRIPCION = NULL, NRO_CAJA = NULL, COD_CLIENTE = NULL, ID_PERSONAL = NULL, FK_PERSONAL_CREACION = NULL,"
                            SQL = SQL & " FECHA_ACTUALIZACION = NULL, FECHA_CREACION = NULL,COD_ESTADO=NULL, COD_INDICE = NULL, FK_INDICES = NULL"
                            SQL = SQL & " Where    ID_CLIENTE_LEGAJO = " & RS!ID_CLIENTE_LEGAJO
                            SQL = SQL & " AND COD_CLIENTE = " & txtClienteEtiquetaBorrar.Text
                            ConLegajo.Execute SQL
                        Else
                            MsgBox "El estado no es el correcto la etiqueta: " & RS!ID_CLIENTE_LEGAJO & "NO sera Borrada"
                        End If
                    Else
                        MsgBox "Tiene Movimiento de elementos la etiqueta: " & RS!ID_CLIENTE_LEGAJO & "NO sera Borrada"
                    End If
                Else
                    MsgBox "Tiene caja Rearchivo la etiqueta: " & RS!ID_CLIENTE_LEGAJO & "NO sera Borrada"
                End If
                RS.MoveNext
            Loop
        End If
        
        MsgBox "Terminado"
End Sub

Private Sub Command10_Click()
Dim SQL As String
    Dim rsCajas As New ADODB.Recordset
    Dim rsContenedor As New ADODB.Recordset
    Dim Etiqueta As String
        
        SQL = "  SELECT ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_CONTENEDOR, FK_ESTADO "
        SQL = SQL & vbCrLf & " From basasql.dbo.Cajas "
        Rem Sql = Sql & vbCrLf & " WHERE ID_CAJA  = " & txtCajaCustodia.Text
        SQL = SQL & vbCrLf & " WHERE ID_CAJA  IN("
        SQL = SQL & vbCrLf & " 697364,754269,808180,821329,821334,821343,823162,825020,841850,850161,863384,863499,863533,870761,891834,893390,893393,893403,899698,909486,925853,927632,927633,928457,931049,936816,938120,942513,942515,942517,944505,944506,951398,951416,951424,965638)"
        
        rsCajas.Open SQL, strConBasa
        
  
  Do While Not rsCajas.EOF
        
      
                    SQL = " SELECT  TOP (1) ID_CONTENEDOR "
                    SQL = SQL & vbCrLf & " From basasql.dbo.CONTENEDOR"
                    SQL = SQL & vbCrLf & " WHERE ESTANTERIA BETWEEN 150 AND 160 "
                    SQL = SQL & vbCrLf & " AND (ESTADO = 1)"
                    SQL = SQL & vbCrLf & " AND (COD_CLIENTE IS NULL)"
                    Set rsContenedor = New ADODB.Recordset
                    rsContenedor.Open SQL, strConBasa
                    SQL = " Update basasql.dbo.CONTENEDOR"
                    SQL = SQL & vbCrLf & " SET COD_CLIENTE =" & rsCajas!FK_CLIENTE
                    SQL = SQL & vbCrLf & ", NRO_CAJA =" & rsCajas!NRO_CAJA
                    SQL = SQL & vbCrLf & ", ESTADO = 2"
                    SQL = SQL & vbCrLf & " Where ID_CONTENEDOR = " & rsContenedor!ID_CONTENEDOR
                    ExecutarSql SQL
             
        
            rsCajas.MoveNext
        Loop
        

    MsgBox "terminado"
End Sub

Private Sub Command2_Click()
Dim SQL As String

SQL = " Update REQUERIMIENTO Set IDESTADO = " & txtRequeEstado.Text
SQL = SQL & " Where IDREQUERIMIENTO = " & txtRequerimiento.Text
ExecutarSql SQL
MsgBox "Terminado", vbInformation

End Sub

Private Sub Command3_Click()
Dim SQL As String
     
    SQL = "  SELECT CAJAS.ID_CAJA, CAJAS.DIGITO_VERIFICADOR, CAJAS.ROLLO"
    SQL = SQL & " FROM CAJAS INNER JOIN"
    SQL = SQL & " LECTURACOLECTOR ON CAJAS.ID_CAJA = LECTURACOLECTOR.CAJA"
    SQL = SQL & " Where LECTURACOLECTOR.NUMERO_LECTURA =" & InputBox("Ingrese el Numero de lectura", "Lectura", 0)
    SQL = SQL & " ORDER BY LECTURACOLECTOR.ORDEN"
     
    frmReportes.ImprimirReporte PasoReportes + "cajasbasaEtiquetas.rpt", SQL, True

End Sub

Private Sub Command4_Click()
Dim RS As New ADODB.Recordset
Dim SQL As String
Dim Lectura As String
Lectura = InputBox("Ingrese el Numero de lectura", "", 0)


SQL = " SELECT        LECTURACOLECTOR.CAJA, LECTURACOLECTOR.CLIENTE, LECTURACOLECTOR.ORDEN, REMITOS_CUERPO.ANULADO"
 SQL = SQL & " FROM            LECTURACOLECTOR INNER JOIN"
  SQL = SQL & "                        REMITOS_DETALLE ON LECTURACOLECTOR.CAJA = REMITOS_DETALLE.DESDE INNER JOIN"
   SQL = SQL & "                        REMITOS_CUERPO ON LECTURACOLECTOR.CLIENTE = REMITOS_CUERPO.ID_CLIENTE AND"
    SQL = SQL & "                      REMITOS_DETALLE.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO"
 SQL = SQL & " Where LECTURACOLECTOR.NUMERO_LECTURA =" & Lectura
  SQL = SQL & " And (REMITOS_CUERPO.ANULADO <> 1)"
 SQL = SQL & " ORDER BY LECTURACOLECTOR.ORDEN"

RS.Open SQL, ConActiva, 0, 1

If Not RS.EOF Then
    Do While Not RS.EOF
        MsgBox "No estan Anulado los Movimientos de la caja " & RS!Caja
        RS.MoveNext
    Loop
    
    Exit Sub
Else

    
    Set RS = New ADODB.Recordset
    
    
    SQL = " SELECT        NUMERO_LECTURA, CAJA, CLIENTE, ORDEN "
 SQL = SQL & "  From LECTURACOLECTOR "
 SQL = SQL & "  Where NUMERO_LECTURA = " & Lectura
 
 SQL = SQL & "  ORDER BY ORDEN"
    
    RS.Open SQL, strConBasa, 0, 1
    Do While Not RS.EOF
    
        SQL = " Update Cajas SET FK_CLIENTE =null, FK_ESTADO =4"
        SQL = SQL & " Where FK_CLIENTE = " & RS!Cliente
        SQL = SQL & "  And NRO_CAJA = " & RS!Caja
        ExecutarSql SQL
        SQL = " DELETE FROM CONTENEDOR "
        SQL = SQL & " Where COD_CLIENTE = " & RS!Cliente
        SQL = SQL & " And NRO_CAJA = " & RS!Caja

        RS.MoveNext
    Loop
    MsgBox "Terminado"
End If



End Sub

Private Sub Command5_Click()
Dim SQL As String
    SQL = " UPDATE CONTENEDOR SET "
    SQL = SQL & vbCrLf & " ESTADO = " & txtCajaEstado
    SQL = SQL & vbCrLf & " WHERE "
    SQL = SQL & " COD_CLIENTE = " & txtClienteCaja.Text
    SQL = SQL & " AND NRO_CAJA = " & txtCajaCaja.Text

     ExecutarSql SQL
End Sub

Private Sub Command6_Click()
'Update Cajas
'Set FK_CLIENTE = 99
'Where (FK_CLIENTE = 76) And (NRO_CAJA = 812425)
'
'
'Update CONTENEDOR
'Set COD_CLIENTE = 99
'Where (COD_CLIENTE = 76) And (NRO_CAJA = 812425)
'
'
'SELECT     REMITOS_CUERPO.ID_CLIENTE, REMITOS_DETALLE.DESDE
'FROM         REMITOS_CUERPO INNER JOIN
'                      REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO
'Where (REMITOS_CUERPO.id_cliente = 76) And (REMITOS_DETALLE.Desde = 812425)
'
'
'



End Sub

Private Sub Command7_Click()
 Dim SQL As String
    SQL = " UPDATE    LEGAJOS"
    SQL = SQL & "  SET LETRA_DESDE = NULL, LETRA_HASTA = NULL, NRO_DESDE = NULL, NRO_HASTA = NULL, FECHA_DESDE = NULL, FECHA_HASTA = NULL,"
    SQL = SQL & " DESCRIPCION = NULL, NRO_CAJA = NULL, COD_CLIENTE = NULL, ID_PERSONAL = NULL, FK_PERSONAL_CREACION = NULL,"
    SQL = SQL & " FECHA_ACTUALIZACION = NULL, FECHA_CREACION = NULL,COD_ESTADO=NULL, COD_INDICE = NULL, FK_INDICES = NULL"
    SQL = SQL & " Where   (ROLLO <> 4760)  and   NRO_CAJA = " & txtBorrarCajasLegajos.Text
    SQL = SQL & " AND COD_CLIENTE = " & txtBorrarclienteLegajos.Text
    If MsgBox("Esta usted seguro de borrar el registro", vbCritical + vbYesNo) = vbYes Then
        ExecutarSql SQL
    End If
    MsgBox "Terminado"
End Sub

Private Sub Command8_Click()
Dim CAJAS As String

Dim concajas As New ADODB.Connection
concajas.Open strConBasa
CAJAS = txtCambioCaja.Text
Dim clienteInicial As Integer
Dim clienteFinal As Integer

clienteInicial = txtClienteInicial.Text
clienteFinal = txtClienteFinal.Text

SQL = " Update dbo.CONTENEDOR"
SQL = SQL & " Set COD_CLIENTE = " & clienteFinal
SQL = SQL & " Where COD_CLIENTE = " & clienteInicial
SQL = SQL & " AND (NRO_CAJA IN (" & CAJAS & "))"

concajas.Execute SQL



SQL = " Update dbo.cajas "
SQL = SQL & "  Set FK_CLIENTE = " & clienteFinal
SQL = SQL & "  WHERE  FK_CLIENTE =  " & clienteInicial
SQL = SQL & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute SQL


SQL = " Update dbo.REFERENCIAS"
SQL = SQL & " Set COD_CLIENTE = " & clienteFinal
SQL = SQL & " Where COD_CLIENTE = " & clienteInicial
SQL = SQL & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute SQL

SQL = "  Update dbo.MOV_CAJAS2 "
SQL = SQL & " Set id_cliente =  " & clienteFinal
SQL = SQL & " Where (Tipo_elemento = 0)"
SQL = SQL & " AND (ELEMENTO IN (" & CAJAS & "))"
SQL = SQL & " AND ID_CLIENTE = " & clienteInicial
concajas.Execute SQL

MsgBox "Terminado"
Exit Sub
salir:
End Sub

Private Sub Command9_Click()
Dim SQL As String

   SQL = "   SELECT     ID_CAJA, DIGITO_VERIFICADOR, FK_CLIENTE, NRO_CAJA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, ORDEN, NUMERO_LECTURA"
 SQL = SQL & " From basasql.dbo.V_CAJAS_LECTURAS"
 SQL = SQL & " Where NUMERO_LECTURA = " & InputBox("Ingrese el Numero de lectura", "Lectura", 0)
 SQL = SQL & " ORDER BY ORDEN"
    
   
    frmReportes.ImprimirReporte PasoReportes + "cajasbasaEtiquetaspos.rpt", SQL, True
End Sub

Private Sub txt_Change()

End Sub

