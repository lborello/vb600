VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cliente"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
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
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   11745
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   59
      Top             =   0
      Width           =   11415
      Begin VB.TextBox txt_ID_Cliente 
         BackColor       =   &H00C0E0FF&
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
         Left            =   1140
         TabIndex        =   91
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtRazonSocial 
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
         Left            =   3180
         MaxLength       =   500
         TabIndex        =   0
         Top             =   240
         Width           =   7395
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   2280
         TabIndex        =   61
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label18 
         Caption         =   "ID Cliente:"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   855
      End
   End
   Begin TabDlg.SSTab sstCliente 
      Height          =   7875
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   13891
      _Version        =   393216
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
      TabCaption(0)   =   "Datos del Cliente"
      TabPicture(0)   =   "frmCliente.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label28"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtNro_Cuit"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtRuta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkfacturaAutomatica"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAceptarCliente"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Contactos"
      TabPicture(1)   =   "frmCliente.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sstContactos"
      Tab(1).Control(1)=   "cmdAceptar_Contactos"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Impuestos"
      TabPicture(2)   =   "frmCliente.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdAceptarTarifas"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "Frame3"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "Frame1"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdAceptarCliente 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   7800
         TabIndex        =   90
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Frame Frame8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   66
         Top             =   3600
         Width           =   9735
         Begin VB.TextBox txtCANTIDADCAJASUMARESTA 
            Height          =   330
            Left            =   3840
            MaxLength       =   30
            TabIndex        =   68
            Top             =   240
            Width           =   2115
         End
         Begin VB.Label Label9 
            Caption         =   "Cantidad de Suma/Resta Cajas en  Facturas:"
            Height          =   315
            Left            =   120
            TabIndex        =   67
            Top             =   300
            Width           =   3675
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   63
         Top             =   2760
         Width           =   9735
         Begin VB.Label lblCantidadCajasCustodia 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   3000
            TabIndex        =   65
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "Cantidad de Cajas en Custodia:"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   360
            Width           =   2655
         End
      End
      Begin TabDlg.SSTab sstContactos 
         Height          =   6555
         Left            =   -74880
         TabIndex        =   62
         Top             =   480
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   11562
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Buscar Contacto"
         TabPicture(0)   =   "frmCliente.frx":0054
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "grdUsuarioCliente"
         Tab(0).Control(1)=   "txtFiltroUsuarioCliente"
         Tab(0).Control(2)=   "cmdModificarContacto"
         Tab(0).Control(3)=   "cmdAgregarContacto"
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Actualizar Contacto"
         TabPicture(1)   =   "frmCliente.frx":0070
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "ctlIndiceUsuario"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame9"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdAceptarContacto"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Command4"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         Begin VB.CommandButton Command4 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6600
            TabIndex        =   89
            Top             =   6120
            Width           =   1095
         End
         Begin VB.CommandButton cmdAceptarContacto 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   88
            Top             =   6120
            Width           =   1095
         End
         Begin VB.CommandButton cmdAgregarContacto 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   -67080
            TabIndex        =   87
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton cmdModificarContacto 
            Caption         =   "Modificar"
            Height          =   375
            Left            =   -68280
            TabIndex        =   84
            Top             =   480
            Width           =   1095
         End
         Begin VB.Frame Frame9 
            Height          =   3315
            Left            =   120
            TabIndex        =   72
            Top             =   2700
            Width           =   8055
            Begin VB.CheckBox chkEnvioReferencia 
               Caption         =   "Envio de Referencia"
               Height          =   375
               Left            =   5640
               TabIndex        =   83
               Top             =   1680
               Width           =   1935
            End
            Begin VB.TextBox txtApellidoNombre 
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
               Left            =   1800
               TabIndex        =   77
               Top             =   720
               Width           =   5775
            End
            Begin VB.TextBox txtCorreo 
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
               Left            =   1800
               TabIndex        =   76
               Top             =   1200
               Width           =   5775
            End
            Begin VB.TextBox txtCod_Indice 
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
               Left            =   1800
               TabIndex        =   75
               Top             =   1680
               Width           =   3675
            End
            Begin VB.TextBox txtUsuario 
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
               Left            =   1800
               TabIndex        =   74
               Top             =   2640
               Width           =   5535
            End
            Begin VB.TextBox txtTelefono 
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
               Left            =   1800
               TabIndex        =   73
               Top             =   2160
               Width           =   5655
            End
            Begin VB.Label lblID_Contacto 
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
               Height          =   375
               Left            =   1800
               TabIndex        =   86
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label11 
               Caption         =   "ID Contacto:"
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
               Left            =   120
               TabIndex        =   85
               Top             =   360
               Width           =   1155
            End
            Begin VB.Label Label22 
               Caption         =   "Apellido y Nombre: "
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
               Left            =   120
               TabIndex        =   82
               Top             =   840
               Width           =   1995
            End
            Begin VB.Label Label20 
               Caption         =   "Correo:"
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
               Left            =   120
               TabIndex        =   81
               Top             =   1260
               Width           =   795
            End
            Begin VB.Label Label12 
               Caption         =   "Sector"
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
               Left            =   120
               TabIndex        =   80
               Top             =   1740
               Width           =   795
            End
            Begin VB.Label Label10 
               Caption         =   "Usuario"
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
               Left            =   120
               TabIndex        =   79
               Top             =   2760
               Width           =   915
            End
            Begin VB.Label Label8 
               Caption         =   "Telefono:"
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
               Left            =   120
               TabIndex        =   78
               Top             =   2280
               Width           =   915
            End
         End
         Begin VB.TextBox txtFiltroUsuarioCliente 
            Height          =   375
            Left            =   -74880
            TabIndex        =   70
            Top             =   480
            Width           =   6495
         End
         Begin MSDataGridLib.DataGrid grdUsuarioCliente 
            Height          =   5175
            Left            =   -74880
            TabIndex        =   69
            Top             =   960
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   9128
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
         Begin Controles.cltIndice ctlIndiceUsuario 
            Height          =   2115
            Left            =   120
            TabIndex        =   71
            Top             =   600
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   3731
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
      End
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   52
         Top             =   420
         Width           =   11175
         Begin VB.TextBox txtCuit 
            Height          =   375
            Left            =   8220
            TabIndex        =   92
            Top             =   1140
            Width           =   2775
         End
         Begin VB.ComboBox cboProvincia 
            Height          =   345
            ItemData        =   "frmCliente.frx":008C
            Left            =   1080
            List            =   "frmCliente.frx":00A2
            TabIndex        =   4
            Text            =   "Combo1"
            Top             =   1140
            Width           =   5055
         End
         Begin VB.TextBox txtNumero 
            Height          =   375
            Left            =   9600
            TabIndex        =   2
            Top             =   240
            Width           =   1395
         End
         Begin VB.TextBox txtCodigoPostal 
            Height          =   375
            Left            =   8220
            TabIndex        =   6
            Top             =   1560
            Width           =   2775
         End
         Begin VB.TextBox txtTelefonos 
            Height          =   375
            Left            =   1080
            TabIndex        =   5
            Top             =   1560
            Width           =   5055
         End
         Begin VB.TextBox txtLocalidad 
            Height          =   375
            Left            =   1080
            TabIndex        =   3
            Top             =   720
            Width           =   9915
         End
         Begin VB.TextBox txtCalle 
            Height          =   375
            Left            =   1080
            TabIndex        =   1
            Top             =   240
            Width           =   7395
         End
         Begin VB.Label Label29 
            Caption         =   "CUIT:"
            Height          =   375
            Index           =   1
            Left            =   6960
            TabIndex        =   93
            Top             =   1260
            Width           =   1095
         End
         Begin VB.Label Label30 
            Caption         =   "Provincia:"
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label4 
            Caption         =   "Numero"
            Height          =   315
            Left            =   8700
            TabIndex        =   57
            Top             =   300
            Width           =   795
         End
         Begin VB.Label Label29 
            Caption         =   "Codigo Postal"
            Height          =   315
            Index           =   0
            Left            =   6960
            TabIndex        =   56
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Teléfonos"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   1620
            Width           =   795
         End
         Begin VB.Label Localidad 
            Caption         =   "Localidad:"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   780
            Width           =   795
         End
         Begin VB.Label Label5 
            Caption         =   "Domicilio:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   360
            Width           =   795
         End
      End
      Begin VB.CheckBox chkfacturaAutomatica 
         Caption         =   "Factura Automatica"
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
         Left            =   2820
         TabIndex        =   51
         Top             =   2340
         Width           =   1755
      End
      Begin VB.TextBox txtRuta 
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
         Left            =   8160
         TabIndex        =   49
         Top             =   1920
         Width           =   1035
      End
      Begin VB.CommandButton cmdAceptar_Contactos 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   -65280
         TabIndex        =   46
         Top             =   7200
         Width           =   1275
      End
      Begin VB.CommandButton cmdAceptarTarifas 
         Caption         =   "Tarifas"
         Height          =   315
         Left            =   -66780
         TabIndex        =   45
         Top             =   5100
         Width           =   1275
      End
      Begin VB.TextBox txtNro_Cuit 
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
         Left            =   7200
         TabIndex        =   44
         Top             =   2400
         Width           =   2235
      End
      Begin VB.Frame Frame4 
         Caption         =   "Servicios"
         Height          =   1275
         Left            =   -74940
         TabIndex        =   28
         Top             =   3480
         Width           =   9795
         Begin VB.TextBox txtRearchivoPorLote 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            Left            =   4740
            TabIndex        =   42
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtRearchivo_Fisico 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            Left            =   1620
            TabIndex        =   40
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtArchivistaCliente 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            Left            =   8280
            TabIndex        =   36
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txtArchivistaPlanta 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            Left            =   8280
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtImagen 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            Left            =   4740
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtPrecinto 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
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
            Left            =   1620
            TabIndex        =   30
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Rearchivo Por Lote"
            Height          =   375
            Index           =   11
            Left            =   2880
            TabIndex        =   43
            Top             =   660
            Width           =   1635
         End
         Begin VB.Label Label14 
            Caption         =   "Rearchivo Fisico"
            Height          =   375
            Index           =   10
            Left            =   120
            TabIndex        =   41
            Top             =   660
            Width           =   1395
         End
         Begin VB.Label Label14 
            Caption         =   "Archivista en el Cliente"
            Height          =   315
            Index           =   9
            Left            =   5880
            TabIndex        =   35
            Top             =   660
            Width           =   2055
         End
         Begin VB.Label Label14 
            Caption         =   "Archivista en planta"
            Height          =   315
            Index           =   8
            Left            =   5880
            TabIndex        =   34
            Top             =   300
            Width           =   1815
         End
         Begin VB.Label Label14 
            Caption         =   "Precio por Imagen"
            Height          =   375
            Index           =   6
            Left            =   2880
            TabIndex        =   31
            Top             =   300
            Width           =   1635
         End
         Begin VB.Label Label14 
            Caption         =   "Precinto:"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   29
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fletes y Consultas"
         Height          =   855
         Left            =   -74880
         TabIndex        =   21
         Top             =   2460
         Width           =   9735
         Begin VB.TextBox txtFleteUrgente 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   8160
            TabIndex        =   26
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtFleteNormal 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   4380
            TabIndex        =   25
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtConsulta 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1320
            TabIndex        =   23
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Flete Urgente"
            Height          =   315
            Index           =   4
            Left            =   6780
            TabIndex        =   27
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label14 
            Caption         =   "Flete Normal"
            Height          =   255
            Index           =   3
            Left            =   3300
            TabIndex        =   24
            Top             =   360
            Width           =   1275
         End
         Begin VB.Label Label14 
            Caption         =   "Consulta "
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   22
            Top             =   420
            Width           =   1035
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Costos Iniciales"
         Height          =   855
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   9735
         Begin VB.TextBox txtCaja 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1380
            TabIndex        =   19
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtCargaLegajo 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   4320
            TabIndex        =   18
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtReferencia 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   8100
            TabIndex        =   16
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Caja:"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   675
         End
         Begin VB.Label Label14 
            Caption         =   "Carga de Legajos:"
            Height          =   255
            Index           =   7
            Left            =   2760
            TabIndex        =   17
            Top             =   300
            Width           =   1455
         End
         Begin VB.Label Label14 
            Caption         =   "Referencia y Trasvase"
            Height          =   315
            Index           =   0
            Left            =   6120
            TabIndex        =   15
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Canon Mensual"
         Height          =   855
         Left            =   -74880
         TabIndex        =   9
         Top             =   1500
         Width           =   9735
         Begin VB.TextBox txtCanonLegajo 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   5820
            TabIndex        =   47
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtAbonoMinimo 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00;(""$"" #.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   8160
            TabIndex        =   38
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtCanonLibro 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   3600
            TabIndex        =   13
            Top             =   300
            Width           =   735
         End
         Begin VB.TextBox txtCanonCaja 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   """$"" #.##0,00;(""$"" #.##0,00)"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   315
            Left            =   1380
            TabIndex        =   12
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Canon Legajo"
            Height          =   435
            Left            =   4680
            TabIndex        =   48
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label16 
            Caption         =   "Abono Mínimo"
            Height          =   315
            Left            =   6900
            TabIndex        =   37
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Canon Libro"
            Height          =   435
            Left            =   2400
            TabIndex        =   11
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label13 
            Caption         =   "Canon Caja"
            Height          =   315
            Left            =   180
            TabIndex        =   10
            Top             =   360
            Width           =   1035
         End
      End
      Begin VB.Label Label28 
         Caption         =   "Ruta"
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
         Left            =   7320
         TabIndex        =   50
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cuit:"
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
         Left            =   6720
         TabIndex        =   39
         Top             =   2460
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   7560
      Width           =   1395
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   1920
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
            Picture         =   "frmCliente.frx":00E3
            Key             =   "Ver+"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":04DD
            Key             =   "Buscar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":077D
            Key             =   "Ver-"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":0B7B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":0F8F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":1357
            Key             =   "Punto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":146C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":188F
            Key             =   "RotarI"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":1CA0
            Key             =   "Vertical"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":20BF
            Key             =   "Sig"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":24C0
            Key             =   "Ant"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":28BE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":2CD4
            Key             =   "Nuevo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":2D8A
            Key             =   "RotarD"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":3194
            Key             =   "Cargar"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":356F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":3937
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":3A29
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":3E19
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":422C
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":45EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":4689
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":472C
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":4B1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":4F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":52E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":56FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":58A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":5C83
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":5D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":6197
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":65A7
            Key             =   "Fin"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":69D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":6B0A
            Key             =   "Aceptar"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":6EC3
            Key             =   "Control"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":6FAE
            Key             =   "Esp. Fax"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":73E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":7518
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":7929
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":7AA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":7EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":82DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":86B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":8AC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":8EC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":9293
            Key             =   "Anular"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":9661
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":9A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":9E36
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":A26A
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":A6AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":A947
            Key             =   "Casa"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCliente.frx":AD4E
            Key             =   "Bandera"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCliente As ADODB.Recordset
Dim WithEvents RsTarifas As ADODB.Recordset
Attribute RsTarifas.VB_VarHelpID = -1
Dim rsCOntactos As ADODB.Recordset
Public ID_Cliente_Maestro As Integer
Dim rsContactos_Actualizar As ADODB.Recordset

Dim rsUsuarioCliente As ADODB.Recordset

Private Sub cboProvincia_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
       SendKeys vbTab
    End If
End Sub


Private Sub cmdAceptar_Contactos_Click()
'Dim sql As String
'    sql = "  Update FACTURA_CONTACTO"
'    sql = sql & vbCrLf & "  SET  "
'    sql = sql & vbCrLf & "  FACTURA_TIPO_ENTREGA= " & ControldatoString(txtTipo_Entrega_Factura.Text)
'    sql = sql & vbCrLf & " ,FACTURA_NOMBRE = " & ControldatoString(txtFactura_Nombre.Text)
'    sql = sql & vbCrLf & " ,FACTURA_CORREO =" & ControldatoString(txtFactura_Correo.Text)
'    sql = sql & vbCrLf & " ,FACTURA_TELEFONOS =" & ControldatoString(txtFactura_Telefono.Text)
'    sql = sql & vbCrLf & " ,FACTURA_DIRECCION =" & ControldatoString(txtFactura_Direccion.Text)
'    sql = sql & vbCrLf & " ,COBRANZA_NOMBRE =" & ControldatoString(txtCobranza_Nombre.Text)
'    sql = sql & vbCrLf & " ,COBRANZA_CORREO =" & ControldatoString(txtCobranza_Correo.Text)
'    sql = sql & vbCrLf & " ,COBRANZA_TELEFONOS =" & ControldatoString(txtCobranza_Telefono.Text)
'    sql = sql & vbCrLf & " ,COBRANZA_TIPO =" & ControldatoString(txtCobranza_Tipo.Text)
'    sql = sql & vbCrLf & " ,COBRANZA_DIRECCION =" & ControldatoString(txtCobranza_Direccion.Text)
'    sql = sql & vbCrLf & "  Where COD_CLIENTE = " & txtID_Cliente.Text
'    ExecutarSql sql

End Sub

Private Sub cmdAceptarCliente_Click()
    
    Dim id_cliente, CANTIDADCAJASUMARESTA  As Integer
    Dim RAZON_SOCIAL, NUMERO, Localidad, ID_PROVINCIA As String
    Dim CALLE, TELEFONOS, COD_POSTAL, Sql  As String
    Dim Cuit     As String
    
On Error GoTo salir
    
    If txt_ID_Cliente.Text <> "" Then
        id_cliente = txt_ID_Cliente.Text
    
    Else
         MsgBox "El ID No puede ser vacia", vbCritical
         Exit Sub
    End If
    
    If Trim(txtRazonSocial.Text) <> "" Then
        RAZON_SOCIAL = "'" & Trim(txtRazonSocial.Text) & "'"
    Else
        MsgBox "La Razon Social No puede ser vacia", vbCritical
        Exit Sub
    End If
    
    
    If Trim(txtCalle.Text) <> "" Then
            CALLE = "'" & Trim(txtCalle.Text) & "'"
    Else
        CALLE = "Null"
        Rem MsgBox "La calle No puede ser vacia", vbCritical
      Rem   Exit Sub
    End If
    
    If Trim(txtNumero.Text) <> "" Then
        NUMERO = "'" & Trim(txtNumero.Text) & "'"
    Else
        NUMERO = 0
    End If
    
    If Trim(txtLocalidad.Text) <> "" Then
        Localidad = "'" & Trim(txtLocalidad.Text) & "'"
    Else
        Localidad = "Null"
    End If
    
    If Trim(cboProvincia.Text) <> "" Then
        ID_PROVINCIA = "'" & Trim(cboProvincia.Text) & "'"
    Else
        
        MsgBox "La provincia No puede ser vacia", vbCritical
        Exit Sub
    End If
    If Trim(txtCodigoPostal.Text) <> "" Then
        COD_POSTAL = "'" & Trim(txtCodigoPostal.Text) & "'"
    Else
         COD_POSTAL = "NULL"
    End If
    
    If Trim(txtTelefonos.Text) <> "" Then
        TELEFONOS = "'" & Trim(txtTelefonos.Text) & "'"
    Else
        TELEFONOS = "Null"
    End If
    
    If IsNumeric(txtCANTIDADCAJASUMARESTA.Text) Then
        CANTIDADCAJASUMARESTA = txtCANTIDADCAJASUMARESTA.Text
        Else
        CANTIDADCAJASUMARESTA = 0
    End If
    If Trim(txtCuit.Text) <> "" Then
        Cuit = "'" & Trim(txtCuit.Text) & "'"
    Else
        Cuit = "NULL"
    End If
    
    
    
    
    If ID_Cliente_Maestro = 0 Then
        Sql = " INSERT INTO CLIENTES "
        Sql = Sql & "  (ID_CLIENTE, RAZON_SOCIAL, CALLE, NUMERO"
        Sql = Sql & "  , LOCALIDAD, ID_PROVINCIA, TELEFONOS, "
        Sql = Sql & "  COD_POSTAL, NRO_Cuit)"
        Sql = Sql & "  VALUES     "
        Sql = Sql & "  (" & id_cliente & "," & RAZON_SOCIAL & "," & CALLE & "," & NUMERO
        Sql = Sql & "," & Localidad & "," & ID_PROVINCIA & "," & TELEFONOS & ","
        Sql = Sql & COD_POSTAL & "," & Cuit & ")"
        ExecutarSql Sql
        Unload Me
    Else
     
      
        Sql = " Update Clientes"
        Sql = Sql & "   SET "
        Sql = Sql & "   RAZON_SOCIAL =" & RAZON_SOCIAL
        Sql = Sql & "  , CALLE =" & CALLE
        Sql = Sql & "  , NUMERO =" & NUMERO
        Sql = Sql & "  , LOCALIDAD =" & Localidad
        Sql = Sql & "  , ID_PROVINCIA =" & ID_PROVINCIA
        Sql = Sql & "  , TELEFONOS =" & TELEFONOS
        Sql = Sql & "   , COD_POSTAL =" & COD_POSTAL
       Sql = Sql & "   , CANTIDADCAJASUMARESTA= " & CANTIDADCAJASUMARESTA
        Sql = Sql & "   Where ID_CLIENTE = " & id_cliente
        ExecutarSql Sql
        Unload Me
     
     
     
      
    End If
    
    Exit Sub
salir:
    MsgBox "Error en el ingreso"
    
End Sub

Private Sub cmdAceptarContacto_Click()
    
    Dim rsMaxContacto As ADODB.Recordset
    
    If lblID_Contacto.Caption = "" Then
        Set rsMaxContacto = New ADODB.Recordset
        
        rsMaxContacto.Open " SELECT     MAX(ID_CLIENTEUSUARIO) AS MAX From CLIENTEUSUARIO", ConActiva, 0, 1
        rsContactos_Actualizar.Fields("ID_CLIENTEUSUARIO") = rsMaxContacto!Max + 1
        rsContactos_Actualizar.Fields("COD_CLIENTE") = ID_Cliente_Maestro
     End If
    
    
    
    rsContactos_Actualizar.Update
    rsUsuarioCliente.Requery
    sstContactos.TabEnabled(1) = False
    sstContactos.Tab = 0
    
End Sub

Private Sub cmdAceptarTarifas_Click()

 Dim Sql As String
  
'sql = " Update TARIFAS_FACTURA"
'sql = sql & vbCrLf & " SET CANON_CAJA =" & ControlDatoNumericTarifas(txtCanonCaja.Text)
'sql = sql & vbCrLf & " , CANON_LIBRO =" & ControlDatoNumericTarifas(txtCanonLibro.Text)
'sql = sql & vbCrLf & " , CANON_LEGAJO=" & ControlDatoNumericTarifas(txtCanonLegajo.Text)
'sql = sql & vbCrLf & " , CAJA =" & ControlDatoNumericTarifas(txtCaja.Text)
'sql = sql & vbCrLf & " , REFERENCIA =" & ControlDatoNumericTarifas(txtReferencia.Text)
'sql = sql & vbCrLf & " , CARGAR_LEGAJOS =" & ControlDatoNumericTarifas(txtCargaLegajo.Text)
'sql = sql & vbCrLf & " , CONSULTA =" & ControlDatoNumericTarifas(txtConsulta)
'sql = sql & vbCrLf & " , FLETE_NORMAL =" & ControlDatoNumericTarifas(txtFleteNormal.Text)
'sql = sql & vbCrLf & " , FLETE_URGENTE =" & ControlDatoNumericTarifas(txtFleteUrgente.Text)
'sql = sql & vbCrLf & " , PRECINTO =" & ControlDatoNumericTarifas(txtPrecinto.Text)
'sql = sql & vbCrLf & " , HORA_ARCHIVISTA_BASA =" & ControlDatoNumericTarifas(txtArchivistaPlanta.Text)
'sql = sql & vbCrLf & " , HORA_ARCHIVISTA_CLIENTE =" & ControlDatoNumericTarifas(txtArchivistaCliente.Text)
'sql = sql & vbCrLf & " , ABONO_MINIMO =" & ControlDatoNumericTarifas(txtAbonoMinimo.Text)
'sql = sql & vbCrLf & " , IMAGEN =" & ControlDatoNumericTarifas(txtImagen.Text)
'sql = sql & vbCrLf & " , REACHIVO_FISICO =" & ControlDatoNumericTarifas(txtRearchivo_Fisico.Text)
'sql = sql & vbCrLf & " , REARCHIVO_LOTE = " & ControlDatoNumericTarifas(txtRearchivoPorLote.Text)
'sql = sql & vbCrLf & "  Where COD_CLIENTE = " & txtID_Cliente.Text
'ExecutarSql sql

End Sub

Private Sub cmdAgregarContacto_Click()
  Dim Sql As String

    Set rsContactos_Actualizar = New ADODB.Recordset
    rsContactos_Actualizar.CursorLocation = adUseClient

    Sql = " SELECT     ID_CLIENTEUSUARIO,COD_INDICE,USUARIO, APELLIDO_NOMBRE, CORREO, TELEFONOS, COD_CLIENTE, REFERENCIAS"
 Sql = Sql & "  From CLIENTEUSUARIO "
Sql = Sql & "  Where COD_CLIENTE = " & ID_Cliente_Maestro
Sql = Sql & "  And ID_CLIENTEUSUARIO = " & rsUsuarioCliente.Fields.Item(0).value

rsContactos_Actualizar.Open Sql, ConActiva, adOpenDynamic, adLockPessimistic


        Set lblID_Contacto.DataSource = rsContactos_Actualizar.DataSource
        lblID_Contacto.DataField = "ID_CLIENTEUSUARIO"
        
        Set txtApellidoNombre.DataSource = rsContactos_Actualizar.DataSource
        txtApellidoNombre.DataField = "APELLIDO_NOMBRE"
        
        Set txtCorreo.DataSource = rsContactos_Actualizar.DataSource
        txtCorreo.DataField = "CORREO"
        
        Set txtCod_Indice.DataSource = rsContactos_Actualizar.DataSource
        txtCod_Indice.DataField = "COD_INDICE"
        
'        Set txtDocumento.DataSource = rsContactos_Actualizar.DataSource
'        txtDocumento.DataField = "DOCUMENTO"
'
        Set txtTelefono.DataSource = rsContactos_Actualizar.DataSource
        txtTelefono.DataField = "TELEFONOS"
        
        
       
        Set txtUsuario.DataSource = rsContactos_Actualizar.DataSource
        txtUsuario.DataField = "USUARIO"
        sstContactos.TabEnabled(1) = True
        sstContactos.Tab = 1
        rsContactos_Actualizar.AddNew
End Sub

Private Sub cmdInsert_Click()
'    Dim rs As New ADODB.Recordset
'    Dim SQL As String
'    Dim Id As Integer
'        SQL = " SELECT MAX(ID_CLIENTE) as Max From Clientes "
'        rs.Open SQL, strConBasa , 0 ,1
'        Id = rs!max + 1
'        SQL = " INSERT INTO CLIENTES "
'        SQL = SQL & vbCrLf & " (ID_CLIENTE, RAZON_SOCIAL, CALLE"
'        SQL = SQL & vbCrLf & " ,NUMERO, TELEFONOS, NRO_CUIT, Localidad )"
'        SQL = SQL & vbCrLf & " VALUES (" & Id & ",'" & UCase(Trim(txtRazonSocial.Text)) & "','" & Trim(txtCalle.Text) & "'"
'        SQL = SQL & vbCrLf & "," & Trim(txtNumero.Text) & ",'" & Trim(txtTelefonos.Text) & "','" & Trim(txtCuit.Text) & "','" & Trim(txtLocalidad.Text) & "')"
'        ExecutarSql SQL
'        MsgBox "El cliente es el " & Id & vbCrLf & " Recuerde reiniciar el sistema Basa Para que tome los cambios", vbInformation

End Sub



Private Sub DatosCliente()

'Dim Sql As String
'Set rsCliente = New ADODB.Recordset
'rsCliente.CursorLocation = adUseClient
'    Sql = " SELECT ID_CLIENTE, RAZON_SOCIAL, CALLE, NUMERO,"
'    Sql = Sql & vbCrLf & " PISO_DEPTO, COD_POSTAL, LOCALIDAD, ID_PROVINCIA,"
'    Sql = Sql & vbCrLf & "    TELEFONOS ,  NRO_CUIT , "
'    Sql = Sql & vbCrLf & " ID_PROVINCIA, PERIODO_FACTURA ,"
'    Sql = Sql & vbCrLf & " TIPO_FACTURA ,CLIENTE_ADMINISTRACION, DETALLE_FACTURACION, TIPO_ENTREGA,RUTA "
'    Sql = Sql & vbCrLf & "  From Clientes "
'     Sql = Sql & vbCrLf & "  ORDER BY ID_CLIENTE "
'   Rem  Sql = Sql & vbCrLf & "  ORDER BY CLIENTE_ADMINISTRACION "
'    rsCliente.Open Sql,ConActiva, adOpenDynamic, adLockOptimistic
'
'Set txtID_Cliente.DataSource = rsCliente.DataSource
'txtID_Cliente.DataField = "ID_CLIENTE"
'
'Set txtRazonSocial.DataSource = rsCliente.DataSource
'txtRazonSocial.DataField = "RAZON_SOCIAL"
'
'Set txtCalle.DataSource = rsCliente.DataSource
'txtCalle.DataField = "CALLE"
'
'Set txtNumero.DataSource = rsCliente.DataSource
'txtNumero.DataField = "NUMERO"
'
'
''Set txtPiso_Depto.DataSource = rsCliente.DataSource
''txtPiso_Depto.DataField = "PISO_DEPTO"
'
'Set txtLocalidad.DataSource = rsCliente.DataSource
'txtLocalidad.DataField = "Localidad"
'
'
'Set txtLocalidad.DataSource = rsCliente.DataSource
'txtLocalidad.DataField = "Localidad"
'
'
'Set txtLocalidad.DataSource = rsCliente.DataSource
'txtLocalidad.DataField = "Localidad"
'
'
'Set txtTelefonos.DataSource = rsCliente.DataSource
'txtTelefonos.DataField = "Telefonos"
'
'Set txtNro_Cuit.DataSource = rsCliente.DataSource
'txtNro_Cuit.DataField = "NRO_CUIT"
'
'
'
'Rem ID_PROVINCIA , PERIODO_FACTURA, ""
'
'
'Set txtTipo_Entrega.DataSource = rsCliente.DataSource
'txtTipo_Entrega.DataField = "Tipo_Entrega"
'
'
'Set txtPeriodo_Factura.DataSource = rsCliente.DataSource
'txtPeriodo_Factura.DataField = "PERIODO_FACTURA"
'
'
'Set txtTipo_Factura.DataSource = rsCliente.DataSource
'txtTipo_Factura.DataField = "TIPO_FACTURA"
'
'Set txtDetalle_Facturacion.DataSource = rsCliente.DataSource
'txtDetalle_Facturacion.DataField = "Detalle_Facturacion"
'
'
'Set txtRuta.DataSource = rsCliente.DataSource
'txtRuta.DataField = "RUTA"

 End Sub

Private Sub cmdModificarContacto_Click()
    Dim Sql As String

    Set rsContactos_Actualizar = New ADODB.Recordset
    rsContactos_Actualizar.CursorLocation = adUseClient

    Sql = " SELECT     ID_CLIENTEUSUARIO,COD_INDICE,USUARIO, APELLIDO_NOMBRE, CORREO, TELEFONOS, COD_CLIENTE, REFERENCIAS"
 Sql = Sql & "  From CLIENTEUSUARIO "
Sql = Sql & "  Where COD_CLIENTE = " & ID_Cliente_Maestro
Sql = Sql & "  And ID_CLIENTEUSUARIO = " & rsUsuarioCliente.Fields.Item(0).value

rsContactos_Actualizar.Open Sql, ConActiva, adOpenDynamic, adLockPessimistic


        Set lblID_Contacto.DataSource = rsContactos_Actualizar.DataSource
        lblID_Contacto.DataField = "ID_CLIENTEUSUARIO"
        
        Set txtApellidoNombre.DataSource = rsContactos_Actualizar.DataSource
        txtApellidoNombre.DataField = "APELLIDO_NOMBRE"
        
        Set txtCorreo.DataSource = rsContactos_Actualizar.DataSource
        txtCorreo.DataField = "CORREO"
        
        Set txtCod_Indice.DataSource = rsContactos_Actualizar.DataSource
        txtCod_Indice.DataField = "COD_INDICE"
        
'        Set txtDocumento.DataSource = rsContactos_Actualizar.DataSource
'        txtDocumento.DataField = "DOCUMENTO"
'
        Set txtTelefono.DataSource = rsContactos_Actualizar.DataSource
        txtTelefono.DataField = "TELEFONOS"
        
        
       
        Set txtUsuario.DataSource = rsContactos_Actualizar.DataSource
        txtUsuario.DataField = "USUARIO"
        sstContactos.TabEnabled(1) = True
        sstContactos.Tab = 1
        
        
End Sub

Private Sub Command3_Click()
'Dim rs As New ADODB.Recordset
'Dim sql As String
'Dim idFACTURA As Long
'
'sql = " SELECT COD_CLIENTE, FECHA, N_COMP, IMPORTE"
'sql = sql & " From FACTURASPENDIENTES ORDER BY COD_CLI, FECHA"
'
'rs.Open sql, strConBasa , 0 ,1
'idFACTURA = 100
'Do While Not rs.EOF
'sql = " INSERT INTO FACTURAS "
' sql = sql & " (ID_FACTURA, TIPO_COMPROBANTE,"
' sql = sql & "   TIPO_FACTURA, NUMERO_FACTURA,"
'    sql = sql & " COD_CLIENTE, FECHA, ESTADO, MONTO_CON_IVA)"
'sql = sql & " VALUES ("
'sql = sql & idFACTURA & ",'Factura',"
'sql = sql & "'" & Mid(rs!N_COMP, 1, 1) & "'," & Mid(rs!N_COMP, 6) & ","
'sql = sql & rs!COD_CLIENTE & ",'" & rs!FECHA & "', 10,'" & rs!IMPORTE & "')"
'   ExecutarSql sql
'   idFACTURA = idFACTURA + 1
'   rs.MoveNext
'   Loop
'


End Sub

Private Sub Command2_Click()
'Dim rs As New ADODB.Recordset
'Dim sql As String
'
'sql = " SELECT ID_CLIENTE, CLIENTE_ADMINISTRACION"
'sql = sql & " From Clientes"
'sql = sql & "  Where (Not (CLIENTE_ADMINISTRACION Is Null))"
'sql = sql & "  ORDER BY ID_CLIENTE "
'
'rs.Open sql, strConBasa , 0 ,1
'
'Do While Not rs.EOF
'sql = " Update FACTURASPENDIENTES"
'sql = sql & "   Set COD_CLIENTE = " & rs!id_cliente
'sql = sql & "   WHERE COD_CLI = '" & rs!CLIENTE_ADMINISTRACION & "'"
'ExecutarSql sql
'    rs.MoveNext
'Loop
End Sub

Private Sub Command4_Click()

    rsUsuarioCliente.Requery
    sstContactos.TabEnabled(1) = False
    sstContactos.Tab = 0
End Sub

Private Sub ctlIndiceUsuario_DblClick()
 txtCod_Indice.Text = ctlIndiceUsuario.Item_Selecionado
End Sub


Private Sub Form_Load()
    Set rsCliente = New ADODB.Recordset
    
    
    Dim Sql As String
    Set rsUsuarioCliente = New ADODB.Recordset
    rsUsuarioCliente.CursorLocation = adUseClient
    txt_ID_Cliente.Enabled = False
    
    If ID_Cliente_Maestro = 0 Then
        sstCliente.TabEnabled(0) = True
        sstCliente.TabEnabled(1) = False
        sstCliente.TabEnabled(2) = False
        Sql = " SELECT     MAX(ID_CLIENTE) AS MaxCliente"
        Sql = Sql & " From Clientes"
        Sql = Sql & "  Where id_cliente <> 999 "
        rsCliente.Open Sql, ConActiva, 0, 1
        txt_ID_Cliente.Text = rsCliente!MaxCliente + 1
        txt_ID_Cliente.Enabled = True
    
    Else
        
        Sql = " SELECT     ID_CLIENTE, RAZON_SOCIAL, CALLE, NUMERO, LOCALIDAD, ID_PROVINCIA, TELEFONOS, COD_POSTAL , CANTIDADCAJASUMARESTA, NRO_CUIT  "
        Sql = Sql & "  From Clientes"
        Sql = Sql & "  Where ID_CLIENTE = " & ID_Cliente_Maestro
        rsCliente.Open Sql, ConActiva, 0, 1
        txt_ID_Cliente.Text = rsCliente!id_cliente
        txtRazonSocial.Text = rsCliente!RAZON_SOCIAL
        If Not IsNull(rsCliente!CALLE) Then
            txtCalle.Text = rsCliente!CALLE
        End If
        
        If Not IsNull(rsCliente!NUMERO) Then
            txtNumero.Text = rsCliente!NUMERO
        End If
        
        If Not IsNull(rsCliente!Localidad) Then
        txtLocalidad.Text = rsCliente!Localidad
        End If
        
        If Not IsNull(rsCliente!ID_PROVINCIA) Then
            cboProvincia.Text = rsCliente!ID_PROVINCIA
        End If
        If Not IsNull(rsCliente!COD_POSTAL) Then
            txtCodigoPostal.Text = rsCliente!COD_POSTAL
        End If
        If Not IsNull(rsCliente!TELEFONOS) Then
            txtTelefonos.Text = rsCliente!TELEFONOS
        End If
 
        If Not IsNull(rsCliente!CANTIDADCAJASUMARESTA) Then
            txtCANTIDADCAJASUMARESTA.Text = rsCliente!CANTIDADCAJASUMARESTA
        Else
            txtCANTIDADCAJASUMARESTA.Text = 0
        End If
    
    If Not IsNull(rsCliente!NRO_Cuit) Then
        txtCuit.Text = rsCliente!NRO_Cuit
    Else
        txtCuit.Text = ""
    End If
    
    
    End If
    
    
    
    Sql = " SELECT CLIENTEUSUARIO.ID_CLIENTEUSUARIO, CLIENTEUSUARIO.APELLIDO_NOMBRE, CLIENTEUSUARIO.CORREO, CLIENTEUSUARIO.TELEFONOS,"
    Sql = Sql & " INDICES.DESCRIPCION"
    Sql = Sql & " FROM CLIENTEUSUARIO LEFT OUTER JOIN"
    Sql = Sql & " INDICES ON CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE AND CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE"
    Sql = Sql & " Where CLIENTEUSUARIO.COD_CLIENTE = " & ID_Cliente_Maestro
    
    rsUsuarioCliente.Open Sql, ConActiva, 0, 1
    
    Set grdUsuarioCliente.DataSource = rsUsuarioCliente.DataSource
    
    
    ctlIndiceUsuario.Actualizar ID_Cliente_Maestro, Sector, False
    
    sstContactos.TabEnabled(1) = False
    
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
' Select Case Button.Caption
' Case "Atras"
'    rsCliente.MovePrevious
'
' Case "Siguiente"
'    If rsCliente.EOF Then
'        MsgBox "Fin de Archivo"
'    Else
'
'    rsCliente.MoveNext
'    End If
' Case "Buscar"
'    rsCliente.Filter = "RAZON_SOCIAL LIKE '%" & InputBox("Ingrese parte de la Razon Social") & "%'"
' Case "Todos"
'    Set rsCliente = New ADODB.Recordset
'    DatosCliente
'  End Select
'
' Dim sql As String
'sql = " SELECT COD_CLIENTE, CANON_CAJA, CANON_LIBRO, CANON_LEGAJO, CAJA, "
'sql = sql & vbCrLf & " REFERENCIA, CARGAR_LEGAJOS, CONSULTA, "
'sql = sql & vbCrLf & " FLETE_NORMAL, FLETE_URGENTE, PRECINTO, "
'sql = sql & vbCrLf & " HORA_ARCHIVISTA_BASA, HORA_ARCHIVISTA_CLIENTE, "
'sql = sql & vbCrLf & " ABONO_MINIMO, IMAGEN, "
'sql = sql & vbCrLf & " REACHIVO_FISICO , REARCHIVO_LOTE,USUARIO "
'sql = sql & vbCrLf & " From TARIFAS_FACTURA "
'
'If txtID_Cliente.Text = "" Then
'Exit Sub
'End If
'Set RsTarifas = New ADODB.Recordset
'RsTarifas.CursorLocation = adUseClient
'RsTarifas.Open sql,ConActiva, adOpenStatic, adLockReadOnly
'RsTarifas.MoveFirst
'RsTarifas.Find "COD_CLIENTE=" & Trim(txtID_Cliente.Text)
'Set txtCanonCaja.DataSource = RsTarifas.DataSource
'    txtCanonCaja.DataField = "CANON_CAJA"
'
'Set txtCanonLibro.DataSource = RsTarifas.DataSource
'    txtCanonLibro.DataField = "CANON_LIBRO"
'
'Set txtCanonLegajo.DataSource = RsTarifas.DataSource
'    txtCanonLegajo.DataField = "CANON_LEGAJO"
'
'
'Set txtCaja.DataSource = RsTarifas.DataSource
'    txtCaja.DataField = "CAJA"
'
'
'
'Rem REFERENCIA, CARGAR_LEGAJOS, CONSULTA,
'
'
'Set txtReferencia.DataSource = RsTarifas.DataSource
'    txtReferencia.DataField = "REFERENCIA"
'
'
'Set txtCargaLegajo.DataSource = RsTarifas.DataSource
'    txtCargaLegajo.DataField = "CARGAR_LEGAJOS"
'
'Set txtConsulta.DataSource = RsTarifas.DataSource
'    txtConsulta.DataField = "CONSULTA"
'
'Rem FLETE_NORMAL, FLETE_URGENTE, PRECINTO,
'
'
'Set txtFleteNormal.DataSource = RsTarifas.DataSource
'    txtFleteNormal.DataField = "FLETE_NORMAL"
'
'Set txtFleteUrgente.DataSource = RsTarifas.DataSource
'    txtFleteUrgente.DataField = "FLETE_URGENTE"
'
'
'Set txtPrecinto.DataSource = RsTarifas.DataSource
'    txtPrecinto.DataField = "Precinto"
'
'
'Rem HORA_ARCHIVISTA_BASA, HORA_ARCHIVISTA_CLIENTE,
'
'
'Set txtArchivistaPlanta.DataSource = RsTarifas.DataSource
'    txtArchivistaPlanta.DataField = "HORA_ARCHIVISTA_BASA"
'
'Set txtArchivistaCliente.DataSource = RsTarifas.DataSource
'    txtArchivistaCliente.DataField = "HORA_ARCHIVISTA_CLIENTE"
'
'Rem " ABONO_MINIMO , IMAGEN"
'
'Set txtAbonoMinimo.DataSource = RsTarifas.DataSource
'    txtAbonoMinimo.DataField = "ABONO_MINIMO"
'
'Set txtImagen.DataSource = RsTarifas.DataSource
'    txtImagen.DataField = "Imagen"
'Rem  REACHIVO_FISICO , REARCHIVO_LOTE "
'
'
'Set txtRearchivo_Fisico.DataSource = RsTarifas.DataSource
'    txtRearchivo_Fisico.DataField = "REACHIVO_FISICO"
'
'Set txtRearchivoPorLote.DataSource = RsTarifas.DataSource
'    txtRearchivoPorLote.DataField = "REARCHIVO_LOTE"
'
'
'
'    sql = " SELECT COD_CLIENTE, FACTURA_TIPO_ENTREGA,FACTURA_NOMBRE, FACTURA_CORREO, "
'    sql = sql & vbCrLf & "    FACTURA_TELEFONOS, FACTURA_DIRECCION, "
'    sql = sql & vbCrLf & " COBRANZA_NOMBRE, COBRANZA_CORREO, "
'    sql = sql & vbCrLf & " COBRANZA_TELEFONOS, COBRANZA_TIPO, "
'    sql = sql & vbCrLf & " COBRANZA_DIRECCION "
'    sql = sql & vbCrLf & " From FACTURA_CONTACTO "
'    sql = sql & vbCrLf & " WHERE COD_CLIENTE = " & Trim(txtID_Cliente.Text)
'Rem rsContactos.UpdateBatch adAffectAllChapters
'Set rsCOntactos = New ADODB.Recordset
' rsCOntactos.Open sql,ConActiva, adOpenDynamic, adLockOptimistic
'
' Rem FACTURA_NOMBRE, FACTURA_CORREO
'Set txtTipo_Entrega_Factura.DataSource = rsCOntactos.DataSource
'txtTipo_Entrega_Factura.DataField = "FACTURA_TIPO_ENTREGA"
'
' Set txtFactura_Nombre.DataSource = rsCOntactos.DataSource
'    txtFactura_Nombre.DataField = "FACTURA_NOMBRE"
'
'   Set txtFactura_Correo.DataSource = rsCOntactos.DataSource
'    txtFactura_Correo.DataField = "FACTURA_CORREO"
'
'
' Rem  FACTURA_TELEFONOS, FACTURA_DIRECCION,
'
' Set txtFactura_Telefono.DataSource = rsCOntactos.DataSource
'    txtFactura_Telefono.DataField = "FACTURA_TELEFONOS"
'
'   Set txtFactura_Direccion.DataSource = rsCOntactos.DataSource
'    txtFactura_Direccion.DataField = "FACTURA_DIRECCION"
'
' Rem COBRANZA_NOMBRE, COBRANZA_CORREO,
'    Set txtCobranza_Nombre.DataSource = rsCOntactos.DataSource
'    txtCobranza_Nombre.DataField = "COBRANZA_NOMBRE"
'
'    Set txtCobranza_Correo.DataSource = rsCOntactos.DataSource
'    txtCobranza_Correo.DataField = "COBRANZA_CORREO"
'
' Rem " COBRANZA_TELEFONOS, COBRANZA_TIPO, "
'
'    Set txtCobranza_Telefono.DataSource = rsCOntactos.DataSource
'    txtCobranza_Telefono.DataField = "COBRANZA_TELEFONOS"
'
'    Set txtCobranza_Tipo.DataSource = rsCOntactos.DataSource
'    txtCobranza_Tipo.DataField = "COBRANZA_TIPO"
'
'  Rem " COBRANZA_DIRECCION "
'    Set txtCobranza_Direccion.DataSource = rsCOntactos.DataSource
'    txtCobranza_Direccion.DataField = "COBRANZA_DIRECCION"


 
End Sub

Private Sub txtTipo_Factura_LostFocus()
'txtTipo_Factura = UCase(txtTipo_Factura)
End Sub

Public Function ControlDatoNumeric(DATO As String) As String
If Not IsNumeric(DATO) Then
    ControlDatoNumeric = 0
Else
    ControlDatoNumeric = "'" & DATO & "'"
End If


End Function

Public Function ControlDatoNumericTarifas(DATO As String) As String
If Not IsNumeric(DATO) Then
    ControlDatoNumericTarifas = 0
Else
    ControlDatoNumericTarifas = "'" & Replace(DATO, ",", ".") & "'"
End If


End Function

Private Sub Label17_Click()

End Sub

Private Sub txt_ID_Cliente_LostFocus()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
         Sql = " SELECT ID_CLIENTE, RAZON_SOCIAL "
         Sql = Sql & " From Clientes Where id_cliente = " & txt_ID_Cliente.Text
         
         rs.Open Sql, ConActiva, 0, 1
        If Not rs.EOF Then
             MsgBox "En  Numero de cliente ya esta en uso", vbCritical
             txt_ID_Cliente.Text = ""
        End If
   
End Sub

Private Sub txtCalle_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
      SendKeys vbTab
    End If
End Sub


Private Sub txtCodigoPostal_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub


Private Sub txtFiltroUsuarioCliente_Change()
If txtFiltroUsuarioCliente.Text <> "" Then
    rsUsuarioCliente.Filter = " APELLIDO_NOMBRE like '%" & txtFiltroUsuarioCliente.Text & "%'"
    
    
Else
    rsUsuarioCliente.Filter = ""
    rsUsuarioCliente.Requery
End If

End Sub


Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub


Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys vbTab
    End If
End Sub


Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys vbTab
    End If

End Sub

Private Sub txtTelefonos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub


