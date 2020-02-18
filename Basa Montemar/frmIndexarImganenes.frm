VERSION 5.00
Object = "{ED512BE6-6629-4FB4-953D-D0C353847163}#1.0#0"; "ImagXpr7.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmIndexarImganenes 
   Caption         =   "DIGITALIZACION"
   ClientHeight    =   10260
   ClientLeft      =   16365
   ClientTop       =   750
   ClientWidth     =   14565
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10260
   ScaleWidth      =   14565
   Begin TabDlg.SSTab SSTab1 
      Height          =   9675
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   17066
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lotes"
      TabPicture(0)   =   "frmIndexarImganenes.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label13"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblCantidad"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "grdLotes"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ctlCliente"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkIndexsarporID"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TXTnOMBREdiRECTORIO"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdChandon"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdCopiarImagenes"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtPasoImagenesFinal"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdBuscar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "CMDuTIL"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cboCampo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtFiltro"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdExportarExcel"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Command10"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Command4"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtHojaRuta"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdCajaDigitalizada"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtCajaTerminada"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Command7"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cmdCopiarMontemar"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Command9"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtLotesExportar"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdleerLegajos"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "chkControlExpreso"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "chkControlExpresoGuiaSucursal"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Command11"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Command2"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmdExtraerIDLocal"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "fraImagen"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cmdImprimirLote"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cmdImprimirCaja"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "cmdImprimirResumen"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "Index"
      TabPicture(1)   =   "frmIndexarImganenes.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCampos(0)"
      Tab(1).Control(1)=   "fraCampos(1)"
      Tab(1).Control(2)=   "fraCampos(2)"
      Tab(1).Control(3)=   "fraCampos(4)"
      Tab(1).Control(4)=   "fraCampos(5)"
      Tab(1).Control(5)=   "fraCampos(3)"
      Tab(1).Control(6)=   "chkCopiarLetra_Numero"
      Tab(1).Control(7)=   "cboOrden"
      Tab(1).Control(8)=   "cmdLoteTerminado"
      Tab(1).Control(9)=   "Command1"
      Tab(1).Control(10)=   "ctlPersonalIndexacion"
      Tab(1).Control(11)=   "ctlVerImagenes1"
      Tab(1).Control(12)=   "grdIndexarImagenes"
      Tab(1).Control(13)=   "Label2"
      Tab(1).Control(14)=   "Label3"
      Tab(1).Control(15)=   "lblLote"
      Tab(1).Control(16)=   "Label4"
      Tab(1).Control(17)=   "lblCliente"
      Tab(1).Control(18)=   "Label5"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmIndexarImganenes.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "GRDHI"
      Tab(2).Control(3)=   "TXTLEGAJOHILE"
      Tab(2).Control(4)=   "CMDbUSCARhILEBRAND"
      Tab(2).Control(5)=   "XTXTCAJAHI"
      Tab(2).Control(6)=   "cmdCopiarexcelhile"
      Tab(2).Control(7)=   "Command8"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmIndexarImganenes.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Command53"
      Tab(3).Control(1)=   "Command52"
      Tab(3).Control(2)=   "Command51"
      Tab(3).Control(3)=   "Command50"
      Tab(3).Control(4)=   "Command49"
      Tab(3).Control(5)=   "Command48"
      Tab(3).Control(6)=   "cmdDuplicadosLibros"
      Tab(3).Control(7)=   "Command47"
      Tab(3).Control(8)=   "Command46"
      Tab(3).Control(9)=   "Command45"
      Tab(3).Control(10)=   "Command44"
      Tab(3).Control(11)=   "Command43"
      Tab(3).Control(12)=   "Command42"
      Tab(3).Control(13)=   "Command41"
      Tab(3).Control(14)=   "Command40"
      Tab(3).Control(15)=   "Command39"
      Tab(3).Control(16)=   "Command38"
      Tab(3).Control(17)=   "Command37"
      Tab(3).Control(18)=   "Command36"
      Tab(3).Control(19)=   "Command35"
      Tab(3).Control(20)=   "Command34"
      Tab(3).Control(21)=   "cmdGodoyCruzCatastroExportarPDF"
      Tab(3).Control(22)=   "Command33"
      Tab(3).Control(23)=   "Command30"
      Tab(3).Control(24)=   "Command29(0)"
      Tab(3).Control(25)=   "Command28"
      Tab(3).Control(26)=   "Command27"
      Tab(3).Control(27)=   "Command26"
      Tab(3).Control(28)=   "Command25"
      Tab(3).Control(29)=   "Command24"
      Tab(3).Control(30)=   "Command23"
      Tab(3).Control(31)=   "Command22"
      Tab(3).Control(32)=   "Command21"
      Tab(3).Control(33)=   "Command20"
      Tab(3).Control(34)=   "Command19"
      Tab(3).Control(35)=   "cmdGodoyCruzCatastroExportartif"
      Tab(3).Control(36)=   "Command18"
      Tab(3).Control(37)=   "Command17"
      Tab(3).Control(38)=   "Command16"
      Tab(3).Control(39)=   "cmdGodoyCruzCatastro"
      Tab(3).Control(40)=   "Command15"
      Tab(3).Control(41)=   "cmdTurismo"
      Tab(3).Control(42)=   "Command14"
      Tab(3).Control(43)=   "txtMesAño"
      Tab(3).Control(44)=   "ctlClienteContar"
      Tab(3).Control(45)=   "txtCajaContar"
      Tab(3).Control(46)=   "cboPasoContar"
      Tab(3).Control(47)=   "Command13"
      Tab(3).Control(48)=   "ImagXpress1"
      Tab(3).Control(49)=   "txtPasoFinalDamsu"
      Tab(3).Control(50)=   "Text1txtPasoDamsu"
      Tab(3).Control(51)=   "Command12"
      Tab(3).Control(52)=   "Label15(2)"
      Tab(3).Control(53)=   "Label15(1)"
      Tab(3).Control(54)=   "Label15(0)"
      Tab(3).Control(55)=   "Label14"
      Tab(3).ControlCount=   56
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmIndexarImganenes.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Command32"
      Tab(4).Control(1)=   "Text1"
      Tab(4).Control(2)=   "Command31"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "frmIndexarImganenes.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdGodoyCruzCajasPersonal"
      Tab(5).ControlCount=   1
      Begin VB.CommandButton Command53 
         Caption         =   "Command53"
         Height          =   495
         Left            =   -69720
         TabIndex        =   149
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton Command52 
         Caption         =   "Command52"
         Height          =   615
         Left            =   -69720
         TabIndex        =   148
         Top             =   8880
         Width           =   1935
      End
      Begin VB.CommandButton Command51 
         Caption         =   "centro Card  Vieja"
         Height          =   615
         Left            =   -72240
         TabIndex        =   147
         Top             =   8760
         Width           =   1935
      End
      Begin VB.CommandButton Command50 
         Caption         =   "Command50"
         Height          =   615
         Left            =   -63840
         TabIndex        =   146
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command49 
         Caption         =   "ExportarlaCaja"
         Height          =   495
         Left            =   -67320
         TabIndex        =   145
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton Command48 
         Caption         =   "Remito Fisico"
         Height          =   495
         Left            =   -63840
         TabIndex        =   144
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdDuplicadosLibros 
         Caption         =   "DuplicadosLibloes"
         Height          =   495
         Left            =   -65160
         TabIndex        =   143
         Top             =   8160
         Width           =   1815
      End
      Begin VB.CommandButton Command47 
         Caption         =   "Command47"
         Height          =   495
         Left            =   -66840
         TabIndex        =   142
         Top             =   8160
         Width           =   1455
      End
      Begin VB.CommandButton cmdGodoyCruzCajasPersonal 
         Caption         =   "Godoy Cruz Cajas Personal"
         Height          =   495
         Left            =   -74760
         TabIndex        =   141
         Top             =   960
         Width           =   2415
      End
      Begin VB.CommandButton Command46 
         Caption         =   "Command46"
         Height          =   495
         Left            =   -63360
         TabIndex        =   140
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton Command45 
         Caption         =   "Osde"
         Height          =   495
         Left            =   -65400
         TabIndex        =   139
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Command44 
         BackColor       =   &H0080C0FF&
         Caption         =   "Command44"
         Height          =   375
         Left            =   -69000
         TabIndex        =   138
         Top             =   8280
         Width           =   1335
      End
      Begin VB.CommandButton Command43 
         Caption         =   "fichas celestes catastro"
         Height          =   495
         Left            =   -72120
         TabIndex        =   137
         Top             =   8160
         Width           =   2295
      End
      Begin VB.CommandButton cmdImprimirResumen 
         Caption         =   "Imprimir resumen"
         Height          =   435
         Left            =   10800
         TabIndex        =   136
         Top             =   1560
         Width           =   1815
      End
      Begin VB.CommandButton cmdImprimirCaja 
         Caption         =   "Imprimir Caja"
         Height          =   435
         Left            =   9120
         TabIndex        =   135
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command42 
         Caption         =   "Command42"
         Height          =   495
         Left            =   -74280
         TabIndex        =   134
         Top             =   8160
         Width           =   1215
      End
      Begin VB.CommandButton Command41 
         BackColor       =   &H008080FF&
         Caption         =   "Command41"
         Height          =   675
         Left            =   -64260
         MaskColor       =   &H000080FF&
         TabIndex        =   133
         Top             =   7440
         Width           =   915
      End
      Begin VB.CommandButton cmdImprimirLote 
         Caption         =   "Imprimir Lote"
         Height          =   435
         Left            =   7560
         TabIndex        =   132
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton Command40 
         Caption         =   "Command40"
         Height          =   735
         Left            =   -67080
         TabIndex        =   131
         Top             =   7380
         Width           =   2295
      End
      Begin VB.Frame fraImagen 
         Caption         =   "Imagen"
         Height          =   675
         Left            =   10320
         TabIndex        =   128
         Top             =   720
         Width           =   2055
         Begin VB.OptionButton optImagenServer 
            Caption         =   "Server"
            Height          =   255
            Left            =   1080
            TabIndex        =   130
            Top             =   300
            Width           =   915
         End
         Begin VB.OptionButton optImagenLocal 
            Caption         =   "Local"
            Height          =   255
            Left            =   180
            TabIndex        =   129
            Top             =   300
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdExtraerIDLocal 
         Caption         =   "ExtraerLocal"
         Height          =   375
         Left            =   10080
         TabIndex        =   127
         Top             =   2640
         Width           =   1635
      End
      Begin VB.CommandButton Command39 
         Caption         =   "EXPORTAR FICHAS CATASTRO"
         Height          =   555
         Left            =   -72360
         TabIndex        =   126
         Top             =   7500
         Width           =   2475
      End
      Begin VB.CommandButton Command38 
         Caption         =   "LEGAJOS GODOY CRUZ"
         Height          =   615
         Left            =   -69480
         TabIndex        =   125
         Top             =   7440
         Width           =   1875
      End
      Begin VB.CommandButton Command37 
         Caption         =   "Command37"
         Height          =   555
         Left            =   -65760
         TabIndex        =   124
         Top             =   3300
         Width           =   1515
      End
      Begin VB.CommandButton Command36 
         Caption         =   "FICHAS CATASTRALES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -74580
         TabIndex        =   123
         Top             =   7440
         Width           =   1815
      End
      Begin VB.CommandButton Command35 
         Caption         =   "Command35"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -66540
         TabIndex        =   122
         Top             =   6420
         Width           =   975
      End
      Begin VB.CommandButton Command34 
         Caption         =   "Command34"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -68340
         TabIndex        =   121
         Top             =   6540
         Width           =   1275
      End
      Begin VB.CommandButton cmdGodoyCruzCatastroExportarPDF 
         Caption         =   "Catastro Exportar pdf"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -65460
         TabIndex        =   120
         Top             =   4800
         Width           =   2595
      End
      Begin VB.CommandButton Command33 
         Caption         =   "Command33"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -72000
         TabIndex        =   119
         Top             =   4020
         Width           =   1455
      End
      Begin VB.CommandButton Command32 
         Caption         =   "Command32"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74700
         TabIndex        =   118
         Top             =   1860
         Width           =   1995
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -72660
         TabIndex        =   117
         Text            =   "Text1"
         Top             =   1080
         Width           =   9075
      End
      Begin VB.CommandButton Command31 
         Caption         =   "Command31"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74640
         TabIndex        =   116
         Top             =   1080
         Width           =   1755
      End
      Begin VB.CommandButton Command30 
         Caption         =   "Command30"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -67680
         TabIndex        =   115
         Top             =   3360
         Width           =   1695
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Command29"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   -69600
         TabIndex        =   114
         Top             =   3300
         Width           =   1695
      End
      Begin VB.CommandButton Command28 
         Caption         =   "medife"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -73260
         TabIndex        =   113
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command27 
         Caption         =   "Command27"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -75000
         TabIndex        =   112
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command26 
         Caption         =   "Command26"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -71520
         TabIndex        =   111
         Top             =   3300
         Width           =   1695
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Command25"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74580
         TabIndex        =   110
         Top             =   6660
         Width           =   1695
      End
      Begin VB.CommandButton Command24 
         Caption         =   "directorios muni godo cruz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -65340
         TabIndex        =   109
         Top             =   6480
         Width           =   2595
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Command23"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -70320
         TabIndex        =   108
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Command22"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -72360
         TabIndex        =   107
         Top             =   6600
         Width           =   1455
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Command21"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   -64860
         TabIndex        =   106
         Top             =   5520
         Width           =   1755
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Command20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -66360
         TabIndex        =   105
         Top             =   5580
         Width           =   1395
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Command19"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -68460
         TabIndex        =   104
         Top             =   5460
         Width           =   1695
      End
      Begin VB.CommandButton cmdGodoyCruzCatastroExportartif 
         Caption         =   "Catastro Exportar tif"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68100
         TabIndex        =   103
         Top             =   4800
         Width           =   2595
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Command18"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   -70320
         TabIndex        =   102
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Command17"
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
         Left            =   -72180
         TabIndex        =   101
         Top             =   5520
         Width           =   1635
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Command16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -72120
         TabIndex        =   100
         Top             =   4740
         Width           =   1695
      End
      Begin VB.CommandButton cmdGodoyCruzCatastro 
         BackColor       =   &H8000000D&
         Caption         =   "GodoyCruzCatastro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -68520
         TabIndex        =   99
         Top             =   3960
         Width           =   3015
      End
      Begin VB.CommandButton Command15 
         Caption         =   "exportar turismo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -69900
         TabIndex        =   98
         Top             =   4740
         Width           =   1755
      End
      Begin VB.CommandButton cmdTurismo 
         Caption         =   "Turismo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   -70440
         TabIndex        =   97
         Top             =   4020
         Width           =   1875
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Command14"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -64800
         TabIndex        =   96
         Top             =   1620
         Width           =   1035
      End
      Begin VB.TextBox txtMesAño 
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
         Left            =   -66660
         TabIndex        =   94
         Text            =   "0"
         Top             =   2640
         Width           =   1155
      End
      Begin Controles.cltGenerico ctlClienteContar 
         Height          =   315
         Left            =   -72060
         TabIndex        =   91
         Top             =   2640
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
      End
      Begin VB.TextBox txtCajaContar 
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
         Left            =   -74160
         TabIndex        =   89
         Text            =   "0"
         Top             =   2580
         Width           =   1215
      End
      Begin VB.ComboBox cboPasoContar 
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
         ItemData        =   "frmIndexarImganenes.frx":00A8
         Left            =   -74160
         List            =   "frmIndexarImganenes.frx":00AA
         TabIndex        =   88
         Text            =   "Combo1"
         Top             =   2220
         Width           =   8715
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Contar Imagenes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -64740
         TabIndex        =   87
         Top             =   2220
         Width           =   1395
      End
      Begin ImagXpr7Ctl.ImagXpress ImagXpress1 
         Height          =   2115
         Left            =   -74460
         TabIndex        =   86
         Top             =   4320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3731
         ErrStr          =   "2917BAFF9E86EAF1B61D37759B64E559"
         ErrCode         =   1057980002
         ErrInfo         =   620834119
         Persistence     =   -1  'True
         _cx             =   3625
         _cy             =   3731
         AutoSize        =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SaveTransparencyColor=   0
         OLEDropMode     =   0
         SaveTIFFCompression=   0
         SaveTransparent =   0
         SaveJPEGProgressive=   0   'False
         SaveJPEGGrayscale=   0   'False
         SaveJPEGLumFactor=   10
         SaveJPEGChromFactor=   10
         SaveJPEGSubSampling=   2
         ViewAntialias   =   -1  'True
         BorderType      =   1
         ViewDithered    =   -1  'True
         AlignH          =   1
         AlignV          =   1
         LoadRotated     =   0
         JPEGEnhDecomp   =   -1  'True
         WMFConvert      =   0   'False
         ProcessImageID  =   1
         OwnDIB          =   -1  'True
         FileTimeout     =   10000
         AsyncPriority   =   0
         LZWPassword     =   ""
         ViewUpdate      =   -1  'True
         TWAINProductName=   ""
         TWAINProductFamily=   ""
         TWAINManufacturer=   ""
         TWAINVersionInfo=   ""
         ViewProgressive =   0   'False
         SaveTIFFByteOrder=   0
         FTPUserName     =   ""
         FTPPassword     =   ""
         ProxyServer     =   ""
         SaveEXIFThumbnailSize=   0
         SaveLJPPrediction=   1
         PDFSwapBlackandWhite=   0   'False
         SaveTIFFRowsPerStrip=   0
         TIFFIFDOffset   =   0
         ViewGrayMode    =   0
         SaveWSQQuant    =   1
         DisplayError    =   0   'False
         EvalMode        =   0
      End
      Begin VB.TextBox txtPasoFinalDamsu 
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
         Left            =   -74700
         TabIndex        =   85
         Text            =   "I:\122-DAMSU\CambioNombre"
         Top             =   1680
         Width           =   6495
      End
      Begin VB.TextBox Text1txtPasoDamsu 
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
         Left            =   -74700
         TabIndex        =   84
         Text            =   "I:\122-DAMSU\DIGITALIZADAS\1057730"
         Top             =   1200
         Width           =   6495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Command12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -67560
         TabIndex        =   83
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pasar Chandor"
         Height          =   375
         Left            =   4560
         TabIndex        =   82
         Top             =   8040
         Width           =   1575
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Command11"
         Height          =   375
         Left            =   1800
         TabIndex        =   81
         Top             =   8880
         Width           =   1215
      End
      Begin VB.CheckBox chkControlExpresoGuiaSucursal 
         Caption         =   "Control Expreso Guia"
         Height          =   375
         Left            =   9840
         TabIndex        =   80
         Top             =   8400
         Width           =   2055
      End
      Begin VB.CheckBox chkControlExpreso 
         Caption         =   "Control Expreso Destino"
         Height          =   375
         Left            =   9840
         TabIndex        =   79
         Top             =   7920
         Width           =   2535
      End
      Begin VB.CommandButton cmdleerLegajos 
         Caption         =   "Completar con legajos"
         Height          =   315
         Left            =   6000
         TabIndex        =   78
         Top             =   8400
         Width           =   1935
      End
      Begin VB.TextBox txtLotesExportar 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   75
         Top             =   1620
         Width           =   5955
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Command9"
         Height          =   375
         Left            =   5880
         TabIndex        =   73
         Top             =   8880
         Width           =   1275
      End
      Begin VB.CommandButton cmdCopiarMontemar 
         Caption         =   "Montemar"
         Height          =   375
         Left            =   3240
         TabIndex        =   72
         Top             =   8040
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Directorio"
         Height          =   375
         Left            =   8760
         TabIndex        =   71
         Top             =   8880
         Width           =   1215
      End
      Begin VB.TextBox txtCajaTerminada 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   69
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdCajaDigitalizada 
         Caption         =   "Caja Terminada"
         Height          =   315
         Left            =   6240
         TabIndex        =   68
         Top             =   7980
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -74580
         TabIndex        =   66
         Top             =   1200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdCopiarexcelhile 
         Caption         =   "Copiar Excel"
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
         Left            =   -67800
         TabIndex        =   63
         Top             =   780
         Width           =   1455
      End
      Begin VB.TextBox XTXTCAJAHI 
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
         Left            =   -74100
         TabIndex        =   62
         Top             =   840
         Width           =   1035
      End
      Begin VB.CommandButton CMDbUSCARhILEBRAND 
         Caption         =   "Buscar"
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
         Left            =   -69480
         TabIndex        =   61
         Top             =   780
         Width           =   1515
      End
      Begin VB.TextBox TXTLEGAJOHILE 
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
         Left            =   -71880
         TabIndex        =   60
         Top             =   840
         Width           =   1875
      End
      Begin MSDataGridLib.DataGrid GRDHI 
         Height          =   7275
         Left            =   -74640
         TabIndex        =   59
         Top             =   1440
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   12832
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
      Begin VB.TextBox txtHojaRuta 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   57
         Top             =   1140
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   4440
         TabIndex        =   53
         Top             =   8880
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Command4"
         Height          =   375
         Left            =   3120
         TabIndex        =   52
         Top             =   8880
         Width           =   1215
      End
      Begin VB.CommandButton cmdExportarExcel 
         Caption         =   "Exportar Excel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7680
         TabIndex        =   51
         Top             =   7440
         Width           =   1695
      End
      Begin VB.TextBox txtFiltro 
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
         Left            =   3540
         TabIndex        =   50
         Top             =   1200
         Width           =   1635
      End
      Begin VB.ComboBox cboCampo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1020
         TabIndex        =   49
         Text            =   "Combo1"
         Top             =   1200
         Width           =   2235
      End
      Begin VB.CommandButton CMDuTIL 
         Caption         =   "UTIL"
         Height          =   375
         Left            =   11160
         TabIndex        =   48
         Top             =   7500
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   5400
         TabIndex        =   36
         Top             =   780
         Width           =   1155
      End
      Begin VB.Frame fraCampos 
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   0
         Left            =   -74760
         OLEDropMode     =   1  'Manual
         TabIndex        =   32
         Top             =   1920
         Width           =   4635
         Begin VB.TextBox txtDato 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   0
            Left            =   1140
            TabIndex        =   33
            Text            =   "Text1"
            Top             =   120
            Width           =   3315
         End
         Begin VB.Label lblTitulo 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   60
            TabIndex        =   34
            Top             =   180
            Width           =   1755
         End
      End
      Begin VB.Frame fraCampos 
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   -73440
         OLEDropMode     =   1  'Manual
         TabIndex        =   29
         Top             =   1920
         Width           =   5535
         Begin VB.TextBox txtDato 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   1980
            TabIndex        =   30
            Text            =   "Text1"
            Top             =   120
            Width           =   3435
         End
         Begin VB.Label lblTitulo 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   180
            Width           =   1755
         End
      End
      Begin VB.Frame fraCampos 
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   2
         Left            =   -71400
         OLEDropMode     =   1  'Manual
         TabIndex        =   26
         Top             =   1920
         Width           =   5535
         Begin VB.TextBox txtDato 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   1980
            TabIndex        =   27
            Text            =   "Text1"
            Top             =   120
            Width           =   3435
         End
         Begin VB.Label lblTitulo 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   180
            Width           =   1755
         End
      End
      Begin VB.Frame fraCampos 
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   4
         Left            =   -68160
         OLEDropMode     =   1  'Manual
         TabIndex        =   23
         Top             =   1920
         Width           =   5415
         Begin VB.TextBox txtDato 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   1860
            TabIndex        =   24
            Text            =   "Text1"
            Top             =   120
            Width           =   3435
         End
         Begin VB.Label lblTitulo 
            BackColor       =   &H80000004&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   60
            TabIndex        =   25
            Top             =   180
            Width           =   1755
         End
      End
      Begin VB.Frame fraCampos 
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   5
         Left            =   -68100
         OLEDropMode     =   1  'Manual
         TabIndex        =   20
         Top             =   1920
         Width           =   5535
         Begin VB.TextBox txtDato 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   5
            Left            =   1980
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   120
            Width           =   3435
         End
         Begin VB.Label lblTitulo 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1755
         End
      End
      Begin VB.Frame fraCampos 
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   3
         Left            =   -69780
         OLEDropMode     =   1  'Manual
         TabIndex        =   17
         Top             =   1920
         Width           =   5535
         Begin VB.TextBox txtDato 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   1980
            TabIndex        =   18
            Text            =   "Text1"
            Top             =   120
            Width           =   3435
         End
         Begin VB.Label lblTitulo 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   180
            Width           =   1755
         End
      End
      Begin VB.CheckBox chkCopiarLetra_Numero 
         Caption         =   "Copiar Letra Numero"
         Height          =   450
         Left            =   -69600
         TabIndex        =   16
         Top             =   960
         Width           =   2115
      End
      Begin VB.ComboBox cboOrden 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmIndexarImganenes.frx":00AC
         Left            =   -73140
         List            =   "frmIndexarImganenes.frx":00AE
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   960
         Width           =   3375
      End
      Begin VB.CommandButton cmdLoteTerminado 
         Caption         =   "Lote Terminado"
         Height          =   315
         Left            =   -65100
         TabIndex        =   13
         Top             =   1620
         Width           =   2595
      End
      Begin VB.TextBox txtPasoImagenesFinal 
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Text            =   "D:\ExportarImagenes\"
         Top             =   7500
         Width           =   4755
      End
      Begin VB.CommandButton cmdCopiarImagenes 
         Caption         =   "Copiar Imagenes"
         Height          =   375
         Left            =   5760
         TabIndex        =   11
         Top             =   7500
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   315
         Left            =   -73560
         TabIndex        =   10
         Top             =   5160
         Width           =   1215
      End
      Begin VB.CommandButton cmdChandon 
         Caption         =   "Crear direc"
         Height          =   375
         Left            =   7320
         TabIndex        =   9
         Top             =   8880
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Command3"
         Height          =   375
         Left            =   480
         TabIndex        =   8
         Top             =   8880
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   180
         TabIndex        =   5
         Top             =   1980
         Width           =   9435
         Begin VB.CheckBox chkInvertirNombre 
            Caption         =   "Invertir Nombre"
            Height          =   375
            Left            =   7680
            TabIndex        =   74
            Top             =   660
            Width           =   1815
         End
         Begin VB.CommandButton cmdUnirLotes 
            Caption         =   "Unir"
            Height          =   315
            Left            =   7800
            TabIndex        =   56
            Top             =   240
            Width           =   675
         End
         Begin VB.TextBox txtUnirNotti 
            Height          =   345
            Left            =   1380
            MultiLine       =   -1  'True
            TabIndex        =   55
            Top             =   600
            Width           =   6135
         End
         Begin VB.TextBox txtUnirFichas 
            Height          =   345
            Left            =   1380
            MultiLine       =   -1  'True
            TabIndex        =   54
            Top             =   180
            Width           =   6135
         End
         Begin VB.Label lblCantidadImagenes 
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
            Left            =   7740
            TabIndex        =   67
            Top             =   420
            Width           =   1575
         End
         Begin VB.Label Label10 
            Caption         =   "Lotes Fichas:"
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "Lotes Notti"
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   660
            Width           =   1215
         End
      End
      Begin VB.TextBox TXTnOMBREdiRECTORIO 
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
         Left            =   1380
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2340
         Width           =   7935
      End
      Begin VB.CheckBox chkIndexsarporID 
         Caption         =   "Indexar por ID"
         Height          =   255
         Left            =   8640
         TabIndex        =   3
         Top             =   780
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Notti"
         Height          =   375
         Left            =   1860
         TabIndex        =   2
         Top             =   8040
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Hospital Notti"
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   8040
         Width           =   1335
      End
      Begin Controles.cltGenerico ctlPersonalIndexacion 
         Height          =   315
         Left            =   -65280
         TabIndex        =   14
         Top             =   960
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
      End
      Begin Controles.ctlVerImagenes ctlVerImagenes1 
         Height          =   4095
         Left            =   -74820
         TabIndex        =   35
         Top             =   5220
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   7223
      End
      Begin Controles.cltGenerico ctlCliente 
         Height          =   375
         Left            =   1020
         TabIndex        =   37
         Top             =   780
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   661
      End
      Begin MSDataGridLib.DataGrid grdLotes 
         Height          =   4095
         Left            =   180
         TabIndex        =   38
         Top             =   3180
         Width           =   12315
         _ExtentX        =   21722
         _ExtentY        =   7223
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   16
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
      Begin MSDataGridLib.DataGrid grdIndexarImagenes 
         Height          =   2295
         Left            =   -74760
         TabIndex        =   39
         Top             =   2700
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   4048
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   16
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
      Begin VB.Label Label15 
         Caption         =   "MesAño"
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
         Index           =   2
         Left            =   -67260
         TabIndex        =   95
         Top             =   2760
         Width           =   1035
      End
      Begin VB.Label Label15 
         Caption         =   "Cliente"
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
         Index           =   1
         Left            =   -72720
         TabIndex        =   93
         Top             =   2700
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Caja"
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   92
         Top             =   2700
         Width           =   1035
      End
      Begin VB.Label Label14 
         Caption         =   "Paso"
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
         Left            =   -74760
         TabIndex        =   90
         Top             =   2220
         Width           =   1335
      End
      Begin VB.Label lblCantidad 
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
         Left            =   10320
         TabIndex        =   77
         Top             =   1140
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Lotes Exportar"
         Height          =   375
         Left            =   240
         TabIndex        =   76
         Top             =   1740
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "Caja:"
         Height          =   255
         Left            =   6780
         TabIndex        =   70
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label Label9 
         Caption         =   "Legajo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72660
         TabIndex        =   65
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Caja:"
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
         Left            =   -74580
         TabIndex        =   64
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Hoja de Ruta"
         Height          =   255
         Left            =   6060
         TabIndex        =   58
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
         Height          =   315
         Left            =   180
         TabIndex        =   47
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Orden de los datos"
         Height          =   315
         Left            =   -74820
         TabIndex        =   46
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Lote:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   45
         Top             =   1620
         Width           =   1035
      End
      Begin VB.Label lblLote 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -73140
         TabIndex        =   44
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   -69420
         TabIndex        =   43
         Top             =   1620
         Width           =   675
      End
      Begin VB.Label lblCliente 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -68640
         TabIndex        =   42
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "Personal de Indexacion"
         Height          =   255
         Left            =   -67320
         TabIndex        =   41
         Top             =   1080
         Width           =   2115
      End
      Begin VB.Label Label8 
         Caption         =   "Paso"
         Height          =   315
         Left            =   300
         TabIndex        =   40
         Top             =   7560
         Width           =   555
      End
   End
   Begin VB.Menu mnuManejoLotes 
      Caption         =   "mnuManejoLotes"
      Begin VB.Menu mnuContarimagenes 
         Caption         =   "Contar Imagenes"
      End
      Begin VB.Menu mnuBorrarLote 
         Caption         =   "Borrar Lote"
      End
      Begin VB.Menu mnuCambiodeNombre 
         Caption         =   "Cambio de Nombre"
      End
      Begin VB.Menu mnuReporteRearchivo 
         Caption         =   "Reporte Rearchivo"
      End
      Begin VB.Menu mnuImportar 
         Caption         =   "Exportar"
         Index           =   22
         Begin VB.Menu mnuMedife 
            Caption         =   "medife"
         End
         Begin VB.Menu MuniGodoyCruzPersonal 
            Caption         =   "Muni Godoy cruz personal"
         End
         Begin VB.Menu mnuPorIdImagen 
            Caption         =   "Por Id de imagen"
         End
         Begin VB.Menu mnuSalud 
            Caption         =   "Salud"
         End
         Begin VB.Menu mnuExportarHilebrand 
            Caption         =   "ExportarHilebrand"
         End
         Begin VB.Menu mnuAriLiquede 
            Caption         =   "AirLiquede"
         End
         Begin VB.Menu mnulacaja 
            Caption         =   "La Caja"
            Begin VB.Menu mnuSinTrack 
               Caption         =   "Sin Track"
            End
            Begin VB.Menu mnuConTrack 
               Caption         =   "Con Track"
            End
            Begin VB.Menu mnuCrearDirectorios 
               Caption         =   "Crear Directorios"
            End
         End
         Begin VB.Menu mnuChandonImport 
            Caption         =   "Chandon"
            Begin VB.Menu mnuExportacion 
               Caption         =   "Exportacion"
            End
            Begin VB.Menu mnuChandonProveedores 
               Caption         =   "Proveedores "
            End
            Begin VB.Menu mnuDDHH 
               Caption         =   "DDHH"
            End
         End
         Begin VB.Menu mnuOsdeDiabeticos 
            Caption         =   "Osde Diabeticos y Psco"
         End
         Begin VB.Menu mnuExpresoLujan1 
            Caption         =   "ExpresoLujan"
            Index           =   23
            Begin VB.Menu mnuExportarEnvio 
               Caption         =   "Exportar Envio"
            End
            Begin VB.Menu mnuControldeimagenes 
               Caption         =   "Control de Imagenes no exportadas"
            End
            Begin VB.Menu mnucontroldeduplicados 
               Caption         =   "Control de duplicados"
            End
            Begin VB.Menu mnuControldecodigos 
               Caption         =   "Control de codigo"
            End
            Begin VB.Menu mnuControlDigitoVerificador 
               Caption         =   "Control de digito Verificador"
            End
         End
         Begin VB.Menu mnuExportCentroCard 
            Caption         =   "CentroCard"
         End
         Begin VB.Menu MNUCOHEN 
            Caption         =   "EXPORTAR Varios"
         End
         Begin VB.Menu mnuZucardiExportar 
            Caption         =   "Zucardi Exportar"
            Begin VB.Menu mnuZucardiFactura 
               Caption         =   "Factura"
            End
            Begin VB.Menu MnuZucardiOrdenes 
               Caption         =   "Ordenes"
            End
         End
         Begin VB.Menu mnuExportarEspañol 
            Caption         =   "Exportar Español"
         End
         Begin VB.Menu mnuExpAndesmar 
            Caption         =   "Andesmar"
            Begin VB.Menu mnuExpAndesmarHojasDeRutas 
               Caption         =   "Hojas de rutas"
            End
         End
         Begin VB.Menu mnuExportaMulti 
            Caption         =   "ExportarCajasMultiAsimple"
         End
      End
      Begin VB.Menu mnuImportar 
         Caption         =   "Importar"
         Index           =   23
         Begin VB.Menu mnuImportMuni 
            Caption         =   "Muni Godoy Cruz"
         End
         Begin VB.Menu mnuBasa 
            Caption         =   "Basa"
            Begin VB.Menu mnuFactura 
               Caption         =   "Factura"
            End
         End
         Begin VB.Menu mnuImportarCentroCard 
            Caption         =   "CentroCard"
         End
         Begin VB.Menu mnuJFH 
            Caption         =   "JFH"
         End
         Begin VB.Menu mnuMontemar 
            Caption         =   "Montemar"
            Begin VB.Menu mnuAcuses9151 
               Caption         =   "Acuses ID 9151 Maxi Acuses"
            End
         End
         Begin VB.Menu mnuAirLiquede 
            Caption         =   "AirLiquede"
         End
         Begin VB.Menu mnuExpresoLujan 
            Caption         =   "Expreso Lujan"
         End
         Begin VB.Menu mnuChandonExport 
            Caption         =   "Chandon"
            Index           =   24
         End
         Begin VB.Menu MnuZucardi 
            Caption         =   "Zucardi"
         End
         Begin VB.Menu mnuEspañol 
            Caption         =   "Español"
         End
         Begin VB.Menu mnuAndesmar 
            Caption         =   "Andesmar"
            Begin VB.Menu mnuhojaderuta 
               Caption         =   "Hojas de Rutas"
            End
         End
      End
      Begin VB.Menu mnuBajarlaprimeraImagen 
         Caption         =   "Baja Primera Imagen"
      End
   End
End
Attribute VB_Name = "frmIndexarImganenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FrameOcupados As Integer
    Dim rsGrilla As ADODB.Recordset
    Dim rsBuscar As ADODB.Recordset
    Dim rsBuscarLote As ADODB.Recordset
    Dim Directorios_Nombres(100) As String
    Dim NRO_Lote As String


Private Sub cboOrden_Click()
 rsGrilla.Sort = cboOrden.Text
    


'    Set grdIndexarImagenes.DataSource = rsGrilla.DataSource
'    grdIndexarImagenes.Rebind
'    grdIndexarImagenes.Refresh
End Sub

Private Sub LaCajaConTrack()
'    Dim Sql As String
'    Dim rsImagenes As New ADODB.Recordset
'    Dim Carpeta As String
'    Dim DocSeparador As MODI.Document
'    Dim DocFichas As MODI.Document
'    Dim DocSave As MODI.Document
'
'    rsBuscar.Requery
'    MousePointer = 11
'    Dim Lotes As String
'
'    Set DocSeparador = New MODI.Document
'    DocSeparador.Create "C:\registro.tif"
'
'
'        Carpeta = InputBox("Carpeta de salida", "", "D:\ExportarImagenes\")
'        Do While Not rsBuscar.EOF
'           Lotes = Lotes & "," & rsBuscar!Lote
'           rsBuscar.MoveNext
'        Loop
'
'        Sql = " SELECT     ID, NRO_DESDE, DIRECTORIO_PASO"
'        Sql = Sql & "  From DOCUMENTOS_DIGITALES "
'        Sql = Sql & "  WHERE FK_DOCUMENTOS_DIGITALES_LOTE IN (" & Mid(Lotes, 2) & ")"
'        Sql = Sql & "  ORDER BY ID"
'
'        Set rsImagenes = New ADODB.Recordset
'        rsImagenes.Open Sql, strConBasa , 0 ,1
'
'
'            FileSystem.MkDir Carpeta
'            Dim i As Integer
'
'            Do While Not rsImagenes.EOF
'                Set DocFichas = New MODI.Document
'                DocFichas.Create PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif"
'               Set DocSave = New MODI.Document
'               DocSave.Create
'
'
'                    For i = 0 To DocFichas.Images.Count - 1
'                     DocSave.Images.Add DocFichas.Images.Item(i), DocFichas.Images.Item(i)
'                    Next
'
'
'                     DocSave.Images.Add DocSeparador.Images.Item(0), DocSave.Images.Item(0)
'
'              If chkInvertirNombre.value = 1 Then
'                    DocSave.SaveAs Carpeta & "\" & rsImagenes!ID & "_" & Trim(rsImagenes!NRO_DESDE) & ".TIF"
'               Else
'                    DocSave.SaveAs Carpeta & "\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
'               End If
'                DocSave.Close
'                rsImagenes.MoveNext
'            Loop
'
'
'MousePointer = 0
'MsgBox "Operacion terminada"
End Sub

Private Sub cmdBajar1Imagen_Click()
'    Dim Sql As String
'    Dim rsImagenes As New ADODB.Recordset
'    Dim docOrigen As MODI.Document
'    Dim docDestino As MODI.Document
'    rsBuscar.Requery
'    MousePointer = 11
'        Do While Not rsBuscar.EOF
'            Sql = "  SELECT ID, COD_CLIENTE, LOTE, COD_ESTADO, DIRECTORIO_PASO "
'            Sql = Sql & " From  DOCUMENTOS_DIGITALES  "
'            Sql = Sql & "  WHERE    COD_CLIENTE = " & rsBuscar!COD_CLIENTE
'            Sql = Sql & "  AND LOTE =  '" & rsBuscar!Lote & "'"
'            Sql = Sql & "  AND COD_ESTADO = " & rsBuscar!Cod_Estado
'            Set rsImagenes = New ADODB.Recordset
'            rsImagenes.Open Sql, strConBasa , 0 ,1
'            Do While Not rsImagenes.EOF
'                Set docOrigen = New MODI.Document
'                docOrigen.Create PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif"
'                Set docDestino = New MODI.Document
'                docDestino.Create
'                docDestino.Images.Add docOrigen.Images(0), docOrigen.Images.Item(0)
'                docDestino.SaveAs txtPasoImagenesFinal.Text & rsImagenes!ID & ".tif"
'                rsImagenes.MoveNext
'                docOrigen.Close
'                docDestino.Close
'            Loop
'       rsBuscar.MoveNext
'Loop
'MousePointer = 0
'MsgBox "Operacion terminada"
End Sub

Private Sub cmdBuscar_Click()
'Dim SQL As String
'Dim CantidadRegistro As Integer
'Set rsBuscar = New ADODB.Recordset
'rsBuscar.CursorLocation = adUseClient
'CantidadRegistro = InputBox("Ingrese la cantidad de registro", "", 100)
'
'ConBasa.CommandTimeout = 300
'SQL = " SELECT    TOP " & CantidadRegistro & " COD_CLIENTE, LOTE, COD_ESTADO, FECHA_INCORPORACION, INDICE,REMITO, count(*) as cantidad"
'SQL = SQL & " FROM DOCUMENTOS_DIGITALES"
'SQL = SQL & "  WHERE COD_CLIENTE = " & ctlCliente.Valor
'    If txtLote.Text <> "" Then
'        SQL = SQL & "  and Lote Like '%" & txtLote.Text & "%'"
'    End If
'
'     If txtEstado.Text <> "" Then
'        SQL = SQL & "  and COD_ESTADO = " & txtEstado.Text
'    End If
'
'    If txtFechaIncorporacion.Text <> "" Then
'       SQL = SQL & "  and  FECHA_INCORPORACION = '" & txtFechaIncorporacion.Text & "'"
'    End If
'
'
'SQL = SQL & " GROUP BY COD_CLIENTE, LOTE, COD_ESTADO, FECHA_INCORPORACION, INDICE,REMITO"
'SQL = SQL & " ORDER BY FECHA_INCORPORACION DESC"
'rsBuscar.Open SQL, strConBasa , 0 ,1
'
'Dim i As Integer
'
'    Set grdLotes.DataSource = rsBuscar.DataSource
'        grdLotes.DataMember = rsBuscar.DataMember
'        grdLotes.Refresh
''        Do While Not rs.EOF
''
''
''            grdLotes2.AddItem i & vbTab & rs!COD_CLIENTE & vbTab & rs!Lote & vbTab & rs!Cod_Estado & vbTab & rs!FECHA_INCORPORACION & vbTab & rs!Indice & vbTab & rs!REMITO & vbTab & rs!CANTIDAD
''             i = i + 1
''            rs.MoveNext
''        Loop
''
'
'
' lblLote.Caption = ""
' lblCliente.Caption = ""
    
    Dim Sql As String
    Dim i As Integer
        Set rsBuscarLote = New ADODB.Recordset
        rsBuscarLote.CursorLocation = adUseClient
        Sql = " SELECT ID_DOCUMENTOS_DIGITALES_LOTE as LOTE , DESCRIPCION, SUB_LOTE as ORDEN1, FK_ESTADO,FK_CAJAS ,  "
        Sql = Sql & " CANTIDAD_IMAGENES , CANTIDAD_ARCHIVOS, FECHA_PREPARACION, "
        Sql = Sql & " FECHA_SCANNER , FECHA_INDEXACION, FECHA_EXPORTACION, TIPO_DOCUMENTO,"
        Sql = Sql & " HOJA_RUTA, FK_LA_CAJA_TOMADOR, FK_INDICES ,  FK_PERSONAL_INDEXACION, FK_PERSONAL_SCANNER, FK_PERSONAL_PREPARACION , remito"
        Sql = Sql & " From DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & " Where FK_CLIENTES = " & ctlCliente.Valor
        If txtFiltro.Text <> "" And cboCampo.Text <> "" Then
            If cboCampo.Text = "LOTE" Then
                Sql = Sql & " and ID_DOCUMENTOS_DIGITALES_LOTE  =" & txtFiltro.Text
            Else
                If cboCampo.ListIndex = -1 Then
                    MsgBox "Ingrese ele campo"
                    Exit Sub
                End If
                If cboCampo.ItemData(cboCampo.ListIndex) = 200 Then
                    Sql = Sql & " and  " & cboCampo.Text & " like '" & txtFiltro.Text & "'"
                Else
                    Sql = Sql & " and  " & cboCampo.Text & " = '" & txtFiltro.Text & "'"
                End If
            End If
        End If
         
         
         If txtHojaRuta.Text <> "" Then
             Sql = Sql & " AND HOJA_RUTA = " & txtHojaRuta.Text
         End If
        Sql = Sql & " order by ID_DOCUMENTOS_DIGITALES_LOTE desc "
        rsBuscarLote.Open Sql, strConBasa, adOpenDynamic, adLockOptimistic
        cboCampo.Clear
        For i = 0 To rsBuscarLote.Fields.Count - 1
            cboCampo.AddItem rsBuscarLote.Fields.Item(i).Name
            cboCampo.ItemData(i) = rsBuscarLote.Fields.Item(i).Type
        Next
        grdLotes.Columns(0).Locked = True
        Set grdLotes.DataSource = rsBuscarLote.DataSource
        grdLotes.DataMember = rsBuscarLote.DataMember
        grdLotes.Refresh
        
        


End Sub



Private Sub CMDbUSCARhILEBRAND_Click()
 Dim Sql As String
 Dim rs As ADODB.Recordset
 Dim Paso As String

Sql = " SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION,"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION , DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"
Sql = Sql & " FROM         DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE ON "
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE     (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 84) AND (DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS =" & XTXTCAJAHI.Text & " ) AND "
Sql = Sql & " (DOCUMENTOS_DIGITALES.LETRA_DESDE LIKE '" & TXTLEGAJOHILE.Text & "')"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"


Set rs = New ADODB.Recordset
Paso = ""
rs.Open Sql, ConActiva, 0, 1

If Not rs.EOF Then
    Paso = rs!Descripcion
End If

Do While Not rs.EOF
    MsgBox "POS " & rs!IMAGEN_ORIGEN
    rs.MoveNext
Loop


If Paso <> "" Then
Sql = " SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION,"
Sql = Sql & " DOCUMENTOS_DIGITALES.PasoOrigen , DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"
Sql = Sql & " FROM         DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE ON "
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE     (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 84) AND (DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS =" & XTXTCAJAHI.Text & " ) AND "
Sql = Sql & " (DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION LIKE '%" & Paso & "%')"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"

rs.Close

Dim RSGG As New ADODB.Recordset
RSGG.CursorLocation = adUseClient
Set RSGG = New ADODB.Recordset
RSGG.Open Sql, ConActiva, adOpenKeyset, adLockReadOnly

 Set GRDHI.DataSource = RSGG.DataSource

End If







End Sub

Private Sub cmdCajaDigitalizada_Click()
        If txtCajaTerminada.Text <> "" And Not IsNull(ctlCliente.Valor) Then
            Dim Sql As String
            Sql = "Update dbo.Cajas "
            Sql = Sql & " Set FK_TIPO_REFERENCIA = 1070"
            Sql = Sql & " Where FK_CLIENTE =" & ctlCliente.Valor
            Sql = Sql & "  And NRO_CAJA = " & txtCajaTerminada.Text
            ExecutarSql Sql
            MsgBox "La caja se registro como DIGITALIZACION TERMINADA", vbInformation
        Else
            MsgBox "Faltan datos"
        End If
End Sub

Private Sub cmdChandon_Click()


'MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
'Do While MyName <> ""   ' Start the loop.
'      ' Use bitwise comparison to make sure MyName is a directory.
'      If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
'         ' Display entry only if it's a directory.
'         MsgBox (MyName)
'      End If
'   MyName = Dir()   ' Get next entry.
'Loop


End Sub

Private Sub cmdCopiarexcelhile_Click()
CopiarDatosGrilla GRDHI
End Sub

Private Sub cmdCopiarImagenes_Click()
    Dim Sql As String
    Dim rsImagenes As New ADODB.Recordset
   
    rsBuscar.Requery
       
        MousePointer = 11
        
        
        
        Do While Not rsBuscar.EOF
        Sql = "  SELECT ID, COD_CLIENTE, LOTE, COD_ESTADO, DIRECTORIO_PASO, Letra_Desde "
        Sql = Sql & " From  DOCUMENTOS_DIGITALES  "
        Sql = Sql & "  WHERE    COD_CLIENTE = " & rsBuscar!COD_CLIENTE
        Sql = Sql & "  AND LOTE =  '" & rsBuscar!lote & "'"
        Sql = Sql & "  AND COD_ESTADO = " & rsBuscar!Cod_Estado
        
'     sql = "    SELECT    ID, COD_CLIENTE, LOTE, COD_ESTADO, DIRECTORIO_PASO, Letra_Desde"
' sql = sql & "  From DOCUMENTOS_DIGITALES"
' sql = sql & "  WHERE (COD_CLIENTE = 147) AND  NRO_desde = 0 "
' sql = sql & "  ORDER BY ID "
        
        
        Set rsImagenes = New ADODB.Recordset
        
        rsImagenes.Open Sql, ConActiva, 0, 1
        
            Do While Not rsImagenes.EOF
               Rem FileSystem.FileCopy PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif", txtPasoImagenesFinal & Trim(rsImagenes!LETRA_DESDE) & "_" & rsImagenes!ID & ".tif"
               FileSystem.FileCopy PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif", txtPasoImagenesFinal & rsImagenes!ID & ".tif"
              rsImagenes.MoveNext
            Loop
       rsBuscar.MoveNext
Loop
MousePointer = 0
MsgBox "Operacion terminada"
End Sub

Private Sub cmdCopiarMontemar_Click()
CopiarMontemar
End Sub

Private Sub cmdDuplicadosLibros_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Cantidad As Integer

Dim NRO_DESDE As Double
Dim IDMAX As Long
Dim FK_DOCUMENTOS_DIGITALES_LOTE As Long
Dim PasoOrigen As String
Dim IMAGEN_ORIGEN As String
Dim FECHA_INCORPORACION As String
Dim Cantidad_Imagenes As Integer
Dim estado As String
Dim DIRECTORIO_PASO As String
Dim i As Integer





Sql = " SELECT DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES.PASOORIGEN, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FECHA_INCORPORACION, DOCUMENTOS_DIGITALES.CANTIDAD_IMAGENES, DOCUMENTOS_DIGITALES.ESTADO,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.DIRECTORIO_PASO"
Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & vbCrLf & " WHERE   (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1156) AND (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10843) AND (CONVERT(nvarchar, "
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE) LIKE '%88888%') AND  LEN(DOCUMENTOS_DIGITALES.NRO_DESDE) = 10 "
Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE"


            rs.Open Sql, ConBasa
            Do While Not rs.EOF
            Cantidad = Mid(rs!NRO_DESDE, 10)
                    For i = 1 To Cantidad
                            If i = 1 Then
                                Sql = "Update DOCUMENTOS_DIGITALES"
                                Sql = Sql & vbCrLf & " SET "
                                Sql = Sql & vbCrLf & " NRO_DESDE = 8888888881"
                                Sql = Sql & vbCrLf & " , ESTADO ='VERIFICAR MANUALMENTE' "
                                Sql = Sql & vbCrLf & " Where ID = " & rs!ID
                                ExecutarSql Sql
                            Else
                                NRO_DESDE = "8888888880" + i
                                IDMAX = MAX_DOCUMENTOS_DIGITALES_2()
                                FK_DOCUMENTOS_DIGITALES_LOTE = rs!FK_DOCUMENTOS_DIGITALES_LOTE
                                PasoOrigen = "'" & Trim(rs!PasoOrigen) & "'"
                                IMAGEN_ORIGEN = "'" & Trim(rs!IMAGEN_ORIGEN) & "'"
                                FECHA_INCORPORACION = "'" & rs!FECHA_INCORPORACION & "'"
                                Cantidad_Imagenes = rs!Cantidad_Imagenes
                                estado = "'VERIFICAR MANUALMENTE'"
                                DIRECTORIO_PASO = BuscarDirectorioPaso(IDMAX)
                                Sql = " INSERT INTO DOCUMENTOS_DIGITALES ( "
                                Sql = Sql & vbCrLf & " NRO_DESDE "
                                Rem SQL = SQL & vbCrLf & " , ID"
                                Sql = Sql & vbCrLf & " , FK_DOCUMENTOS_DIGITALES_LOTE"
                                Sql = Sql & vbCrLf & " , PASOORIGEN"
                                Sql = Sql & vbCrLf & " , IMAGEN_ORIGEN"
                                Sql = Sql & vbCrLf & " , FECHA_INCORPORACION"
                                Sql = Sql & vbCrLf & " , CANTIDAD_IMAGENES"
                                Sql = Sql & vbCrLf & " , ESTADO"
                                Sql = Sql & vbCrLf & " , DIRECTORIO_PASO"
                                Sql = Sql & vbCrLf & " )"
                                Sql = Sql & vbCrLf & " VALUES ( "
                                Sql = Sql & vbCrLf & NRO_DESDE
                                Rem SQL = SQL & vbCrLf & " , " & IDMAX + 1
                                Sql = Sql & vbCrLf & " , " & FK_DOCUMENTOS_DIGITALES_LOTE
                                Sql = Sql & vbCrLf & " , " & PasoOrigen
                                Sql = Sql & vbCrLf & " , " & IMAGEN_ORIGEN
                                Sql = Sql & vbCrLf & " , " & FECHA_INCORPORACION
                                Sql = Sql & vbCrLf & " , " & Cantidad_Imagenes
                                Sql = Sql & vbCrLf & " , " & estado
                                Sql = Sql & vbCrLf & " , '" & Trim(DIRECTORIO_PASO) & "'"
                                Sql = Sql & vbCrLf & " )"
                                ExecutarSql Sql
                                FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF", "\\222.15.19.251\Imagenes\" & DIRECTORIO_PASO & "\" & IDMAX & ".TIF"
                            End If
                    Next
            rs.MoveNext
            Loop



'Dim SQL As String
'    Dim RS As New ADODB.Recordset
'    Dim NombreArchivo  As String
'        SQL = " SELECT     DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_INDICES, DOCUMENTOS_DIGITALES.LETRA_DESDE, "
'        SQL = SQL & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.FK_ID_CATASTRO, "
'        SQL = SQL & " DOCUMENTOS_DIGITALES.LETRA_HASTA , DOCUMENTOS_DIGITALES.Descripcion , DIRECTORIO_PASO "
'        SQL = SQL & " FROM DOCUMENTOS_DIGITALES INNER JOIN "
'        SQL = SQL & " DOCUMENTOS_DIGITALES_LOTE ON "
'        SQL = SQL & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
'        SQL = SQL & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10844) "
'        SQL = SQL & " ORDER BY DOCUMENTOS_DIGITALES.ID "
'        Set RS = New ADODB.Recordset
'        RS.CursorLocation = adUseClient
'        RS.Open SQL, strConBasa, adOpenKeyset, adLockOptimistic
'        Do While Not RS.EOF
'             NombreArchivo = "PADRON " & Format(RS!NRO_DESDE, "0000000") & " ID_" & RS!ID
'            If Dir(PasoImagenes & RS!DIRECTORIO_PASO & "\" & RS!ID & ".tif") <> "" Then
'                FileSystem.FileCopy PasoImagenes & RS!DIRECTORIO_PASO & "\" & RS!ID & ".tif", "D:\FICHAS CELESTES\" & Trim(NombreArchivo) & ".TIF"
'            Else
'                Debug.Print PasoImagenes & RS!DIRECTORIO_PASO & "\" & RS!ID & ".tif"
'            End If
'            RS.MoveNext
'        Loop



End Sub

Private Sub cmdExportarExcel_Click()
CopiarDatosGrilla grdLotes
End Sub

Private Sub cmdExtraerIDLocal_Click()
     
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim Documento As Long
    Dim Nombre As String
    Dim P As Integer
    Dim i As Integer
    Dim Codigo As String
    Dim Archivos As String
    Dim lote As String
    
        MousePointer = 11
        Archivos = ""
        Sql = " SELECT DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER, "
        Sql = Sql & " DOCUMENTOS_DIGITALES.Exportado , DOCUMENTOS_DIGITALES.LETRA_HASTA "
        Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN "
        Sql = Sql & " DOCUMENTOS_DIGITALES ON "
        Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE "
        Sql = Sql & " Where  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE in (" & Mid(txtLotesExportar.Text, 2) & ")"
        Rem  Sql = Sql & " WHERE     (DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE BETWEEN 27073 AND 27171)"
        rs.CursorLocation = adUseClient
        rs.Open Sql, strConBasa
            Do While Not rs.EOF
            lote = rs!ID_DOCUMENTOS_DIGITALES_LOTE
            If Dir("C:\ImagenLocal\" & lote, vbDirectory) = "" Then
             FileSystem.MkDir "C:\ImagenLocal\" & lote
            End If
              If Dir(PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif") <> "" Then
            
                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", "C:\ImagenLocal\" & lote & "\" & rs!ID & ".tif"
            Else
                Archivos = Archivos & vbCrLf & rs!ID & ".tif"
                
            End If
            
                rs.MoveNext
            Loop
            MousePointer = 0
            If Archivos <> "" Then
            
                MsgBox "Archivos no encntrados " & Archivos
            End If
            
            
            
            MsgBox "Terminado"
            





End Sub

Private Sub cmdGodoyCruzCajasPersonal_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Sql = " SELECT  DOCUMENTOS_DIGITALES.COD_CLIENTE, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
Sql = Sql & " DOCUMENTOS_DIGITALES.ID , DOCUMENTOS_DIGITALES.DIRECTORIO_PASO , DOCUMENTOS_DIGITALES.FK_LEGAJO_ETIQUETA "
Sql = Sql & " FROM DOCUMENTOS_DIGITALES INNER JOIN DOCUMENTOS_DIGITALES_LOTE ON "
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS = 1105356) "
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID"

rs.Open Sql, strConBasa

Do While Not rs.EOF

FileSystem.FileCopy "\\222.15.19.251\ImagenesPDF\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".PDF", "D:\ExportarImagenes\1105356\" & Mid(rs!FK_LEGAJO_ETIQUETA, 1, 12) & ".PDF"
rs.MoveNext
Loop



End Sub

Private Sub cmdGodoyCruzCatastro_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
        
        
        
        
        
        
        
        
        
        
        
        
        
        Sql = " SELECT  ID_LEGAJO, COD_INDICE, FK_INDICES "
        Sql = Sql & " , LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA "
        Sql = Sql & " ,FECHA_DESDE, FECHA_HASTA, NRO_CAJA, DESCRIPCION "
        Sql = Sql & " From LEGAJOS "
        Sql = Sql & "  Where (COD_CLIENTE = 1156) "
        Sql = Sql & "  And (FK_INDICES = 10715)"
        Rem sql = sql & "  and nombre is null  "
        Sql = Sql & "   AND NRO_CAJA in( 1105820  )"
        Sql = Sql & "  ORDER BY LETRA_DESDE, LETRA_HASTA "
        
        rs.Open Sql, strConBasa
        
        Do While Not rs.EOF
       If Dir("I:\1156-GODOY CRUZ\2-CATASTRO\20-PLANOS\TERMINADAS\" & rs!NRO_CAJA & "\" & rs!ID_LEGAJO & ".tif", vbNormal) <> "" Then
         
        Rem I:\1156-GODOY CRUZ\2-CATASTRO\20-PLANOS\TERMINADAS
        Rem    If Dir("I:\1156-GODOY CRUZ\CATASTRO\PLANOS\PARA EL CLIENTE\" & rs!ID_LEGAJO & ".tif") <> "" Then
                Sql = " Update LEGAJOS "
                Sql = Sql & " Set Nombre ='SI'  "
                Sql = Sql & " WHERE ID_LEGAJO = " & rs!ID_LEGAJO
                ExecutarSql Sql
           
          Rem FileCopy "I:\1156-GODOY CRUZ\CATASTRO\PLANOS\PARA EL CLIENTE\" & rs!ID_LEGAJO & ".tif", "I:\1156-GODOY CRUZ\CATASTRO\PLANOS\PARA EL CLIENTE\control23022016\" & rs!ID_LEGAJO & ".tif"
            If Dir("D:\CATASTRO\" & rs!NRO_CAJA & "\" & rs!LETRA_DESDE & "\*.TIF") = "" Then
             
            FileSystem.MkDir "D:\CATASTRO\" & rs!NRO_CAJA & "\" & rs!LETRA_DESDE
            End If
            
             
             FileCopy "I:\1156-GODOY CRUZ\2-CATASTRO\20-PLANOS\TERMINADAS\" & rs!NRO_CAJA & "\" & rs!ID_LEGAJO & ".tif", "D:\CATASTRO\" & rs!NRO_CAJA & "\" & rs!LETRA_DESDE & "\" & Trim(rs!LETRA_DESDE) & "_" & Trim(rs!LETRA_HASTA) & "_" & Trim(rs!NRO_DESDE) & ".tif"
              
              Rem  Kill "I:\1156-GODOY CRUZ\CATASTRO\PLANOS\PARA EL CLIENTE\" & rs!ID_LEGAJO & ".tif"
           
           Else
                Sql = " Update LEGAJOS "
                Sql = Sql & " Set Nombre ='NOMUNI'  "
                Sql = Sql & " WHERE ID_LEGAJO = " & rs!ID_LEGAJO
                ExecutarSql Sql
           
           End If
           
            
            rs.MoveNext
        Loop
        


End Sub

Private Sub cmdGodoyCruzCatastroExportar_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
        
        
        
        
Sql = " SELECT LETRA_DESDE "
Sql = Sql & " From LEGAJOS"
Sql = Sql & " Where(COD_CLIENTE = 1156)"
Sql = Sql & " And (FK_INDICES = 10715)"
Sql = Sql & " GROUP BY LETRA_DESDE"
Sql = Sql & " ORDER BY LETRA_DESDE"
        rs.Open Sql, strConBasa
        
        Do While Not rs.EOF
        
        FileSystem.MkDir ("D:\Catastro\" & rs!LETRA_DESDE)
            rs.MoveNext
        Loop
        
        
        
        
        
        
        Sql = " SELECT  ID_LEGAJO, COD_INDICE, FK_INDICES "
        Sql = Sql & " , LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA "
        Sql = Sql & " ,FECHA_DESDE, FECHA_HASTA, NRO_CAJA, DESCRIPCION "
        Sql = Sql & " From LEGAJOS "
        Sql = Sql & "  Where (COD_CLIENTE = 1156) "
        Sql = Sql & "  And (FK_INDICES = 10715) "
       Rem   Sql = Sql & "   AND (NOMBRE = 'NO')"
        Sql = Sql & "  ORDER BY LETRA_DESDE, LETRA_HASTA , NRO_DESDE "
        
        
        Set rs = New ADODB.Recordset
        
        rs.Open Sql, strConBasa
        
        Do While Not rs.EOF
           
           
           FileCopy "I:\1156-GODOY CRUZ\CATASTRO\Control\" & rs!ID_LEGAJO & ".tif", "D:\Catastro\" & rs!LETRA_DESDE & "\" & Trim(rs!LETRA_DESDE) & "_" & Trim(rs!LETRA_HASTA) & "_" & Trim(rs!NRO_DESDE) & ".tif"
           
            
            rs.MoveNext
        Loop
        

End Sub


Private Sub cmdGodoyCruzCatastroExportarPDF_Click()
Dim rs As New ADODB.Recordset
    Dim Sql As String
        
        
        
        
Sql = " SELECT LETRA_DESDE "
Sql = Sql & " From LEGAJOS"
Sql = Sql & " Where(COD_CLIENTE = 1156)"
Sql = Sql & " And (FK_INDICES = 10715)"
Sql = Sql & " GROUP BY LETRA_DESDE"
Sql = Sql & " ORDER BY LETRA_DESDE"
        rs.Open Sql, strConBasa
        
        Do While Not rs.EOF
        
        FileSystem.MkDir ("D:\Catastro\" & rs!LETRA_DESDE)
            rs.MoveNext
        Loop
        
        
        
        
        
        
        Sql = " SELECT  ID_LEGAJO, COD_INDICE, FK_INDICES "
        Sql = Sql & " , LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA "
        Sql = Sql & " ,FECHA_DESDE, FECHA_HASTA, NRO_CAJA, DESCRIPCION "
        Sql = Sql & " From LEGAJOS "
        Sql = Sql & "  Where (COD_CLIENTE = 1156) "
        Sql = Sql & "  And (FK_INDICES = 10715) "
       Rem   Sql = Sql & "   AND (NOMBRE = 'NO')"
        Sql = Sql & "  ORDER BY LETRA_DESDE, LETRA_HASTA , NRO_DESDE "
        
        
        Set rs = New ADODB.Recordset
        
        rs.Open Sql, strConBasa
        
        Do While Not rs.EOF
           
           
           FileCopy "I:\1156-GODOY CRUZ\CATASTRO\PLANOS\PARA EL CLIENTE\Planos PDF\" & rs!ID_LEGAJO & ".pdf", "D:\Catastro\" & rs!LETRA_DESDE & "\" & Trim(rs!LETRA_DESDE) & Trim(rs!LETRA_HASTA) & "_" & Trim(rs!NRO_DESDE) & ".PDF"
           
            
            rs.MoveNext
        Loop
End Sub

Private Sub cmdGodoyCruzCatastroExportartif_Click()
Dim rs As New ADODB.Recordset
    Dim Sql As String
        
        
        
        
Sql = " SELECT LETRA_DESDE "
Sql = Sql & " From LEGAJOS"
Sql = Sql & " Where(COD_CLIENTE = 1156)"
Sql = Sql & " And (FK_INDICES = 10715)"
Sql = Sql & " GROUP BY LETRA_DESDE"
Sql = Sql & " ORDER BY LETRA_DESDE"
        rs.Open Sql, strConBasa
        
     Do While Not rs.EOF
        
        FileSystem.MkDir ("D:\Catastro\" & rs!LETRA_DESDE)
            rs.MoveNext
        Loop
        
        
        
        
        
        
        Sql = " SELECT ID_LEGAJO, COD_INDICE, FK_INDICES "
        Sql = Sql & " , LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA "
        Sql = Sql & " , FECHA_DESDE, FECHA_HASTA, NRO_CAJA, DESCRIPCION "
        Sql = Sql & " From LEGAJOS "
        Sql = Sql & " Where (COD_CLIENTE = 1156) "
        Sql = Sql & " AND (FK_INDICES = 10715) "
        Sql = Sql & " AND (NOMBRE = 'SI')"
        Sql = Sql & " ORDER BY LETRA_DESDE, LETRA_HASTA , NRO_DESDE "
        
        
        Set rs = New ADODB.Recordset
        
        rs.Open Sql, strConBasa
        
        Do While Not rs.EOF
            FileCopy "I:\1156-GODOY CRUZ\CATASTRO\PLANOS\PARA EL CLIENTE\control23022016\" & rs!ID_LEGAJO & ".tif", "D:\Catastro\" & rs!LETRA_DESDE & "\" & Trim(rs!LETRA_DESDE) & "_" & Trim(rs!LETRA_HASTA) & "_" & Trim(rs!NRO_DESDE) & ".tif"
            FileCopy "I:\1156-GODOY CRUZ\CATASTRO\PLANOS\PARA EL CLIENTE\control23022016 pdf\" & rs!ID_LEGAJO & ".pdf", "D:\Catastro\" & rs!LETRA_DESDE & "\" & Trim(rs!LETRA_DESDE) & "_" & Trim(rs!LETRA_HASTA) & "_" & Trim(rs!NRO_DESDE) & ".pdf"
            rs.MoveNext
        Loop
End Sub

Private Sub cmdImprimir_Click()
    Dim Sql As String
    Sql = " SELECT V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.SUB_LOTE, V_DOCUMENTOS_LOTES_PRODUCION.FK_CLIENTES,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.DESCRIPCION, V_DOCUMENTOS_LOTES_PRODUCION.COSTOPREPARACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_PREPARACION, V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_PREPARACION, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FECHA_PREPARACION, V_DOCUMENTOS_LOTES_PRODUCION.FECHA_SCANNER, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_SCANNER, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_SCANNER,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FECHA_REORDENAR, V_DOCUMENTOS_LOTES_PRODUCION.FECHA_INDEXACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_REORDENAR, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_REORDENAR,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.COSTODIGITALIZACION, V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_INDEXACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_INDEXACION, V_DOCUMENTOS_LOTES_PRODUCION.COSTOARMADO,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.COSTOINDEXACION ,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_CAJAS, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_DIRECTORIO_SALIDA"
    Sql = Sql & " FROM   basasql.dbo.V_DOCUMENTOS_LOTES_PRODUCION V_DOCUMENTOS_LOTES_PRODUCION"
    Sql = Sql & " Where V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE in(" & Mid(txtLotesExportar.Text, 2) & ")"
    Sql = Sql & " ORDER BY V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE"
    frmReportes.ImprimirReporte PasoReportes & "rpt_Digitalizacion.rpt", Sql, True
        
        
End Sub

Private Sub cmdImprimirCaja_Click()
 
 
 
 Dim Sql As String
 
 If txtCajaTerminada.Text <> "" Then
    Sql = " SELECT V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.SUB_LOTE, V_DOCUMENTOS_LOTES_PRODUCION.FK_CLIENTES,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.DESCRIPCION, V_DOCUMENTOS_LOTES_PRODUCION.COSTOPREPARACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_PREPARACION, V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_PREPARACION, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FECHA_PREPARACION, V_DOCUMENTOS_LOTES_PRODUCION.FECHA_SCANNER, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_SCANNER, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_SCANNER,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FECHA_REORDENAR, V_DOCUMENTOS_LOTES_PRODUCION.FECHA_INDEXACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_REORDENAR, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_REORDENAR,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.COSTODIGITALIZACION, V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_INDEXACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_INDEXACION, V_DOCUMENTOS_LOTES_PRODUCION.COSTOARMADO,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.COSTOINDEXACION ,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_CAJAS, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_DIRECTORIO_SALIDA"
    Sql = Sql & " FROM   basasql.dbo.V_DOCUMENTOS_LOTES_PRODUCION V_DOCUMENTOS_LOTES_PRODUCION"
    Sql = Sql & " Where V_DOCUMENTOS_LOTES_PRODUCION.FK_CAJAS = " & txtCajaTerminada.Text
    Sql = Sql & " ORDER BY V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE"
    frmReportes.ImprimirReporte PasoReportes & "rpt_Digitalizacion_Caja.rpt", Sql, True
 Else
  MsgBox "Ingrese la caja"
 End If
    
End Sub

Private Sub cmdImprimirLote_Click()
    Dim Sql As String
    Sql = " SELECT V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.SUB_LOTE, V_DOCUMENTOS_LOTES_PRODUCION.FK_CLIENTES,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.DESCRIPCION, V_DOCUMENTOS_LOTES_PRODUCION.COSTOPREPARACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_PREPARACION, V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_PREPARACION, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FECHA_PREPARACION, V_DOCUMENTOS_LOTES_PRODUCION.FECHA_SCANNER, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_SCANNER, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_SCANNER,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FECHA_REORDENAR, V_DOCUMENTOS_LOTES_PRODUCION.FECHA_INDEXACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_REORDENAR, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_REORDENAR,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.COSTODIGITALIZACION, V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_INDEXACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_INDEXACION, V_DOCUMENTOS_LOTES_PRODUCION.COSTOARMADO,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.COSTOINDEXACION ,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_CAJAS, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_DIRECTORIO_SALIDA"
    Sql = Sql & " FROM   basasql.dbo.V_DOCUMENTOS_LOTES_PRODUCION V_DOCUMENTOS_LOTES_PRODUCION"
    Sql = Sql & " Where V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE in(" & Mid(txtLotesExportar.Text, 2) & ")"
    Sql = Sql & " ORDER BY V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE"
    frmReportes.ImprimirReporte PasoReportes & "rpt_Digitalizacion.rpt", Sql, True
        
End Sub


Private Sub cmdImprimirResumen_Click()
Dim Sql As String
 
 If txtCajaTerminada.Text <> "" Then
    Sql = " SELECT V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.SUB_LOTE, V_DOCUMENTOS_LOTES_PRODUCION.FK_CLIENTES,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.DESCRIPCION, V_DOCUMENTOS_LOTES_PRODUCION.COSTOPREPARACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_PREPARACION, V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_PREPARACION, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FECHA_PREPARACION, V_DOCUMENTOS_LOTES_PRODUCION.FECHA_SCANNER, "
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_SCANNER, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_SCANNER,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FECHA_REORDENAR, V_DOCUMENTOS_LOTES_PRODUCION.FECHA_INDEXACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_REORDENAR, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_REORDENAR,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.COSTODIGITALIZACION, V_DOCUMENTOS_LOTES_PRODUCION.FK_PERSONAL_INDEXACION,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_PERSONAL_INDEXACION, V_DOCUMENTOS_LOTES_PRODUCION.COSTOARMADO,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.COSTOINDEXACION ,"
    Sql = Sql & " V_DOCUMENTOS_LOTES_PRODUCION.FK_CAJAS, V_DOCUMENTOS_LOTES_PRODUCION.NOMBRE_DIRECTORIO_SALIDA"
    Sql = Sql & " FROM   basasql.dbo.V_DOCUMENTOS_LOTES_PRODUCION V_DOCUMENTOS_LOTES_PRODUCION"
    Sql = Sql & " Where V_DOCUMENTOS_LOTES_PRODUCION.FK_CAJAS = " & txtCajaTerminada.Text
    Sql = Sql & " ORDER BY V_DOCUMENTOS_LOTES_PRODUCION.ID_DOCUMENTOS_DIGITALES_LOTE"
    frmReportes.ImprimirReporte PasoReportes & "rpt_Digitalizacion_control.rpt", Sql, True
 Else
  MsgBox "Ingrese la caja"
 End If
End Sub

Private Sub cmdleerLegajos_Click()

Dim lote As String
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rslegajos As New ADODB.Recordset
Dim ID_LEGAJO As Long



Sql = " SELECT     FK_DOCUMENTOS_DIGITALES_LOTE, COD_LOTE, DESCRIPCION, NRO_DESDE,letra_DESDE , ID"
Sql = Sql & "  From DOCUMENTOS_DIGITALES"
Sql = Sql & "  Where FK_DOCUMENTOS_DIGITALES_LOTE in(" & Mid(txtLotesExportar.Text, 2) & ")"
Sql = Sql & "  ORDER BY FK_DOCUMENTOS_DIGITALES_LOTE "

Set rs = New ADODB.Recordset
rs.Open Sql, strConBasa

Do While Not rs.EOF
     Sql = " SELECT     ID_LEGAJO, COD_INDICE, FK_INDICES, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, CLIENTE_LEGAJO,"
    Sql = Sql & " Descripcion , NRO_CAJA, COD_CLIENTE "
    Sql = Sql & " From LEGAJOS "
    If IsNumeric(rs!LETRA_DESDE) Then
    Sql = Sql & " where ID_LEGAJO = " & CLng(rs!LETRA_DESDE)
    ID_LEGAJO = CLng(rs!LETRA_DESDE)
    Else
        If IsNull(rs!LETRA_DESDE) Then
            Sql = Sql & " where ID_LEGAJO = " & rs!NRO_DESDE
            ID_LEGAJO = rs!NRO_DESDE
        Else
         GoTo echo
        End If
    End If
    
    Set rslegajos = New ADODB.Recordset
    
    rslegajos.Open Sql, strConBasa, 0, 1
    
     If Not rslegajos.EOF Then
     
            Sql = "  Update DOCUMENTOS_DIGITALES "
            Sql = Sql & " SET "
            Sql = Sql & " INDICE ='" & rslegajos!Cod_Indice & "'"
            Sql = Sql & " , COD_CLIENTE =" & rslegajos!COD_CLIENTE
            Sql = Sql & " , LETRA_DESDE ='" & rslegajos!LETRA_DESDE & "'"
            Sql = Sql & " , LETRA_HASTA ='" & rslegajos!LETRA_HASTA & "'"
            Sql = Sql & " , NRO_DESDE =" & rslegajos!NRO_DESDE
            Sql = Sql & " , NRO_HASTA =" & rslegajos!NRO_HASTA
            If Not IsNull(rslegajos!FECHA_DESDE) Then
             Sql = Sql & " , FECHA_DESDE =" & FechaFormato(rslegajos!FECHA_DESDE)
            End If
            If Not IsNull(rslegajos!FECHA_HASTA) Then
                Sql = Sql & " , FECHA_HASTA =" & FechaFormato(rslegajos!FECHA_HASTA)
            End If
            
            Sql = Sql & " , FK_ID_LEGAJO =" & ID_LEGAJO
            Sql = Sql & " , DESCRIPCION ='" & rslegajos!Descripcion & "'"
            Sql = Sql & " ,  COD_ESTADO =100"
            Sql = Sql & " Where ID = " & rs!ID
            ExecutarSql Sql
            
     End If
     
    



echo:
    

    rs.MoveNext
Loop


MsgBox "Terminado"
End Sub

Private Sub cmdLoteTerminado_Click()
Dim Sql As String
If IsNull(ctlPersonalIndexacion.Valor) Then

    MsgBox "Ingrese el personal"
    Exit Sub
End If
Sql = " Update DOCUMENTOS_DIGITALES_LOTE "
Sql = Sql & "  Set FK_ESTADO = 100 "
Sql = Sql & "  , FK_PERSONAL_INDEXACION =" & ctlPersonalIndexacion.Valor
Sql = Sql & "  , FECHA_INDEXACION = " & SysDateMinutoSegundo
Sql = Sql & "  Where ID_DOCUMENTOS_DIGITALES_LOTE = " & Trim(lblLote.Caption)

ExecutarSql Sql
MsgBox "Lote terminado"
End Sub

Private Sub cmdOcr_Click()
'
'
'Dim Sql As String
'    Dim rsImagenes As New ADODB.Recordset
'
'    Dim docOrigen As MODI.Document
'            Set docOrigen = New MODI.Document
'
'
'
'
'            Dim MyName As String
'
'    rsBuscar.Requery
'
'        MousePointer = 11
'
'        Dim NombreArchivo As String
'        Dim Manzana As String
'        Dim Parcela As String
'        Dim SubParcela As String
'        Dim Divi As String
'        Dim i As Integer
'        Do While Not rsBuscar.EOF
'        Sql = "  SELECT ID, DIRECTORIO_PASO "
'        Sql = Sql & " From  DOCUMENTOS_DIGITALES  "
'        Sql = Sql & "  WHERE    COD_CLIENTE = " & rsBuscar!COD_CLIENTE
'        Sql = Sql & "  AND LOTE =  '" & rsBuscar!Lote & "'"
'        Sql = Sql & "  AND COD_ESTADO = " & rsBuscar!Cod_Estado
'        Sql = Sql & "  AND OCR IS NULL "
'        Sql = Sql & "  ORDER BY ID "
'        Set rsImagenes = New ADODB.Recordset
'
'        rsImagenes.Open Sql,ConActiva, adOpenForwardOnly, adLockReadOnly
'
''            Do While Not rsImagenes.EOF
''               docOrigen.Create PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif"
''               docOrigen.OCR miLANG_SPANISH, True
''               docOrigen.Save
''               Sql = " Update DOCUMENTOS_DIGITALES"
''               Sql = Sql & "  Set OCR = 1"
''               Sql = Sql & " Where ID = " & rsImagenes!ID
''               ExecutarSql Sql
''               cmdOcr.Caption = I
''               I = I + 1
''               cmdOcr.Refresh
''
''               rsImagenes.MoveNext
''            Loop
'       rsBuscar.MoveNext
'Loop
'MousePointer = 0
'' cmdOcr.Caption = "OCR"
'MsgBox "Operacion terminada"



End Sub

Private Sub cmdTurismo_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Etiqueta As Long





Sql = " SELECT     ID, NOMBRE_ARCHIVO, NRO_CAJA, CANTIDAD_IMAGEN, LOTEHORA, FK_CLIENTE, MESAÑO, ID_PIEZAS"
Sql = Sql & "  From CANTIDAD_IMAGEN"
Sql = Sql & "  Where (FK_CLIENTE = 405)   AND (ID_PIEZAS IS NULL)"

rs.Open Sql, strConBasa

Do While Not rs.EOF
    Etiqueta = Mid(rs!Nombre_Archivo, 1, 7)
    Sql = " SELECT id From pieza_administrativa "
    Sql = Sql & " Where DIGITOVERIFICADOR = " & Etiqueta
    
    
' Sql = " SELECT id "
'  Sql = Sql & "   , denominacion"
'  Sql = Sql & "   ,ruta"
'  Sql = Sql & "   ,    idpiezaadministrativa"
'Rem   Sql = Sql & "   , enbasa "
'Sql = Sql & " From documentacion_pieza "
' Sql = Sql & "  Where ruta like '" & Etiqueta & "%'"
'
    Set rs2 = New ADODB.Recordset
    rs2.Open Sql, strConBasa
    If Not rs2.EOF Then
        Sql = " UPDATE    CANTIDAD_IMAGEN "
        Sql = Sql & " Set ID_PIEZAS = " & rs2!ID
        Sql = Sql & " , FECHA_ACT='29/12/2015'"
        Sql = Sql & " Where ID = " & rs!ID
        Sql = Sql & " AND (ID_PIEZAS IS NULL)"
        ExecutarSql Sql
    End If
    
    
    rs.MoveNext
Loop


End Sub

Private Sub cmdUnirLotes_Click()
UnirFlyersNotti
End Sub

Private Sub cmdRotar1_Click()

'Dim sql As String
'    Dim rsImagenes As New ADODB.Recordset
'
'    Dim docOrigen As MODI.Document
'            Set docOrigen = New MODI.Document
'      Dim i As Integer
'
'
'
'            Dim MyName As String
'
'    rsBuscar.Requery
'
'        MousePointer = 11
'
'
'        Dim NombreArchivo As String
'        Dim Manzana As String
'        Dim Parcela As String
'        Dim SubParcela As String
'        Dim Divi As String
'
'        Do While Not rsBuscar.EOF
'        sql = "  SELECT ID, DIRECTORIO_PASO "
'        sql = sql & " From  DOCUMENTOS_DIGITALES  "
'        sql = sql & "  WHERE    COD_CLIENTE = " & rsBuscar!COD_CLIENTE
'        sql = sql & "  AND LOTE =  '" & rsBuscar!lote & "'"
'        sql = sql & "  AND COD_ESTADO = " & rsBuscar!Cod_Estado
'        sql = sql & "  AND rotar IS NULL "
'        sql = sql & "  ORDER BY ID "
'        Set rsImagenes = New ADODB.Recordset
'
'        rsImagenes.Open sql, ConActiva, adOpenForwardOnly, adLockReadOnly
'
'            Do While Not rsImagenes.EOF
'               docOrigen.Create PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif"
'               docOrigen.Images.Item(0).Rotate 180
'               docOrigen.Save
'               sql = " Update DOCUMENTOS_DIGITALES"
'               sql = sql & "  Set ROTAR = 1"
'               sql = sql & " Where ID = " & rsImagenes!ID
'               ExecutarSql sql
'               cmdOcr.Caption = i
'               i = i + 1
'               cmdOcr.Refresh
'
'               rsImagenes.MoveNext
'            Loop
'       rsBuscar.MoveNext
'Loop
'MousePointer = 0
'cmdOcr.Caption = "OCR"
'MsgBox "Operacion terminada"
 End Sub

Private Sub CMDuTIL_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String
'
'Sql = " SELECT     DESCRIPCION, FK_CLIENTES, FK_CAJAS,ID_DOCUMENTOS_DIGITALES_LOTE"
'Sql = Sql & " From DOCUMENTOS_DIGITALES_LOTE "
'Sql = Sql & "  ORDER BY FK_CLIENTES, FK_CAJAS, DESCRIPCION "
'
'
'Sql = " SELECT LACAJALOTES$.LOTE, LACAJALOTES$.TIPO_DOC, LACAJALOTES$.CANTIDAD, LACAJALOTES$.RUTA, LACAJALOTES$.HOJA_CONTROL,"
'Sql = Sql & " LACAJALOTES$.FECHA_EXPOR , LACAJALOTES$.DESCRIPCION, LA_CAJA_TOMADOR.ID_TOMADOR"
'Sql = Sql & " FROM       LACAJALOTES$ INNER JOIN"
'Sql = Sql & " LA_CAJA_TOMADOR ON LACAJALOTES$.TOMADOR = LA_CAJA_TOMADOR.DESCRIPCION"


Sql = " SELECT     INDICES.COD_CLIENTE, INDICES.INDICE, INDICES.ID"
Sql = Sql & " FROM         INDICE_DIGITALIZACION INNER JOIN"
              Sql = Sql & "        INDICES ON INDICE_DIGITALIZACION.COD_CLIENTE = INDICES.COD_CLIENTE AND INDICE_DIGITALIZACION.COD_INDICE = INDICES.INDICE"
                      

rs.Open Sql, ConActiva, 0, 1
Dim Registros As Long

Do While Not rs.EOF

    
   ConBasa.CommandTimeout = 6000
    
    
    Sql = "    Update INDICE_DIGITALIZACION"
Sql = Sql & " Set FK_INDICES = " & rs!ID
Sql = Sql & " WHERE     COD_CLIENTE = " & rs!COD_CLIENTE
Sql = Sql & " AND COD_INDICE ='" & rs!Indice & "'"
    
    
 Registros = ExecutarSql(Sql)

    If Registros = 0 Then
        MsgBox "LLLL"
    End If
    
    

    rs.MoveNext
Loop





End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String

Sql = "SELECT     FECHA_INCORPORACION, FECHA_INCORPORACION2, ID "
Sql = Sql & " From DOCUMENTOS_DIGITALES "
Sql = Sql & " Where (Not (FECHA_INCORPORACION Is Null)) ORDER BY ID "

rs.CursorLocation = adUseClient

rs.Open Sql, ConActiva, adOpenKeyset, adLockBatchOptimistic


Do While Not rs.EOF
If IsDate(Trim(rs!FECHA_INCORPORACION)) Then
    
    ExecutarSql " Update DOCUMENTOS_DIGITALES SET  FECHA_INCORPORACION2 = " & FechaServerTipo(rs!FECHA_INCORPORACION) & " Where ID = " & rs!ID
    
    Else
    
    End If

    rs.MoveNext
Loop


End Sub

Private Sub cmdMuniCapital_Click()
Dim Sql As String
    Dim rsImagenes As New ADODB.Recordset
   
    rsBuscar.Requery
       
        MousePointer = 11
        
        Dim NombreArchivo As String
        Dim Manzana As String
        Dim Parcela As String
        Dim SubParcela As String
        Dim Divi As String
        
        Do While Not rsBuscar.EOF
        Sql = "  SELECT ID, COD_CLIENTE, LOTE, COD_ESTADO, DIRECTORIO_PASO,  NRO_DESDE, NRO_HASTA, NRO_UNO, NRO_DOS "
        Sql = Sql & " From  DOCUMENTOS_DIGITALES  "
        Sql = Sql & "  WHERE    COD_CLIENTE = " & rsBuscar!COD_CLIENTE
        Sql = Sql & "  AND LOTE =  '" & rsBuscar!lote & "'"
        Sql = Sql & "  AND COD_ESTADO = " & rsBuscar!Cod_Estado
        Sql = Sql & "  ORDER BY nro_desde, nro_hasta, nro_uno, nro_dos"
        
        


        
        
        Set rsImagenes = New ADODB.Recordset
        
        rsImagenes.Open Sql, ConActiva, 0, 1
        
            Do While Not rsImagenes.EOF
                If Not IsNull(rsImagenes!NRO_DESDE) Then
                    Manzana = Format(rsImagenes!NRO_DESDE, "0000")
                 Else
                    Manzana = "0000"
                 
                 End If
                 
                 If Not IsNull(rsImagenes!NRO_HASTA) Then
                    Parcela = Format(rsImagenes!NRO_HASTA, "000")
                 Else
                    Parcela = "000"
                 End If
                 
                 If Not IsNull(rsImagenes!NRO_UNO) Then
                    SubParcela = Format(rsImagenes!NRO_UNO, "000")
                 Else
                    SubParcela = "000"
                 
                 End If
                 
                 If Not IsNull(rsImagenes!NRO_DOS) Then
                    Divi = Format(rsImagenes!NRO_DOS, "000")
                  Else
                    Divi = "000"
                 End If
                 
                 
                 
            
               If Dir(txtPasoImagenesFinal & "MZA " & rsImagenes!NRO_DESDE, vbDirectory) = "" Then
                FileSystem.MkDir txtPasoImagenesFinal & "MZA " & rsImagenes!NRO_DESDE
               
               End If
               FileSystem.FileCopy PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif", txtPasoImagenesFinal & "MZA " & rsImagenes!NRO_DESDE & "\" & Manzana & Parcela & SubParcela & Divi & "_" & rsImagenes!ID & ".tif"
               rsImagenes.MoveNext
            Loop
       rsBuscar.MoveNext
Loop
MousePointer = 0
MsgBox "Operacion terminada"
End Sub

Private Sub ChandonExportacion()
'    Dim conAcces As New ADODB.Connection
'    Dim rs As ADODB.Recordset
'    Dim RsSql As ADODB.Recordset
'    conAcces.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Mis Sistemas\chandon\chandon.mdb;Persist Security Info=False"
'
'        Dim sql As String
'
'
'
'
'
'    Dim MaxLegajo As Integer
'    Set rs = New ADODB.Recordset
'
'        sql = " SELECT max(Legajos.id_Legajo) as MaxLegajo FROM Legajos "
'        rs.Open sql, conAcces
'
'   If Not rs.EOF Then
'        MaxLegajo = rs!MaxLegajo + 1
'   Else
'
'
'   End If
'
'
'
'Dim MaxImagenes As Integer
'
'    Set rs = New ADODB.Recordset
'
'    sql = " SELECT max(Imagenes.id_imagen) as maxImagen FROM Imagenes "
'    rs.Open sql, conAcces
'    If Not rs.EOF Then
'        MaxImagenes = rs!maxImagen + 1
'    End If
'
'
'    sql = " SELECT  COD_CLIENTE, INDICE, FECHA_INCORPORACION, ID, LETRA_DESDE, LETRA_HASTA, DIRECTORIO_PASO, nombre "
'        sql = sql & " FROM DOCUMENTOS_DIGITALES "
'        sql = sql & " WHERE COD_CLIENTE = 47 "
'        sql = sql & " AND (INDICE = N'007012') "
'        sql = sql & " AND FECHA_INCORPORACION = '29/12/2008'"
'sql = sql & " ORDER BY ID"
'
'    Set RsSql = New ADODB.Recordset
'    RsSql.Open sql, strConBasa , 0 ,1
'    Do While Not RsSql.EOF
'            sql = "  INSERT INTO Legajos ( id_Legajo, Tipo_Operacion, Bodega )"
'            sql = sql & " Values ( "
'            sql = sql & MaxLegajo & ",'" & Trim(RsSql!LETRA_DESDE) & "','" & Trim(RsSql!Nombre) & "')"
'            conAcces.Execute sql
'
'            sql = "  INSERT INTO Imagenes ( id_imagen, Paso_Origen, id_Tipo_Doc, id_legajo)"
'            sql = sql & " Values ( "
'            sql = sql & MaxImagenes & ",'" & Trim(RsSql!ID) & "',3," & MaxLegajo & " )"
'            conAcces.Execute sql
'
'            sql = "  INSERT INTO Imagenes ( id_imagen, Paso_Origen, id_Tipo_Doc, id_legajo)"
'            sql = sql & " Values ( "
'            sql = sql & MaxImagenes & ",'" & Trim(RsSql!ID) & "',1," & MaxLegajo & " )"
'            Rem conAcces.Execute sql
'
'             sql = "  INSERT INTO DatosImagenes ( id_imagen, id_Campo, Valor_CampoT)"
'            sql = sql & " Values ( "
'            sql = sql & MaxImagenes & ",3,'" & Trim(RsSql!LETRA_DESDE) & "')"
'            conAcces.Execute sql
'            sql = "  INSERT INTO DatosImagenes ( id_imagen, id_Campo, Valor_CampoT)"
'            sql = sql & " Values ( "
'            sql = sql & MaxImagenes & ",1,'" & Trim(RsSql!LETRA_HASTA) & "')"
'            conAcces.Execute sql
'
'            FileCopy "\\Base\Imagenes\" & RsSql!DIRECTORIO_PASO & "\" & RsSql!ID & ".TIF", "D:\Mis Sistemas\chandon\Imagenes\" & MaxImagenes & ".TIF"
'
'        RsSql.MoveNext
'        MaxImagenes = MaxImagenes + 1
'        MaxLegajo = MaxLegajo + 1
'    Loop
'
'
'
'

End Sub

Private Sub Command10_Click()
 
        Dim ApExcel As Excel.Application
        Rem Dim ApExcel As Object
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
        Dim B_Error As Boolean
        Dim i As Integer
        Dim msgError  As String
        Dim UltimaFila As Integer
        Dim ErrorDoc As String
        Dim ERRORCAJA As String
        Dim ValidarIndiceSecundario As Boolean
        Dim Sql As String
        Dim rs As ADODB.Recordset
        
        
       On Error GoTo salir
        'abrir hoja excel
        Set ApExcel = New Excel.Application
        
        Set libroEx = Excel.Workbooks.Open("D:\Mis Sistemas\La Caja\Asegurados_2_11_112_5020957892801.XLS")
        Set hojaEx = libroEx.Worksheets.Item(1)
        Dim Documento As Long
        
        
        For i = 1 To 5700
            Documento = hojaEx.Cells(i, 3)
            
  If Trim(hojaEx.Cells(i, 25)) = "" Then
            
                        Set rs = New ADODB.Recordset
                        
                        
                        Sql = "      SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO"
                        Sql = Sql & " FROM         DOCUMENTOS_DIGITALES INNER JOIN"
                        Sql = Sql & "    DOCUMENTOS_DIGITALES_LOTE ON"
                        Sql = Sql & "   DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
                        Sql = Sql & " Where                 (DOCUMENTOS_DIGITALES_LOTE.TIPO_DOCUMENTO = 'SOLICITUD')"
                        Sql = Sql & " AND (COD_CLIENTE = 163) "
                        Sql = Sql & " And NRO_DESDE = " & Documento
                        
                        rs.Open Sql, ConActiva, 0, 1
                        
                        Do While Not rs.EOF
                         If FileSystem.FileLen(PasoImagenes & "\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif") < 70000 And FileSystem.FileLen(PasoImagenes & "\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif") > 28000 Then
                            hojaEx.Cells(i, 1) = rs!ID
                            hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\" & Documento & ".tif"
                            FileCopy PasoImagenes & "\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", "D:\Mis Sistemas\La Caja\" & Documento & ".tif"
                            Exit Do
                          End If
                            rs.MoveNext
                        Loop
       End If
        Debug.Print Documento
        Debug.Print i
        Next
        libroEx.Save
        libroEx.Close
            ApExcel.Quit
         Set hojaEx = Nothing
            Set libroEx = Nothing
            Set ApExcel = Nothing
            
salir:
        
        
'        'Control de formato
'        If Not (Mid(Cells(4, 3), 1, 18) <> "Nombre:" Or Mid(Cells(4, 3), 1, 18) <> "Nombre y Apellido:") Then
'            MsgBox "Error en el formato", vbInformation
'            Control_Excel_Cliente = True
'
'            Set hojaEx = Nothing
'            Set libroEx = Nothing
'            Set ApExcel = Nothing
'            Exit Function
'        End If
' hojaEx.Cells(i, 1) = rs!ID
''                 hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\" & txtLote.Text & "\" & NombreArchivo
End Sub

Private Sub Command11_Click()
MsgBox DigitoVerificadorExpreso("13100303519480202")
End Sub

Private Sub Command12_Click()
'Dim direcOrig(90) As String
'Dim direcFin(90) As String
'
'Dim i As Integer
'Dim Caja As Long
'Dim sFolderPath As String
'Dim sArchivo As String
'Dim PasoFinal As String
'Dim PasoTeleform As String
'Dim PasoDig As String
'Dim PasoDirSubDir As String
'
'
'If Trim(txtCaja.Text) <> "" Then
'    Caja = txtCaja.Text
'Else
'    MsgBox "Ingrese la caja"
'    Exit Sub
'End If
'
'
'
'   PasoDig = txtPasoFinalDamsu
'   sArchivo = Dir(PasoDig & "\*", vbDirectory)
'
'    If sArchivo = "" Then
'        MsgBox "No existe la caja" & Caja & " en " & PasoDig
'        Exit Sub
'    End If
'
'    Do While sArchivo <> ""
'
'        If Len(sArchivo) > 8 Then
'            direcOrig(i) = sArchivo
'            direcFin(i) = Mid(sArchivo, 1, 7) & Format(Mid(sArchivo, 8), "000")
'        Else
'            direcOrig(i) = ""
'            direcFin(i) = ""
'        End If
'            i = i + 1
'            sArchivo = Dir
'    Loop
'
''PasoTeleform = "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\PARA TELEFORM"
''
''If Dir(PasoTeleform & "\" & Caja, vbDirectory) = "" Then
'' FileSystem.MkDir PasoTeleform & "\" & Caja
'' PasoFinal = PasoTeleform & "\" & Caja
''Else
''    PasoFinal = PasoTeleform & "\" & Caja
''End If
''
''    Dim ArchivoNombreFinal As String
''        For i = 0 To 90
''           If Trim(direcOrig(i)) <> "" Then
''              If Len(Trim(direcOrig(i))) > 2 Then
''                PasoDirSubDir = PasoDig & Caja & "\" & direcOrig(i)
''                sArchivo = Dir(PasoDirSubDir & "\*.tif")
''                Do While sArchivo <> ""
''                   ArchivoNombreFinal = Mid(sArchivo, 1, Len(sArchivo) - 4)
''                   ArchivoNombreFinal = Format(ArchivoNombreFinal, "000")
''                   FileCopy PasoDirSubDir & "\" & sArchivo, PasoFinal & "\" & direcFin(i) & ArchivoNombreFinal & ".TIF"
''                   sArchivo = Dir
''                Loop
''              End If
''            End If
''        Next
''    MsgBox "Terminados"
End Sub

Private Sub Command13_Click()
        Dim MyName As String
        Dim Sql As String
       
       
       Dim Nombre_Archivo As String
       Dim NRO_CAJA As String
       Dim CANTIDAD_IMAGEN As Integer
       Dim LOTEHORA As String
       Dim FK_CLIENTE As String
       Dim MesAño As String
       MesAño = txtMesAño.Text
        
        FK_CLIENTE = ctlClienteContar.Valor
        LOTEHORA = Format(Now, "YYMMDDHHm")
        MyName = Dir(cboPasoContar.Text & "\" & txtCajaContar & "\*.tif")    ' Retrieve the first entry.
        Do While MyName <> ""   ' Start the loop.
            Nombre_Archivo = MyName
            NRO_CAJA = txtCajaContar
            ImagXpress1.FileName = cboPasoContar.Text & "\" & txtCajaContar & "\" & MyName
            CANTIDAD_IMAGEN = ImagXpress1.Pages
            Sql = " Insert "
            Sql = Sql & " Into CANTIDAD_IMAGEN("
            Sql = Sql & " NOMBRE_ARCHIVO"
            Sql = Sql & ", NRO_CAJA"
            Sql = Sql & ", CANTIDAD_IMAGEN"
            Sql = Sql & ", LOTEHORA"
            Sql = Sql & ", FK_CLIENTE"
            Sql = Sql & ", MESAÑO )"
            Sql = Sql & " VALUES ("
            Sql = Sql & "'" & Nombre_Archivo & "'"
            Sql = Sql & "," & NRO_CAJA
            Sql = Sql & "," & CANTIDAD_IMAGEN
            Sql = Sql & ",'" & LOTEHORA & "'"
            Sql = Sql & "," & FK_CLIENTE
             Sql = Sql & ",'" & MesAño & "')"
            ExecutarSql Sql
            MyName = Dir()   ' Get next entry.
            
            
        Loop

MsgBox "TERMINADO"

End Sub

Private Function Directorios(MyPath As String) As Integer

Dim MyName As String
Dim i As Integer
Dim D As String



    MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
    Do While MyName <> ""   ' Start the loop.
           i = i + 1
           Directorios_Nombres(i) = MyName
          D = D & vbCrLf & MyName
       MyName = Dir()   ' Get next entry.
    Loop
    Clipboard.Clear
    Clipboard.SetText D


End Function

Private Sub Command14_Click()
Dim a(100) As String
Directorios ("D:\Dansu\")
  
  
  
End Sub

Private Sub Command15_Click()
Dim s As String
Dim Sql As String
Dim rs As New ADODB.Recordset


Sql = " SELECT      NOMBRE_ARCHIVO, ID_PIEZAS"
Sql = Sql & " From CANTIDAD_IMAGEN"
Sql = Sql & "  Where (FK_CLIENTE = 405) AND (Not (ID_PIEZAS Is Null)) AND FECHA_ACT='29/12/2015'"


Sql = " SELECT     ID, NOMBRE_ARCHIVO, NRO_CAJA, CANTIDAD_IMAGEN, LOTEHORA, FK_CLIENTE, MESAÑO, ID_PIEZAS, FECHA_ACT"
Sql = Sql & "  From CANTIDAD_IMAGEN"
Sql = Sql & "  WHERE     (FK_CLIENTE = 405) and FECHA_ACT='29/12/2015' AND (NOT (ID_PIEZAS IS NULL))"

Sql = Sql & " ORDER BY ID DESC"


rs.Open Sql, strConBasa


Do While Not rs.EOF
   s = s & vbCrLf & "INSERT INTO documentacion_pieza_administrativa (denominacion, ruta, idpiezaadministrativa) VALUES ('TEST', '" & Mid(rs!Nombre_Archivo, 1, 7) & ".PDF" & "', " & rs!ID_PIEZAS & ");"
   rs.MoveNext
Loop
Clipboard.Clear
Clipboard.SetText s

End Sub

Private Sub Command16_Click()

Dim Paso As String
Dim Sql As String
Dim MyName As String


    MyName = Dir("D:\CHANDONPDF\*.pdf")   ' Retrieve the first entry.
    
    Do While MyName <> ""   ' Start the loop.
                Sql = " INSERT INTO basasql.dbo.DIR "
                Sql = Sql & vbCrLf & " ("
                Sql = Sql & vbCrLf & "PASO_COMPLETO "
                Sql = Sql & vbCrLf & ", ARCHIVO "
                Sql = Sql & vbCrLf & ",DIRECTORIO "
                Sql = Sql & vbCrLf & ")"
                Sql = Sql & vbCrLf & "VALUES("
                Sql = Sql & vbCrLf & "'" & "D:\CHANDONPDF\" & MyName & "'"
                Sql = Sql & vbCrLf & ",'" & MyName & "'"
                Sql = Sql & vbCrLf & ",'" & Mid(MyName, 1, 6) & "'"
                Sql = Sql & vbCrLf & ")"
                ExecutarSql Sql
                MyName = Dir()   ' Get next entry.
    Loop


End Sub

Private Sub Command17_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
Sql = "SELECT     DIRECTORIO"
Sql = Sql & " From Dir"
Sql = Sql & "  GROUP BY DIRECTORIO"
rs.Open Sql, strConBasa

Do While Not rs.EOF
    FileSystem.MkDir "D:\CHANDONPDF\CUENTAS\" & rs!directorio
    rs.MoveNext
Loop



End Sub

Private Sub Command18_Click()
Dim Paso As String
Dim Sql As String
Dim MyName As String


    MyName = Dir("D:\CHANDONPDF\*.pdf")   ' Retrieve the first entry.
    
    Do While MyName <> ""   ' Start the loop.
          
             
              FileCopy "D:\CHANDONPDF\" & MyName, "D:\CHANDONPDF\CUENTAS\" & Mid(MyName, 1, 6) & "\" & MyName
             
       MyName = Dir()   ' Get next entry.
    Loop

End Sub

Private Sub Command19_Click()

'    Dim rs As New ADODB.Recordset
'    Dim Sql As String
'
'
'
'
'
'    rs.Open Sql, strConBasa
'
'
'
'Dim Paso As String
'Dim Sql As String
'Dim MyName As String
'
'
'    MyName = Dir("D:\CHANDONPDF\*.pdf")   ' Retrieve the first entry.
'
'    Do While MyName <> ""   ' Start the loop.
'
'
'        Sql = " SELECT     LETRA_DESDE"
'        Sql = Sql & " From LEGAJOS "
'        Sql = Sql & " Where (COD_CLIENTE = 1156) "
'        Sql = Sql & " And (FK_INDICES = 10715)"
'        Sql = Sql & " GROUP BY LETRA_DESDE"
'        Sql = Sql & " ORDER BY LETRA_DESDE"
'
'
'
'              ExecutarSql Sql
'
'       MyName = Dir()   ' Get next entry.
'    Loop





End Sub

Private Sub Command2_Click()
'MyName = Dir(MyPath, vbDirectory)   ' Retrieve the first entry.
'Do While MyName <> ""   ' Start the loop.
'      ' Use bitwise comparison to make sure MyName is a directory.
'      If (GetAttr(MyPath & MyName) And vbDirectory) = vbDirectory Then
'         ' Display entry only if it's a directory.directo
'         MsgBox (MyName)
'      End If
'   MyName = Dir()   ' Get next entry.
'Loop
End Sub

Private Sub Command20_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Legajo As String




    Sql = " SELECT     ID,NRO_DESDE2, LEN(NRO_DESDE2) AS Expr1, Suspense_File, BatchPgDta"
    Sql = Sql & " FROM         TELEFORM_DIGITAL"
    Sql = Sql & " Where (BatchNo = 1936) And (Len(NRO_DESDE2) = 8)"
    Sql = Sql & " ORDER BY NRO_DESDE2"
    
rs.Open Sql, strConBasa
    
    
   Do While Not rs.EOF
        Legajo = Mid(rs!NRO_DESDE2, 1, 7)
        Sql = " UPDATE    TELEFORM_DIGITAL "
        Sql = Sql & " Set NRO_DESDE2 = " & Legajo
        Sql = Sql & " Where ID = " & rs!ID
        ExecutarSql Sql
        rs.MoveNext
     Loop
    
    
    
    
End Sub

Private Sub Command21_Click()
Dim Sql As String
Dim LAGO As Long

Dim rs As New ADODB.Recordset

        Sql = " SELECT  [ID]"
        Sql = Sql & "   ,[ID_LEGAJO]"
        Sql = Sql & "   ,[NRO_DESDE]"
        Sql = Sql & "   ,[LETRA_DESDE]"
        Sql = Sql & "   ,[BARRA]"
        Sql = Sql & "   ,[BatchPgDta]"
        Sql = Sql & "   ,[TAMAÑIO]"
        Sql = Sql & "   From [basasql].[dbo].[LegajosMunicipalidadGodoyCruz]"

rs.Open Sql, strConBasa

Do While Not rs.EOF
        LAGO = FileLen("I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\DIGITALIZACION\UNa\" & Mid(rs!BatchPgDta, 1, 10))
        Sql = " Update basasql.dbo.LegajosMunicipalidadGodoyCruz"
        Sql = Sql & "  Set TAMAÑIO = " & LAGO
        Sql = Sql & "  Where ID = " & rs!ID
        ExecutarSql Sql
        rs.MoveNext
Loop




End Sub

Private Sub Command22_Click()


Dim Sql As String
Dim LAGO As Long

Dim rs As New ADODB.Recordset

Dim RS1 As New ADODB.Recordset

        
        
         Sql = " SELECT     ID_LEGAJO, COUNT(*) AS Expr1"
Sql = Sql & "   From LegajosMunicipalidadGodoyCruz"
Sql = Sql & "   GROUP BY ID_LEGAJO"
Sql = Sql & "   HAVING      (COUNT(*) = 4)"
        

rs.Open Sql, strConBasa

Do While Not rs.EOF
        
        Sql = " SELECT     ID_LEGAJO AS Expr1, TAMAÑIO, ID"
        Sql = Sql & "   From LegajosMunicipalidadGodoyCruz"
        Sql = Sql & "   Where ID_LEGAJO =  " & rs!ID_LEGAJO
                      
        Sql = Sql & "   ORDER BY TAMAÑIO DESC "
        
        Set RS1 = New ADODB.Recordset
        
        RS1.Open Sql, strConBasa
        RS1.MoveNext
              
        Sql = " Update basasql.dbo.LegajosMunicipalidadGodoyCruz"
        Sql = Sql & "  Set activo = 0"
        Sql = Sql & "  Where ID = " & RS1!ID
        ExecutarSql Sql
        
        RS1.MoveNext
              
        Sql = " Update basasql.dbo.LegajosMunicipalidadGodoyCruz"
        Sql = Sql & "  Set activo = 0"
        Sql = Sql & "  Where ID = " & RS1!ID
        ExecutarSql Sql
        
         RS1.MoveNext
              
        Sql = " Update basasql.dbo.LegajosMunicipalidadGodoyCruz"
        Sql = Sql & "  Set activo = 0"
        Sql = Sql & "  Where ID = " & RS1!ID
        ExecutarSql Sql
        
        rs.MoveNext
Loop


End Sub

Private Sub Command23_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset




Sql = "SELECT     ID_LEGAJO , ID AS Expr2, NRO_DESDE AS Expr3, LETRA_DESDE AS Expr4, BARRA AS Expr5, BatchPgDta, TAMAÑIO AS Expr7,"
Sql = Sql & " activo As Expr8"
Sql = Sql & " From LegajosMunicipalidadGodoyCruz"
Sql = Sql & " Where (activo = 0)"


rs.Open Sql, strConBasa


Do While Not rs.EOF
    FileCopy "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\DIGITALIZACION\UNa\" & Mid(rs!BatchPgDta, 1, 10), "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\DIGITALIZACION\duplicadas\" & rs!ID_LEGAJO & " _ " & Mid(rs!BatchPgDta, 1, 10)
    Kill "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\DIGITALIZACION\UNa\" & Mid(rs!BatchPgDta, 1, 10)


    rs.MoveNext
Loop





End Sub

Private Sub Command24_Click()


Dim Sql As String
Dim Paso As String
Dim rs As New ADODB.Recordset
Sql = " SELECT     LEGAJOS.NRO_DESDE, PERSONALGODOYCRUZ.NOMBRE"
Sql = Sql & " FROM TELEFORM_DIGITAL INNER JOIN"
Sql = Sql & " LEGAJOS ON TELEFORM_DIGITAL.NRO_DESDE2 = LEGAJOS.ID_LEGAJO INNER JOIN"
Sql = Sql & " PERSONALGODOYCRUZ ON LEGAJOS.NRO_DESDE = PERSONALGODOYCRUZ.NRO_DOCUMENTO"
Sql = Sql & " Where (TELEFORM_DIGITAL.BatchNo = 1936) And (LEGAJOS.COD_CLIENTE = 1156)"
Sql = Sql & " GROUP BY LEGAJOS.NRO_DESDE, PERSONALGODOYCRUZ.NOMBRE"

rs.Open Sql, strConBasa

Paso = "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\PARA ENTREGAR\DOCUMENTOS\"

Do While Not rs.EOF
    MkDir Paso & Format(rs!NRO_DESDE, "00000000") & "  " & rs!Nombre
    rs.MoveNext
Loop




End Sub

Private Sub Command25_Click()
Dim Sql As String

Dim rs As New ADODB.Recordset
Dim largo As Long

Sql = " SELECT [ID]   ,[BatchPgDta]"
Sql = Sql & "      ,[NRO_DESDE2]      ,[NRO_DESDE]"
 Sql = Sql & "     ,[NRO_HASTA]     ,[LETRA_DESDE]"
 Sql = Sql & "      ,[LETRA_HASTA]      ,[DESCRIPCION]       ,[TAMAÑO]"
 Sql = Sql & "  From [basasql].[dbo].[TELEFORM_DIGITAL_MUNI]"
 
 rs.Open Sql, strConBasa
 
 
 Do While Not rs.EOF
 
 If Not IsNull(rs!BatchPgDta) Then
   largo = FileLen("I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\DIGITALIZACION\completas\" & Mid(rs!BatchPgDta, 1, 10))
   Sql = " Update TELEFORM_DIGITAL_MUNI"
   Sql = Sql & "  Set TAMAÑO = " & largo
   Sql = Sql & "  Where ID = " & rs!ID
   ExecutarSql Sql
   End If
   
   rs.MoveNext
 Loop
 
 
 
 
 
 
 
End Sub

Private Sub Command26_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset



Sql = " SELECT     NRO_DESDE2, COUNT(*) AS CANT"
Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL_MUNI"
Sql = Sql & "  GROUP BY NRO_DESDE2"
Sql = Sql & "  ORDER BY CANT"

rs.Open Sql, strConBasa

Do While Not rs.EOF

   Sql = "  Update TELEFORM_DIGITAL_MUNI"
Sql = Sql & "  Set cantidad = " & rs!cant
Sql = Sql & "  Where NRO_DESDE2 = " & rs!NRO_DESDE2

    ExecutarSql Sql


    rs.MoveNext
Loop




End Sub

Private Sub Command27_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
Dim Paso As String
        Sql = " SELECT LEGAJOS.ID_LEGAJO, TELEFORM_DIGITAL_MUNI.BatchPgDta,"
        Sql = Sql & " TELEFORM_DIGITAL_MUNI.TAMAÑO , PERSONALGODOYCRUZ.Nro_documento, "
        Sql = Sql & " PERSONALGODOYCRUZ.Nombre , TELEFORM_DIGITAL_MUNI.cantidad"
        Sql = Sql & " FROM TELEFORM_DIGITAL_MUNI INNER JOIN"
        Sql = Sql & " LEGAJOS ON TELEFORM_DIGITAL_MUNI.NRO_DESDE2 = LEGAJOS.ID_LEGAJO INNER JOIN"
        Sql = Sql & " PERSONALGODOYCRUZ ON LEGAJOS.NRO_DESDE = PERSONALGODOYCRUZ.NRO_DOCUMENTO"
        Sql = Sql & " Where (TELEFORM_DIGITAL_MUNI.cantidad = 4 )"
        Sql = Sql & " ORDER BY LEGAJOS.ID_LEGAJO, TELEFORM_DIGITAL_MUNI.TAMAÑO DESC"
        
        rs.Open Sql, strConBasa
Paso = "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\DIGITALIZACION\completas\"
Do While Not rs.EOF
    If Dir(Paso & Mid(rs!BatchPgDta, 1, 10)) <> "" Then
    FileCopy Paso & Mid(rs!BatchPgDta, 1, 10), "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\PARA ENTREGAR\ARCHIVOS\" & Format(rs!Nro_documento, "0000000000") & "  " & rs!Nombre & ".TIF"
    Kill Paso & Mid(rs!BatchPgDta, 1, 10)
     End If
     
    
    rs.MoveNext
    
    If Dir(Paso & Mid(rs!BatchPgDta, 1, 10)) <> "" Then
    FileCopy Paso & Mid(rs!BatchPgDta, 1, 10), "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\PARA ENTREGAR\Duplicados\" & Format(rs!Nro_documento, "0000000000") & "  " & rs!Nombre & "_2.TIF"
    Kill Paso & Mid(rs!BatchPgDta, 1, 10)
     End If
     
    
    rs.MoveNext
    
    
    If Dir(Paso & Mid(rs!BatchPgDta, 1, 10)) <> "" Then
    FileCopy Paso & Mid(rs!BatchPgDta, 1, 10), "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\PARA ENTREGAR\Duplicados\" & Format(rs!Nro_documento, "0000000000") & "  " & rs!Nombre & "_3.TIF"
    Kill Paso & Mid(rs!BatchPgDta, 1, 10)
     End If
     
    
    rs.MoveNext
    
    If Dir(Paso & Mid(rs!BatchPgDta, 1, 10)) <> "" Then
    FileCopy Paso & Mid(rs!BatchPgDta, 1, 10), "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\PARA ENTREGAR\Duplicados\" & Format(rs!Nro_documento, "0000000000") & "  " & rs!Nombre & "_4.TIF"
    Kill Paso & Mid(rs!BatchPgDta, 1, 10)
     End If
     
    
    rs.MoveNext
    

Loop


End Sub

Private Sub Command28_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset

Dim sArchivo As String

sArchivo = Dir("I:\1051-MEDIFE\EN PROCESO\1097276\*.pdf")
sArchivo = Dir("I:\1051-MEDIFE\COBRADO AL 31-03-2016\1110245\*.pdf")

Do While sArchivo <> ""
        Sql = " SELECT     ID_LEGAJO, FK_INDICES, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, ETIQUETA"
        Sql = Sql & "  From LEGAJOS "
        Sql = Sql & "  Where ETIQUETA = '" & Mid(sArchivo, 1, 12) & "'"
        Set rs = New ADODB.Recordset
        rs.Open Sql, strConBasa
        If Not rs.EOF Then
             Rem FileCopy "I:\1051-MEDIFE\EN PROCESO\1097276\" & sArchivo, "I:\1051-MEDIFE\PARA ENTREGAR\" & Format(RS!ID_LEGAJO, "00000000") & "_" & Format(RS!NRO_DESDE, "0000000") & "  " & Trim(RS!LETRA_DESDE) & ".PDF"
             FileCopy "I:\1051-MEDIFE\EN PROCESO\1097276\" & sArchivo, "I:\1051-MEDIFE\COBRADO AL 31-03-2016\1110245\" & Format(rs!ID_LEGAJO, "00000000") & "_" & Format(rs!NRO_DESDE, "0000000") & "  " & Trim(rs!LETRA_DESDE) & ".PDF"
        End If

     sArchivo = Dir
Loop


End Sub

'Private Sub Command29_Click()
'Dim paso As String
'Dim Dir_Caja(200) As String
'Dim Dir_Etiqueta(200) As String
'Dim sArchivo As String
'Dim c As Integer
'Dim e As Integer
'Dim sql As String
'paso = "I:\1051-MEDIFE\DIGITALIZADAS\"
'
'
'sArchivo = Dir(paso & "*", vbDirectory)
'
'sql = "DELETE FROM basasql.dbo.DIR_ETIQUETA"
'ExecutarSql sql
'
'c = 1
'Do While sArchivo <> ""
'
'         If Len(sArchivo) > 3 Then
'            Dir_Caja(c) = sArchivo
'            c = c + 1
'        End If
'        sArchivo = Dir
'Loop
'
'For c = 1 To c
'
'   paso = "I:\1051-MEDIFE\DIGITALIZADAS\" & Dir_Caja(c)
'
'
'    sArchivo = Dir(paso & "\*", vbDirectory)
'    e = 1
'        Do While sArchivo <> ""
'
'                 If Len(sArchivo) > 3 And Dir_Caja(c) <> "" Then
'                    Dir_Etiqueta(e) = sArchivo
'                    e = e + 1
'                    sql = "INSERT INTO basasql.dbo.DIR_ETIQUETA"
'                    sql = sql & " ( ETIQUETA , CAJA )"
'                    sql = sql & "  VALUES( " & sArchivo & "," & Dir_Caja(c) & ")"
'                    ExecutarSql sql
'
'                End If
'                sArchivo = Dir
'        Loop
'
'Next
'
'
'
'End Sub

Private Sub Command3_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String

Sql = " SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO, NRO_REMITO"
Sql = Sql & " FROM         REMITOS_CUERPO INNER JOIN"
Sql = Sql & "                       DOCUMENTOS_DIGITALES ON REMITOS_CUERPO.NRO_REMITO = DOCUMENTOS_DIGITALES.NRO_DESDE"
Sql = Sql & "  WHERE     (REMITOS_CUERPO.ID_CLIENTE = 115) AND (REMITOS_CUERPO.FECHA > CONVERT(DATETIME, '2008-01-01 00:00:00', 102)) AND"
                      Sql = Sql & "  (REMITOS_CUERPO.TIPO = 2)  AND (DOCUMENTOS_DIGITALES.COD_CLIENTE = 83)"
Sql = Sql & "  ORDER BY REMITOS_CUERPO.TIPO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.ESTADO"

'sql = "  SELECT     COD_CLIENTE AS Expr1, ID, NRO_DESDE , DIRECTORIO_PASO"
'sql = sql & " From DOCUMENTOS_DIGITALES"
'sql = sql & "  WHERE     (COD_CLIENTE = 83) AND (NRO_DESDE IN (22843, 23014, 23261, 23361, 23361, 23237, 23409, 23380, 23615, 23723, 23857, 23926, 23944, 24027,"
' sql = sql & " 24202, 24208, 24126, 24042, 23852, 24186, 24185, 24280, 24282, 24244, 24588, 24738, 25044, 25367, 25559, 25587, 25759))"
'sql = sql & " ORDER BY ID"




rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
    
        FileSystem.FileCopy PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", txtPasoImagenesFinal & rs!NRO_REMITO & ".tif"

    rs.MoveNext
Loop

End Sub

Private Sub Command30_Click()
    Dim Paso As String
    Dim Dir_Caja(200) As String
    Dim Dir_Etiqueta(200) As String
    Dim sArchivo As String
    Dim C As Integer
    Dim e As Integer
    Dim Sql As String
            Paso = "I:\1156-GODOY CRUZ\0101- LEGAJOS DE PERSONAL\PARA ENTREGAR\control\ENTREGA FINAL\"
            
            
            sArchivo = Dir(Paso & "*", vbDirectory)
            
            Sql = "DELETE FROM basasql.dbo.DIR_ETIQUETA"
            ExecutarSql Sql
            
            C = 1
            Do While sArchivo <> ""
                   If Len(sArchivo) > 3 Then
'                        Dir_Etiqueta(e) = sArchivo
'                        e = e + 1
                        Sql = "INSERT INTO basasql.dbo.DIR_ETIQUETA"
                        Sql = Sql & " ( ETIQUETA , CAJA )"
                        Sql = Sql & "  VALUES( " & CLng(Mid(sArchivo, 1, 10)) & ",0)"
                        ExecutarSql Sql
                    End If
                    
                    sArchivo = Dir
            Loop
            
            
End Sub

Private Sub Command31_Click()
Dim Paso As String
Dim PasoCaja As String
Dim Dir_Caja(300) As String
Dim Dir_Etiqueta(400) As String
Dim sArchivo As String
Dim C As Integer
Dim e As Integer
Dim Sql As String
Paso = "I:\1156-GODOY CRUZ\CATASTRO\CEDULAS CATASTRALES\DIGITALIZADAS\"


sArchivo = Dir(Paso & "*", vbDirectory)

Sql = "DELETE FROM basasql.dbo.DIR_CATASTRO"
ExecutarSql Sql

C = 1
Do While sArchivo <> ""

         If Len(sArchivo) > 3 Then
            Dir_Caja(C) = sArchivo
            C = C + 1
        End If
        sArchivo = Dir
Loop

For C = 1 To C

   PasoCaja = Paso & Dir_Caja(C)


    sArchivo = Dir(PasoCaja & "\*", vbDirectory)
    e = 1
        Do While sArchivo <> ""

                 If Len(sArchivo) > 3 And Dir_Caja(C) <> "" Then
                    Dir_Etiqueta(e) = sArchivo
                    e = e + 1
                    
                   Sql = " INSERT INTO basasql.dbo.DIR_CATASTRO"
                   Sql = Sql & " (DIRECTORIO"
                   Sql = Sql & " , CAJA"
                   Sql = Sql & " , ARCHIVO"
                   Sql = Sql & " )"
                   Sql = Sql & " VALUES "
                   Sql = Sql & " ('" & Dir_Caja(C) & "'"
                   Sql = Sql & " ,'" & Dir_Caja(C) & "'"
                   Sql = Sql & " ,'" & sArchivo & "'"
                   Sql = Sql & " )"
                    ExecutarSql Sql
                End If
                sArchivo = Dir
        Loop

Next



End Sub

Private Sub Command32_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim PasoInicio As String
Dim PasoFin As String


PasoInicio = "I:\1156-GODOY CRUZ\CATASTRO\CEDULAS CATASTRALES\DIGITALIZADAS\"
PasoFin = "I:\1156-GODOY CRUZ\CATASTRO\CEDULAS CATASTRALES\UNIFICADA\"
Sql = "SELECT    ID, DIRECTORIO, CAJA, ARCHIVO"
Sql = Sql & "  From basasql.dbo.DIR_CATASTRO"
Sql = Sql & "   ORDER BY DIRECTORIO, ARCHIVO"
rs.Open Sql, strConBasa
Do While Not rs.EOF
    FileCopy PasoInicio & rs!Caja & "/" & rs!Archivo, PasoFin & "/" & Mid(rs!Caja, 1, 8) & Format(rs!ID, "00000000") & ".TIF"
    rs.MoveNext
Loop


End Sub

Private Sub Command33_Click()


 Dim PasoOriginal As String
 Dim PasoDir As String
 Dim Nombre_Archivo  As String
 
        PasoOriginal = "I:\1156-GODOY CRUZ\CATASTRO\CEDULAS CATASTRALES\UNIFICADA\"
        PasoDir = Dir(PasoOriginal & "*.tif")
        
        Do While PasoDir <> ""   ' Start the loop.
             Nombre_Archivo = PasoDir
            
            ImagXpress1.InsertPage PasoOriginal & Nombre_Archivo, "D:\test\1.tif", 1
            ImagXpress1.DeletePage "D:\test\1.tif", 2
            ImagXpress1.SaveFile

            FileCopy "D:\test\1.tif", "D:\test\u\" & Nombre_Archivo
            Kill "D:\test\1.tif"
            FileCopy "D:\test\11.tif", "D:\test\1.tif"
            PasoDir = Dir()   ' Get next entry.
            
            
        Loop




End Sub

Private Sub Command34_Click()
Dim PasoTiff As String
Dim PasoPDF As String
Dim Archivo As String
Dim sArchivo As String
Dim Sql As String


PasoTiff = "I:\1156-GODOY CRUZ\0010- LEGAJOS DE PERSONAL\PARA ENTREGAR\control\ENTREGA FINAL\"
PasoPDF = "\\Pcmorfeo2-pc\d\Final Ocr\"


'sArchivo = FileSystem.Dir(PasoTiff & "*", vbDirectory)
'Do While sArchivo <> ""
'     If Len(sArchivo) > 5 Then
'            Archivo = Mid(sArchivo, 1, Len(sArchivo) - 4) & ".pdf"
'            sql = "INSERT INTO DIR"
'            sql = sql & "("
'            sql = sql & " PASO_COMPLETO"
'            sql = sql & ", ARCHIVO"
'            sql = sql & ")"
'            sql = sql & " VALUES("
'            sql = sql & "'" & PasoTiff & sArchivo & "'"
'            sql = sql & ",'" & Archivo & "')"
'            ExecutarSql sql
'        End If
'
'      sArchivo = Dir
'Loop

Dim rs As New ADODB.Recordset

'sql = " SELECT     ID, PASO_COMPLETO, ARCHIVO, DIRECTORIO"
'sql = sql & " From Dir"
'
'
'rs.Open sql, strConBasa
'
'
'Do While Not rs.EOF
'    If Dir(PasoPDF & rs!Archivo) <> "" Then
'        sql = " Update Dir Set directorio ='SI' Where ID = " & rs!ID
'    Else
'        sql = " Update Dir Set directorio ='NO' Where ID = " & rs!ID
'    End If
'ExecutarSql sql
'
'    rs.MoveNext
'
'    Loop

Sql = " SELECT     DIRECTORIO, ID, PASO_COMPLETO, ARCHIVO"
Sql = Sql & "  From Dir"
Sql = Sql & "  WHERE     (DIRECTORIO = N'NO')"

rs.Open Sql, strConBasa
Dim pasofilan As String

pasofilan = "I:\1156-GODOY CRUZ\0010- LEGAJOS DE PERSONAL\PARA ENTREGAR\control\pasar\"



Do While Not rs.EOF
FileCopy rs!PASO_COMPLETO, pasofilan & Mid(rs!Archivo, 1, Len(rs!Archivo) - 4) & ".TIF"

    rs.MoveNext
Loop



End Sub

Private Sub Command35_Click()

Dim PasoTiff As String
Dim PasoPDF As String
Dim Archivo As String
Dim sArchivo As String
Dim Sql As String


PasoTiff = "I:\1156-GODOY CRUZ\0010- LEGAJOS DE PERSONAL\PARA ENTREGAR\control\ENTREGA FINAL\"
PasoPDF = "\\Pcmorfeo2-pc\d\Final Ocr\"


sArchivo = FileSystem.Dir("C:\Lectura\*.TXT")
Do While sArchivo <> ""
     FileCopy "C:\Lectura\" & sArchivo, "C:\Lectura\txt\" & Mid(sArchivo, 1, Len(sArchivo) - 4) & ".txt"
      
      sArchivo = Dir
Loop


End Sub


Private Sub Command36_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset

Sql = " SELECT DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_INDICES,LETRA_DESDE,  DOCUMENTOS_DIGITALES.NRO_DESDE,DOCUMENTOS_DIGITALES.NRO_HASTA, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.FK_ID_CATASTRO"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES INNER JOIN DOCUMENTOS_DIGITALES_LOTE ON"
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10738) AND (DOCUMENTOS_DIGITALES.FK_ID_CATASTRO IS NULL)"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID"


Sql = " SELECT        DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_INDICES, DOCUMENTOS_DIGITALES.LETRA_DESDE,"
Sql = Sql & "                          DOCUMENTOS_DIGITALES.NRO_HASTA, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.FK_ID_CATASTRO,"
Sql = Sql & "                          DOCUMENTOS_DIGITALES.LETRA_HASTA , DOCUMENTOS_DIGITALES.Descripcion"
Sql = Sql & "  FROM            DOCUMENTOS_DIGITALES INNER JOIN"
 Sql = Sql & "                         DOCUMENTOS_DIGITALES_LOTE ON"
 Sql = Sql & "                         DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  Where (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10738) And (DOCUMENTOS_DIGITALES.FK_ID_CATASTRO Is Null)"
Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES.ID"



    
    rs.CursorLocation = adUseClient
    rs.Open Sql, strConBasa, adOpenStatic, adLockBatchOptimistic
    
      
   Do While Not rs.EOF
   
   
   
   
                
'        If Not IsNull(RS!NRO_DESDE) And Not IsNull(RS!NRO_HASTA) Then
'
'             SQL = "SELECT ID_FICHAS, PADRON, NOMENCLATURA"
'             SQL = SQL & " From MUNIGODOYCRUZCATASTRO "
'             SQL = SQL & " WHERE PADRON = " & RS!NRO_HASTA
'             SQL = SQL & " AND NOMENCLATURA LIKE '%" & RS!NRO_DESDE & "%'"
'
'
'
'
'            SQL = "SELECT ID_FICHAS, PADRON,  NOMENCLATURA, TITULAR, CALLE, NROCALLE, BARRIO "
'             SQL = SQL & " From MUNIGODOYCRUZCATASTRO "
'             SQL = SQL & " WHERE PADRON = " & RS!NRO_HASTA
'            Rem SQL = SQL & " AND nrocalle LIKE '" & Trim(RS!LETRA_DESDE) & "'"
'
'
'
'            Set rs2 = New ADODB.Recordset
'            rs2.CursorLocation = adUseClient
'             rs2.Open SQL, strConBasa, adOpenStatic, adLockBatchOptimistic
'
'
'             If Not IsNull(RS!NRO_DESDE) Then
''                 If CLng(RS!NRO_HASTA) <> 0 Then
'                     If Not rs2.EOF Then
''                         SQL = " Update DOCUMENTOS_DIGITALES"
''                         SQL = SQL & " Set FK_ID_CATASTRO = " & rs2!ID_FICHAS
''                         SQL = SQL & " Where ID = " & RS!ID
'
'                         SQL = " Update DOCUMENTOS_DIGITALES"
'                         SQL = SQL & " Set Letra_hasta  = '" & Format(rs2!ID_FICHAS, "0000000") & " &P&: " & rs2!padron & "  &N&:" & Trim(rs2!NOMENCLATURA) & " CALLE: " & Trim(Trim(rs2!CALLE) & " " & Trim(rs2!BARRIO)) & "  NROC&:" & rs2!NROCALLE & "'"
'                         SQL = SQL & " Where ID = " & RS!ID
'                         ExecutarSql SQL
'                     End If
''                 End If
'             End If
'            End If
        rs.MoveNext
    Loop
    
  

End Sub

Private Sub Command37_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Caja As String
Dim ant As String

Sql = " SELECT ID, BatchPgDta "
Sql = Sql & " From TELEFORM_DIGITAL "
Sql = Sql & "  Where (BatchNo = 2198) "
rs.Open Sql, strConBasa

Do While Not rs.EOF
    Caja = "1056205"
    ant = Mid(rs!BatchPgDta, 8)
    Sql = " Update TELEFORM_DIGITAL "
    Sql = Sql & " SET BatchPgDta ='" & Caja & ant & "'"
    Sql = Sql & " Where ID = " & rs!ID
    ExecutarSql Sql
    

    rs.MoveNext
Loop


End Sub

Private Sub Command38_Click()
'Update LEGAJOS
'Set REGISTRO_VERIFICADO = 1
'Where (ID_LEGAJO = 1524)
'
'
Dim Sql As String
Dim rs As New ADODB.Recordset

Dim NombreArchivo As String

Sql = "SELECT     LEGAJOS.ID_LEGAJO, INDICES.DESCRIPCION, LEGAJOS.FK_INDICES, LEGAJOS.COD_INDICE, LEGAJOS.COD_CLIENTE, LEGAJOS.NRO_DESDE,"
Sql = Sql & " LEGAJOS.LETRA_DESDE ,  LEGAJOS.CONTROL_EXPORT, LEGAJOS.NRO_CAJA , ETIQUETA "
Sql = Sql & " FROM         LEGAJOS INNER JOIN"
Sql = Sql & " INDICES ON LEGAJOS.COD_INDICE = INDICES.INDICE AND LEGAJOS.COD_CLIENTE = INDICES.COD_CLIENTE"
Sql = Sql & " WHERE    "

Sql = Sql & " ORDER BY LEGAJOS.COD_INDICE"


Sql = " SELECT        DOCUMENTOS_DIGITALES.DIRECTORIO_PASO, DOCUMENTOS_DIGITALES.ID"
Sql = Sql & "  FROM            DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE ON"
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  WHERE        (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1156) AND (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10843) AND"
Sql = Sql & "          (DOCUMENTOS_DIGITALES.DESCRIPCION = '7777777777')"


Sql = " SELECT        DOCUMENTOS_DIGITALES.DIRECTORIO_PASO, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_DESDE"
Sql = Sql & " FROM            DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & "                          DOCUMENTOS_DIGITALES_LOTE ON"
Sql = Sql & "                          DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE        (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1156) AND (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10843) AND"
Sql = Sql & "                          (DOCUMENTOS_DIGITALES.NRO_DESDE < 94310)"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE"


rs.Open Sql, strConBasa

    Do While Not rs.EOF
        NombreArchivo = Format(rs!NRO_DESDE, "0000000000") & "_" & Format(rs!ID, "0000000000")
             FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", "C:\Catastro Viejo\" & NombreArchivo & ".tif"
             
        rs.MoveNext
    Loop



End Sub

Private Sub Command39_Click()

    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim NombreArchivo  As String
        
        Sql = "SELECT DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_INDICES, DOCUMENTOS_DIGITALES.NRO_DESDE,"
        Sql = Sql & "   DOCUMENTOS_DIGITALES.NRO_HASTA, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.FK_ID_CATASTRO, MUNIGODOYCRUZCATASTRO.PADRON,"
        Sql = Sql & "   MUNIGODOYCRUZCATASTRO.NOMENCLATURA, MUNIGODOYCRUZCATASTRO.TITULAR, MUNIGODOYCRUZCATASTRO.CALLE,"
        Sql = Sql & "   MUNIGODOYCRUZCATASTRO.NROCALLE , MUNIGODOYCRUZCATASTRO.BARRIO ,  DOCUMENTOS_DIGITALES.DIRECTORIO_PASO "
        Sql = Sql & "   FROM DOCUMENTOS_DIGITALES INNER JOIN"
        Sql = Sql & "   DOCUMENTOS_DIGITALES_LOTE ON"
        Sql = Sql & "   DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & "   MUNIGODOYCRUZCATASTRO ON DOCUMENTOS_DIGITALES.FK_ID_CATASTRO = MUNIGODOYCRUZCATASTRO.ID_FICHAS"
        Sql = Sql & "   Where (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10738) And (Not (DOCUMENTOS_DIGITALES.FK_ID_CATASTRO Is Null))"
        Sql = Sql & "   ORDER BY DOCUMENTOS_DIGITALES.ID"



Sql = " SELECT        DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_INDICES, DOCUMENTOS_DIGITALES.LETRA_DESDE,"
Sql = Sql & "                          DOCUMENTOS_DIGITALES.NRO_HASTA, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.FK_ID_CATASTRO,"
Sql = Sql & "                          DOCUMENTOS_DIGITALES.LETRA_HASTA , DOCUMENTOS_DIGITALES.Descripcion , DIRECTORIO_PASO  "
Sql = Sql & "  FROM            DOCUMENTOS_DIGITALES INNER JOIN"
 Sql = Sql & "                         DOCUMENTOS_DIGITALES_LOTE ON"
 Sql = Sql & "                         DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  Where (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10738) And (DOCUMENTOS_DIGITALES.FK_ID_CATASTRO Is Null)"
Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES.ID"


        Set rs = New ADODB.Recordset
        
        
        rs.CursorLocation = adUseClient

      
        rs.Open Sql, strConBasa, adOpenKeyset, adLockOptimistic
        
        Rem 0031735 &P&: 30596  &N&:BD578200 CALLE: LAUTARO SUPE  NROC&:3529
        
        Do While Not rs.EOF
            Rem NombreArchivo = "PADRON " & Format(RS!padron, "000000") & " NOMECLATURA " & RS!NOMENCLATURA & " " & Trim(RS!titular) & " calle_ " & RS!CALLE & " " & RS!NROCALLE & " ID_" & RS!ID
                        NombreArchivo = "PADRON " & Format(rs!NRO_HASTA, "000000") & "   " & Mid(rs!LETRA_HASTA, 19)
            
            NombreArchivo = "PADRON " & Format(rs!NRO_HASTA, "000000")
            
            NombreArchivo = Replace(NombreArchivo, "&N&:", "NOMECLATURA  ")
            NombreArchivo = Replace(NombreArchivo, "NROC&:", "NRO ")
            NombreArchivo = Replace(NombreArchivo, ":", " ")
            NombreArchivo = Replace(NombreArchivo, ".", " ")
            Rem NombreArchivo = Replace(NombreArchivo, "&N&:", "NOMECLATURA")
            
            If Dir(PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif") <> "" Then
               Rem FileSystem.FileCopy PasoImagenes & RS!DIRECTORIO_PASO & "\" & RS!ID & ".tif", "D:\NOMECLATURA\" & NombreArchivo & ".TIF"
               
                              FileSystem.FileCopy PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", "D:\NOMECLATURA\1" & ".TIF"
                 
             
                 
                 FileCopy "\\222.15.19.251\ImagenesPDF" & "\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".pdf", "D:\NOMECLATURA\" & Trim(NombreArchivo) & ".pdf"
                
                Rem ExecutarSql "UPDATE DOCUMENTOS_DIGITALES Set Exportado =  '15/06/2016' Where ID = " & RS!ID
            Else
            Debug.Print PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif"
            Rem ExecutarSql "UPDATE DOCUMENTOS_DIGITALES Set Exportado =  '01/01/2000' Where ID = " & RS!ID
            End If
            
            
            
            rs.MoveNext
        Loop
        


End Sub

Private Sub Command4_Click()
    Dim Sql As String
    Dim rsImagenes As New ADODB.Recordset
   

       
        MousePointer = 11
        
        Dim NombreArchivo As String
        Dim Manzana As String
        Dim Parcela As String
        Dim SubParcela As String
        Dim Divi As String
        

        
      Sql = "  SELECT   DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE,"
         Sql = Sql & "  DOCUMENTOS_DIGITALES_LOTE.LA_CAJA_HOJA_CONTROL, DOCUMENTOS_DIGITALES_LOTE.FK_LA_CAJA_TOMADOR,"
          Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.TIPO_DOCUMENTO , DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO"
  Sql = Sql & " FROM         DOCUMENTOS_DIGITALES INNER JOIN"
                  Sql = Sql & "       DOCUMENTOS_DIGITALES_LOTE ON"
                        Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
  Sql = Sql & " WHERE     (DOCUMENTOS_DIGITALES_LOTE.FK_LA_CAJA_TOMADOR = 325) AND (DOCUMENTOS_DIGITALES_LOTE.TIPO_DOCUMENTO = 'SOLICITUD') AND"
                  Sql = Sql & "       (DOCUMENTOS_DIGITALES.NRO_DESDE > 10)"
  Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE"
        

ConBasa.CommandTimeout = 300
        
        
        Set rsImagenes = New ADODB.Recordset
        
        rsImagenes.Open Sql, ConActiva, 0, 1
        
            Do While Not rsImagenes.EOF
               FileSystem.FileCopy PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif", "D:\ExportarImagenes\" & rsImagenes!NRO_DESDE & "_" & rsImagenes!ID & ".tif"
               rsImagenes.MoveNext
            Loop
MousePointer = 0
MsgBox "Operacion terminada"
End Sub

Private Sub Command40_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim ID As String

Sql = " SELECT NRO_DESDE2, BatchPgDta"
Sql = Sql & " From TELEFORM_DIGITAL"
Sql = Sql & " WHERE BatchNo IN (2490, 2494)"
Sql = Sql & " ORDER BY NRO_DESDE2"

rs.Open Sql, strConBasa

Do While Not rs.EOF
    ID = Mid(rs!BatchPgDta, 1, 7)
    Sql = " Update DOCUMENTOS_DIGITALES"
    Sql = Sql & " Set NRO_HASTA = " & rs!NRO_DESDE2
    Sql = Sql & "  Where NRO_HASTA IS NULL and id = " & ID
    ExecutarSql Sql
    rs.MoveNext
Loop


End Sub

Private Sub Command41_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim Desde As Double
    Dim Hasta As Double
     Dim i As Integer
     
    Rem 557
    
'    desde = 5570000
'    For i = 1 To 1000
'    desde = desde + 1
'    Hasta = desde + 9999
'
'        Sql = " Insert "
'        Sql = Sql & " Into DIRECTORIOS_IMAGENES(desde, Hasta, DIRECTORIO_PASO)"
'        Sql = Sql & " VALUES (" & desde & "," & Hasta & ",'" & desde & "-" & Hasta & "')"
'        ExecutarSql Sql
'     desde = Hasta
'    Next
'
    
   
    
    
    Sql = " SELECT         DIRECTORIO_PASO"
    Sql = Sql & " From DIRECTORIOS_IMAGENES"
    Sql = Sql & "  Where ID > 557"
     Sql = Sql & " ORDER BY HASTA"
    
    rs.Open Sql, strConBasa
    
    Do While Not rs.EOF
      FileSystem.MkDir "\\222.15.19.251\ImagenesSimpleTif\" & Trim(rs!DIRECTORIO_PASO)
      FileSystem.MkDir "\\222.15.19.251\Imagenes\" & Trim(rs!DIRECTORIO_PASO)
      FileSystem.MkDir "\\222.15.19.251\ImagenesPDF\" & Trim(rs!DIRECTORIO_PASO)
      rs.MoveNext
    Loop



End Sub

Private Sub Command42_Click()
Dim Sql As String

Dim PasoInicio As String
Dim PasoFin As String
Dim sArchivo As String
Dim rs As New ADODB.Recordset
Dim ID  As Long
PasoInicio = "I:\1156-GODOY CRUZ\1-RRHH\ACTUALIZACIONES DE PERSONAL\para exportar easp\"


PasoFin = "I:\1156-GODOY CRUZ\1-RRHH\ACTUALIZACIONES DE PERSONAL\para exportar easp\Transformado\"

sArchivo = Dir(PasoInicio & "*.PDF", vbDirectory)
Do While sArchivo <> ""
      ID = 0
        If Len(sArchivo) = 16 Then
            ID = Mid(sArchivo, 6, 7)
            Else
           
        End If
        If Len(sArchivo) = 11 Then
            ID = Mid(sArchivo, 1, 7)
        End If
        
        Set rs = New ADODB.Recordset
            Sql = " SELECT LEGAJOS.ID_LEGAJO, LEGAJOS.NRO_HASTA, LEGAJOS.FECHA_DESDE, LEGAJOS.FECHA_HASTA, LEGAJOS.DESCRIPCION, LEGAJOS.NRO_CAJA,"
            Sql = Sql & vbCrLf & " LEGAJOS.COD_ESTADO, INDICES.ID_CODIGO_DOCUMENTO, LEGAJOS.NRO_DESDE, LEGAJOS.LETRA_HASTA, LEGAJOS.LETRA_DESDE, INDICES.INDICE,"
            Sql = Sql & vbCrLf & " LEGAJOS.Etiqueta , LEGAJOS.UNIFICACION_ID_LEGAJOS"
            Sql = Sql & vbCrLf & " FROM LEGAJOS INNER JOIN INDICES ON LEGAJOS.FK_INDICES = INDICES.ID"
            Sql = Sql & vbCrLf & " WHERE LEGAJOS.COD_INDICE LIKE '001001%' AND  LEGAJOS.COD_CLIENTE = 1156"
            Sql = Sql & vbCrLf & " AND LEGAJOS.ID_LEGAJO = " & ID
            rs.Open Sql, strConBasa
            If Not rs.EOF Then
                   FileSystem.FileCopy PasoInicio & sArchivo, PasoFin & rs!Etiqueta & ".PDF"
               Rem  Kill PasoInicio & sArchivo
'Sql = " Update LEGAJOS"
'Sql = Sql & " Set REGISTRO_VERIFICADO = 1"
'Sql = Sql & "  Where ID_LEGAJO = " & RS!ID_LEGAJO
' ExecutarSql Sql

                
              End If
              
            
        
        
        
        sArchivo = Dir
Loop






End Sub

Private Sub Command43_Click()


'Dim SQL As String
'Dim RS As New ADODB.Recordset
'Dim cantidad As Integer
'
'Dim NRO_DESDE As Long
'Dim IDMAX As Long
'Dim FK_DOCUMENTOS_DIGITALES_LOTE As Long
'Dim PasoOrigen As String
'Dim IMAGEN_ORIGEN As String
'Dim FECHA_INCORPORACION As String
'Dim CANTIDAD_IMAGENES As Integer
'Dim estado As String
'Dim DIRECTORIO_PASO As String
'Dim i As Integer
'
''
''SQL = " SELECT DOCUMENTOS_DIGITALES.NRO_DESDE AS NRO_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
''SQL = SQL & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE , DOCUMENTOS_DIGITALES.DIRECTORIO_PASO "
''SQL = SQL & "  FROM            DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
''SQL = SQL & "  DOCUMENTOS_DIGITALES ON"
''SQL = SQL & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
''SQL = SQL & "  WHERE        (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES IN (10844)) AND (DOCUMENTOS_DIGITALES.NRO_DESDE > 9999999990)"
''SQL = SQL & "  ORDER BY NRO_DESDE"
'
'
'
'SQL = " SELECT DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
'SQL = SQL & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES.PASOORIGEN, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN,"
'SQL = SQL & vbCrLf & " DOCUMENTOS_DIGITALES.FECHA_INCORPORACION, DOCUMENTOS_DIGITALES.CANTIDAD_IMAGENES, DOCUMENTOS_DIGITALES.ESTADO,"
'SQL = SQL & vbCrLf & " DOCUMENTOS_DIGITALES.DIRECTORIO_PASO"
'SQL = SQL & vbCrLf & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON"
'SQL = SQL & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
'SQL = SQL & vbCrLf & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_INDICES IN (10844) AND (DOCUMENTOS_DIGITALES.NRO_DESDE > 9999999990)"
'SQL = SQL & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE"
'
'
'RS.Open SQL, ConBasa
'
'Do While Not RS.EOF
'   cantidad = Mid(RS!NRO_DESDE, 10)
'
'
'
'
'
'
'   For i = 1 To cantidad
'
'
'
'   If i = 1 Then
'        SQL = "Update DOCUMENTOS_DIGITALES"
'        SQL = SQL & vbCrLf & " SET "
'        SQL = SQL & vbCrLf & " NRO_DESDE =" & 1111111111
'        SQL = SQL & vbCrLf & " , ESTADO ='VERIFICAR MANUALMENTE' "
'        SQL = SQL & vbCrLf & " Where ID = " & RS!ID
'        ExecutarSql SQL
'   Else
'
'    NRO_DESDE = 1111111110 + i
'    IDMAX = MAX_DOCUMENTOS_DIGITALES_2()
'    FK_DOCUMENTOS_DIGITALES_LOTE = RS!FK_DOCUMENTOS_DIGITALES_LOTE
'    PasoOrigen = "'" & Trim(RS!PasoOrigen) & "'"
'    IMAGEN_ORIGEN = "'" & Trim(RS!IMAGEN_ORIGEN) & "'"
'    FECHA_INCORPORACION = "'" & RS!FECHA_INCORPORACION & "'"
'    CANTIDAD_IMAGENES = RS!CANTIDAD_IMAGENES
'    estado = "'VERIFICAR MANUALMENTE'"
'    DIRECTORIO_PASO = BuscarDirectorioPaso(IDMAX)
'
'    SQL = " INSERT INTO DOCUMENTOS_DIGITALES ( "
'    SQL = SQL & vbCrLf & " NRO_DESDE "
'    Rem SQL = SQL & vbCrLf & " , ID"
'    SQL = SQL & vbCrLf & " , FK_DOCUMENTOS_DIGITALES_LOTE"
'    SQL = SQL & vbCrLf & " , PASOORIGEN"
'    SQL = SQL & vbCrLf & " , IMAGEN_ORIGEN"
'    SQL = SQL & vbCrLf & " , FECHA_INCORPORACION"
'    SQL = SQL & vbCrLf & " , CANTIDAD_IMAGENES"
'    SQL = SQL & vbCrLf & " , ESTADO"
'    SQL = SQL & vbCrLf & " , DIRECTORIO_PASO"
'    SQL = SQL & vbCrLf & " )"
'    SQL = SQL & vbCrLf & " VALUES ( "
'    SQL = SQL & vbCrLf & NRO_DESDE
'    Rem SQL = SQL & vbCrLf & " , " & IDMAX + 1
'    SQL = SQL & vbCrLf & " , " & FK_DOCUMENTOS_DIGITALES_LOTE
'    SQL = SQL & vbCrLf & " , " & PasoOrigen
'    SQL = SQL & vbCrLf & " , " & IMAGEN_ORIGEN
'    SQL = SQL & vbCrLf & " , " & FECHA_INCORPORACION
'    SQL = SQL & vbCrLf & " , " & CANTIDAD_IMAGENES
'    SQL = SQL & vbCrLf & " , " & estado
'    SQL = SQL & vbCrLf & " , '" & Trim(DIRECTORIO_PASO) & "'"
'    SQL = SQL & vbCrLf & " )"
'    ExecutarSql SQL
'
'    FileCopy "\\222.15.19.251\Imagenes\" & RS!DIRECTORIO_PASO & "\" & RS!ID & ".TIF", "\\222.15.19.251\Imagenes\" & DIRECTORIO_PASO & "\" & IDMAX & ".TIF"
'

'   End If
'
'Next
'
'
'
' RS.MoveNext
'Loop



'Dim Sql As String
'    Dim rs As New ADODB.Recordset
'    Dim NombreArchivo  As String
'        Sql = " SELECT     DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_INDICES, DOCUMENTOS_DIGITALES.LETRA_DESDE, "
'        Sql = Sql & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.FK_ID_CATASTRO, "
'        Sql = Sql & " DOCUMENTOS_DIGITALES.LETRA_HASTA , DOCUMENTOS_DIGITALES.Descripcion , DIRECTORIO_PASO "
'        Sql = Sql & " FROM DOCUMENTOS_DIGITALES INNER JOIN "
'        Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE ON "
'        Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
'        Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10844) "
'        Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID "
'        Set rs = New ADODB.Recordset
'        rs.CursorLocation = adUseClient
'        rs.Open Sql, strConBasa, adOpenKeyset, adLockOptimistic
'        Do While Not rs.EOF
'             NombreArchivo = "PADRON " & Format(rs!NRO_DESDE, "0000000") & " ID_" & rs!ID
'            If Dir(PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif") <> "" Then
'                FileSystem.FileCopy PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", "D:\FICHAS CELESTES\" & Trim(NombreArchivo) & ".TIF"
'            Else
'                Debug.Print PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif"
'            End If
'            rs.MoveNext
'        Loop
   End Sub

Private Sub Command44_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset

Sql = "SELECT DOCUMENTOS_DIGITALES_LOTE.FECHA_VERIFICADO_RECEPCION, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO,"
Sql = Sql & " DOCUMENTOS_DIGITALES.estado"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 86)"
Sql = Sql & " And (DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE < 26831)"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID"
rs.CursorLocation = adUseClient

rs.Open Sql, strConBasa, adOpenForwardOnly, adLockOptimistic


Do While Not rs.EOF

        FileSystem.FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF", "\\222.15.19.248\limpieza\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF"
        Sql = " UPDATE  DOCUMENTOS_DIGITALES SET  ESTADO ='IMAGEN BORRADA'  Where ID = " & rs!ID
        ExecutarSql Sql
        
        FileSystem.Kill "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF"
        rs.MoveNext
  Loop
  

End Sub

Private Sub Command45_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim ViejoLote As String
    Dim PasoActual As String
        Sql = " SELECT DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO ,  DOCUMENTOS_DIGITALES.ID as ID_IMAGEN , "
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, INDICES.DESCRIPCION AS DESCRIPCION_INDICE,DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"
        Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & vbCrLf & " INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
        Sql = Sql & vbCrLf & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 86) "
        Sql = Sql & vbCrLf & " AND (DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS = 1218282) "
       Rem  Sql = Sql & vbCrLf & " AND (DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION = '20160801')"
        Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"
    
    
    rs.Open Sql, strConBasa
    
    Do While Not rs.EOF
     If ViejoLote <> rs!ID_DOCUMENTOS_DIGITALES_LOTE Then
        FileSystem.MkDir "D:\OSDE\" & rs!ID_DOCUMENTOS_DIGITALES_LOTE & "_" & Trim(rs!DESCRIPCION_INDICE)
        PasoActual = "D:\OSDE\" & rs!ID_DOCUMENTOS_DIGITALES_LOTE & "_" & Trim(rs!DESCRIPCION_INDICE)
        ViejoLote = rs!ID_DOCUMENTOS_DIGITALES_LOTE
     End If
     
        FileSystem.FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID_imagen & ".TIF", PasoActual & "\" & Format(Trim(rs!IMAGEN_ORIGEN), "00000000") & ".TIF"
        rs.MoveNext
    Loop

End Sub

Private Sub Command46_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Sql = " SELECT        TOP (1000) ID, DESDE, HASTA, DIRECTORIO_PASO"
Sql = Sql & " From DIRECTORIOS_IMAGENES"
Sql = Sql & "  Where ID < 252"

rs.Open Sql, strConBasa

Do While Not rs.EOF
    FileSystem.MkDir "\\222.15.19.251\ImagenesPDF\" & Trim(rs!DIRECTORIO_PASO)
    rs.MoveNext
Loop



End Sub

Private Sub Command47_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String



Sql = " SELECT  DOCUMENTOS_DIGITALES.ID AS ID_IMAGEN, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO AS PASO, DOCUMENTOS_DIGITALES.PasoControl,"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 155)"
Sql = Sql & " ORDER BY ID_IMAGEN "


Rem /****** Script para el comando SelectTopNRows de SSMS  ******/
Sql = " SELECT  DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES.ID as ID_imagen, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO as paso"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1156) "
Sql = Sql & " And (DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = 10738)"



        rs.CursorLocation = adUseClient
        rs.Open Sql, strConBasa, adOpenForwardOnly, adLockReadOnly
        Do While Not rs.EOF
            If Dir("\\222.15.19.251\ImagenesPDF\" & rs!Paso & "\" & rs!ID_imagen & ".pdf") <> "" Then
                FileSystem.FileCopy "\\222.15.19.251\ImagenesPDF\" & rs!Paso & "\" & rs!ID_imagen & ".pdf", "C:\Cedulas Catastrales\" & rs!ID_imagen & ".pdf"
                Sql = " UPDATE  DOCUMENTOS_DIGITALES SET  PasoControl = 'NO ESTA' Where ID = " & rs!ID_imagen
                ExecutarSql Sql
            Else
                Sql = " UPDATE  DOCUMENTOS_DIGITALES SET  PasoControl = 'NO ESTA' Where ID = " & rs!ID_imagen
                ExecutarSql Sql
            End If
            rs.MoveNext
        Loop




End Sub

Private Sub Command48_Click()
Dim Sql As String
Dim i As Long
Dim R As Long
R = 70400
For i = 1 To 3000
    Sql = " Insert Top(1000)"
    Sql = Sql & " Into REMITOS_FISICOS(NRO_REMITO_FISICO, FORMATO, ID_REMITO_FISICO , FORMATOBARRA  )"
    Sql = Sql & " VALUES (" & R + i & ",'0001-000" & CStr(R + i) & "'," & R + i & ",'0001000" & CStr(R + i) & "')"
    ConBasa.Execute Sql
Next

End Sub

'Private Sub Command49_Click()
'    Dim Sql As String
'    Dim rs As New ADODB.Recordset
'    Dim rsGeneral As New ADODB.Recordset
'
'    Dim s_Directorio As String
'    Dim H_s_Directorio As String
'    Dim H_HOJA_RUTA As String
'    Dim H_TOMADOR As String
'    Dim Nombre_Archivo As String
'
'
'
'
'
'    Sql = " SELECT DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.LETRA_HASTA, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_DESDE,"
'    Sql = Sql & " DOCUMENTOS_DIGITALES.NRO_HASTA, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO, DOCUMENTOS_DIGITALES_LOTE.LOTE_ESTADO,"
'    Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
'    Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
'    Sql = Sql & " DOCUMENTOS_DIGITALES ON DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
'    Sql = Sql & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS IN (1066385, 1066388, 1066389, 1066391, 1066392, 1066395, 1066441, 1066494, 1066499, 1066502, 1066517, 1136651, 1136652, 1136654, 1139767, 1053895,"
'    Sql = Sql & " 1053896, 1171785, 936151, 1053879, 1053880, 1065480, 1064704, 1064762, 1064764, 1064765)) AND (DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN = 1)"
'    Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.LETRA_DESDE"
'
'    rsGeneral.Open Sql, strConBasa
'
'Do While rsGeneral.EOF
'
' If H_HOJA_RUTA <> rsGeneral!NRO_DESDE Then
'    H_HOJA_RUTA = rsGeneral!NRO_DESDE
'    If Dir("D\LA CAJA\" & H_HOJA_RUTA, vbDirectory) = "" Then
'        MkDir ("D:\LA CAJA\" & H_HOJA_RUTA)
'    End If
'
'    If H_TOMADOR <> rsGeneral!LETRA_DESDE Then
'        H_TOMADOR = rsGeneral!LETRA_DESDE
'        If Dir("D\LA CAJA\" & H_HOJA_RUTA & "\" & H_TOMADOR, vbDirectory) = "" Then
'            MkDir ("D\LA CAJA\" & H_HOJA_RUTA & "\" & H_TOMADOR)
'        End If
'    End If
' End If
'
'
'         Sql = " SELECT        TOP (1000) ID, FK_DOCUMENTOS_DIGITALES_LOTE, LETRA_DESDE, NRO_DESDE,   DIRECTORIO_PASO, IMAGEN_ORIGEN"
'Sql = Sql & "  From DOCUMENTOS_DIGITALES"
'Sql = Sql & "  Where FK_DOCUMENTOS_DIGITALES_LOTE = " & rsGeneral!ID_DOCUMENTOS_DIGITALES_LOTE
'Sql = Sql & "  ORDER BY ID"
'
'         Set rs = New ADODB.Recordset
'         rs.Open Sql, strConBasa, adOpenForwardOnly, adLockOptimistic
'
'         Do While Not rs.EOF
'            s_Directorio = "D\LA CAJA\" & H_HOJA_RUTA & "\" & H_TOMADOR & "\"
'
'
'            If rs!IMAGEN_ORIGEN > 1 Then
'               Nombre_Archivo = "0000 Caratula "
'            Else
'                Nombre_Archivo = rs!ID
'            End If
'
'                  FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", s_Directorio & Nombre_Archivo & ".tif"
'
'            rs.MoveNext
'        Loop
'
'End Sub

Private Sub Command5_Click()

'
'Dim Sql As String
'    Dim rsImagenes As New ADODB.Recordset
'    Dim rs  As ADODB.Recordset
'    Dim docDestino As MODI.Document
'    rsBuscar.Requery
'    MousePointer = 11
'        Do While Not rsBuscar.EOF
'            Sql = "  SELECT ID, COD_CLIENTE, LOTE, COD_ESTADO, DIRECTORIO_PASO,NRO_DESDE  "
'            Sql = Sql & " From  DOCUMENTOS_DIGITALES  "
'            Sql = Sql & "  WHERE    COD_CLIENTE = " & rsBuscar!COD_CLIENTE
'            Sql = Sql & "  AND LOTE =  '" & rsBuscar!Lote & "'"
'            Sql = Sql & "  AND COD_ESTADO = " & rsBuscar!Cod_Estado
'            Set rsImagenes = New ADODB.Recordset
'            rsImagenes.Open Sql, strConBasa , 0 ,1
'            Do While Not rsImagenes.EOF
'            If rsImagenes!NRO_DESDE <> 0 Then
'                    Sql = " SELECT     COD_CLIENTE, NRO_DESDE, ID"
'                    Sql = Sql & " From DOCUMENTOS_DIGITALES"
'                    Sql = Sql & " Where COD_CLIENTE = 163 "
'                    Sql = Sql & " And NRO_DESDE = " & rsImagenes!NRO_DESDE
'                    Sql = Sql & " And (NOTI = 1)"
'                    Set rs = New ADODB.Recordset
'                    rs.Open Sql, strConBasa , 0 ,1
'                    If Not rs.EOF Then
'
'                       Rem  MsgBox rsImagenes!NRO_DESDE
'
'                 Sql = "   Update DOCUMENTOS_DIGITALES Set IMAGEN_NOTTI = " & rs!ID & " Where ID = " & rsImagenes!ID
'
'                        ExecutarSql Sql
'
'                        Sql = "   Update DOCUMENTOS_DIGITALES Set IMAGEN_NOTTI = " & rsImagenes!ID & " Where ID = " & rs!ID
'
'                        ExecutarSql Sql
'
'                    End If
'            End If
'                rsImagenes.MoveNext
'            Loop
'       rsBuscar.MoveNext
'Loop
'MousePointer = 0
'MsgBox "Operacion terminada"



End Sub

Private Sub UnirNottiFicha()

'
'
'    Dim DocSeparador As MODI.Document
'    Dim DocFichas As MODI.Document
'    Dim DocFlyers As MODI.Document
'    Dim DocSave As MODI.Document
'
'        Dim Sql As String
'        Dim rsImagenes As ADODB.Recordset
'        Dim rsNoti As ADODB.Recordset
'
'    MousePointer = 11
'
'
'
'    Set DocSeparador = New MODI.Document
'    DocSeparador.Create "C:\registro.tif"
'        Dim i As Integer
'
'
'    If Dir(txtPasoImagenesFinal.Text & "\" & Trim(TXTnOMBREdiRECTORIO.Text), vbDirectory) = "" Then
'        FileSystem.MkDir txtPasoImagenesFinal.Text & "\" & Trim(TXTnOMBREdiRECTORIO.Text)
'    End If
'
'
'
'
'         rsBuscar.Requery
'         Do While Not rsBuscar.EOF
'            Sql = "  SELECT ID, COD_CLIENTE, LOTE, COD_ESTADO, DIRECTORIO_PASO, LETRA_DESDE,NRO_DESDE, IMAGEN_NOTTI "
'            Sql = Sql & " From  DOCUMENTOS_DIGITALES  "
'            Sql = Sql & "  WHERE    COD_CLIENTE = " & rsBuscar!COD_CLIENTE
'            Sql = Sql & "  AND LOTE =  '" & rsBuscar!Lote & "'"
'            Sql = Sql & "  AND COD_ESTADO = " & rsBuscar!Cod_Estado
'            Sql = Sql & "  order by NRO_DESDE "
'            Set rsImagenes = New ADODB.Recordset
'            rsImagenes.Open Sql, strConBasa , 0 ,1
'            Do While Not rsImagenes.EOF
'                    Set DocSeparador = New MODI.Document
'                    DocSeparador.Create "C:\registro.tif"
'                    Set DocFichas = New MODI.Document
'                    DocFichas.Create PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif"
'                    Set DocSave = New MODI.Document
'                    DocSave.Create
'                    If IsNull(rsImagenes!IMAGEN_NOTTI) Then
'                            DocSave.Images.Add DocFichas.Images.Item(0), DocFichas.Images.Item(0)
'                            DocSave.Images.Add DocSeparador.Images.Item(0), DocSeparador.Images.Item(0)
'                            DocSave.SaveAs txtPasoImagenesFinal.Text & Trim(TXTnOMBREdiRECTORIO.Text) & "\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
'                        Else
'                            Sql = "   SELECT     COD_CLIENTE, LOTE, IMAGEN_NOTTI,DIRECTORIO_PASO, ID"
'                            Sql = Sql & "  From DOCUMENTOS_DIGITALES"
'                            Sql = Sql & "  WHERE   iD= " & rsImagenes!IMAGEN_NOTTI
'                               Set rsNoti = New ADODB.Recordset
'                                rsNoti.Open Sql, strConBasa , 0 ,1
'                                If Not rsNoti.EOF Then
'
'                                  Do While Not rsNoti.EOF
'                                    Set DocFlyers = New MODI.Document
'                                    DocFlyers.Create PasoImagenes & rsNoti!DIRECTORIO_PASO & "\" & rsNoti!ID & ".tif"
'                                    For i = 0 To DocFlyers.Images.Count - 1
'                                        DocSave.Images.Add DocFlyers.Images.Item(i), DocFlyers.Images.Item(i)
'                                    Next
'
'                                    rsNoti.MoveNext
'                                   Loop
'
'                                    DocSave.Images.Add DocFichas.Images.Item(0), DocFichas.Images.Item(0)
'                                    DocSave.Images.Add DocSeparador.Images.Item(0), DocSeparador.Images.Item(0)
'                                    DocSave.SaveAs txtPasoImagenesFinal.Text & Trim(TXTnOMBREdiRECTORIO.Text) & "\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
'
'                                Else
'                                    DocSave.Images.Add DocFichas.Images.Item(0), DocFichas.Images.Item(0)
'                                    DocSave.Images.Add DocSeparador.Images.Item(0), DocSeparador.Images.Item(0)
'                                    DocSave.SaveAs txtPasoImagenesFinal.Text & Trim(TXTnOMBREdiRECTORIO.Text) & "\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
'
'                                End If
'                        End If
'                        rsImagenes.MoveNext
'                    Loop
'                rsBuscar.MoveNext
'                Loop
'
''
''
''            rsImagenes.Open SQL, strConBasa , 0 ,1
''            Set docOrigen = New MODI.Document
''            docOrigen.Create PasoImagenes & rsBuscar!DIRECTORIO_PASO & "\" & rsBuscar!ID & ".tif"
''                If Not rsImagenes.EOF Then
''                    Set docDestino = New MODI.Document
''                    docDestino.Create PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif"
''                    docDestino.Images.Add docOrigen.Images.Item(0), docDestino.Images.Item(0)
''                End If
''                docDestino.SaveAs "D:\ExportarImagenes\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
''                rsBuscar.MoveNext
''
''
''
''                Set docDestino = New MODI.Document
''                docDestino.Create PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif"
''                docDestino.Images.Add docOrigen.Images.Item(0), docDestino.Images.Item(0)
''                If rsImagenes!NRO_DESDE < 10 Or IsNull(rsImagenes!NRO_DESDE) Then
''                     docDestino.SaveAs txtPasoImagenesFinal.Text & Trim(TXTnOMBREdiRECTORIO.Text) & "\0" & "_" & rsImagenes!ID & ".TIF"
''                Else
''                     docDestino.SaveAs txtPasoImagenesFinal.Text & Trim(TXTnOMBREdiRECTORIO.Text) & "\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
''                End If
''                Rem FileCopy PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif", txtPasoImagenesFinal.Text & Trim(TXTnOMBREdiRECTORIO.Text) & "\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
''               rsImagenes.MoveNext
''            Loop
''       rsBuscar.MoveNext
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''
''        SQL = "   SELECT     COD_CLIENTE, LOTE, IMAGEN_NOTTI,DIRECTORIO_PASO, ID"
''        SQL = SQL & "  From DOCUMENTOS_DIGITALES"
''        SQL = SQL & "  WHERE     (COD_CLIENTE = 163) AND (LOTE LIKE N'5100%') AND (NOT (IMAGEN_NOTTI IS NULL))"
''        Set rsBuscar = New ADODB.Recordset
''        rsBuscar.Open SQL, strConBasa , 0 ,1
''
''        Do While Not rsBuscar.EOF
''            SQL = "  SELECT ID, COD_CLIENTE, LOTE, COD_ESTADO, DIRECTORIO_PASO, LETRA_DESDE,NRO_DESDE "
''            SQL = SQL & " From  DOCUMENTOS_DIGITALES  "
''            SQL = SQL & "  WHERE    id=" & rsBuscar!IMAGEN_NOTTI
''            Set rsImagenes = New ADODB.Recordset
''            rsImagenes.Open SQL, strConBasa , 0 ,1
''            Set docOrigen = New MODI.Document
''            docOrigen.Create PasoImagenes & rsBuscar!DIRECTORIO_PASO & "\" & rsBuscar!ID & ".tif"
''                If Not rsImagenes.EOF Then
''                    Set docDestino = New MODI.Document
''                    docDestino.Create PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif"
''                    docDestino.Images.Add docOrigen.Images.Item(0), docDestino.Images.Item(0)
''                End If
''                docDestino.SaveAs "D:\ExportarImagenes\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
''                rsBuscar.MoveNext
''        Loop
'        MousePointer = 0
'        MsgBox "Operacion terminada"
End Sub

Private Sub Command50_Click()
            Dim Sql As String
            Dim rs As New ADODB.Recordset
            Dim rslegajos As ADODB.Recordset
            Dim ID_LEGAJO As String
            
            Sql = " SELECT  DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER, DOCUMENTOS_DIGITALES_LOTE.REMITO, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE,"
            Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.CANTIDAD_IMAGENES, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.FK_LEGAJO_ETIQUETA, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
            Sql = Sql & " LEN(DOCUMENTOS_DIGITALES.FK_LEGAJO_ETIQUETA) AS Expr1, DOCUMENTOS_DIGITALES.ID"
            Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
            Sql = Sql & " DOCUMENTOS_DIGITALES ON DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
            Rem Sql = Sql & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER BETWEEN CONVERT(DATETIME, '2017-09-29 00:00:00', 102) AND CONVERT(DATETIME, '2017-10-31 00:00:00', 102)) AND"
            Sql = Sql & " WHERE "
            Sql = Sql & " (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 279) and LETRA_DESDE is null "
            
            rs.Open Sql, ConBasa
            
            Do While Not rs.EOF
                ID_LEGAJO = ""
                
                
                
            If Len(rs!FK_LEGAJO_ETIQUETA) = 12 Then
            ID_LEGAJO = Mid(rs!FK_LEGAJO_ETIQUETA, 5, 8)
            End If
            If Len(rs!FK_LEGAJO_ETIQUETA) = 13 Then
            ID_LEGAJO = Mid(rs!FK_LEGAJO_ETIQUETA, 5, 8)
            End If
            
            If ID_LEGAJO <> "" Then
            
                Set rslegajos = New ADODB.Recordset
                Sql = " SELECT LETRA_DESDE, NRO_DESDE From LEGAJOS"
                Sql = Sql & " Where ID_LEGAJO = " & ID_LEGAJO
                Sql = Sql & " And (COD_CLIENTE = 279)"
                rslegajos.Open Sql, strConBasa
                
                If Not rslegajos.EOF Then
                    Sql = " Update DOCUMENTOS_DIGITALES"
                    Sql = Sql & " SET LETRA_DESDE ='" & rslegajos!LETRA_DESDE & "'"
                    Sql = Sql & " , NRO_DESDE =" & rslegajos!NRO_DESDE
                    Sql = Sql & " Where ID = " & rs!ID
                    Sql = Sql & " And (LETRA_DESDE Is Null) And (NRO_DESDE Is Null)"
                    ExecutarSql Sql
                End If
            End If
                rs.MoveNext
            Loop
            
            




End Sub

Private Sub Command51_Click()
Dim Sql As String
            Dim rs As New ADODB.Recordset
            Dim rslegajos As ADODB.Recordset
            Dim ID_LEGAJO As String
            
            Sql = " SELECT  DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER, DOCUMENTOS_DIGITALES_LOTE.REMITO, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE,"
            Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.CANTIDAD_IMAGENES, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_VIEJA.FK_LEGAJO_ETIQUETA, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
            Sql = Sql & " LEN(DOCUMENTOS_DIGITALES_VIEJA.FK_LEGAJO_ETIQUETA) AS Expr1, DOCUMENTOS_DIGITALES_VIEJA.ID"
            Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
            Sql = Sql & " DOCUMENTOS_DIGITALES_VIEJA ON DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_VIEJA.FK_DOCUMENTOS_DIGITALES_LOTE"
            Sql = Sql & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER BETWEEN CONVERT(DATETIME, '2017-09-29 00:00:00', 102) AND CONVERT(DATETIME, '2017-10-31 00:00:00', 102)) AND"
 Sql = Sql & " (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 279) and LETRA_DESDE is null "
            Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER"
            
           Sql = " SELECT        DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER, DOCUMENTOS_DIGITALES_LOTE.REMITO, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE,"
           Sql = Sql & "               DOCUMENTOS_DIGITALES_LOTE.CANTIDAD_IMAGENES, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_VIEJA.FK_LEGAJO_ETIQUETA, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
           Sql = Sql & "               LEN(DOCUMENTOS_DIGITALES_VIEJA.FK_LEGAJO_ETIQUETA) AS Expr1, DOCUMENTOS_DIGITALES_VIEJA.ID, DOCUMENTOS_DIGITALES_VIEJA.NRO_DESDE"
Sql = Sql & " FROM            DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
  Sql = Sql & "                        DOCUMENTOS_DIGITALES_VIEJA ON DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_VIEJA.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 279) And (DOCUMENTOS_DIGITALES_VIEJA.NRO_DESDE Is Null) And (Not (DOCUMENTOS_DIGITALES_VIEJA.FK_LEGAJO_ETIQUETA Is Null))"
Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER"
            
            
            rs.Open Sql, ConBasa
            
            Do While Not rs.EOF
                ID_LEGAJO = ""
                
                
                
            If Len(rs!FK_LEGAJO_ETIQUETA) = 12 Then
            ID_LEGAJO = Mid(rs!FK_LEGAJO_ETIQUETA, 5, 8)
            End If
            If Len(rs!FK_LEGAJO_ETIQUETA) = 13 Then
            ID_LEGAJO = Mid(rs!FK_LEGAJO_ETIQUETA, 5, 8)
            End If
            
            If ID_LEGAJO <> "" Then
            
                Set rslegajos = New ADODB.Recordset
                Sql = " SELECT LETRA_DESDE, NRO_DESDE From LEGAJOS"
                Sql = Sql & " Where ID_LEGAJO = " & ID_LEGAJO
                Sql = Sql & " And (COD_CLIENTE = 279)"
                rslegajos.Open Sql, strConBasa
                
                If Not rslegajos.EOF Then
                    Sql = " Update DOCUMENTOS_DIGITALES_VIEJA"
                    Sql = Sql & " SET LETRA_DESDE ='" & rslegajos!LETRA_DESDE & "'"
                    Sql = Sql & " , NRO_DESDE =" & rslegajos!NRO_DESDE
                    Sql = Sql & " Where ID = " & rs!ID
                    Sql = Sql & " And (LETRA_DESDE Is Null) And (NRO_DESDE Is Null)"
                    ExecutarSql Sql
                End If
            End If
                rs.MoveNext
            Loop
            
            

End Sub

Private Sub Command52_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String

rs.CursorLocation = adUseClient
Sql = "  SELECT        DIRECTORIO_PASO, ID, COPIADA"
Sql = Sql & " From DOCUMENTOS_DIGITALES"
Sql = Sql & " WHERE        id > " & InputBox("Ingrese el id de imagen de inicio")
Sql = Sql & " ORDER BY ID"




Sql = " SELECT         ID_CENTROCARD.ID, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO"
Sql = Sql & "  FROM            ID_CENTROCARD INNER JOIN"
Sql = Sql & "                          DOCUMENTOS_DIGITALES ON ID_CENTROCARD.ID = DOCUMENTOS_DIGITALES.ID"
Sql = Sql & "  ORDER BY ID_CENTROCARD.ID"


rs.Open Sql, ConActiva, adOpenKeyset, adLockOptimistic


Do While Not rs.EOF
 If Dir("X:\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif") <> "" Then
    FileSystem.FileCopy "X:\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF"
    Sql = " Update DOCUMENTOS_DIGITALES"
    Sql = Sql & " SET  COPIADA ='Copiada'"
    Sql = Sql & "  Where ID = " & rs!ID
Else
    Sql = " Update DOCUMENTOS_DIGITALES"
    Sql = Sql & " SET  COPIADA ='No se encontro'"
    Sql = Sql & "  Where ID = " & rs!ID
End If
  
  ExecutarSql Sql
  
    rs.MoveNext
Loop

End Sub

Private Sub Command53_Click()
Dim Sql As String
Dim rs As ADODB.Recordset



Sql = " SELECT CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, CONTENEDOR.ID_CONTENEDOR"
Sql = Sql & " FROM   CONTENEDOR LEFT OUTER JOIN"
Sql = Sql & "  CAJAS ON CONTENEDOR.COD_CLIENTE = CAJAS.FK_CLIENTE AND CONTENEDOR.NRO_CAJA = CAJAS.NRO_CAJA"
Sql = Sql & "  Where (CAJAS.FK_CLIENTE Is Null) "
Sql = Sql & "  And (Not (CONTENEDOR.COD_CLIENTE Is Null))"
rs.Open Sql, strConBasa

 Do While Not rs.EOF
 
        Sql = "  Update CONTENEDOR"
        Sql = Sql & " SET            COD_CLIENTE_ERROR =    COD_CLIENTE , NRO_CAJA_ERROR = NRO_CAJA "
        Sql = Sql & " Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR
        ExecutarSql Sql
 Sql = " Update CONTENEDOR"
 Sql = Sql & "  SET                COD_CLIENTE = 2000"
 Sql = Sql & "  Where ID_CONTENEDOR = " & rs!ID_CONTENEDOR
     ExecutarSql Sql
    rs.MoveNext
 Loop
 

End Sub

Private Sub Command7_Click()

    Dim direc(90) As String
    Dim i As Integer
    
    Dim sFolderPath As String
    sFolderPath = "C:\Windows\"
   
 Dim sArchivo As String

sArchivo = Dir(sFolderPath & "*", vbDirectory)
Do While sArchivo <> ""
        If GetAttr(sFolderPath & sArchivo) = 16 Then
            direc(i) = sArchivo
            i = i + 1
        End If
        sArchivo = Dir
Loop

For i = 0 To 90
   
   If Trim(direc(i)) <> "" Then
   MsgBox direc(i)
    End If
Next




'If GetAttr("C:\" & Fld) = 16 Then
'MsgBox ("This is a folder")
'End If
'
'I tried it, works perfectly.
'
'    For Each subfolder In FSfolder.SubFolders
'        DoEvents
'        i = i + 1
'        Debug.Print subfolder
'    Next subfolder
'    Set FSfolder = Nothing
'    MsgBox "Total sub folders in " & sFolderPath & " : " & i

End Sub

Private Sub Command8_Click()
    Dim Sql As String
    Dim F As Folders
    
    
    Dim rs As New ADODB.Recordset
    Dim NUMERO As Integer
        Sql = " SELECT     PASOORIGEN, IMAGEN_ORIGEN, ID"
        Sql = Sql & " From DOCUMENTOS_DIGITALES"
        Sql = Sql & " Where (COD_CLIENTE = 84) And (NRO_CAJA > 200)"
        Sql = Sql & " And (IMAGEN_ORIGEN Is Null)"
        Sql = Sql & "  ORDER BY NRO_CAJA"
        rs.Open Sql, ConActiva, 0, 1
        Do While Not rs.EOF
            NUMERO = Mid(rs!PasoOrigen, Len(Trim(rs!PasoOrigen)) - 7, 4)
            Sql = "  Update DOCUMENTOS_DIGITALES Set IMAGEN_ORIGEN = " & NUMERO
            Sql = Sql & " Where ID = " & rs!ID
            ExecutarSql Sql
            rs.MoveNext
        Loop
End Sub

Private Sub Command9_Click()
ExportarCOHEN
End Sub

Private Sub ctlVerImagenes1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Source.Top = ctlVerImagenes1.Top + Y
Source.Left = ctlVerImagenes1.Left + X
Rem Source.SetFocus
End Sub

Private Sub ExpresoLujan_Click(Index As Integer)
Rem ExpresoLujanSacarImagenesConError
 ExpresoLujanExportarCodigo




End Sub

Private Sub Form_Load()
ctlCliente.TipoControl = Cliente
ctlClienteContar.TipoControl = Cliente
Dim i As Integer
LimpiarCampos

Dim CAN As Integer
ctlPersonalIndexacion.TipoControl = Personal
For i = 0 To txtDato.Count - 1
           txtDato.Item(i).FontSize = 16
           txtDato.Item(i).Refresh
           fraCampos.Item(i).Height = 675
            txtDato.Item(i).Height = 435
        Next
        
        
       
    Dim El_Directorio As Folder
    
    
    
    DoEvents
    
     
        
End Sub

Private Sub imgRemitOsde_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
Source.Top = X
End Sub


'Private Sub RellenarNodosPlantillas(ByVal El_Directorio As Folder)
'
'    On Error GoTo errsub
'        Dim Contador As String
'
'        'Variable del tipo Folder
'        Dim SubDirectorio As Folder
'
'
'        Contador = "1"
'        'Recorrer los subdirectorios
'        For Each SubDirectorio In El_Directorio.SubFolders
'            'Agregar el Path
'           Rem  Set nodeX = ArbolPlantillas.Nodes.Add(, tvwChild, "A" + contador, SubDirectorio.Name)
'            Call RellenarSubNodos(SubDirectorio, Contador)
'            'Sigue listando los directorios
'            RellenarNodosPlantillas SubDirectorio
'            Contador = Str(CInt(Contador) + 1)
'        Next
'
'    Exit Sub
'
'    'Error
'errsub:
'    'Error de permiso denegado
'    If Err.Number = 70 Then
'        Resume Next
'    ElseIf Err.Number = 91 Then
'        Screen.MousePointer = vbDefault
'        Exit Sub
'    Else
'        MsgBox Err.Description, vbCritical
'        Exit Sub
'
'    End If
'End Sub
'Código:
'Private Sub RellenarSubNodos(ByVal SubDirectorio As String, Contador As String)
'    Dim nodeX As Node
'    Dim contadorS
'    Dim El_Archivo As File
'    Dim El_Directorio As Folder
'    Dim Fso As FileSystemObject
'    Dim DirDocumentos As String
'
'    'Nuevo objeto felisystemobjetc
'    Set Fso = New FileSystemObject
'
'    'Obtiene el directorio
'    DirDocumentos = V_DAtendidos + "\" + CStr(Me.T_NHistoria) + "\Documentos"
'    Set El_Directorio = Fso.GetFolder(SubDirectorio)
'
'    contadorS = "1"
'
'    'Listar los ficheros de esta carpeta
'    For Each El_Archivo In El_Directorio.Files
'        Set nodeX = ArbolPlantillas.Nodes.Add("A" + Contador, tvwChild, "B" + contadorS, El_Archivo.Name)
'        contadorS = Str(CInt(contadorS) + 1)
'    Next El_Archivo
'
'
'End Sub
'
Private Sub Form_Resize()
On Error GoTo salir
SSTab1.Height = frmIndexarImganenes.Height - 100
SSTab1.Width = frmIndexarImganenes.Width - 100
ctlVerImagenes1.Width = SSTab1.Width - 200
ctlVerImagenes1.Height = SSTab1.Height - ctlVerImagenes1.Top - 1500
salir:
End Sub

Private Sub grdIndexarImagenes_Click()
PonerImagenLocal
End Sub
       
       
       
       
Private Sub PegarDatos()

 Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet


 Dim Min As Long
    Dim Max As Long
       
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Add
    Set hojaEx = libroEx.Worksheets.Item(1)
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim PasoInicial As String
    
         PasoInicial = "D:\ExportarImagenes\"

    Dim i As Integer
    Dim R As Excel.Range
    Dim h As Excel.Hyperlinks
    i = 1

                   
            
            hojaEx.Cells(1, 1) = "Tipo Documento"
            hojaEx.Cells(1, 2) = "Dato"
            hojaEx.Cells(1, 3) = "Nombre"
            hojaEx.Cells(1, 4) = "Caja"
            
        
        
      

        Dim sgrabar As String

        


Sql = " SELECT  INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.Nombre, DOCUMENTOS_DIGITALES.ID,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE , DOCUMENTOS_DIGITALES.DIRECTORIO_PASO, DOCUMENTOS_DIGITALES.NRO_CAJA"
Sql = Sql & vbCrLf & "  FROM   INDICES INNER JOIN"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES ON INDICES.COD_CLIENTE = DOCUMENTOS_DIGITALES.COD_CLIENTE AND"
Sql = Sql & vbCrLf & " INDICES.Indice = DOCUMENTOS_DIGITALES.Indice"
Sql = Sql & vbCrLf & "  Where DOCUMENTOS_DIGITALES.COD_CLIENTE = " & ctlCliente.Valor
Sql = Sql & vbCrLf & "  ORDER BY INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.NRO_CAJA"

Dim NombreArchivo As String
      
        i = 2
        rs.Open Sql, ConActiva, 0, 1

            Do While Not rs.EOF
                i = i + 1
               Rem NombreArchivo = Trim(RS!LETRA_DESDE) & "_" & CStr(RS!ID) & ".tif"
                NombreArchivo = CStr(rs!ID) & ".tif"
             hojaEx.Cells(i, 1) = rs!Descripcion
                   hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\" & rs!NRO_CAJA & "\" & NombreArchivo
                
                If IsNull(rs!LETRA_DESDE) Then
                    hojaEx.Cells(i, 2) = rs!NRO_DESDE
                Else
                    hojaEx.Cells(i, 2) = rs!LETRA_DESDE
                End If
                hojaEx.Cells(i, 3) = rs!Nombre
                If Dir(PasoInicial & rs!NRO_CAJA, vbDirectory) = "" Then
                    FileSystem.MkDir PasoInicial & rs!NRO_CAJA
                Else

                End If
                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & rs!NRO_CAJA & "\" & NombreArchivo
                hojaEx.Cells(i, 4) = rs!NRO_CAJA
                hojaEx.Cells(i, 5) = rs!ID
                
                rs.MoveNext
            Loop

           libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY") & "2.xls"
           libroEx.Close
           ApExcel.Quit
           Set ApExcel = Nothing
           Set libroEx = Nothing
  


End Sub

Private Sub grdIndexarImagenes_DblClick()
Dim Paso
Paso = PasoImagenes & rsGrilla!DIRECTORIO_PASO & "\" & rsGrilla!ID & ".TIF"
Clipboard.Clear
Clipboard.SetText Paso
MsgBox "Paso Copiado"
End Sub

Private Sub grdLotes_Click()
    txtLotesExportar.Text = txtLotesExportar.Text & "," & grdLotes.Text
End Sub

Private Sub grdLotes_DblClick()
 Dim Cliente As Integer
 Dim lote As String
 Dim i As Integer
 Dim Orden As String
 
 lote = grdLotes.Columns("LOTE").Text
 lblLote.Caption = grdLotes.Columns("LOTE").Text
 lblCliente.Caption = ctlCliente.Valor
 Dim rs As New ADODB.Recordset
 Dim Sql As String
 Dim R As Integer
 Dim TitulosSql As String
 On Error GoTo salir
 
 If optImagenLocal.value = False And optImagenServer.value = False Then
        MsgBox "Por favor defina el modo de ver la imagen Local o Server", vbCritical
        Exit Sub
   End If
 
    Sql = " SELECT     FK_INDICES, COD_CLIENTE, COD_INDICE, COPIAR_LETRA_DESDE_NRO_DESDE"
    Sql = Sql & vbCrLf & " , LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA"
    Sql = Sql & vbCrLf & " , DESCRIPCION , Nombre, FECHA_HASTA, FECHA_DESDE "
    Sql = Sql & vbCrLf & " From INDICE_DIGITALIZACION"
    Sql = Sql & vbCrLf & " WHERE FK_INDICES = " & grdLotes.Columns("FK_INDICES").Text
    
    
    
    rs.Open Sql, strConBasa
    LimpiarCampos
    chkCopiarLetra_Numero.value = 0
    Orden = ""

If Not rs.EOF Then
    For R = 3 To rs.Fields.Count - 1
        
       If Not IsNull(rs.Fields(R).value) Then
            If rs!COPIAR_LETRA_DESDE_NRO_DESDE = "1" And rs.Fields(R).Name = "COPIAR_LETRA_DESDE_NRO_DESDE" Then
                chkCopiarLetra_Numero.value = 1
            Else
                fraCampos.Item(FrameOcupados).Visible = True
                lblTitulo.Item(FrameOcupados).Visible = True
                lblTitulo.Item(FrameOcupados).Caption = rs.Fields(R).value
                txtDato.Item(FrameOcupados).DataField = rs.Fields(R).value
                txtDato.Item(FrameOcupados).Tag = rs.Fields(R).Name
                txtDato.Item(FrameOcupados).Text = ""
                FrameOcupados = FrameOcupados + 1
                TitulosSql = TitulosSql & "," & rs.Fields(R).Name & " AS " & rs.Fields(R).value
                If Orden = "" Then
                    Orden = rs.Fields(R).Name
                End If
            End If
            
        End If
    Next

End If
          
 Set rsGrilla = New ADODB.Recordset
 


    If chkControlExpreso.value = 0 Then
        Sql = "SELECT ID, DIRECTORIO_PASO "
        Sql = Sql & TitulosSql
        Sql = Sql & " FROM DOCUMENTOS_DIGITALES"
        Sql = Sql & " WHERE  FK_DOCUMENTOS_DIGITALES_LOTE = " & grdLotes.Columns("LOTE").Text
        Sql = Sql & " order by LETRA_HASTA "
    Else
        Sql = "SELECT ID, DIRECTORIO_PASO "
        Sql = Sql & TitulosSql
        Sql = Sql & " FROM DOCUMENTOS_DIGITALES"
        Sql = Sql & " WHERE   ID in ( " & ControlExpreso(grdLotes.Columns("LOTE").Text) & ")"
        Sql = Sql & " order by ID "
    End If


If chkControlExpresoGuiaSucursal.value = 1 Then
    Sql = "SELECT ID, DIRECTORIO_PASO "
    Sql = Sql & TitulosSql
    Sql = Sql & " FROM DOCUMENTOS_DIGITALES"
    Sql = Sql & " WHERE   ID in ( " & ControlExpresoGuia(grdLotes.Columns("LOTE").Text) & ")"
    Sql = Sql & " order by ID "

End If




rsGrilla.CursorLocation = adUseClient

rsGrilla.Open Sql, strConBasa, adOpenKeyset, adLockOptimistic

Rem rsGrilla.Open Sql, strConBasa, adOpenKeyset, adLockPessimistic



Set grdIndexarImagenes.DataSource = rsGrilla.DataSource
grdIndexarImagenes.ReBind
grdIndexarImagenes.Refresh

For i = 0 To fraCampos.Count - 1
    If fraCampos.Item(i).Visible = True Then
        Set txtDato.Item(i).DataSource = rsGrilla.DataSource
    End If
    Next

cboOrden.Clear
For i = 0 To rsGrilla.Fields.Count - 1
    cboOrden.AddItem rsGrilla.Fields(i).Name
Next


grdIndexarImagenes.Columns(0).Visible = True
grdIndexarImagenes.Columns(0).Locked = True
grdIndexarImagenes.Columns(1).Visible = False

ctlVerImagenes1.ZonnFijo 0.4
If Not rsGrilla.EOF Then
    
    
   Rem  ctlVerImagenes1.PonerImagen PasoImagenes & "\" & rsGrilla!DIRECTORIO_PASO & "\" & rsGrilla!ID & ".TIF"
PonerImagenLocal

Else
    MsgBox "No hay Mas"
End If


SSTab1.Tab = 1
 
 Exit Sub
salir:
 MsgBox Err.Description
End Sub

'Private Sub grdLotes()
' Dim Cliente As Integer
' Dim lote As String
' Dim i As Integer
' Dim Orden As String
'
' lote = grdLotes.Columns("LOTE").Text
' lblLote.Caption = grdLotes.Columns("LOTE").Text
' lblCliente.Caption = ctlCliente.Valor
' Dim rs As New ADODB.Recordset
' Dim Sql As String
' Dim R As Integer
' Dim TitulosSql As String
' On Error GoTo salir
'
'
'
'    Sql = " SELECT     FK_INDICES, COD_CLIENTE, COD_INDICE, COPIAR_LETRA_DESDE_NRO_DESDE"
'    Sql = Sql & vbCrLf & " , LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA"
'    Sql = Sql & vbCrLf & " , DESCRIPCION , Nombre, FECHA_HASTA, FECHA_DESDE "
'    Sql = Sql & vbCrLf & " From INDICE_DIGITALIZACION"
'    Sql = Sql & vbCrLf & " WHERE FK_INDICES = " & grdLotes.Columns("FK_INDICES").Text
'
'
'
'    rs.Open Sql, strConBasa
'    LimpiarCampos
'    chkCopiarLetra_Numero.value = 0
'    Orden = ""
'
'If Not rs.EOF Then
'    For R = 3 To rs.Fields.Count - 1
'
'        If Not IsNull(rs.Fields(R).value) Then
'            If rs!COPIAR_LETRA_DESDE_NRO_DESDE = "1" And rs.Fields(R).Name = "COPIAR_LETRA_DESDE_NRO_DESDE" Then
'                chkCopiarLetra_Numero.value = 1
'            Else
'            fraCampos.Item(FrameOcupados).Visible = True
'            lblTitulo.Item(FrameOcupados).Visible = True
'            lblTitulo.Item(FrameOcupados).Caption = rs.Fields(R).value
'            txtDato.Item(FrameOcupados).DataField = rs.Fields(R).value
'            txtDato.Item(FrameOcupados).Tag = rs.Fields(R).Name
'            txtDato.Item(FrameOcupados).Text = ""
'            FrameOcupados = FrameOcupados + 1
'            TitulosSql = TitulosSql & "," & rs.Fields(R).Name & " AS " & rs.Fields(R).value
'            If Orden = "" Then
'                Orden = rs.Fields(R).Name
'            End If
'            End If
'
'        End If
'    Next
'
'End If
'
' Set rsGrilla = New ADODB.Recordset
'
'
'
'If chkControlExpreso.value = 0 Then
'
''Sql = "SELECT     ID, DIRECTORIO_PASO"
''Sql = Sql & TitulosSql
''Sql = Sql & " FROM DOCUMENTOS_DIGITALES"
''Sql = Sql & " WHERE    FK_DOCUMENTOS_DIGITALES_LOTE = " & grdLotes.Columns("LOTE").Text
''Sql = Sql & " order by  " & Orden
'Sql = "SELECT     ID, DIRECTORIO_PASO"
'Sql = Sql & TitulosSql
'Sql = Sql & " FROM DOCUMENTOS_DIGITALES"
'Sql = Sql & " WHERE  FK_DOCUMENTOS_DIGITALES_LOTE = " & grdLotes.Columns("LOTE").Text
'Sql = Sql & " order by nro_desde "
'
'
'Else
'
'Sql = "SELECT     ID, DIRECTORIO_PASO"
'Sql = Sql & TitulosSql
'Sql = Sql & " FROM DOCUMENTOS_DIGITALES"
'Sql = Sql & " WHERE   ID in ( " & ControlExpreso(grdLotes.Columns("LOTE").Text) & ")"
'Sql = Sql & " order by ID "
'
'
'
'End If
'
'
'If chkControlExpresoGuiaSucursal.value = 1 Then
'    Sql = "SELECT     ID, DIRECTORIO_PASO"
'    Sql = Sql & TitulosSql
'    Sql = Sql & " FROM DOCUMENTOS_DIGITALES"
'    Sql = Sql & " WHERE   ID in ( " & ControlExpresoGuia(grdLotes.Columns("LOTE").Text) & ")"
'    Sql = Sql & " order by ID "
'
'End If
'
'
'
'
'rsGrilla.CursorLocation = adUseClient
'
'rsGrilla.Open Sql, strConBasa, adOpenKeyset, adLockOptimistic
'
'Rem rsGrilla.Open Sql, strConBasa, adOpenKeyset, adLockPessimistic
'
'
'
'Set grdIndexarImagenes.DataSource = rsGrilla.DataSource
'grdIndexarImagenes.Rebind
'grdIndexarImagenes.Refresh
'
'For i = 0 To fraCampos.Count - 1
'    If fraCampos.Item(i).Visible = True Then
'        Set txtDato.Item(i).DataSource = rsGrilla.DataSource
'    End If
'    Next
'
'cboOrden.Clear
'For i = 0 To rsGrilla.Fields.Count - 1
'    cboOrden.AddItem rsGrilla.Fields(i).Name
'Next
'
'
'grdIndexarImagenes.Columns(0).Visible = True
'grdIndexarImagenes.Columns(0).Locked = True
'grdIndexarImagenes.Columns(1).Visible = False
'ctlVerImagenes1.ZonnFijo 0.4
'If Not rsGrilla.EOF Then
'ctlVerImagenes1.PonerImagen PasoImagenes & "\" & rsGrilla!DIRECTORIO_PASO & "\" & rsGrilla!ID & ".TIF"
'Else
'MsgBox "No hay Mas"
'End If
'
'
'SSTab1.Tab = 1
'
' Exit Sub
'salir:
' MsgBox Err.Description
'End Sub
Private Sub ExportarHilebrand()

 Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet


 Dim Min As Long
    Dim Max As Long
       
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Add
    Set hojaEx = libroEx.Worksheets.Item(1)
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim PasoInicial As String
    
         PasoInicial = txtPasoImagenesFinal

    Dim i As Integer
    Dim R As Excel.Range
    Dim h As Excel.Hyperlinks
    i = 1

                   
            
            hojaEx.Cells(1, 1) = "Link Imagen"
            hojaEx.Cells(1, 2) = "Nombre Imagen"
            hojaEx.Cells(1, 3) = "Caja"
            hojaEx.Cells(1, 4) = "Legajo"
            
        
        
      

        Dim sgrabar As String

        




      


Sql = " SELECT     INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_HASTA,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO,   DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS AS NRO_CAJA,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.LETRA_DESDE AS Expr1, DOCUMENTOS_DIGITALES.LETRA_HASTA,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.Lote , DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES.Nombre"
Sql = Sql & vbCrLf & "  FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & vbCrLf & " INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
Sql = Sql & vbCrLf & " WHERE   (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 84) "
Sql = Sql & vbCrLf & "  AND DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in( " & Mid(txtLotesExportar, 2) & ") "
Sql = Sql & vbCrLf & "  ORDER BY DOCUMENTOS_DIGITALES.NRO_CAJA"


Dim NombreArchivo As String
      
        i = 2
        rs.Open Sql, ConActiva, 0, 1

            Do While Not rs.EOF
                i = i + 1
                NombreArchivo = Trim(rs!LETRA_DESDE) & "_" & CStr(rs!ID) & ".tif"
                hojaEx.Cells(i, 1) = NombreArchivo
                hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\0\" & rs!NRO_CAJA & "\" & NombreArchivo
                
                If IsNull(rs!LETRA_DESDE) Then
                    hojaEx.Cells(i, 2) = rs!NRO_DESDE
                Else
                    hojaEx.Cells(i, 2) = rs!LETRA_DESDE
                End If
                hojaEx.Cells(i, 2) = rs!ID
                hojaEx.Cells(i, 3) = rs!NRO_CAJA
                hojaEx.Cells(i, 4) = Trim(rs!LETRA_DESDE)
                If Dir(PasoInicial & rs!NRO_CAJA, vbDirectory) = "" Then
                    FileSystem.MkDir PasoInicial & rs!NRO_CAJA
                Else

                End If

                 FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & rs!NRO_CAJA & "\" & NombreArchivo
                hojaEx.Cells(i, 4) = rs!NRO_CAJA
                hojaEx.Cells(i, 5) = rs!ID
                
                rs.MoveNext
            Loop

           libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY_SS") & ".xls"
             
           libroEx.Close
           ApExcel.Quit
           Set ApExcel = Nothing
           Set libroEx = Nothing
           MsgBox "terminado"
           
  


End Sub


'
'
'
'
'
'
'

Private Sub ExportarCOHEN()

 Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet


 Dim Min As Long
    Dim Max As Long
       
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Add
    Set hojaEx = libroEx.Worksheets.Item(1)
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim PasoInicial As String
    PasoInicial = txtPasoImagenesFinal.Text

    Dim i As Integer
    Dim R As Excel.Range
    Dim h As Excel.Hyperlinks
    i = 1
        hojaEx.Cells(1, 1) = "Link Imagen"
        hojaEx.Cells(1, 2) = "Nombre Imagen"
        hojaEx.Cells(1, 3) = "Caja"
        hojaEx.Cells(1, 4) = "Legajo"
        Dim sgrabar As String
        Sql = " SELECT     INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_HASTA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO,   DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS AS NRO_CAJA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.LETRA_DESDE AS Expr1, DOCUMENTOS_DIGITALES.LETRA_HASTA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.Lote , DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES.Nombre"
        Sql = Sql & vbCrLf & "  FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES ON"
        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & vbCrLf & " INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
         Sql = Sql & vbCrLf & "  Where      DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in( " & Mid(txtLotesExportar, 2) & ") "
        Rem   Sql = Sql & vbCrLf & "  Where      DOCUMENTOS_DIGITALES.FILTRO = 18072013"
        Sql = Sql & vbCrLf & "  ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE "
        Dim NombreArchivo As String
        i = 2
        rs.Open Sql, ConActiva, 0, 1

            Do While Not rs.EOF
                i = i + 1
                 NombreArchivo = Trim(rs!NRO_DESDE) & " _ " & CStr(rs!ID) & ".tif"
                Rem  NombreArchivo = Trim(rs!NRO_DESDE) & ".tif"
                
                NombreArchivo = Replace(NombreArchivo, Chr(10), "")
                hojaEx.Cells(i, 1) = NombreArchivo
                hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\" & NombreArchivo
                
                               
                hojaEx.Cells(i, 2) = Trim(rs!NRO_DESDE)
                hojaEx.Cells(i, 3) = Trim(rs!LETRA_DESDE)
               
'               If Dir(PasoInicial & rs!NRO_CAJA, vbDirectory) = "" Then
'                    FileSystem.MkDir PasoInicial & rs!NRO_CAJA
'                Else
'
'                End If

                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & NombreArchivo
                hojaEx.Cells(i, 4) = rs!NRO_CAJA
                hojaEx.Cells(i, 5) = rs!ID
                
                rs.MoveNext
            Loop

           libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY_ss") & ".xls"
           libroEx.Close
           ApExcel.Quit
           Set ApExcel = Nothing
           Set libroEx = Nothing
           
           MsgBox "Terminado"
  


End Sub


Private Sub ExportarChandonddhh()

 Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet


 Dim Min As Long
    Dim Max As Long
       
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Add
    Set hojaEx = libroEx.Worksheets.Item(1)
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim PasoInicial As String
    PasoInicial = txtPasoImagenesFinal.Text

    Dim i As Integer
    Dim R As Excel.Range
    Dim h As Excel.Hyperlinks
    i = 1
        hojaEx.Cells(1, 1) = "Link Imagen"
        hojaEx.Cells(1, 2) = "Nombre Imagen"
        hojaEx.Cells(1, 3) = "Caja"
        hojaEx.Cells(1, 4) = "Legajo"
        Dim sgrabar As String
        Sql = " SELECT     INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_HASTA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO,   DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS AS NRO_CAJA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.LETRA_DESDE AS Expr1, DOCUMENTOS_DIGITALES.LETRA_HASTA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.Lote , DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES.Nombre"
        Sql = Sql & vbCrLf & "  FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES ON"
        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & vbCrLf & " INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
         Sql = Sql & vbCrLf & "  Where      DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in( " & Mid(txtLotesExportar, 2) & ") "
        Rem   Sql = Sql & vbCrLf & "  Where      DOCUMENTOS_DIGITALES.FILTRO = 18072013"
        Sql = Sql & vbCrLf & "  ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE "
        Dim NombreArchivo As String
        i = 2
        rs.Open Sql, ConActiva, 0, 1

            Do While Not rs.EOF
                i = i + 1
                 NombreArchivo = Trim(rs!LETRA_DESDE) & " _ " & Trim(rs!LETRA_HASTA) & " _ " & CStr(rs!ID) & ".tif"
                Rem  NombreArchivo = Trim(rs!NRO_DESDE) & ".tif"
                
                NombreArchivo = Replace(NombreArchivo, Chr(10), "")
                hojaEx.Cells(i, 1) = NombreArchivo
                hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\" & NombreArchivo
                
                               
                hojaEx.Cells(i, 2) = Trim(rs!LETRA_DESDE)
                hojaEx.Cells(i, 3) = Trim(rs!LETRA_HASTA)
               
'               If Dir(PasoInicial & rs!NRO_CAJA, vbDirectory) = "" Then
'                    FileSystem.MkDir PasoInicial & rs!NRO_CAJA
'                Else
'
'                End If

                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & NombreArchivo
                hojaEx.Cells(i, 4) = rs!NRO_CAJA
                hojaEx.Cells(i, 5) = rs!ID
                
                rs.MoveNext
            Loop

           libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY_ss") & ".xls"
           libroEx.Close
           ApExcel.Quit
           Set ApExcel = Nothing
           Set libroEx = Nothing
           
           MsgBox "Terminado"
  


End Sub


Public Sub LimpiarCampos()
Dim i As Integer
For i = 0 To fraCampos.Count - 1
   fraCampos.Item(i).Visible = False
   txtDato.Item(i).Text = ""
   txtDato.Item(i).DataField = ""
   lblTitulo.Item(i).Caption = ""
   lblTitulo.Item(FrameOcupados).Caption = ""
   txtDato.Item(FrameOcupados).DataField = ""
   txtDato.Item(FrameOcupados).Tag = ""
   txtDato.Item(FrameOcupados).Text = ""
   
Next
FrameOcupados = 0
End Sub

Private Sub grdLotes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuManejoLotes

End If

End Sub



Private Sub mnuAcuses9151_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Documento As Long
Dim Nombre As String
Dim P As Integer



Sql = "SELECT DOCUMENTOS_DIGITALES.ID"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE  (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 40)  "
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.LETRA_DESDE IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.FECHA_DESDE IS NULL)"
rs.Open Sql, strConBasa
Do While Not rs.EOF

Sql = " SELECT     ID, ENBASA, letra_desde , nro_desde, BatchPgDta"
Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
Sql = Sql & " ORDER BY ID"

Set rs2 = New ADODB.Recordset

rs2.Open Sql, strConBasa

Dim i As Integer
If Not rs2.EOF Then
    If Not IsNumeric(rs2!NRO_DESDE) Then
        Documento = 0
        Else
            If Not IsNull(rs2!NRO_DESDE) Then
                Documento = rs2!NRO_DESDE
                If IsNull(rs2!LETRA_DESDE) Then
                Nombre = "Null"
                Else
                    Nombre = Replace(Trim(rs2!LETRA_DESDE), Chr(10), " // ")
                    Nombre = Replace(Nombre, ",", " ")
                    Nombre = "'" & Nombre & "'"
                End If
                
                Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
                Sql = Sql & " SET "
                Sql = Sql & " LETRA_DESDE =" & Nombre
                Sql = Sql & ", LETRA_HASTA =" & Nombre
                Sql = Sql & ", NRO_DESDE =" & Documento
                Sql = Sql & ", NRO_HASTA =" & Documento
                Sql = Sql & " Where ID = " & rs!ID
                Sql = Sql & " and   ( (NRO_DESDE IS NULL)) "
                ExecutarSql Sql
            End If
     End If
 End If

 
 rs.MoveNext

Loop


MsgBox "Terminado"

End Sub

Private Sub mnuAirLiquede_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Documento As Long
Dim Nombre As String
Dim P As Integer



Sql = "SELECT DOCUMENTOS_DIGITALES.ID"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 155"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.DESCRIPCION IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.LETRA_DESDE IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.FECHA_DESDE IS NULL)"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID "

Dim Sucursal As Integer
Dim REMITO As Long
Dim fecha As String
Dim Cliente As String



rs.Open Sql, strConBasa
Do While Not rs.EOF


'829064
'
'Letra_desde  , nro_desde =  Numero de remito
'nro_hasta = sucursal 155
'FECHA_DESDE = fecha
'
'
'en didgital
'
'NRO_DESDE , Cliente
' FECHA_DESDE,
'                     LETRA_DESDE Sucursal, Remito


Sql = " SELECT     ID, ENBASA,NRO_DESDE, FECHA_DESDE ,  LETRA_DESDE , LETRA_HASTA, BatchPgDta"
Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
Sql = Sql & " ORDER BY ID"

Set rs2 = New ADODB.Recordset

rs2.Open Sql, strConBasa
Dim estado As String

estado = "RECONOCIDO"
Dim i As Integer
If Not rs2.EOF Then

If InStr(1, Trim(rs2!LETRA_DESDE), "~") = 0 Then
    If Len(Trim(rs2!LETRA_DESDE)) = 13 Then
       If IsNumeric(Mid(Trim(rs2!LETRA_DESDE), 1, 4)) Then
            Sucursal = Mid(Trim(rs2!LETRA_DESDE), 1, 4)
       Else
            Sucursal = 0
       End If
       If IsNumeric(Mid(Trim(rs2!LETRA_DESDE), 6)) Then
            REMITO = Mid(Trim(rs2!LETRA_DESDE), 6)
        Else
            REMITO = 0
        End If
        
    Else
        Sucursal = 0
        REMITO = 0
        estado = "VERIFICAR MANUALMENTE"
    End If
Else
    Sucursal = 0
    REMITO = 0
    estado = "VERIFICAR MANUALMENTE"

End If


 
If Len(Trim(rs2!LETRA_HASTA)) = 8 Then

    fecha = Trim(rs2!LETRA_HASTA)
    fecha = Mid(fecha, 1, 2) & "/" & Mid(fecha, 4, 2) & "/" & Mid(fecha, 7, 2)
 If IsDate(fecha) Then
 fecha = "'" & fecha & "'"
 Else
 fecha = "Null"
 estado = "VERIFICAR MANUALMENTE"
 End If
Else

fecha = "Null"
estado = "VERIFICAR MANUALMENTE"
End If

If IsNull(rs2!NRO_DESDE) Then
 Cliente = 0
 estado = "VERIFICAR MANUALMENTE"
Else
Cliente = "'" & Replace(rs2!NRO_DESDE, " ", "") & "'"
End If

 


     Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
Sql = Sql & " SET "
Sql = Sql & " NRO_DESDE =" & REMITO
Sql = Sql & "  ,NRO_HASTA =" & Sucursal
Sql = Sql & " , LETRA_HASTA =" & Cliente
Sql = Sql & ", FECHA_DESDE =" & fecha
Sql = Sql & ",ESTADO ='" & estado & "'"

Sql = Sql & " Where ID = " & rs!ID

ExecutarSql Sql
 
 End If
 

 
 rs.MoveNext

Loop


MsgBox "Terminado"


End Sub




Private Sub mnuAriLiquede_Click()
Export_AirLiquide
End Sub

Private Sub mnuBorrarLote_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim PasoImagen As String
    Sql = " SELECT LOTE, ID, DIRECTORIO_PASO,COD_CLIENTE, COD_ESTADO "
    Sql = Sql & " From DOCUMENTOS_DIGITALES"
       Sql = Sql & " WHERE  FK_DOCUMENTOS_DIGITALES_LOTE  = '" & grdLotes.Columns(0).Text & "'"
   Rem  sql = sql & " AND COD_ESTADO = 0 "
    Sql = Sql & "  order by ID"
    rs.Open Sql, strConBasa
    
    
    
    If MsgBox("Esta usted seguro de borra el lote " & grdLotes.Columns(0).Text, vbYesNo) = vbYes Then
            Do While Not rs.EOF
                PasoImagen = PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif"
                If Dir(PasoImagen) <> "" Then
                    FileSystem.Kill PasoImagen
                End If
                ExecutarSql " DELETE   FROM DOCUMENTOS_DIGITALES Where  ID =  " & rs!ID
                rs.MoveNext
            Loop
            
            
           ExecutarSql "  DELETE  FROM DOCUMENTOS_DIGITALES_LOTE Where  ID_DOCUMENTOS_DIGITALES_LOTE =  " & grdLotes.Columns(0).Text

    End If
    rs.Close
    
 cmdBuscar_Click
 MsgBox "Lote Borrado"
 
End Sub


Private Sub mnuChandonExport_Click(Index As Integer)
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Documento As Long
Dim Nombre As String
Dim P As Integer



Sql = "SELECT DOCUMENTOS_DIGITALES.ID"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 47"
Sql = Sql & " and DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE  in(22292,22291)"
'
'Sql = Sql & " AND (DOCUMENTOS_DIGITALES.DESCRIPCION IS NULL)"
'Sql = Sql & " AND (DOCUMENTOS_DIGITALES.LETRA_DESDE IS NULL)"
'Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)"
'Sql = Sql & " AND (DOCUMENTOS_DIGITALES.FECHA_DESDE IS NULL)"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID "

Dim Proveedor As Long



rs.Open Sql, strConBasa
Do While Not rs.EOF


Sql = " SELECT     ID, ENBASA,NRO_DESDE, FECHA_DESDE ,  LETRA_DESDE , LETRA_hasta ,BatchPgDta"
Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
Sql = Sql & " ORDER BY ID"

Set rs2 = New ADODB.Recordset

rs2.Open Sql, strConBasa

Dim i As Integer
If Not rs2.EOF Then

Proveedor = 0
  
If IsNumeric(rs2!NRO_DESDE) Then
    Proveedor = rs2!NRO_DESDE
    
Else
    Proveedor = "NULL"
    
End If

Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
Sql = Sql & " SET "
Sql = Sql & " NRO_DESDE =" & Proveedor
Sql = Sql & "  ,NRO_HASTA =" & Proveedor
Sql = Sql & " , letra_desde ='" & Proveedor & "'"
Sql = Sql & "  ,letra_hasta ='" & Proveedor & "'"
Sql = Sql & " Where ID = " & rs!ID

ExecutarSql Sql
 
 End If
 

 
 rs.MoveNext

Loop


MsgBox "Terminado"
End Sub

Private Sub mnuChandonProveedores_Click()
        
            Dim Sql As String
            Dim rsImagenes As ADODB.Recordset
            Dim rsNoti As ADODB.Recordset
            Dim NumeroAnterior As Long
        
            MousePointer = 11
        
            Dim Max As Long

            Dim rs As New ADODB.Recordset
            Dim i As Integer
            Dim R As Integer
            Dim BANDERA As Boolean
            Set rsBuscar = New ADODB.Recordset
        
            Sql = "  SELECT DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_INDICES, INDICES.DESCRIPCION, "
            Sql = Sql & "   DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.LETRA_DESDE,  DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO, "
            Sql = Sql & "   DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS , DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE "
            Sql = Sql & "   FROM DOCUMENTOS_DIGITALES INNER JOIN "
            Sql = Sql & "   DOCUMENTOS_DIGITALES_LOTE ON "
            Sql = Sql & "   DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE INNER JOIN "
            Sql = Sql & "   INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
            Sql = Sql & "   Where DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 47"
            Sql = Sql & "  and  (FK_INDICES = 9221)"
            Sql = Sql & "   ORDER BY  NRO_DESDE "
            rsBuscar.Open Sql, strConBasa, 0, 1
            R = 1
            Dim fecha As String
            Dim Orden  As String
               Do While Not rsBuscar.EOF
                    fecha = FileSystem.FileDateTime(PasoImagenes & rsBuscar!DIRECTORIO_PASO & "\" & rsBuscar!ID & ".tif")
                    fecha = Format(fecha, "DD_MM_YYYY")
                    Orden = Format(rsBuscar!NRO_DESDE, "000000")
                  Rem   FileCopy PasoImagenes & rsBuscar!DIRECTORIO_PASO & "\" & rsBuscar!ID & ".tif", txtPasoImagenesFinal.Text & Orden & " Fecha " & fecha & ".TIF"
                     FileCopy "\\222.15.19.251\ImagenesPDF\" & rsBuscar!DIRECTORIO_PASO & "\" & rsBuscar!ID & ".PDF", txtPasoImagenesFinal.Text & Orden & " Fecha " & fecha & " " & rsBuscar!ID & ".PDF"
                    rsBuscar.MoveNext
                Loop
                
                
                
                
                
                MousePointer = 0
                MsgBox "Operacion terminada"
End Sub

Private Sub MNUCOHEN_Click()
ExportarCOHEN
End Sub

Private Sub mnuContarimagenes_Click()
 Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim PasoImagen As String
    Sql = " SELECT LOTE, ID, DIRECTORIO_PASO,COD_CLIENTE, COD_ESTADO "
    Sql = Sql & " From DOCUMENTOS_DIGITALES"
     Sql = Sql & " WHERE  FK_DOCUMENTOS_DIGITALES_LOTE  = '" & grdLotes.Columns(0).Text & "'"
       Sql = Sql & "  order by ID"
    rs.Open Sql, ConActiva, adOpenStatic, adLockReadOnly
    
    Dim cantidadImagenes As Integer
    Dim cantidadArchivos As Integer
    
            Do While Not rs.EOF
                PasoImagen = PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif"
           cantidadImagenes = cantidadImagenes + frmManejoArchivo.cantidadImagenes(PasoImagen)
              cantidadArchivos = cantidadArchivos + 1
              rs.MoveNext
            Loop
            
            
  Sql = "   Update DOCUMENTOS_DIGITALES_LOTE SET     CANTIDAD_IMAGENES =" & cantidadImagenes
  Sql = Sql & " , CANTIDAD_ARCHIVOS =" & cantidadArchivos
 Sql = Sql & "  Where ID_DOCUMENTOS_DIGITALES_LOTE = " & grdLotes.Columns(0).Text

            
         ExecutarSql Sql

    
    rs.Close
    
 cmdBuscar_Click
 MsgBox "Actualizado"
End Sub

Private Sub mnuConTrack_Click()
LaCajaConTrack
End Sub

Private Sub mnuControldecodigos_Click()
ExpresoLujanControlCodigo
End Sub

Private Sub mnucontroldeduplicados_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim Duplicados As String

Sql = "SELECT       LETRA_HASTA"
Sql = Sql & " From basasql.dbo.DOCUMENTOS_DIGITALES"
Sql = Sql & "  GROUP BY FK_DOCUMENTOS_DIGITALES_LOTE, LETRA_HASTA"
Sql = Sql & "  HAVING      (COUNT(*) > 1)"
Sql = Sql & " AND  FK_DOCUMENTOS_DIGITALES_LOTE = " & InputBox("Ingrese el numero de id de lote")
    Clipboard.Clear

rs.Open Sql, strConBasa

    Do While Not rs.EOF
         Duplicados = Trim(Duplicados) & vbCrLf & Trim(rs!LETRA_HASTA)
        Rem Duplicados = Duplicados & vbCrLf & rs!ID
        rs.MoveNext
    Loop
    Clipboard.SetText Duplicados
    MsgBox "Terminado"
End Sub

Private Sub mnuControldeimagenes_Click()
ExpresoLujanSacarImagenID
End Sub

Private Sub mnuControlDigitoVerificador_Click()


Dim Sql As String
Dim rs As New ADODB.Recordset

Sql = " SELECT     DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.LETRA_HASTA, DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION,"
Sql = Sql & " DOCUMENTOS_DIGITALES.DESCRIPCION AS Expr1, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_DESDE"
Sql = Sql & " FROM         DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE ON"
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE     (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401) AND  DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE =" & Mid(txtLotesExportar.Text, 2)


rs.Open Sql, strConBasa

Do While Not rs.EOF
    If DigitoVerificadorExpreso(CStr(rs!LETRA_HASTA)) = Mid(Trim(CStr(rs!LETRA_HASTA)), 18, 1) Then
        Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
        Sql = Sql & " SET "
        Sql = Sql & "  LETRA_DESDE =" & rs!LETRA_HASTA
        Sql = Sql & " Where ID = " & rs!ID
        ExecutarSql Sql
    Else
        Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
        Sql = Sql & " SET "
        Sql = Sql & "  LETRA_DESDE =0"
        Sql = Sql & " Where ID = " & rs!ID
        ExecutarSql Sql
    End If
    rs.MoveNext
Loop

MsgBox "TERMINADO"

End Sub

Private Sub mnuCrearDirectorios_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset



Sql = " SELECT        DOCUMENTOS_DIGITALES.LETRA_DESDE"
Sql = Sql & " FROM            DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & "                          DOCUMENTOS_DIGITALES ON"
Sql = Sql & "                          DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE        (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 163) AND (DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS IN (1064766, 1064766, 1064767, 1064768, 1064769,"
Sql = Sql & "                          1064770, 1064771, 1064772, 1064773)) AND (DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN = 1) AND"
Sql = Sql & "                          (DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE <> 38534)"
Sql = Sql & "  GROUP BY DOCUMENTOS_DIGITALES.LETRA_DESDE"
'rs.Open sql, ConBasa
'
'Do While Not rs.EOF
'    MkDir ("C:\La Caja\" & UCase(Trim(rs!LETRA_DESDE)))
'    rs.MoveNext
'Loop


    Sql = " SELECT LACAJA20170207.LETRA_DESDE, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES.ID,"
    Sql = Sql & " DOCUMENTOS_DIGITALES.DIRECTORIO_PASO"
    Sql = Sql & " FROM LACAJA20170207 INNER JOIN DOCUMENTOS_DIGITALES "
    Sql = Sql & " ON LACAJA20170207.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " WHERE        (LACAJA20170207.LETRA_DESDE = 'MUNICIPALIDAD DE LINCOLN')"
    Sql = Sql & " ORDER BY LACAJA20170207.LETRA_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"



Dim Contador As Integer
Dim ContadorImagen1 As Integer
Dim Titulo_Anteriror As String
ContadorImagen1 = 0
Titulo_Anteriror = ""
rs.Open Sql, ConBasa

Do While Not rs.EOF
   If Titulo_Anteriror = Trim(rs!LETRA_DESDE) Then
        Contador = Contador + 1
        If rs!IMAGEN_ORIGEN = 1 Then
           ContadorImagen1 = ContadorImagen1 + 1
        End If
        FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF", "C:\La caja\" & Titulo_Anteriror & "\" & Format(Contador, "000000") & ".TIF"
        
  Else
        Titulo_Anteriror = Trim(rs!LETRA_DESDE)
        Contador = 1
        If rs!IMAGEN_ORIGEN = 1 And ContadorImagen1 <> 1 Then
            ContadorImagen1 = 1
            FileCopy "\\222.15.19.251\Imagenes\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF", "C:\La caja\" & Titulo_Anteriror & "\000001 CARATUTA.TIF"
        End If
   End If
   
   
   
 rs.MoveNext

Loop



End Sub

Private Sub mnuDDHH_Click()
    ExportarChandonddhh

End Sub

Private Sub mnuEspañol_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Documento As Long
Dim Nombre As String
Dim P As Integer



Sql = "SELECT DOCUMENTOS_DIGITALES.ID"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE  (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1)  "
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)"
rs.Open Sql, strConBasa
Do While Not rs.EOF

Sql = " SELECT     ID, ENBASA, letra_desde , nro_desde, BatchPgDta"
Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
Sql = Sql & " ORDER BY ID"
Debug.Print rs!ID

Set rs2 = New ADODB.Recordset

rs2.Open Sql, strConBasa

Dim i As Integer
If Not rs2.EOF Then
    If IsNumeric(rs2!NRO_DESDE) Then
                
                If Not IsNull(rs2!NRO_DESDE) Then
                    Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
                    Sql = Sql & " SET "
                    Sql = Sql & " NRO_DESDE =" & rs2!NRO_DESDE
                    Sql = Sql & ", NRO_HASTA =" & rs2!NRO_DESDE
                    Sql = Sql & " Where ID = " & rs!ID
                    Sql = Sql & " and NRO_DESDE IS NULL"
                    ExecutarSql Sql
                End If
            End If
  
 End If

 
 rs.MoveNext

Loop


MsgBox "Terminado"

End Sub

Private Sub mnuExpAndesmarHojasDeRutas_Click()


    Dim Sql As String
    Dim Caja As Long
    Dim Encabezado As String
    Dim NombreArchivoImagen As String
    Dim NombreArchivoImagenPDF As String
    Dim NombreArchivoImagenTIF As String
    Dim NombreArchivoTxt As String
    
    Dim LoteNombreArchivoTxt As String
    Dim PasoInicio As String
    Dim PasoImagenes As String
    Dim PasoExpImagenes  As String
    Dim PasoExp As String
    Dim Datos As String
    Dim CantImagenes As Long
    Dim ControlXML As String
    Dim PasoExpRaiz As String
    Dim LoteAndesmar As String
    
    
    Set rsGrilla = New ADODB.Recordset
    
LoteAndesmar = InputBox("INGRESE EL LOTE")

Sql = " SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.CANTIDAD_IMAGENES , DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS"
Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES INNER JOIN "
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE ON "
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & vbCrLf & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 403) "
Sql = Sql & vbCrLf & " AND DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in ( " & LoteAndesmar & ")"
Sql = Sql & vbCrLf & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE > 100)"
Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.ID"
    
    
    
    Set rsGrilla = New ADODB.Recordset
        rsGrilla.Open Sql, strConBasa
        Caja = rsGrilla!FK_CAJAS
        LoteNombreArchivoTxt = "LOTE" & Caja
       
        PasoExp = "I:\0403-ANDESMAR EXPRESS\0008-HOJAS DE RUTAS\EXPORTADAS" & "\" & "LOTE" & Caja
    If Dir(PasoExp, vbDirectory) = "" Then
        FileSystem.MkDir PasoExp
        PasoExpRaiz = PasoExp
        PasoExpImagenes = PasoExpRaiz & "\IMAGENES"
        FileSystem.MkDir PasoExpImagenes
     Else
        MsgBox "Los directorios ya existen"
        Exit Sub
    End If

    NombreArchivoTxt = PasoExpRaiz & "\" & LoteNombreArchivoTxt & ".TXT"
    Open NombreArchivoTxt For Append As #1
    Encabezado = Chr(34) & "DocumentFileName" & Chr(34) & "," & Chr(34) & "PageCount" & Chr(34) & "," & Chr(34) & "Guia" & Chr(34) & "," & Chr(34) & "Fecha" & Chr(34) & "," & Chr(34) & "Sucursal" & Chr(34) & "," & Chr(34) & "Caja" & Chr(34) & "," & Chr(34) & "Lote" & Chr(34)
    Print #1, Encabezado
     PasoImagenes = "\\222.15.19.251\Imagenes\"
        
        Do While Not rsGrilla.EOF
            CantImagenes = CantImagenes + 1
            
             NombreArchivoImagen = rsGrilla!NRO_DESDE & "_" & Caja & "_" & rsGrilla!ID
             NombreArchivoImagenPDF = NombreArchivoImagen & ".pdf"
            NombreArchivoImagenTIF = NombreArchivoImagen & ".TIF"
            Datos = Chr(34) & "//" & LoteNombreArchivoTxt & "/IMAGENES/" & NombreArchivoImagenPDF & Chr(34) & "," & Chr(34) & rsGrilla!Cantidad_Imagenes & Chr(34) & "," & Chr(34) & rsGrilla!NRO_DESDE & Chr(34) & "," & Chr(34) & Format(Now, "DD/MM/YYYY") & Chr(34) & "," & Chr(34) & "MENDOZA" & Chr(34) & "," & Chr(34) & Caja & Chr(34) & "," & Chr(34) & LoteNombreArchivoTxt & Chr(34)
            Print #1, Datos
            
               FileCopy PasoImagenes & BuscarDirectorioPaso(rsGrilla!ID) & "\" & rsGrilla!ID & ".tif", PasoExpImagenes & "\" & NombreArchivoImagenTIF
             
            rsGrilla.MoveNext
        Loop
    Close #1
    ControlXML = "<Batch>" & vbCr
    ControlXML = ControlXML & " <Statistics>" & vbCr
    ControlXML = ControlXML & "     <DocumentCount>" & CantImagenes & "</DocumentCount>" & vbCr
    ControlXML = ControlXML & " </Statistics>" & vbCr
    ControlXML = ControlXML & "</Batch>" & vbCr
    Open PasoExpRaiz & "\" & LoteNombreArchivoTxt & ".XML" For Append As #1
        Print #1, ControlXML
    Close #1
    MsgBox "terminados"






'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'Dim rs As New ADODB.Recordset
'        Dim Sql As String
'        Dim PasoInicial As String
'        Dim i As Integer
'        Dim sgrabar As String
'        Dim DATO  As String
'        Dim NombreArchivo As String
'        Dim Paso As String
'        Dim Año As String
'        Dim Mes As String
'        Dim Dia As String
'        Dim FechaActulizacion As String
'        Dim lote As String
'
'lote = InputBox("Ingrese el lote")
'    If Not IsNumeric(lote) Then
'
'     MsgBox "el lote no existe"
'     Exit Sub
'
'
'    End If
'
'            Sql = " SELECT DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_DESDE"
'            Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES INNER JOIN"
'            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE ON"
'            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
'            Sql = Sql & vbCrLf & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 403) "
'            Sql = Sql & vbCrLf & " and DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in ( " & lote & ")"
'            Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.ID"
'Paso = txtPasoImagenesFinal.Text
'
'      rs.Open Sql, strConBasa
'
'    Do While Not rs.EOF
'
'                DATO = rs!NRO_DESDE
'                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", Paso & DATO & ".tif"
'
'
'        rs.MoveNext
'    Loop
'
'           MsgBox "Terminado"
End Sub

Private Sub mnuExportacion_Click()
ChandonExportacion
End Sub

Private Sub mnuExportaMulti_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim Nombre As String
    Dim Indice As String

Sql = " SELECT  ID ,   BatchNo,  Suspense_File, BatchPgDta"
Sql = Sql & vbCrLf & " From TELEFORM_DIGITAL"
Sql = Sql & vbCrLf & " Where (BatchNo = 1030)"
Sql = Sql & vbCrLf & " ORDER BY BatchPgDta"


rs.Open Sql, strConBasa
            
            
            Do While Not rs.EOF
            
            Indice = Replace(Mid(rs!BatchPgDta, 14), "[", "")
            Indice = Replace(Indice, "]", "")
              Indice = Format(Indice, "0000")
                Nombre = Mid(rs!BatchPgDta, 1, 9) & "_" & Indice
                Rem FileCopy "\\PCTELEMEMO1" & Mid(rs!Suspense_File, 3), "I:\0403-ANDESMAR EXPRESS\0003-GUIAS CONFORMADAS\PARA TELEFORM\1036571\" & Nombre & ".tif"
                
               Sql = " Update TELEFORM_DIGITAL"
                Sql = Sql & " SET BatchPgDta ='" & Nombre & ".tif'"
                Sql = Sql & "  Where (BatchNo = 1030) And ID =" & rs!ID
                ExecutarSql Sql
                rs.MoveNext
            Loop



MsgBox "Terminado"



End Sub

Private Sub mnuExportarEnvio_Click()
ExpresoLujanExportarCodigo
End Sub

Private Sub mnuExportarEspañol_Click()
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim PasoInicial As String
        Dim i As Integer
        Dim sgrabar As String
        Dim DATO  As String
        Dim NombreArchivo As String
            Sql = " SELECT DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.NRO_DESDE "
            Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES INNER JOIN"
            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE ON"
            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
            Sql = Sql & vbCrLf & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1) "
            Sql = Sql & vbCrLf & " and DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in( " & InputBox("Ingrese el lote") & ")"
            Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.ID"
            rs.Open Sql, strConBasa
            Do While Not rs.EOF
                If IsNull(rs!LETRA_DESDE) Then
                    DATO = "E" & Format(rs!NRO_DESDE, "0000000") & "_01"
                Else
                    DATO = "E" & Format(rs!NRO_DESDE, "0000000") & "_0" & Trim(rs!LETRA_DESDE)
                End If
                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", txtPasoImagenesFinal.Text & "/" & DATO & ".tif"
                rs.MoveNext
            Loop
           MsgBox "Terminado"
End Sub

'Private Sub mnuExportar Hillebrand_Click()
'ExportarHilebrand
'End Sub
'
Private Sub LaCajaSinTrack()
    Dim Sql As String
    Dim rsImagenes As New ADODB.Recordset
    Dim Carpeta As String

    
    MousePointer = 11
    Dim Lotes As String
    
    
        
        Carpeta = InputBox("Carpeta de salida", "", "D:\ExportarImagenes\")
'        Do While Not rsBuscar.EOF
'           Lotes = Lotes & "," & rsBuscar!Lote
'           rsBuscar.MoveNext
'        Loop
       
        Sql = " SELECT     ID, NRO_DESDE, DIRECTORIO_PASO"
        Sql = Sql & "  From DOCUMENTOS_DIGITALES "
        Sql = Sql & "  WHERE FK_DOCUMENTOS_DIGITALES_LOTE IN (" & Mid(txtLotesExportar, 2) & ")"
        Sql = Sql & "  ORDER BY ID"
            
        Set rsImagenes = New ADODB.Recordset
        rsImagenes.Open Sql, ConActiva, 0, 1
        
            
           Rem  FileSystem.MkDir Carpeta
            
            Do While Not rsImagenes.EOF
                If chkInvertirNombre.value = 1 Then
                    FileCopy PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif", Carpeta & "\" & rsImagenes!ID & "_" & Trim(rsImagenes!NRO_DESDE) & ".TIF"
                Else
                    FileCopy PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif", Carpeta & "\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
                End If
                rsImagenes.MoveNext
            Loop
       

MousePointer = 0
MsgBox "Operacion terminada"
End Sub

Private Sub mnuExportarHilebrand_Click()


'    Dim Min As Long
'    Dim Max As Long
'    Dim RS As New ADODB.Recordset
'    Dim SQL As String
'    Dim PasoInicial As String
'    PasoInicial = txtPasoImagenesFinal
'    Dim i As Integer
'    Dim R As Excel.Range
'    Dim h As Excel.Hyperlinks
'    i = 1
'    Dim sgrabar As String
'    SQL = " SELECT     INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_HASTA,"
'    SQL = SQL & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO,   DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS AS NRO_CAJA,"
'    SQL = SQL & vbCrLf & " DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.LETRA_DESDE AS Expr1, DOCUMENTOS_DIGITALES.LETRA_HASTA,"
'    SQL = SQL & vbCrLf & " DOCUMENTOS_DIGITALES.Lote , DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES.Nombre"
'    SQL = SQL & vbCrLf & "  FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
'    SQL = SQL & vbCrLf & " DOCUMENTOS_DIGITALES ON"
'    SQL = SQL & vbCrLf & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
'    SQL = SQL & vbCrLf & " INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
'    SQL = SQL & vbCrLf & " WHERE   (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 84) "
'    SQL = SQL & vbCrLf & "  AND DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in( " & Mid(txtLotesExportar, 2) & ") "
'    SQL = SQL & vbCrLf & "  ORDER BY DOCUMENTOS_DIGITALES.NRO_CAJA"
'    Dim NombreArchivo As String
'    i = 2
'    RS.Open SQL, ConActiva, 0, 1
'        Do While Not RS.EOF
'                i = i + 1
'                NombreArchivo = Trim(RS!LETRA_DESDE) & "_" & CStr(RS!ID) & ".tif"
'                hojaEx.Cells(i, 1) = NombreArchivo
'                hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\0\" & RS!NRO_CAJA & "\" & NombreArchivo
'
'                If IsNull(RS!LETRA_DESDE) Then
'                    hojaEx.Cells(i, 2) = RS!NRO_DESDE
'                Else
'                    hojaEx.Cells(i, 2) = RS!LETRA_DESDE
'                End If
'                hojaEx.Cells(i, 2) = RS!ID
'                hojaEx.Cells(i, 3) = RS!NRO_CAJA
'                hojaEx.Cells(i, 4) = Trim(RS!LETRA_DESDE)
'                If Dir(PasoInicial & RS!NRO_CAJA, vbDirectory) = "" Then
'                    FileSystem.MkDir PasoInicial & RS!NRO_CAJA
'                Else
'
'                End If
'
'                 FileCopy PasoImagenes & BuscarDirectorioPaso(RS!ID) & "\" & RS!ID & ".tif", PasoInicial & RS!NRO_CAJA & "\" & NombreArchivo
'                hojaEx.Cells(i, 4) = RS!NRO_CAJA
'                hojaEx.Cells(i, 5) = RS!ID
'
'                RS.MoveNext
'            Loop
'
'           libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY_SS") & ".xls"
'
'           libroEx.Close
'           ApExcel.Quit
'           Set ApExcel = Nothing
'           Set libroEx = Nothing
'           MsgBox "terminado"
'
  


End Sub

Private Sub mnuExportCentroCard_Click()
ExportarCentroCard
End Sub

Private Sub mnuExpresoLujan_Click()
  ExpresoLujanImportarCodigo
 
 Rem  EXPRESOCONTROLCOMPLETO
 Rem  nExpresoLujanControlCodigo
  Rem ExpresoLujanSacarImagenesConError
Rem ActualizarExpreso

Rem  E xpresoLujanCodigo
Rem  ExpresoLujan2
'
End Sub

Private Sub mnuFactura_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Etiqueta As Long
Dim Documento As Long
Dim Nombre As String
Dim P As Integer






Sql = " SELECT DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.NRO_HASTA, DOCUMENTOS_DIGITALES.ID"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 83)"
Sql = Sql & "  And (DOCUMENTOS_DIGITALES.NRO_DESDE Is Null) And (DOCUMENTOS_DIGITALES.NRO_HASTA Is Null)"
Sql = Sql & "  And  DOCUMENTOS_DIGITALES.ID > 1756604"
Sql = Sql & " order by DOCUMENTOS_DIGITALES.ID "

rs.Open Sql, strConBasa
Do While Not rs.EOF

Sql = " SELECT     ID, FACTURA, FECHA, ENBASA,  BatchPgDta"
Sql = Sql & " From basasql.dbo.TELEFORM_FACTURA"
Sql = Sql & " WHERE  ( BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
Sql = Sql & " ORDER BY ID "

Set rs2 = New ADODB.Recordset

rs2.Open Sql, strConBasa

Dim i As Integer
    If Not rs2.EOF Then
        If IsNumeric(rs2!FACTURA) Then
            Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
            Sql = Sql & " SET "
            Sql = Sql & " NRO_DESDE =" & rs2!FACTURA
            Sql = Sql & " , NRO_HASTA =" & rs2!FACTURA
            If Not IsNull(rs2!fecha) Then
                Sql = Sql & " , DESCRIPCION ='" & Replace(rs2!fecha, "'", "") & "'"
            End If
            Sql = Sql & " Where ID = " & rs!ID
            ExecutarSql Sql
        
        End If
    End If
 

 
 rs.MoveNext

Loop


MsgBox "Terminado"


End Sub

Private Sub mnuhojaderuta_Click()
 Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim Etiqueta As Long
    Dim Documento As Long
    Dim Nombre As String
    Dim P As Integer
    
    
    
        Sql = " SELECT DOCUMENTOS_DIGITALES.ID "
        Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & " DOCUMENTOS_DIGITALES ON"
        Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & " WHERE  (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 403)"
         Rem Sql = Sql & " AND   DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE > 23031"
        Sql = Sql & " Order by DOCUMENTOS_DIGITALES.ID "
        
        
        
        
        rs.Open Sql, strConBasa
        
        
        
        Do While Not rs.EOF
            Sql = " SELECT     ID, ENBASA, NRO_DESDE2, NRO_HASTA , BatchPgDta"
            Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
            Sql = Sql & " WHERE  (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
            Sql = Sql & " ORDER BY ID"
            Set rs2 = New ADODB.Recordset
            rs2.Open Sql, strConBasa
            Dim i As Integer
                   If Not rs2.EOF Then
                       If IsNumeric(rs2!NRO_DESDE2) Then
                           Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
                           Sql = Sql & " SET "
                           Sql = Sql & " NRO_DESDE =" & rs2!NRO_DESDE2
                           'If Len(rs2!NRO_HASTA) < 9 Then
                           'Sql = Sql & ", NRO_HASTA =" & rs2!NRO_HASTA
                           'Else
                           'Sql = Sql & ", NRO_HASTA = 0" & Mid(rs2!NRO_HASTA, 6, 7)
                           'End If
                           Sql = Sql & " Where ID = " & rs!ID
                           ExecutarSql Sql
                       End If
                   End If
            rs.MoveNext
        Loop
        MsgBox "Terminado"

End Sub

Private Sub mnuImportarCentroCard_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim Etiqueta As Long
    Dim Documento As Long
    Dim Nombre As String
    Dim P As Integer
        Sql = " SELECT DOCUMENTOS_DIGITALES.ID "
        Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & " DOCUMENTOS_DIGITALES ON"
        Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & " WHERE  (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 279)"
        Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)"
        Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_HASTA IS NULL)"
        Rem Sql = Sql & " AND   DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE > 23031"
        Sql = Sql & " Order by DOCUMENTOS_DIGITALES.ID "
        rs.Open Sql, strConBasa
        Do While Not rs.EOF
            Sql = " SELECT     ID, ENBASA, NRO_DESDE, NRO_HASTA , BatchPgDta"
            Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
            Sql = Sql & " WHERE  (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
            Sql = Sql & " ORDER BY ID"
            Set rs2 = New ADODB.Recordset
            rs2.Open Sql, strConBasa
            Dim i As Integer
                   If Not rs2.EOF Then
                       If IsNumeric(rs2!NRO_DESDE) Then
                           Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
                           Sql = Sql & " SET "
                           Sql = Sql & " NRO_DESDE =" & rs2!NRO_DESDE
                           'If Len(rs2!NRO_HASTA) < 9 Then
                           'Sql = Sql & ", NRO_HASTA =" & rs2!NRO_HASTA
                           'Else
                           'Sql = Sql & ", NRO_HASTA = 0" & Mid(rs2!NRO_HASTA, 6, 7)
                           'End If
                           Sql = Sql & " Where ID = " & rs!ID
                           ExecutarSql Sql
                       End If
                   End If
            rs.MoveNext
        Loop
        MsgBox "Terminado"
End Sub

Private Sub mnuImportMuni_Click()
  Dim Sql As String
        Dim rs As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
        Dim Documento As Long
        Dim Nombre As String
        Dim P As Integer
        Dim i As Integer
        Dim Codigo As String
        Dim conDoc As ADODB.Connection
        
       Sql = " SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER,"
       Sql = Sql & "               DOCUMENTOS_DIGITALES.Exportado , DOCUMENTOS_DIGITALES.LETRA_HASTA"
Sql = Sql & " FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & "                       DOCUMENTOS_DIGITALES ON"
 Sql = Sql & "                     DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"

Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1156) And  (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL) "
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER DESC"
        
     Set conDoc = New ADODB.Connection
        conDoc.Open strConBasa

            Set rs = New ADODB.Recordset
            rs.Open Sql, conDoc
            
            Do While Not rs.EOF
                Sql = " SELECT     ID, ENBASA, NRO_DESDE , NRO_HASTA ,BatchPgDta"
                Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
                Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
                Sql = Sql & " ORDER BY ID DESC "
                Set rs2 = New ADODB.Recordset
                rs2.Open Sql, strConBasa
                Codigo = 0
                If Not rs2.EOF Then
                     If Not IsNull(rs2!NRO_DESDE) Then
                                        Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
                                        Sql = Sql & " SET "
                                        Sql = Sql & "  NRO_DESDE =" & rs2!NRO_DESDE
                                        Sql = Sql & " Where ID = " & rs!ID
                                        conDoc.Execute Sql
                    End If
                    If Not IsNull(rs2!NRO_HASTA) Then
                                        Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
                                        Sql = Sql & " SET "
                                        Sql = Sql & "  NRO_HASTA =" & rs2!NRO_HASTA
                                        Sql = Sql & " Where ID = " & rs!ID
                                        ExecutarSql Sql
                    End If
                End If
                
                rs.MoveNext
            Loop
            MsgBox "Terminado"
End Sub

Private Sub mnuJFH_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Documento As Long
Dim Nombre As String
Dim P As Integer



Sql = "SELECT DOCUMENTOS_DIGITALES.ID"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 84"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.DESCRIPCION IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.LETRA_DESDE IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.FECHA_DESDE IS NULL)"

rs.Open Sql, strConBasa
Do While Not rs.EOF

Sql = " SELECT     ID, ENBASA, letra_desde, BatchPgDta"
Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
Sql = Sql & " ORDER BY ID"

Set rs2 = New ADODB.Recordset

rs2.Open Sql, strConBasa

Dim i As Integer
If Not rs2.EOF Then

If Not IsNull(rs2!LETRA_DESDE) Then
   
    Nombre = UCase(Trim(rs2!LETRA_DESDE))
    
    Nombre = "'" & Nombre & "'"
 Else
 Nombre = "'NULL'"
 End If
 
 
     Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
Sql = Sql & " SET "
Sql = Sql & " LETRA_DESDE =" & Nombre
Sql = Sql & ", LETRA_HASTA =" & Nombre
Sql = Sql & " Where ID = " & rs!ID
 
ExecutarSql Sql
 
 End If
 

 
 rs.MoveNext

Loop


MsgBox "Terminado"
End Sub


Private Sub mnuMedife_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim DATO As String


Sql = " SELECT LEGAJOS.LETRA_DESDE, LEGAJOS.LETRA_HASTA, LEGAJOS.NRO_DESDE, LEGAJOS.NRO_HASTA, FK_LEGAJO_ETIQUETA ,"
Sql = Sql & "  DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES.ID AS ID_IMAGEN, FK_CAJAS ,"
Sql = Sql & "  DOCUMENTOS_DIGITALES.DIRECTORIO_PASO"
Sql = Sql & "  FROM LEGAJOS INNER JOIN"
Sql = Sql & "  DOCUMENTOS_DIGITALES_LOTE ON LEGAJOS.NRO_CAJA = DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS INNER JOIN"
Sql = Sql & "  DOCUMENTOS_DIGITALES ON DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE AND"
Sql = Sql & "  LEGAJOS.Etiqueta = SUBSTRING(DOCUMENTOS_DIGITALES.FK_LEGAJO_ETIQUETA, 1, 12)"
Sql = Sql & "  Where (LEGAJOS.NRO_CAJA = 1318775)"

rs.Open Sql, strConBasa

Do While Not rs.EOF
 DATO = Format(rs!NRO_DESDE, "00000000") & " " & Trim(rs!LETRA_DESDE) & " " & Trim(rs!FK_LEGAJO_ETIQUETA) & ".PDF"
 FileCopy "\\222.15.19.251\ImagenesPDF\" & Trim(rs!DIRECTORIO_PASO) & "\" & rs!ID_imagen & ".PDF", "C:\ExportarImagenes\" & rs!FK_CAJAS & "\" & DATO
rs.MoveNext
Loop




End Sub

Private Sub mnuOsdeDiabeticos_Click()
Dim Min As Long
    Dim Max As Long
    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet


    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Add
    Set hojaEx = libroEx.Worksheets.Item(1)
    Dim rs As New ADODB.Recordset
    Dim Sql As String



    Dim i As Integer
    Dim R As Excel.Range
    Dim h As Excel.Hyperlinks
    i = 1

        hojaEx.Columns("A:A").ColumnWidth = 3.57
        hojaEx.Columns("B:B").ColumnWidth = 7.14
        hojaEx.Columns("C:C").ColumnWidth = 15.57
        hojaEx.Columns("C:C").NumberFormat = "@"
        hojaEx.Columns("D:D").ColumnWidth = 8.57
        hojaEx.Columns("E:E").ColumnWidth = 11.23
        hojaEx.Columns("G:G").ColumnWidth = 9.43
        hojaEx.Columns("H:H").ColumnWidth = 7.57
        hojaEx.Columns("I:I").ColumnWidth = 6.71
        hojaEx.Columns("I:I").NumberFormat = "@"
        hojaEx.Columns("J:J").ColumnWidth = 6.71
        hojaEx.Range("J1:j1000").Select
        hojaEx.Columns("J:J").NumberFormat = "@"
        hojaEx.Columns("K:K").ColumnWidth = 7.43
        hojaEx.Columns("K:K").NumberFormat = "@"


            hojaEx.Cells(1, 1) = "i04t"
            hojaEx.Cells(1, 2) = "filial"
            hojaEx.Cells(1, 3) = "nomfoto"
            hojaEx.Cells(1, 4) = "cantfotos"
            hojaEx.Cells(1, 5) = "ic"
            hojaEx.Cells(1, 6) = "filler0"
            hojaEx.Cells(1, 7) = "feccarg"
            hojaEx.Cells(1, 8) = "nrobasa"
            hojaEx.Cells(1, 9) = "criterio"
            hojaEx.Cells(1, 10) = "nrotram"
            hojaEx.Cells(1, 11) = "nrotram"






        Dim sgrabar As String



Sql = "  SELECT     DOCUMENTOS_DIGITALES.COD_CLIENTE, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.NRO_DESDE,"
Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES.ID , DOCUMENTOS_DIGITALES.NRO_CAJA, DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION"
Sql = Sql & vbCrLf & "   FROM         DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & vbCrLf & "   DOCUMENTOS_DIGITALES_LOTE ON"
Sql = Sql & vbCrLf & "   DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & vbCrLf & " WHERE     (DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE IN (" & InputBox("Ingrese los numeros de lotes separadospor ,") & "))"
Sql = Sql & vbCrLf & "   ORDER BY DOCUMENTOS_DIGITALES.ID"



        i = 1
        Dim ValorDiabetico As String
        If InputBox("Ingrese 1 Para diabeticos y 2 para Psicopatologias") = "1" Then
                    ValorDiabetico = "DB"
                Else
                    ValorDiabetico = "SR"
                End If

        rs.Open Sql, ConActiva, 0, 1
Dim directorio As String
directorio = InputBox("Directorio de Exportacion")
            Do While Not rs.EOF
                i = i + 1

                If Dir(txtPasoImagenesFinal & directorio, vbDirectory) = "" Then
                 FileSystem.MkDir txtPasoImagenesFinal & directorio
                 FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", txtPasoImagenesFinal & directorio & "\" & Format(rs!ID, "0000000000000") & ".tif"
                Else
                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", txtPasoImagenesFinal & directorio & "\" & Format(rs!ID, "0000000000000") & ".tif"
                End If
                sgrabar = "i04t"
                hojaEx.Cells(i, 1) = sgrabar
                hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), CStr(Format(rs!ID, "0000000000000")) & ".jpg"
                hojaEx.Cells(i, 2) = rs!ID
                hojaEx.Cells(i, 3) = "'" & CStr(Format(rs!ID, "0000000000000"))
                hojaEx.Cells(i, 4) = "1"
                hojaEx.Cells(i, 5) = Trim(rs!LETRA_DESDE)
                hojaEx.Cells(i, 6) = 1
                Rem rs!nro_desde
                hojaEx.Cells(i, 7) = directorio
                hojaEx.Cells(i, 8) = CStr(rs!NRO_CAJA) & "11"
                hojaEx.Cells(i, 9) = ValorDiabetico
                hojaEx.Cells(i, 10) = "0"
                rs.MoveNext
            Loop





           libroEx.SaveAs txtPasoImagenesFinal.Text & directorio & ".xls"
           libroEx.Close
           ApExcel.Quit
        Set ApExcel = Nothing
        Set libroEx = Nothing
End Sub

Private Sub mnuPorIdImagen_Click()
        
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim Documento As Long
    Dim Nombre As String
    Dim P As Integer
    Dim i As Integer
    Dim Codigo As String
    
        
        
        Sql = " SELECT DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER, "
        Sql = Sql & " DOCUMENTOS_DIGITALES.Exportado , DOCUMENTOS_DIGITALES.LETRA_HASTA "
        Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN "
        Sql = Sql & " DOCUMENTOS_DIGITALES ON "
        Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE "
       Rem  Sql = Sql & " Where  (DOCUMENTOS_DIGITALES.ID in (" & InputBox("Ingrese los id separados por , ") & "))"
        Rem Sql = Sql & " Where  DOCUMENTOS_DIGITALES.ID in (  )"
       Sql = Sql & " WHERE ( DOCUMENTOS_DIGITALES.ID  IN (2041253, 2041473, 2048493, 2048493, 2048494, 2048494, 2055085, 2075106, 2075106, 2081256, 2081256, 2081256, 2081337, 2081337, 2089336, 2089336,"
       Sql = Sql & "                  2089336, 2095024, 2095024, 2095160, 2095164, 2122966, 2122966, 2122981, 2122981, 2123061, 2123061, 2129989, 2129989, 2135553, 2135553, 2136008,"
       Sql = Sql & "                  2145898, 2145898, 2145942, 2145942, 2146082, 2150073, 2150073, 2150113, 2150497, 2157933, 2157933, 2166906, 2166906, 2167013, 2167121, 2179385,"
       Sql = Sql & "                  2179385, 2179558, 2184514, 2184514, 2201809, 2201809, 2201938, 2207819, 2207819, 2225867, 2225867, 2226089, 2232540, 2232540, 2239249, 2239249,"
       Sql = Sql & "                    2239289, 2239289, 2239361, 2265101, 2265791, 2265791, 2265829, 2265829, 2273949, 2273949, 2274060, 2313775, 2313775, 2332530, 2332530, 2332556,"
       Sql = Sql & "                    2332556, 2332638, 2332670, 2332681, 2332712, 2340213, 2344530, 2344552, 2344552, 2344584, 2344584, 2344601, 2344601, 2344602, 2344655, 2354859,"
       Sql = Sql & "                    2354859, 2354891, 2354891, 2354979, 2360688, 2360688, 2360694, 2360694, 2360796, 2372734, 2384080, 2411705))"
        
        Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER DESC "
        
       Rem  /****** Script for SelectTopNRows command from SSMS  ******/
 Sql = " SELECT        DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_INICIAL20180131_BORRAR20180216.NRO_DESDE, DOCUMENTOS_DIGITALES_INICIAL20180131_BORRAR20180216.LETRA_DESDE,"
  Sql = Sql & "                       DOCUMENTOS_DIGITALES_INICIAL20180131_BORRAR20180216.ID , DOCUMENTOS_DIGITALES_INICIAL20180131_BORRAR20180216.DIRECTORIO_PASO"
Sql = Sql & " FROM            DOCUMENTOS_DIGITALES_INICIAL20180131_BORRAR20180216 INNER JOIN"
 Sql = Sql & "                         DOCUMENTOS_DIGITALES_LOTE ON DOCUMENTOS_DIGITALES_INICIAL20180131_BORRAR20180216.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 279) and  id > 1284183 "
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES_INICIAL20180131_BORRAR20180216.ID"
        
        
        
        rs.Open Sql, strConBasa
            Do While Not rs.EOF
               Rem  FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", txtPasoImagenesFinal & rs!ID & ".tif"
                FileCopy "Y:\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", "Y:\centrocard\" & Format(rs!NRO_DESDE, "000000000") & "_" & rs!ID & ".tif"
                
                rs.MoveNext
            Loop
            MsgBox "Terminado"
            




End Sub

Private Sub mnuReporteRearchivo_Click()
    Dim Sql As String
        Sql = "SELECT     DESCRIPCION, LETRA_DESDE, NRO_DESDE, COD_CLIENTE, REMITO, COD_ESTADO, Nombre"
        Sql = Sql & "  From V_REARCHIVO_DIGITAL"
        Sql = Sql & "  WHERE   REMITO LIKE   '%" & InputBox("ingrese el remito") & "%'"
        frmReportes.ImprimirReporte PasoReportes & "rptRearchivoDigital2.rpt", Sql, True
    

End Sub

Private Sub mnuSalud_Click()
Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet


 Dim Min As Long
    Dim Max As Long
       
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Add
    Set hojaEx = libroEx.Worksheets.Item(1)
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim PasoInicial As String
    PasoInicial = txtPasoImagenesFinal.Text

    Dim i As Integer
    Dim R As Excel.Range
    Dim h As Excel.Hyperlinks
    i = 1
        hojaEx.Cells(1, 1) = "Link Imagen"
        hojaEx.Cells(1, 2) = "Nombre Imagen"
        hojaEx.Cells(1, 3) = "Numero"
        hojaEx.Cells(1, 4) = "Letra"
        hojaEx.Cells(1, 5) = "Año"
        hojaEx.Cells(1, 6) = "Copia"
        Dim sgrabar As String
        Sql = " SELECT     INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_HASTA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO,DOCUMENTOS_DIGITALES.fecha_desde  , DOCUMENTOS_DIGITALES.descripcion ,   DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS AS NRO_CAJA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.LETRA_DESDE AS Expr1, DOCUMENTOS_DIGITALES.LETRA_HASTA,"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.Lote , DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES.Nombre"
        Sql = Sql & vbCrLf & "  FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES ON"
        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & vbCrLf & " INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
        Sql = Sql & vbCrLf & "  Where      DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in( " & Mid(txtLotesExportar, 2) & ") "
        Sql = Sql & vbCrLf & "  ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE "
        Dim NombreArchivo As String
        i = 1
        MousePointer = 11
                
                rs.Open Sql, ConActiva, 0, 1
                
                
                If Dir(PasoInicial, vbDirectory) = "" Then
                    FileSystem.MkDir PasoInicial
                Else

                End If


            Do While Not rs.EOF
                i = i + 1
                NombreArchivo = Trim(rs!NRO_DESDE) & "-" & Trim(rs!LETRA_DESDE) & "-" & Format(Trim(rs!FECHA_DESDE), "YYYY") & "_     " & CStr(rs!ID) & ".tif"
                hojaEx.Cells(i, 1) = NombreArchivo
                hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\" & NombreArchivo
                
           hojaEx.Cells(i, 2) = Trim(rs!NRO_DESDE) & "-" & Trim(rs!LETRA_DESDE) & "-" & Format(Trim(rs!FECHA_DESDE), "YYYY")
        hojaEx.Cells(i, 3) = Trim(rs!NRO_DESDE)
        hojaEx.Cells(i, 4) = Trim(rs!LETRA_DESDE)
        hojaEx.Cells(i, 5) = Format(Trim(rs!FECHA_DESDE), "yyyy")
        hojaEx.Cells(i, 6) = Trim(rs!Descripcion)
                               

                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & NombreArchivo
                
                rs.MoveNext
            Loop

           libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY_ss") & ".xls"
           libroEx.Close
           ApExcel.Quit
           Set ApExcel = Nothing
           Set libroEx = Nothing
           MousePointer = 0
           
End Sub

Private Sub mnuSinTrack_Click()
LaCajaSinTrack
End Sub

Private Sub mnuZucardi_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Documento As Long
Dim Nombre As String
Dim P As Integer



Sql = "SELECT DOCUMENTOS_DIGITALES.ID"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1123"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID "

Dim Sucursal As Integer
Dim REMITO As Long
Dim fecha As String
Dim Cliente As String



rs.Open Sql, strConBasa
Do While Not rs.EOF

Sql = " SELECT     ID, ENBASA,NRO_DESDE ,BatchPgDta"
Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
Sql = Sql & " ORDER BY ID"

Set rs2 = New ADODB.Recordset

rs2.Open Sql, strConBasa

Dim i As Integer
        If Not rs2.EOF Then
            If IsNumeric(Trim(rs2!NRO_DESDE)) Then
                Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
                Sql = Sql & " SET "
                Sql = Sql & " NRO_DESDE =" & rs2!NRO_DESDE
                Sql = Sql & " Where ID = " & rs!ID
                ExecutarSql Sql
            End If
        End If
 

 
 rs.MoveNext

Loop


MsgBox "Terminado"


End Sub

Private Sub mnuZucardiFactura_Click()
Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim PasoInicial As String
        Dim i As Integer
        Dim sgrabar As String
        Dim DATO  As String
        Dim NombreArchivo As String
        Dim Paso As String
        Dim Año As String
        Dim Mes As String
        Dim Dia As String
        Dim FechaActulizacion As String
        Dim lote As String

lote = InputBox("Ingrese el lote")
    If Not IsNumeric(lote) Then
    
     MsgBox "el lote no existe"
     Exit Sub
     
    
    End If
    
            Sql = " SELECT DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_DESDE"
            Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES INNER JOIN"
            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE ON"
            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
            Sql = Sql & vbCrLf & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1123) "
            Sql = Sql & vbCrLf & " and DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in ( " & lote & ")"
            Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.ID"
Paso = txtPasoImagenesFinal.Text
      
      rs.Open Sql, strConBasa
      
    Do While Not rs.EOF
        
                DATO = Trim(rs!LETRA_DESDE)
                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", Paso & DATO & ".tif"
        
        
        rs.MoveNext
    Loop

           MsgBox "Terminado"
End Sub

Private Sub MnuZucardiOrdenes_Click()
Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim PasoInicial As String
        Dim i As Integer
        Dim sgrabar As String
        Dim DATO  As String
        Dim NombreArchivo As String
        Dim Paso As String
        Dim Año As String
        Dim Mes As String
        Dim Dia As String
        Dim FechaActulizacion As String
        Dim lote As String

Rem lote = InputBox("Ingrese el lote")
'    If Not IsNumeric(lote) Then
'
'     MsgBox "el lote no existe"
'     Exit Sub
     
    
   Rem  End If
    Rem 1114415 , 1114419
            Sql = " SELECT DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_DESDE, FK_legajo_etiqueta , FK_CAJAS "
            Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES INNER JOIN"
            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE ON"
            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
            Sql = Sql & vbCrLf & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 1123) "
            Sql = Sql & vbCrLf & " and DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS in (" & InputBox("ingresela Caja", "Cajas ,0") & ")"
            Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.ID"
Paso = txtPasoImagenesFinal.Text
      
      rs.Open Sql, strConBasa
      
    Do While Not rs.EOF
        If Dir(Paso & "\" & rs!FK_CAJAS, vbDirectory) = "" Then
    MkDir Paso & "\" & rs!FK_CAJAS
        End If
        
        
        
                DATO = Format(rs!NRO_DESDE, "0000000") & "_" & rs!FK_LEGAJO_ETIQUETA
                FileCopy "\\222.15.19.251\ImagenesPDF\" & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".PDF", Paso & "\" & rs!FK_CAJAS & "\" & DATO & ".PDF"
        
        
        rs.MoveNext
    Loop

           MsgBox "Terminado"

End Sub

Private Sub MuniGodoyCruzPersonal_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset

    Sql = " SELECT DOCUMENTOS_DIGITALES.DIRECTORIO_PASO, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES.ID,"
    Sql = Sql & " DOCUMENTOS_DIGITALES.FK_LEGAJO_ETIQUETA "
    Sql = Sql & " FROM  DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON"
    Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS = 1105358)"
    Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.NRO_HASTA"

    Do While Not rs.EOF
        
'        If Dir("\\222.15.19.251\ImagenesPDF\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".pdf") <> "" Then
'            FileSystem.FileCopy ("D:\ExportarImagenes\" & rs!FK_CAJAS & "\" & rs!FK_LEGAJO_ETIQUETA & ".PDF")
'
'        Else
'            MsgBox " NO existe archivo " & rs!DIRECTORIO_PASO & "\" & rs!ID & ".pdf"
'        End If
'
        
    
    
        rs.MoveNext
    Loop
    



End Sub

Private Sub SSTab1_DblClick()
CopiarDatosGrilla GRDHI
End Sub

Private Sub txtDato_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo salir
Dim Sql As String
    If KeyAscii = 13 Then
        Dim i As Integer
        Dim Cantidad As Integer
            For i = 0 To fraCampos.Count - 1
                If fraCampos.Item(i).Visible = True Then
                   Cantidad = i
                   
                End If
                
            Next
        
        If Index + 1 > Cantidad Then
        
        If chkCopiarLetra_Numero.value = 1 Then
            Copiar_Letra_Numero
        End If
            If IsNull(ctlPersonalIndexacion.Valor) Then
                MsgBox "FALTA EL PERSONAL INDEXACION"
                Exit Sub
            End If
                Sql = " Update DOCUMENTOS_DIGITALES"
                Sql = Sql & " SET FECHA_INDEXACION = " & SysDateMinutoSegundo
                Sql = Sql & ",PERSONAL_INDEXACION = " & ctlPersonalIndexacion.Valor
                Sql = Sql & "  Where ID = " & rsGrilla.Fields(0).value
                ExecutarSql Sql
                Rem rsGrilla.Update
                ProximaImagen
                txtDato.Item(0).SetFocus
        Else
        txtDato.Item(Index + 1).SetFocus
        End If
    End If
    
    If KeyAscii = 43 Then
        For i = 0 To txtDato.Count - 1
           txtDato.Item(i).FontSize = CInt(txtDato.Item(i).FontSize) + 1
           txtDato.Item(i).Refresh
        Next
        KeyAscii = 0
    End If

    
Exit Sub
salir:
If Err.Number = 6160 Then
MsgBox "Fin de archivo"
Else
MsgBox Err.Description
End If

    

End Sub

Public Sub ProximaImagen()
 On Error GoTo salir
 Dim posci As Long
 Dim Paso As String
If rsGrilla.EOF Then
Else

       rsGrilla.MoveNext
    
 Rem grdIndexarImagenes.Rebind
    
     If optImagenLocal.value = 0 Then
               Paso = PasoImagenes & "\" & rsGrilla!DIRECTORIO_PASO & "\" & rsGrilla!ID & ".TIF"
            Else
               Paso = "C:\ImagenLocal\" & rsGrilla!lote & "\" & rsGrilla!ID & ".TIF"
            End If
            
            If Dir(Paso) <> "" Then
                ctlVerImagenes1.PonerImagen Paso
            Else
                MsgBox "No existe la imagen se copio el paso "
                Clipboard.Clear
                Clipboard.SetText Paso
            End If
    
    End If
   Exit Sub
salir:
If Err.Number = 6160 Then
MsgBox "Fin de archivo"

End If

 
End Sub

Public Sub Copiar_Letra_Numero()
Dim i As Integer
Dim DATO As String
 For i = 0 To txtDato.Count - 1
    If txtDato.Item(i).Tag = "LETRA_DESDE" Then
        DATO = txtDato.Item(i).Text
        Exit For
    End If
 Next
 For i = 0 To txtDato.Count - 1
    If txtDato.Item(i).Tag = "NRO_DESDE" Then
       If IsNumeric(DATO) Then
            txtDato.Item(i).Text = CDbl(DATO)
        End If
        Exit For
    End If
 Next
 
 
End Sub

Public Sub Export_AirLiquide()
     Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
    
    
     Dim Min As Long
        Dim Max As Long
    
        
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim PasoInicial As String
    
        PasoInicial = txtPasoImagenesFinal.Text
        Dim i As Integer
        Dim R As Excel.Range
        Dim h As Excel.Hyperlinks
        i = 1
    Dim directorio  As String
    Dim AnioAnterior As Integer
    AnioAnterior = 0
    
'    If txtLotesExportar.Text = "" Then
'        MsgBox "Ingrese los lotes"
'        Exit Sub
'    End If
      lblCantidad.Caption = 0
    
   Rem  directorio = Format(Now, "DDMMYYYY")

               
    
    
    
            Dim sgrabar As String
    
    
    
    
            Sql = " SELECT    LETRA_DESDE, ID, NRO_HASTA, NRO_DESDE, DIRECTORIO_PASO, NRO_CAJA, FECHA_DESDE, LETRA_DESDE AS Expr1, LETRA_HASTA, LOTE, "
'              Sql = Sql & vbCrLf & "  IMAGEN_ORIGEN, COD_CLIENTE"
'             Sql = Sql & vbCrLf & "  FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
'             Sql = Sql & vbCrLf & "         DOCUMENTOS_DIGITALES ON"
'                      Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
'            Sql = Sql & vbCrLf & " where DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 155)"
'            Rem FK_DOCUMENTOS_DIGITALES_LOTE IN (" & Mid(txtLotesExportar.Text, 2) & ")"
'            Sql = Sql & vbCrLf & " ORDER BY NRO_HASTA"
'
'
'            Sql = " SELECT     DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.DESCRIPCION,"
'            Sql = Sql & vbCrLf & "           DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.LETRA_HASTA, DOCUMENTOS_DIGITALES.NRO_DESDE,"
'            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_HASTA , DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.FECHA_HASTA"
'            Sql = Sql & vbCrLf & " FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
'            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES ON"
'            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
'            Sql = Sql & vbCrLf & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 155)"
'            Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE"
            
            
            Sql = "  SELECT  Lote,IMAGEN_ORIGEN,  ID, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.DESCRIPCION,"
            Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.LETRA_HASTA, DOCUMENTOS_DIGITALES.NRO_DESDE,"
            Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES.NRO_HASTA , DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.FECHA_HASTA ,DIRECTORIO_PASO "
            Sql = Sql & vbCrLf & "  FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
            Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES ON"
            Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
            Sql = Sql & vbCrLf & "  Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 155) "
           Rem Sql = Sql & vbCrLf & " and  FK_DOCUMENTOS_DIGITALES_LOTE IN (" & Mid(txtLotesExportar.Text, 2) & ")"
            Sql = Sql & vbCrLf & "  and     (FECHA_DESDE > CONVERT(DATETIME, '2015-12-31 00:00:00', 102)) "
            Rem SQL = SQL & vbCrLf & "  and     NRO_DESDE >= 101897 "
            
           Rem  Sql = Sql & vbCrLf & "  ORDER BY ID"
            Sql = Sql & vbCrLf & "  ORDER BY DOCUMENTOS_DIGITALES.FECHA_DESDE "
            Rem SQL = SQL & vbCrLf & "  ORDER BY FECHA_DESDE , DOCUMENTOS_DIGITALES.NRO_DESDE"
            
            
            
    Dim C As Long
    Dim anio As String
    Dim Mes  As String
    Dim NombreArchivo As String
    Dim BanderaPrimera As Boolean
    BanderaPrimera = True
    PasoInicial = InputBox("Ingrese el paso final de las imagenes", "Arliquide", "D:\Arliquide")
            
            i = 2
            rs.CursorLocation = adUseClient
            rs.Open Sql, strConBasa, adOpenDynamic, adLockReadOnly
            


     If Dir(PasoInicial, vbDirectory) = "" Then
                FileSystem.MkDir PasoInicial

                      
                 
                    End If
                Do While Not rs.EOF
                
                
                
                
                    
                    anio = Format(rs!FECHA_DESDE, "YYYY")
                    Mes = Format(rs!FECHA_DESDE, "MM")
                    
                    If AnioAnterior <> anio Then
                       If BanderaPrimera = False Then
                        C = 2
                            libroEx.SaveAs PasoInicial & "\" & AnioAnterior & ".xls"
                            libroEx.Close
                            ApExcel.Quit
                            Set ApExcel = Nothing
                            Set libroEx = Nothing
                       End If
                       BanderaPrimera = False
                       AnioAnterior = anio
                       
                        Rem abrir hoja excel
                        Set ApExcel = New Excel.Application
                        Set libroEx = Excel.Workbooks.Add
                        Set hojaEx = libroEx.Worksheets.Item(1)
                        hojaEx.Cells(1, 1) = "Imagen"
                        hojaEx.Cells(1, 2) = "Sucursal"
                        hojaEx.Cells(1, 3) = "Remito"
                        hojaEx.Cells(1, 4) = "Cliente"
                        hojaEx.Cells(1, 5) = "Fecha"
                        hojaEx.Cells(1, 6) = "Lote"
                     End If
        
                    
                    If Dir(PasoInicial & "\" & anio, vbDirectory) = "" Then
                        MkDir (PasoInicial & "\" & anio)
                        MkDir (PasoInicial & "\" & anio & "\" & Mes)
                    Else
                    
                        If Dir(PasoInicial & "\" & anio & "\" & Mes, vbDirectory) = "" Then
                            MkDir (PasoInicial & "\" & anio & "\" & Mes)
                        End If
                        
                        
                    End If
                    directorio = PasoInicial & "\" & anio & "\" & Mes
                    
                    
                    C = C + 1
    
                    NombreArchivo = Format(rs!NRO_HASTA, "0000") & "_" & Trim(Format(rs!NRO_DESDE, "00000000")) & "_" & rs!ID & ".tif"
                     hojaEx.Cells(C, 1) = rs!ID
                     hojaEx.Cells(C, 1).Hyperlinks.Add hojaEx.Cells(C, 1), ".\" & anio & "\" & Mes & "\" & NombreArchivo
    
                    If IsNull(rs!NRO_HASTA) Then
                        hojaEx.Cells(C, 2) = ""
                    Else
                        hojaEx.Cells(C, 2) = Trim(rs!NRO_HASTA)
                    End If
    
                    If IsNull(rs!NRO_DESDE) Then
                        hojaEx.Cells(C, 3) = ""
                    Else
                        hojaEx.Cells(C, 3) = Trim(rs!LETRA_DESDE)
                    End If
    
                    If IsNull(rs!LETRA_HASTA) Then
                        hojaEx.Cells(C, 4) = ""
                    Else
                        hojaEx.Cells(C, 4) = Trim(rs!LETRA_HASTA)
                    End If
    
    
                    If IsNull(rs!FECHA_DESDE) Then
                        hojaEx.Cells(C, 5) = ""
                    Else
                        hojaEx.Cells(C, 5) = CDate(rs!FECHA_DESDE)
                    End If
                     hojaEx.Cells(C, 6) = Trim(rs!lote)
    
    If Dir(PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif") <> "" Then
                       
                    FileCopy PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", directorio & "\" & NombreArchivo
                    Else
                    Debug.Print rs!ID
                    End If
                    
                    rs.MoveNext
                    lblCantidad.Caption = lblCantidad.Caption + 1
                    lblCantidad.Refresh
                Loop
    lblCantidad.Caption = "Terminado"
     Rem luis libroEx.SaveAs PasoInicial & "\" & AnioAnterior & ".xls"
                         Rem   libroEx.Close
                          Rem   ApExcel.Quit
                           Rem  Set ApExcel = Nothing
                            Rem Set libroEx = Nothing
               
  
End Sub

'Public Sub Export_AirLiquide()
'     Dim ApExcel As Excel.Application
'        Dim libroEx As Excel.Workbook
'        Dim hojaEx As Excel.Worksheet
'
'
'     Dim Min As Long
'        Dim Max As Long
'
'        Rem abrir hoja excel
'        Set ApExcel = New Excel.Application
'        Set libroEx = Excel.Workbooks.Add
'        Set hojaEx = libroEx.Worksheets.Item(1)
'        Dim rs As New ADODB.Recordset
'        Dim sql As String
'        Dim PasoInicial As String
'
'             PasoInicial = "D:\ExportarImagenes\"
'
'        Dim i As Integer
'        Dim R As Excel.Range
'        Dim H As Excel.Hyperlinks
'        i = 1
'
'
'
'                hojaEx.Cells(1, 1) = "Imagen"
'                hojaEx.Cells(1, 2) = "Sucursal"
'                hojaEx.Cells(1, 3) = "Remito"
'                hojaEx.Cells(1, 4) = "Cliente"
'                hojaEx.Cells(1, 5) = "Fecha"
'                hojaEx.Cells(1, 6) = "Lote"
'
'
'
'            Dim sgrabar As String
'
'
'
'
'            sql = " SELECT    LETRA_DESDE, ID, NRO_HASTA, NRO_DESDE, DIRECTORIO_PASO, NRO_CAJA, FECHA_DESDE, LETRA_DESDE AS Expr1, LETRA_HASTA, LOTE, "
'              sql = sql & vbCrLf & "  IMAGEN_ORIGEN, COD_CLIENTE"
'           sql = sql & vbCrLf & " FROM   DOCUMENTOS_DIGITALES"
'            sql = sql & vbCrLf & " where FK_DOCUMENTOS_DIGITALES_LOTE IN (" & InputBox("Ingrese los numeros de lotes separadospor ,") & ")"
'            sql = sql & vbCrLf & " ORDER BY NRO_HASTA"
'
'    Dim NombreArchivo As String
'
'            i = 2
'            rs.Open sql, strConBasa , 0 ,1
'
'            Dim Directorio As String
'Directorio = InputBox("Directorio de Exportacion")
'
'                Do While Not rs.EOF
'                    i = i + 1
'
'                    NombreArchivo = Format(rs!NRO_HASTA, "0000") & "_" & Trim(Format(rs!LETRA_DESDE, "00000000")) & "_" & rs!ID & ".tif"
'                     hojaEx.Cells(i, 1) = rs!ID
'                     hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\" & Directorio & "\" & NombreArchivo
'
'                    If IsNull(rs!NRO_HASTA) Then
'                        hojaEx.Cells(i, 2) = ""
'                    Else
'                        hojaEx.Cells(i, 2) = rs!NRO_HASTA
'                    End If
'
'                    If IsNull(rs!NRO_DESDE) Then
'                        hojaEx.Cells(i, 3) = ""
'                    Else
'                        hojaEx.Cells(i, 3) = rs!NRO_DESDE
'                    End If
'
'                    If IsNull(rs!LETRA_HASTA) Then
'                        hojaEx.Cells(i, 4) = ""
'                    Else
'                        hojaEx.Cells(i, 4) = Trim(rs!LETRA_HASTA)
'                    End If
'
'
'                    If IsNull(rs!FECHA_DESDE) Then
'                        hojaEx.Cells(i, 5) = ""
'                    Else
'                        hojaEx.Cells(i, 5) = rs!FECHA_DESDE
'                    End If
'                    hojaEx.Cells(i, 6) = Trim(rs!Lote) & "-" & Trim(rs!IMAGEN_ORIGEN)
'
'
'                    If Dir(PasoInicial & Directorio, vbDirectory) = "" Then
'                        FileSystem.MkDir PasoInicial & Directorio
'                    End If
'
'                    FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & Directorio & "\" & NombreArchivo
'
'
'
'                    rs.MoveNext
'                Loop
'
'               libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY") & "2.xls"
'               libroEx.Close
'               ApExcel.Quit
'               Set ApExcel = Nothing
'               Set libroEx = Nothing
'
'End Sub
'
Public Sub Export_COHEN()
     Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
    
    
     Dim Min As Long
        Dim Max As Long
    
        Rem abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Add
        Set hojaEx = libroEx.Worksheets.Item(1)
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim PasoInicial As String
    
             PasoInicial = "E:\ExportarImagenes\"
    
        Dim i As Integer
        Dim R As Excel.Range
        Dim h As Excel.Hyperlinks
        i = 1
    
    
    
                hojaEx.Cells(1, 1) = "Imagen"
                hojaEx.Cells(1, 2) = "Sucursal"
                hojaEx.Cells(1, 3) = "Remito"
                hojaEx.Cells(1, 4) = "Cliente"
                hojaEx.Cells(1, 5) = "Fecha"
                hojaEx.Cells(1, 6) = "Lote"
    
    
    
            Dim sgrabar As String
    
    
    
    
            Sql = " SELECT    LETRA_DESDE, ID, NRO_HASTA, NRO_DESDE, DIRECTORIO_PASO, NRO_CAJA, FECHA_DESDE, LETRA_DESDE AS Expr1, LETRA_HASTA, LOTE, "
              Sql = Sql & vbCrLf & "  IMAGEN_ORIGEN, COD_CLIENTE"
           Sql = Sql & vbCrLf & " FROM   DOCUMENTOS_DIGITALES"
            Sql = Sql & vbCrLf & " where FK_DOCUMENTOS_DIGITALES_LOTE IN (" & InputBox("Ingrese los numeros de lotes separadospor ,") & ")"
            Sql = Sql & vbCrLf & " ORDER BY NRO_HASTA"
    
    Dim NombreArchivo As String
    
            i = 2
            rs.Open Sql, ConActiva, 0, 1
            
            Dim directorio As String
directorio = InputBox("Directorio de Exportacion")
    
                Do While Not rs.EOF
                    i = i + 1
    
                    NombreArchivo = Format(rs!NRO_HASTA, "0000") & "_" & Trim(Format(rs!LETRA_DESDE, "00000000")) & "_" & rs!ID & ".tif"
                     hojaEx.Cells(i, 1) = rs!ID
                     hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\" & directorio & "\" & NombreArchivo
    
                    If IsNull(rs!NRO_HASTA) Then
                        hojaEx.Cells(i, 2) = ""
                    Else
                        hojaEx.Cells(i, 2) = rs!NRO_HASTA
                    End If
    
                    If IsNull(rs!NRO_DESDE) Then
                        hojaEx.Cells(i, 3) = ""
                    Else
                        hojaEx.Cells(i, 3) = rs!NRO_DESDE
                    End If
    
                    If IsNull(rs!LETRA_HASTA) Then
                        hojaEx.Cells(i, 4) = ""
                    Else
                        hojaEx.Cells(i, 4) = Trim(rs!LETRA_HASTA)
                    End If
    
    
                    If IsNull(rs!FECHA_DESDE) Then
                        hojaEx.Cells(i, 5) = ""
                    Else
                        hojaEx.Cells(i, 5) = rs!FECHA_DESDE
                    End If
                    hojaEx.Cells(i, 6) = Trim(rs!lote) & "-" & Trim(rs!IMAGEN_ORIGEN)
    
    
                    If Dir(PasoInicial & directorio, vbDirectory) = "" Then
                        FileSystem.MkDir PasoInicial & directorio
                    End If
    
                    FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & directorio & "\" & NombreArchivo
    
    
    
                    rs.MoveNext
                Loop
    
               libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY") & "2.xls"
               libroEx.Close
               ApExcel.Quit
               Set ApExcel = Nothing
               Set libroEx = Nothing
  
End Sub


'Public Function PegadoImagen(ID_ficha As Long) As MODI.Document
'
'
'Dim Sql As String
'    Dim rsImagenes As New ADODB.Recordset
'    Dim docOrigen As MODI.Document
'    Dim docDestino As MODI.Document
'
'    MousePointer = 11
'
'    If Dir(txtPasoImagenesFinal.Text & "\" & Trim(TXTnOMBREdiRECTORIO.Text), vbDirectory) = "" Then
'    FileSystem.MkDir txtPasoImagenesFinal.Text & "\" & Trim(TXTnOMBREdiRECTORIO.Text)
'    End If
'
'        Sql = "   SELECT     COD_CLIENTE, LOTE, IMAGEN_NOTTI,DIRECTORIO_PASO, ID"
'        Sql = Sql & "  From DOCUMENTOS_DIGITALES"
'        Sql = Sql & "  WHERE     (COD_CLIENTE = 163) AND (LOTE LIKE N'5100%') AND (NOT (IMAGEN_NOTTI IS NULL))"
'
'        Set rsBuscar = New ADODB.Recordset
'        rsBuscar.Open Sql, strConBasa , 0 ,1
'
'        Do While Not rsBuscar.EOF
'
'
'
'
'            Sql = "  SELECT ID, COD_CLIENTE, LOTE, COD_ESTADO, DIRECTORIO_PASO, LETRA_DESDE,NRO_DESDE "
'            Sql = Sql & " From  DOCUMENTOS_DIGITALES  "
'            Sql = Sql & "  WHERE    id=" & rsBuscar!IMAGEN_NOTTI
'
'
'            Set rsImagenes = New ADODB.Recordset
'            rsImagenes.Open Sql, strConBasa , 0 ,1
'
'
'            Set docOrigen = New MODI.Document
'            docOrigen.Create PasoImagenes & rsBuscar!DIRECTORIO_PASO & "\" & rsBuscar!ID & ".tif"
'            If Not rsImagenes.EOF Then
'                Set docDestino = New MODI.Document
'                docDestino.Create PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif"
'                docDestino.Images.Add docOrigen.Images.Item(0), docDestino.Images.Item(0)
'             End If
'
'
'                docDestino.SaveAs "D:\ExportarImagenes\" & Trim(rsImagenes!NRO_DESDE) & "_" & rsImagenes!ID & ".TIF"
'
'       rsBuscar.MoveNext
'Loop
'MousePointer = 0
'MsgBox "Operacion terminada"
'
'End Function
'
Public Sub UnirFlyersNotti()
  
    
'    Dim DocSeparador As MODI.Document
'    Dim DocFichas As MODI.Document
'    Dim DocFlyers As MODI.Document
'    Dim DocSave As MODI.Document
'    Dim Sql As String
'    Dim rsFlyers As ADODB.Recordset
'    Dim rsNotti As ADODB.Recordset
'    Dim lOTESFICHAS As String
'    Dim LotesNotti As String
'    Dim Carpeta As String
'    Dim SOLICITUD As Long
'    Carpeta = InputBox("Ingrese las paso de exportacion", "Paso", "D:\ExportarImagenes\")
'     Carpeta = Trim(Carpeta)
'    lOTESFICHAS = txtUnirFichas.Text
'    LotesNotti = txtUnirNotti.Text
'    MousePointer = 11
'
'   FileSystem.MkDir Carpeta
'
'    Set DocSeparador = New MODI.Document
'    DocSeparador.Create "C:\registro.tif"
'    Dim i As Integer
'
'         Sql = " SELECT     ID, FK_DOCUMENTOS_DIGITALES_LOTE, NRO_DESDE,NRO_HASTA,  DIRECTORIO_PASO"
'         Sql = Sql & "  From DOCUMENTOS_DIGITALES"
'         Sql = Sql & "  WHERE  FK_DOCUMENTOS_DIGITALES_LOTE IN (" & lOTESFICHAS & ")"
'         Sql = Sql & "  order by ID"
'
'            Set rsFlyers = New ADODB.Recordset
'            rsFlyers.Open Sql, strConBasa , 0 ,1
'            Dim j As Integer
'            Do While Not rsFlyers.EOF
'                    Set DocSeparador = New MODI.Document
'                    DocSeparador.Create "C:\registro.tif"
'                    Set DocFichas = New MODI.Document
'                    DocFichas.Create PasoImagenes & rsFlyers!DIRECTORIO_PASO & "\" & rsFlyers!ID & ".tif"
'                    Set DocSave = New MODI.Document
'                    DocSave.Create
'
'                    If IsNull(rsFlyers!NRO_HASTA) Then
'                    SOLICITUD = 1111111111
'                    Else
'                        If rsFlyers!NRO_HASTA < 10 Then
'                            SOLICITUD = 1111111111
'                        Else
'                            SOLICITUD = rsFlyers!NRO_HASTA
'                        End If
'                    End If
'
'                    Sql = " SELECT  ID, FK_DOCUMENTOS_DIGITALES_LOTE, NRO_DESDE, DIRECTORIO_PASO"
'                    Sql = Sql & "   From DOCUMENTOS_DIGITALES"
'                    Sql = Sql & "   WHERE  FK_DOCUMENTOS_DIGITALES_LOTE IN (" & LotesNotti & ")"
'                    Sql = Sql & "   AND ( NRO_HASTA  = " & SOLICITUD
'                    Sql = Sql & "   OR  (LETRA_HASTA = '" & SOLICITUD & "')) "
'                    Sql = Sql & "   ORDER BY ID"
'
'
'                    Set rsNotti = New ADODB.Recordset
'
'                    rsNotti.Open Sql, strConBasa , 0 ,1
'
'                         If rsNotti.EOF Then
'
'                         If rsFlyers!NRO_DESDE > 100 Then
'
'                         Sql = " SELECT  ID, FK_DOCUMENTOS_DIGITALES_LOTE, NRO_DESDE, DIRECTORIO_PASO"
'                    Sql = Sql & "   From DOCUMENTOS_DIGITALES"
'                    Sql = Sql & "   WHERE  FK_DOCUMENTOS_DIGITALES_LOTE IN (" & LotesNotti & ")"
'                    Sql = Sql & "   AND  NRO_DESDE  = " & rsFlyers!NRO_DESDE
'                    Sql = Sql & "   ORDER BY ID"
'
'
'
'
'
'                    Set rsNotti = New ADODB.Recordset
'
'                    rsNotti.Open Sql, strConBasa , 0 ,1
'                         End If
'                         End If
'
'
'
'
'                     If Not rsNotti.EOF Then
'                            Do While Not rsNotti.EOF
'                                Set DocFlyers = New MODI.Document
'                                DocFlyers.Create PasoImagenes & rsNotti!DIRECTORIO_PASO & "\" & rsNotti!ID & ".tif"
'                                For i = 0 To DocFlyers.Images.Count - 1
'                                     DocSave.Images.Add DocFlyers.Images.Item(i), DocFlyers.Images.Item(i)
'                                 Next
'                                rsNotti.MoveNext
'                            Loop
'
'
'                             If DocFichas.Images.Count = 1 Then
'                                DocSave.Images.Add DocFichas.Images.Item(0), DocFichas.Images.Item(0)
'                             Else
'                              For i = 0 To DocFichas.Images.Count - 1
'                               DocSave.Images.Add DocFichas.Images.Item(i), DocFichas.Images.Item(i)
'                              Next
'                             End If
'
'                             DocSave.Images.Add DocSeparador.Images.Item(0), DocSeparador.Images.Item(0)
'
'                             DocSave.SaveAs "c:\LUIS.TIF"
'                             If chkInvertirNombre.value = 1 Then
'                                DocSave.SaveAs Carpeta & "\" & rsFlyers!ID & "_" & Trim(rsFlyers!NRO_DESDE) & ".TIF"
'                             Else
'                                 DocSave.SaveAs Carpeta & "\" & Trim(rsFlyers!NRO_DESDE) & "_" & rsFlyers!ID & ".TIF"
'                             End If
'                             DocSave.Close
'                             DocFlyers.Close
'
'                    Else
'                             If DocFichas.Images.Count = 1 Then
'                             DocSave.Images.Add DocFichas.Images.Item(0), DocFichas.Images.Item(0)
'                             Else
'                              For i = 0 To DocFichas.Images.Count - 1
'                               DocSave.Images.Add DocFichas.Images.Item(i), DocFichas.Images.Item(i)
'                              Next
'                             End If
'
'                             DocSave.Images.Add DocSeparador.Images.Item(0), DocSeparador.Images.Item(0)
'                             If chkInvertirNombre.value = 1 Then
'                                DocSave.SaveAs Carpeta & "\" & rsFlyers!ID & "_" & Trim(rsFlyers!NRO_DESDE) & ".TIF"
'
'                             Else
'                                DocSave.SaveAs Carpeta & "\" & Trim(rsFlyers!NRO_DESDE) & "_" & rsFlyers!ID & ".TIF"
'                             End If
'                    End If
'        rsFlyers.MoveNext
'       Loop
'
'
'        MousePointer = 0
'        MsgBox "Operacion terminada"
End Sub

Private Sub txtDato_LostFocus(Index As Integer)

txtDato(Index).Text = Trim(txtDato(Index).Text)


End Sub

Private Sub txtUnirFichas_Change()
txtUnirFichas.Text = Replace(Replace(txtUnirFichas.Text, vbCrLf, ""), " ", "")
End Sub


Private Sub txtUnirNotti_Change()
txtUnirNotti.Text = Replace(Replace(txtUnirNotti.Text, vbCrLf, ""), " ", "")
End Sub



Public Sub ExportOsdeVitalicia()
Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet


 Dim Min As Long
    Dim Max As Long
       
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Add
    Set hojaEx = libroEx.Worksheets.Item(1)
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim PasoInicial As String
    
         PasoInicial = "E:\ExportarImagenes\"

    Dim i As Integer
    Dim R As Excel.Range
    Dim h As Excel.Hyperlinks
    i = 1

                   
            
            hojaEx.Cells(1, 1) = "Link Imagen"
            hojaEx.Cells(1, 2) = "Nombre Imagen"
            hojaEx.Cells(1, 3) = "Caja"
            hojaEx.Cells(1, 4) = "Legajo"
            
        
        
      

        Dim sgrabar As String

        




      


Sql = " SELECT     INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.NRO_HASTA,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO,   DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS AS NRO_CAJA,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.FECHA_DESDE, DOCUMENTOS_DIGITALES.LETRA_DESDE AS Expr1, DOCUMENTOS_DIGITALES.LETRA_HASTA,"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES.Lote , DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES.Nombre"
Sql = Sql & vbCrLf & "  FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & vbCrLf & " INDICES ON DOCUMENTOS_DIGITALES_LOTE.FK_INDICES = INDICES.ID"
Sql = Sql & vbCrLf & "  Where  FK_DOCUMENTOS_DIGITALES_LOTE in(" & InputBox("Ingrese los numeros de lote separados por ,") & ")"
Sql = Sql & vbCrLf & "  ORDER BY DOCUMENTOS_DIGITALES.NRO_CAJA"


Dim NombreArchivo As String
      
        i = 2
        rs.Open Sql, ConActiva, 0, 1

            Do While Not rs.EOF
                i = i + 1
                NombreArchivo = rs!NRO_DESDE & " " & Trim(rs!LETRA_DESDE) & "_" & CStr(rs!ID) & ".tif"
                hojaEx.Cells(i, 1) = NombreArchivo
                hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\20090717\" & NombreArchivo
                
                If IsNull(rs!LETRA_DESDE) Then
                    hojaEx.Cells(i, 2) = rs!NRO_DESDE
                Else
                    hojaEx.Cells(i, 2) = rs!LETRA_DESDE
                End If
                hojaEx.Cells(i, 2) = rs!ID
                hojaEx.Cells(i, 3) = rs!NRO_CAJA
                hojaEx.Cells(i, 4) = Trim(rs!LETRA_DESDE)
                

                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & "\20090717\" & NombreArchivo
                hojaEx.Cells(i, 4) = rs!NRO_CAJA
                hojaEx.Cells(i, 5) = rs!ID
                
                rs.MoveNext
            Loop

           libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY") & ".xls"
           libroEx.Close
           ApExcel.Quit
           Set ApExcel = Nothing
           Set libroEx = Nothing
  

End Sub

Public Sub CopiarMontemar()

 Dim Sql As String
    Dim rsImagenes As New ADODB.Recordset
   

       
        MousePointer = 11
        
        
        

        
Sql = "   SELECT     DOCUMENTOS_DIGITALES_LOTE.REMITO, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, "
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION, DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_DESDE, "
Sql = Sql & " DOCUMENTOS_DIGITALES.NRO_DESDE , DOCUMENTOS_DIGITALES.DIRECTORIO_PASO "

Sql = Sql & " FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN DOCUMENTOS_DIGITALES ON   DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
  Sql = Sql & " WHERE     (DOCUMENTOS_DIGITALES_LOTE.REMITO IN ('0001-00032866', '0001-00032877', '0001-00032865')) "

 Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.REMITO "

Dim REMITO As String
        
        
        Set rsImagenes = New ADODB.Recordset
        
        Dim Nombre As String
        rsImagenes.Open Sql, ConActiva, 0, 1
        Dim nuevoPaso As String
        
            Do While Not rsImagenes.EOF
             If REMITO <> rsImagenes!REMITO Then
             REMITO = rsImagenes!REMITO
              FileSystem.MkDir txtPasoImagenesFinal & "\" & REMITO
              nuevoPaso = txtPasoImagenesFinal & REMITO & "\"
             End If
            If Dir(PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif") <> "" Then
                Debug.Print rsImagenes!ID
                Nombre = ""
                If IsNull(rsImagenes!LETRA_DESDE) Then
                  Nombre = "No tiene"
                Else
                 Nombre = Trim(rsImagenes!LETRA_DESDE)
                End If
                
                  
               FileSystem.FileCopy PasoImagenes & rsImagenes!DIRECTORIO_PASO & "\" & rsImagenes!ID & ".tif", nuevoPaso & rsImagenes!NRO_DESDE & " " & Trim(rsImagenes!LETRA_DESDE) & "_" & rsImagenes!ID & ".tif"
              
            Else
            
                MsgBox "error"
                
            End If
              rsImagenes.MoveNext
            Loop
       Rem rsBuscar.MoveNext

MousePointer = 0
MsgBox "Operacion terminada"
End Sub






Public Function ControlExpreso(lote As Long) As String


Dim Sql As String
Dim ID_imagen As String


Dim rs As New ADODB.Recordset
Dim rsImagen As New ADODB.Recordset

Sql = " SELECT     LETRA_DESDE, COUNT(*) AS Expr1"
Sql = Sql & " From basasql.dbo.DOCUMENTOS_DIGITALES"
Sql = Sql & " Where FK_DOCUMENTOS_DIGITALES_LOTE = " & lote
Sql = Sql & " GROUP BY LETRA_DESDE"
Sql = Sql & " HAVING      (COUNT(*) < 4)"


rs.Open Sql, strConBasa

 Do While Not rs.EOF
 
 Sql = "  SELECT     ID, LETRA_DESDE From DOCUMENTOS_DIGITALES"
Sql = Sql & " WHERE    FK_DOCUMENTOS_DIGITALES_LOTE = " & lote
Sql = Sql & "  AND (LETRA_DESDE = '" & rs!LETRA_DESDE & "')"


Set rsImagen = New ADODB.Recordset


    rsImagen.Open Sql, strConBasa
    Do While Not rsImagen.EOF
        
    ID_imagen = ID_imagen & "," & rsImagen!ID
        
    
        rsImagen.MoveNext
    
    Loop
    
 
 
 
 
    rs.MoveNext
 
 Loop

ControlExpreso = Mid(ID_imagen, 2)

End Function

Public Function ControlExpresoGuia(lote As Long) As String


Dim Sql As String
Dim ID_imagen As String


Dim rs As New ADODB.Recordset
Dim rsImagen As New ADODB.Recordset


 
 Sql = "  SELECT     ID, LETRA_DESDE From DOCUMENTOS_DIGITALES"
Sql = Sql & " WHERE    FK_DOCUMENTOS_DIGITALES_LOTE = " & lote
Sql = Sql & "  AND (LETRA_DESDE = '" & rs!LETRA_DESDE & "')"
Sql = Sql & " AND ((NRO_DESDE = 0) OR (NRO_HASTA = 0)) "


Set rsImagen = New ADODB.Recordset


    rsImagen.Open Sql, strConBasa
    Do While Not rsImagen.EOF
        
    ID_imagen = ID_imagen & "," & rsImagen!ID
        
    
        rsImagen.MoveNext
    
    Loop
    
 

Rem ControlExpreso = Mid(ID_imagen, 2)

End Function


Public Sub ExpresoLujanExportarCodigo()
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Dim PasoInicial As String
        Dim i As Integer
        Dim sgrabar As String
        Dim DATO  As String
        Dim NombreArchivo As String
        Dim Paso As String
        Dim Año As String
        Dim Mes As String
        Dim Dia As String
        Dim FechaActulizacion As String
        Dim Caja As String
        Dim Strlotes As String
        
Caja = InputBox("Ingrese la caja")


Sql = " SELECT DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES INNER JOIN DOCUMENTOS_DIGITALES_LOTE ON"
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401 "
Sql = Sql & " AND DOCUMENTOS_DIGITALES.ESTADO = 'LISTA PARA EXPORTAR'"
Sql = Sql & " AND DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS = " & Caja
Sql = Sql & " GROUP BY DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"

Set rs = New ADODB.Recordset
            rs.Open Sql, strConBasa
   Do While Not rs.EOF
        Strlotes = Strlotes & "," & rs!ID_DOCUMENTOS_DIGITALES_LOTE
    rs.MoveNext
   Loop
   
            
            

Sql = " SELECT        DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO , DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
Sql = Sql & " DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE,"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES , DOCUMENTOS_DIGITALES.NRO_DESDE as LETRA_HASTA, DOCUMENTOS_DIGITALES.estado"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES INNER JOIN DOCUMENTOS_DIGITALES_LOTE ON"
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401"
Sql = Sql & "  AND ESTADO  ='LISTA PARA EXPORTAR' "
Sql = Sql & "  AND  FK_CAJAS =" & Caja
Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES.ESTADO DESC"

      Dim cant As Integer
            i = 2
            Set rs = New ADODB.Recordset
            rs.Open Sql, ConActiva, 0, 1
            Paso = "C:/ExpresoLujan"
            Set rs = New ADODB.Recordset
            rs.Open Sql, strConBasa
            FechaActulizacion = InputBox("Ingrese La fecha de Exportacion", "Fecha de exportado", Format(Now, "yyyymmdd") & "01")
    Do While Not rs.EOF
        DATO = Trim(rs!LETRA_HASTA)
        If Len(DATO) = 18 Then
                Sql = "  Update basasql.dbo.DOCUMENTOS_DIGITALES"
                Sql = Sql & " Set Exportado ='" & FechaActulizacion & "'"
                 Sql = Sql & " , ESTADO ='EXPORTADA'"
                Sql = Sql & " , descripcion = 'Exportado : " & FechaActulizacion & "'"
                Sql = Sql & " Where ID = " & rs!ID
                ExecutarSql Sql
                Año = "20" & CStr(Mid(DATO, 1, 2))
                Mes = Mid(DATO, 3, 2)
                Dia = Mid(DATO, 5, 2)
                If Dir(Paso & "/" & Año, vbDirectory) = "" Then
                    FileSystem.MkDir (Paso & "/" & Año)
                End If
                If Dir(Paso & "/" & Año & "/" & Mes, vbDirectory) = "" Then
                    FileSystem.MkDir (Paso & "/" & Año & "/" & Mes)
                End If
                If Dir(Paso & "/" & Año & "/" & Mes & "/" & Dia, vbDirectory) = "" Then
                    FileSystem.MkDir (Paso & "/" & Año & "/" & Mes & "/" & Dia)
                End If
                If Dir(PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif") <> "" Then
                 FileCopy PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif", Paso & "/" & Año & "/" & Mes & "/" & Dia & "/" & DATO & ".tif"
               Else
                MsgBox "NO SE ENCONTRO EL ARCHIVO" & rs!ID
                
               End If
               cant = cant + 1
        
        Debug.Print rs!ID
        Else
        MsgBox "El largo de campo no es el correcto", "verificque en indice"
        End If
        rs.MoveNext
    Loop
    
    
    Sql = " Update DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " SET LOTE_ESTADO ='EXPORTADO'"
    Sql = Sql & " , FECHA_EXPORTACION=" & SysDateMinutoSegundo
    Sql = Sql & "  Where ID_DOCUMENTOS_DIGITALES_LOTE IN( " & Mid(Strlotes, 2) & ")"
        
    ExecutarSql Sql
    
         MsgBox cant
           MsgBox "Terminado"
  

End Sub



Public Function DigitoVerificadorExpreso(DATO As String) As Integer

    Dim Valor(17) As Integer
    Dim Datos(17) As Integer
    Dim Total As Integer
    Dim strTotal As String
    Dim i As Integer
    Dim DIGITO As Integer
    
DATO = Trim(DATO)
     For i = 1 To 17
     
           Datos(i) = CInt(Mid(DATO, i, 1))
          Valor(i) = Datos(i) * 3
           i = i + 1
           
           If i > 17 Then
                Exit For
           End If
           
           Datos(i) = CInt(Mid(DATO, i, 1))
          Valor(i) = Datos(i) * 1
    Next
  

        For i = 1 To 17
            Total = Total + Valor(i)
        Next
        
        strTotal = Total
        
       Rem  MsgBox Len(strTotal)
        
      If strTotal = "100" Then
        
           
            DIGITO = 0
            
        Exit Function
        End If
        
        If Len(strTotal) = 2 Then
            Total = Int(Str(CInt(Mid(strTotal, 1, 1)) + 1) & "0")
            DIGITO = CInt(strTotal) - Total
            DIGITO = DIGITO * -1
            If DIGITO = 10 Then
                DIGITO = 0
            End If
            
            DigitoVerificadorExpreso = DIGITO
        End If
        
        If Len(strTotal) = 3 Then
        
            Total = Mid(strTotal, 1, 1) & Int(Str(CInt(Mid(strTotal, 2, 1)) + 1) & "0")
            
            DIGITO = CInt(strTotal) - Total
            
             DIGITO = DIGITO * -1
            If DIGITO = 10 Then
                DIGITO = 0
            End If
            
            DigitoVerificadorExpreso = DIGITO
        End If
        
       
   
    
   

End Function

Public Sub ExpresoLujanCodigo()

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim Documento As Long
Dim Nombre As String
Dim P As Integer



Sql = "SELECT DOCUMENTOS_DIGITALES.ID"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES ON"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.DESCRIPCION IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.LETRA_DESDE IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.FECHA_DESDE IS NULL)"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID "

Dim Sucursal As Integer
Dim guia As Long
Dim Destino As String



rs.Open Sql, strConBasa
Do While Not rs.EOF


Sql = " SELECT     ID, ENBASA,NRO_DESDE, FECHA_DESDE ,  LETRA_DESDE , LETRA_hasta ,BatchPgDta"
Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
Sql = Sql & " ORDER BY ID"

Set rs2 = New ADODB.Recordset

rs2.Open Sql, strConBasa

Dim i As Integer
Dim Dato_Letra_Desde As String
    If Not rs2.EOF Then
            If DigitoVerificadorExpreso(Mid(CStr(rs2!NRO_DESDE), 1, 17)) = Mid(CStr(rs2!NRO_DESDE), 1, 17) Then
                guia = rs2!NRO_DESDE
            Else
                guia = 0
            End If
     Else
        guia = 0
    End If
    
        Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
        Sql = Sql & " SET "
        Sql = Sql & " NRO_DESDE =" & guia
        Sql = Sql & " Where ID = " & rs!ID
        ExecutarSql Sql
        rs.MoveNext
    Loop
       MsgBox "Terminado"
    End Sub

Public Sub ActualizarExpreso()
    Dim Sql As String
    Dim rs As ADODB.Recordset




Sql = " SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_HASTA, LEN(DOCUMENTOS_DIGITALES.LETRA_HASTA) AS largo,"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES , DOCUMENTOS_DIGITALES.NRO_DESDE, ExpresoLujan.GSCAEAN"
Sql = Sql & " FROM DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE ON"
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & " EXPRESOLUJAN ON DOCUMENTOS_DIGITALES.NRO_DESDE = EXPRESOLUJAN.GSCANUMGUI"
Sql = Sql & " WHERE (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401) AND (LEN(DOCUMENTOS_DIGITALES.LETRA_HASTA) <> 18) AND"
Sql = Sql & " (DOCUMENTOS_DIGITALES.NRO_DESDE > 1)"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE"







Set rs = New ADODB.Recordset
rs.Open Sql, strConBasa

 Do While Not rs.EOF
    Sql = " Update DOCUMENTOS_DIGITALES"
    Sql = Sql & "  SET LETRA_HASTA ='" & Trim(rs!GSCAEAN) & "'"
    Sql = Sql & " Where ID = " & rs!ID
    ExecutarSql Sql
    rs.MoveNext
 Loop
 






End Sub

Public Sub ExpresoLujanSacarImagenesConError()
Dim Sql As String
Dim rs As New ADODB.Recordset

Sql = "SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_HASTA, LEN(DOCUMENTOS_DIGITALES.LETRA_HASTA) AS largo,"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES , DOCUMENTOS_DIGITALES.NRO_DESDE"
Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE ON"
Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401) And (Len(DOCUMENTOS_DIGITALES.LETRA_HASTA) <> 18)"
Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES.NRO_DESDE "



Sql = " SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_HASTA, LEN(DOCUMENTOS_DIGITALES.LETRA_HASTA) AS Expr1,"
Sql = Sql & "                      DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.NRO_HASTA,"
Sql = Sql & "                      DOCUMENTOS_DIGITALES.FECHA_DESDE , DOCUMENTOS_DIGITALES.FECHA_HASTA"
Sql = Sql & " FROM         DOCUMENTOS_DIGITALES INNER JOIN"
 Sql = Sql & "                     DOCUMENTOS_DIGITALES_LOTE ON"
  Sql = Sql & "                    DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401) And (DOCUMENTOS_DIGITALES.Exportado Is Null) And (DOCUMENTOS_DIGITALES.NRO_DESDE < 1 Or DOCUMENTOS_DIGITALES.NRO_DESDE Is Null)"
Sql = Sql & " AND (DOCUMENTOS_DIGITALES.ID IN (1484571, 1485632, 1486068, 1486069, 1486146, 1486496, 1487319,"
Sql = Sql & "                      1488016, 1488141, 1488257, 1488267, 1488280, 1488282, 1488378, 1488516, 1488562, 1489361, 1490217, 1490548, 1493524, 1500794, 1500821, 1501308,"
Sql = Sql & "                      1501342, 1501611, 1502072, 1505318, 1506872, 1509074, 1511568, 1511652, 1513743, 1513796, 1513801, 1513814, 1513815, 1513853, 1513909, 1514004,"
 Sql = Sql & "                     1514019, 1514163, 1514180, 1514191, 1514192, 1514196, 1514205))"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID"



Sql = "  SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_HASTA, LEN(DOCUMENTOS_DIGITALES.LETRA_HASTA) AS Expr1,"
Sql = Sql & "                       DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.NRO_HASTA,"
 Sql = Sql & "                      DOCUMENTOS_DIGITALES.FECHA_DESDE , DOCUMENTOS_DIGITALES.FECHA_HASTA"
Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & "                       DOCUMENTOS_DIGITALES_LOTE ON"
 Sql = Sql & "                      DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  WHERE     (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401) AND (DOCUMENTOS_DIGITALES.Exportado IS NULL)  AND (DOCUMENTOS_DIGITALES.ID IN (1484571, 1485632, 1486068, 1486069, 1486146, 1486496, 1487319,"
Sql = Sql & "                       1488016, 1488141, 1488257, 1488267, 1488280, 1488282, 1488378, 1488516, 1488562, 1489361, 1490217, 1490548, 1493524, 1500794, 1500821, 1501308,"
Sql = Sql & "                       1501342, 1501611, 1502072, 1505318, 1506872, 1509074, 1511568, 1511652, 1513743, 1513796, 1513801, 1513814, 1513815, 1513853, 1513909, 1514004,"
 Sql = Sql & "                      1514019, 1514163, 1514180, 1514191, 1514192, 1514196, 1514205))"
Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES.ID"




Sql = "  SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES.LETRA_HASTA, LEN(DOCUMENTOS_DIGITALES.LETRA_HASTA) AS Expr1,"
Sql = Sql & " DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.NRO_DESDE, DOCUMENTOS_DIGITALES.NRO_HASTA,"
Sql = Sql & " DOCUMENTOS_DIGITALES.FECHA_DESDE , DOCUMENTOS_DIGITALES.FECHA_HASTA"
Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & "                       DOCUMENTOS_DIGITALES_LOTE ON"
 Sql = Sql & "                      DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & "  WHERE     (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401)  AND( (LEN(DOCUMENTOS_DIGITALES.LETRA_HASTA)"
Sql = Sql & "                       <> 18) OR (DOCUMENTOS_DIGITALES.LETRA_HASTA) IS NULL)"
Sql = Sql & "  ORDER BY DOCUMENTOS_DIGITALES.ID"




rs.Open Sql, strConBasa

Do While Not rs.EOF

Sql = " Delete "
Sql = Sql & " From TELEFORM_DIGITAL"
Sql = Sql & " WHERE     (BatchPgDta LIKE  '" & rs!ID & "%')"
ExecutarSql Sql

FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", "D:\ExpresoLujan\" & rs!ID & ".tif"

rs.MoveNext

Loop


End Sub

Public Sub ExpresoLujanImportarCodigo()
        Dim Sql As String
        Dim rs As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
        Dim Documento As Long
        Dim Nombre As String
        Dim P As Integer
        Dim i As Integer
        Dim Codigo As String
        
       Sql = " SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER,"
       Sql = Sql & "               DOCUMENTOS_DIGITALES.Exportado , DOCUMENTOS_DIGITALES.LETRA_HASTA"
Sql = Sql & " FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & "                       DOCUMENTOS_DIGITALES ON"
 Sql = Sql & "                     DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"

Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401) And   ( LEN(DOCUMENTOS_DIGITALES.LETRA_HASTA) < 2 or DOCUMENTOS_DIGITALES.LETRA_HASTA is null) and  (DOCUMENTOS_DIGITALES.Exportado Is Null)"
      Sql = Sql & " and DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in (" & InputBox("ingrese la orden") & ")"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER DESC"
        
        
'            Sql = "SELECT DOCUMENTOS_DIGITALES.ID"
'            Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
'            Sql = Sql & " DOCUMENTOS_DIGITALES ON"
'            Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
'            Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401"
'    Rem         Sql = Sql & " and DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in ( 22470)"
'            Sql = Sql & " AND (DOCUMENTOS_DIGITALES.LETRA_HASTA IS NULL)"
'            Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID "
            
            
            
            rs.Open Sql, strConBasa
            
            Do While Not rs.EOF
                Sql = " SELECT     ID, ENBASA, NRO_DESDE , NRO_HASTA ,  BatchPgDta"
                Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
                Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
                Sql = Sql & " ORDER BY ID DESC "
                Set rs2 = New ADODB.Recordset
                rs2.Open Sql, strConBasa
                Codigo = 0
                If Not rs2.EOF Then
                    
                        
                
                  If Not IsNull(rs2!NRO_DESDE) Then
                    If Len(Replace(Trim(rs2!NRO_DESDE), " ", "")) = 18 Then
                        Codigo = Trim(rs2!NRO_DESDE)
                    End If
                  End If
                
                If Codigo = 0 Then
                
                  If Not IsNull(rs2!NRO_HASTA) Then
                    If Len(Replace(Trim(rs2!NRO_HASTA), " ", "")) = 18 Then
                        Codigo = Trim(rs2!NRO_HASTA)
                    End If
                  End If
                
                End If
                
                  If Codigo > 0 Then
                    If DigitoVerificadorExpreso(CStr(Codigo)) = Mid(Trim(CStr(Codigo)), 18, 1) Then
                                    Codigo = Trim(Codigo)
                                    Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
                                    Sql = Sql & " SET "
                                    Sql = Sql & "  LETRA_HASTA ='" & Codigo & "'"
                                    Sql = Sql & " Where ID = " & rs!ID
                                    ExecutarSql Sql
                    End If
                End If
                 
                 End If
                
                rs.MoveNext
              Loop
            MsgBox "Terminado"
End Sub

Public Sub ExpresoLujanControlCodigo()

        Dim Sql As String
        Dim rs As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
        Dim Documento As Long
        Dim Nombre As String
        Dim P As Integer
        Dim i As Integer
        Dim Codigo As String
            
            
           Sql = " SELECT DOCUMENTOS_DIGITALES.ID , DOCUMENTOS_DIGITALES.LETRA_HASTA"
           Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES_LOTE"
           Sql = Sql & vbCrLf & " INNER JOIN DOCUMENTOS_DIGITALES "
           Sql = Sql & vbCrLf & " ON DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
           Sql = Sql & vbCrLf & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401"
           Sql = Sql & vbCrLf & " AND (LEN(DOCUMENTOS_DIGITALES.LETRA_HASTA) = 18)"
           Sql = Sql & vbCrLf & " AND DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE > " & InputBox("Ingrese el nenor numero de lote")
           Sql = Sql & vbCrLf & " ORDER BY DOCUMENTOS_DIGITALES.ID"

            
            
            rs.Open Sql, strConBasa
            Do While Not rs.EOF
                If DigitoVerificadorExpreso(rs!LETRA_HASTA) <> Mid(Trim(rs!LETRA_HASTA), 18, 1) Then
                        Sql = "  Update basasql.dbo.DOCUMENTOS_DIGITALES"
                        Sql = Sql & " SET  Exportado = 'NO'"
                        Sql = Sql & ", LETRA_HASTA='9999' "
                        Sql = Sql & " Where ID = " & rs!ID
                        ExecutarSql Sql
                End If
                
                
                rs.MoveNext
            Loop
            MsgBox "Terminado"
End Sub


Public Sub ExpresoLujanImportaGuia()

'Dim Sql As String
'Dim Rs As New ADODB.Recordset
'Dim rs2 As New ADODB.Recordset
'Dim Documento As Long
'Dim Nombre As String
'Dim P As Integer
'
'
'
'Sql = "SELECT DOCUMENTOS_DIGITALES.ID"
'Sql = Sql & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
'Sql = Sql & " DOCUMENTOS_DIGITALES ON"
'Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
'Sql = Sql & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401"
' Sql = Sql & " and DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE in ( 22470)"
'
''Sql = Sql & " AND (DOCUMENTOS_DIGITALES.DESCRIPCION IS NULL)"
''Sql = Sql & " AND (DOCUMENTOS_DIGITALES.LETRA_DESDE IS NULL)"
''Sql = Sql & " AND (DOCUMENTOS_DIGITALES.NRO_DESDE IS NULL)"
''Sql = Sql & " AND (DOCUMENTOS_DIGITALES.FECHA_DESDE IS NULL)"
'Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES.ID "
'
'Dim Sucursal As Integer
'Dim guia As Long
'Dim Destino As String
'
'
'
'Rs.Open Sql, strConBasa
'Do While Not Rs.EOF
'
'
'Sql = " SELECT     ID, ENBASA,NRO_DESDE, FECHA_DESDE ,  LETRA_DESDE , LETRA_hasta ,BatchPgDta"
'Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
'Sql = Sql & " WHERE     (BatchPgDta = '" & Rs!ID & ".TIF[ 1 ]')"
'Sql = Sql & " ORDER BY ID"
'
'Set rs2 = New ADODB.Recordset
'
'rs2.Open Sql, strConBasa
'
'Dim i As Integer
'Dim Dato_Letra_Desde As String
'Dim Codigo As String
'If Not rs2.EOF Then
'
'Sucursal = 0
'   Destino = "NULL"
'   guia = 0
'
'If Len(Trim(rs2!LETRA_HASTA)) = 2 Then
'    Destino = "'" & Replace(UCase(Trim(rs2!LETRA_HASTA)), "I", "J") & "'"
'
'Else
'    Destino = "NULL"
'
'End If
'
'If IsNull(rs2!LETRA_DESDE) Then
'Dato_Letra_Desde = 0
'Else
'
' Dato_Letra_Desde = Replace(rs2!LETRA_DESDE, " ", "0")
' End If
'
'If Len(Trim(Dato_Letra_Desde)) = 13 Or Len(Trim(Dato_Letra_Desde)) = 12 Then
'
'   If IsNumeric(Mid(Dato_Letra_Desde, 1, 3)) Then
'       If IsNumeric(Trim(Mid(Dato_Letra_Desde, 4, 1))) Then
'        Sucursal = Trim(Mid(Dato_Letra_Desde, 4, 1))
'        Else
'            Sucursal = 0
'        End If
'
'    Else
'        Sucursal = 0
'    End If
'
'    If IsNumeric(Mid(Dato_Letra_Desde, 6)) Then
'        guia = Mid(Dato_Letra_Desde, 6)
'    Else
'        guia = 0
'    End If
'
'Else
'
'
'    guia = 0
'     Sucursal = 0
'End If
'
'
'   If guia < 0 Then
'   guia = guia * -1
'   End If
'
'
'
'
'
'   If guia = 0 Then
'   guia = guia
'   End If
'Codigo = "'0'"
'     If Len(CStr(rs2!NRO_DESDE)) = 18 Then
'
''    If DigitoVerificadorExpreso(CStr(rs2!NRO_DESDE)) = Mid(CStr(rs!NRO_DESDE), 1, 17) Then
'
'    Codigo = "'" & CStr(rs2!NRO_DESDE) & "'"
'
''    End If
'End If
'
'
'
'
'
'
'     Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
'Sql = Sql & " SET "
''Sql = Sql & " NRO_DESDE =" & guia
''Sql = Sql & "  ,NRO_HASTA =" & Sucursal
''Sql = Sql & " , LETRA_DESDE =" & Destino
'Sql = Sql & "  LETRA_HASTA =" & Codigo
'Sql = Sql & " Where ID = " & Rs!ID
'
'ExecutarSql Sql
'
' End If
'
'
'
' Rs.MoveNext
'
'Loop


MsgBox "Terminado"


End Sub

Public Sub ExpresoLujanSacarImagenID()

    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim PasoInicial As String
    Dim NombreArchivo As String

        PasoInicial = txtPasoImagenesFinal.Text
        Sql = " SELECT DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION,"
        Sql = Sql & " DOCUMENTOS_DIGITALES.LETRA_HASTA , DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN, DOCUMENTOS_DIGITALES.Exportado, DOCUMENTOS_DIGITALES.ID"
        Sql = Sql & " FROM DOCUMENTOS_DIGITALES INNER JOIN"
        Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE ON"
        Sql = Sql & " DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401) "
        Sql = Sql & " And (DOCUMENTOS_DIGITALES.Exportado Is Null)"
        
        Sql = Sql & " AND (DOCUMENTOS_DIGITALES.LETRA_HASTA = '2')"
        
        Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS,"
        Sql = Sql & " DOCUMENTOS_DIGITALES_LOTE.DESCRIPCION,"
        Sql = Sql & " DOCUMENTOS_DIGITALES.IMAGEN_ORIGEN"
        rs.Open Sql, strConBasa
            Do While Not rs.EOF
                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & rs!ID & ".tif"
                rs.MoveNext
            Loop

           MsgBox "Terminado"

End Sub

Public Sub EXPRESOCONTROLCOMPLETO()
Dim Sql As String
        Dim rs As New ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
        Dim Documento As Long
        Dim Nombre As String
        Dim P As Integer
        Dim i As Integer
        Dim Codigo As String
        
Sql = " SELECT     DOCUMENTOS_DIGITALES.ID, DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER,"
Sql = Sql & "               DOCUMENTOS_DIGITALES.Exportado , DOCUMENTOS_DIGITALES.LETRA_HASTA"
Sql = Sql & " FROM         DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
Sql = Sql & "                       DOCUMENTOS_DIGITALES ON"
Sql = Sql & "                     DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
Sql = Sql & " Where (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 401) And (DOCUMENTOS_DIGITALES.Exportado Is Null)"
Sql = Sql & " ORDER BY DOCUMENTOS_DIGITALES_LOTE.FECHA_SCANNER DESC"
        
 
            Dim RS_GSCAEAN As New ADODB.Recordset
            Dim NRO_GUIA As Long
            
            rs.Open Sql, strConBasa
            Do While Not rs.EOF
                Sql = " SELECT     ID, ENBASA, NRO_DESDE ,DESCRIPCION , LETRA_DESDE , BatchPgDta"
                Sql = Sql & " From basasql.dbo.TELEFORM_DIGITAL"
                Sql = Sql & " WHERE     (BatchPgDta = '" & rs!ID & ".TIF[ 1 ]')"
                Sql = Sql & " ORDER BY ID DESC "
                Set rs2 = New ADODB.Recordset
                rs2.Open Sql, strConBasa


 If rs!ID = 1516757 Then
 MsgBox "ssss"
 End If
 
                If Not rs2.EOF Then
                 If Len(rs2!Descripcion) > 1 Then
                    If IsNumeric(Mid(rs2!Descripcion, 6)) Then
                        NRO_GUIA = Mid(rs2!Descripcion, 6)
                        Sql = "SELECT      GSCAEAN"
                        Sql = Sql & " From basasql.dbo.EXPRESOLUJANCOMPLETA"
                        Sql = Sql & " WHERE     GSCANUMGUI = " & NRO_GUIA
                
                        Set RS_GSCAEAN = New ADODB.Recordset
                        RS_GSCAEAN.Open Sql, strConBasa
                
                            If Not RS_GSCAEAN.EOF Then
                                Sql = " Update basasql.dbo.DOCUMENTOS_DIGITALES"
                                Sql = Sql & " SET "
                                Sql = Sql & "  LETRA_HASTA =" & Trim(RS_GSCAEAN!GSCAEAN)
                                Sql = Sql & " Where ID = " & rs!ID
                                ExecutarSql Sql
                            End If
                     End If
                 End If
              End If
                rs.MoveNext
            Loop
            MsgBox "Terminado"
End Sub

Public Sub ExportarCentroCard()


 Dim Min As Long
    Dim Max As Long
       
  
    Dim rs As New ADODB.Recordset
    Dim RSlOTES As New ADODB.Recordset
    Dim Strlotes As String
    Dim Sql As String
    Dim PasoInicial As String
    Dim fecha As String
    Dim NombreArchivo As String
    Dim Doc As String
    PasoInicial = txtPasoImagenesFinal.Text

    Dim i As Long
    Dim R As Excel.Range
    Dim h As Excel.Hyperlinks
    Dim CajasExportar As String
    i = 1
        Dim sgrabar As String
       CajasExportar = InputBox("Ingrese la/s caja/s a exportar")
        Sql = " SELECT   DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE "
        Sql = Sql & vbCrLf & " FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & vbCrLf & " DOCUMENTOS_DIGITALES ON DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & vbCrLf & " WHERE DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 279 "
         Sql = Sql & vbCrLf & " AND DOCUMENTOS_DIGITALES.ESTADO = N'LISTA PARA EXPORTAR'"
        Sql = Sql & vbCrLf & " AND FK_CAJAS IN(" & CajasExportar & ")"
        Sql = Sql & vbCrLf & " GROUP BY DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE"
        RSlOTES.Open Sql, strConBasa
        
         Do While Not RSlOTES.EOF
            Strlotes = Strlotes & "," & RSlOTES!ID_DOCUMENTOS_DIGITALES_LOTE
            RSlOTES.MoveNext
         Loop
         
         Dim CajaVieja As Long

            
     
        Sql = " SELECT DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE, DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES, DOCUMENTOS_DIGITALES.FK_LEGAJO_ETIQUETA, DOCUMENTOS_DIGITALES.ESTADO,"
        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES.ID AS IDIMAGEN, DOCUMENTOS_DIGITALES.DIRECTORIO_PASO, DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS, DOCUMENTOS_DIGITALES.NRO_DESDE,"
        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES.LETRA_DESDE"
        Sql = Sql & vbCrLf & "  FROM DOCUMENTOS_DIGITALES_LOTE INNER JOIN"
        Sql = Sql & vbCrLf & "  DOCUMENTOS_DIGITALES ON DOCUMENTOS_DIGITALES_LOTE.ID_DOCUMENTOS_DIGITALES_LOTE = DOCUMENTOS_DIGITALES.FK_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & vbCrLf & " WHERE        (DOCUMENTOS_DIGITALES_LOTE.FK_CLIENTES = 279)"
         Sql = Sql & vbCrLf & " AND (DOCUMENTOS_DIGITALES.ESTADO = N'LISTA PARA EXPORTAR')"
        Sql = Sql & vbCrLf & " AND DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS IN (" & CajasExportar & ") "
        Sql = Sql & vbCrLf & "  ORDER BY DOCUMENTOS_DIGITALES_LOTE.FK_CAJAS"
            
            
            Dim PasoCaja As String
        
         rs.CursorLocation = adUseClient
            rs.Open Sql, strConBasa, , adOpenDynamic, adLockReadOnly
        
            Do While Not rs.EOF
                i = i + 1
                
                If CajaVieja = rs!FK_CAJAS Then
                
                Else
                CajaVieja = rs!FK_CAJAS
                     FileSystem.MkDir (txtPasoImagenesFinal.Text & CajaVieja)
                    
                    
                    PasoCaja = txtPasoImagenesFinal.Text & CajaVieja & "\"
                End If
                If rs!IDIMAGEN = 5628171 Then
                 MsgBox "aaa"
                End If
                
                
                fecha = FileSystem.FileDateTime(PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".tif")
                fecha = Format(fecha, "DD_MM_YYYY")
                Doc = Format(Trim(rs!NRO_DESDE), "0000000000")
                NombreArchivo = Doc & " _ " & fecha & "   " & CStr(rs!IDIMAGEN) & ".tif"
                If Dir(PasoCaja & Doc, vbDirectory) = "" Then
                    FileSystem.MkDir PasoCaja & Doc
                    FileCopy PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".tif", PasoCaja & Doc & "\" & NombreArchivo
                Else
                    FileCopy PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!IDIMAGEN & ".tif", PasoCaja & Doc & "\" & NombreArchivo
                End If
                 
                    Sql = " Update DOCUMENTOS_DIGITALES SET ESTADO ='EXPORTADO'"
                    Sql = Sql & " Where ID = " & rs!IDIMAGEN
                    ExecutarSql Sql
                 rs.MoveNext
            Loop
   Sql = " Update DOCUMENTOS_DIGITALES_LOTE"
    Sql = Sql & " SET LOTE_ESTADO ='EXPORTADO'"
    Sql = Sql & "  Where ID_DOCUMENTOS_DIGITALES_LOTE IN( " & Mid(Strlotes, 2) & ")"
    
    ExecutarSql Sql
 
           MsgBox "Terminado"
  
End Sub




Public Sub PonerImagenLocal()
Dim Paso As String
    On Error GoTo salir
   
   
   
   
            If optImagenLocal.value = 0 Then
               Paso = PasoImagenes & "\" & rsGrilla!DIRECTORIO_PASO & "\" & rsGrilla!ID & ".TIF"
            Else
               Paso = "C:\ImagenLocal\" & rsGrilla!lote & "\" & rsGrilla!ID & ".TIF"
            End If
            
            If Dir(Paso) <> "" Then
                ctlVerImagenes1.PonerImagen Paso
            Else
                MsgBox "No existe la imagen se copio el paso "
                Clipboard.Clear
                Clipboard.SetText Paso
            End If
   Exit Sub
salir:
MsgBox "No existe la imagen"
End Sub
Private Function MAX_DOCUMENTOS_DIGITALES_2() As Long
 Dim rs As New ADODB.Recordset
    rs.Open "SELECT MAX(ID) AS MaxDoc FROM DOCUMENTOS_DIGITALES", ConActiva, 0, 1
    MAX_DOCUMENTOS_DIGITALES_2 = rs!Maxdoc + 1
End Function


