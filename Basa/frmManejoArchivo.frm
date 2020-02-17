VERSION 5.00
Object = "{ED512BE6-6629-4FB4-953D-D0C353847163}#1.0#0"; "ImagXpr7.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmManejoArchivo 
   Caption         =   "Natalia"
   ClientHeight    =   9930
   ClientLeft      =   840
   ClientTop       =   1665
   ClientWidth     =   14055
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   14055
   Begin TabDlg.SSTab SSTab1 
      Height          =   8835
      Left            =   0
      TabIndex        =   9
      Top             =   60
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   15584
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
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
      TabCaption(0)   =   "Subir Imagenenes"
      TabPicture(0)   =   "frmManejoArchivo.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtFechaIngresoLote"
      Tab(0).Control(1)=   "cmdInsertarLotes"
      Tab(0).Control(2)=   "Frame5"
      Tab(0).Control(3)=   "ImagXpress1"
      Tab(0).Control(4)=   "Command7"
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(6)=   "txtCantidadLotes"
      Tab(0).Control(7)=   "Command11"
      Tab(0).Control(8)=   "ctlIndiceDigitalizacion"
      Tab(0).Control(9)=   "cmdArchivoOrigen"
      Tab(0).Control(10)=   "txtDescripcion"
      Tab(0).Control(11)=   "txtCaja"
      Tab(0).Control(12)=   "ctlcliente"
      Tab(0).Control(13)=   "mskRemito"
      Tab(0).Control(14)=   "Label6"
      Tab(0).Control(15)=   "Label24"
      Tab(0).Control(16)=   "Label23"
      Tab(0).Control(17)=   "Label13"
      Tab(0).Control(18)=   "lblIndice"
      Tab(0).Control(19)=   "Label11"
      Tab(0).Control(20)=   "Label12"
      Tab(0).Control(21)=   "lblPasoImagenesOrigenes"
      Tab(0).Control(22)=   "Label9"
      Tab(0).Control(23)=   "Label7"
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Backup"
      TabPicture(1)   =   "frmManejoArchivo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label27"
      Tab(1).Control(1)=   "Label28"
      Tab(1).Control(2)=   "Label29"
      Tab(1).Control(3)=   "Command9"
      Tab(1).Control(4)=   "txtBackupPasoOrigen"
      Tab(1).Control(5)=   "txtBackupPasoDestino"
      Tab(1).Control(6)=   "txtbackupComenzarendvd"
      Tab(1).Control(7)=   "grdDatosBackup"
      Tab(1).Control(8)=   "cmdCopiarDVD"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmManejoArchivo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(2)=   "Label16"
      Tab(2).Control(3)=   "Frame1"
      Tab(2).Control(4)=   "txtPasoDestino"
      Tab(2).Control(5)=   "chk"
      Tab(2).Control(6)=   "txtImagenesCajas"
      Tab(2).Control(7)=   "cmdExportar"
      Tab(2).Control(8)=   "TxtIndice"
      Tab(2).Control(9)=   "Frame2"
      Tab(2).Control(10)=   "comPaso"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "OSDE AFILIACIONES"
      TabPicture(3)   =   "frmManejoArchivo.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label19"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label20"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label21"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "txtOSDEArchivoHistorico"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtOsdePasoDirectorios"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtOSDELote"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdOsdeProcesarTexto"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      TabCaption(4)   =   "Lotes Generales"
      TabPicture(4)   =   "frmManejoArchivo.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      Begin VB.TextBox txtFechaIngresoLote 
         Height          =   435
         Left            =   -73560
         TabIndex        =   83
         Top             =   2400
         Width           =   2595
      End
      Begin VB.CommandButton cmdInsertarLotes 
         Caption         =   "Insertar Lotes"
         Height          =   435
         Left            =   -70020
         TabIndex        =   81
         Top             =   2400
         Width           =   1635
      End
      Begin VB.Frame Frame5 
         Caption         =   "Fecha"
         Height          =   2175
         Left            =   -70260
         TabIndex        =   76
         Top             =   3540
         Width           =   1995
         Begin VB.TextBox txt_FECHA_REORDENAR 
            Height          =   375
            Left            =   120
            TabIndex        =   80
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txt_FECHA_INDEXACION 
            Height          =   330
            Left            =   120
            TabIndex        =   79
            Top             =   1140
            Width           =   1695
         End
         Begin VB.TextBox txt_FECHA_PREPARACION 
            Height          =   375
            Left            =   120
            TabIndex        =   78
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txt_FECHA_SCANNER 
            Height          =   375
            Left            =   120
            TabIndex        =   77
            Top             =   720
            Width           =   1695
         End
      End
      Begin ImagXpr7Ctl.ImagXpress ImagXpress1 
         Height          =   615
         Left            =   -63660
         TabIndex        =   67
         Top             =   6360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1085
         ErrStr          =   "2917BAFF9E86EAF1B61D37759B64E559"
         ErrCode         =   624630059
         ErrInfo         =   -1890158385
         Persistence     =   -1  'True
         _cx             =   1296
         _cy             =   1085
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
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   315
         Left            =   -65040
         TabIndex        =   66
         Top             =   7920
         Width           =   1335
      End
      Begin VB.CommandButton cmdCopiarDVD 
         Caption         =   "Copiar DVD"
         Height          =   375
         Left            =   -67080
         TabIndex        =   65
         Top             =   1980
         Width           =   2235
      End
      Begin MSFlexGridLib.MSFlexGrid grdDatosBackup 
         Height          =   2235
         Left            =   -74700
         TabIndex        =   64
         Top             =   2820
         Width           =   10740
         _ExtentX        =   18944
         _ExtentY        =   3942
         _Version        =   393216
         Cols            =   5
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
      Begin VB.TextBox txtbackupComenzarendvd 
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
         Left            =   -72480
         TabIndex        =   63
         Top             =   1980
         Width           =   5115
      End
      Begin VB.TextBox txtBackupPasoDestino 
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
         Left            =   -72480
         TabIndex        =   61
         Top             =   1500
         Width           =   5115
      End
      Begin VB.TextBox txtBackupPasoOrigen 
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
         Left            =   -72480
         TabIndex        =   59
         Top             =   1080
         Width           =   5115
      End
      Begin VB.Frame Frame3 
         Caption         =   "Personal"
         Height          =   2175
         Left            =   -74940
         TabIndex        =   57
         Top             =   3540
         Width           =   4455
         Begin Controles.cltGenerico CTL_FK_PERSONAL_REORDENAR 
            Height          =   375
            Left            =   1440
            TabIndex        =   71
            Top             =   1560
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   661
         End
         Begin Controles.cltGenerico CTL_FK_PERSONAL_PREPARACION 
            Height          =   375
            Left            =   1440
            TabIndex        =   73
            Top             =   300
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   661
         End
         Begin Controles.cltGenerico CTL_FK_PERSONAL_SCANNER 
            Height          =   375
            Left            =   1440
            TabIndex        =   74
            Top             =   720
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   661
         End
         Begin Controles.cltGenerico CTL_FK_PERSONAL_INDEXACION 
            Height          =   375
            Left            =   1440
            TabIndex        =   75
            Top             =   1140
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   661
         End
         Begin VB.Label Label31 
            Caption         =   "Reordenar"
            Height          =   255
            Left            =   120
            TabIndex        =   72
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "Indexación"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   1140
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "Preparación"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Digitalización"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   1275
         End
      End
      Begin VB.TextBox txtCantidadLotes 
         Height          =   330
         Left            =   -70920
         TabIndex        =   55
         Text            =   "0"
         Top             =   2040
         Width           =   2475
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Command11"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -65040
         TabIndex        =   53
         Top             =   7440
         Width           =   1275
      End
      Begin VB.CommandButton cmdOsdeProcesarTexto 
         Caption         =   "Procesar Texto"
         Height          =   495
         Left            =   4560
         TabIndex        =   52
         Top             =   2460
         Width           =   2415
      End
      Begin VB.TextBox txtOSDELote 
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
         Left            =   1800
         TabIndex        =   51
         Top             =   1920
         Width           =   8655
      End
      Begin VB.TextBox txtOsdePasoDirectorios 
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
         Left            =   1800
         TabIndex        =   49
         Text            =   "I:\86-osde\Imagenes\Osde\Afiliaciones\Mendoza\"
         Top             =   1440
         Width           =   11655
      End
      Begin VB.TextBox txtOSDEArchivoHistorico 
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
         Left            =   1800
         TabIndex        =   47
         Text            =   "I:\86-osde\Imagenes\Osde\Afiliaciones\Mendoza\Historico Mendoza.txt"
         Top             =   900
         Width           =   11535
      End
      Begin MSComDlg.CommonDialog comPaso 
         Left            =   -70560
         Top             =   2700
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   36
         Top             =   780
         Width           =   10455
         Begin VB.CommandButton cmdPasoDestino 
            Caption         =   "..."
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
            Left            =   9360
            TabIndex        =   43
            Top             =   840
            Width           =   375
         End
         Begin VB.CommandButton cmdPasoOrigen 
            Caption         =   "..."
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
            Left            =   9360
            TabIndex        =   42
            Top             =   360
            Width           =   375
         End
         Begin VB.CommandButton cmdUnirImagenes 
            Caption         =   "Unir Imagenes"
            Height          =   375
            Left            =   7800
            TabIndex        =   41
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtPasoDestinoUnion 
            Height          =   375
            Left            =   1560
            TabIndex        =   39
            Text            =   "D:\ExportarImagenes\5020_MUNI ROSARIO_0090\Luis\"
            Top             =   840
            Width           =   7575
         End
         Begin VB.TextBox txtPasoOrigen 
            Height          =   375
            Left            =   1560
            TabIndex        =   37
            Text            =   "D:\ExportarImagenes\5020_MUNI ROSARIO_0090\00000001\"
            Top             =   360
            Width           =   7575
         End
         Begin VB.Label lblCantidadImagenes 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   2760
            TabIndex        =   45
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Cantidad de imagens Procesadas"
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   1440
            Width           =   2655
         End
         Begin VB.Label Label17 
            Caption         =   "Paso Destino"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lblPasoOrigen 
            Caption         =   "Paso Origen"
            Height          =   255
            Left            =   240
            TabIndex        =   38
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtIndice 
         Height          =   375
         Left            =   -70200
         TabIndex        =   35
         Top             =   3600
         Width           =   3135
      End
      Begin VB.CommandButton cmdExportar 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   -66780
         TabIndex        =   33
         Top             =   4020
         Width           =   1635
      End
      Begin VB.TextBox txtImagenesCajas 
         Height          =   375
         Left            =   -70200
         TabIndex        =   32
         Top             =   4080
         Width           =   3135
      End
      Begin VB.CheckBox chk 
         Caption         =   "Exportar Base Completa"
         Height          =   315
         Left            =   -66720
         TabIndex        =   30
         Top             =   3600
         Width           =   2595
      End
      Begin VB.TextBox txtPasoDestino 
         Height          =   375
         Left            =   -70200
         TabIndex        =   29
         Text            =   "D:\ExportarImagenes\"
         Top             =   3120
         Width           =   6315
      End
      Begin VB.Frame Frame1 
         Caption         =   "Nombre del Archivo"
         Height          =   1455
         Left            =   -74820
         TabIndex        =   24
         Top             =   3120
         Width           =   2415
         Begin VB.OptionButton optLetra_ID 
            Caption         =   "Letra_ desde mas ID"
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
            Left            =   240
            TabIndex        =   27
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton optID 
            Caption         =   "ID"
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
            Left            =   240
            TabIndex        =   26
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton optNumeroCorrelativo 
            Caption         =   "Numerico Corelativo"
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
            Left            =   240
            TabIndex        =   25
            Top             =   1080
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Backup"
         Height          =   375
         Left            =   -67080
         TabIndex        =   23
         Top             =   1080
         Width           =   2175
      End
      Begin Controles.cltIndice ctlIndiceDigitalizacion 
         Height          =   4155
         Left            =   -68160
         TabIndex        =   18
         Top             =   720
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   7329
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
      Begin VB.CommandButton cmdArchivoOrigen 
         Caption         =   "...."
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
         Left            =   -68640
         TabIndex        =   17
         Top             =   5040
         Width           =   315
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   330
         Left            =   -73560
         TabIndex        =   14
         Top             =   1140
         Width           =   5175
      End
      Begin VB.TextBox txtCaja 
         Height          =   330
         Left            =   -73560
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin Controles.cltGenerico ctlcliente 
         Height          =   315
         Left            =   -73560
         TabIndex        =   11
         Top             =   720
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
      End
      Begin MSMask.MaskEdBox mskRemito 
         Height          =   375
         Left            =   -70860
         TabIndex        =   21
         Top             =   1560
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "0001-000#####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Ingreso"
         Height          =   315
         Left            =   -74880
         TabIndex        =   82
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label Label29 
         Caption         =   "Backup a partir del DVD:"
         Height          =   375
         Left            =   -74700
         TabIndex        =   62
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label28 
         Caption         =   "Paso Destino"
         Height          =   255
         Left            =   -74640
         TabIndex        =   60
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label27 
         Caption         =   "Paso Origen:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   58
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Label Label24 
         Caption         =   "C. Lotes"
         Height          =   255
         Left            =   -71700
         TabIndex        =   56
         Top             =   2100
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "Descripción"
         Height          =   255
         Left            =   -74880
         TabIndex        =   54
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label21 
         Caption         =   "Lote"
         Height          =   315
         Left            =   180
         TabIndex        =   50
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label20 
         Caption         =   "Paso Directorios"
         Height          =   315
         Left            =   180
         TabIndex        =   48
         Top             =   1500
         Width           =   1395
      End
      Begin VB.Label Label19 
         Caption         =   "Archivo Historico"
         Height          =   315
         Left            =   180
         TabIndex        =   46
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Indice"
         Height          =   255
         Left            =   -72240
         TabIndex        =   34
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "Copiar Imagenes Cajas"
         Height          =   255
         Left            =   -72240
         TabIndex        =   31
         Top             =   4140
         Width           =   1995
      End
      Begin VB.Label Label14 
         Caption         =   "Paso Exportar"
         Height          =   315
         Left            =   -72240
         TabIndex        =   28
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Remito:"
         Height          =   255
         Left            =   -71700
         TabIndex        =   22
         Top             =   1620
         Width           =   675
      End
      Begin VB.Label lblIndice 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -73560
         TabIndex        =   20
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label11 
         Caption         =   "Indice"
         Height          =   255
         Left            =   -74700
         TabIndex        =   19
         Top             =   1980
         Width           =   1155
      End
      Begin VB.Label Label12 
         Caption         =   "Paso:"
         Height          =   195
         Left            =   -74940
         TabIndex        =   16
         Top             =   5160
         Width           =   435
      End
      Begin VB.Label lblPasoImagenesOrigenes 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   -74400
         TabIndex        =   15
         Top             =   5100
         Width           =   5655
      End
      Begin VB.Label Label9 
         Caption         =   "Cliente"
         Height          =   255
         Left            =   -74700
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Caja"
         Height          =   255
         Left            =   -74760
         TabIndex        =   10
         Top             =   1500
         Width           =   555
      End
   End
   Begin MSComDlg.CommonDialog comArchivoOrigen 
      Left            =   1500
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
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
      Left            =   5160
      TabIndex        =   4
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   4560
      Width           =   3435
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   2580
      Width           =   2955
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   2640
      Width           =   2715
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1740
      TabIndex        =   0
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "Label3"
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
      Left            =   480
      TabIndex        =   8
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
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
      Left            =   300
      TabIndex        =   7
      Top             =   3060
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   1980
      Width           =   1875
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
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
      Left            =   300
      TabIndex        =   5
      Top             =   1020
      Width           =   1335
   End
   Begin VB.Menu mnuArbol 
      Caption         =   "Arbol"
      Begin VB.Menu mnuBuscarIndice 
         Caption         =   "Buscar Indice"
      End
   End
End
Attribute VB_Name = "frmManejoArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdArchivoOrigen_Click()
    comArchivoOrigen.ShowOpen
    lblPasoImagenesOrigenes.Caption = Replace(comArchivoOrigen.FileName, comArchivoOrigen.FileTitle, "") & "*.tif"

End Sub

Private Sub cmdCopiarDVD_Click()
 Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim i As Double
    Dim CD As Integer
    Dim Paso As String
    Dim ImagenInicio As String
    

    Sql = " SELECT  ID, TAMANIO, BACKUP_IMAGEN, DIRECTORIO_PASO"
    Sql = Sql & " From DOCUMENTOS_DIGITALES"
    Sql = Sql & " WHERE     (BACKUP_IMAGEN IS NULL) OR "
    Sql = Sql & " BACKUP_IMAGEN =  " & txtbackupComenzarendvd.Text
    Sql = Sql & " ORDER BY ID"
    
    ConBasa.CommandTimeout = 300
 CD = txtbackupComenzarendvd.Text
Paso = txtBackupPasoDestino.Text



 rs.CursorLocation = adUseClient
 rs.Open Sql, ConActiva, 0, 1
Rem rs.Open sql, strConBasa , 0 ,1
Do While Not rs.EOF
   
   If Dir(txtBackupPasoDestino.Text & "DVD" & CD, vbDirectory) = "" Then
        FileSystem.MkDir txtBackupPasoDestino.Text & "DVD" & CD
    End If
   
    If Dir(txtBackupPasoDestino.Text & "DVD" & CD & "\" & rs!DIRECTORIO_PASO, vbDirectory) = "" Then
        FileSystem.MkDir txtBackupPasoDestino.Text & "DVD" & CD & "\" & rs!DIRECTORIO_PASO
      
    End If
    
    On Error GoTo salir
    If Dir(txtBackupPasoOrigen.Text & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF") = "" Then
        Debug.Print txtBackupPasoOrigen.Text & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF"
    End If
    If Dir(txtBackupPasoOrigen.Text & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF") <> "" Then
        FileSystem.FileCopy txtBackupPasoOrigen.Text & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF", txtBackupPasoDestino.Text & "DVD" & CD & "\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF"
    Else
        MsgBox "No existe el archivo " & rs!ID
    End If


salir:
   
    
    Debug.Print rs!ID
  
   
    rs.MoveNext
   
Loop



End Sub

Private Sub cmdExportar_Click()
' Dim ApExcel As Excel.Application
'    Dim libroEx As Excel.Workbook
'    Dim hojaEx As Excel.Worksheet
'
'
' Dim Min As Long
'    Dim Max As Long
'
'    'abrir hoja excel
'    Set ApExcel = New Excel.Application
'    Set libroEx = Excel.Workbooks.Add
'    Set hojaEx = libroEx.Worksheets.Item(1)
'    Dim rs As New ADODB.Recordset
'    Dim sql As String
'    Dim PasoInicial As String
'
'         PasoInicial = "D:\ExportarImagenes\"
'
'    Dim i As Integer
'    Dim R As Excel.Range
'    Dim h As Excel.Hyperlinks
'    i = 1
'
'
'
'            hojaEx.Cells(1, 1) = "Tipo Documento"
'            hojaEx.Cells(1, 2) = "Dato"
'            hojaEx.Cells(1, 3) = "Nombre"
'            hojaEx.Cells(1, 4) = "Caja"
'
'
'
'
'
'        Dim sgrabar As String
'
'
'
'
'sql = " SELECT  INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.LETRA_DESDE, DOCUMENTOS_DIGITALES.Nombre, DOCUMENTOS_DIGITALES.ID,"
'sql = sql & vbCrLf & " DOCUMENTOS_DIGITALES.NRO_DESDE , DOCUMENTOS_DIGITALES.DIRECTORIO_PASO, DOCUMENTOS_DIGITALES.NRO_CAJA"
'sql = sql & vbCrLf & "  FROM   INDICES INNER JOIN"
'sql = sql & vbCrLf & " DOCUMENTOS_DIGITALES ON INDICES.COD_CLIENTE = DOCUMENTOS_DIGITALES.COD_CLIENTE AND"
'sql = sql & vbCrLf & " INDICES.Indice = DOCUMENTOS_DIGITALES.Indice"
'sql = sql & vbCrLf & "  Where DOCUMENTOS_DIGITALES.COD_CLIENTE = " & ctlcliente.Valor
'sql = sql & vbCrLf & "  ORDER BY INDICES.DESCRIPCION, DOCUMENTOS_DIGITALES.NRO_CAJA"
'
'Dim NombreArchivo As String
'
'        i = 2
'        rs.Open sql, strConBasa , 0 ,1
'
'            Do While Not rs.EOF
'                i = i + 1
'
'                If optLetra_ID.Value = 1 Then
'                    NombreArchivo = Trim(rs!LETRA_DESDE) & "_" & CStr(rs!ID) & ".tif"
'                End If
'                If optID.Value = 1 Then
'                    NombreArchivo = CStr(rs!ID) & ".tif"
'                End If
'                If optNumeroCorrelativo.Value = 1 Then
'                    NombreArchivo = i & ".tif"
'                End If
'
'                hojaEx.Cells(i, 1) = rs!DESCRIPCION
'                hojaEx.Cells(i, 1).Hyperlinks.Add hojaEx.Cells(i, 1), ".\" & rs!NRO_CAJA & "\" & NombreArchivo
'
'                If IsNull(rs!LETRA_DESDE) Then
'                    hojaEx.Cells(i, 2) = rs!NRO_DESDE
'                Else
'                    hojaEx.Cells(i, 2) = rs!LETRA_DESDE
'                End If
'                hojaEx.Cells(i, 3) = rs!Nombre
'                If Dir(PasoInicial & rs!NRO_CAJA, vbDirectory) = "" Then
'                    FileSystem.MkDir PasoInicial & rs!NRO_CAJA
'                Else
'
'                End If
'                If rs!REMITO = TXTRENITO Then
'
'                FileCopy PasoImagenes & BuscarDirectorioPaso(rs!ID) & "\" & rs!ID & ".tif", PasoInicial & rs!NRO_CAJA & "\" & NombreArchivo
'                hojaEx.Cells(i, 4) = rs!NRO_CAJA
'                hojaEx.Cells(i, 5) = rs!ID
'
'                rs.MoveNext
'            Loop
'
'           libroEx.SaveAs PasoInicial & Format(Now, "DD_MM_YYYY") & "2.xls"
'           libroEx.Close
'           ApExcel.Quit
'           Set ApExcel = Nothing
'           Set libroEx = Nothing
  



End Sub

Private Sub osdediABETICOS()

'    Dim Min As Long
'    Dim Max As Long
'
'    'abrir hoja excel
'    Set ApExcel = New Excel.Application
'    Set libroEx = Excel.Workbooks.Add
'    Set hojaEx = libroEx.Worksheets.Item(1)
'    Dim rs As New ADODB.Recordset
'    Dim sql As String
'
'         TxtDestino.Text = TxtDestino.Text & txtLote.Text & "\"
'
'    Dim i As Integer
'    Dim R As Excel.Range
'    Dim h As Excel.Hyperlinks
'    i = 1
'
'        hojaEx.Columns("A:A").ColumnWidth = 3.57
'        hojaEx.Columns("B:B").ColumnWidth = 7.14
'        hojaEx.Columns("C:C").ColumnWidth = 15.57
'        hojaEx.Columns("C:C").NumberFormat = "@"
'        hojaEx.Columns("D:D").ColumnWidth = 8.57
'        hojaEx.Columns("E:E").ColumnWidth = 11.23
'        hojaEx.Columns("G:G").ColumnWidth = 9.43
'        hojaEx.Columns("H:H").ColumnWidth = 7.57
'        hojaEx.Columns("I:I").ColumnWidth = 6.71
'        hojaEx.Columns("I:I").NumberFormat = "@"
'        hojaEx.Columns("J:J").ColumnWidth = 6.71
'        hojaEx.Range("J1:j1000").Select
'        hojaEx.Columns("J:J").NumberFormat = "@"
'        hojaEx.Columns("K:K").ColumnWidth = 7.43
'        hojaEx.Columns("K:K").NumberFormat = "@"
'
'
'            hojaEx.Cells(1, 1) = "i04t"
'            hojaEx.Cells(1, 2) = "filial"
'            hojaEx.Cells(1, 3) = "nomfoto"
'            hojaEx.Cells(1, 4) = "cantfotos"
'            hojaEx.Cells(1, 5) = "ic"
'            hojaEx.Cells(1, 6) = "filler0"
'            hojaEx.Cells(1, 7) = "feccarg"
'            hojaEx.Cells(1, 8) = "nrobasa"
'            hojaEx.Cells(1, 9) = "criterio"
'            hojaEx.Cells(1, 10) = "nrotram"
'            hojaEx.Cells(1, 11) = "nrotram"
'
'        PegarDatos
'           libroEx.SaveAs TxtDestino.Text & txtLote.Text & ".xls"
'           libroEx.Close
'           ApExcel.Quit
'        Set ApExcel = Nothing
'        Set libroEx = Nothing
  
       
End Sub


Private Sub cmdInsertarLotes_Click()
InsertarLotes
End Sub

Private Sub cmdOsdeProcesarTexto_Click()
    Dim MyName As String
    Dim Archivo As String
    Dim Directorios As New Collection
    Dim i As Integer
    Dim Paso As String
    MousePointer = 11
    Paso = txtOsdePasoDirectorios & txtOSDELote & "\"
    MyName = Dir(Paso & "F*.*", vbDirectory)
        Do While MyName <> ""
            Directorios.Add MyName
            MyName = Dir()
        Loop
        Open Trim(txtOSDEArchivoHistorico.Text) For Append As #2
        For i = 1 To Directorios.Count
            If Dir(Paso & Directorios.Item(i) & "\Index.dat") <> "" Then
            Close #1
                Open Paso & Directorios.Item(i) & "\Index.dat" For Input As #1
                Do While Not EOF(1)
                    Line Input #1, Temp
                    Print #2, txtOSDELote.Text & ";" & Directorios.Item(i) & ";" & Temp
                    DoEvents
                Loop
            End If
            Close #1
        Next
        Close #2
        MousePointer = 0
        MsgBox "terminado"
        
        
End Sub

Private Sub cmdPasoDestino_Click()
comPaso.ShowOpen
txtPasoDestinoUnion.Text = Replace(comPaso.FileName, comPaso.FileTitle, "")
End Sub

Private Sub cmdPasoOrigen_Click()
comPaso.ShowOpen
txtPasoOrigen.Text = Replace(comPaso.FileName, comPaso.FileTitle, "")
End Sub

Private Sub InsertarLotes()
    Dim ID_DOCUMENTOS_DIGITALES_LOTE As Long
    Dim SUB_LOTE As Integer
    Dim FK_CLIENTES As Integer
    Dim FK_INDICES As Integer
    Dim FK_CAJAS As Long
    Dim FK_ESTADO As Integer
    Dim TIPO_DOCUMENTO As String
    Dim Descripcion As String
    Dim REMITO As String
  
  
    Dim FECHA_PREPARACION As String
    Dim FK_PERSONAL_PREPARACION   As Integer
    Dim FK_PERSONAL_SCANNER As Integer
    Dim FECHA_SCANNER As String
    Dim HOJA_RUTA As Long
    Dim FK_LA_CAJA_TOMADOR As Integer
    Dim FK_PERSONAL_REORDENAR As Integer
    Dim FECHA_REORDENAR As String
 
    
    
    Dim Sql As String

On Error GoTo salir:



Rem Dim ID_DOCUMENTOS_DIGITALES_LOTE As Long
    
    
    If txtDescripcion.Text <> "" Then
        Descripcion = Trim(txtDescripcion.Text)
    Else
        MsgBox "Ingrese txtDescripcion"
        Exit Sub
    End If
    
    If Not IsNull(ctlCliente.Valor) Then
        FK_CLIENTES = ctlCliente.Valor
    Else
        MsgBox "Ingrese el Cliente"
        Exit Sub
    End If
    
    If lblIndice.Caption <> "" Then
        FK_INDICES = lblIndice.Caption
    Else
        MsgBox "Ingrese el Indice"
        Exit Sub
    End If
    
    If txtCaja.Text <> "" Then
        FK_CAJAS = txtCaja.Text
    Else
        MsgBox "Ingrese La Caja"
        Exit Sub
    End If
 
    FK_ESTADO = 0
 
     
    
     If txtDescripcion.Text <> "" Then
        Descripcion = "'" & Trim(txtDescripcion.Text) & "'"
     Else
        MsgBox "Ingrese la descripción"
        Exit Sub
     End If
    REMITO = "'" & mskRemito.Text & "'"
    
    
    

    
    If IsDate(txt_FECHA_PREPARACION.Text) Then
        FECHA_PREPARACION = FechaFormato(txt_FECHA_PREPARACION.Text)
    Else
        MsgBox "Ingrese la Fecha de preparacion"
        Exit Sub
    End If
    
    If IsDate(txt_FECHA_SCANNER.Text) Then
        FECHA_SCANNER = FechaFormato(txt_FECHA_SCANNER.Text)
    Else
        MsgBox "Ingrese el Fecha scanner"
        Exit Sub
    End If
    
     
    If IsDate(txt_FECHA_INDEXACION.Text) Then
        FECHA_INDEXACION = FechaFormato(txt_FECHA_INDEXACION.Text)
    Else
        MsgBox "Ingrese el Fecha FECHA_INDEXACION"
        Exit Sub
    End If
    
    
     If IsDate(txt_FECHA_REORDENAR.Text) Then
        FECHA_REORDENAR = FechaFormato(txt_FECHA_REORDENAR.Text)
    Else
        MsgBox "Ingrese el  FECHA_REORDENAR"
        Exit Sub
    End If
    
    
    If Not IsNull(CTL_FK_PERSONAL_PREPARACION.Valor) Then
        FK_PERSONAL_PREPARACION = CTL_FK_PERSONAL_PREPARACION.Valor
    Else
        MsgBox "Ingrese el Personal preparacion"
        Exit Sub
    End If
    
    If Not IsNull(CTL_FK_PERSONAL_SCANNER.Valor) Then
        FK_PERSONAL_SCANNER = CTL_FK_PERSONAL_SCANNER.Valor
    Else
        MsgBox "Ingrese el Personal scanner"
        Exit Sub
    End If
    
    If Not IsNull(CTL_FK_PERSONAL_INDEXACION.Valor) Then
        FK_PERSONAL_INDEXACION = CTL_FK_PERSONAL_INDEXACION.Valor
    Else
        MsgBox "Ingrese el Personal FK_PERSONAL_INDEXACION"
        Exit Sub
    End If
    
    If Not IsNull(CTL_FK_PERSONAL_REORDENAR.Valor) Then
        FK_PERSONAL_REORDENAR = CTL_FK_PERSONAL_REORDENAR.Valor
    Else
        MsgBox "Ingrese el Personal FK_PERSONAL_REORDENAR"
        Exit Sub
    End If

Dim CantidadLotes As Integer
If txtCantidadLotes.Text = "0" Then
    MsgBox "Ingrese CANTIDAD LOTES"
Else
 CantidadLotes = txtCantidadLotes.Text
End If




For i = 1 To CantidadLotes

        Sql = " Insert INTO DOCUMENTOS_DIGITALES_LOTE "
        Sql = Sql & vbCrLf & "("
        Sql = Sql & vbCrLf & "  SUB_LOTE"
        Sql = Sql & vbCrLf & " , DESCRIPCION"
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
        Sql = Sql & vbCrLf & " VALUES ("
        Sql = Sql & vbCrLf & i
        Sql = Sql & vbCrLf & " , " & Descripcion
        Sql = Sql & vbCrLf & " , " & FK_CLIENTES
        Sql = Sql & vbCrLf & " , " & FK_INDICES
        Sql = Sql & vbCrLf & " , " & FK_CAJAS
        Sql = Sql & vbCrLf & " , " & FK_ESTADO
        Sql = Sql & vbCrLf & " , " & FK_PERSONAL_PREPARACION
        Sql = Sql & vbCrLf & " , " & FK_PERSONAL_SCANNER
        Sql = Sql & vbCrLf & " , " & FK_PERSONAL_INDEXACION
        Sql = Sql & vbCrLf & " , " & FK_PERSONAL_REORDENAR
        Sql = Sql & vbCrLf & " , " & REMITO
        Sql = Sql & vbCrLf & " , 0 "
        Sql = Sql & vbCrLf & " , 0 "
        Sql = Sql & vbCrLf & " , " & FECHA_PREPARACION
        Sql = Sql & vbCrLf & " , " & FECHA_SCANNER
        Sql = Sql & vbCrLf & " , " & FECHA_INDEXACION
        Sql = Sql & vbCrLf & " , " & FECHA_REORDENAR
        Sql = Sql & vbCrLf & " , 'CREADO'"
        Sql = Sql & vbCrLf & " , '" & txtFechaIngresoLote.Text & "'"
        Sql = Sql & vbCrLf & " )"
    
     ExecutarSql Sql
Next

MsgBox "Terminado"

'Dim rsMax As New ADODB.Recordset
'
' rsMax.Open "SELECT MAX(ID_DOCUMENTOS_DIGITALES_LOTE) AS MaxDoc From DOCUMENTOS_DIGITALES_LOTE ", ConActiva, 0, 1
'
' ID_DOCUMENTOS_DIGITALES_LOTE = rsMax!Maxdoc
'
'
'
'
'Dim PasoOrigen As String
'Dim IMAGEN_ORIGEN As Integer
'Dim TAMANIO As Long
'Dim DIRECTORIO_PASO As String
'Dim CANTIDAD_IMAGENES_ARCHIVO As Integer
'Dim NombreImagen As String
'Dim NRO_DESDE As String
'Dim LETRA_DESDE As String
'
'
'
'PasoOrigen = Replace(lblPasoImagenesOrigenes.Caption, "*.tif", "")
'
'
'
'       Dim MyName As String
'Dim ID As Long
'Dim NRO_CAJA As Long
'Dim PasoOrigenCompleto As String
'
'
'            MyName = Dir(lblPasoImagenesOrigenes.Caption)
'            CAN = 0
'             PasoOriginal = Replace(lblPasoImagenesOrigenes.Caption, "*.tif", "")
'
'             Do While MyName <> ""
'                    CAN = CAN + 1
'                    NombreImagen = Replace(MyName, ".tif", "")
'                    NombreImagen = Replace(NombreImagen, ".TIF", "")
'
'                    Rem IMAGEN_ORIGEN = NombreImagen
'                    IMAGEN_ORIGEN = Mid(NombreImagen, 1, 4)
'                    TAMANIO = FileLen(PasoOrigen & MyName)
'                    CANTIDAD_IMAGENES_ARCHIVO = cantidadImagenes(PasoOriginal & MyName)
'                    ID = MAX_DOCUMENTOS_DIGITALES
'                    DIRECTORIO_PASO = BuscarDirectorioPaso(ID)
'                    PasoOrigenCompleto = Mid(lblPasoImagenesOrigenes.Caption, 1, Len(lblPasoImagenesOrigenes.Caption) - 5) & MyName
'                    NRO_DESDE = "NULL"
'                    LETRA_DESDE = "NULL"
'                    INSERT_DOCUMENTOS_DIGITALES ID, "'" & PasoOrigenCompleto & "'", IMAGEN_ORIGEN, NRO_CAJA, TAMANIO, _
'                    "'" & DIRECTORIO_PASO & "'", CANTIDAD_IMAGENES_ARCHIVO, ID_DOCUMENTOS_DIGITALES_LOTE, NRO_DESDE, LETRA_DESDE
'                    FileCopy PasoOrigen & MyName, PasoImagenes & DIRECTORIO_PASO & "\" & ID & ".TIF"
'
'
'
'
'
'                    Rem FileCopy PasoOrigen & MyName, "X:\IMAGENES\" & DIRECTORIO_PASO & "\" & ID & ".TIF"
'                    MyName = Dir()
'            Loop
'Dim rsCantidades As ADODB.Recordset
'Dim Cantidad_Archivos As Integer
'Dim Cantidad_Imagenes As Integer
'
'
'Set rsCantidades = New ADODB.Recordset
'Sql = " SELECT     COUNT(*) AS CANTIDAD_ARCHIVOS"
'Sql = Sql & " From DOCUMENTOS_DIGITALES "
'Sql = Sql & " Where FK_DOCUMENTOS_DIGITALES_LOTE = " & ID_DOCUMENTOS_DIGITALES_LOTE
'rsCantidades.Open Sql, ConActiva, 0, 1
'Cantidad_Archivos = rsCantidades!Cantidad_Archivos
'
'
'Set rsCantidades = New ADODB.Recordset
'Sql = " SELECT     SUM(CANTIDAD_IMAGENES) AS CANTIDAD_IMAGENES"
'Sql = Sql & " From DOCUMENTOS_DIGITALES "
'Sql = Sql & " Where FK_DOCUMENTOS_DIGITALES_LOTE = " & ID_DOCUMENTOS_DIGITALES_LOTE
'
'rsCantidades.Open Sql, ConActiva, 0, 1
'Cantidad_Imagenes = rsCantidades!Cantidad_Imagenes
'
'Sql = "  Update DOCUMENTOS_DIGITALES_LOTE"
'Sql = Sql & " SET CANTIDAD_IMAGENES = " & Cantidad_Imagenes & ", CANTIDAD_ARCHIVOS =" & Cantidad_Archivos
'Sql = Sql & " Where ID_DOCUMENTOS_DIGITALES_LOTE = " & ID_DOCUMENTOS_DIGITALES_LOTE
' ExecutarSql Sql
'            MousePointer = 0
'            MsgBox "Cantidad de imagenes " & CAN
            Exit Sub
salir:
MsgBox Err.Description
MsgBox PasoImagenes & DIRECTORIO_PASO & "\" & ID & ".TIF"
Clipboard.SetText PasoImagenes & DIRECTORIO_PASO
            
            

End Sub

Private Sub cmdUnirImagenes_Click()
' Dim docFrente As MODI.Document
'  Dim docAtras As MODI.Document
'
'
'       Dim i As Integer
'       Dim C As Integer
'       C = 0
'
'       If txtPasoDestinoUnion.Text = txtPasoOrigen.Text Then
'        MsgBox "Los Paso SOn Iguales"
'        Exit Sub
'       End If
'
'       If Dir(txtPasoDestinoUnion.Text, vbDirectory) = "" Then
'         FileSystem.MkDir txtPasoDestinoUnion.Text
'       Else
'          If MsgBox("El directorio Ya existe quiere continuar", vbYesNo) = vbYes Then
'          Else
'
'            Exit Sub
'          End If
'
'       End If
'
'            Dim MyName As String
'            MyName = Dir(txtPasoOrigen.Text & "0*.tif", vbDirectory)
'            MousePointer = 11
'            For i = 1 To 700
'                MyName = Dir(txtPasoOrigen.Text & Format(i, "00000000") & ".tif", vbDirectory)
'
'                If MyName <> "" Then
'                    Set docFrente = New MODI.Document
'                    docFrente.Create txtPasoOrigen.Text & Format(i, "00000000") & ".tif"
'                    i = i + 1
'                    Set docAtras = New MODI.Document
'                    docAtras.Create txtPasoOrigen.Text & Format(i, "00000000") & ".tif"
'                    docAtras.Images.Add docFrente.Images.Item(0), docAtras.Images.Item(0)
'                     C = C + 1
'                     lblCantidadImagenes.Caption = C
'                     lblCantidadImagenes.Refresh
'                    docAtras.SaveAs txtPasoDestinoUnion.Text & Format(C, "00000000") & ".tif"
'
'                 Else
'                    MsgBox "Operacion terminada cantidad de archivos: " & C
'                    MousePointer = 0
'                    Exit Sub
'                End If
'            Next
'            CAN = 0
'             Do While MyName <> ""
'               docOrigen.Create txtPasoOrigen.Text & MyName
'               MyName = Dir()
'            Loop
'
'
'                Set docDestino = New MODI.Document
'                docDestino.Create


End Sub

Private Sub Command1_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String
rs.CursorLocation = adUseClient

Sql = " SELECT     ID, DESDE, HASTA, DIRECTORIO_PASO"
Sql = Sql & " From DIRECTORIOS_IMAGENES"

rs.Open Sql, ConActiva, adOpenKeyset, adLockOptimistic

Do While Not rs.EOF
    FileSystem.MkDir "D:\DIRECTORIOS\" & rs!DIRECTORIO_PASO
'    RS!DIRECTORIO_PASO = Format(CStr(RS!DESDE), "0000000") & "-" & Format(CStr(RS!HASTA), "0000000")
'    RS.Update
    rs.MoveNext
Loop




End Sub


Private Sub Command10_Click()
'    Dim docOrigen As MODI.Document
'            Set docOrigen = New MODI.Document
'
'
'
'
'            Dim MyName As String
'            MyName = Dir(txtPasoOCR.Text & "0*.tif", vbDirectory)
'            CAN = 0
'             Do While MyName <> ""
'               docOrigen.Create txtPasoOCR.Text & MyName
'               docOrigen.OCR miLANG_SPANISH, True
'               docOrigen.Save
''               Update DOCUMENTOS_DIGITALES
''Set OCR = 1
''Where (ID = 10)
'
'               MyName = Dir()
'            Loop
'
'
'                Set docDestino = New MODI.Document
'                docDestino.Create
End Sub

Private Sub Command11_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
Dim cantidad As Integer


Sql = " SELECT     ID, CANTIDAD_IMAGENES, DIRECTORIO_PASO"
Sql = Sql & " From DOCUMENTOS_DIGITALES"
Sql = Sql & " Where (CANTIDAD_IMAGENES Is Null) and id > 7000"
Sql = Sql & " ORDER BY ID"

rs.Open Sql, ConActiva, 0, 1
Do While Not rs.EOF
cantidad = 0
    cantidad = cantidadImagenes(PasoImagenes & rs!DIRECTORIO_PASO & "\" & rs!ID & ".tif")

Sql = " Update DOCUMENTOS_DIGITALES "
Sql = Sql & " Set CANTIDAD_IMAGENES = " & cantidad
Sql = Sql & "  Where ID = " & rs!ID
ExecutarSql Sql
Debug.Print rs!ID
    rs.MoveNext
Loop

End Sub

Private Sub Command3_Click()
Dim i As Long
Dim Paso As String

For i = 10 To 1000000
    Paso = BuscarDirectorioPaso(CLng(i))

Next


End Sub

Private Sub Command4_Click()


Dim rs As New ADODB.Recordset
Dim Sql As String
rs.CursorLocation = adUseClient
Dim TAMAÑO As Long
Dim coD_DIRECTORIO As Long
Dim Paso As String

Sql = " SELECT     id, Cod_cliente"
Sql = Sql & " FROM DOCUMENTOS_DIGITALESUNICOS2 "
Sql = Sql & "  WHERE     (Cod_cliente IN (83, 84, 86, 95))"
Sql = Sql & "  ORDER BY id"
rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
If Dir("\\server001\imagenesSql\" & rs!ID & ".TIF") <> "" Then

 TAMAÑO = FileSystem.FileLen("\\server001\imagenesSql\" & rs!ID & ".TIF")
 coD_DIRECTORIO = rs!ID
 Paso = BuscarDirectorioPaso(coD_DIRECTORIO)

Sql = " INSERT INTO IMAGENESOK   (ID, COD_ESTADO, TAMANIO, COD_DIRECTORIO, DIRECTORIO_PASO)"
Sql = Sql & " VALUES (" & rs!ID & ",10," & TAMAÑO & "," & coD_DIRECTORIO & ",'" & Paso & "' )"

FileSystem.FileCopy "\\server001\imagenesSql\" & rs!ID & ".TIF", "\\Base\SQL\" & Paso & "\" & rs!ID & ".TIF"
 ExecutarSql Sql
 Else
 Debug.Print rs!ID
 End If
rs.MoveNext
Loop



End Sub

Private Sub Command5_Click()

Dim rs As New ADODB.Recordset
rs.CursorLocation = adUseClient
Dim Sql As String
Dim Legajo As String


Sql = " SELECT     ELEMENTO ,  ID"
Sql = Sql & " From REARCHIVO_DIGITAL_DETALLE"
Sql = Sql & "   WHERE NOT ISNULL ELEMENTO"
Sql = Sql & "  ORDER BY ID"


Sql = "  SELECT     ID, ELEMENTO"
Sql = Sql & " From REARCHIVO_DIGITAL_DETALLE"
Sql = Sql & " Where (Not (ELEMENTO Is Null))"
Sql = Sql & " ORDER BY ID"


Sql = "  SELECT     COD_DOCUMENTO, ID, ELEMENTO,COD_DOCUMENTO"
Sql = Sql & " From REARCHIVO_DIGITAL_DETALLE"
Sql = Sql & " Where (Not (ELEMENTO Is Null))"
Sql = Sql & " ORDER BY COD_DOCUMENTO"



rs.Open Sql, ConActiva, adOpenKeyset, adLockOptimistic


Dim CODCLIENTE As Integer
Dim CODINTERCLIENTE As Integer


Do While Not rs.EOF
If CODCLIENTE = rs!COD_DOCUMENTO Then
Else
    CODCLIENTE = rs!COD_DOCUMENTO
    CODINTERCLIENTE = BUSCARCODCLIUENTE(CODCLIENTE, 40)
End If
    If Mid(rs!Elemento, 1, Len(CODINTERCLIENTE)) = CStr(CODINTERCLIENTE) Then

    rs!Elemento = Replace(rs!Elemento, Str(CODINTERCLIENTE), "")
    rs.Update
    End If
    rs.MoveNext
Loop

End Sub

Public Function BUSCARCODCLIUENTE(NUMERODOCUMENTO As Integer, Cliente As Integer) As Integer

Dim Sql As String
Dim rs As New ADODB.Recordset

Sql = "SELECT     COD_CLIENTE, ID_CODIGO_DOCUMENTO, DESCRIPCION, CODIGO_INTERNO_CLIENTE"
Sql = Sql & " From INDICES"
Sql = Sql & "  WHERE     COD_CLIENTE =  " & Cliente
Sql = Sql & "  AND ID_CODIGO_DOCUMENTO = " & NUMERODOCUMENTO

rs.Open Sql, ConActiva, 0, 1

If Not rs.EOF Then
    If IsNull(rs!CODIGO_INTERNO_CLIENTE) Then
        BUSCARCODCLIUENTE = 0
    Else
        BUSCARCODCLIUENTE = rs!CODIGO_INTERNO_CLIENTE
    End If
    
End If

End Function

Private Sub Command6_Click()


Sql = " SELECT     ID, PASO_INTERNO, ELEMENTO, NOMBRE_ARCHIVO, COD_DOCUMENTO"
Sql = Sql & " From REARCHIVO_DIGITAL_DETALLE"
Sql = Sql & " Where (Not (NOMBRE_ARCHIVO Is Null))"
Sql = Sql & " ORDER BY ID "



Dim rs As New ADODB.Recordset

rs.CursorLocation = adUseClient
Dim TAMAÑO As Long
Dim coD_DIRECTORIO As Long
Dim Paso As String

Sql = " SELECT     ID, PASO_INTERNO, ELEMENTO, NOMBRE_ARCHIVO, COD_DOCUMENTO"
Sql = Sql & " From REARCHIVO_DIGITAL_DETALLE"
Sql = Sql & " Where (Not (NOMBRE_ARCHIVO Is Null))"
Sql = Sql & " ORDER BY ID "



rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
If Dir("D:\Montemar\Directorios\" & rs!PASO_INTERNO & "\" & rs!ID & ".TIF") <> "" Then

 TAMAÑO = FileSystem.FileLen("D:\Montemar\Directorios\" & rs!PASO_INTERNO & "\" & rs!ID & ".TIF")
 coD_DIRECTORIO = rs!ID
 Paso = BuscarDirectorioPaso(coD_DIRECTORIO)

Sql = " INSERT INTO IMAGENESOK   (ID, COD_ESTADO, TAMANIO, COD_DIRECTORIO, DIRECTORIO_PASO)"
Sql = Sql & " VALUES (" & rs!ID & ",10," & TAMAÑO & "," & coD_DIRECTORIO & ",'" & Paso & "' )"

FileSystem.FileCopy "D:\Montemar\Directorios\" & rs!PASO_INTERNO & "\" & rs!ID & ".TIF", "\\Finanzas\Imagenes\" & Paso & "\" & rs!ID & ".TIF"
 ExecutarSql (Sql)
 Else
 Debug.Print rs!ID
 End If
rs.MoveNext
Loop



End Sub

Private Function INSERT_DOCUMENTOS_DIGITALES(ID As Long, PasoOrigen As String, IMAGEN_ORIGEN As Integer _
 , NRO_CAJA As Long, TAMANIO As Long, DIRECTORIO_PASO As String, _
 Cantidad_Imagenes As Integer, FK_DOCUMENTOS_DIGITALES_LOTE As Long _
 , NRO_DESDE As String, LETRA_DESDE As String) As Integer
Dim Sql As String
Cod_Estado = 0
FECHA_INCORPORACION = SysDate

'INSERT INTO basasql.dbo.DOCUMENTOS_DIGITALES
'                      (PASOORIGEN, IMAGEN_ORIGEN, FECHA_INCORPORACION, NRO_CAJA, COD_ESTADO, TAMANIO, DIRECTORIO_PASO, CANTIDAD_IMAGENES,
'                      FK_DOCUMENTOS_DIGITALES_LOTE, NRO_DESDE, LETRA_DESDE)
'VALUES     (,,,,,,,,,,)

        Sql = "  INSERT INTO basasql.dbo.DOCUMENTOS_DIGITALES"
        Sql = Sql & vbCrLf & " ("
        Sql = Sql & vbCrLf & " PASOORIGEN"
        Sql = Sql & vbCrLf & " , IMAGEN_ORIGEN"
        Sql = Sql & vbCrLf & " , FECHA_INCORPORACION"
        Sql = Sql & vbCrLf & " , NRO_CAJA"
        Sql = Sql & vbCrLf & " , COD_ESTADO"
        Sql = Sql & vbCrLf & " , TAMANIO"
        Sql = Sql & vbCrLf & " , DIRECTORIO_PASO "
        Sql = Sql & vbCrLf & " , CANTIDAD_IMAGENES"
        Sql = Sql & vbCrLf & " , FK_DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & vbCrLf & " , NRO_DESDE"
        Sql = Sql & vbCrLf & " , LETRA_HASTA"
        Sql = Sql & vbCrLf & " ) "
        Sql = Sql & vbCrLf & " VALUES "
        Sql = Sql & vbCrLf & " ("
        Sql = Sql & vbCrLf & PasoOrigen
        Sql = Sql & vbCrLf & "," & IMAGEN_ORIGEN
        Sql = Sql & vbCrLf & "," & FECHA_INCORPORACION
        Sql = Sql & vbCrLf & "," & NRO_CAJA
        Sql = Sql & vbCrLf & "," & Cod_Estado
        Sql = Sql & vbCrLf & "," & TAMANIO
        Sql = Sql & vbCrLf & "," & DIRECTORIO_PASO
        Sql = Sql & vbCrLf & "," & Cantidad_Imagenes
        Sql = Sql & vbCrLf & "," & FK_DOCUMENTOS_DIGITALES_LOTE
        Sql = Sql & vbCrLf & "," & NRO_DESDE
        Sql = Sql & vbCrLf & "," & LETRA_DESDE
        Sql = Sql & vbCrLf & ")"




ExecutarSql Sql
End Function

Private Function MAX_DOCUMENTOS_DIGITALES() As Long
 Dim rs As New ADODB.Recordset
    rs.Open "SELECT MAX(ID) AS MaxDoc FROM DOCUMENTOS_DIGITALES", ConActiva, 0, 1
    MAX_DOCUMENTOS_DIGITALES = rs!Maxdoc + 1
End Function


Private Sub Command8_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset

Sql = " SELECT  REARCHIVO_DIGITAL_DETALLE.ID, INDICES.INDICE, REARCHIVO_DIGITAL.FECHA, REARCHIVO_DIGITAL.NRO_CAJA,"
  Sql = Sql & " REARCHIVO_DIGITAL_DETALLE.Lote , REARCHIVO_DIGITAL_DETALLE.ELEMENTO, REARCHIVO_DIGITAL_DETALLE.DESCRIPCION"
  Sql = Sql & "  FROM         REARCHIVO_DIGITAL INNER JOIN"
   Sql = Sql & "                      REARCHIVO_DIGITAL_DETALLE ON REARCHIVO_DIGITAL.ID = REARCHIVO_DIGITAL_DETALLE.COD_REARCHIVO_DIGITAL INNER JOIN"
     Sql = Sql & "                    INDICES ON REARCHIVO_DIGITAL.COD_CLIENTE = INDICES.COD_CLIENTE AND"
      Sql = Sql & "                   REARCHIVO_DIGITAL_DETALLE.COD_DOCUMENTO = INDICES.ID_CODIGO_DOCUMENTO"


Sql = " SELECT     REARCHIVO_DIGITAL_DETALLE.ID, REARCHIVO_DIGITAL_DETALLE.ELEMENTO, REARCHIVO_DIGITAL_DETALLE.DESCRIPCION,"
   Sql = Sql & "                     DOCUMENTOS_DIGITALES.LETRA_DESDE, REARCHIVO_DIGITAL_DETALLE.COD_DOCUMENTO, DOCUMENTOS_DIGITALES.COD_CLIENTE,"
      Sql = Sql & "                  INDICES.COD_CLIENTE AS Expr1, INDICES.INDICE,  REARCHIVO_DIGITAL_DETALLE.LOTE,  REARCHIVO_DIGITAL.NRO_CAJA"
 

   Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
     Sql = Sql & "                     REARCHIVO_DIGITAL_DETALLE ON DOCUMENTOS_DIGITALES.ID = REARCHIVO_DIGITAL_DETALLE.ID INNER JOIN"
     Sql = Sql & "                     INDICES ON REARCHIVO_DIGITAL_DETALLE.COD_DOCUMENTO = INDICES.ID_CODIGO_DOCUMENTO INNER JOIN"
      Sql = Sql & "                    REARCHIVO_DIGITAL ON REARCHIVO_DIGITAL_DETALLE.COD_REARCHIVO_DIGITAL = REARCHIVO_DIGITAL.ID"
   Sql = Sql & "  WHERE     (DOCUMENTOS_DIGITALES.LETRA_DESDE IS NULL) AND (NOT (REARCHIVO_DIGITAL_DETALLE.ELEMENTO IS NULL)) AND"
       Sql = Sql & "                   (INDICES.COD_CLIENTE = 40)"


Sql = " SELECT     REARCHIVO_DIGITAL_DETALLE.ID, REARCHIVO_DIGITAL_DETALLE.ELEMENTO, REARCHIVO_DIGITAL_DETALLE.DESCRIPCION,"
 Sql = Sql & "                      DOCUMENTOS_DIGITALES.LETRA_DESDE, REARCHIVO_DIGITAL_DETALLE.COD_DOCUMENTO, DOCUMENTOS_DIGITALES.COD_CLIENTE,"
 Sql = Sql & "                      REARCHIVO_DIGITAL_DETALLE.Lote , REARCHIVO_DIGITAL_DETALLE.COD_REARCHIVO_DIGITAL, INDICES.Indice"
Sql = Sql & "  FROM         DOCUMENTOS_DIGITALES INNER JOIN"
Sql = Sql & "                       REARCHIVO_DIGITAL_DETALLE ON DOCUMENTOS_DIGITALES.ID = REARCHIVO_DIGITAL_DETALLE.ID INNER JOIN"
 Sql = Sql & "                      INDICES ON REARCHIVO_DIGITAL_DETALLE.COD_DOCUMENTO = INDICES.ID_CODIGO_DOCUMENTO"
Sql = Sql & " WHERE     (DOCUMENTOS_DIGITALES.LETRA_DESDE IS NULL) AND (NOT (REARCHIVO_DIGITAL_DETALLE.ELEMENTO IS NULL)) AND"
Sql = Sql & "                       (INDICES.COD_CLIENTE = 40)"



rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
   Sql = "  Update DOCUMENTOS_DIGITALES"
Sql = Sql & " SET  COD_CLIENTE = 40"
Sql = Sql & ", LETRA_DESDE ='" & rs!Elemento & "'"
If IsNull(rs!Elemento) Then
    Sql = Sql & ", NRO_DESDE =null"
Else
    Sql = Sql & ", NRO_DESDE =" & CDbl(rs!Elemento)
End If
If Not IsNull(rs!Descripcion) Then
Sql = Sql & ", Nombre ='" & rs!Descripcion & "'"
End If
Rem Sql = Sql & ", NRO_CAJA =" & Rs!NRO_CAJA
Sql = Sql & ", NRO_CAJA =0"
Sql = Sql & ", LOTE ='" & rs!lote & "'"
Sql = Sql & ", INDICE ='" & Trim(rs!Indice) & "'"
Sql = Sql & " Where ID = " & rs!ID
ExecutarSql Sql
    rs.MoveNext
Loop



End Sub

Private Sub Command7_Click()

'Dim rs As New ADODB.Recordset
'Dim Sql As String
'Dim imagen As String
'
'rs.CursorLocation = adUseClient
'
'Sql = " SELECT     COD_CLIENTE, LOTE, IMAGEN_ORIGEN, PERSONAL_INDEXACION, Nombre_Archivo_Origen, PASOORIGEN, ID"
'Sql = Sql & " From DOCUMENTOS_DIGITALES"
'Sql = Sql & " Where (COD_CLIENTE = 83) And (Not (Lote Is Null)) And (IMAGEN_ORIGEN Is Null)"
'
'
'rs.Open Sql,ConActiva, adOpenKeyset, adLockOptimistic
'Do While Not rs.EOF
'imagen = Mid(Trim(rs!PasoOrigen), Len(Trim(rs!PasoOrigen)) - 11, 8)
'If IsNumeric(imagen) Then
'    rs!IMAGEN_ORIGEN = CInt(imagen)
'    rs.Update
'    End If
'
'    rs.MoveNext
'Loop

Rem ContarGrilla

End Sub

Private Sub Command9_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim i As Double
    Dim CD As Integer
    Dim Paso As String
    Dim ImagenInicio As String
    

    Sql = " SELECT  ID, TAMANIO, BACKUP_IMAGEN, DIRECTORIO_PASO"
    Sql = Sql & " From DOCUMENTOS_DIGITALES"
    Sql = Sql & " WHERE     (BACKUP_IMAGEN IS NULL) OR "
    Sql = Sql & " BACKUP_IMAGEN > " & txtbackupComenzarendvd.Text
    Sql = Sql & " ORDER BY ID"
    
    ConBasa.CommandTimeout = 300
 CD = txtbackupComenzarendvd.Text
Paso = txtBackupPasoDestino.Text



 rs.CursorLocation = adUseClient
 rs.Open Sql, ConActiva, adOpenKeyset, adLockOptimistic
Rem rs.Open sql, strConBasa , 0 ,1
Do While Not rs.EOF
    i = i + rs!TAMANIO
    If Dir(Paso & "DVD" & CD, vbDirectory) = "" Then
        FileSystem.MkDir Paso & "DVD" & CD
        ImagenInicio = rs!ID
    End If
    If i > 4000000000# Then
        grdDatosBackup.AddItem CD & vbTab & ImagenInicio & vbTab & rs!ID & vbTab & Round(i / 1024 / 1024 / 1024, 2)
        CD = CD + 1
        FileSystem.MkDir Paso & "DVD" & CD
         ImagenInicio = rs!ID
        i = 0
    End If
    If Dir(txtBackupPasoDestino.Text & "DVD" & CD & "\" & rs!DIRECTORIO_PASO, vbDirectory) = "" Then
        FileSystem.MkDir txtBackupPasoDestino.Text & "DVD" & CD & "\" & rs!DIRECTORIO_PASO
    End If
    
    On Error GoTo salir
    If Dir(txtBackupPasoOrigen.Text & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF") = "" Then
        Debug.Print txtBackupPasoOrigen.Text & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF"
    End If
     FileSystem.FileCopy txtBackupPasoOrigen.Text & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF", txtBackupPasoDestino.Text & "DVD" & CD & "\" & rs!DIRECTORIO_PASO & "\" & rs!ID & ".TIF"


salir:
    If Err.Number <> 0 Then
       rs!BACKUP_IMAGEN = -1
      Err.Clear
    Else
       rs!BACKUP_IMAGEN = CD
    End If
    
    Debug.Print rs!ID
  
    rs.Update
    rs.MoveNext
   
Loop





End Sub

Private Sub ctlCliente_Click()
ctlIndiceDigitalizacion.Actualizar ctlCliente.Valor, Nulo, 0
lblIndice.Caption = ""
End Sub

Private Sub ctlIndiceDigitalizacion_DblClick()
 Dim rs As New ADODB.Recordset
 Dim Sql As String
 Sql = " SELECT     ID From INDICES "
Sql = Sql & " WHERE  COD_CLIENTE = " & ctlCliente.Valor
Sql = Sql & "  AND INDICE = '" & ctlIndiceDigitalizacion.Item_Selecionado & "'"


 rs.Open Sql, ConActiva, 0, 1
 lblIndice.Caption = rs!ID
 

End Sub

Private Sub ctlIndiceDigitalizacion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuArbol
 End If
End Sub

Private Sub Form_Load()
ctlCliente.TipoControl = Cliente
CTL_FK_PERSONAL_PREPARACION.TipoControl = Personal
CTL_FK_PERSONAL_SCANNER.TipoControl = Personal
CTL_FK_PERSONAL_INDEXACION.TipoControl = Personal
CTL_FK_PERSONAL_REORDENAR.TipoControl = Personal


txtBackupPasoOrigen.Text = PasoImagenes
grdDatosBackup.ColWidth(1) = 1600
grdDatosBackup.ColWidth(2) = 1600
grdDatosBackup.ColWidth(3) = 1600
grdDatosBackup.ColWidth(4) = 1600
grdDatosBackup.ColWidth(0) = 1600

txt_FECHA_PREPARACION.Text = Format(Now, "DD/MM/YYYY")
txt_FECHA_SCANNER.Text = txt_FECHA_PREPARACION.Text
txt_FECHA_INDEXACION.Text = txt_FECHA_PREPARACION.Text
txt_FECHA_REORDENAR.Text = txt_FECHA_PREPARACION.Text
CTL_FK_PERSONAL_PREPARACION.Valor = 17
CTL_FK_PERSONAL_SCANNER.Valor = 17
CTL_FK_PERSONAL_INDEXACION.Valor = 17
CTL_FK_PERSONAL_REORDENAR.Valor = 17
txtFechaIngresoLote.Text = Format(Now, "dd/MM/YYYY HH:mm:ss")

End Sub

Public Function cantidadImagenes(Paso As String) As Integer
ImagXpress1.FileName = Paso
cantidadImagenes = ImagXpress1.Pages

End Function

Private Sub mnuBuscarIndice_Click()
    ctlIndiceDigitalizacion.BuscarIndice InputBox("Ingrese el indice"), True
End Sub

Private Sub txtLote_Change()
lblIndice.Caption = ""
End Sub

Private Sub txtNro_Caja_Change()
lblIndice.Caption = ""
End Sub

Public Function ValidarDatos() As Boolean
ValidarDatos = True
If IsNull(ctlCliente.Valor) Then
    ValidarDatos = False
    MsgBox "El cliente es incorrecto"
End If
If IsNull(ctlOperadorScanner.Valor) Then
    ValidarDatos = False
    MsgBox "El Operador del scanner"
End If

If txtFechaProcesoScanner.Text = "" Then
    ValidarDatos = False
    MsgBox "La fecha no es la correcta"
End If

If lblIndice.Caption = "" Then
    ValidarDatos = False
    MsgBox "El Indice no es el correcto"
End If

 If txtNro_Caja.Text = "" Then
       ValidarDatos = False
    MsgBox "La Caja no es el correcto"
 
 End If
 
 If txtLote.Text = "" Then
     ValidarDatos = False
    MsgBox "EL Lote no es el correcto"
 End If
 
 
If lblPasoImagenesOrigenes.Caption = "" Then
    ValidarDatos = False
    MsgBox "EL paso no es el correcto"

End If









End Function


Public Function DirControl(Paso) As Boolean
    DirControl = False
If Dir(Paso) = "" Then
   DirControl = True
End If

End Function

