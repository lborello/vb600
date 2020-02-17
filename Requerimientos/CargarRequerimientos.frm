VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCargarRequerimientos 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Carga de Requerimientos"
   ClientHeight    =   10785
   ClientLeft      =   120
   ClientTop       =   90
   ClientWidth     =   14820
   DrawStyle       =   5  'Transparent
   FillColor       =   &H00C00000&
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "CargarRequerimientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10785
   ScaleWidth      =   14820
   Begin VB.Frame fraTravase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   180
      TabIndex        =   16
      Top             =   4620
      Width           =   13995
      Begin VB.TextBox txtDescripcion 
         Height          =   4395
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   12975
      End
      Begin VB.TextBox txtCantidadElemento 
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1980
         TabIndex        =   18
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label lblCantidadcajas 
         Caption         =   "Cantidad"
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
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   360
         Width           =   1755
      End
   End
   Begin VB.Frame fraCajas 
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
      Height          =   6255
      Left            =   120
      TabIndex        =   9
      Top             =   3900
      Width           =   14115
      Begin VB.CommandButton Command2 
         Caption         =   "Leer datos"
         Height          =   315
         Left            =   6480
         TabIndex        =   42
         Top             =   300
         Width           =   1080
      End
      Begin VB.CommandButton cmdNotificarPorCorreo 
         Caption         =   "Notificar por Correo"
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Top             =   5100
         Width           =   1815
      End
      Begin VB.CommandButton cmdColector 
         Caption         =   "Lectura"
         Height          =   315
         Left            =   5520
         TabIndex        =   29
         Top             =   300
         Width           =   840
      End
      Begin VB.TextBox txtDescripcionCajaLibro 
         Height          =   2115
         Left            =   2340
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   3960
         Width           =   10935
      End
      Begin VB.CommandButton cmdBorrarCaja 
         Caption         =   "Borrar"
         Height          =   315
         Left            =   3240
         Picture         =   "CargarRequerimientos.frx":058A
         TabIndex        =   11
         Top             =   300
         Width           =   840
      End
      Begin VB.TextBox txtCajaLibroDesde 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Top             =   300
         Width           =   840
      End
      Begin VB.TextBox txtCajaLibroHasta 
         BackColor       =   &H00C0FFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   300
         Width           =   840
      End
      Begin MSFlexGridLib.MSFlexGrid grdCajasLibros 
         Height          =   2415
         Left            =   300
         TabIndex        =   15
         Top             =   780
         Width           =   13395
         _ExtentX        =   23627
         _ExtentY        =   4260
         _Version        =   393216
         Cols            =   6
         ForeColor       =   8388608
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
      Begin VB.Label lblcantCajasLibros 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   5040
         TabIndex        =   24
         Top             =   300
         Width           =   405
      End
      Begin VB.Label Label11 
         Caption         =   "Label11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   75
         Left            =   840
         TabIndex        =   23
         Top             =   900
         Width           =   15
      End
      Begin VB.Label lbldesc 
         Caption         =   "Descripción"
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
         Left            =   480
         TabIndex        =   21
         Top             =   4500
         Width           =   1155
      End
      Begin VB.Label lblCajaLibro 
         Caption         =   "Desde:"
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
         TabIndex        =   14
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label6 
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
         Height          =   195
         Left            =   1620
         TabIndex        =   13
         Top             =   360
         Width           =   600
      End
      Begin VB.Label lblcajas 
         Caption         =   "Cantidad:"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   4140
         TabIndex        =   12
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame fraCupones 
      Caption         =   "Cupones"
      Height          =   6255
      Left            =   120
      TabIndex        =   52
      Top             =   3900
      Width           =   14115
      Begin MSFlexGridLib.MSFlexGrid grdCupones 
         Height          =   3015
         Left            =   360
         TabIndex        =   54
         Top             =   1080
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmdCargarPlanilla 
         Caption         =   "Cargar Planilla"
         Height          =   495
         Left            =   240
         TabIndex        =   53
         Top             =   420
         Width           =   2295
      End
   End
   Begin VB.CheckBox chkFlete 
      Caption         =   "Flete"
      Height          =   315
      Left            =   4860
      TabIndex        =   51
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CheckBox chkCobrar 
      Caption         =   "Cobrar"
      Height          =   315
      Left            =   2700
      TabIndex        =   50
      Top             =   1980
      Width           =   1995
   End
   Begin VB.CheckBox chkEncontrado 
      Caption         =   "Encontrado"
      Height          =   315
      Left            =   600
      TabIndex        =   49
      Top             =   1980
      Width           =   1635
   End
   Begin MSComctlLib.StatusBar staRequerimiento 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   30
      Top             =   10170
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Estado"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
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
   End
   Begin MSComctlLib.ImageList img16x16 
      Left            =   10560
      Top             =   300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":0BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":0EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":1BA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":1EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":21D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":2628
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":2A7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":2D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":2EEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":3900
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":4312
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":4D24
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":5736
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":6148
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":6B5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":756C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":7F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":8990
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":93A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":9DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":A7C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":B1D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CargarRequerimientos.frx":BBEA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrCambioImagen 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5460
      Top             =   2280
   End
   Begin VB.Frame fraTipoRecepcion 
      Height          =   1275
      Left            =   120
      TabIndex        =   26
      Top             =   660
      Width           =   14115
      Begin VB.Frame Frame2 
         Caption         =   "Recepción"
         Height          =   975
         Left            =   300
         TabIndex        =   35
         Top             =   180
         Width           =   6255
         Begin MSMask.MaskEdBox calFechaRecepciom 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   300
            TabIndex        =   47
            Top             =   420
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox maskHorafax 
            Height          =   315
            Left            =   3480
            TabIndex        =   36
            Top             =   420
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648384
            Enabled         =   0   'False
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            Caption         =   "Hora "
            Height          =   255
            Left            =   3600
            TabIndex        =   37
            Top             =   120
            Width           =   555
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Compromiso de entrega"
         Height          =   975
         Left            =   6900
         TabIndex        =   32
         Top             =   180
         Width           =   6555
         Begin VB.ComboBox cboHoraDia 
            BackColor       =   &H00FFC0C0&
            Height          =   345
            ItemData        =   "CargarRequerimientos.frx":C5FC
            Left            =   2940
            List            =   "CargarRequerimientos.frx":C606
            TabIndex        =   34
            Text            =   "Combo1"
            Top             =   420
            Width           =   2895
         End
         Begin MSMask.MaskEdBox calFechaCompromiso 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            Height          =   315
            Left            =   300
            TabIndex        =   48
            Top             =   420
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            Caption         =   "Hora"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4020
            TabIndex        =   33
            Top             =   120
            Width           =   615
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   600
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   14115
      _ExtentX        =   24897
      _ExtentY        =   1058
      ButtonWidth     =   1296
      ButtonHeight    =   953
      ImageList       =   "img16x16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo Requerimiento"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modif."
            Key             =   "Modificacion"
            Object.ToolTipText     =   "Modificación de un Requerimiento"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anular"
            Key             =   "Anular"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aceptar"
            Key             =   "Aceptar"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Esp. Fax"
            Key             =   "Esp. Fax"
            Object.ToolTipText     =   "Esperando Fax"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   6
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Control"
            Key             =   "Control"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fin"
            Key             =   "Fin"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrControlFax 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5460
      Top             =   2760
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1575
      Left            =   300
      TabIndex        =   2
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
   Begin VB.TextBox txtEstadoRequerimiento 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7920
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame fraDatosRequerimiento 
      Caption         =   "Datos del Requerimiento"
      Height          =   1515
      Left            =   120
      TabIndex        =   3
      Top             =   2340
      Width           =   14115
      Begin Controles.ctlClienteUsuario ctlClienteUsuario 
         Height          =   375
         Left            =   5880
         TabIndex        =   46
         Top             =   600
         Width           =   6195
         _ExtentX        =   10927
         _ExtentY        =   661
      End
      Begin Controles.cltGenerico ctlCliente 
         Height          =   435
         Left            =   5880
         TabIndex        =   45
         Top             =   180
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   767
      End
      Begin Controles.cltGenerico ctlTipoRequerimiento 
         Height          =   315
         Left            =   780
         TabIndex        =   44
         Top             =   240
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
      End
      Begin Controles.cltGenerico ctlPersonal 
         Height          =   375
         Left            =   780
         TabIndex        =   43
         Top             =   600
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   661
      End
      Begin VB.ComboBox cboSucursal 
         BackColor       =   &H00C0E0FF&
         Height          =   345
         ItemData        =   "CargarRequerimientos.frx":C61A
         Left            =   1260
         List            =   "CargarRequerimientos.frx":C627
         TabIndex        =   41
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CommandButton cmdBuscarAsignacion 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11040
         TabIndex        =   38
         Top             =   180
         Width           =   375
      End
      Begin MSMask.MaskEdBox maskHoraLimite 
         Height          =   315
         Left            =   4380
         TabIndex        =   7
         Top             =   2460
         Width           =   15
         _ExtentX        =   26
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   5
         Format          =   "hh:mm "
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox maskDiaLImite 
         Height          =   315
         Left            =   1200
         TabIndex        =   6
         Top             =   2400
         Width           =   0
         _ExtentX        =   0
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label9 
         Caption         =   "Sucursal Resolucion"
         Height          =   435
         Left            =   120
         TabIndex        =   40
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Tomo:"
         Height          =   315
         Left            =   60
         TabIndex        =   39
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblsector 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   5880
         TabIndex        =   28
         Top             =   1020
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "Sector:"
         Height          =   195
         Left            =   4920
         TabIndex        =   27
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo:"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Solicita"
         Height          =   255
         Left            =   4920
         TabIndex        =   5
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   315
         Left            =   4920
         TabIndex        =   4
         Top             =   240
         Width           =   675
      End
   End
   Begin MSComDlg.CommonDialog PasoPlanilla 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "Z:\Administracion\BUSQUEDA DISCO\*.xls"
   End
   Begin VB.Label lblID_fax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   22
      Top             =   8460
      Width           =   1455
   End
End
Attribute VB_Name = "frmCargarRequerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Const PasoBuscarFax = "C:\Fax\"
    Const PasoGrabarFax = "C:\Fax\Dos\"
    Dim PasoImagen  As String
    Dim Revisar As Boolean
    Dim EspFax As Boolean
    Dim BARRAFORM As Integer
    Dim ControlCajasConLegajos As Boolean
'Win API to determine display capabilities
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long


Private Sub ControlCajasVacias()
Dim sql As String
        If Not IsNull(ctlCliente.Valor) Then
            
            
            If ctlTipoRequerimiento.Valor = 7 Then
            Dim rsVacias As ADODB.Recordset
            Set rsVacias = New ADODB.Recordset
            sql = " SELECT IDREQUERIMIENTO From Requerimiento "
            sql = sql & vbCrLf & " Where IDTIPOREQUERIMIENTO = 7 AND IDESTADO < 5 AND ANULADO Is Null "
            sql = sql & vbCrLf & " AND ID_CLIENTE = " & ctlCliente.Valor
            rsVacias.Open sql, ConActiva, 0, 1
            If Not rsVacias.EOF Then
                MsgBox "Existen requerimientos pendientes " & rsVacias!IDREQUERIMIENTO, vbCritical
                ctlTipoRequerimiento.Valor = Null
                ctlCliente.Valor = Null
            End If
        End If
        End If

End Sub

Private Sub ctlClientesRequerimiento_KeyPress(KeyAscii As Integer)
 Dim i As Integer
' For i = 0 To ctlClientesRequerimiento.ListCount
'    If UCase(Chr(KeyAscii)) = UCase(Mid(ctlClientesRequerimiento.List(i), 7, 1)) Then
'        ctlClientesRequerimiento.ListIndex = i
'        Exit For
'    End If
' Next
 
End Sub
Private Sub cmdBorrarCaja_Click()
    Dim Grilla() As String
    ReDim Grilla(grdCajasLibros.Rows - 1, grdCajasLibros.Cols - 1)
    Dim R As Long
    Dim c As Long
    Dim RM As Long
    Dim Cm As Long
    RM = grdCajasLibros.Rows - 1
    Cm = grdCajasLibros.Cols - 1
    
   If txtCajaLibroDesde = "" Then
        If MsgBox("Usted quiuere Borrar todas las cajas", vbQuestion + vbYesNo) = vbYes Then
             BorraGrilla
        End If
   Else
        If MsgBox("Usted quiere borra la Caja/Libro " & txtCajaLibroDesde, vbQuestion + vbYesNo) = vbYes Then
                If txtCajaLibroDesde <> "" Then
                    For R = 1 To grdCajasLibros.Rows - 1
                        For c = 1 To grdCajasLibros.Cols - 1
                            If grdCajasLibros.TextMatrix(R, c) = txtCajaLibroDesde Then
                                grdCajasLibros.TextMatrix(R, c) = ""
                            End If
                        Next
                    Next
                End If
                For R = 1 To grdCajasLibros.Rows - 1
                    For c = 1 To grdCajasLibros.Cols - 1
                        If grdCajasLibros.TextMatrix(R, c) <> "" Then
                            Grilla(R, c) = grdCajasLibros.TextMatrix(R, c)
                        End If
                    Next
                Next
                BorraGrilla
                For R = 0 To RM
                    For c = 0 To Cm
                    If Grilla(R, c) <> "" Then
                        CargarGrilla (Grilla(R, c))
                    End If
                    Next
                Next
        End If
    End If
ContarCajas
End Sub

Private Sub cmdCargarPlanilla_Click()




        PasoPlanilla.ShowOpen
     

        BusquedaDisco (PasoPlanilla.FileName)

End Sub

Private Sub cmdCopiarExcel_Click()




End Sub

Private Sub Command1_Click()

Dim emailOutlookApp As Outlook.Application
Dim emailNameSpace As Outlook.Namespace
Dim emailFolder As Outlook.MAPIFolder
Dim emailItem As Outlook.MailItem
Dim EmailRecipient As Recipient
Dim emailItem2 As Outlook.MailItem

'-----Open Outlook in a background process and the Inbox Folder-----
Set emailOutlookApp = CreateObject("Outlook.Application")
Set emailNameSpace = emailOutlookApp.GetNamespace("MAPI")
Set emailFolder = emailNameSpace.GetDefaultFolder(olFolderInbox)

MsgBox emailFolder.Folders.Count
Dim i As Integer

For i = 1 To emailFolder.Folders.Count
    MsgBox emailFolder.Folders.Item(i).Name

Next

MsgBox emailFolder.Folders.Item("SUPER").Items.Count

Dim sql As String

For i = 1 To emailFolder.Folders.Item("SUPER").Items.Count

    Set emailItem2 = emailFolder.Folders.Item("SUPER").Items(i)
   Rem  MsgBox emailItem2.Body
    emailItem2.Categories = "REQUE:" & i & " id " & emailItem2.EntryID
    Rem emailItem2.FlagStatus = olFlagComplete
    
     emailItem2.FlagRequest = i
     
     sql = " INSERT INTO dbo.CORREOS"
      sql = sql & " (ENTRYID, ENVIADO, ASUNTO, CUERPO, KF_USUARIO)"
sql = sql & " VALUES ('" & emailItem2.InternetCodepage & emailItem2.SenderEmailAddress & emailItem2.Subject & "','" & emailItem2.SenderEmailAddress & "','" & emailItem2.Subject & "','" & Replace(Mid(emailItem2.Body, 1, 2000), "'", "`") & "'," & ctlPersonal.Valor & ")"
         ExecutarSql sql
     
    emailItem2.Save
Next


Set emailNameSpace = Nothing
Set emailFolder = Nothing
Set emailItem = Nothing
Set emailOutlookApp = Nothing

MsgBox "ok"

End Sub

Private Sub Command2_Click()
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
        CargarGrilla CStr(dato)
   End If
    
    
'
'    datoInicio = Mid(L, comienzo)
'    espacio = InStr(datoInicio, "&")
'    dato = Mid(datoInicio, 1, espacio - 1)
'    comienzo = espacio + 1
    Rem CargarGrilla CStr(dato)
Next
dato = Mid(L, Inicio, i - Inicio)
CargarGrilla CStr(dato)
ContarCajas
Exit Sub

salir:
MsgBox Err.Description
   
End Sub

Private Sub cmdBuscarAsignacion_Click()
    Dim rs As New ADODB.Recordset
    Dim sql As String
    If Not IsNull(ctlCliente.Valor) Then
        Dim dato As String
        sql = " SELECT COMPROMISO_ENTREGA, ID_CLIENTE, FECHAENTREGA "
        sql = sql & " From Requerimiento "
        
        sql = " SELECT     COMPROMISO_ENTREGA, ID_CLIENTE, FECHAENTREGA,"
        sql = sql & "     Tiporequerimiento.DESCRIPCION"
        sql = sql & "  FROM         REQUERIMIENTO INNER JOIN"
        sql = sql & "             TIPOREQUERIMIENTO ON REQUERIMIENTO.IDTIPOREQUERIMIENTO = TIPOREQUERIMIENTO.IDTIPOREQUERIMIENTO"
        sql = sql & " WHERE  ID_CLIENTE = " & ctlCliente.Valor
        sql = sql & " AND FECHAENTREGA >= " & FechaFormato(Format(SysDate, "DD/MM/YYYY"))
        sql = sql & "  ORDER BY  FECHAENTREGA , COMPROMISO_ENTREGA , REQUERIMIENTO.IDTIPOREQUERIMIENTO  "
        rs.Open sql, ConActiva, 0, 1
        
        If rs.EOF Then
            MsgBox "No existen Compromisos", vbInformation
                    
        Else
            Do While Not rs.EOF
                dato = dato & " Fecha " & rs!FECHAENTREGA & " Turno: " & rs!COMPROMISO_ENTREGA & "  Tipo: " & Trim(rs!DESCRIPCION) & vbCrLf
                rs.MoveNext
            Loop
            MsgBox dato
        End If
        
    Else
        MsgBox "Colocar el cliente"
    End If

End Sub

Private Sub cmdColector_Click()
    Dim rs2 As New ADODB.Recordset
    Dim sql As String
    Dim Lectura As Long
    Dim Cliente As Integer
   On Error GoTo salir
        Lectura = InputBox("Por Favor Ingrese el numero de Lectura ", "Lectura", 0)
        
        
        If ctlTipoRequerimiento.Valor = 1 Or ctlTipoRequerimiento.Valor = 3 Then
        
        sql = " SELECT NUMERO_LECTURA, CAJA, CLIENTE, ORDEN From LECTURACOLECTOR "
        sql = sql & " Where NUMERO_LECTURA = " & Lectura
        sql = sql & " AND CLIENTE < 9000 "
        sql = sql & " ORDER BY ORDEN "
        rs2.Open sql, ConActiva, 0, 1
        Cliente = CInt(ctlCliente.Valor)
        Do While Not rs2.EOF
            
            If Cliente = rs2!Cliente Then
                CargarGrilla CStr(rs2!Caja)
            Else
                MsgBox "Cliente incorrecto"
            End If
            rs2.MoveNext
        Loop
        
        End If
        
        If ctlTipoRequerimiento.Valor = 10 Or ctlTipoRequerimiento.Valor = 11 Then
        
        
         sql = "  SELECT     ID_Lectura_Legajo, Cod_Legajo_cliente, Cliente"
         sql = sql & " From LECTURA_LEGAJO"
         sql = sql & " Where ID_Lectura_Legajo = " & Lectura
        
        
        rs2.Open sql, ConActiva, 0, 1
        Cliente = CInt(ctlCliente.Valor)
        Do While Not rs2.EOF
            
            If Cliente = rs2!Cliente Then
                CargarGrilla CStr(rs2!Cod_Legajo_cliente)
            Else
                MsgBox "Cliente incorrecto"
            End If
            rs2.MoveNext
        Loop
        
        End If
        
        txtCajaLibroDesde = ""
        txtCajaLibroDesde.SelStart = 0
        ContarCajas
salir:
End Sub

Private Sub cmdEnviarCorreo_Click()

End Sub

Private Sub cmdNotificarPorCorreo_Click()
Dim rs As New ADODB.Recordset
Dim sql As String

sql = " SELECT     ID_CLIENTEUSUARIO, CORREO"
sql = sql & " From CLIENTEUSUARIO"
sql = sql & "  Where ID_CLIENTEUSUARIO = " & ctlClienteUsuario.Valor
rs.Open sql, ConActiva, 0, 1
 If Not rs.EOF Then
     If Not IsNull(rs!correo) Then
        SendMail rs!correo, "Elementos en consultas Banco de Archivos", "Estimado Cliente Le informamos que los siguiente elementos ya se  encuentran en consulta" & vbCrLf & txtDescripcionCajaLibro.Text & vbCrLf & "Estamos a su disposición"
        MsgBox "El correo se envio con exito", vbInformation
     Else
        MsgBox "El usuario CLiente No tiene correo"
     End If
     
 End If
 

End Sub

Private Sub ctlCliente_Click()
    If Not IsNull(ctlCliente.Valor) Then
        ctlClienteUsuario.Clear
        ctlClienteUsuario.Valor = Null
        ctlClienteUsuario.LlenarConCliente CInt(ctlCliente.Valor)
        If ctlTipoRequerimiento.Valor = 7 Then
            Rem ControlCajasVacias
        End If
    End If
End Sub
Private Sub ctlClienteUsuario_LostFocusClienteUsuario()
 lblsector.Caption = ctlClienteUsuario.Sector
End Sub

Private Sub ctlTipoRecepcion_Click()
'    If staRequerimiento.Panels.Item("Estado").Text = "" Then
'        MsgBox "El estado es incorrecto", vbInformation
'        Exit Sub
'    End If
'
'
'    fraCajas.Visible = False
'    fraTravase.Visible = False
'    LimpiarCampos
'    maskFechafax.Text = Format(Date, "DD/MM/YYYY")
'    maskHorafax.Text = Format(Time, "HH:MM")
'    ctlTipoRequerimiento.Valor = Null
'    Select Case ctlTipoRecepcion.Valor
'     Case 1 ' fax institucional
'        fraDatosRequerimiento.Visible = False
'        fraInstitucional.Visible = True
'     Case 2 ' fax requerimiento
'        fraDatosRequerimiento.Visible = True
'        fraInstitucional.Visible = False
'     Case 3 'orden Interna
'
'        frmCargarRequerimientos.Caption = " Cargar Requerimiento " & "Nuevo"
'        lblID_fax = ""
'        maskFechafax.Enabled = False
'        maskHorafax.Enabled = False
'        maskFechafax.Text = Format(Date, "DD/MM/YYYY")
'        maskHorafax.Text = Format(Time, "HH:MM")
'        fraDatosRequerimiento.Visible = True
'        fraInstitucional.Visible = False
'        PasoImagen = ""
'     Case 4 'Telefonicamente
'        lblID_fax = ""
'        Rem Grabar = Nuevo
'        frmCargarRequerimientos.Caption = " Cargar Requerimiento " & "Nuevo"
'        fraDatosRequerimiento.Visible = True
'        fraInstitucional.Visible = False
'        maskFechafax.Enabled = True
'        maskHorafax.Enabled = True
'        PasoImagen = ""
'     End Select
End Sub

Private Sub ctlPersonal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If IsNull(ctlPersonal.Valor) Then
            ctlPersonal.Valor = 99
        End If
    
    End If

End Sub

Private Sub ctlTipoRequerimiento_Click()
    LimpiarCampos
    fraCupones.Visible = False
    
If staRequerimiento.Panels(1).Text = "" Then
    MsgBox "Verifique el estado" & " ¿Es nuevo? "
    Exit Sub
End If

    If IsNull(ctlTipoRequerimiento.Valor) Then
        Exit Sub
    End If
        
    Select Case ctlTipoRequerimiento.Valor
    Case 1, 3, 9, 27
            fraCajas.Visible = True
            fraTravase.Visible = False
            PresentacionCajaLibro "Cajas"
    Case 2, 4
            fraCajas.Visible = True
            fraTravase.Visible = False
            PresentacionCajaLibro "Libro"
    
            
    Case 8
            fraCajas.Visible = False
            fraTravase.Visible = True
            
    Case 10, 11
            fraCajas.Visible = True
            fraTravase.Visible = False
            PresentacionCajaLibro "Legajo"
       Case 25
    fraCupones.Visible = True
    fraCajas.Visible = False
    
    
            fraTravase.Visible = False
                 
            
            Case Else
            fraCajas.Visible = False
            fraTravase.Visible = True
   
    End Select
End Sub

Private Sub Form_Activate()
frmCargarRequerimientos.WindowState = 2
ControlCajasConLegajos = False
End Sub

Private Sub Form_Load()
    LimpiarCampos
    lblID_fax = ""
    staRequerimiento.Panels.Item("Estado").Text = "Nuevo"
    ctlCliente.TipoControl = Cliente
    ctlPersonal.TipoControl = PERSONAL
    calFechaRecepciom.Text = Format(SysDate_DD_MM_YYYY, "dd/mm/yyyy")
    If Format(SysDate_DD_MM_YYYY, "ddd") = "vie" Then
        calFechaCompromiso.Text = DateAdd("D", 3, calFechaRecepciom.Text)
    Else
       calFechaCompromiso.Text = DateAdd("D", 1, Format(SysDate_DD_MM_YYYY, "dd/mm/yyyy"))
    End If
   
    If Format(SysDate_DD_MM_YYYY_mm_ss, "HH") > 13 Then
        cboHoraDia.ListIndex = 1
    Else
        cboHoraDia.ListIndex = 0
    End If
    maskHorafax.Text = Format(SysDate_DD_MM_YYYY_mm_ss, "HH:mm")
    Rem ctlTipoRecepcion.TipoControl = Tipo_Recepcion
    ctlTipoRequerimiento.TipoControl = Tipo_Requerimiento
    Rem fraCajas.Visible = False
    fraTravase.Visible = False
    fraDatosRequerimiento.Visible = True
    ctlPersonal.Valor = MDIfrmInicio.StaInicio.Panels.Item(2).Text
   cboSucursal.Text = Sucursal
    fraCupones.Visible = False
    
    
     
End Sub



Private Sub Timer2_Timer()
    Revisar = True
End Sub
Private Sub Timer3_Timer()
 
End Sub

Private Sub tmrCambioImagen_Timer()
'On Error GoTo Salir
'    Dim MyName As String
'    MyName = Dir(PasoImagen & "*.dcx", vbDirectory)
'
'    If MyName = "" Then
'        EspFax = False
'        tmrCambioImagen.Enabled = False
'        PasoImagen = ""
'        Exit Sub
'    Else
'        PasoImagen = MyName
'    End If
'   Beep
'   BARRAFORM = BARRAFORM + 1
'   If BARRAFORM < 2 Then
'        frmCargarRequerimientos.Icon = LoadPicture("\\Server1basa\Sistemas\Iconos\Fax.ico")
'   Else
'        frmCargarRequerimientos.Icon = LoadPicture("\\Server1basa\Sistemas\Iconos\Fax1.ico")
'        BARRAFORM = 0
'   End If
'Salir:
End Sub

Private Sub tmrControlFax_Timer()
'    Dim MyName
'    On Error GoTo Salir
'        MyName = Dir(PasoBuscarFax & "*.dcx", vbDirectory)
'        If MyName = "" Then
'           tmrCambioImagen.Enabled = False
'         Else
'            tmrCambioImagen.Enabled = True
'        End If
'Salir:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo salir:

    Dim ESTADO As String
              
               Select Case Button.Key
               Case "Anular"
               
                If MDIfrmInicio.StaInicio.Panels(2).Text = 19 Or MDIfrmInicio.StaInicio.Panels(2).Text = 31 Or MDIfrmInicio.StaInicio.Panels(2).Text = 48 Then
                    AnularRequerimiento InputBox("Ingrese el Nº de requerimiento para Anular")
                Else
                    MsgBox "No permitido para este usuario"
                End If

               
               

               Case "Nuevo"
                    LimpiarCampos
                    
                    lblID_fax = ""
                    staRequerimiento.Panels.Item("Estado").Text = "Nuevo"
                    fraDatosRequerimiento.Visible = True
               Case "Modificacion"
                    staRequerimiento.Panels.Item("Estado").Text = "Modificacion"
               Case "Aceptar"
                    ESTADO = staRequerimiento.Panels.Item("Estado").Text
                    If Validar Then
                            Select Case ESTADO
                            Case "Nuevo"
                            If cboHoraDia.Text = "MAÑANA" Then
                                If DateDiff("d", SysDate_DD_MM_YYYY, calFechaCompromiso.Text) = 1 Then
                                    If Format(SysDate_DD_MM_YYYY_mm_ss, "HH") > 14 Then
                                        If MsgBox("QUE NO SE DEBE CARGAR REQUERIMIENTOS DESPUES DE LAS 14HS PARA EL DIA SIGUIENTE A LA MAÑAÑA " & vbCrLf & "QUIERE CONTINUAR", vbYesNo) = vbYes Then
                                        
                                        
                                        Else
                                        Rem Exit Sub
                                        End If
                                     End If
                                 End If
                            End If
                            If cboHoraDia.Text = "TARDE." Then
                                If DateDiff("d", SysDate_DD_MM_YYYY, calFechaCompromiso.Text) = 0 Then
                                    If Format(SysDate_DD_MM_YYYY_mm_ss, "HH") > 11 Then
                                        MsgBox "QUE NO SE DEBE CARGAR REQUERIMIENTOS DESPUES DE LAS 11 PARA EL MISMO DIA A LA TARDE "
                                        Rem Exit Sub

                                     End If
                                 End If
                            End If
                                GrabarNuevo
                            Case "Modificacion"
                                GrabarModificacion
                            End Select
                           staRequerimiento.Panels.Item("Estado").Text = ""
                     End If
                     If Format(SysDate_DD_MM_YYYY, "ddd") = "vie" Then
        calFechaCompromiso.Text = DateAdd("D", 3, Format(SysDate_DD_MM_YYYY, "dd/mm/yyyy"))
    Else
       calFechaCompromiso.Text = DateAdd("D", 1, Format(SysDate_DD_MM_YYYY, "dd/mm/yyyy"))
    End If
   
    If Format(SysDate_DD_MM_YYYY_mm_ss, "HH") > 13 Then
        cboHoraDia.ListIndex = 1
    Else
        cboHoraDia.ListIndex = 0
    End If
                     
                Case "Esp. Fax"
                     EsperarFax
                Case "Buscar"
                    Rem Grabar = EsperandoFax
                    frmCargarRequerimientos.Caption = " Cargar Requerimiento " & "EsperandoFax"
                    
                Case "Fin"
                    frmCargarFechaHora.Show
                Case "Control"
                    frmControlEstados.Show
                    frmControlEstados.SetFocus
                    
                End Select
                  Rem oleImgEdit1.Refresh
                  Exit Sub
salir:
MsgBox Err.Description

End Sub

Private Sub Toolbar1_ButtonDropDown(ByVal Button As MSComctlLib.Button)
If Toolbar1.Buttons.Item("Esp. Fax").Value = tbrUnpressed Then
   Toolbar1.Buttons.Item("Esp. Fax").Value = tbrPressed
   Else
   Toolbar1.Buttons.Item("Esp. Fax").Value = tbrUnpressed
   End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
If Toolbar1.Buttons.Item("Esp. Fax").Value = tbrUnpressed Then
   Toolbar1.Buttons.Item("Esp. Fax").Value = tbrPressed
   Else
   Toolbar1.Buttons.Item("Esp. Fax").Value = tbrUnpressed
   End If
End Sub
Private Sub txtCajaLibroDesde_KeyPress(KeyAscii As Integer)
    Dim F As Long
        If KeyAscii = 13 Then
            If txtCajaLibroHasta = "" Then
                If IsNumeric(txtCajaLibroDesde.Text) Then
                    CargarGrilla CStr(txtCajaLibroDesde)
                    txtCajaLibroDesde = ""
                    txtCajaLibroDesde.SelStart = 0
                    ContarCajas
                Else
                    If UCase(Mid(txtCajaLibroDesde.Text, 1, 2)) = "L1" Then
                        CargarGrilla CStr(CLng(Mid(txtCajaLibroDesde, 6)))
                        txtCajaLibroDesde = ""
                        txtCajaLibroDesde.SelStart = 0
                        ContarCajas
                    Else
                        MsgBox "Error en dato", vbInformation
                        txtCajaLibroDesde.Text = ""
                    End If
                    
                End If
             End If
        End If
End Sub
Private Sub txtCajaLibroHasta_Change()
    If Not IsNumeric(txtCajaLibroHasta) And txtCajaLibroHasta <> "" Then
           txtCajaLibroHasta = Mid(txtCajaLibroHasta, 1, Len(txtCajaLibroHasta) - 1)
    End If
End Sub
Public Sub CargarGrilla(Valor As String)
    Dim c As Integer
    Dim R As Integer
    Dim RsEstado As ADODB.Recordset
    Dim RsReferencia As ADODB.Recordset
    Dim rsSector As ADODB.Recordset
    Dim rsCajaConLegajos As New ADODB.Recordset
    Dim sql As String
    Dim strSector As String
    Dim LargoSector As Integer
    Dim ID_CLIENTE As Integer
    Dim TipoElemento As String
    ID_CLIENTE = ctlCliente.Valor
    Dim rsLegajoDes As ADODB.Recordset
    On Error GoTo salir
    
    If Not IsNumeric(Valor) Then
     MsgBox " No es un numero "
     Exit Sub
     
    
    End If
    
    If ctlTipoRequerimiento.Valor = 1 Or ctlTipoRequerimiento.Valor = 3 Then  ' consulta en planta
            sql = " SELECT     NRO_CAJA, COD_CLIENTE"
            sql = sql & " From basasql.dbo.LEGAJOS"
            sql = sql & " Where Cod_cliente =  " & ctlCliente.Valor
            sql = sql & " And NRO_CAJA = " & Valor
            sql = sql & " And COD_ESTADO = 2"
            rsCajaConLegajos.Open sql, strConBasa
                If Not rsCajaConLegajos.EOF Then
                    If MsgBox("Caja marcada como Legajos " & Valor & " ¿desea continuar? ", vbYesNo) = vbYes Then
                        Exit Sub
                    End If
                End If
            End If


    
        sql = " SELECT     REQUERIMIENTO.ID_CLIENTE, REQUERIMIENTO.IDESTADO, REQUERIMIENTO.ANULADO, REQUERIMIENTO.IDREQUERIMIENTO,"
        sql = sql & "  REQUELIBOSCAJAS.CAJASLIBROS"
        sql = sql & "  FROM     REQUERIMIENTO INNER JOIN "
         sql = sql & "              REQUELIBOSCAJAS ON REQUERIMIENTO.IDREQUERIMIENTO = REQUELIBOSCAJAS.IDREQUERIMIENTOS"
        
        sql = sql & "  WHERE    REQUERIMIENTO.ID_CLIENTE =  " & ctlCliente.Valor
        sql = sql & "  AND (REQUERIMIENTO.IDESTADO < 5) "
        sql = sql & "  AND (REQUERIMIENTO.ANULADO IS NULL) AND"
        sql = sql & "  REQUELIBOSCAJAS.CAJASLIBROS = " & Valor
           Set RsEstado = New ADODB.Recordset
           RsEstado.Open sql, ConActiva, 0, 1
           
           If Not RsEstado.EOF Then
                MsgBox "Ya existe un requerimiento para este numero de elemento " & Valor & " Req:" & RsEstado!IDREQUERIMIENTO, vbCritical
                Exit Sub
           End If
       Select Case ctlTipoRequerimiento.Valor
        Case 2, 4

            Set RsEstado = New ADODB.Recordset
            RsEstado.Open "Select * From libros where cod_cliente= " & ID_CLIENTE & " and  NRO_LIBRO_INTERNO  = " & Valor, ConActiva, 0, 1
            TipoElemento = "El Libro"
        Case 1, 3, 9, 27
            
            Set RsEstado = New ADODB.Recordset
            
'
'            sql = " SELECT     NRO_CAJA, COD_CLIENTE"
'            sql = sql & " From contenedor "
'            sql = sql & " Where NRO_CAJA = " & Valor
'            sql = sql & " And Cod_cliente = " & ID_CLIENTE
'
'
'           RsEstado.Open sql, strConBasa, 0, 1
'
'            If ctlTipoRequerimiento.Valor <> 9 Then ' consulta en planta
'                If RsEstado.EOF Then
'                Else
'                    If ControlCajasConLegajos = True Then
'                        GoTo CARGAR
'                    Else
'                        If MsgBox("Caja marcada como Legajos " & Valor & " ¿desea continuar? ", vbYesNo) = vbYes Then
'                            If MDIfrmInicio.StaInicio.Panels(2).Text = 82 Or MDIfrmInicio.StaInicio.Panels(2).Text = 31 Or MDIfrmInicio.StaInicio.Panels(2).Text = 17 Then
'                                ControlCajasConLegajos = True
'                                CargarGrilla (Valor)
'                            Else
'                                MsgBox " No esta Autorizado "
'                                Exit Sub
'                            End If
'                        Else
'                            Exit Sub
'                        End If
'                    End If
'                End If
'            End If
'
'            Set RsEstado = New ADODB.Recordset
            
            RsEstado.Open "Select * From contenedor where cod_cliente= " & ID_CLIENTE & " and nro_caja = " & Valor, ConActiva, 0, 1
            Set rsSector = New ADODB.Recordset
            sql = " SELECT COD_INDICE"
            sql = sql & " From CLIENTEUSUARIO "
            sql = sql & "  Where ID_CLIENTEUSUARIO = " & ctlClienteUsuario.Valor
            rsSector.Open sql, ConActiva, 0, 1
            strSector = rsSector!Cod_Indice
            LargoSector = Len(strSector)
            Set RsReferencia = New ADODB.Recordset
            sql = "SELECT INDICE From REFERENCIAS "
            sql = sql & " WHERE COD_CLIENTE =" & ID_CLIENTE
            sql = sql & " AND NRO_CAJA = " & Valor
            RsReferencia.Open sql, ConActiva, 0, 1
            If RsReferencia.EOF Then
               MsgBox "La Caja No posee Referencia", vbCritical
            End If
           
            
            Do While Not RsReferencia.EOF
                If Mid(RsReferencia!Indice, 1, LargoSector) = strSector Then
                
                Else
                    MsgBox "La Caja " & Valor & " pertenece a otro sector ", vbCritical
                    Exit Do
                End If
            
                RsReferencia.MoveNext
            Loop
            
            
            TipoElemento = "La Caja"
        Case 10, 11
        
       Legajos_RecalcularCaracteres_DescripcionRemito " Where ID_CLIENTE_LEGAJO =" & Valor & " And COD_CLIENTE = " & ID_CLIENTE
            
            Set RsEstado = New ADODB.Recordset
            RsEstado.Open " SELECT COD_ESTADO  as estado From LEGAJOS Where ID_CLIENTE_LEGAJO =" & Valor & " And COD_CLIENTE = " & ID_CLIENTE, ConActiva, 0, 1
            TipoElemento = "El Legajo"
            
            
        End Select

        If Not RsEstado.EOF Then
           If IsNull(RsEstado!ESTADO) Then
                MsgBox "El estado  es Nulo"
                Exit Sub
           End If
           
           If CInt(RsEstado!ESTADO) <> 2 Then
                MsgBox " El Elemento " & Valor & " no esta en el ESTADO CORRECTO"
                
                If ctlTipoRequerimiento.Valor = 10 Or ctlTipoRequerimiento.Valor = 10 Then
                    
                    Set rsLegajoDes = New ADODB.Recordset
                    rsLegajoDes.Open "SELECT COD_ESTADO ,  NRO_DESDE, LETRA_DESDE  From LEGAJOS " & " WHERE  COD_CLIENTE = " & ID_CLIENTE & " AND ID_CLIENTE_LEGAJO = " & Valor, ConActiva, 0, 1
                    
                    If Not rsLegajoDes.EOF Then
                            txtDescripcionCajaLibro.Text = txtDescripcionCajaLibro.Text & vbCrLf & TipoElemento & " " & rsLegajoDes!NRO_DESDE & " " & rsLegajoDes!LETRA_DESDE & " se encuentra en consulta "
                    End If
                    
                
               
                Else
                    txtDescripcionCajaLibro.Text = txtDescripcionCajaLibro.Text & TipoElemento & " Nº " & Valor & " se encuentra en consulta "
                End If
                
                Exit Sub
           End If
           
        Else
            MsgBox "No Tiene este elemento: " & Valor, vbInformation
            Exit Sub
        End If

        For R = 1 To grdCajasLibros.Rows - 1
            For c = 1 To grdCajasLibros.Cols - 1
                If grdCajasLibros.TextMatrix(R, c) = Valor Then
                    MsgBox TipoElemento & "  " & Valor & " ya esta Cargada", vbInformation
                    txtCajaLibroDesde = ""
                    txtCajaLibroHasta = ""
                    Exit Sub
                End If
                If grdCajasLibros.TextMatrix(R, c) = "" Then
                    grdCajasLibros.TextMatrix(R, c) = Valor
                    Exit Sub
                End If
            Next
        Next
        grdCajasLibros.AddItem ""
        grdCajasLibros.TextMatrix(grdCajasLibros.Rows - 1, 1) = Valor
        Exit Sub
        
salir:
        
        MsgBox Err.Description
End Sub

Private Sub txtCajaLibroHasta_KeyPress(KeyAscii As Integer)
    Dim F As Long
        If KeyAscii = 13 Then
            If IsNumeric(txtCajaLibroDesde) And CLng(txtCajaLibroDesde) < CLng(txtCajaLibroHasta) Then
                For F = txtCajaLibroDesde To txtCajaLibroHasta
                 CargarGrilla CStr(F)
                 ContarCajas
                txtCajaLibroDesde = ""
                txtCajaLibroHasta = ""
                txtCajaLibroDesde.SetFocus
                Next
            End If
        End If
End Sub

Public Sub ContarCajas()
    Dim R As Long
    Dim c As Long
    Dim ContadorCaja As Long
         For R = 1 To grdCajasLibros.Rows - 1
            For c = 1 To grdCajasLibros.Cols - 1
                If grdCajasLibros.TextMatrix(R, c) <> "" Then
                 ContadorCaja = ContadorCaja + 1
                End If
            Next
         Next
        lblcantCajasLibros.Caption = ContadorCaja
        txtCajaLibroDesde = ""
        txtCajaLibroHasta = ""
End Sub

Public Sub BorraGrilla()
        grdCajasLibros.Clear
        grdCajasLibros.Rows = 2
        grdCajasLibros.TextMatrix(0, 1) = "Cajas"
        grdCajasLibros.TextMatrix(0, 2) = "Cajas"
        grdCajasLibros.TextMatrix(0, 3) = "Cajas"
        grdCajasLibros.TextMatrix(0, 4) = "Cajas"
        grdCajasLibros.TextMatrix(0, 5) = "Cajas"
       
End Sub

Public Sub PresentacionCajaLibro(Titulo As String)
    grdCajasLibros.ColWidth(0) = 100
    grdCajasLibros.ColWidth(1) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(2) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(3) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(4) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(5) = (grdCajasLibros.Width - 210) / 5
    
    grdCajasLibros.ColAlignment(1) = 4
    grdCajasLibros.ColAlignment(2) = 4
    grdCajasLibros.ColAlignment(3) = 4
    grdCajasLibros.ColAlignment(4) = 4
    grdCajasLibros.ColAlignment(5) = 4
    
    
    grdCajasLibros.TextMatrix(0, 1) = Titulo
    grdCajasLibros.TextMatrix(0, 2) = Titulo
    grdCajasLibros.TextMatrix(0, 3) = Titulo
    grdCajasLibros.TextMatrix(0, 4) = Titulo
    grdCajasLibros.TextMatrix(0, 5) = Titulo
   
    lblCajaLibro = Titulo
    lblcantCajasLibros = ""
    fraCajas.Caption = Titulo
End Sub

Public Sub LimpiarCampos()
    txtCajaLibroDesde.Text = ""
    txtCajaLibroHasta.Text = ""
    txtCantidadElemento.Text = ""
    txtDescripcion.Text = ""
    txtDescripcionCajaLibro.Text = ""
    lblcantCajasLibros = ""
   Rem txtMotivo.Text = ""
    Rem txtNombreEmpresa.Text = ""
    lblsector.Caption = ""
    ControlCajasConLegajos = False
    ctlCliente.Valor = Null
    ctlPersonal.Valor = Null
    ctlClienteUsuario.Valor = Null
    maskDiaLImite.Mask = ""
    maskDiaLImite.Text = ""
    maskDiaLImite.Mask = "##/##/####"
    maskHoraLimite.Mask = ""
    maskHoraLimite.Text = ""
    maskHoraLimite.Mask = "##:##"
    
    txtEstadoRequerimiento = ""
    Rem fraCajas.Visible = False
   
'     frmCargarRequerimientos.Height = mdiEntradaSalida.Height - 1200
'     frmCargarRequerimientos.Width = fraDatosRequerimiento.Width + 200
    fraTravase.Visible = False
    BorraGrilla
End Sub
Public Sub GrabarNuevo()
    Dim IDFAX, IDREQUERIMIENTO As Long
    IDFAX = 0
'    strConBasa , 0 ,1.BeginTrans

Dim conReq As New ADODB.Connection

On Error GoTo salir:

conReq.Open strConBasa
conReq.BeginTrans
    On Error GoTo salir
       
        If ctlTipoRequerimiento.Valor = 11 Or ctlTipoRequerimiento.Valor = 10 Then
                If lblcantCajasLibros.Caption = "" Then
                    lblcantCajasLibros.Caption = 0
                End If
            End If
       
            IDREQUERIMIENTO = InsertarRequerimiento(IDFAX, conReq)
           
          If lblcantCajasLibros.Caption <> "" Then
                If lblcantCajasLibros.Caption = 0 Then
                
                Else
                    InsertarRequerimientoDetalle IDREQUERIMIENTO, conReq
                End If
            Else
                InsertarRequerimientoDetalle IDREQUERIMIENTO, conReq
            End If
            Dim sql As String
            
            InsertarHistoricoEstadoRequerimiento IDREQUERIMIENTO
                        If cboHoraDia.Text = "MAÑANA" Then
                                If DateDiff("d", SysDate_DD_MM_YYYY, calFechaCompromiso.Text) = 1 Then
                                    If Format(SysDate_DD_MM_YYYY_mm_ss, "HH") > 14 Then
                                        sql = " Update dbo.Requerimiento "
                                        sql = sql & " SET DESCRIPCION_ACTUALIZADA =6 "
                                        sql = sql & " Where IDREQUERIMIENTO = " & IDREQUERIMIENTO
                                        conReq.Execute sql
                                     End If
                                 End If
                            End If
       
        conReq.CommitTrans
        
        Rem chkElementosEncontrados.Value = 0
        MsgBox "El requerimiento se realizó con exito Nº " & IDREQUERIMIENTO, vbInformation
        Rem fraDatosRequerimiento.Visible = False
        LimpiarCampos
        ctlTipoRequerimiento.Valor = Null
        ctlTipoRequerimiento.SetFocus
    Exit Sub
salir:
  Rem chkElementosEncontrados.Value = 0
     conReq.RollbackTrans
     MsgBox "El requerimiento NO se realizo", vbCritical
End Sub
Private Sub txtDescripcionCajaLibro_DblClick()
    If lbldesc.Top = 0 Then
        lbldesc.Top = 3540
        txtDescripcionCajaLibro.Top = 3780
        txtDescripcionCajaLibro.Height = 495
    Else
        lbldesc.Top = 0
        txtDescripcionCajaLibro.Top = 360
        txtDescripcionCajaLibro.Height = 3915
    End If
End Sub

Public Sub GrabarModificacion()
'    Dim I As Integer
'    Dim DESCRIPCION As String
'    On Error GoTo ERROR:
'    strConBasa , 0 ,1.BeginTrans
'      '------------------------------------- TABLA REQUERIMIENTO  -----------------------------------
'    If Not (txtEstadoRequerimiento = 1 And lblcantCajasLibros.Caption <> "") Then
'        MsgBox "ESTE REQUERIMIETO NO PUDE SER MODIFICADO"
'    End If
'    IDREQUERIMIENTO = CLng(TXTNumeroRequerimiento)
'    DESCRIPCION = "'" & UCase(Trim(txtDescripcionCajaLibro)) & "'"
'    If lblcantCajasLibros.Caption <> "" Then
'        IDESTADO = 2
'        CANTIDAD = lblcantCajasLibros.Caption
'    Else
'        IDESTADO = 1
'        CANTIDAD = "0"
'    End If
'    If cboTipoRequerimiento.ListIndex = 4 Or cboTipoRequerimiento.ListIndex = 5 Or cboTipoRequerimiento.ListIndex = 6 Then 'cajas vacias
'        If txtCantidadCajas <> "" Then
'            CANTIDAD = txtCantidadCajas
'            IDESTADO = 2
'        End If
'    End If
'        Sql = " UPDATE REQUERIMIENTO "
'        Sql = Sql & vbCrLf & " SET IDESTADO = " & IDESTADO
'        Sql = Sql & vbCrLf & ", DESCRIPCION =  " & DESCRIPCION
'        Sql = Sql & vbCrLf & ", CANTIDAD =  " & CANTIDAD
'        Sql = Sql & vbCrLf & " WHERE IDREQUERIMIENTO  =" & NumeroRequerimiento
'        ExecutarSql (Sql)
' '-----------------------------------------------------------------------------------------
'
' '-----------------------------------  TABLA REQUELIBOSCAJAS ---------------------------------
'
'
'        Sql = "DELETE  REQUELIBOSCAJAS "
'        Sql = Sql & vbCrLf & " where IDREQUERIMIENTOS =" & NumeroRequerimiento
'        ExecutarSql (Sql)
'
'
'        Dim c As Integer
'        Dim R As Integer
'        Dim Cajas, libros As Long
'            For R = 1 To grdCajasLibros.Rows - 1
'                For c = 1 To grdCajasLibros.Cols - 1
'                    If grdCajasLibros.TextMatrix(R, c) <> "" Then
'                        Sql = " INSERT INTO REQUELIBOSCAJAS ("
'                        Sql = Sql & vbCrLf & " IDREQUERIMIENTOS, CAJASLIBROS )"
'                        Sql = Sql & vbCrLf & " VALUES ("
'                        Sql = Sql & vbCrLf & IDREQUERIMIENTO & "," & grdCajasLibros.TextMatrix(R, c) & " )"
'                        ExecutarSql (Sql)
'                    End If
'                Next
'            Next
' ' ----------------------------------------------------------------------------------------------
'            IDPERSONAL = CLng(Mid(cboTomo.List(cboTomo.ListIndex), 1, 2))
'
' '------------------------------------ H_ESTADO_REQUE ----------------------------------
'    Sql = "  INSERT INTO H_ESTADO_REQUE ("
'    Sql = Sql & vbCrLf & " IDREQUERIMIENTO, IDESTADO, IDPERSONAL,"
'    Sql = Sql & vbCrLf & " CONTADOR, FECHA )"
'    Sql = Sql & vbCrLf & "  VALUES ("
'    Sql = Sql & vbCrLf & IDREQUERIMIENTO & "," & IDESTADO & "," & IDPERSONAL & ","
'    Sql = Sql & vbCrLf & 1 & "," & SysDate & ")"
'    ExecutarSql (Sql)
' '--------------------------------------------------------------------------------
'
'    ' TABLA FAX
'    lblID_fax = ""
'    strConBasa , 0 ,1.CommitTrans
'    Set rsMaxFax = Nothing
'    LimpiarCampos
'    fraCajas.Visible = False
'    fraTravase.Visible = False
'    fraDatosRequerimiento.Visible = False
'    fraInstitucional.Visible = False
'    If PasoImagen <> "" Then
'      If Dir(PasoImagen) <> "" Then
'        Kill PasoImagen
'      End If
'    End If
'  Rem  PonerImagen ("\\Server1basa\fax\EsperandoFax.bmp")
'    Rem oleImgEdit1.Zoom = 75
'    Grabar = EsperandoFax
'    frmCargarRequerimientos.Caption = " Cargar Requerimiento " & "EsperandoFax"
'
'    Exit Sub
'ERROR:
'    strConBasa , 0 ,1.Rollback
'    End Select
'    Grabar = EsperandoFax
'    frmCargarRequerimientos.Caption = " Cargar Requerimiento " & "EsperandoFax"
End Sub


Public Function Validar() As Boolean
        Validar = False
       
            If ctlPersonal.Valor = Null Then
                MsgBox "Falta el dato quien tomo", vbCritical
                Exit Function
            End If
            If IsNull(ctlCliente.Valor) Then
                 MsgBox "Falta el dato cliente", vbCritical
                 Exit Function
            End If
            
            If IsNull(ctlTipoRequerimiento.Valor) Then
                MsgBox "Falta el dato Tipo requerimiento ", vbCritical
                Exit Function
            Else
                If ctlTipoRequerimiento.Valor = 20 Or ctlTipoRequerimiento.Valor = 24 Then
                   If txtCantidadElemento.Text < 20 Then
                       MsgBox "INGRESAR LA CANTIDAD DE KILOMETROS", vbInformation
                       Exit Function
                   End If
                End If
            End If
            
            
            
            
            If maskHorafax.ClipText = "" Then
                MsgBox "Falta el dato hora de entrada", vbCritical
                Exit Function
            End If
            Select Case ctlTipoRequerimiento.Valor
            Case 5, 6, 7, 12, 13, 14, 15, 16, 17, 18, 19, 8
                If txtCantidadElemento.Text = "" Then
                    MsgBox "Falta La cantidad", vbCritical
                    Exit Function
                End If
            End Select
       
        Validar = True
End Function



Public Sub AnularRequerimiento(IDREQUERIMIENTO As Long)
    Dim sql As String
    Dim rsAnular As New ADODB.Recordset
    Dim R As Integer
    
    sql = " SELECT IDREQUERIMIENTO, IDESTADO, IDREMITO"
    sql = sql & " From Requerimiento Where idRequerimiento = " & IDREQUERIMIENTO
    rsAnular.Open sql, ConActiva, 0, 1
    If Not rsAnular.EOF Then
         If IsNull(rsAnular!IDREMITO) Then
            sql = " Update Requerimiento Set ANULADO = 1 Where IDREQUERIMIENTO = " & IDREQUERIMIENTO
             R = ExecutarSql(sql)
            If R = 1 Then
                MsgBox "El requerimiento fue Anulado", vbInformation
            Else
                MsgBox "El requerimiento NO fue Anulado", vbCritical
            End If
         Else
            sql = " Update Requerimiento Set ANULADO = 1 Where IDREQUERIMIENTO = " & IDREQUERIMIENTO
             R = ExecutarSql(sql)
            frmRemitoEntrada.AnularRemito rsAnular!IDREMITO
            If R = 1 Then
                MsgBox "El requerimiento fue Anulado", vbInformation
            Else
                MsgBox "El requerimiento NO fue Anulado", vbCritical
            End If
         End If
    Else
        MsgBox "El requerimiento NO fue Anulado", vbCritical
    End If
    
        
    
   
    End Sub

Public Sub EsperarFax()
'Dim DataTimeFax As String
'    Dim MyName As String
'                    MyName = Dir(PasoBuscarFax & "*.dcx", vbDirectory)
'                    If MyName <> "" Then
'                            Shell "c:\i_view32.exe " & PasoBuscarFax & MyName, vbMaximizedFocus
'                            Rem Grabar = Nuevo
'                            PasoImagen = PasoBuscarFax & MyName
'                            frmCargarRequerimientos.Caption = " Cargar Requerimiento " & "Nuevo"
'                            frmCargarRequerimientos.Icon = LoadPicture("\\Server1basa\Sistemas\Iconos\Fax.ico")
'                            DataTimeFax = Format(CDate(FileDateTime(PasoBuscarFax & MyName)), "DD/MM/YYYY HH:MM:SS")
'                            Rem AppActivate "Requerimientos"
'                            frmCargarRequerimientos.WindowState = 0
'                            DoEvents
'                            EspFax = False
'                            tmrCambioImagen.Enabled = False
'                            maskFechafax.Text = Format(DataTimeFax, "DD/MM/YYYY")
'                            maskHorafax.Text = Format(DataTimeFax, "HH:MM")
'                            PasoImagen = PasoBuscarFax & MyName
'                            fraCajas.Visible = False
'                            fraDatosRequerimiento.Visible = False
'                            fraInstitucional.Visible = False
'                            fraTravase.Visible = False
'                            LimpiarCampos
'                            ctlTipoRequerimiento.Valor = Null
'
'                    Else
'                        PasoImagen = ""
'                        MsgBox "NO existen fax pendientes", vbInformation
'                    End If
End Sub

Public Function InsertarRequerimiento(ByVal IDFAX As Long, conReq As ADODB.Connection) As Long
 Dim IDREQUERIMIENTO As Long
 Dim ID_CLIENTE, IDPERSONAL, TOMO As Integer
 Dim IDTIPORECEPCION, IDESTADO, IDTIPOREQUERIMIENTO  As Integer
 Dim Sector, DESCRIPCION, SOLICITANTE, FECHARECEPCION, FECHAENTREGA, COMPROMISO_ENTREGA, FK_SUCURSAL  As String
 Dim CANTIDAD, COD_USUARIO_CLIENTE As Integer
 Dim Flete As String
 Dim Cobrar As String
 
 Dim EnvioTarde As String
 
 Dim sql As String

    
    
    ID_CLIENTE = ctlCliente.Valor
    IDPERSONAL = ctlPersonal.Valor
    IDTIPORECEPCION = 4
    IDESTADO = ResolucionEstado
    IDTIPOREQUERIMIENTO = ctlTipoRequerimiento.Valor
    IDFAX = IDFAX
    Sector = "'" & ctlClienteUsuario.Sector & "'"
    DESCRIPCION = ResolucionDescripcion
    SOLICITANTE = "'" & ctlClienteUsuario.DESCRIPCION & "'"
    FECHAENTREGA = FechaFormato(calFechaCompromiso.Text)
    COMPROMISO_ENTREGA = "'" & Trim(cboHoraDia.Text) & "'"
     FK_SUCURSAL = "'" & Trim(cboSucursal.Text) & "'"
     Flete = chkFlete.Value
     Cobrar = chkCobrar.Value
 EnvioTarde = 0
    
    
    
    
    
    
 
    FECHARECEPCION = FechaFormato(calFechaRecepciom.Text)
    
    
   
   If chkEncontrado.Value = 1 Then
   IDESTADO = 4
   End If
   Dim MaxRequer As Long
    
    CANTIDAD = ResolucionCantidad
    COD_USUARIO_CLIENTE = ctlClienteUsuario.Valor
    TOMO = ctlPersonal.Valor
   MaxRequer = MaxIDRequerimiento
    sql = "  INSERT INTO REQUERIMIENTO"
    sql = sql & vbCrLf & " (  ID_CLIENTE, IDPERSONAL,"
    sql = sql & vbCrLf & " IDTIPORECEPCION, IDESTADO, IDTIPOREQUERIMIENTO,"
    sql = sql & vbCrLf & " IDFAX, SECTOR, DESCRIPCION, SOLICITANTE,"
    sql = sql & vbCrLf & " FECHARECEPCION, CANTIDAD, COD_USUARIO_CLIENTE,"
    sql = sql & vbCrLf & " TOMO,ENVIOTARDE,FECHAENTREGA ,COMPROMISO_ENTREGA, FECHA_SISTEMA, FK_SUCURSAL, FLETE , COBRAR )"
    sql = sql & vbCrLf & " VALUES ("
    sql = sql & vbCrLf & ID_CLIENTE & "," & IDPERSONAL & ","
    sql = sql & vbCrLf & IDTIPORECEPCION & "," & IDESTADO & "," & IDTIPOREQUERIMIENTO & ","
    sql = sql & vbCrLf & IDFAX & "," & Sector & "," & DESCRIPCION & "," & SOLICITANTE & ","
    sql = sql & vbCrLf & FECHARECEPCION & "," & CANTIDAD & "," & COD_USUARIO_CLIENTE & ","
    sql = sql & vbCrLf & TOMO & "," & EnvioTarde & "," & FECHAENTREGA & "," & COMPROMISO_ENTREGA & "," & SysDate_mm_ss & "," & FK_SUCURSAL
   sql = sql & vbCrLf & "," & Cobrar & "," & Flete & ")"
    conReq.Execute (sql)
    
    
    Dim rs As New ADODB.Recordset
    
    sql = " SELECT     MAX(IDREQUERIMIENTO) AS MaxReq From Requerimiento "
    
    rs.Open sql, conReq
    
   InsertarRequerimiento = rs!MaxReq
    
    
    
    If DESCRIPCION <> "NULL" Then
       Insert_Requerimiento_Historico_Descripcion InsertarRequerimiento, CStr(DESCRIPCION), CInt(IDPERSONAL), SysDate_mm_ss, conReq
    End If
End Function

Public Function ResolucionEstado() As Integer
    ResolucionEstado = 2
    Select Case ctlTipoRequerimiento.Valor
    Case 1, 2, 3, 4, 9, 10, 11, 7, 27
        ResolucionEstado = 2
    Case 6, 15, 5
        ResolucionEstado = 5
    Case 8
        ResolucionEstado = 1
    End Select
    
'    If chkElementosEncontrados.Value = 1 Then
'        ResolucionEstado = 4
'    End If
    
End Function
Public Function ResolucionCantidad() As Integer
 ResolucionCantidad = 0

    Select Case ctlTipoRequerimiento.Valor
    Case 1, 3, 9, 10, 11, 2, 4, 27
    
          ResolucionCantidad = lblcantCajasLibros.Caption
    Case 25
    ResolucionCantidad = grdCupones.Rows - 1
    Case Else
        ResolucionCantidad = txtCantidadElemento.Text
    End Select
End Function

Public Function ResolucionDescripcion() As String
    ResolucionDescripcion = ""
    Select Case ctlTipoRequerimiento.Valor
    Case 1, 3, 9, 10, 11, 2, 4, 27
        If txtDescripcionCajaLibro.Text = "" Then
           ResolucionDescripcion = "Null"
        Else
            ResolucionDescripcion = "'" & Replace(Trim(UCase(txtDescripcionCajaLibro.Text)), "'", "") & "'"
        End If
    Case Else
        If Trim(txtDescripcion.Text) = "" Then
            ResolucionDescripcion = "Null"
        Else
            ResolucionDescripcion = "'" & Replace(Trim(UCase(txtDescripcion.Text)), "'", "") & "'"
        End If
    End Select
    
    
    
    
    
End Function

Public Function InsertarFax() As Long
'    Dim MaxFax As Long
'    Dim Motivo As String
'    Dim NombreEmpresa As String
'    Dim DESCRIPCION As String
'    InsertarFax = 0
'    Dim SQL As String
'    If PasoImagen <> "" Then
'        MaxFax = MaxIDFax
'         FileCopy PasoImagen, PasoGrabarFax & MaxFax & ".dcx"
''         If txtNombreEmpresa.Text <> "" Then
''            NombreEmpresa = "'" & Trim(UCase(txtNombreEmpresa)) & "'"
''        Else
''            NombreEmpresa = "Null"
''        End If
'        If txtMotivo.Text <> "" Then
'            DESCRIPCION = "'" & Trim(UCase(txtMotivo)) & "'"
'        Else
'            DESCRIPCION = "Null"
'        End If
'            SQL = "INSERT INTO FAX"
'            SQL = SQL & vbCrLf & "( IDFAX, NOMBRE, "
'            SQL = SQL & vbCrLf & " DESCRIPCION,  FECHA )"
'            SQL = SQL & vbCrLf & " VALUES ( " & MaxFax & "," & NombreEmpresa & ","
'            SQL = SQL & vbCrLf & DESCRIPCION & "," & SysDate & ")"
'            ExecutarSql (SQL)
'            Kill PasoImagen
'            InsertarFax = MaxFax
'       End If
End Function

Public Sub InsertarRequerimientoDetalle(IDREQUERIMIENTO As Long, conreque As ADODB.Connection)
    Dim c As Integer
    Dim R As Integer
    Dim rsVacias As ADODB.Recordset
    Dim CajaInicio As Long
    Dim j As Long
    Dim Cajafin As Long
    Dim Cajas, libros As Long
    Dim sql As String
    Dim DEPOSITO As String
        Select Case ctlTipoRequerimiento.Valor
        Case 1, 2, 3, 4, 9, 10, 11, 27
            With grdCajasLibros
                  For R = 1 To .Rows - 1
                      For c = 1 To .Cols - 1
                          
                         If ctlTipoRequerimiento.Valor = 1 Or ctlTipoRequerimiento.Valor = 3 Then
                               If .TextMatrix(R, c) <> "" Then
                                    DEPOSITO = "'" & ObtenerDeposito(.TextMatrix(R, c), ctlCliente.Valor, "CAJA") & "'"
                                    sql = " INSERT INTO REQUELIBOSCAJAS ("
                                    sql = sql & vbCrLf & " IDREQUERIMIENTOS, CAJASLIBROS, DEPOSITO, ESTADO )"
                                    sql = sql & vbCrLf & " VALUES ( "
                                    sql = sql & vbCrLf & IDREQUERIMIENTO & "," & .TextMatrix(R, c) & ","
                                    sql = sql & vbCrLf & DEPOSITO & ", NULL )"
                                    conreque.Execute (sql)
                                End If
                          Else
                                If ctlTipoRequerimiento.Valor = 10 Or ctlTipoRequerimiento.Valor = 11 Then
                                    If .TextMatrix(R, c) <> "" Then
                                        DEPOSITO = "'" & ObtenerDeposito(.TextMatrix(R, c), ctlCliente.Valor, "LEGAJO") & "'"
                                        sql = " INSERT INTO REQUELIBOSCAJAS ("
                                        sql = sql & vbCrLf & " IDREQUERIMIENTOS, CAJASLIBROS, DEPOSITO, ESTADO )"
                                        sql = sql & vbCrLf & " VALUES ( "
                                        sql = sql & vbCrLf & IDREQUERIMIENTO & "," & .TextMatrix(R, c) & ","
                                        sql = sql & vbCrLf & DEPOSITO & " , NULL)"
                                        conreque.Execute (sql)
                                    End If
                                Else
                                    If .TextMatrix(R, c) <> "" Then
                                        DEPOSITO = "'" & ObtenerDeposito(.TextMatrix(R, c), ctlCliente.Valor, "CAJA") & "'"
                                        sql = " INSERT INTO REQUELIBOSCAJAS ("
                                        sql = sql & vbCrLf & " IDREQUERIMIENTOS, CAJASLIBROS, DEPOSITO, ESTADO )"
                                        sql = sql & vbCrLf & " VALUES ( "
                                        sql = sql & vbCrLf & IDREQUERIMIENTO & "," & .TextMatrix(R, c) & ","
                                        sql = sql & vbCrLf & DEPOSITO & ", NULL )"
                                        conreque.Execute (sql)
                                    End If
                                End If
                           End If
  
                      Next
                  Next
            End With
        Case 7
'            Set rsVacias = New ADODB.Recordset
'            Sql = " Select max(nro_caja)MaxNumeroCaja "
'            Sql = Sql & vbCrLf & " FROM Contenedor "
'            Sql = Sql & vbCrLf & " Where Cod_Cliente = " & ctlCliente.Valor
'            rsVacias.Open Sql, strConBasa , 0 ,1
'
'                If IsNull(rsVacias!MaxNumeroCaja) Then
'                    CajaInicio = 1
'                    Cajafin = CInt(txtCantidadElemento.Text)
'                Else
'                    CajaInicio = CLng(rsVacias!MaxNumeroCaja) + 1
'                    Cajafin = CLng(rsVacias!MaxNumeroCaja) + CInt(txtCantidadElemento.Text)
'                End If
'                For j = CajaInicio To Cajafin
'                    Sql = " INSERT INTO REQUELIBOSCAJAS ("
'                    Sql = Sql & vbCrLf & " IDREQUERIMIENTOS, CAJASLIBROS )"
'                    Sql = Sql & vbCrLf & " VALUES ("
'                    Sql = Sql & vbCrLf & IDREQUERIMIENTO & "," & j & " )"
'                    ExecutarSql (Sql)
'                Next

        Case 25
        Dim i As Integer
            For i = 2 To grdCupones.Rows - 1
            
            
                     sql = " INSERT INTO REQUELIBOSCAJAS ("
                    sql = sql & vbCrLf & " IDREQUERIMIENTOS, CAJASLIBROS , detalle )"
                    sql = sql & vbCrLf & " VALUES ("
                    sql = sql & vbCrLf & IDREQUERIMIENTO & "," & grdCupones.TextMatrix(i, 1) & " ,'" & Trim(grdCupones.TextMatrix(i, 2)) & "')"
                    ExecutarSql (sql)
            
            Next
            
            
        

        End Select
End Sub

Public Sub InsertarHistoricoEstadoRequerimiento(IDREQUERIMIENTO As Long)
    Dim sql As String
        sql = "  INSERT INTO H_ESTADO_REQUE ("
        sql = sql & vbCrLf & " IDREQUERIMIENTO, IDESTADO, IDPERSONAL,"
        sql = sql & vbCrLf & " CONTADOR, FECHA )"
        sql = sql & vbCrLf & "  VALUES ( "
        sql = sql & vbCrLf & IDREQUERIMIENTO & "," & ResolucionEstado & "," & ctlPersonal.Valor & ","
        sql = sql & vbCrLf & 1 & "," & SysDate & ")"
        ExecutarSql (sql)

End Sub

Public Function ObtenerDeposito(NumeroElemento As Long, Cliente As Integer, TIPO As String) As String
    
    
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    If TIPO = "CAJA" Then
        sql = " SELECT     DEPOSITO "
        sql = sql & " From Cajas "
        sql = sql & " Where FK_CLIENTE = " & Cliente
        sql = sql & " And NRO_CAJA = " & NumeroElemento
            
    Else
       sql = "SELECT     CAJAS.DEPOSITO"
       sql = sql & "  FROM         LEGAJOS INNER JOIN"
       sql = sql & "  CAJAS ON LEGAJOS.NRO_CAJA = CAJAS.NRO_CAJA AND LEGAJOS.COD_CLIENTE = CAJAS.FK_CLIENTE"
       sql = sql & "  Where LEGAJOS.ID_CLIENTE_LEGAJO =  " & NumeroElemento
       sql = sql & "  And lEGAJOS.Cod_cliente = " & Cliente
    End If
    
    
    rs.Open sql, ConActiva, 0, 1
    If Not rs.EOF Then
        If Not IsNull(rs!DEPOSITO) Then
            ObtenerDeposito = UCase(Trim(rs!DEPOSITO))
        Else
            ObtenerDeposito = "NO TIENE"
        End If
        
        
    End If
    
    
    

End Function

Public Sub BusquedaDisco(Paso As String)


    Dim ApExcel As Excel.Application
    Dim libroEx As Excel.Workbook
    Dim hojaEx As Excel.Worksheet
    
    
    On Error GoTo salir
    'abrir hoja excel
    Set ApExcel = New Excel.Application
    Set libroEx = Excel.Workbooks.Open(Paso)
    Set hojaEx = libroEx.Worksheets.Item(1)
    TituloGrillaCupones
    Dim datos As String

    Dim EMISOR As Integer
    Dim SM As Integer
    Dim FECHA_LIMITE As Integer
    Dim NRO_TARJETA As Integer
    Dim Fecha As Integer
    Dim TRANSACCION As Integer
    Dim CUPON As Integer
    Dim LOTE As Integer
    Dim IMPORTE As Integer
    Dim REGION  As Integer
    Dim OBSERVACIONES As Integer
    
    EMISOR = 1
    FECHA_LIMITE = 2
    SM = 3
    NRO_TARJETA = 4
    Fecha = 5
    TRANSACCION = 6
    CUPON = 6
    LOTE = 8
    IMPORTE = 8
    REGION = 10
    OBSERVACIONES = 11
    Dim DETALLE As String
    Dim sql    As String
 Dim i As Integer
 
    Dim rsBuscar As New ADODB.Recordset
    
    hojaEx.Columns(1).Select
  

With hojaEx

  datos = vbTab & "CAJA" & vbTab & "Detalle"
  
 
        For i = 2 To 500

            If .Cells(i, SM) <> "" Then
 
 
 
 

                sql = " SELECT    REFERENCIAS.NRO_CAJA"
                sql = sql & " FROM         REFERENCIAS INNER JOIN"
                sql = sql & " INDICES ON REFERENCIAS.COD_CLIENTE = INDICES.COD_CLIENTE AND REFERENCIAS.INDICE = INDICES.INDICE"
                sql = sql & "  Where (REFERENCIAS.COD_CLIENTE = 1197)"
                sql = sql & "  AND INDICES.ID_CODIGO_DOCUMENTO =  " & (.Cells(i, SM) + 20000)
                sql = sql & "  AND ('" & hojaEx.Cells(i, Fecha) & "' BETWEEN REFERENCIAS.FECHA_DESDE AND   REFERENCIAS.FECHA_HASTA)"
                 
                 Set rsBuscar = New ADODB.Recordset
                 rsBuscar.Open sql, strConBasa
                 
                 DETALLE = "EMISOR:" & hojaEx.Cells(i, EMISOR) & " SM:" & hojaEx.Cells(i, SM) & " FECHA:" & hojaEx.Cells(i, Fecha) & " IMPORTE:" & hojaEx.Cells(i, IMPORTE) & " N° CUPON:" & hojaEx.Cells(i, CUPON) & " NRO_TARJETA :" & hojaEx.Cells(i, NRO_TARJETA) & " FECHA_LIMITE:" & hojaEx.Cells(i, FECHA_LIMITE) & " Original:" & hojaEx.Cells(i, OBSERVACIONES)
                 
                 If rsBuscar.EOF Then
                 datos = datos & vbCrLf & i & vbTab & "0" & vbTab & "NO ENCONTRADO " & DETALLE
                    grdCupones.AddItem i & vbTab & "0" & vbTab & "NO ENCONTRADO " & DETALLE
                 
                 Else
                   Do While Not rsBuscar.EOF
                        grdCupones.AddItem i & vbTab & rsBuscar!NRO_CAJA & vbTab & DETALLE
                        datos = datos & vbCrLf & i & vbTab & rsBuscar!NRO_CAJA & vbTab & DETALLE
                        rsBuscar.MoveNext
                    Loop
                 End If
                 
        End If
                 
                 Next
                 



End With
libroEx.Close

Clipboard.Clear
Clipboard.SetText datos
MsgBox "Los datos fueron copiados"
Exit Sub

salir:

MsgBox Err.Description
End Sub

Public Sub TituloGrillaCupones()

grdCupones.Cols = 3
grdCupones.Clear

grdCupones.Refresh
grdCupones.ColWidth(1) = 1500
grdCupones.ColWidth(2) = 9500
grdCupones.Rows = 1
grdCupones.AddItem 0 & vbTab & "CAJA" & vbTab & "DETALLE"


End Sub
