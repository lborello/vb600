VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmfacturacion 
   ClientHeight    =   9870
   ClientLeft      =   -30
   ClientTop       =   840
   ClientWidth     =   13110
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   13110
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13035
      _ExtentX        =   22992
      _ExtentY        =   16960
      _Version        =   393216
      Tabs            =   5
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
      TabCaption(0)   =   "Parametros"
      TabPicture(0)   =   "frmfacturacion.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame8"
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(3)=   "Frame3"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Facturación"
      TabPicture(1)   =   "frmfacturacion.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Cliente"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblTipoFactura"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label24"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label10"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cboTipoComprobante"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtNumeroFactura"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtFechaFactura"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtMesFacturacion"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chkNoBorraGrilla"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdInforme"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame9"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Command5"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "grdfactura"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "ctlClienteFactura"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "Fletes"
      TabPicture(2)   =   "frmfacturacion.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label25"
      Tab(2).Control(1)=   "grdFletes"
      Tab(2).Control(2)=   "cmdActualizarFletes"
      Tab(2).Control(3)=   "txtFechaFlete"
      Tab(2).Control(4)=   "chkSoloPendientes"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Recibos"
      TabPicture(3)   =   "frmfacturacion.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblReciboTotalFacturas"
      Tab(3).Control(1)=   "Label40"
      Tab(3).Control(2)=   "Frame6"
      Tab(3).Control(3)=   "grdReciboFacuta"
      Tab(3).Control(4)=   "Frame7"
      Tab(3).Control(5)=   "cmdAceptarRecibo"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "facturasDatas"
      TabPicture(4)   =   "frmfacturacion.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label61"
      Tab(4).Control(1)=   "Label62"
      Tab(4).Control(2)=   "grdFacturaCustodia"
      Tab(4).Control(3)=   "cmdFacturacionCustodia"
      Tab(4).Control(4)=   "Command45"
      Tab(4).Control(5)=   "txtClienteCustodia"
      Tab(4).Control(6)=   "txtFechaDatas"
      Tab(4).Control(7)=   "chkPasarTodas"
      Tab(4).Control(8)=   "cmdActualizarCliente"
      Tab(4).ControlCount=   9
      Begin VB.CommandButton cmdActualizarCliente 
         Caption         =   "Actualizar Cliente"
         Height          =   375
         Left            =   -64440
         TabIndex        =   185
         Top             =   840
         Width           =   1455
      End
      Begin VB.CheckBox chkPasarTodas 
         Caption         =   "Pasar Todas"
         Height          =   255
         Left            =   -67800
         TabIndex        =   184
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtFechaDatas 
         Height          =   375
         Left            =   -71640
         TabIndex        =   182
         Text            =   "01/07/2016"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtClienteCustodia 
         Height          =   375
         Left            =   -73800
         TabIndex        =   181
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command45 
         Caption         =   "Command4"
         Height          =   375
         Left            =   -66480
         TabIndex        =   179
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdFacturacionCustodia 
         Caption         =   "Facturacion Custodia"
         Height          =   375
         Left            =   -70200
         TabIndex        =   178
         Top             =   720
         Width           =   2055
      End
      Begin MSDataGridLib.DataGrid grdFacturaCustodia 
         Height          =   7455
         Left            =   -74640
         TabIndex        =   177
         Top             =   1440
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   13150
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
      Begin Controles.cltGenerico ctlClienteFactura 
         Height          =   375
         Left            =   960
         TabIndex        =   173
         Top             =   960
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   661
      End
      Begin MSDataGridLib.DataGrid grdfactura 
         Height          =   3855
         Left            =   300
         TabIndex        =   172
         Top             =   5400
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6800
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
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   435
         Left            =   4980
         TabIndex        =   166
         Top             =   9660
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame Frame9 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4515
         Left            =   60
         TabIndex        =   144
         Top             =   9720
         Width           =   12615
         Begin VB.CommandButton Command2 
            Caption         =   "Command2"
            Height          =   435
            Left            =   12300
            TabIndex        =   154
            Top             =   480
            Width           =   195
         End
         Begin VB.CommandButton cmdGrabarFactura 
            Caption         =   "Aceptar Factura"
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
            Left            =   11160
            TabIndex        =   153
            Top             =   4020
            Width           =   1395
         End
         Begin VB.TextBox txtDescripcion_Factura 
            Height          =   555
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   152
            Top             =   3840
            Width           =   10935
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H80000013&
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
            Left            =   2400
            TabIndex        =   151
            Text            =   "Text1"
            Top             =   3420
            Width           =   1095
         End
         Begin VB.TextBox txtCodigo 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   780
            TabIndex        =   150
            Top             =   480
            Width           =   555
         End
         Begin VB.CommandButton cmdInsertFacturacion 
            Caption         =   "..."
            Height          =   345
            Left            =   11940
            TabIndex        =   149
            Top             =   480
            Width           =   315
         End
         Begin VB.TextBox txtTotal 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   11040
            ScrollBars      =   2  'Vertical
            TabIndex        =   148
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtPrecioUnitario 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   10140
            TabIndex        =   147
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtDescripcion 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1380
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   146
            Top             =   480
            Width           =   8655
         End
         Begin VB.TextBox txtCantidad 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   145
            Top             =   480
            Width           =   615
         End
         Begin MSFlexGridLib.MSFlexGrid grdFacturacion 
            Height          =   2355
            Left            =   0
            TabIndex        =   171
            Top             =   960
            Width           =   12375
            _ExtentX        =   21828
            _ExtentY        =   4154
            _Version        =   393216
            Cols            =   5
            BackColor       =   16711401
            WordWrap        =   -1  'True
            AllowUserResizing=   3
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
         Begin VB.Label Label12 
            Caption         =   "IVA "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4200
            TabIndex        =   165
            Top             =   3480
            Width           =   315
         End
         Begin VB.Label lblIVA 
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
            Height          =   315
            Left            =   4620
            TabIndex        =   164
            Top             =   3420
            Width           =   915
         End
         Begin VB.Label lblTotal 
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
            Height          =   375
            Left            =   11520
            TabIndex        =   163
            Top             =   3480
            Width           =   915
         End
         Begin VB.Label Label11 
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   10980
            TabIndex        =   162
            Top             =   3540
            Width           =   555
         End
         Begin VB.Label lblSubTotal 
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
            Height          =   315
            Left            =   1080
            TabIndex        =   161
            Top             =   3420
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "SubTotal"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   160
            Top             =   3480
            Width           =   795
         End
         Begin VB.Label Label5 
            Caption         =   "Codigo"
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
            Left            =   840
            TabIndex        =   159
            Top             =   240
            Width           =   555
         End
         Begin VB.Label Label4 
            Caption         =   "Total"
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
            Left            =   11340
            TabIndex        =   158
            Top             =   180
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Precio/U"
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
            Left            =   10320
            TabIndex        =   157
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
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
            Left            =   3300
            TabIndex        =   156
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Label1 
            Caption         =   "Cant"
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
            TabIndex        =   155
            Top             =   240
            Width           =   435
         End
      End
      Begin VB.CommandButton cmdInforme 
         Caption         =   "..."
         Height          =   315
         Left            =   10020
         TabIndex        =   143
         Top             =   1020
         Width           =   315
      End
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   2235
         Left            =   -71160
         TabIndex        =   131
         Top             =   5580
         Width           =   8295
         Begin VB.TextBox txtIncremento 
            Height          =   375
            Left            =   2160
            TabIndex        =   136
            Top             =   780
            Width           =   1875
         End
         Begin VB.TextBox txtAbonoMinimo 
            Height          =   375
            Left            =   2160
            TabIndex        =   134
            Top             =   360
            Width           =   1875
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   435
            Left            =   1260
            TabIndex        =   132
            Top             =   1380
            Width           =   1815
         End
         Begin VB.Label Label57 
            Caption         =   "Porcentaje de incremento : "
            Height          =   315
            Left            =   180
            TabIndex        =   135
            Top             =   840
            Width           =   1935
         End
         Begin VB.Label Label56 
            Caption         =   "Abono Minimo:"
            Height          =   255
            Left            =   180
            TabIndex        =   133
            Top             =   480
            Width           =   1155
         End
      End
      Begin VB.CommandButton cmdAceptarRecibo 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   -68340
         TabIndex        =   124
         Top             =   6420
         Width           =   1635
      End
      Begin VB.Frame Frame7 
         Caption         =   "Recibos"
         Height          =   1695
         Left            =   -74460
         TabIndex        =   99
         Top             =   780
         Width           =   11655
         Begin VB.TextBox txtRetencionesSUSS 
            Height          =   315
            Left            =   6540
            TabIndex        =   125
            Text            =   "0"
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtBanco 
            Height          =   315
            Left            =   9120
            TabIndex        =   122
            Top             =   1200
            Width           =   2355
         End
         Begin VB.TextBox txtReciboNumero 
            Height          =   315
            Left            =   4140
            TabIndex        =   121
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txtReciboValores 
            Height          =   315
            Left            =   1260
            TabIndex        =   120
            Text            =   "0"
            Top             =   780
            Width           =   1095
         End
         Begin VB.TextBox txtRetencionesIVA 
            Height          =   315
            Left            =   1260
            TabIndex        =   119
            Text            =   "0"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtRetencionesGanancias 
            Height          =   315
            Left            =   4140
            TabIndex        =   115
            Text            =   "0"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtRetencionesIngresosBrutos 
            Height          =   315
            Left            =   4140
            TabIndex        =   113
            Text            =   "0"
            Top             =   780
            Width           =   1095
         End
         Begin VB.TextBox txtReciboFecha 
            Height          =   315
            Left            =   6540
            TabIndex        =   106
            Top             =   300
            Width           =   1155
         End
         Begin VB.TextBox txtNumero_Respaldo 
            Height          =   315
            Left            =   9120
            TabIndex        =   104
            Top             =   720
            Width           =   2415
         End
         Begin VB.ComboBox cboTipoPago 
            Height          =   315
            Left            =   9120
            Style           =   2  'Dropdown List
            TabIndex        =   100
            Top             =   300
            Width           =   2415
         End
         Begin VB.Label lbl_ID_RECIBO 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   1260
            TabIndex        =   128
            Top             =   300
            Width           =   1155
         End
         Begin VB.Label Label55 
            Caption         =   "ID_ Recibo"
            Height          =   315
            Left            =   120
            TabIndex        =   127
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label54 
            Caption         =   "Ret. SUSS"
            Height          =   315
            Left            =   5460
            TabIndex        =   126
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label53 
            Caption         =   "Banco"
            Height          =   315
            Left            =   7980
            TabIndex        =   123
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label48 
            Caption         =   "Total"
            Height          =   315
            Left            =   5460
            TabIndex        =   118
            Top             =   1260
            Width           =   915
         End
         Begin VB.Label lblReciboTotal 
            BorderStyle     =   1  'Fixed Single
            Height          =   315
            Left            =   6540
            TabIndex        =   117
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label52 
            Caption         =   "Ret. Ganacias:"
            Height          =   315
            Left            =   2520
            TabIndex        =   116
            Top             =   1260
            Width           =   1395
         End
         Begin VB.Label Label51 
            Caption         =   "Ret. Ingresos Brutos:"
            Height          =   315
            Left            =   2520
            TabIndex        =   114
            Top             =   840
            Width           =   1515
         End
         Begin VB.Label Label50 
            Caption         =   "Ret. IVA"
            Height          =   315
            Left            =   120
            TabIndex        =   112
            Top             =   1200
            Width           =   675
         End
         Begin VB.Label Label49 
            Caption         =   "Fecha:"
            Height          =   315
            Left            =   5460
            TabIndex        =   107
            Top             =   360
            Width           =   555
         End
         Begin VB.Label lblNumero_Respaldo 
            Caption         =   "Tipo de Pago:"
            Height          =   315
            Left            =   7980
            TabIndex        =   105
            Top             =   780
            Width           =   1155
         End
         Begin VB.Label Label44 
            Caption         =   "Valores:"
            Height          =   315
            Left            =   120
            TabIndex        =   103
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label Label42 
            Caption         =   "Tipo de Pago:"
            Height          =   315
            Left            =   7980
            TabIndex        =   102
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label39 
            Caption         =   "Nº Recibo"
            Height          =   315
            Left            =   2580
            TabIndex        =   101
            Top             =   360
            Width           =   1155
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdReciboFacuta 
         Height          =   2175
         Left            =   -74520
         TabIndex        =   98
         Top             =   3840
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
         _Version        =   393216
      End
      Begin VB.Frame Frame6 
         Caption         =   "BuscarFactura"
         Height          =   1275
         Left            =   -74460
         TabIndex        =   86
         Top             =   2520
         Width           =   7755
         Begin VB.CommandButton cmdReciboFacturaGrilla 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   6660
            TabIndex        =   108
            Top             =   780
            Width           =   915
         End
         Begin VB.TextBox txtReciboTipoFactura 
            Height          =   375
            Left            =   1140
            MaxLength       =   1
            TabIndex        =   89
            Top             =   300
            Width           =   375
         End
         Begin VB.TextBox txtreciboNumeroFactura 
            Height          =   375
            Left            =   2520
            TabIndex        =   88
            Top             =   300
            Width           =   735
         End
         Begin VB.CommandButton cmdReciboBuscarFactura 
            Caption         =   "..."
            Height          =   375
            Left            =   3300
            TabIndex        =   87
            Top             =   300
            Width           =   315
         End
         Begin VB.Label lblCod_Cliente_Recibo 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1140
            TabIndex        =   130
            Top             =   780
            Width           =   675
         End
         Begin VB.Label Label47 
            Caption         =   "Tipo Factura"
            Height          =   315
            Left            =   60
            TabIndex        =   97
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label46 
            Caption         =   "Nº Factura"
            Height          =   315
            Left            =   1680
            TabIndex        =   96
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label45 
            Caption         =   "Monto"
            Height          =   315
            Left            =   3840
            TabIndex        =   95
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblReciboMontoFactura 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   4440
            TabIndex        =   94
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label43 
            Caption         =   "ID Factura"
            Height          =   315
            Left            =   5820
            TabIndex        =   93
            Top             =   300
            Width           =   855
         End
         Begin VB.Label lblReciboIDFactura 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   6660
            TabIndex        =   92
            Top             =   300
            Width           =   915
         End
         Begin VB.Label Label41 
            Caption         =   "Razon Social"
            Height          =   315
            Left            =   60
            TabIndex        =   91
            Top             =   840
            Width           =   975
         End
         Begin VB.Label lblReciboRazonSocial 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1860
            TabIndex        =   90
            Top             =   780
            Width           =   4755
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Recibos"
         Height          =   1875
         Left            =   -71040
         TabIndex        =   79
         Top             =   3480
         Width           =   8055
         Begin VB.CommandButton cmdReciboResponsable 
            Caption         =   "CrearRecibos"
            Height          =   375
            Left            =   1920
            TabIndex        =   111
            Top             =   1140
            Width           =   1635
         End
         Begin Controles.cltGenerico ctlPersonalRecibo 
            Height          =   375
            Left            =   1680
            TabIndex        =   84
            Top             =   180
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   661
         End
         Begin VB.TextBox txtReciboFin 
            Height          =   375
            Left            =   5940
            TabIndex        =   81
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox txtReciboInicio 
            Height          =   375
            Left            =   1680
            TabIndex        =   80
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label38 
            Caption         =   "Responsable"
            Height          =   315
            Left            =   180
            TabIndex        =   85
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label Label36 
            Caption         =   "Recibo Fin"
            Height          =   315
            Left            =   4860
            TabIndex        =   83
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label34 
            Caption         =   "Recibo Inicio"
            Height          =   315
            Left            =   180
            TabIndex        =   82
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkNoBorraGrilla 
         Caption         =   "No Borrar Grilla"
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
         Left            =   10560
         TabIndex        =   78
         Top             =   960
         Width           =   1875
      End
      Begin VB.Frame Frame4 
         Caption         =   "Anular Factura"
         Height          =   2475
         Left            =   -71040
         TabIndex        =   60
         Top             =   900
         Width           =   8295
         Begin VB.CommandButton cmdPendienteCobro 
            Caption         =   "Pendiente de Cobro"
            Height          =   435
            Left            =   2520
            TabIndex        =   129
            Top             =   1980
            Width           =   1335
         End
         Begin VB.CommandButton cmdEnviadaPorCorreo 
            Caption         =   "Envida por Correo"
            Height          =   435
            Left            =   3960
            TabIndex        =   76
            Top             =   1980
            Width           =   1335
         End
         Begin VB.CommandButton cmdFacturaEntregada 
            Caption         =   "Entregada por Basa"
            Height          =   435
            Left            =   5400
            TabIndex        =   75
            Top             =   1980
            Width           =   1335
         End
         Begin VB.CommandButton cmdAnular 
            Caption         =   "Anular"
            Height          =   435
            Left            =   6780
            TabIndex        =   74
            Top             =   1980
            Width           =   1335
         End
         Begin VB.TextBox txtDescripcionAnulacion 
            Height          =   735
            Left            =   1140
            TabIndex        =   72
            Top             =   1140
            Width           =   6915
         End
         Begin VB.CommandButton cmdAnularFacturaBuscar 
            Caption         =   "..."
            Height          =   375
            Left            =   3960
            TabIndex        =   65
            Top             =   240
            Width           =   315
         End
         Begin VB.TextBox txtFacturaAnuladaNumero 
            Height          =   375
            Left            =   3000
            TabIndex        =   64
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtTipo_Factura_Anulada 
            Height          =   375
            Left            =   1140
            MaxLength       =   1
            TabIndex        =   63
            Top             =   300
            Width           =   375
         End
         Begin VB.Label Label37 
            Caption         =   "Descripcion"
            Height          =   315
            Left            =   60
            TabIndex        =   73
            Top             =   1380
            Width           =   975
         End
         Begin VB.Label lblRazonSocialAnulada 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   1140
            TabIndex        =   71
            Top             =   720
            Width           =   6915
         End
         Begin VB.Label Label35 
            Caption         =   "Razon Social"
            Height          =   315
            Left            =   60
            TabIndex        =   70
            Top             =   780
            Width           =   975
         End
         Begin VB.Label ID_Factura_Anulada 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   7200
            TabIndex        =   69
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label33 
            Caption         =   "ID Factura"
            Height          =   315
            Left            =   6360
            TabIndex        =   68
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lblMontoFactura 
            BorderStyle     =   1  'Fixed Single
            Height          =   375
            Left            =   5220
            TabIndex        =   67
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label32 
            Caption         =   "Monto"
            Height          =   315
            Left            =   4680
            TabIndex        =   66
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label31 
            Caption         =   "Numero Factura"
            Height          =   315
            Left            =   1740
            TabIndex        =   62
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label30 
            Caption         =   "Tipo Factura"
            Height          =   315
            Left            =   60
            TabIndex        =   61
            Top             =   360
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Parametros"
         Height          =   2475
         Left            =   -74820
         TabIndex        =   53
         Top             =   900
         Width           =   3615
         Begin VB.CommandButton cmdActualizarParametroFactura 
            Caption         =   "Actualizar"
            Height          =   375
            Left            =   2100
            TabIndex        =   77
            Top             =   1620
            Width           =   1335
         End
         Begin VB.TextBox txtMesServicio 
            Height          =   375
            Left            =   1500
            TabIndex        =   59
            Top             =   300
            Width           =   1935
         End
         Begin VB.TextBox txtFecha_Hasta 
            Height          =   375
            Left            =   1500
            TabIndex        =   57
            Top             =   1140
            Width           =   1935
         End
         Begin VB.TextBox txtFecha_Desde 
            Height          =   375
            Left            =   1500
            TabIndex        =   55
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label29 
            Caption         =   "Mes Servicio"
            Height          =   315
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label28 
            Caption         =   "Fecha Hasta"
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   56
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Fecha Desde"
            Height          =   315
            Left            =   120
            TabIndex        =   54
            Top             =   780
            Width           =   1095
         End
      End
      Begin VB.TextBox txtMesFacturacion 
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
         Left            =   10560
         TabIndex        =   52
         Text            =   "200711"
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Información Factura"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   240
         TabIndex        =   46
         Top             =   3720
         Width           =   12615
         Begin VB.TextBox txtPeriodoActual 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   5940
            TabIndex        =   50
            Top             =   780
            Width           =   2175
         End
         Begin VB.TextBox txtDetalleFacturacion 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   47
            Top             =   300
            Width           =   8535
         End
         Begin VB.Label Label26 
            Caption         =   "Periodo Actual"
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
            Left            =   4560
            TabIndex        =   51
            Top             =   900
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   " Periodo Anterior Facturado"
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
            Left            =   9480
            TabIndex        =   49
            Top             =   180
            Width           =   2415
         End
         Begin VB.Label lblPeriodoAnteriorFacturado 
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
            Height          =   435
            Left            =   8820
            TabIndex        =   48
            Top             =   480
            Width           =   3555
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   16
         Top             =   1740
         Width           =   12675
         Begin VB.CheckBox chkCopiarImagenes 
            Caption         =   "Copiar Imagenes"
            Height          =   315
            Left            =   1380
            TabIndex        =   176
            Top             =   1500
            Width           =   1935
         End
         Begin VB.CommandButton cmdImprimirInforme 
            Caption         =   "Imprimir Informe"
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
            Left            =   10680
            TabIndex        =   175
            Top             =   1500
            Width           =   1695
         End
         Begin VB.CommandButton cmdExportExcel 
            Caption         =   "Export Excel"
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
            Left            =   8820
            TabIndex        =   174
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton cmdCuponesDisco 
            Caption         =   "..."
            Height          =   315
            Left            =   7080
            TabIndex        =   168
            Top             =   960
            Width           =   315
         End
         Begin VB.CommandButton cmdFacturacionMensual 
            Caption         =   "Planilla General"
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
            Left            =   10680
            TabIndex        =   167
            Top             =   1080
            Width           =   1695
         End
         Begin VB.CommandButton cmdCargaLegajos 
            BackColor       =   &H80000013&
            Caption         =   "..."
            Height          =   315
            Left            =   4560
            MaskColor       =   &H80000013&
            Style           =   1  'Graphical
            TabIndex        =   140
            Top             =   1020
            Width           =   315
         End
         Begin VB.CommandButton cmdRearchivo 
            BackColor       =   &H80000013&
            Caption         =   "..."
            Height          =   315
            Left            =   2160
            MaskColor       =   &H00C0C000&
            Style           =   1  'Graphical
            TabIndex        =   137
            Top             =   1020
            Width           =   315
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Command3"
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
            Left            =   12120
            TabIndex        =   42
            Top             =   660
            Width           =   375
         End
         Begin VB.CommandButton cmdCantidadConsultasDetalle 
            Caption         =   "..."
            Height          =   315
            Left            =   4560
            TabIndex        =   37
            Top             =   660
            Width           =   315
         End
         Begin VB.CommandButton cmdCantidadCajaDetalle 
            Caption         =   "..."
            Height          =   315
            Left            =   2160
            TabIndex        =   23
            Top             =   300
            Width           =   315
         End
         Begin VB.CommandButton cmdCantidadLibrosDetalle 
            Caption         =   "..."
            Height          =   315
            Left            =   4560
            TabIndex        =   22
            Top             =   300
            Width           =   315
         End
         Begin VB.CommandButton cmsCantidadCajasVacias 
            Caption         =   "..."
            Height          =   315
            Left            =   2160
            TabIndex        =   21
            Top             =   660
            Width           =   315
         End
         Begin VB.CommandButton cmdCantidadFletesNormales 
            Caption         =   "..."
            Height          =   315
            Left            =   7080
            TabIndex        =   20
            Top             =   600
            Width           =   315
         End
         Begin VB.CommandButton cmdCantidadFletesUrgentes 
            Caption         =   "..."
            Height          =   315
            Left            =   9420
            TabIndex        =   19
            Top             =   660
            Width           =   315
         End
         Begin VB.CommandButton cmdCantidadLegajos 
            BackColor       =   &H80000013&
            Caption         =   "..."
            Height          =   315
            Left            =   7080
            MaskColor       =   &H80000013&
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   315
         End
         Begin VB.CommandButton cmdCantidadCajasCrecimientoMes 
            BackColor       =   &H80000013&
            Caption         =   "..."
            Height          =   315
            Left            =   9420
            MaskColor       =   &H80000013&
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   300
            Width           =   315
         End
         Begin MSMask.MaskEdBox mskFecha_Desde 
            Height          =   315
            Left            =   10980
            TabIndex        =   40
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
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
         Begin MSMask.MaskEdBox mskFecha_Hasta 
            Height          =   315
            Left            =   10980
            TabIndex        =   41
            Top             =   660
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
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
         Begin VB.Label lblCuponesDisco 
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
            Height          =   315
            Left            =   6420
            TabIndex        =   170
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label58 
            Caption         =   "Cupones Disco"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4980
            TabIndex        =   169
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label60 
            Caption         =   "Car Legajos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2580
            TabIndex        =   142
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblCargaLegajo 
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
            Height          =   315
            Left            =   3780
            TabIndex        =   141
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label Label59 
            Caption         =   "Rearchivo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   139
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblRearchivo 
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
            Height          =   315
            Left            =   1380
            TabIndex        =   138
            Top             =   1020
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Cant. Cajas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label20 
            Caption         =   "Fecha Fin:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9840
            TabIndex        =   44
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "Fecha Inicio:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9840
            TabIndex        =   43
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label21 
            Caption         =   " Fletes Normales"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   39
            Top             =   720
            Width           =   1395
         End
         Begin VB.Label Label18 
            Caption         =   "Cant. de Legajos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4980
            TabIndex        =   38
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label lblCantidad_Cajas 
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
            Height          =   315
            Left            =   1380
            TabIndex        =   36
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Cant. Libros:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2580
            TabIndex        =   35
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblCantidad_Libros 
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
            Height          =   315
            Left            =   3780
            TabIndex        =   34
            Top             =   300
            Width           =   735
         End
         Begin VB.Label lblCantidadCajasVacias 
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
            Height          =   315
            Left            =   1380
            TabIndex        =   33
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "Cajas Vacias:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblCantidadDesarchivos 
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
            Height          =   315
            Left            =   3780
            TabIndex        =   31
            Top             =   660
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   " Consultas:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2520
            TabIndex        =   30
            Top             =   720
            Width           =   915
         End
         Begin VB.Label lbl_FletesNormales 
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
            Height          =   315
            Left            =   6420
            TabIndex        =   29
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lbl_FletesUrgentes 
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
            Height          =   315
            Left            =   8820
            TabIndex        =   28
            Top             =   660
            Width           =   555
         End
         Begin VB.Label Label23 
            Caption         =   " Fletes Urgentes"
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
            Left            =   7380
            TabIndex        =   27
            Top             =   660
            Width           =   1395
         End
         Begin VB.Label lblCantidad_Legajos 
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
            Height          =   315
            Left            =   6420
            TabIndex        =   26
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "Cajas Crec."
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
            Left            =   7560
            TabIndex        =   25
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblCajasCrecimientoMes 
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
            Height          =   315
            Left            =   8820
            TabIndex        =   24
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.CheckBox chkSoloPendientes 
         Alignment       =   1  'Right Justify
         Caption         =   "Solo los pendiente"
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
         Left            =   -71940
         TabIndex        =   15
         Top             =   900
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txtFechaFlete 
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
         Left            =   -73680
         TabIndex        =   13
         Top             =   840
         Width           =   1635
      End
      Begin VB.CommandButton cmdActualizarFletes 
         Caption         =   "Actualizar"
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
         Left            =   -69840
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid grdFletes 
         Height          =   6915
         Left            =   -74880
         TabIndex        =   11
         Top             =   1380
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   12197
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   16
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
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
      Begin VB.TextBox txtFechaFactura 
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
         Left            =   8460
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtNumeroFactura 
         Height          =   315
         Left            =   6300
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cboTipoComprobante 
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
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label62 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   -72480
         TabIndex        =   183
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label61 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   180
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label40 
         Caption         =   "Total Facturas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69780
         TabIndex        =   110
         Top             =   6060
         Width           =   1395
      End
      Begin VB.Label lblReciboTotalFacturas 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -68280
         TabIndex        =   109
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label Label25 
         Caption         =   "Fecha Flete:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74880
         TabIndex        =   14
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
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
         Left            =   7800
         TabIndex        =   9
         Top             =   1500
         Width           =   555
      End
      Begin VB.Label Label24 
         Caption         =   "Nº Factura:"
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
         Left            =   5400
         TabIndex        =   7
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "Tipo de Comprobante"
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
         Left            =   120
         TabIndex        =   5
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Tipo Factura:"
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
         Left            =   3840
         TabIndex        =   4
         Top             =   1500
         Width           =   795
      End
      Begin VB.Label lblTipoFactura 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4620
         TabIndex        =   3
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label15 
         Caption         =   " "
         Height          =   315
         Left            =   4920
         TabIndex        =   2
         Top             =   840
         Width           =   1035
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
         Left            =   180
         TabIndex        =   1
         Top             =   1020
         Width           =   675
      End
   End
   Begin VB.Menu mnuGrilla 
      Caption         =   "Grilla"
      Visible         =   0   'False
      Begin VB.Menu mnuGrillaModificar 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnuGrillaBorrar 
         Caption         =   "Borrar"
      End
      Begin VB.Menu mnuBorrarTodo 
         Caption         =   "Borrar Todo"
      End
   End
   Begin VB.Menu mnuFlete 
      Caption         =   "Fletes"
      Visible         =   0   'False
      Begin VB.Menu mnuNuevoFlete 
         Caption         =   "Nuevo Flete"
      End
      Begin VB.Menu mnuFleteUnificado 
         Caption         =   "Flete Unificado"
      End
      Begin VB.Menu mnuFleteSinCosto 
         Caption         =   "Flete Sin Costo"
      End
   End
   Begin VB.Menu mnuInformes 
      Caption         =   "Informes"
      Begin VB.Menu mnuFacturacion 
         Caption         =   "Facturacion"
         Begin VB.Menu mnuFacturasCobrar 
            Caption         =   "Factura a Cobra"
         End
         Begin VB.Menu mnuControlfacturacion 
            Caption         =   "Control Facturacion"
         End
         Begin VB.Menu mnuInformePorCliente 
            Caption         =   "Informe Por Cliente"
         End
         Begin VB.Menu mnuInformeFacturacion 
            Caption         =   "Informe facturacion"
            Begin VB.Menu mnuOrdenadoFactura 
               Caption         =   "Ordenado por Factura"
            End
            Begin VB.Menu mnuOrdenadoPorCliente 
               Caption         =   "Ordenado por Cliente"
            End
         End
      End
      Begin VB.Menu mnuRetenciones 
         Caption         =   "Retenciones"
         Begin VB.Menu mnuControlRetenciones 
            Caption         =   "Control Retenciones"
         End
      End
      Begin VB.Menu mnuRecibos 
         Caption         =   "Recibos"
         Begin VB.Menu mnuListadoRecibos 
            Caption         =   "Listado de recibos"
         End
         Begin VB.Menu mnuListadoReciboPorFechaCarga 
            Caption         =   "Listado de recibos por fecha de Carga"
         End
         Begin VB.Menu mnuReciboRendicion 
            Caption         =   "Rendición"
         End
      End
      Begin VB.Menu mnuCobranzas 
         Caption         =   "Cobranzas"
         Begin VB.Menu mnuListadoCobranzas 
            Caption         =   "Listado de  Cobranzas"
         End
         Begin VB.Menu mnuCompromisoPago 
            Caption         =   "Compromiso de pago"
         End
      End
   End
   Begin VB.Menu mnuGrillarecibos 
      Caption         =   "Grilla_Recibos"
      Visible         =   0   'False
      Begin VB.Menu mnuReciboBorrarTodo 
         Caption         =   "Borrar Todo"
      End
   End
End
Attribute VB_Name = "frmfacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RowGrilla As Integer
Dim RemitoFlete As Long
Dim RowBo As Long
Dim AbonoMinimo As Long
Dim COD_CLIENTE As Integer
Dim FECHA_INICIO As String
Dim FECHA_FIN As String
Dim MES_SERVICIO As Integer
Dim PasoFacturaImagenes As String
Dim RsFactura As New ADODB.Recordset
Enum Tipo_Orden
    FISICO = 1
    lote = 2

    
End Enum
Private Sub cboTipoPago_Click()
    lblNumero_Respaldo.Visible = False
    txtNumero_Respaldo.Visible = False
    If cboTipoPago.Text <> "CONTADO" Then
        lblNumero_Respaldo.Caption = "Numero:"
        lblNumero_Respaldo.Visible = True
        txtNumero_Respaldo.Visible = True
    End If
End Sub

Private Sub cmdAceptarRecibo_Click()
    Dim Sql As String
    Dim i As Integer
    Dim TIPO_PAGO, fecha, NUMERO_RESPALDO, BANCO    As String
    Dim MONTO_TOTAL, ESTADO_RECIBO, RETENCION_GANANCIAS As String
    Dim RETENCION_IVA, RENTENCION_INGRESOS_BRUTOS, RENTENCION_SUSS   As String




If lblReciboTotalFacturas.Caption = lblReciboTotal.Caption Then
Else
    MsgBox "No concuerdan los Montos", vbInformation
    Exit Sub
End If

    If cboTipoPago.Text <> "" Then
       TIPO_PAGO = "'" & Trim(cboTipoPago.Text) & "'"
    Else
       MsgBox "Falta el tipo de pago"
       Exit Sub
    End If
    
    
    If txtReciboFecha.Text <> "" Then
       fecha = "'" & Trim(txtReciboFecha.Text) & "'"
    Else
       MsgBox "Falta la fecha"
       Exit Sub
    End If
    
    If txtNumero_Respaldo.Text <> "" Then
      NUMERO_RESPALDO = "'" & txtNumero_Respaldo.Text & "'"
    Else
      NUMERO_RESPALDO = "Null"
    End If
      
     If txtBanco.Text <> "" Then
        BANCO = "'" & Trim(txtBanco.Text) & "'"
     Else
        BANCO = "Null"
    End If
             
     If lblReciboTotal.Caption <> "" Then
        MONTO_TOTAL = "'" & Trim(lblReciboTotal.Caption) & "'"
     Else
        MONTO_TOTAL = "Null"
        MsgBox "Monto del recibo"
        Exit Sub
     End If
     ESTADO_RECIBO = "100"
     
     If Trim(txtRetencionesGanancias.Text) <> "" Then
        RETENCION_GANANCIAS = "'" & Trim(txtRetencionesGanancias.Text) & "'"
     Else
        RETENCION_GANANCIAS = "Null"
     End If
     
     If Trim(txtRetencionesIVA.Text) <> "" Then
        RETENCION_IVA = "'" & Trim(txtRetencionesIVA.Text) & "'"
     Else
        RETENCION_IVA = "Null"
     End If
     
     If Trim(txtRetencionesIngresosBrutos.Text) <> "" Then
        RENTENCION_INGRESOS_BRUTOS = "'" & Trim(txtRetencionesIngresosBrutos.Text) & "'"
     Else
        RENTENCION_INGRESOS_BRUTOS = "Null"
     End If
      
      If Trim(txtRetencionesSUSS.Text) <> "" Then
         RENTENCION_SUSS = "'" & Trim(txtRetencionesSUSS.Text) & "'"
      Else
        RENTENCION_SUSS = "Null"
      End If
      
        If lblCod_Cliente_Recibo.Caption = "" Then
            MsgBox "Falta el codigo del cliente"
            Exit Sub
        End If
        
    ConBasa.BeginTrans
On Error GoTo salir:

Sql = " Update RECIBOS "
Sql = Sql & vbCrLf & " SET TIPO_PAGO =" & TIPO_PAGO
Sql = Sql & vbCrLf & " , FECHA =" & fecha
Sql = Sql & vbCrLf & " , NUMERO_RESPALDO =" & NUMERO_RESPALDO
Sql = Sql & vbCrLf & " , BANCO =" & BANCO
Sql = Sql & vbCrLf & " , MONTO_TOTAL =" & MONTO_TOTAL
Sql = Sql & vbCrLf & " , ESTADO_RECIBO =" & ESTADO_RECIBO
Sql = Sql & vbCrLf & " , RETENCION_GANANCIAS =" & RETENCION_GANANCIAS
Sql = Sql & vbCrLf & " , RETENCION_IVA =" & RETENCION_IVA
Sql = Sql & vbCrLf & " , RENTENCION_INGRESOS_BRUTOS =" & RENTENCION_INGRESOS_BRUTOS
Sql = Sql & vbCrLf & " , RENTENCION_SUSS =" & RENTENCION_SUSS
Sql = Sql & vbCrLf & " , FECHA_CARGA= " & SysDate
Sql = Sql & vbCrLf & " , COD_CLIENTE= " & lblCod_Cliente_Recibo.Caption
Sql = Sql & vbCrLf & "  Where ID_Recibo = " & lbl_ID_RECIBO.Caption

ExecutarSql Sql

         For i = 1 To grdReciboFacuta.Rows - 1
                Sql = " Update FACTURAS "
                Sql = Sql & " SET COD_RECIBO =" & lbl_ID_RECIBO.Caption
                Sql = Sql & " , ESTADO =100 "
                Sql = Sql & "  Where ID_FACTURA = " & grdReciboFacuta.TextMatrix(i, 1)
                 ExecutarSql Sql
        Next
        ConBasa.CommitTrans
        lblReciboIDFactura.Caption = ""
    txtReciboTipoFactura.Text = ""
    txtreciboNumeroFactura.Text = ""
    lblReciboMontoFactura.Caption = ""
    lblReciboRazonSocial.Caption = ""
        mnuReciboBorrarTodo_Click
        
            txtReciboNumero.Text = ""
            cboTipoPago.ListIndex = -1
            txtReciboFecha.Text = ""
            txtNumero_Respaldo.Text = ""
            txtBanco.Text = ""
            txtReciboValores.Text = "0"
            txtRetencionesGanancias.Text = "0"
            txtRetencionesIVA.Text = "0"
            txtRetencionesIngresosBrutos.Text = "0"
            txtRetencionesSUSS.Text = "0"
            lblReciboTotal.Caption = ""
            lblReciboTotalFacturas.Caption = ""
            
        MsgBox "La Acutalizacion se realizo con exito" & vbCrLf & "NUMERO " & lbl_ID_RECIBO, vbInformation
        lbl_ID_RECIBO.Caption = ""
        
       
Exit Sub
salir:
 ConBasa.RollbackTrans
 MsgBox "Error En la actualizacion", vbCritical

End Sub

Private Sub cmdActualizarCliente_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim ConClienteElecronica As New ADODB.Connection
    Dim conData As New ADODB.Connection
    conData.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=datas"
        
    
    
    
        ConClienteElecronica.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=False;User ID=sa;Initial Catalog=factura_electronica;Data Source=222.15.19.150"
        
        ConClienteElecronica.Execute " DELETE FROM CLIENTE "
        
        Sql = " SELECT IDCLIENTE, NOMBRE ,DOMICILIO ,NUMEROCALLE"
        Sql = Sql & " ,PROVINCIA ,IVA ,CUIT , FACTURAR, LOCALIDAD"
        Sql = Sql & " From Cliente "
        rs.Open Sql, conData, adOpenForwardOnly, adLockReadOnly
        
        
        
  
  Do While Not rs.EOF
   Rem MsgBox Rs!Nombre
   
   Sql = " Insert Into factura_electronica.dbo.Cliente("
   Sql = Sql & " IDCliente"
   Sql = Sql & " , Nombre"
   Sql = Sql & " , DOMICILIO"
   Sql = Sql & " , NUMEROCALLE"
   Sql = Sql & " , PROVINCIA"
   Sql = Sql & " , IVA"
   Sql = Sql & " , Cuit"
   Sql = Sql & " , FACTURAR"
   Sql = Sql & " , Localidad"
   Sql = Sql & " )"
   Sql = Sql & "  VALUES ("
    Sql = Sql & rs!IDCliente
   Sql = Sql & " , '" & Trim(rs!Nombre) & "'"
   
   If Len(rs!DOMICILIO) < 2 Then
   Sql = Sql & " , NULL"
   Else
    Sql = Sql & " , '" & Replace(Trim(rs!DOMICILIO), "'", "´") & "'"
   End If
   
   Sql = Sql & " , '" & Trim(rs!NUMEROCALLE) & "'"
   If Len(rs!PROVINCIA) = 0 Then
   Sql = Sql & " , NULL"
   Else
   Sql = Sql & " , '" & Trim(rs!PROVINCIA) & "'"
   End If
   Sql = Sql & " , '" & Trim(rs!IVA) & "'"
   Sql = Sql & " , '" & Trim(rs!Cuit) & "'"
   Sql = Sql & " , '" & Trim(rs!FACTURAR) & "'"
   If Len(rs!Localidad) = 0 Then
     Sql = Sql & " , NULL"
   Else
     Sql = Sql & " , '" & Trim(rs!Localidad) & "'"
   End If
   
   
 
   
   
   Sql = Sql & " )"
   ConClienteElecronica.Execute Sql
    
    rs.MoveNext
  Loop
  
  MsgBox "terminado"
  
End Sub

Private Sub cmdActualizarFletes_Click()
    CargarFletes
End Sub

Private Sub cmdActualizarParametroFactura_Click()
    Dim Sql As String
    Sql = " Update FACTURA_PARAMETROS"
    Sql = Sql & " SET FECHA_DESDE = '" & txtFecha_Desde.Text & "'"
    Sql = Sql & " , FECHA_HASTA = '" & txtFecha_Hasta.Text & "'"
    Sql = Sql & " , MESFACTURACION =" & txtMesServicio.Text
    ExecutarSql Sql
    Unload Me
End Sub

Private Sub cmdAnular_Click()
    
    Dim Sql As String
    Dim Control As Integer
    If Trim(txtDescripcionAnulacion.Text) = "" Then
        MsgBox "No existe descripcion " & vbCrLf & "NO se realiso la operacion"
        Exit Sub
    End If
    
        Sql = " Update FACTURAS"
        Sql = Sql & vbCrLf & " SET ESTADO = 0,"
        Sql = Sql & vbCrLf & "  DESCRIPCION_ANULADO ='" & txtDescripcionAnulacion.Text & "'"
        Sql = Sql & vbCrLf & " Where ID_FACTURA = " & ID_Factura_Anulada.Caption
        Sql = Sql & vbCrLf & " AND ESTADO IN(10,20,30,40 ) "
       Control = ExecutarSql(Sql)
    If Control = 0 Then
        MsgBox "La Actualizacion NO se realizo con exito", vbCritical
    Else
        MsgBox "La Actualizacion se realizo con exito", vbInformation
    End If
         txtDescripcionAnulacion.Text = ""
         ID_Factura_Anulada.Caption = ""
         lblRazonSocialAnulada.Caption = ""

End Sub

Private Sub cmdAnularFacturaBuscar_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
    Sql = " SELECT FACTURAS.ID_FACTURA, FACTURAS.TIPO_FACTURA, "
    Sql = Sql & vbCrLf & " FACTURAS.NUMERO_FACTURA, FACTURAS.MONTO_CON_IVA, "
    Sql = Sql & vbCrLf & " Clientes.RAZON_SOCIAL, Clientes.id_cliente "
    Sql = Sql & vbCrLf & "  From FACTURAS, Clientes "
    Sql = Sql & vbCrLf & "  Where FACTURAS.COD_CLIENTE = Clientes.id_cliente"
    Sql = Sql & vbCrLf & "  AND FACTURAS.TIPO_FACTURA = '" & UCase(txtTipo_Factura_Anulada.Text) & "' "
    Sql = Sql & vbCrLf & "  AND FACTURAS.NUMERO_FACTURA =" & txtFacturaAnuladaNumero.Text
    rs.Open Sql, ConActiva, 0, 1

If Not rs.EOF Then
    lblMontoFactura.Caption = rs!MONTO_CON_IVA
    ID_Factura_Anulada.Caption = rs!ID_FACTURA
    lblRazonSocialAnulada.Caption = rs!id_cliente & " \ " & rs!RAZON_SOCIAL
Else

End If



End Sub

Private Sub cmdCantidadCajaDetalle_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim DATO As String

     
     
    Sql = " SELECT NRO_REMITO, NRO_REM_PROV, FECHA, CANTIDAD, TIPO "
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " Where id_cliente =" & ctlClienteFactura.Valor
    Sql = Sql & vbCrLf & " And TIPO in( 0 ,3)"
    Sql = Sql & vbCrLf & " And ANULADO Is Null"
    Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
     Sql = Sql & vbCrLf & " AND FECHA >  " & FechaFormato(mskFecha_Desde.Text)
      Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(mskFecha_Hasta.Text)
    Sql = Sql & vbCrLf & " Order by NRO_REMITO"

    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    DATO = ctlClienteFactura.Descripcion & vbCrLf
     DATO = DATO & " Crecimiento Mensual de cajas " & vbCrLf & " Desde el : " & mskFecha_Desde.Text & " hasta " & mskFecha_Hasta.Text & vbCrLf
     DATO = DATO & " Remito " & vbTab & " Remito PROV " & vbTab & "Fecha" & vbTab & " Cantidad " & vbCrLf
    Do While Not rs.EOF
            If rs!TIPO = 0 Then
                DATO = DATO & rs!NRO_REMITO & vbTab & Trim(rs!NRO_REM_PROV) & vbTab & rs!fecha & vbTab & rs!cantidad & vbCrLf
             Else
                DATO = DATO & rs!NRO_REMITO & vbTab & Trim(rs!NRO_REM_PROV) & vbTab & rs!fecha & vbTab & rs!cantidad * -1 & vbCrLf
            End If
        
        rs.MoveNext
    Loop
    
    DATO = DATO & " Cajas sumas resta " & CajasSumaResta(ctlClienteFactura.Valor)
    
    Clipboard.Clear
    Clipboard.SetText DATO
    
    MsgBox "Los datos fueron Copiados"
    
            Exit Sub
salir:

MsgBox Err.Description
End Sub

Private Sub cmdCantidadCajasCrecimientoMes_Click()
FECHA_INICIO = mskFecha_Desde.Text
FECHA_FIN = mskFecha_Hasta.Text
CAJAS_CRECIMIENTO_MES_andre ctlClienteFactura.Valor, ""
MsgBox "Copiado"

End Sub

Private Sub cmdCantidadConsultasDetalle_Click()

    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim DirectorioImagenes As String
    Dim EstaRemito As String
    DirectorioImagenes = InputBox("Ingrese el directorio de los remitos    ")
    
    If Dir("c:/" & DirectorioImagenes, vbDirectory) = "" Then
         MkDir "c:/" & DirectorioImagenes
    End If
    
    

    
   Sql = "  SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.IMAGEN, REMITOS_CUERPO.CANTIDAD,"
   Sql = Sql & vbCrLf & "                    CLIENTEUSUARIO.APELLIDO_NOMBRE"
Sql = Sql & vbCrLf & " FROM         REMITOS_CUERPO INNER JOIN"
Sql = Sql & vbCrLf & "                      CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
Sql = Sql & vbCrLf & " WHERE     (REMITOS_CUERPO.ID_CLIENTE = " & ctlClienteFactura.Valor & ") AND (REMITOS_CUERPO.TIPO IN (1)) AND (REMITOS_CUERPO.ANULADO IS NULL) AND"
 Sql = Sql & vbCrLf & "   FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
  

 Sql = Sql & vbCrLf & "         AND              (REMITOS_CUERPO.OPERACION = 1)"
Sql = Sql & vbCrLf & " ORDER BY REMITOS_CUERPO.FECHA DESC"


    
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    DATO = ctlClienteFactura.Descripcion & vbCrLf
    DATO = DATO & " Consultas  " & vbCrLf & " Desde el : " & mskFecha_Desde.Text & " hasta " & mskFecha_Hasta.Text & vbCrLf
    DATO = DATO & vbTab & "  Remito " & vbTab & " Cantidad " & vbTab & " Fecha " & vbCrLf
    Dim NombreArchivo As String
    
    Do While Not rs.EOF
    
    If Dir("c:/" & DirectorioImagenes, vbDirectory) = "" Then
        MkDir ("c:/" & DirectorioImagenes)
   
    End If
   
    If IsNull(rs!Imagen) Then
      EstaRemito = "No existe la imagem del remito "
      MsgBox "No existe imagen " & rs!NRO_REMITO
      Else
      EstaRemito = ""
      
      NombreArchivo = Dir("Z:\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REMITO & "\*.tif")
      If NombreArchivo <> "" Then
      FileSystem.FileCopy "Z:\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REMITO & "\" & NombreArchivo, "c:/" & DirectorioImagenes & "/" & rs!NRO_REMITO & ".tif"
     Else
        MsgBox "No existe imagen " & rs!NRO_REMITO
     End If
     End If
         DATO = DATO & vbTab & rs!NRO_REMITO & vbTab & rs!cantidad & vbTab & rs!fecha & vbTab & rs!APELLIDO_NOMBRE & vbTab & "Consulta" & vbTab & EstaRemito & vbCrLf
        rs.MoveNext
    Loop
    
    
     Sql = "  SELECT  REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA,IMAGEN , REMITOS_CUERPO.CANTIDAD, CLIENTEUSUARIO.APELLIDO_NOMBRE "
    Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO INNER JOIN"
     Sql = Sql & vbCrLf & "  CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
    Sql = Sql & vbCrLf & "  Where OPERACION = 1 and  TIPO = 3 "
    Sql = Sql & vbCrLf & " and  REMITOS_CUERPO.ID_CLIENTE = " & ctlClienteFactura.Valor
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ANULADO IS NULL "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
    
    
   

    Sql = Sql & vbCrLf & " ORDER BY FECHA DESC "
     Dim NombreArchivo2 As String
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    Do While Not rs.EOF
     If IsNull(rs!Imagen) Then
      MsgBox "No existe la imagem del remito" & rs!NRO_REMITO
      Else
       If Dir("Z:\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REMITO & "\" & rs!Imagen) = "" Then
        MsgBox " Falta el remito " & rs!NRO_REMITO
       Else
      
      NombreArchivo2 = Dir("Z:\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REMITO & "\*.tif")
             
             
             If NombreArchivo2 <> "" Then
                    FileSystem.FileCopy "Z:\Administracion\Imagenes_Internas\Remitos\" & rs!NRO_REMITO & "\" & NombreArchivo2, "c:/" & DirectorioImagenes & "/" & rs!NRO_REMITO & ".tif"
             Else
             
             MsgBox " Falta el remito " & rs!NRO_REMITO
             End If
             
        
        
        End If
        
     End If
     
        DATO = DATO & vbTab & rs!NRO_REMITO & vbTab & rs!cantidad & vbTab & rs!fecha & vbTab & rs!APELLIDO_NOMBRE & vbTab & "Baja" & vbCrLf
        rs.MoveNext
    Loop
    
    
    
    Clipboard.Clear
    Clipboard.SetText DATO
    
    MsgBox "Los datos fueron Copiados"
    
End Sub

Private Sub cmdCantidadFletesNormales_Click()

    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    
    
    Sql = " SELECT REMITOS_CUERPO.COD_FLETE, REMITOS_CUERPO.NRO_REMITO,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ESTADO,"
    Sql = Sql & vbCrLf & "  REMITO_ESTADOS.DESCRIPCION,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.CANTIDAD,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.FECHA,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ANULADO"
    Sql = Sql & vbCrLf & "  From REMITOS_CUERPO, REMITO_ESTADOS"
    Sql = Sql & vbCrLf & "  WHERE REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ID_CLIENTE = " & ctlClienteFactura.Valor
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ANULADO IS NULL "
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ESTADO = 0 "
    Sql = Sql & vbCrLf & "  AND NOT COD_FLETE is null  "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
    Sql = Sql & vbCrLf & " ORDER BY FECHA DESC "
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    DATO = ctlClienteFactura.Descripcion & vbCrLf
    DATO = DATO & " Fletes Normales " & vbCrLf & " Desde el : " & mskFecha_Desde.Text & " hasta " & mskFecha_Hasta.Text & vbCrLf
    DATO = DATO & "Flete" & vbTab & "  Remito " & vbTab & " Estado " & vbTab & " Cantidad " & vbTab & " Fecha " & vbCrLf
    Do While Not rs.EOF
        DATO = DATO & rs!Cod_Flete & vbTab & rs!NRO_REMITO & vbTab & rs!Descripcion & vbTab & rs!cantidad & vbTab & rs!fecha & vbCrLf
        rs.MoveNext
    Loop
    Clipboard.Clear
    Clipboard.SetText DATO
    
    MsgBox "Los datos fueron Coipiados"
    
     
     
     

End Sub

Private Sub cmdCantidadFletesUrgentes_Click()
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Sql = " SELECT REMITOS_CUERPO.COD_FLETE, REMITOS_CUERPO.NRO_REMITO,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ESTADO,"
    Sql = Sql & vbCrLf & "  REMITO_ESTADOS.DESCRIPCION,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.CANTIDAD,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.FECHA,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ANULADO"
    Sql = Sql & vbCrLf & "  From REMITOS_CUERPO, REMITO_ESTADOS"
    Sql = Sql & vbCrLf & "  WHERE REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ID_CLIENTE = " & ctlClienteFactura.Valor
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ANULADO IS NULL "
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ESTADO = 1 "
    Sql = Sql & vbCrLf & "  AND NOT COD_FLETE is null  "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
    Sql = Sql & vbCrLf & " ORDER BY FECHA DESC "
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    DATO = ctlClienteFactura.Descripcion & vbCrLf
    DATO = DATO & " Fletes Normales " & vbCrLf & " Desde el : " & mskFecha_Desde.Text & " hasta " & mskFecha_Hasta.Text & vbCrLf
    DATO = DATO & "Flete" & vbTab & "  Remito " & vbTab & " Estado " & vbTab & " Cantidad " & vbTab & " Fecha " & vbCrLf
    Do While Not rs.EOF
        DATO = DATO & rs!Cod_Flete & vbTab & rs!NRO_REMITO & vbTab & rs!Descripcion & vbTab & rs!cantidad & vbTab & rs!fecha & vbCrLf
        rs.MoveNext
    Loop
    Clipboard.Clear
    Clipboard.SetText DATO
    
    MsgBox "Los datos fueron Coipiados"
    
End Sub


Private Sub cmdCantidadLegajos_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo salir:

        Sql = " SELECT COUNT(*) as Cantidad_Legajos"
        Sql = Sql & vbCrLf & " From LEGAJOS"
        Sql = Sql & vbCrLf & "  WHERE COD_CLIENTE = " & ctlClienteFactura.Valor
        Sql = Sql & vbCrLf & " AND NOT ( NRO_CAJA IS NULL)"
        rs.Open Sql, ConActiva, 0, 1
        If Not IsNull(rs!CANTIDAD_LEGAJOS) Then
            lblCantidad_Legajos.Caption = rs!CANTIDAD_LEGAJOS
        Else
            lblCantidad_Legajos.Caption = 0
        End If
        
        Exit Sub
salir:

MsgBox Err.Description

End Sub

Private Function LEGAJOS_CANTIDAD(COD_CLIENTE As Long) As Long
    Dim Sql As String
    Dim rs As New ADODB.Recordset
        Sql = " SELECT COUNT(*) as Cantidad_Legajos"
        Sql = Sql & vbCrLf & " From LEGAJOS"
        Sql = Sql & vbCrLf & "  WHERE COD_CLIENTE = " & COD_CLIENTE
        rs.Open Sql, ConActiva, 0, 1
        If Not IsNull(rs!CANTIDAD_LEGAJOS) Then
            LEGAJOS_CANTIDAD = rs!CANTIDAD_LEGAJOS
        Else
            LEGAJOS_CANTIDAD = 0
        End If

End Function




Private Sub cmdCargaLegajos_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset

On Error GoTo salir:

Sql = " SELECT     count(FECHA_ACTUALIZACION) as cantidad"
Sql = Sql & " From LEGAJOS"

Sql = Sql & " WHERE COD_CLIENTE = " & ctlClienteFactura.Valor

 Sql = Sql & " AND FECHA_ACTUALIZACION BETWEEN '" & mskFecha_Desde.Text
Sql = Sql & "' AND '" & mskFecha_Hasta & "'"


rs.Open Sql, ConActiva, 0, 1

If Not rs.EOF Then
      lblCargaLegajo.Caption = rs!cantidad
    
End If
Exit Sub
salir:
MsgBox Err.Description
End Sub

Private Function LEGAJOS_CARGA(COD_CLIENTE As Long) As Long
    Dim Sql As String
    Dim rs As New ADODB.Recordset

    Sql = " SELECT count(*) as cantidad"
    Sql = Sql & " From LEGAJOS"
    Sql = Sql & " WHERE COD_CLIENTE = " & COD_CLIENTE
    Sql = Sql & " AND  FECHA_CREACION BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & " AND " & FechaFormato(FECHA_FIN)
    rs.Open Sql, ConActiva, 0, 1
    If Not rs.EOF Then
        LEGAJOS_CARGA = rs!cantidad
     Else
        LEGAJOS_CARGA = 0
    End If

End Function


Private Sub cmdCuponesDisco_Click()
Dim rs As ADODB.Recordset
Dim DATO As String

Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
Dim Sql As String
Dim CAN As Integer
Sql = " SELECT     IDREQUERIMIENTO, IDTIPOREQUERIMIENTO, CONVERT(char, FECHAENTREGA, 103) AS FECHA, CANTIDAD"
Sql = Sql & " From REQUERIMIENTO"
Sql = Sql & " Where (IDTIPOREQUERIMIENTO = 26)"
Sql = Sql & " And (id_cliente = 1197)"
Sql = Sql & vbCrLf & "  AND FECHAENTREGA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
Sql = Sql & " And (ANULADO Is Null)"



    DATO = DATO & " cupones " & vbCrLf & " Desde el : " & mskFecha_Desde.Text & " hasta " & mskFecha_Hasta.Text & vbCrLf
    DATO = DATO & " IDREQUERIMIENTO " & vbTab & " FECHA " & vbTab & "CANTIDAD" & vbCrLf
    rs.Open Sql, strConBasa
    
    Do While Not rs.EOF
        DATO = DATO & rs!IDREQUERIMIENTO & vbTab & Trim(rs!fecha) & vbTab & rs!cantidad & vbCrLf
        CAN = CAN + rs!cantidad
        rs.MoveNext
    Loop
    lblCuponesDisco.Caption = CAN
    Clipboard.Clear
    Clipboard.SetText DATO
    MsgBox " Los datos fueron copiados "
    
    
End Sub

Private Sub cmdEnviadaPorCorreo_Click()
    Dim Sql As String
    Dim Control As Integer
    
    
    Sql = " Update FACTURAS"
    Sql = Sql & vbCrLf & " SET ESTADO = 30"
    Sql = Sql & vbCrLf & " Where ID_FACTURA = " & ID_Factura_Anulada.Caption
    Sql = Sql & vbCrLf & " AND ESTADO =10 "
    Control = ExecutarSql(Sql)
    If Control = 0 Then
       MsgBox "La Actualizacion NO se realizo con exito", vbCritical
    Else
       MsgBox "La Actualizacion se realizo con exito", vbInformation
    End If
    txtDescripcionAnulacion.Text = ""
    ID_Factura_Anulada.Caption = ""
    lblRazonSocialAnulada.Caption = ""

End Sub

Private Sub cmdExportExcel_Click()
    CopiarDatosGrilla grdfactura
End Sub

Private Sub cmdFacturacionCustodia_Click()
        Dim Sql As String
        Dim conData As New ADODB.Connection
        Dim rsTEM_IVA_DATA As New ADODB.Recordset
        Dim fecha As Long
        
            fecha = DateDiff("d", "28/12/1800", txtFechaDatas.Text)
            conData.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=datas"
            Sql = "SELECT FacturaABC, NumeroFactura, MesFacturacion, AnoFacturacion  "
            Sql = Sql & " , Impresa ,  DATEADD(DD, FechaFacturacion, '01/01/1800')   as   Fecha "
            
            grdFacturaCustodia.RecordSelectors = True
            grdFacturaCustodia.Columns(1).Button = True
            Sql = "SELECT  FacturaABC, NumeroFactura,  IDCliente "
            Sql = Sql & " , CUIT , Subtotal, IVAInscripto "
            Sql = Sql & " , TotalFacturado , MesFacturacion , AnoFacturacion "
            Sql = Sql & " , NombreCliente, FechaFacturacion , DetallePago , impresa "
            Sql = Sql & " From factura  "
            Sql = Sql & " Where FechaFacturacion = " & fecha
            If txtClienteCustodia.Text <> "" Then
                Sql = Sql & " AND  IDCliente = " & txtClienteCustodia.Text
            End If
         Sql = Sql & " Order by FacturaABC , NumeroFactura "
            Set RsFactura = New ADODB.Recordset
            RsFactura.CursorLocation = adUseClient
            RsFactura.Open Sql, conData
            
    Rem         RsFactura.Sort = "NumeroFactura"
            Set grdFacturaCustodia.DataSource = RsFactura.DataSource

            conData.Execute Sql
End Sub

Private Sub cmdFacturacionMensual_Click()
    Facturacion_Mensual
End Sub

Private Sub cmdFacturaEntregada_Click()
   Dim Sql As String
   Dim Control As Integer
    
    
        Sql = " Update FACTURAS"
        Sql = Sql & vbCrLf & " SET ESTADO = 20"
        Sql = Sql & vbCrLf & " Where ID_FACTURA = " & ID_Factura_Anulada.Caption
        Sql = Sql & vbCrLf & " AND ESTADO =10 "
       
        
       Control = ExecutarSql(Sql)
         If Control = 0 Then
        MsgBox "La Actualizacion NO se realizo con exito", vbCritical
    Else
            MsgBox "La Actualizacion se realizo con exito", vbInformation
    End If
         txtDescripcionAnulacion.Text = ""
         ID_Factura_Anulada.Caption = ""
         lblRazonSocialAnulada.Caption = ""

End Sub

Private Sub cmdImprimirInforme_Click()
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        Sql = " SELECT   * "
        Sql = Sql & " From basasql.dbo.FACTURACION"
        Sql = Sql & " ORDER BY COD_CLIENTE"
        frmReportes.ImprimirReporte PasoReportes & "Facturacion_Jose.rpt", Sql, True
        MsgBox "Operacion terminada"
End Sub

Private Sub cmdInforme_Click()
mnuInformePorCliente_Click
End Sub

Private Sub cmdPendienteCobro_Click()
Dim fecha As String
Dim Detalle_Pago As String
Dim Sql As String
Dim Control As Integer
fecha = InputBox("Ingrese la fecha de compromiso de pago")

 If Not IsDate(fecha) Then
  MsgBox "Fecha incorrrecta"
    Exit Sub
 End If
 
 Detalle_Pago = InputBox("Ingrese el Detalle de Pago")
 If Trim(Detalle_Pago) = "" Then
    Detalle_Pago = "NULL"
 Else
    Detalle_Pago = "'" & Trim(Detalle_Pago) & "'"
 End If
 
 
 
    Sql = " Update FACTURAS"
    Sql = Sql & " SET FECHA_PAGO = '" & fecha & "'"
    Sql = Sql & "  , ESTADO =40"
      Sql = Sql & "  , Detalle_Pago = " & Detalle_Pago
    Sql = Sql & " Where ID_FACTURA = " & ID_Factura_Anulada.Caption
     Sql = Sql & " and ESTADO in(10,20,30,40)  "
   
    Control = ExecutarSql(Sql)
         If Control = 0 Then
        MsgBox "La Actualizacion NO se realizo con exito", vbCritical
    Else
            MsgBox "La Actualizacion se realizo con exito", vbInformation
    End If
    
End Sub

Private Sub cmdRearchivo_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset

On Error GoTo salir
Sql = "SELECT  SUM(CANTIDAD) AS Cantidad"
Sql = Sql & " From ORDENAR_DOCUMENTACION"
Sql = Sql & " WHERE COD_CLIENTE = " & ctlClienteFactura.Valor
Sql = Sql & " AND FECHA BETWEEN '" & mskFecha_Desde.Text
Sql = Sql & "' AND '" & mskFecha_Hasta & "'"
rs.Open Sql, ConActiva, 0, 1

If Not rs.EOF Then
      lblRearchivo.Caption = rs!cantidad
    
End If

Dim DATO As String

DATO = "Rearchivo Fisico " & vbCrLf
DATO = DATO & "Cliente " & ctlClienteFactura.Descripcion
Set rs = New ADODB.Recordset
Sql = " SELECT     COD_CLIENTE, COD_REMITO_PRO,FECHA, SUM(CANTIDAD) AS Cantidad"
Sql = Sql & " From ORDENAR_DOCUMENTACION "
Sql = Sql & " WHERE  FECHA BETWEEN '" & mskFecha_Desde.Text
Sql = Sql & "' AND '" & mskFecha_Hasta & "'"
Sql = Sql & "  GROUP BY COD_CLIENTE, COD_REMITO_PRO,FECHA"
Sql = Sql & "  Having COD_CLIENTE = " & ctlClienteFactura.Valor
Sql = Sql & "  ORDER BY FECHA "


rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
        DATO = DATO & rs!COD_CLIENTE & vbTab & rs!COD_REMITO_PRO & vbTab & rs!fecha & vbTab & rs!cantidad & vbCrLf
        rs.MoveNext
    Loop
    Clipboard.Clear
    Clipboard.SetText DATO
    MsgBox "Los Datos Fueron Copiados"



Exit Sub
salir:
MsgBox Err.Description

End Sub


Private Sub cmdReciboBuscarFactura_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
    Sql = " SELECT FACTURAS.ID_FACTURA, FACTURAS.TIPO_FACTURA, "
    Sql = Sql & vbCrLf & " FACTURAS.NUMERO_FACTURA, FACTURAS.MONTO_CON_IVA, "
    Sql = Sql & vbCrLf & " Clientes.RAZON_SOCIAL, Clientes.id_cliente "
    Sql = Sql & vbCrLf & "  From FACTURAS, Clientes "
    Sql = Sql & vbCrLf & "  Where FACTURAS.COD_CLIENTE = Clientes.id_cliente"
     Sql = Sql & vbCrLf & "  AND (FACTURAS.ESTADO < 50) "
   Sql = Sql & vbCrLf & "  AND COD_RECIBO IS NULL "
    Sql = Sql & vbCrLf & "  AND FACTURAS.TIPO_FACTURA = '" & UCase(txtReciboTipoFactura.Text) & "' "
    Sql = Sql & vbCrLf & "  AND FACTURAS.NUMERO_FACTURA =" & txtreciboNumeroFactura.Text
    rs.Open Sql, ConActiva, 0, 1
    lblCod_Cliente_Recibo.Caption = ""
If Not rs.EOF Then
    lblReciboMontoFactura.Caption = rs!MONTO_CON_IVA
    lblReciboIDFactura.Caption = rs!ID_FACTURA
    lblReciboRazonSocial.Caption = rs!RAZON_SOCIAL
    lblCod_Cliente_Recibo.Caption = rs!id_cliente
Else
 MsgBox "Verifique la Factura", vbCritical
 Exit Sub
End If

End Sub

Private Sub cmdReciboFacturaGrilla_Click()
If grdReciboFacuta.Rows = 2 And grdReciboFacuta.TextMatrix(1, 0) = "" Then
    grdReciboFacuta.TextMatrix(1, 0) = 1
    grdReciboFacuta.TextMatrix(1, 1) = lblReciboIDFactura.Caption
    grdReciboFacuta.TextMatrix(1, 2) = txtReciboTipoFactura.Text
    grdReciboFacuta.TextMatrix(1, 3) = txtreciboNumeroFactura.Text
    grdReciboFacuta.TextMatrix(1, 4) = lblReciboMontoFactura.Caption
    grdReciboFacuta.TextMatrix(1, 5) = lblReciboRazonSocial.Caption
Else
    grdReciboFacuta.AddItem grdReciboFacuta.Rows & vbTab & lblReciboIDFactura.Caption & vbTab & txtReciboTipoFactura.Text & vbTab & txtreciboNumeroFactura.Text & vbTab & lblReciboMontoFactura.Caption & vbTab & lblReciboRazonSocial.Caption
End If
SumarTotalReciboFactura
    lblReciboIDFactura.Caption = ""
    txtReciboTipoFactura.Text = ""
    txtreciboNumeroFactura.Text = ""
    lblReciboMontoFactura.Caption = ""
    lblReciboRazonSocial.Caption = ""
End Sub

Private Sub cmdReciboResponsable_Click()
    Dim rsMax As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim ID_Recibo As Long
    Dim R As Integer
        Sql = " SELECT NUMERO_RECIBO"
        Sql = Sql & " From RECIBOS "
        Sql = Sql & "  Where NUMERO_RECIBO = " & txtReciboInicio.Text
        
        rs.Open Sql, ConActiva, 0, 1
        If Not rs.EOF Then
            MsgBox "El recibo ya esta creado", vbCritical
            Exit Sub
        End If
        Sql = " SELECT MAX(ID_RECIBO) as MaxRecibo From RECIBOS "

If IsNull(ctlPersonalRecibo.Valor) Then
    MsgBox "Falta el responsable"
    Exit Sub
Else
End If


    rsMax.Open Sql, ConActiva, 0, 1
    ID_Recibo = rsMax!MaxRecibo + 1
        For R = txtReciboInicio.Text To txtReciboFin.Text
            
            Set rs = New ADODB.Recordset
            Sql = " SELECT NUMERO_RECIBO"
        Sql = Sql & " From RECIBOS "
        Sql = Sql & "  Where NUMERO_RECIBO = " & R
            
            rs.Open Sql, ConActiva, 0, 1
            
            If rs.EOF Then
                Sql = "  INSERT INTO RECIBOS"
                Sql = Sql & vbCrLf & "  (ID_RECIBO, "
                Sql = Sql & vbCrLf & "  NUMERO_RECIBO,"
                Sql = Sql & vbCrLf & "  RESPONSABLE,"
                Sql = Sql & vbCrLf & "  ESTADO_RECIBO )"
                Sql = Sql & vbCrLf & " VALUES ("
                Sql = Sql & vbCrLf & ID_Recibo & ","
                Sql = Sql & vbCrLf & R & ","
                Sql = Sql & vbCrLf & ctlPersonalRecibo.Valor & ",10 )"
                ExecutarSql Sql
                ID_Recibo = ID_Recibo + 1
            Else
                MsgBox "El recibo " & R & " ya existe"
            End If
            
                
        Next
MsgBox "Tarea Completa", vbInformation


End Sub



Private Sub cmsCantidadCajasVacias_Click()
Dim Sql As String
Dim rs As ADODB.Recordset
Dim DATO As String
     
     Sql = " SELECT NRO_REMITO, NRO_REM_PROV, FECHA, CANTIDAD "
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " Where id_cliente =" & ctlClienteFactura.Valor
    Sql = Sql & vbCrLf & " And TIPO = 2"
    Sql = Sql & vbCrLf & " And ANULADO Is Null"
    Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
    Sql = Sql & vbCrLf & " Order by NRO_REMITO"

    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    DATO = ctlClienteFactura.Descripcion & vbCrLf
     DATO = DATO & " Cajas Vacias " & vbCrLf & " Desde el : " & mskFecha_Desde.Text & " hasta " & mskFecha_Hasta.Text & vbCrLf
     DATO = DATO & " Remito " & vbTab & "Fecha" & vbTab & " Cantidad " & vbCrLf
     
    Do While Not rs.EOF
        DATO = DATO & rs!NRO_REMITO & vbTab & rs!fecha & vbTab & rs!cantidad & vbCrLf
        rs.MoveNext
    Loop
    Clipboard.Clear
    Clipboard.SetText DATO
    
    MsgBox "Los datos fueron Coipiados"
    
End Sub

Private Sub cmdGrabarFactura_Click()
    Dim Sql As String
    Dim rsMax As New ADODB.Recordset
    rsMax.Open "SELECT max(ID_FACTURA) as MaxFactura FROM FACTURAS ", ConActiva, 0, 1
    Dim conFactura As New ADODB.Connection
    conFactura.Open strConBasa
    conFactura.CursorLocation = adUseClient
    Dim ID_FACTURA As Integer
    Dim TIPO_COMPROBANTE As String
    Dim Tipo_Factura As String
    Dim NUMERO_FACTURA As Long
    Dim COD_CLIENTE As Long
    Dim FECHA_FACTURA As String
    Dim MONTO_SIN_IVA As String
    Dim MONTO_CON_IVA As String
    Dim IVA As String
    Dim Descripcion As String
    Dim estado As Integer
    
    On Error GoTo salir:
    
    conFactura.BeginTrans
    ID_FACTURA = rsMax!maxFactura + 1
    TIPO_COMPROBANTE = cboTipoComprobante.Text
    Tipo_Factura = lblTipoFactura.Caption
    NUMERO_FACTURA = txtNumeroFactura.Text
    COD_CLIENTE = ctlClienteFactura.Valor
    FECHA_FACTURA = txtFechaFactura.Text
    estado = 10
   If BaseOracle = True Then
    MONTO_SIN_IVA = "'" & lblSubTotal.Caption & "'"
    MONTO_CON_IVA = "'" & lblTotal.Caption & "'"
    IVA = "'" & lblIVA.Caption & "'"
    Else
    MONTO_SIN_IVA = "'" & Replace(lblSubTotal.Caption, ",", ".") & "'"
    MONTO_CON_IVA = "'" & Replace(lblTotal.Caption, ",", ".") & "'"
    IVA = "'" & Replace(lblIVA.Caption, ",", ".") & "'"
    End If
    If txtDescripcion_Factura.Text <> "" Then
        Descripcion = "'" & txtDescripcion_Factura.Text & "'"
    Else
        Descripcion = "NULL"
    End If
    
    
Sql = " INSERT INTO FACTURAS "
Sql = Sql & vbCrLf & " (ID_FACTURA,TIPO_COMPROBANTE , "
Sql = Sql & vbCrLf & " TIPO_FACTURA, NUMERO_FACTURA, "
Sql = Sql & vbCrLf & " COD_CLIENTE, FECHA,"
Sql = Sql & vbCrLf & " ESTADO, MONTO_SIN_IVA,"
Sql = Sql & vbCrLf & "  MONTO_CON_IVA, IVA"
Sql = Sql & vbCrLf & " ,DESCRIPCION , MESFACTURACION ,PERIODOFACTURACION )"
Sql = Sql & vbCrLf & "  VALUES  "
Sql = Sql & vbCrLf & " (" & ID_FACTURA & ",'" & TIPO_COMPROBANTE & "',"
Sql = Sql & vbCrLf & "'" & Tipo_Factura & "'," & NUMERO_FACTURA & ","
Sql = Sql & vbCrLf & COD_CLIENTE & ",'" & FECHA_FACTURA & "',"
Sql = Sql & vbCrLf & estado & "," & MONTO_SIN_IVA
Sql = Sql & vbCrLf & "," & MONTO_CON_IVA & "," & IVA & ","
Sql = Sql & vbCrLf & Descripcion & ",'" & txtMesFacturacion & "'"
Sql = Sql & vbCrLf & ",'" & "DESDE " & mskFecha_Desde.Text & "HASTA " & mskFecha_Hasta.Text & "')"
conFactura.Execute Sql

Sql = " Update Clientes "
Sql = Sql & vbCrLf & "  SET PERIODO_FACTURA =" & ControldatoString("Fecha Inicio:" & mskFecha_Desde.Text & " Fecha hasta:" & mskFecha_Hasta.Text)
Sql = Sql & vbCrLf & " , DETALLE_FACTURACION =" & ControldatoString(txtDetalleFacturacion.Text)
Sql = Sql & vbCrLf & "  Where id_cliente = " & COD_CLIENTE

conFactura.Execute Sql


Dim Item, cantidad, Codigo, detalle, IMPORTE_UNITARIO, IMPORTE_TOTAL As String
Dim i As Integer

For i = 1 To grdFacturacion.Rows - 1
    Item = grdFacturacion.TextMatrix(i, 0)
    cantidad = grdFacturacion.TextMatrix(i, 1)
    Codigo = grdFacturacion.TextMatrix(i, 2)
    detalle = grdFacturacion.TextMatrix(i, 3)
    
    If BaseOracle = True Then
        IMPORTE_UNITARIO = grdFacturacion.TextMatrix(i, 4)
        IMPORTE_TOTAL = grdFacturacion.TextMatrix(i, 5)
    Else
        IMPORTE_UNITARIO = Replace(grdFacturacion.TextMatrix(i, 4), ",", ".")
        IMPORTE_TOTAL = Replace(grdFacturacion.TextMatrix(i, 5), ",", ".")
    End If
    
    
    Sql = " INSERT INTO FACTURAS_DETALLES "
    Sql = Sql & vbCrLf & " (COD_FACTURA, ITEM, CANTIDAD, CODIGO, DETALLE,"
    Sql = Sql & vbCrLf & " IMPORTE_UNITARIO, IMPORTE_TOTAL)"
    Sql = Sql & vbCrLf & "  VALUES (" & ID_FACTURA
    Sql = Sql & vbCrLf & "," & Item
    Sql = Sql & vbCrLf & ",'" & cantidad & "'"
    Sql = Sql & vbCrLf & ",'" & Codigo & "'"
    Sql = Sql & vbCrLf & ",'" & detalle & "'"
    Sql = Sql & vbCrLf & ",'" & IMPORTE_UNITARIO & "'"
    Sql = Sql & vbCrLf & ",'" & IMPORTE_TOTAL & "')"
        conFactura.Execute Sql
    
    
Next
    conFactura.CommitTrans
MsgBox "factura  " & ID_FACTURA
    If MsgBox("Usted quiere imprimir la factura", vbYesNo, "Imprimir Factura") = vbYes Then
        If lblTipoFactura.Caption = "A" Then
            ImprimirFacura_A CLng(ID_FACTURA)
        Else
            ImprimirFacura_B CLng(ID_FACTURA)
        End If
    End If

 LimpiarCampos
  grdFacturacion.Clear
    grdFacturacion.Rows = 2
    TitulosGrilla
 ctlClienteFactura.SetFocus
 conFactura.Close
 Exit Sub
salir:
conFactura.RollbackTrans
conFactura.Close
MsgBox "error en grabacion Factura", vbCritical
End Sub

Private Sub cmdInsertFacturacion_Click()
    InsertDatoFactura
    txtCantidad.Text = ""
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    txtPrecioUnitario.Text = ""
    txtTotal.Text = ""
    txtCantidad.SetFocus
End Sub

Private Sub Command2_Click()
    Dim Sql As String
    Dim ID_FACTURA As Integer
    Dim rs As ADODB.Recordset
    Dim MesRemplazo As String
    ID_FACTURA = InputBox("Ingrese el numero de ID Factura")
    MesRemplazo = InputBox("Ingrese el mes a remplazar")
    
Sql = " SELECT FACTURAS.ID_FACTURA, FACTURAS_DETALLES.ITEM, FACTURAS_DETALLES.CANTIDAD, FACTURAS_DETALLES.CODIGO,"
Sql = Sql & " FACTURAS_DETALLES.DETALLE , FACTURAS_DETALLES.IMPORTE_UNITARIO, FACTURAS_DETALLES.IMPORTE_TOTAL"
Sql = Sql & " FROM  FACTURAS INNER JOIN"
Sql = Sql & " FACTURAS_DETALLES ON FACTURAS.ID_FACTURA = FACTURAS_DETALLES.COD_FACTURA"
Sql = Sql & " Where FACTURAS.ID_FACTURA = " & ID_FACTURA
Sql = Sql & " ORDER BY FACTURAS.ID_FACTURA DESC, FACTURAS_DETALLES.ITEM"
Set rs = New ADODB.Recordset
rs.Open Sql, ConActiva, 0, 1

Dim i As Integer
grdFacturacion.Clear
TitulosGrilla

i = 0
With grdFacturacion

.Rows = 1
Do While Not rs.EOF
.AddItem ""
    
        grdFacturacion.TextMatrix(.Rows - 1, 0) = i
        grdFacturacion.TextMatrix(.Rows - 1, 1) = rs!cantidad
        grdFacturacion.TextMatrix(.Rows - 1, 2) = Trim(rs!Codigo)
        grdFacturacion.TextMatrix(.Rows - 1, 3) = Replace(Trim(rs!detalle), MesRemplazo, txtPeriodoActual.Text)
        grdFacturacion.TextMatrix(.Rows - 1, 4) = Trim(rs!IMPORTE_UNITARIO)
        grdFacturacion.TextMatrix(.Rows - 1, 5) = Trim(rs!IMPORTE_TOTAL)
      
    
    
    rs.MoveNext
  Loop
  End With
    
    Dim Valor As Double

 
 
 Sumar
        
        
        
        
End Sub
Private Function sGetTitleAuthors() As String
Dim rstParent   As ADODB.Recordset
Dim rstChild    As ADODB.Recordset
Dim sBuf        As String
   
Const CONNECT_PUBS = "PROVIDER=MSDataShape;DATA PROVIDER=SQLOLEDB;" & _
    "SERVER=;DATABASE=pubs;UID=sa;PWD=;"
Const SHAPE_TITLEAUTHORS = _
    "SHAPE {SELECT au_id, au_lname, au_fname FROM authors} " & _
    "APPEND ({SELECT au_id, title FROM titleauthor TA, titles TS " & _
             "WHERE TA.title_id = TS.title_id} " & _
            "AS title_chap RELATE au_id TO au_id)"
            
    '----- create rowsets
    Set rstParent = New ADODB.Recordset
    rstParent.Open SHAPE_TITLEAUTHORS, CONNECT_PUBS
            
    '----- process parent rowset
    Do While Not rstParent.EOF
        sBuf = sBuf & rstParent("au_id") & vbTab & _
            rstParent("au_lname") & ", " & rstParent("au_fname") & vbCrLf
            
        '----- process chapter of child rowset
        Set rstChild = rstParent("title_chap").value
        Do While Not rstChild.EOF
            sBuf = sBuf & vbTab & vbTab & rstChild("title") & vbCrLf
            rstChild.MoveNext
        Loop
        rstParent.MoveNext
    Loop
    sGetTitleAuthors = sBuf
End Function


    
    


Private Sub Command1_Click()

Dim Sql As String
Dim i As Integer


Sql = " CREATE TABLE " & "TARIFAS_FACTURA_" & Format(date, "DDMMYYYY") & " AS"
Sql = Sql & " SELECT COD_CLIENTE, CANON_CAJA, CANON_LIBRO,"
 Sql = Sql & "    CANON_LEGAJO, CAJA, REFERENCIA, CARGAR_LEGAJOS,"
 Sql = Sql & "    CONSULTA, FLETE_NORMAL, FLETE_URGENTE, PRECINTO,"
  Sql = Sql & "   HORA_ARCHIVISTA_BASA, HORA_ARCHIVISTA_CLIENTE,"
  Sql = Sql & "   ABONO_MINIMO, IMAGEN, REACHIVO_FISICO,"
  Sql = Sql & "   REARCHIVO_LOTE , LICITACION, Usuario"
Sql = Sql & " From TARIFAS_FACTURA"

ExecutarSql Sql



Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient


    Sql = " SELECT CANON_CAJA, CANON_LIBRO, CANON_LEGAJO, CAJA,"
    Sql = Sql & " REFERENCIA, CARGAR_LEGAJOS, CONSULTA,"
    Sql = Sql & " FLETE_NORMAL, FLETE_URGENTE, PRECINTO,"
    Sql = Sql & " HORA_ARCHIVISTA_BASA, HORA_ARCHIVISTA_CLIENTE,"
    Sql = Sql & " IMAGEN, REACHIVO_FISICO,"
    Sql = Sql & " REARCHIVO_LOTE"
    Sql = Sql & " From TARIFAS_FACTURA"
    Sql = Sql & " Where (LICITACION Is Null)"
    Sql = Sql & " ORDER BY COD_CLIENTE"

rs.Open Sql, ConActiva, adOpenDynamic, adLockPessimistic

        Do While Not rs.EOF
            For i = 0 To rs.Fields.Count - 1
                If IsNull(rs.Fields.Item(i).value) Then
                   rs.Fields.Item(i).value = 0
                Else
                   rs.Fields.Item(i).value = rs.Fields.Item(i).value * txtIncremento.Text
                End If
             Next
            rs.Update
            rs.MoveNext
        Loop



'
'CREATE TABLE "Incremento10"
'AS
'SELECT COD_CLIENTE, CANON_CAJA, CANON_LIBRO,
'    CANON_LEGAJO, CAJA, REFERENCIA, CARGAR_LEGAJOS,
'    CONSULTA, FLETE_NORMAL, FLETE_URGENTE, PRECINTO,
'    HORA_ARCHIVISTA_BASA, HORA_ARCHIVISTA_CLIENTE,
'    ABONO_MINIMO, IMAGEN, REACHIVO_FISICO,
'    REARCHIVO_LOTE , LICITACION, Usuario
'From TARIFAS_FACTURA
'
'
'
'
'Dim Rs As New ADODB.Recordset
'Dim Sql As String
'Dim Mensaje As String
'
'Sql = " SELECT APELLIDO_NOMBRE, CORREO"
'Sql = Sql & " From CLIENTEUSUARIO"
'Sql = Sql & "  Where (Not (correo Is Null))"
'Sql = Sql & "  ORDER BY APELLIDO_NOMBRE"
'
'Mensaje = " El personal de Banco de Archivos s.a  "
'Mensaje = Mensaje & vbCrLf & " Queremos  hacerle llegar nuestros más sinceros deseos, "
'Mensaje = Mensaje & vbCrLf & " que sea esta Navidad motivo de muchas felicidades."
'Mensaje = Mensaje & vbCrLf & " Y el Año Nuevo una esperanza de éxito y prosperidad. "
'Mensaje = Mensaje & vbCrLf & " Paz y Amor en estas Fiestas"
'Mensaje = Mensaje & vbCrLf
'Mensaje = Mensaje & vbCrLf
'Mensaje = Mensaje & vbCrLf
'Mensaje = Mensaje & vbCrLf & "Saludos Cordiales"
'
'
'
' Rs.Open Sql, strConBasa , 0 ,1
'
'
'Do While Not Rs.EOF
'    SendMail Rs!correo, "Felices Fiestas", Mensaje
'    Rs.MoveNext
'Loop



End Sub

Private Sub Command3_Click()
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim CAJAS As Long
    Dim CajasBajas As Long
    Dim LIBROS As Long
    Dim LibrosBajas As Long


If mskFecha_Desde.Text = "__/__/____" Then
       MsgBox "FECHA INICIO"
    Exit Sub
End If



'__________________Inicio Cajas ____________________________________
        
        Sql = " SELECT SUM(CANTIDAD) As cantidadCajas "
        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
        Sql = Sql & vbCrLf & " Where id_cliente =" & ctlClienteFactura.Valor
        Sql = Sql & vbCrLf & " And TIPO = 0"
        Sql = Sql & vbCrLf & " And ANULADO Is Null"
        Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
        If ctlClienteFactura.Valor > 1000 Then
            Sql = Sql & vbCrLf & " AND FECHA >  " & FechaFormato("29/04/2014")
        End If
        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(FECHA_FIN)
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
        CAJAS = 0
        Else
        If IsNull(rs!CANTIDADCAJAS) Then
            CAJAS = 0
        Else
            CAJAS = rs!CANTIDADCAJAS
        End If
        End If
        ' Bajas
        Sql = "  SELECT SUM(CANTIDAD) AS BajasCajas "
        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
        Sql = Sql & vbCrLf & " Where id_cliente = " & ctlClienteFactura.Valor
        Sql = Sql & vbCrLf & "  And (TIPO = 3) And (ANULADO Is Null)"
        Sql = Sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 0"
        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(FECHA_FIN)
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
            CajasBajas = 0
         Else
            If IsNull(rs!BajasCajas) Then
                CajasBajas = 0
            Else
                CajasBajas = rs!BajasCajas
            End If
        End If
        
        
                
        
        lblCantidad_Cajas = (CLng(CAJAS) - CLng(CajasBajas)) + CajasSumaResta(ctlClienteFactura.Valor)
        CambioColor lblCantidad_Cajas
        
        
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
'
'
'
'
'
'
'        Sql = " SELECT SUM(CANTIDAD) As cantidadCajas "
'        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
'        Sql = Sql & vbCrLf & " Where id_cliente =" & ctlClienteFactura.Valor
'        Sql = Sql & vbCrLf & " And TIPO = 0"
'        Sql = Sql & vbCrLf & " And ANULADO Is Null"
'        Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
'        If COD_CLIENTE > 1000 Then
'            Sql = Sql & vbCrLf & " AND FECHA >  " & FechaFormato("29/04/2014")
'        End If
'
'        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(mskFecha_Hasta.Text)
'        Set rs = New ADODB.Recordset
'        rs.Open Sql, ConActiva, 0, 1
'        If rs.EOF Then
'        CAJAS = 0
'        Else
'        If IsNull(rs!CANTIDADCAJAS) Then
'            CAJAS = 0
'        Else
'            CAJAS = rs!CANTIDADCAJAS
'        End If
'        End If
'        ' Bajas
'        Sql = "  SELECT SUM(CANTIDAD) AS BajasCajas "
'        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
'        Sql = Sql & vbCrLf & " Where id_cliente = " & ctlClienteFactura.Valor
'        Sql = Sql & vbCrLf & "  And (TIPO = 3) And (ANULADO Is Null)"
'        Sql = Sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 0"
'         If COD_CLIENTE > 1000 Then
'            Sql = Sql & vbCrLf & " AND FECHA >  " & FechaFormato("29/04/2014")
'        End If
'
'        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(mskFecha_Hasta.Text)
'        Set rs = New ADODB.Recordset
'        rs.Open Sql, ConActiva, 0, 1
'        If rs.EOF Then
'            CajasBajas = 0
'         Else
'            If IsNull(rs!BajasCajas) Then
'                CajasBajas = 0
'            Else
'                CajasBajas = rs!BajasCajas
'            End If
'        End If
'
'        lblCantidad_Cajas = CLng(CAJAS) - CLng(CajasBajas) + CajasSumaResta(ctlClienteFactura.Valor)
'        CambioColor lblCantidad_Cajas
'
'_________________________Fin Cajas ____________________________________

    Sql = " SELECT SUM(CANTIDAD) As CantidadCajasMes "
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " Where id_cliente =" & ctlClienteFactura.Valor
    Sql = Sql & vbCrLf & " And TIPO = 0"
    Sql = Sql & vbCrLf & " And ANULADO Is Null"
    Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)

    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    If rs.EOF Then
        lblCajasCrecimientoMes = 0
        
    Else
        If IsNull(rs!CantidadCajasMes) Then
            lblCajasCrecimientoMes = 0
            
        Else
            lblCajasCrecimientoMes.Caption = rs!CantidadCajasMes
            
        End If
     End If

CambioColor lblCajasCrecimientoMes


'_________________________Inicio Libros_____________________________________

        Sql = " SELECT SUM(CANTIDAD) As CantidadLibros"
        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
        Sql = Sql & vbCrLf & " Where id_cliente =" & ctlClienteFactura.Valor
        Sql = Sql & vbCrLf & " And TIPO = 0"
        Sql = Sql & vbCrLf & " And ANULADO Is Null"
        Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 1"
        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(mskFecha_Hasta.Text)
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
            LIBROS = 0
        Else
            If IsNull(rs!CantidadLibros) Then
                LIBROS = 0
            Else
                LIBROS = rs!CantidadLibros
            End If
        End If
        ' Bajas
        Sql = "  SELECT SUM(CANTIDAD) AS LibrosBajas "
        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
        Sql = Sql & vbCrLf & " Where id_cliente = " & ctlClienteFactura.Valor
        Sql = Sql & vbCrLf & "  And (TIPO = 3) And (ANULADO Is Null)"
        Sql = Sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 1"
        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(mskFecha_Hasta.Text)
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
            LibrosBajas = 0
         Else
            If IsNull(rs!LibrosBajas) Then
                LibrosBajas = 0
            Else
                LibrosBajas = rs!LibrosBajas
            End If
        End If
        
         
        lblCantidad_Libros.Caption = CLng(LIBROS) - CLng(LibrosBajas)
        
        CambioColor lblCantidad_Libros
        



'CajasVacias

Sql = "  SELECT SUM(CANTIDAD)as CantidadCajasVacias"
Sql = Sql & vbCrLf & "  From REMITOS_CUERPO"
Sql = Sql & vbCrLf & "  Where id_cliente = " & ctlClienteFactura.Valor
Sql = Sql & vbCrLf & " And (TIPO = 2) And (ANULADO Is Null)"
Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
Set rs = New ADODB.Recordset
rs.Open Sql, ConActiva, 0, 1
     If rs.EOF Then
        lblCantidadCajasVacias.Caption = 0
    Else
        If Not IsNull(rs!CantidadCajasVacias) Then
            lblCantidadCajasVacias.Caption = rs!CantidadCajasVacias
         Else
            lblCantidadCajasVacias.Caption = 0
         End If
         
    End If
    
    CambioColor lblCantidadCajasVacias
    
    
    
    
 ' COnsultas o Desarchivo
 
 
 Sql = "  SELECT SUM(CANTIDAD)as CantidadDesarchivo"
Sql = Sql & vbCrLf & " From REMITOS_CUERPO"
Sql = Sql & vbCrLf & " Where id_cliente =  " & ctlClienteFactura.Valor
Sql = Sql & vbCrLf & " And TIPO  in( 1 )  And ANULADO Is Null"
Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
Sql = Sql & vbCrLf & "  AND OPERACION  = 1  "
    
    Set rs = New ADODB.Recordset
rs.Open Sql, ConActiva, 0, 1
     If rs.EOF Then
        lblCantidadDesarchivos.Caption = 0
    Else
         If IsNull(rs!cantidadDesarchivo) Then
            lblCantidadDesarchivos.Caption = 0
         Else
            lblCantidadDesarchivos.Caption = rs!cantidadDesarchivo
         End If
    End If
    
 Sql = "  SELECT SUM(CANTIDAD)as CantidadDesarchivo"
Sql = Sql & vbCrLf & " From REMITOS_CUERPO"
Sql = Sql & vbCrLf & " Where id_cliente =  " & ctlClienteFactura.Valor
Sql = Sql & vbCrLf & " And TIPO  =  3  And ANULADO Is Null"
Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)



    
    Set rs = New ADODB.Recordset
rs.Open Sql, ConActiva, 0, 1
     If rs.EOF Then
        lblCantidadDesarchivos.Caption = 0
    Else
         If IsNull(rs!cantidadDesarchivo) Then
           
         Else
            lblCantidadDesarchivos.Caption = lblCantidadDesarchivos.Caption + rs!cantidadDesarchivo
         End If
    End If
    
    
    
    
    CambioColor lblCantidadDesarchivos







 ' control fletes
 Sql = " SELECT COD_FLETE"
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " where  COD_FLETE is null "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN '" & mskFecha_Desde.Text & "'"
    Sql = Sql & vbCrLf & "  AND '" & mskFecha_Hasta.Text & "'"
    
    
    Set rs = New ADODB.Recordset
'    rs.Open sql, strConBasa , 0 ,1
'    If Not rs.EOF Then
'        MsgBox "Error en asisgnacion de fletes", vbCritical
'
'    End If


' Fletes Normales



    Sql = "  SELECT COUNT(DISTINCT COD_FLETE) as FleteNormal"
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " WHERE ID_CLIENTE = " & ctlClienteFactura.Valor
    Sql = Sql & vbCrLf & " AND Estado = 0 "
    Sql = Sql & vbCrLf & " AND COD_FLETE <> 0 "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
    
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    If rs.EOF Then
        lbl_FletesNormales.Caption = 0
    Else
         If IsNull(rs!FleteNormal) Then
            lbl_FletesNormales.Caption = 0
         Else
            lbl_FletesNormales.Caption = rs!FleteNormal
         End If
    End If
    

 CambioColor lbl_FletesNormales



' Fletes Urgentes

    Sql = "  SELECT COUNT(DISTINCT COD_FLETE) as FleteUrgente"
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " WHERE ID_CLIENTE = " & ctlClienteFactura.Valor
    Sql = Sql & vbCrLf & " AND Estado = 1 "
    Sql = Sql & vbCrLf & " AND COD_FLETE <> 0 "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(mskFecha_Hasta.Text)
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    If rs.EOF Then
        lbl_FletesUrgentes.Caption = 0
    Else
         If IsNull(rs!FleteUrgente) Then
            lbl_FletesUrgentes.Caption = 0
         Else
            lbl_FletesUrgentes.Caption = rs!FleteUrgente
         End If
    End If

CambioColor lbl_FletesUrgentes

'
'
'
'
'

'
'SELECT SUM(CANTIDAD) AS EXPR1
'From REMITOS_CUERPO
'WHERE (ID_CLIENTE = 40) AND ANULADO IS NULL AND
'    tipo = 0 AND (FECHA BETWEEN TO_DATE('01/07/2007',
'    'DD/MM/YYYY') AND TO_DATE('30/10/2007', 'DD/MM/YYYY'))
'GROUP BY TIPO
'
'
' Cajas vacias
'
'
' SELECT SUM(CANTIDAD) AS EXPR1
'From REMITOS_CUERPO
'WHERE (ID_CLIENTE = 40) AND ANULADO IS NULL AND
'    TIPO = 2 AND (FECHA BETWEEN TO_DATE('01/07/2007',
'    'DD/MM/YYYY') AND TO_DATE('30/10/2007', 'DD/MM/YYYY'))
'
'
'
'
'

End Sub

Private Sub Command6_Click()
 Dim i As Integer
 For i = 1 To 69
 
 MsgBox Chr(i)
 
 Next
   MsgBox Asc("g")
   MsgBox Chr(Asc("g"))
End Sub

Private Sub CargarFletes()
Dim Sql As String
Dim rs As ADODB.Recordset
Dim Filtro As String
On Error GoTo salir

 RemitoFlete = 0
    Sql = " SELECT CLIENTES.ID_CLIENTE, CLIENTES.RAZON_SOCIAL,"
    Sql = Sql & vbCrLf & "     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA,"
    Sql = Sql & vbCrLf & "     REMITO_ESTADOS.DESCRIPCION,"
    Sql = Sql & vbCrLf & "     REMITOS_CUERPO.CANTIDAD, REMITOS_CUERPO.COD_FLETE"
    Sql = Sql & vbCrLf & "  From REMITOS_CUERPO, REMITO_ESTADOS, Clientes"
    Sql = Sql & vbCrLf & "  WHERE REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
    Sql = Sql & vbCrLf & "  (REMITOS_CUERPO.TIPO = 1) AND"
    Sql = Sql & vbCrLf & "  (REMITOS_CUERPO.OPERACION = 1) "
    If chkSoloPendientes.value = 1 Then
        Sql = Sql & vbCrLf & " AND COD_FLETE is null"
    End If
    Sql = Sql & vbCrLf & " AND REMITOS_CUERPO.FECHA > " & FechaServerTipo(txtFechaFlete.Text)
    Sql = Sql & vbCrLf & " ORDER BY REMITOS_CUERPO.Fecha , REMITOS_CUERPO.ID_CLIENTE,  REMITOS_CUERPO.NRO_REMITO"




   Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
    rs.Open Sql, strConBasa

Set grdFletes.DataSource = rs.DataSource
grdFletes.Columns.Item(0).Width = 700
grdFletes.Columns.Item(1).Width = 3500
grdFletes.Columns.Item(2).Width = 900
grdFletes.Columns.Item(3).Width = 1300
grdFletes.Columns.Item(4).Width = 900
grdFletes.Columns.Item(5).Width = 900
grdFletes.Columns.Item(6).Width = 900

If rs.RecordCount > RowBo And RowBo <> 0 Then
    grdFletes.Bookmark = RowBo
    grdFletes.Refresh
End If


salir:
End Sub

Private Sub ImprimirFacura_B(ID_FACTURA As Long)


    
    Dim Sql As String
    Sql = " SELECT V_FACTURA.FECHA, V_FACTURA.RAZON_SOCIAL, V_FACTURA.CALLE"
    Sql = Sql & vbCrLf & " , V_FACTURA.CANTIDAD, V_FACTURA.DETALLE,V_FACTURA.TIPO_COMPROBANTE "
    Sql = Sql & vbCrLf & " , V_FACTURA.IMPORTE_UNITARIO,V_FACTURA.IMPORTE_TOTAL "
    Sql = Sql & vbCrLf & " , V_FACTURA.DESCRIPCION , V_FACTURA.MONTO_SIN_IVA"
    Sql = Sql & vbCrLf & " , V_FACTURA.MONTO_CON_IVA , V_FACTURA.IVA, V_FACTURA.COD_FACTURA"
    Sql = Sql & vbCrLf & " , V_FACTURA.NRO_CUIT, V_FACTURA.ID_CLIENTE, V_FACTURA.NUMERO_FACTURA"
    Sql = Sql & vbCrLf & " , V_FACTURA.TIPO_FACTURA,V_FACTURA.TIPO_ENTREGA"
    Sql = Sql & vbCrLf & " FROM  V_FACTURA "
    Sql = Sql & vbCrLf & " WHERE  V_FACTURA.TIPO_FACTURA='B'"
    Sql = Sql & vbCrLf & " AND V_FACTURA.COD_FACTURA = " & ID_FACTURA
    Sql = Sql & vbCrLf & " ORDER BY V_FACTURA.NUMERO_FACTURA"
    frmReportes.ImprimirReporte PasoReportes & "rptFactura B.rpt", Sql, False, , , 2
    
End Sub
Private Sub ImprimirFacura_A(ID_FACTURA As Long)
    Dim Sql As String
    Sql = " SELECT V_FACTURA.FECHA, V_FACTURA.RAZON_SOCIAL, V_FACTURA.CALLE"
    Sql = Sql & vbCrLf & " , V_FACTURA.CANTIDAD, V_FACTURA.DETALLE,V_FACTURA.TIPO_COMPROBANTE "
    Sql = Sql & vbCrLf & " , V_FACTURA.IMPORTE_UNITARIO,V_FACTURA.IMPORTE_TOTAL "
    Sql = Sql & vbCrLf & " , V_FACTURA.DESCRIPCION , V_FACTURA.MONTO_SIN_IVA"
    Sql = Sql & vbCrLf & " , V_FACTURA.MONTO_CON_IVA , V_FACTURA.IVA, V_FACTURA.COD_FACTURA"
    Sql = Sql & vbCrLf & " , V_FACTURA.NRO_CUIT, V_FACTURA.ID_CLIENTE, V_FACTURA.NUMERO_FACTURA"
    Sql = Sql & vbCrLf & " , V_FACTURA.TIPO_FACTURA,V_FACTURA.TIPO_ENTREGA"
    Sql = Sql & vbCrLf & " FROM  V_FACTURA "
    Sql = Sql & vbCrLf & " WHERE  V_FACTURA.TIPO_FACTURA='A'"
    Sql = Sql & vbCrLf & " AND V_FACTURA.COD_FACTURA = " & ID_FACTURA
    Sql = Sql & vbCrLf & " ORDER BY V_FACTURA.NUMERO_FACTURA"
     Rem frmReportes.ImprimirReporte PasoReportes & "rptFactura A.rpt", sql, False, , , 2
   frmReportes.ImprimirReporte PasoReportes & "rptFactura A Impresora Custodia.rpt", Sql, False, , , 2
    
End Sub



Private Sub Command4_Click()
Dim rs As New ADODB.Recordset
Dim Tipo_Factura As String
Dim NUMERO As Integer
Dim Sql As String


Tipo_Factura = InputBox("Ingrese el Tipo Factura")
NUMERO = InputBox("Ingrese en numero de factura")

Sql = " SELECT ID_FACTURA, TIPO_FACTURA, NUMERO_FACTURA"
Sql = Sql & " From FACTURAS"
Sql = Sql & " WHERE TIPO_FACTURA = '" & Tipo_Factura & "'"
Sql = Sql & "  AND NUMERO_FACTURA >= " & NUMERO
Sql = Sql & "  ORDER BY NUMERO_FACTURA "
rs.Open Sql, ConActiva, 0, 1


Do While Not rs.EOF
  
    If Tipo_Factura = "A" Then
    
        If MsgBox("Usted quiere imprimir la factura A Numero " & rs!NUMERO_FACTURA, vbYesNo) = vbYes Then
            ImprimirFacura_A rs!ID_FACTURA
        Else
            Exit Sub
        End If
    End If
    
    
    If Tipo_Factura = "B" Then
     If MsgBox("Usted quiere imprimir la factura B Numero " & rs!NUMERO_FACTURA, vbYesNo) = vbYes Then
        ImprimirFacura_B rs!ID_FACTURA
     End If
    End If

    rs.MoveNext
Loop



End Sub

Private Sub Command45_Click()
    Dim i As Integer
    
     Dim fila As Integer
     Dim TipoFactura As String
    
    
    Dim FacturaABC As String
    Dim NumeroFactura As Long
    Dim IDCliente As Integer
    
    Dim Cuit As String
    Dim Subtotal As String
    Dim IVAInscripto As String
    
    Dim TotalFacturado As String
    Dim MesFacturacion As Integer
    Dim AnoFacturacion As Integer
    
    Dim NombreCliente As String
    Dim FechaFacturacion As Long
    
    Dim File As Integer


MsgBox grdFacturaCustodia.SelBookmarks.Count
    
    If chkPasarTodas.value = 1 Then
    RsFactura.MoveFirst
    Do While Not RsFactura.EOF
        FacturaABC = RsFactura!FacturaABC
        NumeroFactura = RsFactura!NumeroFactura
        IDCliente = RsFactura!IDCliente
        Cuit = RsFactura!Cuit
        Subtotal = RsFactura!Subtotal
        IVAInscripto = RsFactura!IVAInscripto
        TotalFacturado = RsFactura!TotalFacturado
        MesFacturacion = Mid(txtFechaDatas.Text, 4, 2)
        AnoFacturacion = Mid(txtFechaDatas.Text, 7, 4)
        NombreCliente = RsFactura!NombreCliente
        FechaFacturacion = RsFactura!FechaFacturacion
        InsertarFacturaElectronica NumeroFactura, IDCliente, FacturaABC, Cuit, Subtotal, IVAInscripto, TotalFacturado, MesFacturacion, AnoFacturacion, NombreCliente, FechaFacturacion
        RsFactura.MoveNext
    Loop
    
    
    Else
    
    
    For i = 0 To grdFacturaCustodia.SelBookmarks.Count - 1
        
        File = grdFacturaCustodia.SelBookmarks.Item(i)
        
        grdFacturaCustodia.Row = File - 1
        grdFacturaCustodia.Col = 0
        FacturaABC = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 1
        NumeroFactura = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 2
        IDCliente = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 3
        Cuit = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 4
        Subtotal = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 5
        IVAInscripto = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 6
        TotalFacturado = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 7
        MesFacturacion = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 8
        AnoFacturacion = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 9
        NombreCliente = grdFacturaCustodia.Text
        grdFacturaCustodia.Col = 10
        FechaFacturacion = grdFacturaCustodia.Text
        Rem ActualizarFacturasCustodia FacturaABC, NumeroFactura
        InsertarFacturaElectronica NumeroFactura, IDCliente, FacturaABC, Cuit, Subtotal, IVAInscripto, TotalFacturado, MesFacturacion, AnoFacturacion, NombreCliente, FechaFacturacion
    Next

End If
 
MsgBox "Terminado"

End Sub

Private Sub Command5_Click()
'Dim rs As New ADODB.Recordset
'Dim Sql As String
'Sql = " SELECT     ID_CLIENTE, RAZON_SOCIAL, NOFACTURAR"
'Sql = Sql & " From Clientes"
'Sql = Sql & "  Where NOFACTURAR IS NULL "
'Sql = Sql & " Order BY ID_CLIENTE "
'
'rs.Open Sql, strConBasa , 0 ,1
'
'Do While Not rs.EOF
'
'COD_CLIENTE = rs!id_cliente
'FECHA_INICIO = "01/07/2008"
'FECHA_FIN = "31/07/2008"
'MES_SERVICIO = 6
'
'CAJAS_CANTIDAD
'CAJAS_MES
'CAJAS_VACIAS
'Consultas
'CARGA_LEGAJOS
'FLETES_NORMALES
'FLETES_URGENTES
'REARCHIVO_FISICO
'IMAGENES
'rs.MoveNext
'
'Loop


End Sub

Private Sub Command7_Click()
     Dim rsFacturacion As New ADODB.Recordset
    Dim rsFletes As New ADODB.Recordset
    Dim RsDigital As New ADODB.Recordset
    Dim rsMaxRemito As New ADODB.Recordset
    Dim Sql As String
    Dim Cantidades As Long
    Dim RemitosVacias As String
    Dim CantidadImagen As Integer
    Dim RemitosCrecimiento As String
    Dim RemitosOrden As String
    Dim RemitosImagenes As String
    Dim RemitosBajas As String
    
    CantidadImagen = 0
On Error GoTo salir:

Dim R As Integer
Dim C As Integer
        Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
        
        Set ApExcel = New Excel.Application
      Rem    Set libroEx = Excel.Workbooks.Open(strPasoPlanillas & "factura.xls")
         Set libroEx = Excel.Workbooks.Open("D:\fa.xls")
        Set hojaEx = libroEx.Worksheets.Item(1)
    
    Dim Filtro As String
   Filtro = InputBox("Ingrese los Nº de clientes separados por ," & vbCrLf & "Para todos los clientes 0", "Filtro Cliente", 0)
    Sql = " SELECT     COD_CLIENTE_CABECERA , ID_CLIENTE, RAZON_SOCIAL, NOFACTURAR ,DETALLE_FACTURACION"
    Sql = Sql & " From Clientes "
    Sql = Sql & " where NOFACTURAR is null  "
    If Filtro <> "0" Then
        Sql = Sql & " AND ID_CLIENTE In ( " & Filtro & ")"
    End If
    Sql = Sql & " Order by  ID_CLIENTE "
    
    
rsFacturacion.Open Sql, ConActiva, 0, 1

    FECHA_INICIO = mskFecha_Desde.Text
    FECHA_FIN = mskFecha_Hasta.Text
    
    
    Sql = " SELECT    NRO_REMITO, FECHA, ID_CLIENTE , COD_FLETE"
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO"
    Sql = Sql & vbCrLf & " Where (Cod_Flete Is Null)"
    Sql = Sql & vbCrLf & " And ANULADO Is Null"
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & "  AND (REMITOS_CUERPO.TIPO = 1) AND"
    Sql = Sql & vbCrLf & "  (REMITOS_CUERPO.OPERACION = 1) "
    
   
    
    
     Set rsFletes = New ADODB.Recordset
    
    Dim flete As String
    rsFletes.Open Sql, ConActiva, 0, 1
    
    If Not rsFletes.EOF Then
                    flete = flete & " ; " & rsFletes!Cod_Flete
          rsFletes.MoveNext
    Else
       MsgBox "Atencion  faltan procesar los fletes los datos seran copiados a memoria" & vbCrLf & flete, vbCritical
        flete = "NRO_REMITO" & vbTab & "FECHA" & vbTab & "ID_CLIENTE"
        Do While Not rsFletes.EOF
                    
            
             flete = flete & " vbcrlf  " & rsFletes!NRO_REMITO & vbTab & rsFletes!fecha & vbTab & rsFletes!id_cliente
       rsFletes.MoveNext
        Loop
        Clipboard.Clear
        Clipboard.SetText flete
        
     If MsgBox("Usted Quiere continuar sin procesar los fletes", vbYesNo) = vbNo Then
     Exit Sub
     
     Else
     End If
     
    End If
    
    
    
    
    
       
        
Sql = "    SELECT     REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.ID_CLIENTE, REQUERIMIENTO.IDTIPOREQUERIMIENTO,"
Sql = Sql & vbCrLf & " REQUERIMIENTO.Cantidad_Imagenes , REQUERIMIENTO.FECHARECEPCION, Clientes.NOFACTURAR"
Sql = Sql & vbCrLf & " FROM         REQUERIMIENTO INNER JOIN"
Sql = Sql & vbCrLf & " CLIENTES ON REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE"
Sql = Sql & vbCrLf & " WHERE     (REQUERIMIENTO.IDTIPOREQUERIMIENTO IN (13, 14)) AND (REQUERIMIENTO.CANTIDAD_IMAGENES IS NULL) "
 If Filtro <> "0" Then
         Sql = Sql & " AND REQUERIMIENTO.ID_CLIENTE In ( " & Filtro & ")"
         Else
         Sql = Sql & vbCrLf & " AND ( CLIENTES.NOFACTURAR IS NULL)"
        End If

Sql = Sql & vbCrLf & " AND (REQUERIMIENTO.ANULADO IS NULL) AND REQUERIMIENTO.FECHARECEPCION BETWEEN " & FechaFormato("01/09/2010") & " AND " & FechaFormato("30/09/2010")

        Sql = Sql & vbCrLf & "  AND FECHARECEPCION BETWEEN " & FechaFormato(FECHA_INICIO)
        Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
        
        Set RsDigital = New ADODB.Recordset
        
        RsDigital.Open Sql, ConActiva, 0, 1
        Dim ErrorDigital As String
        
        ErrorDigital = ""
        
        
         Do While Not RsDigital.EOF
            ErrorDigital = ErrorDigital & " " & RsDigital!IDREQUERIMIENTO
            RsDigital.MoveNext
         Loop
         
         If ErrorDigital <> "" Then
            MsgBox "Falta procesar los requerimientos digitales Nº :" & ErrorDigital
            Exit Sub
         End If
    
        
        
    
    If DateDiff("d", mskFecha_Hasta.Text, Now) < 1 Then
        MsgBox "Atención NO se puede procesar por que pueden existir errores" & vbCrLf & "Por favor Modifique la Fecha Hasta", vbCritical
        Exit Sub
    End If
    
    Sql = " SELECT     MAX(NRO_REMITO) AS MaxRemito From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " WHERE FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    
    Set rsMaxRemito = New ADODB.Recordset
    
    
    rsMaxRemito.Open Sql, ConActiva, 0, 1
    

  
  hojaEx.Cells(2, 3) = "Movimientos del " & FECHA_INICIO & " hasta " & FECHA_FIN
hojaEx.Cells(3, 3) = " Ultimo Remito " & rsMaxRemito!MaxRemito

R = 7
 Do While Not rsFacturacion.EOF
 If IsNull(rsFacturacion!COD_CLIENTE_CABECERA) Then
 hojaEx.Cells(R, 1) = 0
 Else
    hojaEx.Cells(R, 1) = rsFacturacion!COD_CLIENTE_CABECERA
    End If
    hojaEx.Cells(R, 2) = rsFacturacion!id_cliente
    hojaEx.Cells(R, 3) = Mid(Trim(rsFacturacion!RAZON_SOCIAL), 1, 30)
     If rsFacturacion!id_cliente = 39 Then
        hojaEx.Cells(R, 4) = 16955
     Else
        hojaEx.Cells(R, 4) = CAJAS_CANTIDAD(rsFacturacion!id_cliente)
     End If
    
    hojaEx.Cells(R, 5) = " Cant.: " & CAJAS_CRECIMIENTO_MES(rsFacturacion!id_cliente, RemitosCrecimiento)
    If RemitosCrecimiento <> "" Then
        hojaEx.Cells(R, 5) = hojaEx.Cells(R, 5) & " \ " & " Rem. Man:" & " \ " & Trim(RemitosCrecimiento)
    End If
    
    hojaEx.Cells(R, 6) = "Vacias: " & CAJAS_VACIAS(rsFacturacion!id_cliente, RemitosVacias)
    If RemitosVacias <> "" Then
        hojaEx.Cells(R, 6) = hojaEx.Cells(R, 6) & " \ " & "Rem. Sis:" & " \ " & Trim(RemitosVacias)
    End If
        RemitosBajas = ""
        hojaEx.Cells(R, 7) = "Cant:" & BajasMensuales(rsFacturacion!id_cliente, RemitosBajas)
     
     
     If RemitosBajas <> "" Then
        hojaEx.Cells(R, 7) = hojaEx.Cells(R, 7) & vbCrLf & "Rem.:" & " \ " & Trim(RemitosBajas)
     End If
     hojaEx.Cells(R, 8) = RECAMBIO_CAJAS(rsFacturacion!id_cliente)
    hojaEx.Cells(R, 9) = LIBROS_CANTIDAD(rsFacturacion!id_cliente)
    hojaEx.Cells(R, 10) = LEGAJOS_CANTIDAD(rsFacturacion!id_cliente)
    hojaEx.Cells(R, 11) = LEGAJOS_CARGA(rsFacturacion!id_cliente)
    If rsFacturacion!COD_CLIENTE_CABECERA = 34 Then
        hojaEx.Cells(R, 12) = ConsultasCajasLegajos(rsFacturacion!id_cliente) & " ConsultaPlanta " & CONSULTAS_EN_PLANTA(rsFacturacion!id_cliente, "")
    Else
       Rem  hojaEx.Cells(R, 12) = Consultas(rsFacturacion!id_cliente) + CONSULTAS_EN_PLANTA(rsFacturacion!id_cliente, "")
    End If
    hojaEx.Cells(R, 13) = CONSULTAS_DIGITALES(rsFacturacion!id_cliente, CantidadImagen)
    hojaEx.Cells(R, 14) = CantidadImagen
    hojaEx.Cells(R, 15) = FLETES_NORMALES(rsFacturacion!id_cliente)
    hojaEx.Cells(R, 16) = FLETES_URGENTES(rsFacturacion!id_cliente)
    hojaEx.Cells(R, 17) = "Cant:" & ORDEN_DOCUMENTACION(rsFacturacion!id_cliente, RemitosOrden, FISICO)
    If RemitosOrden <> "" Then
            hojaEx.Cells(R, 17) = hojaEx.Cells(R, 17) & " \ " & "Rem. Man:" & " \ " & Trim(RemitosOrden)
    End If
    
    hojaEx.Cells(R, 18) = "Cant.:" & ORDEN_DOCUMENTACION(rsFacturacion!id_cliente, RemitosOrden, lote)
    If RemitosOrden <> "" Then
        hojaEx.Cells(R, 18) = hojaEx.Cells(R, 18) & " \ " & " Rem. Man: " & " \ " & Trim(RemitosOrden)
    End If
    
    hojaEx.Cells(R, 19) = "Cant:." & IMAGENES(rsFacturacion!id_cliente, RemitosImagenes)
    If RemitosImagenes <> "" Then
        hojaEx.Cells(R, 19) = hojaEx.Cells(R, 19) & " \ " & "Rem. Man:" & " \ " & Trim(RemitosImagenes)
    End If
    
    
    hojaEx.Cells(R, 22) = Precintos(rsFacturacion!id_cliente, mskFecha_Desde, mskFecha_Hasta)
    
    
    
    
       Rem  hojaEx.Cells(R, 18) = rsFacturacion!DETALLE_FACTURACION
R = R + 1

    rsFacturacion.MoveNext
 Loop
hojaEx.SaveAs "C:\Factura\Movimientos desde " & Format(FECHA_INICIO, "dd_MM_yy") & " Hasta " & Format(FECHA_FIN, "DD_MM_YY") & " Fecha de emision " & Format(Now, "DD_MM_YY  HH mm") & ".xls"
 libroEx.Close
            ApExcel.Quit
 MsgBox "Operacion terminada"
      Exit Sub
        
salir:
     
     MsgBox "Error " & Err.Description
        
        
End Sub

Private Sub Command8_Click()
 
    
    

End Sub

Private Sub ctlClienteFactura_LostFocus()

   Dim Sql As String
   
   If IsNull(ctlClienteFactura.Valor) Then
    Exit Sub
    
   End If
   
   
   
   LimpiarCampos
   If chkNoBorraGrilla.value = 1 Then
   Else
   
    grdFacturacion.Clear
    grdFacturacion.Rows = 2
    TitulosGrilla
    End If
    
    
   
   Sql = " SELECT TIPO_FACTURA, DETALLE_FACTURACION,"
   Sql = Sql & " NRO_CUIT, PERIODO_FACTURA "
    Sql = Sql & "  From Clientes "
    Sql = Sql & "  Where id_cliente = " & ctlClienteFactura.Valor
    Dim rs As New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    If rs.EOF Then
        MsgBox "No existe el cliente"
        Exit Sub
    End If
    
  Rem  lblTipoFactura.Caption = rs!Tipo_Factura
    
    If IsNull(rs!DETALLE_FACTURACION) Then
        txtDetalleFacturacion.Text = ""
    Else
        txtDetalleFacturacion.Text = rs!DETALLE_FACTURACION
    End If
    
    If IsNull(rs!PERIODO_FACTURA) Then
        lblPeriodoAnteriorFacturado.Caption = ""
    Else
        lblPeriodoAnteriorFacturado.Caption = rs!PERIODO_FACTURA
    End If
    
    
    Sql = "  SELECT MAX(NUMERO_FACTURA) AS MaxFactura"
    Sql = Sql & " From FACTURAS"
    Sql = Sql & " WHERE TIPO_FACTURA = '" & lblTipoFactura.Caption & "'"
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
   Rem txtNumeroFactura.Text = rs!maxFactura + 1
    txtFechaFactura.Text = Format(Now, "DD/MM/YYYY")
   
  
End Sub

Private Sub Form_Load()
   ctlClienteFactura.TipoControl = 0
    ctlPersonalRecibo.TipoControl = 1
   
   Rem cboTipoComprobante.ListIndex = 0
    mskFecha_Hasta.Text = DateAdd("D", -2, Format(Now, "DD/MM/YYYY"))
    TitulosGrilla
    TitulosGrillaRecibo
    txtFechaFlete.Text = DateAdd("D", -30, Format(Now, "DD/MM/YYYY"))
   Dim rs As New ADODB.Recordset
    mskFecha_Desde.Text = DateAdd("D", -30, Format(Now, "DD/MM/YYYY"))
    mskFecha_Hasta.Text = DateAdd("D", 29, mskFecha_Desde.Text)
    txtMesFacturacion.Text = Format(Now, "MM")
    txtFechaDatas.Text = Format(Now, "DD/MM/YYYY")

End Sub

Private Sub grdFacturacion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
RowGrilla = grdFacturacion.Row
 PopupMenu mnuGrilla
  
End If

End Sub

Private Sub grdFletes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
    If grdFletes.Col = 2 Then
       RemitoFlete = grdFletes.Text
       RowBo = grdFletes.Bookmark
       PopupMenu mnuFlete
       
    End If
 End If
 
End Sub

Private Sub grdReciboFacuta_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuGrillarecibos
End If

End Sub

Private Sub lbl_FletesNormales_Click()
 txtCantidad.Text = lbl_FletesNormales.Caption
    txtCodigo.Text = "FN"
    txtCodigo.SetFocus
End Sub

Private Sub lbl_FletesUrgentes_Click()
 txtCantidad.Text = lbl_FletesUrgentes.Caption
    txtCodigo.Text = "FU"
    txtCodigo.SetFocus
End Sub

Private Sub lblCajasCrecimientoMes_Click()
    txtCantidad.Text = lblCajasCrecimientoMes.Caption
    txtCodigo.Text = "RE"
    txtCodigo.SetFocus
End Sub

Private Sub lblCantidad_Cajas_Click()
    txtCantidad.Text = lblCantidad_Cajas.Caption
    txtCodigo.Text = "CC"
    txtCodigo.SetFocus
    
End Sub

Private Sub lblCantidad_Legajos_Click()
txtCantidad.Text = lblCantidad_Legajos.Caption
    txtCodigo.Text = "CLG"
    txtCodigo.SetFocus
End Sub

Private Sub lblCantidad_Libros_Click()
    txtCantidad.Text = lblCantidad_Libros.Caption
    txtCodigo.Text = "CL"
    txtCodigo.SetFocus
End Sub

Private Sub lblCantidadCajasVacias_Click()
    txtCantidad.Text = lblCantidadCajasVacias.Caption
    txtCodigo.Text = "CA"
    txtCodigo.SetFocus
End Sub

Private Sub lblCantidadDesarchivos_Click()
    txtCantidad.Text = lblCantidadDesarchivos.Caption
    txtCodigo.Text = "CO"
    txtCodigo.SetFocus
End Sub

Private Sub mnuBorrarTodo_Click()
    grdFacturacion.Clear
    grdFacturacion.Rows = 2
    TitulosGrilla
End Sub

Private Sub mnuControlfacturacion_Click()
 Dim rs As New ADODB.Recordset
 Dim rsFac As ADODB.Recordset
 Dim Sql As String
 Dim Mes As String
 Dim DatoFac As String
 Dim DatoCliente As String
 Dim CantFactura As Integer
Sql = " SELECT ID_CLIENTE, RAZON_SOCIAL"
Sql = Sql & " From Clientes "
Sql = Sql & "  ORDER BY ID_CLIENTE"
rs.Open Sql, ConActiva, 0, 1
 Mes = InputBox("Ingrese el mes de Facturacion")
           DatoCliente = "Nº " & vbTab & "Razon Social" & vbTab & "Fecha" & vbTab & " Factura " & vbTab & "Monto" & vbCrLf
 Do While Not rs.EOF
        Sql = "  SELECT FACTURAS.FECHA, FACTURAS.NUMERO_FACTURA, "
        Sql = Sql & vbCrLf & " FACTURAS.MONTO_CON_IVA, FACTURAS.TIPO_FACTURA,"
        Sql = Sql & vbCrLf & " FACTURA_ESTADO.DESCRIPCION AS DesEstado"
        Sql = Sql & vbCrLf & " From FACTURAS, FACTURA_ESTADO"
        Sql = Sql & vbCrLf & " WHERE FACTURAS.ESTADO = FACTURA_ESTADO.ID_ESTADO "
        Sql = Sql & vbCrLf & " AND FACTURAS.COD_CLIENTE = " & rs!id_cliente
        Sql = Sql & vbCrLf & " AND FACTURAS.TIPO_COMPROBANTE = 'Factura'"
        Sql = Sql & vbCrLf & " AND FACTURAS.MESFACTURACION =" & Mes
        Sql = Sql & vbCrLf & " ORDER BY FACTURAS.ID_FACTURA DESC "
       
       DatoFac = ""
        Set rsFac = New ADODB.Recordset
            rsFac.Open Sql, ConActiva, 0, 1
            
            CantFactura = 1
            If rsFac.EOF Then
            DatoCliente = DatoCliente & rs!id_cliente & vbTab & Mid(Trim(rs!RAZON_SOCIAL), 1, 50) & vbTab & vbTab & vbTab & vbTab & "NO FACTURADO" & vbCrLf
                
            Else
                Do While Not rsFac.EOF
            
                  If CantFactura = 1 Then
                    DatoFac = DatoFac & rsFac!fecha & vbTab & rsFac!Tipo_Factura & " \ " & rsFac!NUMERO_FACTURA & vbTab & rsFac!MONTO_CON_IVA & vbTab & rsFac!DesEstado
                  Else
                    DatoFac = DatoFac & vbCrLf & vbTab & vbTab & rsFac!fecha & vbTab & rsFac!Tipo_Factura & " \ " & rsFac!NUMERO_FACTURA & vbTab & rsFac!MONTO_CON_IVA & vbTab & rsFac!DesEstado
                  End If
                  CantFactura = CantFactura + 1
                    rsFac.MoveNext
                Loop
            
            DatoCliente = DatoCliente & rs!id_cliente & vbTab & Mid(Trim(rs!RAZON_SOCIAL), 1, 50) & vbTab & DatoFac & vbCrLf
            
            End If

           
            
    
    rs.MoveNext
 Loop
 
 Clipboard.Clear
 Clipboard.SetText DatoCliente
 
 
 
End Sub

Private Sub mnuControlRetenciones_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    rs.CursorLocation = adUseClient

    Sql = " SELECT RECIBOS.ID_RECIBO, RECIBOS.COD_CLIENTE,"
    Sql = Sql & vbCrLf & "    CLIENTES.RAZON_SOCIAL, RECIBOS.FECHA,"
    Sql = Sql & vbCrLf & " RECIBOS.NUMERO_RECIBO, RECIBOS.ESTADO_RECIBO,"
    Sql = Sql & vbCrLf & " RECIBOS.RETENCION_GANANCIAS AS GANANCIAS,"
    Sql = Sql & vbCrLf & " RECIBOS.RETENCION_IVA AS RETIVA,"
    Sql = Sql & vbCrLf & " RECIBOS.RENTENCION_INGRESOS_BRUTOS AS RETIB,"
    Sql = Sql & vbCrLf & " RECIBOS.RENTENCION_SUSS AS RETSUSS"
    Sql = Sql & vbCrLf & " From RECIBOS, Clientes"
    Sql = Sql & vbCrLf & " WHERE RECIBOS.COD_CLIENTE = CLIENTES.ID_CLIENTE AND"
    Sql = Sql & vbCrLf & " ((RECIBOS.ESTADO_RECIBO = 100) AND"
    Sql = Sql & vbCrLf & " (RECIBOS.RETENCION_GANANCIAS <> 0) OR"
    Sql = Sql & vbCrLf & "(RECIBOS.RETENCION_IVA <> 0) OR"
    Sql = Sql & vbCrLf & " (RECIBOS.RENTENCION_INGRESOS_BRUTOS <> 0) OR"
    Sql = Sql & vbCrLf & " (RECIBOS.RENTENCION_SUSS <> 0))"
    Sql = Sql & vbCrLf & " ORDER BY RECIBOS.ID_RECIBO"


rs.Open Sql, ConActiva, 0, 1

frmInforme.CargarInforme "Retenciones", rs
frmInforme.Show
End Sub

Private Sub mnuFacturasCobrar_Click()
 
Dim rs As New ADODB.Recordset
Dim Sql As String
    rs.CursorLocation = adUseClient
    Sql = "   SELECT FACTURAS.FECHA_PAGO,"
    Sql = Sql & vbCrLf & "    FACTURA_ESTADO.DESCRIPCION, CLIENTES.RAZON_SOCIAL,"
    Sql = Sql & vbCrLf & "   FACTURAS.TIPO_FACTURA, FACTURAS.NUMERO_FACTURA,"
    Sql = Sql & vbCrLf & "   FACTURAS.MONTO_CON_IVA"
    Sql = Sql & vbCrLf & " From FACTURAS, Clientes, FACTURA_ESTADO"
    Sql = Sql & vbCrLf & " WHERE FACTURAS.COD_CLIENTE = CLIENTES.ID_CLIENTE AND"
    Sql = Sql & vbCrLf & "    FACTURAS.ESTADO = FACTURA_ESTADO.ID_ESTADO AND"
    Sql = Sql & vbCrLf & "    (FACTURAS.ESTADO = 40)"
    Sql = Sql & vbCrLf & " ORDER BY FACTURAS.FECHA_PAGO"
    
    
    Sql = "  SELECT FACTURAS.FECHA_PAGO, FACTURAS.DETALLE_PAGO, "
    Sql = Sql & vbCrLf & "  FACTURA_ESTADO.DESCRIPCION, CLIENTES.RAZON_SOCIAL, "
    Sql = Sql & vbCrLf & "  FACTURAS.TIPO_FACTURA, FACTURAS.NUMERO_FACTURA, "
    Sql = Sql & vbCrLf & "  FACTURAS.MONTO_CON_IVA, "
    Sql = Sql & vbCrLf & "  FACTURA_CONTACTO.COBRANZA_TIPO "
    Sql = Sql & vbCrLf & "  FROM FACTURAS, CLIENTES, FACTURA_ESTADO, "
    Sql = Sql & vbCrLf & "  FACTURA_CONTACTO "
    Sql = Sql & vbCrLf & "  WHERE FACTURAS.COD_CLIENTE = CLIENTES.ID_CLIENTE AND "
    Sql = Sql & vbCrLf & "  FACTURAS.ESTADO = FACTURA_ESTADO.ID_ESTADO AND "
    Sql = Sql & vbCrLf & "  FACTURAS.COD_CLIENTE = FACTURA_CONTACTO.COD_CLIENTE "
    Sql = Sql & vbCrLf & "  AND (FACTURAS.ESTADO = 40) "
    Sql = Sql & vbCrLf & " ORDER BY FACTURAS.FECHA_PAGO "
     
    rs.Open Sql, ConActiva, 0, 1
    frmInforme.CargarInforme "Factura", rs
    frmInforme.Show

End Sub

Private Sub mnuFleteSinCosto_Click()
    Dim Sql As String
    If RemitoFlete <> 0 Then
       Sql = " Update REMITOS_CUERPO Set COD_FLETE = 0 Where NRO_REMITO = " & RemitoFlete
       ExecutarSql Sql
    End If
    CargarFletes
    RemitoFlete = 0
End Sub

Private Sub mnuFleteUnificado_Click()
 Dim flete As Long
 Dim Sql As String
If RemitoFlete <> 0 Then
    flete = InputBox("Ingrese el numero de flete", "Flete Unificado")
    Sql = " Update REMITOS_CUERPO Set COD_FLETE =" & flete & "  Where NRO_REMITO = " & RemitoFlete
    ExecutarSql Sql
    CargarFletes
End If

End Sub

Private Sub mnuGrillaBorrar_Click()
 grdFacturacion.RemoveItem RowGrilla
 Sumar
End Sub

Private Sub mnuGrillaModificar_Click()
Dim DATO As String
DATO = InputBox("Ingrese el nuevo valor", , Trim(grdFacturacion.TextMatrix(grdFacturacion.RowSel, grdFacturacion.ColSel)))
grdFacturacion.TextMatrix(grdFacturacion.RowSel, grdFacturacion.ColSel) = DATO
Sumar
End Sub

Private Sub mnuInformeFacturacion_Click()
    Dim Sql As String
   

End Sub

Private Sub mnuInformePorCliente_Click()



    Dim Sql As String
    
    
    
Sql = "    SELECT V_FACTURA.COD_FACTURA, V_FACTURA.ID_CLIENTE, "
Sql = Sql & vbCrLf & "   V_FACTURA.RAZON_SOCIAL, V_FACTURA.TIPO_COMPROBANTE,"
Sql = Sql & vbCrLf & "   V_FACTURA.NUMERO_FACTURA, V_FACTURA.TIPO_FACTURA,"
Sql = Sql & vbCrLf & "   V_FACTURA.FECHA, V_FACTURA.MONTO_SIN_IVA,"
Sql = Sql & vbCrLf & "   V_FACTURA.MONTO_CON_IVA, V_FACTURA.IVA,"
Sql = Sql & vbCrLf & "   V_FACTURA.ITEM, V_FACTURA.CANTIDAD, V_FACTURA.DETALLE,"
Sql = Sql & vbCrLf & "  V_FACTURA.IMPORTE_UNITARIO, V_FACTURA.IMPORTE_TOTAL,"
Sql = Sql & vbCrLf & "  V_FACTURA.DESCRIPCION , V_FACTURA.Estado"
Sql = Sql & vbCrLf & "   FROM   V_FACTURA "
Sql = Sql & vbCrLf & "   Where V_FACTURA.id_cliente = " & ctlClienteFactura.Valor
Sql = Sql & vbCrLf & "   ORDER BY V_FACTURA.NUMERO_FACTURA "
    frmReportes.ImprimirReporte PasoReportes & "rptInformeDetalleFactura.rpt", Sql, True
    
End Sub

Private Sub mnuListadoReciboPorFecha_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
    rs.CursorLocation = adUseClient
    
       Sql = "   SELECT PERSONAL.NOMBRE, PERSONAL.APELLIDO,"
      Sql = Sql & vbCrLf & "    RECIBOS.ID_RECIBO, RECIBOS.NUMERO_RECIBO,"
     Sql = Sql & vbCrLf & "     RECIBOS.TIPO_PAGO, RECIBOS.FECHA,"
       Sql = Sql & vbCrLf & "   RECIBOS.ESTADO_RECIBO"
   Sql = Sql & vbCrLf & "   From PERSONAL, RECIBOS"
   Sql = Sql & vbCrLf & "   Where PERSONAL.IDPERSONAL = RECIBOS.Responsable"
   
   Sql = Sql & vbCrLf & "   ORDER BY RECIBOS.NUMERO_RECIBO"
    
rs.Open Sql, ConActiva, 0, 1

frmInforme.CargarInforme "Factura", rs
frmInforme.Show
End Sub

Private Sub mnuListadoCobranzas_Click()
 Dim Sql As String

 Sql = " SELECT V_COBRANZA.COD_CLIENTE, V_COBRANZA.RAZON_SOCIAL,"
 Sql = Sql & vbCrLf & " V_COBRANZA.TIPO_COMPROBANTE, "
 Sql = Sql & vbCrLf & "  V_COBRANZA.COBRANZA_TELEFONOS , V_COBRANZA.TIPO_FACTURA,"
 Sql = Sql & vbCrLf & "  V_COBRANZA.NUMERO_FACTURA, V_COBRANZA.Fecha,"
 Sql = Sql & vbCrLf & "  V_COBRANZA.ESTADO_DESCRIPCION, V_COBRANZA.MONTO_CON_IVA, "
 Sql = Sql & vbCrLf & "  V_COBRANZA.FACTURA_TELEFONOS, V_COBRANZA.COBRANZA_TIPO,"
 Sql = Sql & vbCrLf & "  V_COBRANZA.FECHA_PAGO,V_COBRANZA.DETALLE_PAGO"
 Sql = Sql & vbCrLf & "  FROM   BASA.V_COBRANZA V_COBRANZA"
 Sql = Sql & vbCrLf & "  ORDER BY V_COBRANZA.COD_CLIENTE, V_COBRANZA.FECHA"
 frmReportes.ImprimirReporte PasoReportes & "rptcobranza.rpt", Sql, True
 
End Sub

Private Sub mnuListadoReciboPorFechaCarga_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String
    Dim FECHA_DESDE As String
    Dim FECHA_HASTA As String
    rs.CursorLocation = adUseClient
    FECHA_DESDE = InputBox("Ingrese la fecha Desde")
    FECHA_HASTA = InputBox("Ingrese la fecha Hasta")

    
    Sql = " SELECT PERSONAL.APELLIDO, FECHA_CARGA,"
   Sql = Sql & vbCrLf & "  RECIBOS.NUMERO_RECIBO, RECIBOS.TIPO_PAGO,"
    Sql = Sql & vbCrLf & "  RECIBOS.FECHA, RECIBOS.MONTO_TOTAL,"
    Sql = Sql & vbCrLf & "  RECIBOS.NUMERO_RESPALDO , RECIBOS.BANCO"
Sql = Sql & vbCrLf & "  From RECIBOS, PERSONAL"
Sql = Sql & vbCrLf & "  Where RECIBOS.Responsable = PERSONAL.IDPERSONAL"
    
    
    Sql = Sql & vbCrLf & " AND FECHA_CARGA "
    Sql = Sql & vbCrLf & " BETWEEN '" & FECHA_DESDE & "'"
    Sql = Sql & vbCrLf & " AND '" & FECHA_HASTA & "'"
    Sql = Sql & vbCrLf & " ORDER BY FECHA_CARGA, TIPO_PAGO "
    rs.Open Sql, ConActiva, 0, 1

frmInforme.CargarInforme "Factura", rs
frmInforme.Show
End Sub

Private Sub mnuListadoRecibos_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
        rs.CursorLocation = adUseClient
        Sql = "   SELECT PERSONAL.NOMBRE, PERSONAL.APELLIDO,"
        Sql = Sql & vbCrLf & "    RECIBOS.ID_RECIBO, RECIBOS.NUMERO_RECIBO,"
        Sql = Sql & vbCrLf & "    RECIBOS.TIPO_PAGO, RECIBOS.FECHA,"
        Sql = Sql & vbCrLf & "    RECIBOS.ESTADO_RECIBO, RECIBOS.NUMERO_RESPALDO,"
        Sql = Sql & vbCrLf & "    RECIBOS.BANCO, RECIBOS.MONTO_TOTAL,"
        Sql = Sql & vbCrLf & "    FACTURAS.TIPO_FACTURA, FACTURAS.NUMERO_FACTURA,"
        Sql = Sql & vbCrLf & "    Clientes.RAZON_SOCIAL"
        Sql = Sql & vbCrLf & "    From Clientes, FACTURAS, PERSONAL, Recibos"
        Sql = Sql & vbCrLf & "    WHERE CLIENTES.ID_CLIENTE (+) = FACTURAS.COD_CLIENTE AND"
        Sql = Sql & vbCrLf & "    FACTURAS.COD_RECIBO (+) = RECIBOS.ID_RECIBO AND"
        Sql = Sql & vbCrLf & "    PERSONAL.IDPERSONAL (+) = RECIBOS.RESPONSABLE"
        Sql = Sql & vbCrLf & "    ORDER BY RECIBOS.NUMERO_RECIBO"
        rs.Open Sql, ConActiva, 0, 1
        frmInforme.CargarInforme "Factura", rs
        frmInforme.Show
End Sub

Private Sub mnuNuevoFlete_Click()
Dim rsMax As ADODB.Recordset
Dim MaxFlete As Long
Dim Sql As String
Set rsMax = New ADODB.Recordset


rsMax.Open " SELECT MAX(COD_FLETE) as MaxFlete From REMITOS_CUERPO ", ConActiva, 0, 1
MaxFlete = rsMax!MaxFlete + 1

If RemitoFlete <> 0 Then
    Sql = " Update REMITOS_CUERPO Set COD_FLETE =" & MaxFlete & "  Where NRO_REMITO = " & RemitoFlete
    ExecutarSql Sql
End If
CargarFletes

End Sub



Private Sub mnuOrdenadoFactura_Click()
 Dim rs As New ADODB.Recordset
 Dim Sql As String
    rs.CursorLocation = adUseClient
        Sql = " SELECT FACTURAS.MESFACTURACION, CLIENTES.ID_CLIENTE,"
        Sql = Sql & vbCrLf & "     CLIENTES.RAZON_SOCIAL, FACTURAS.TIPO_COMPROBANTE,"
        Sql = Sql & vbCrLf & "      FACTURAS.TIPO_FACTURA, FACTURAS.NUMERO_FACTURA,"
        Sql = Sql & vbCrLf & "     FACTURAS.MONTO_SIN_IVA, FACTURAS.MONTO_CON_IVA,"
        Sql = Sql & vbCrLf & "     FACTURA_ESTADO.DESCRIPCION"
        Sql = Sql & vbCrLf & "   From FACTURAS, FACTURA_ESTADO, Clientes"
        Sql = Sql & vbCrLf & "   WHERE FACTURAS.ESTADO = FACTURA_ESTADO.ID_ESTADO AND"
        Sql = Sql & vbCrLf & "     FACTURAS.COD_CLIENTE = Clientes.id_cliente"
        Sql = Sql & vbCrLf & "   ORDER BY FACTURAS.TIPO_FACTURA,"
        Sql = Sql & vbCrLf & "     FACTURAS.NUMERO_FACTURA"
rs.Open Sql, ConActiva, 0, 1

frmInforme.CargarInforme "Factura", rs
frmInforme.Show
End Sub

Private Sub mnuOrdenadoPorCliente_Click()
 Dim rs As New ADODB.Recordset
    rs.CursorLocation = adUseClient
 Dim Sql As String

        Sql = " SELECT FACTURAS.MESFACTURACION, CLIENTES.ID_CLIENTE,"
        Sql = Sql & vbCrLf & "     CLIENTES.RAZON_SOCIAL, FACTURAS.TIPO_COMPROBANTE,"
        Sql = Sql & vbCrLf & "      FACTURAS.TIPO_FACTURA, FACTURAS.NUMERO_FACTURA,"
        Sql = Sql & vbCrLf & "     FACTURAS.MONTO_SIN_IVA, FACTURAS.MONTO_CON_IVA,"
        Sql = Sql & vbCrLf & "     FACTURA_ESTADO.DESCRIPCION"
        Sql = Sql & vbCrLf & "   From FACTURAS, FACTURA_ESTADO, Clientes"
        Sql = Sql & vbCrLf & "   WHERE FACTURAS.ESTADO = FACTURA_ESTADO.ID_ESTADO AND"
        Sql = Sql & vbCrLf & "     FACTURAS.COD_CLIENTE = Clientes.id_cliente AND "
        Sql = Sql & vbCrLf & "   FACTURAS.ESTADO > 1 "
        Sql = Sql & vbCrLf & "   ORDER BY CLIENTES.RAZON_SOCIAL"
rs.Open Sql, ConActiva, 0, 1

frmInforme.CargarInforme "Factura", rs
frmInforme.Show
End Sub

Private Sub SumarTotalReciboFactura()
Dim i As Integer
Dim Datos As Double

For i = 1 To grdReciboFacuta.Rows - 1
    Datos = Datos + CDbl(grdReciboFacuta.TextMatrix(i, 4))

Next
lblReciboTotalFacturas.Caption = Datos
End Sub

Private Sub mnuReciboBorrarTodo_Click()

    grdReciboFacuta.Clear
    grdReciboFacuta.Rows = 2
    TitulosGrillaRecibo
End Sub

Private Sub mnuReciboRendicion_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String
    Dim Recibos As String
    
    rs.CursorLocation = adUseClient
    Recibos = InputBox("Ingrese los recibos separos por ,")
    

    
    Sql = "  SELECT RECIBOS.COD_CLIENTE, CLIENTES.RAZON_SOCIAL,"
    Sql = Sql & vbCrLf & "   RECIBOS.TIPO_PAGO, RECIBOS.FECHA,"
    Sql = Sql & vbCrLf & "   RECIBOS.NUMERO_RECIBO,RECIBOS.MONTO_TOTAL, "
    Sql = Sql & vbCrLf & "   Recibos.NUMERO_RESPALDO"
    Sql = Sql & vbCrLf & "   From Recibos, Clientes"
    Sql = Sql & vbCrLf & "   WHERE RECIBOS.COD_CLIENTE = CLIENTES.ID_CLIENTE AND"
    Sql = Sql & vbCrLf & "   RECIBOS.NUMERO_RECIBO IN (" & Recibos & ")"
    Sql = Sql & vbCrLf & "   ORDER BY RECIBOS.TIPO_PAGO,RECIBOS.NUMERO_RECIBO"

    
    rs.Open Sql, ConActiva, 0, 1

frmInforme.CargarInforme "Factura", rs
frmInforme.Show
End Sub

Private Sub txtCantidad_LostFocus()
    Cantidad_Por_Unitario
End Sub
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
'CANON_CAJA CC
'CANON_LIBRO CL
'CANON_Legajos CLG
'Caja CA
'Referencia RE
'CARGAR_LEGAJOS CLE
'Consulta CO
'FLETE_NORMAL FN
'FLETE_URGENTE FU
'PRECINTO PR
'HORA_ARCHIVISTA_BASA HAB
'HORA_ARCHIVISTA_CLIENTE HAC
'ABONO_MINIMO AM
 Dim Tarifa As String



If KeyAscii = 13 Then
                Dim RsTarifa As New ADODB.Recordset
                Dim Sql As String
                Sql = " SELECT COD_CLIENTE, CANON_CAJA, CANON_LIBRO, CANON_LEGAJO, CAJA,  "
                Sql = Sql & vbCrLf & " REFERENCIA, CARGAR_LEGAJOS, CONSULTA, "
                Sql = Sql & vbCrLf & " FLETE_NORMAL, FLETE_URGENTE, PRECINTO, "
                Sql = Sql & vbCrLf & " HORA_ARCHIVISTA_BASA, HORA_ARCHIVISTA_CLIENTE, "
                Sql = Sql & vbCrLf & " ABONO_MINIMO "
                Sql = Sql & vbCrLf & " From TARIFAS_FACTURA "
                Sql = Sql & vbCrLf & " Where COD_CLIENTE = " & ctlClienteFactura.Valor
                RsTarifa.Open Sql, ConActiva, 0, 1
                If RsTarifa.EOF Then
                    MsgBox "El cliente No tiene tarifa", vbCritical
                    Exit Sub
                End If
                
                
                
                Tarifa = 0
                AbonoMinimo = 0
            Select Case Trim(UCase(txtCodigo.Text))
            Case "CC"
                txtDescripcion.Text = "Canon por guarda y custodia de cajas mes de " & txtPeriodoActual.Text
                Tarifa = RsTarifa!CANON_CAJA
                AbonoMinimo = RsTarifa!ABONO_MINIMO
            Case "CL"
                 txtDescripcion.Text = "Canon por guarda y custodia de libros mes de " & txtPeriodoActual.Text
                 Tarifa = RsTarifa!CANON_LIBRO
            Case "CLG"
                txtDescripcion.Text = "Canon por guarda y custodia de legajos mes de " & txtPeriodoActual.Text
                 Tarifa = RsTarifa!CANON_LEGAJO
            Case "CA"
                  txtDescripcion.Text = "Provisión de cajas mes de " & txtPeriodoActual.Text
                  Tarifa = RsTarifa!Caja
            Case "RE"
                 txtDescripcion.Text = "Alta y referencia en el sistema informático mes de " & txtPeriodoActual.Text
                 Tarifa = RsTarifa!Referencia
            Case "CLE"
                txtDescripcion.Text = "Carga de legajos mes de " & txtPeriodoActual.Text
                Tarifa = RsTarifa!CARGAR_LEGAJOS
            Case "CO"
                txtDescripcion.Text = "Desarchivo de cajas / libros / legajos mes de " & txtPeriodoActual.Text
                   
                Tarifa = RsTarifa!Consulta
            Case "FN"
                txtDescripcion.Text = "Fletes Normales"
                Tarifa = RsTarifa!FLETE_NORMAL
            Case "FU"
                txtDescripcion.Text = "Fletes Urgentes"
                Tarifa = RsTarifa!FLETE_URGENTE
            Case "PR"
                txtDescripcion.Text = "Precintos"
                Tarifa = RsTarifa!PRECINTO
            Case "HAB"
                txtDescripcion.Text = "Horas de archivista en planta"
                Tarifa = RsTarifa!HORA_ARCHIVISTA_BASA
            Case "HAC"
                txtDescripcion.Text = "Horas de archivista en el cliente"
                Tarifa = RsTarifa!HORA_ARCHIVISTA_CLIENTE
            Case "AM"
                txtDescripcion.Text = "Abono minimo mes de " & txtPeriodoActual.Text
                Tarifa = RsTarifa!ABONO_MINIMO

            End Select
            If lblTipoFactura.Caption = "B" Then
                 txtPrecioUnitario.Text = Format(Tarifa * 1.21, "#####.00")
                 If Mid(txtPrecioUnitario.Text, 1, 1) = "," Then
                  txtPrecioUnitario.Text = "0" & txtPrecioUnitario.Text
                 End If
                 
            Else
                txtPrecioUnitario.Text = Tarifa
            End If
            
            Cantidad_Por_Unitario
            If (txtCodigo.Text = "AM" And txtPrecioUnitario.Text = "") Then
            
            Else
            cmdInsertFacturacion.SetFocus
            End If
            

End If

End Sub


Public Sub Cantidad_Por_Unitario()
    If IsNumeric(txtPrecioUnitario.Text) And IsNumeric(txtCantidad.Text) Then
        txtTotal.Text = Format(txtCantidad.Text * txtPrecioUnitario, "#####.00")
        If CLng(txtTotal.Text) < AbonoMinimo And txtCodigo.Text = "CC" Then
           If MsgBox("Usted quiere colocar el abono minimo", vbYesNo) = vbYes Then
                    txtCodigo.Text = "AM"
                    txtTotal.Text = ""
                    txtCantidad.Text = 1
                    txtPrecioUnitario.Text = ""
                    txtDescripcion.Text = ""
                    txtCodigo.SetFocus
                    
           End If
           
        
        End If
    Else
        txtTotal.Text = ""
    End If
End Sub

Private Sub txtMontoRecibo_Change()

End Sub

Private Sub txtPrecioUnitario_LostFocus()
    Cantidad_Por_Unitario
End Sub

Public Sub Sumar()
Dim i As Integer
Dim Valor As Double



 

For i = 1 To grdFacturacion.Rows - 1
    grdFacturacion.TextMatrix(i, 5) = grdFacturacion.TextMatrix(i, 4) * grdFacturacion.TextMatrix(i, 1)
    grdFacturacion.TextMatrix(i, 0) = i
    Valor = Valor + CDbl(grdFacturacion.TextMatrix(i, 5))
 Next

 



If Trim(lblTipoFactura.Caption) = "A" Then
    lblSubTotal.Caption = Format(Valor, "#######.00")
    lblTotal.Caption = Format(Valor * 1.21, "#######.00")
    lblIVA.Caption = Format(lblTotal.Caption - lblSubTotal.Caption, "#######.00")
Else
    lblSubTotal.Caption = Format(Valor, "#######.00")
    lblTotal.Caption = Format(Valor, "#######.00")
    lblIVA.Caption = 0
End If

End Sub

Public Sub InsertDatoFactura()
If grdFacturacion.Rows = 2 And grdFacturacion.TextMatrix(1, 1) = "" Then
        grdFacturacion.TextMatrix(1, 0) = 1
        grdFacturacion.TextMatrix(1, 1) = txtCantidad.Text
        grdFacturacion.TextMatrix(1, 2) = txtCodigo.Text
        grdFacturacion.TextMatrix(1, 3) = txtDescripcion.Text
        grdFacturacion.TextMatrix(1, 4) = txtPrecioUnitario.Text
        grdFacturacion.TextMatrix(1, 5) = txtTotal.Text
       Else
        grdFacturacion.AddItem "0" & vbTab & txtCantidad.Text & vbTab & txtCodigo.Text & vbTab & txtDescripcion.Text & vbTab & txtPrecioUnitario & vbTab & txtTotal
    End If
    Dim Valor As Double
 Dim i As Integer
 
 
 Sumar
End Sub

Public Sub TitulosGrillaRecibo()
Dim i As Integer
With grdReciboFacuta

.Cols = 6
.ColWidth(0) = 600
.ColWidth(1) = 1000
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 2800
.ColAlignment(0) = flexAlignCenterCenter
.ColAlignment(1) = flexAlignCenterCenter
.ColAlignment(2) = flexAlignCenterCenter
.ColAlignment(3) = flexAlignCenterCenter
.ColAlignment(4) = flexAlignCenterCenter
.ColAlignment(5) = 0


.RowHeight(0) = 400

.TextMatrix(0, 0) = "Item"
.TextMatrix(0, 1) = "ID Factura"
.TextMatrix(0, 2) = "Tipo Factura"
.TextMatrix(0, 3) = "Nº Factura"
.TextMatrix(0, 4) = "Monto"
.TextMatrix(0, 5) = "Razon Social"
grdFacturacion.Row = 0
    .Col = 0
    .CellFontSize = 8
    .CellBackColor = &H80000013
    .CellFontBold = True



End With

End Sub

Public Sub TitulosGrilla()
Dim i As Integer
With grdFacturacion

.Cols = 6
.ColWidth(0) = 600
.ColWidth(1) = 700
.ColWidth(2) = 900
.ColWidth(3) = 6000
.ColWidth(4) = 1100
.ColWidth(5) = 700

.ColAlignment(0) = flexAlignCenterCenter
.ColAlignment(1) = flexAlignCenterCenter
.ColAlignment(2) = flexAlignCenterCenter
Rem .ColAlignment(3) = flexAlignCenterCenter
.ColAlignment(4) = flexAlignCenterCenter
.ColAlignment(5) = flexAlignCenterCenter

.RowHeight(0) = 400

grdFacturacion.TextMatrix(0, 0) = "ITEM"
grdFacturacion.TextMatrix(0, 1) = "CANT."
grdFacturacion.TextMatrix(0, 2) = "CODIGO"
grdFacturacion.TextMatrix(0, 3) = "DESCRIPCION"
grdFacturacion.TextMatrix(0, 4) = "PRECIO/U"
grdFacturacion.TextMatrix(0, 5) = "TOTAL"
grdFacturacion.Row = 0
For i = 0 To .Cols - 1
    .Col = i
    .CellFontSize = 8
    .CellBackColor = &H80000013
    .CellFontBold = True

Next

End With

End Sub

Public Sub LimpiarCampos()
   Rem mskFecha_Desde.Text = "__/__/____"
   Rem mskFecha_Hasta.Text = "__/__/____"
    
    lblCantidad_Cajas.Caption = ""
    lblCajasCrecimientoMes.Caption = ""
    
    lblCantidad_Libros.Caption = ""
    lblCantidadCajasVacias.Caption = ""
    
    lblCantidadDesarchivos.Caption = ""
    lbl_FletesNormales.Caption = ""
    
    lbl_FletesUrgentes.Caption = ""
    lblCantidad_Legajos.Caption = ""
    
    txtDetalleFacturacion.Text = ""
    lblPeriodoAnteriorFacturado.Caption = ""
    
    txtCantidad.Text = ""
    txtCodigo.Text = ""
    txtDescripcion.Text = ""
    txtPrecioUnitario.Text = ""
    txtTotal.Text = ""
    
    txtTotal.Text = ""
    lblIVA.Caption = ""
    
    lblTipoFactura.Caption = ""
    txtNumeroFactura.Text = ""
    txtFechaFactura.Text = ""
    
    lblSubTotal.Caption = ""
    lblPeriodoAnteriorFacturado.Caption = ""
    lblTotal.Caption = ""
    
    txtDescripcion_Factura.Text = ""
    
    
    lblRearchivo.Caption = ""
End Sub

Private Sub txtReciboNumero_LostFocus()
 Dim rs As New ADODB.Recordset
 Dim Sql As String
     If txtReciboNumero.Text = "" Then
        Exit Sub
     End If
     
    Sql = " SELECT ID_RECIBO , NUMERO_RECIBO, ESTADO_RECIBO "
    Sql = Sql & " From RECIBOS"
    Sql = Sql & " Where NUMERO_RECIBO = " & txtReciboNumero.Text
    Sql = Sql & " And ESTADO_RECIBO = 10"
    
    rs.Open Sql, ConActiva, 0, 1
    
    lblCod_Cliente_Recibo.Caption = ""
    If rs.EOF Then
        MsgBox "Recibo No disponible", vbCritical
        txtReciboNumero.Text = ""
        lbl_ID_RECIBO.Caption = ""
        Exit Sub
     Else
        lbl_ID_RECIBO.Caption = rs!ID_Recibo
    End If
    
    
    
    
 
End Sub

Public Sub SumarRecibo()
Dim TotalRecibo As Double
If txtReciboValores.Text <> "" Then
lblReciboTotal.Caption = CDbl(txtReciboValores.Text) + CDbl(txtRetencionesIVA.Text) + CDbl(txtRetencionesIngresosBrutos.Text) + CDbl(txtRetencionesGanancias.Text) + CDbl(txtRetencionesSUSS.Text)
End If


End Sub

Private Sub txtReciboValores_LostFocus()
 SumarRecibo
End Sub

Private Sub txtRetencionesGanancias_LostFocus()
SumarRecibo
End Sub

Private Sub txtRetencionesIngresosBrutos_LostFocus()
SumarRecibo
End Sub

Private Sub txtRetencionesIVA_LostFocus()
SumarRecibo
End Sub

Private Sub txtRetencionesSUSS_Change()
SumarRecibo
End Sub

Public Sub CambioColor(lblDato As Label)
If lblDato.Caption = 0 Then
    lblDato.BackColor = &H8000000F
Else
 lblDato.BackColor = &H80000018
End If


End Sub

Public Function CAJAS_CANTIDAD(COD_CLIENTE As Integer) As Long
        
        Dim rs As ADODB.Recordset
        Dim CAJAS As Long
        Dim CajasBajas As Long
        Dim Sql As String


        Sql = " SELECT SUM(CANTIDAD) As cantidadCajas "
        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
        Sql = Sql & vbCrLf & " Where id_cliente =" & COD_CLIENTE
        Sql = Sql & vbCrLf & " And TIPO = 0"
        Sql = Sql & vbCrLf & " And ANULADO Is Null"
        Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
        If COD_CLIENTE > 1000 Then
            Sql = Sql & vbCrLf & " AND FECHA >  " & FechaFormato("29/04/2014")
        End If
        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(FECHA_FIN)
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
        CAJAS = 0
        Else
        If IsNull(rs!CANTIDADCAJAS) Then
            CAJAS = 0
        Else
            CAJAS = rs!CANTIDADCAJAS
        End If
        End If
        ' Bajas
        Sql = "  SELECT SUM(CANTIDAD) AS BajasCajas "
        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
        Sql = Sql & vbCrLf & " Where id_cliente = " & COD_CLIENTE
        Sql = Sql & vbCrLf & "  And (TIPO = 3) And (ANULADO Is Null)"
        Sql = Sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 0"
        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(FECHA_FIN)
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
            CajasBajas = 0
         Else
            If IsNull(rs!BajasCajas) Then
                CajasBajas = 0
            Else
                CajasBajas = rs!BajasCajas
            End If
        End If
         If COD_CLIENTE > 1000 Then
           Rem  CajasBajas = 0
        End If
        
        CAJAS_CANTIDAD = CLng(CAJAS) - CLng(CajasBajas)
        
        
        

End Function

Public Function CAJAS_VACIAS(COD_CLIENTE As Integer, Remitos As String) As Integer
    Dim rs As ADODB.Recordset
    Dim DATO As String
    Dim cantidad As Integer
    Dim Sql As String
    DATO = ""
     
    Sql = " SELECT NRO_REMITO, NRO_REM_PROV, FECHA, CANTIDAD "
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " Where id_cliente =" & COD_CLIENTE
    Sql = Sql & vbCrLf & " And TIPO = 2"
    Sql = Sql & vbCrLf & " And ANULADO Is Null"
    Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & " Order by NRO_REMITO"
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    DATO = ""
    Do While Not rs.EOF
        Rem DATO = DATO & rs!NRO_REMITO & " / "
        DATO = DATO & rs!NRO_REMITO & vbCrLf
        cantidad = cantidad + rs!cantidad
        rs.MoveNext
    Loop
        CAJAS_VACIAS = cantidad
        Remitos = Trim(DATO)
End Function

Public Function CAJAS_CRECIMIENTO_MES(COD_CLIENTE As Integer, Remitos As String) As Integer
    
    
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim cantidad As Integer
    Dim DatoClip As String
    
    Remitos = ""
    
    DATO = ""
    CAJAS_CRECIMIENTO_MES = 0
    Sql = " SELECT NRO_REMITO, NRO_REM_PROV, FECHA, CANTIDAD , IMAGEN "
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " Where id_cliente =" & COD_CLIENTE
    Sql = Sql & vbCrLf & " And TIPO = 0"
    Sql = Sql & vbCrLf & " And ANULADO Is Null"
    Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
    Sql = Sql & vbCrLf & " AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & " AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & " Order by NRO_REMITO"
    
    DatoClip = ctlClienteFactura.Descripcion & vbCrLf
    DatoClip = DatoClip & " Crecimiento Mensual de cajas " & vbCrLf & " Desde el : " & mskFecha_Desde.Text & " hasta " & mskFecha_Hasta.Text & vbCrLf
   
    If chkCopiarImagenes.value = True Then
    
    
    End If
    
    
DatoClip = DatoClip & vbCrLf & "cantidad" & vbTab & "Fecha" & vbTab & "NRO_REM_PROV" & vbTab & "NRO_REMITO"
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    cantidad = 0
    Do While Not rs.EOF
        DATO = DATO & Replace(rs!NRO_REM_PROV, "0001-000", "") & "/" & vbCr
        cantidad = cantidad + rs!cantidad
        DatoClip = DatoClip & vbCrLf & rs!cantidad & vbTab & rs!fecha & vbTab & rs!NRO_REM_PROV & vbTab & rs!NRO_REMITO
        rs.MoveNext
    Loop
    
    Remitos = Trim(DATO)
   CAJAS_CRECIMIENTO_MES = cantidad
       
End Function
Public Function CAJAS_CRECIMIENTO_MES_andre(COD_CLIENTE As Integer, Remitos As String) As Integer
    
    
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim cantidad As Integer
    Dim DatoClip As String
    
    Remitos = ""
    
    DATO = ""
    CAJAS_CRECIMIENTO_MES_andre = 0
    Sql = " SELECT NRO_REMITO, NRO_REM_PROV, FECHA, CANTIDAD "
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " Where id_cliente =" & COD_CLIENTE
    Sql = Sql & vbCrLf & " And TIPO = 0"
    Sql = Sql & vbCrLf & " And ANULADO Is Null"
    Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
    Sql = Sql & vbCrLf & " AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & " AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & " Order by NRO_REMITO"
    
    DatoClip = ctlClienteFactura.Descripcion & vbCrLf
    DatoClip = DatoClip & " Crecimiento Mensual de cajas " & vbCrLf & " Desde el : " & mskFecha_Desde.Text & " hasta " & mskFecha_Hasta.Text & vbCrLf
   
    
    
DatoClip = DatoClip & vbCrLf & "cantidad" & vbTab & "Fecha" & vbTab & "NRO_REM_PROV" & vbTab & "NRO_REMITO"
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    cantidad = 0
    Do While Not rs.EOF
        DATO = DATO & Replace(rs!NRO_REM_PROV, "0001-000", "") & " \ "
        cantidad = cantidad + rs!cantidad
        DatoClip = DatoClip & vbCrLf & rs!cantidad & vbTab & rs!fecha & vbTab & rs!NRO_REM_PROV & vbTab & rs!NRO_REMITO
        rs.MoveNext
    Loop
    Clipboard.Clear
    
    Remitos = Trim(DATO)
   CAJAS_CRECIMIENTO_MES_andre = cantidad
    Clipboard.SetText DatoClip
End Function

Public Function ConsultasTodas(COD_CLIENTE As Integer) As Integer
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim cantidad As Integer
    
    
    
    Sql = "  SELECT  REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, "
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.CANTIDAD "
    Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO "
    Sql = Sql & vbCrLf & "  Where OPERACION = 1 And TIPO = 1 "
    Sql = Sql & vbCrLf & "  AND  REMITOS_CUERPO.ID_CLIENTE = " & COD_CLIENTE
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ANULADO IS NULL "
        Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & " ORDER BY NRO_REMITO "
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    cantidad = 0
    Do While Not rs.EOF
        DATO = DATO & " " & rs!NRO_REMITO & vbCrLf
        cantidad = cantidad + rs!cantidad
        rs.MoveNext
    
    Loop
        ConsultasTodas = cantidad
    
    
End Function

Public Function ConsultasCajas(COD_CLIENTE As Integer) As Integer
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim cantidad As Integer
    
    
    
    Sql = "  SELECT  REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, "
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.CANTIDAD "
    Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO "
    Sql = Sql & vbCrLf & "  Where OPERACION = 1 And TIPO = 1 "
    Sql = Sql & vbCrLf & "  AND  REMITOS_CUERPO.ID_CLIENTE = " & COD_CLIENTE
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ANULADO IS NULL "
    Sql = Sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 0 "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & " ORDER BY NRO_REMITO "
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    cantidad = 0
    Do While Not rs.EOF
        DATO = DATO & " " & rs!NRO_REMITO & vbCrLf
        cantidad = cantidad + rs!cantidad
        rs.MoveNext
    
    Loop
        ConsultasCajas = cantidad
    
    
End Function
Public Function ConsultasLegajos(COD_CLIENTE As Integer) As Integer
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim cantidad As Integer
    
    
    
    Sql = "  SELECT  REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, "
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.CANTIDAD "
    Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO "
    Sql = Sql & vbCrLf & "  Where OPERACION = 1 And TIPO = 1 "
    Sql = Sql & vbCrLf & "  AND  REMITOS_CUERPO.ID_CLIENTE = " & COD_CLIENTE
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ANULADO IS NULL "
    Sql = Sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 3 "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & " ORDER BY NRO_REMITO "
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    cantidad = 0
    Do While Not rs.EOF
        DATO = DATO & " " & rs!NRO_REMITO & vbCrLf
        cantidad = cantidad + rs!cantidad
        rs.MoveNext
    
    Loop
        ConsultasLegajos = cantidad
    
    
End Function

Public Function ConsultasLibros(COD_CLIENTE As Integer) As Integer
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim cantidad As Integer
    
    
    
    Sql = "  SELECT  REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, "
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.CANTIDAD "
    Sql = Sql & vbCrLf & " FROM REMITOS_CUERPO "
    Sql = Sql & vbCrLf & "  Where OPERACION = 1 And TIPO = 1 "
    Sql = Sql & vbCrLf & "  AND  REMITOS_CUERPO.ID_CLIENTE = " & COD_CLIENTE
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ANULADO IS NULL "
    Sql = Sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 1 "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & " ORDER BY NRO_REMITO "
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    cantidad = 0
    Do While Not rs.EOF
        DATO = DATO & " " & rs!NRO_REMITO & vbCrLf
        cantidad = cantidad + rs!cantidad
        rs.MoveNext
    Loop
        ConsultasLibros = cantidad
    
    
End Function
Public Function ConsultasCajasLegajos(COD_CLIENTE As Integer) As String
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim cantidad As Integer
    
    
    
    
    
      Sql = " SELECT     SUM(CANTIDAD) AS Cantidad"
      Sql = Sql & vbCrLf & " From REMITOS_CUERPO"
      Sql = Sql & vbCrLf & " WHERE     (OPERACION = 1) "
      Sql = Sql & vbCrLf & " AND (TIPO = 1) "
      Sql = Sql & vbCrLf & " AND ID_CLIENTE = " & COD_CLIENTE
      Sql = Sql & vbCrLf & " AND (ANULADO IS NULL) "
      Sql = Sql & vbCrLf & " AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
      Sql = Sql & vbCrLf & " AND " & FechaFormato(FECHA_FIN)
      Sql = Sql & vbCrLf & " AND (COD_TIPO_ALMACENAMIENTO = 0)"
    
    
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    cantidad = 0
    If Not rs.EOF Then
                If IsNull(rs!cantidad) Then
            cantidad = 0
        Else
            cantidad = rs!cantidad
            End If

        
    End If
 
    
        ConsultasCajasLegajos = "Cajas :" & cantidad
        
     Sql = " SELECT     SUM(CANTIDAD) AS Cantidad"
      Sql = Sql & vbCrLf & " From REMITOS_CUERPO"
      Sql = Sql & vbCrLf & " WHERE     (OPERACION = 1) "
      Sql = Sql & vbCrLf & " AND (TIPO = 1) "
      Sql = Sql & vbCrLf & " AND ID_CLIENTE = " & COD_CLIENTE
      Sql = Sql & vbCrLf & " AND (ANULADO IS NULL) "
      Sql = Sql & vbCrLf & " AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
      Sql = Sql & vbCrLf & " AND " & FechaFormato(FECHA_FIN)
      Sql = Sql & vbCrLf & " AND (COD_TIPO_ALMACENAMIENTO = 3)"
    
    
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    cantidad = 0
    If Not rs.EOF Then
                If IsNull(rs!cantidad) Then
            cantidad = 0
        Else
            cantidad = rs!cantidad
            End If

        
    End If
         ConsultasCajasLegajos = ConsultasCajasLegajos & "  legajos :" & cantidad
        
        
    
    Sql = " SELECT     SUM(CANTIDAD) AS Cantidad"
      Sql = Sql & vbCrLf & " From REMITOS_CUERPO"
      Sql = Sql & vbCrLf & " WHERE     (OPERACION = 1) "
      Sql = Sql & vbCrLf & " AND (TIPO = 1) "
      Sql = Sql & vbCrLf & " AND ID_CLIENTE = " & COD_CLIENTE
      Sql = Sql & vbCrLf & " AND (ANULADO IS NULL) "
      Sql = Sql & vbCrLf & " AND fECHA BETWEEN " & FechaFormato(FECHA_INICIO)
      Sql = Sql & vbCrLf & " AND " & FechaFormato(FECHA_FIN)
      Sql = Sql & vbCrLf & " AND (COD_TIPO_ALMACENAMIENTO = 1)"
    
    
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    cantidad = 0
    If Not rs.EOF Then
        If IsNull(rs!cantidad) Then
            cantidad = 0
        Else
            cantidad = rs!cantidad
            End If
        
    End If
         ConsultasCajasLegajos = ConsultasCajasLegajos & "  Libros :" & cantidad
        
    
    
End Function


Public Function FLETES_NORMALES(COD_CLIENTE As Integer)
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim cantidad As Integer
    Dim Cod_Flete As Long
    
    
    Sql = " SELECT REMITOS_CUERPO.COD_FLETE, REMITOS_CUERPO.NRO_REMITO,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ESTADO,"
    Sql = Sql & vbCrLf & "  REMITO_ESTADOS.DESCRIPCION,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.CANTIDAD,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.FECHA,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ANULADO"
    Sql = Sql & vbCrLf & "  From REMITOS_CUERPO, REMITO_ESTADOS"
    Sql = Sql & vbCrLf & "  WHERE REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ID_CLIENTE = " & COD_CLIENTE
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ANULADO IS NULL "
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ESTADO = 0 "
    Sql = Sql & vbCrLf & "  AND NOT COD_FLETE IS NULL  "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & "  ORDER BY COD_FLETE  "
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    
    
    Do While Not rs.EOF
    If Cod_Flete <> rs!Cod_Flete Then
        DATO = DATO & " " & rs!NRO_REMITO
        cantidad = cantidad + 1
        Cod_Flete = rs!Cod_Flete
    Else
        DATO = DATO & " " & rs!NRO_REMITO
    End If
        rs.MoveNext
    Loop
      
      FLETES_NORMALES = cantidad
    
End Function


Public Function FLETES_URGENTES(COD_CLIENTE As Integer) As Integer
    Dim DATO As String
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Dim cantidad As Integer
    Dim Cod_Flete As Long
    
    
    Sql = " SELECT REMITOS_CUERPO.COD_FLETE, REMITOS_CUERPO.NRO_REMITO,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ESTADO,"
    Sql = Sql & vbCrLf & "  REMITO_ESTADOS.DESCRIPCION,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.CANTIDAD,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ID_CLIENTE, REMITOS_CUERPO.FECHA,"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ANULADO"
    Sql = Sql & vbCrLf & "  From REMITOS_CUERPO, REMITO_ESTADOS"
    Sql = Sql & vbCrLf & "  WHERE REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
    Sql = Sql & vbCrLf & "  REMITOS_CUERPO.ID_CLIENTE = " & COD_CLIENTE
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ANULADO IS NULL "
    Sql = Sql & vbCrLf & "  AND REMITOS_CUERPO.ESTADO = 1 "
    Sql = Sql & vbCrLf & "  AND NOT COD_FLETE IS NULL  "
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & "  ORDER BY COD_FLETE  "
    
    Set rs = New ADODB.Recordset
    rs.Open Sql, ConActiva, 0, 1
    
    
    Do While Not rs.EOF
    If Cod_Flete <> rs!Cod_Flete Then
        DATO = DATO & " " & rs!NRO_REMITO
        cantidad = cantidad + 1
        Cod_Flete = rs!Cod_Flete
    Else
        DATO = DATO & " " & rs!NRO_REMITO
    End If
        rs.MoveNext
    Loop
    FLETES_URGENTES = cantidad

End Function


Public Sub REARCHIVO_FISICO()

        Dim DATO As String
        Dim rs As New ADODB.Recordset
        Dim cantidad As Integer
        Dim Sql As String
        Set rs = New ADODB.Recordset
            Sql = " SELECT  COD_REMITO_PRO, SUM(CANTIDAD) AS Cantidad"
            Sql = Sql & " From ORDENAR_DOCUMENTACION "
            Sql = Sql & " WHERE  "
            Sql = Sql & " COD_CLIENTE = " & COD_CLIENTE
            Sql = Sql & "  AND Fecha Between '" & FECHA_INICIO
            Sql = Sql & "' AND '" & FECHA_FIN & "'"
            Sql = Sql & "  GROUP BY COD_CLIENTE, COD_REMITO_PRO"
            rs.Open Sql, ConActiva, 0, 1
            Do While Not rs.EOF
                    DATO = DATO & " " & rs!COD_REMITO_PRO
                    cantidad = cantidad + rs!cantidad
                    rs.MoveNext
            Loop
            If cantidad <> 0 Then
                INSERTAR_FACTURA_CUSTODIA 10, cantidad, DATO, MES_SERVICIO
            End If
            
            
End Sub

Public Sub CARGA_LEGAJOS()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim cantidad As Integer
        Sql = " SELECT     count(*) as Cantidad"
        Sql = Sql & " From LEGAJOS"
        Sql = Sql & " WHERE COD_CLIENTE = " & COD_CLIENTE
        Sql = Sql & " AND FECHA_ACTUALIZACION BETWEEN '" & FECHA_INICIO
        Sql = Sql & "' AND '" & FECHA_FIN & "'"
        rs.Open Sql, ConActiva, 0, 1
        If Not rs.EOF Then
            cantidad = rs!cantidad
        End If
        
        If cantidad <> 0 Then
           INSERTAR_FACTURA_CUSTODIA 3, cantidad, "", MES_SERVICIO
        End If
End Sub

Public Sub INSERTAR_FACTURA_CUSTODIA(CODIGO_CUSTODIA_DETALLE, cantidad, Remitos, MES_SERVICIO)
Dim Sql As String
Dim COD_CLIENTE_CUSTODIA As Integer
COD_CLIENTE_CUSTODIA = COD_CLIENTE + 5000
Sql = " INSERT INTO FACTURACUSTODIA"
Sql = Sql & " ( COD_CLIENTE_CUSTODIA, CODIGO_CUSTODIA_DETALLE, "
Sql = Sql & "  CANTIDAD, REMITOS, MES_SERVICIO)"
Sql = Sql & "  VALUES  "
Sql = Sql & " (" & COD_CLIENTE_CUSTODIA & "," & CODIGO_CUSTODIA_DETALLE & ","
Sql = Sql & cantidad & ",'" & Remitos & "'," & MES_SERVICIO & ")"
ExecutarSql Sql


End Sub

Public Function IMAGENES(COD_CLIENTE As Integer, Remitos As String) As Long


        Dim DATO As String
        Dim rs As New ADODB.Recordset
        Dim cantidad As Long
        Dim CONlEGAJOS As ADODB.Connection
        
        Dim Sql As String
        cantidad = 0
        Set CONlEGAJOS = New ADODB.Connection
        CONlEGAJOS.CommandTimeout = 400
        CONlEGAJOS.Open strConBasa
        Set rs = New ADODB.Recordset
        
        Sql = " SELECT  REMITO,  sum(CANTIDAD_IMAGENES) as CANTIDAD"
        Sql = Sql & " From DOCUMENTOS_DIGITALES_LOTE"
        Sql = Sql & "  Where FK_CLIENTES  = " & COD_CLIENTE
        Sql = Sql & "  AND FECHA_SCANNER Between " & FechaFormato(FECHA_INICIO)
        Sql = Sql & "  AND " & FechaFormato(FECHA_FIN)
        Sql = Sql & " GROUP BY REMITO "

            rs.Open Sql, CONlEGAJOS, 0, 1
            Do While Not rs.EOF
                    DATO = DATO & Replace(Trim(rs!REMITO), "0001-000", "") & " \ "
                    cantidad = cantidad + rs!cantidad
                    rs.MoveNext
            Loop
            
            IMAGENES = cantidad
            Remitos = Trim(DATO)
        

End Function

Public Function Cantidad_Cajas()

'
''__________________Inicio Cajas ____________________________________
'        sql = " SELECT SUM(CANTIDAD) As cantidadCajas "
'        sql = sql & vbCrLf & " From REMITOS_CUERPO "
'        sql = sql & vbCrLf & " Where id_cliente =" & ctlClienteFactura.Valor
'        sql = sql & vbCrLf & " And TIPO = 0"
'        sql = sql & vbCrLf & " And ANULADO Is Null"
'        sql = sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 0"
'        sql = sql & vbCrLf & " AND FECHA <=  '" & mskFecha_Hasta.Text & "'"
'        Set rs = New ADODB.Recordset
'        rs.Open sql, strConBasa , 0 ,1
'        If rs.EOF Then
'        Cajas = 0
'        Else
'        If IsNull(rs!CantidadCajas) Then
'            Cajas = 0
'        Else
'            Cajas = rs!CantidadCajas
'        End If
'        End If
'        ' Bajas
'        sql = "  SELECT SUM(CANTIDAD) AS BajasCajas "
'        sql = sql & vbCrLf & " From REMITOS_CUERPO "
'        sql = sql & vbCrLf & " Where id_cliente = " & ctlClienteFactura.Valor
'        sql = sql & vbCrLf & "  And (TIPO = 3) And (ANULADO Is Null)"
'        sql = sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 0"
'        sql = sql & vbCrLf & " AND FECHA <=  '" & mskFecha_Hasta.Text & "'"
'        Set rs = New ADODB.Recordset
'        rs.Open sql, strConBasa , 0 ,1
'        If rs.EOF Then
'            CajasBajas = 0
'         Else
'            If IsNull(rs!BajasCajas) Then
'                CajasBajas = 0
'            Else
'                CajasBajas = rs!BajasCajas
'            End If
'        End If
'
'        lblCantidad_Cajas = CLng(Cajas) - CLng(CajasBajas)
'        CambioColor lblCantidad_Cajas
End Function

Public Sub Cantidad_libros()
'sql = " SELECT SUM(CANTIDAD) As CantidadLibros"
'        sql = sql & vbCrLf & " From REMITOS_CUERPO "
'        sql = sql & vbCrLf & " Where id_cliente =" & ctlClienteFactura.Valor
'        sql = sql & vbCrLf & " And TIPO = 0"
'        sql = sql & vbCrLf & " And ANULADO Is Null"
'        sql = sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 1"
'        Set rs = New ADODB.Recordset
'        rs.Open sql, strConBasa , 0 ,1
'        If rs.EOF Then
'            Libros = 0
'        Else
'            If IsNull(rs!CantidadLibros) Then
'                Libros = 0
'            Else
'                Libros = rs!CantidadLibros
'            End If
'        End If
'        ' Bajas
'        sql = "  SELECT SUM(CANTIDAD) AS LibrosBajas "
'        sql = sql & vbCrLf & " From REMITOS_CUERPO "
'        sql = sql & vbCrLf & " Where id_cliente = " & ctlClienteFactura.Valor
'        sql = sql & vbCrLf & "  And (TIPO = 3) And (ANULADO Is Null)"
'        sql = sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 1"
'        Set rs = New ADODB.Recordset
'        rs.Open sql, strConBasa , 0 ,1
'        If rs.EOF Then
'            LibrosBajas = 0
'         Else
'            If IsNull(rs!LibrosBajas) Then
'                LibrosBajas = 0
'            Else
'                LibrosBajas = rs!LibrosBajas
'            End If
'        End If
'
'
'        lblCantidad_Libros.Caption = CLng(Libros) - CLng(LibrosBajas)
'
'        CambioColor lblCantidad_Libros
        
End Sub

Public Function LIBROS_CANTIDAD(COD_CLIENTE As Integer)
        Dim Sql As String
        Dim rs As New ADODB.Recordset
        Dim LibrosAlta As Integer
        Dim LibrosBajas As Integer
        
        
        Sql = " SELECT SUM(CANTIDAD) As CantidadLibros"
        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
        Sql = Sql & vbCrLf & " Where id_cliente =" & COD_CLIENTE
        Sql = Sql & vbCrLf & " AND TIPO = 0"
        Sql = Sql & vbCrLf & " AND ANULADO Is Null"
        Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 1"
        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(FECHA_FIN)
        
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
            LibrosAlta = 0
        Else
            If IsNull(rs!CantidadLibros) Then
                LibrosAlta = 0
            Else
                LibrosAlta = rs!CantidadLibros
            End If
        End If
        ' Bajas
        Sql = "  SELECT SUM(CANTIDAD) AS LibrosBajas "
        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
        Sql = Sql & vbCrLf & " Where id_cliente = " & COD_CLIENTE
        Sql = Sql & vbCrLf & " AND (TIPO = 3) And (ANULADO Is Null)"
        Sql = Sql & vbCrLf & " AND COD_TIPO_ALMACENAMIENTO = 1"
        Sql = Sql & vbCrLf & " AND FECHA <=  " & FechaFormato(FECHA_FIN)
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        If rs.EOF Then
            LibrosBajas = 0
         Else
            If IsNull(rs!LibrosBajas) Then
                LibrosBajas = 0
            Else
                LibrosBajas = rs!LibrosBajas
            End If
        End If
        
         
  LIBROS_CANTIDAD = CLng(LibrosAlta) - CLng(LibrosBajas)
        
 
        
End Function

Public Function CONSULTAS_DIGITALES(COD_CLIENTE As Integer, cantidadImagenes As Integer) As Integer

Dim Sql As String
Dim rs As ADODB.Recordset


Set rs = New ADODB.Recordset
CONSULTAS_DIGITALES = 0
cantidadImagenes = 0
Sql = " SELECT IDREQUERIMIENTO, FECHARECEPCION, CANTIDAD, CANTIDAD_IMAGENES"
Sql = Sql & " From REQUERIMIENTO "
Sql = Sql & " WHERE ID_CLIENTE = " & COD_CLIENTE
Sql = Sql & " AND  IDTIPOREQUERIMIENTO in (13,14) "
 Sql = Sql & " AND    (ANULADO IS NULL)"
Sql = Sql & " AND  FECHARECEPCION Between " & FechaFormato(FECHA_INICIO)
Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)



Set rs = New ADODB.Recordset

rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
    CONSULTAS_DIGITALES = CONSULTAS_DIGITALES + rs!cantidad
    If Not IsNull(rs!Cantidad_Imagenes) Then
    cantidadImagenes = cantidadImagenes + rs!Cantidad_Imagenes
    End If
    rs.MoveNext
Loop




End Function

Public Function CONSULTAS_EN_PLANTA(COD_CLIENTE As Integer, REQUERIMIENTO As String) As Integer

Dim Sql As String
Dim rs As ADODB.Recordset


Set rs = New ADODB.Recordset
CONSULTAS_EN_PLANTA = 0

Sql = " SELECT IDREQUERIMIENTO, FECHARECEPCION, CANTIDAD "
Sql = Sql & " From REQUERIMIENTO  "
Sql = Sql & " WHERE ID_CLIENTE =  " & COD_CLIENTE
Sql = Sql & " AND  IDTIPOREQUERIMIENTO in (9) "
Sql = Sql & " AND  (ANULADO IS NULL)"
Sql = Sql & " AND  FECHARECEPCION Between " & FechaFormato(FECHA_INICIO)
Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)



Set rs = New ADODB.Recordset

rs.Open Sql, ConActiva, 0, 1
CONSULTAS_EN_PLANTA = 0
REQUERIMIENTO = ""
Do While Not rs.EOF
    REQUERIMIENTO = REQUERIMIENTO & " REQ:" & rs!IDREQUERIMIENTO
    If Not IsNull(rs!cantidad) Then
    CONSULTAS_EN_PLANTA = CONSULTAS_EN_PLANTA + rs!cantidad
    End If
    rs.MoveNext
Loop




End Function

Public Function ORDEN_DOCUMENTACION(COD_CLIENTE As Integer, Remitos As String, TIPO As Tipo_Orden) As Integer
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    ORDEN_DOCUMENTACION = 0
    Remitos = ""
    

 
 
        Sql = " SELECT     COD_REMITO_PRO, SUM(CANTIDAD) AS cantidad"
        Sql = Sql & "  From ORDENAR_DOCUMENTACION"
        Sql = Sql & "  Where COD_CLIENTE = " & COD_CLIENTE
        If TIPO = FISICO Then
            Sql = Sql & " AND COD_TIPO_ORDEN = 'FISICO'"
       End If
        If TIPO = lote Then
            Sql = Sql & " AND COD_TIPO_ORDEN = 'LOTE'"
        End If
        Sql = Sql & " AND FECHA BETWEEN " & FechaFormato(mskFecha_Desde.Text)
        Sql = Sql & " AND " & FechaFormato(mskFecha_Hasta)
        Sql = Sql & "   GROUP BY COD_REMITO_PRO"
        Sql = Sql & "  ORDER BY COD_REMITO_PRO"
    rs.Open Sql, ConActiva, 0, 1
    Do While Not rs.EOF
        Remitos = Remitos & Replace(rs!COD_REMITO_PRO, "0001-000", "") & vbCrLf
        ORDEN_DOCUMENTACION = ORDEN_DOCUMENTACION + rs!cantidad
        rs.MoveNext
    Loop
    Remitos = Trim(Remitos)
    
End Function

Public Function ImagenesProcesadas(COD_CLIENTE As Integer, Remitos As String) As Long
    Dim Sql As String
    Dim rs As ADODB.Recordset



Sql = " SELECT COD_CLIENTE, FECHA_INCORPORACION, CANTIDAD_IMAGENES, REMITO"
Sql = Sql & " From DOCUMENTOS_DIGITALES"
Sql = Sql & " WHERE     (COD_CLIENTE = 40) "
Sql = Sql & " AND (FECHA_INCORPORACION BETWEEN CONVERT(DATETIME, '2008-01-01 00:00:00', 102) AND CONVERT(DATETIME,"
                      '2008-01-30 00:00:00', 102))


End Function

Public Function RECAMBIO_CAJAS(FK_CLIENTE As Integer) As Integer
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Sql = " SELECT COUNT(*) AS CANTIDAD "
    Sql = Sql & vbCrLf & " From dbo.RECAMBIO_CAJA "
    Sql = Sql & vbCrLf & " Where FK_CLIENTE = " & FK_CLIENTE
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    
    rs.Open Sql, ConActiva, 0, 1
    
   If rs.EOF Then
        RECAMBIO_CAJAS = 0
   Else
        RECAMBIO_CAJAS = rs!cantidad
   End If
   


End Function

Public Function BajasMensuales(FK_CLIENTE As Integer, RemitosBajas As String) As Long

    Dim rs As New ADODB.Recordset
    Dim cantidad As Integer
    Dim Remitos As String
    Dim Sql As String

        Sql = "  SELECT NRO_REMITO, CANTIDAD "
        Sql = Sql & vbCrLf & " From REMITOS_CUERPO "
        Sql = Sql & vbCrLf & " Where id_cliente = " & FK_CLIENTE
        Sql = Sql & vbCrLf & "  And (TIPO = 3) And (ANULADO Is Null)"
        Sql = Sql & vbCrLf & "  AND COD_TIPO_ALMACENAMIENTO = 0 "
        Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
        Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
        Set rs = New ADODB.Recordset
        rs.Open Sql, ConActiva, 0, 1
        RemitosBajas = ""
cantidad = 0
        Do While Not rs.EOF
              cantidad = cantidad + rs!cantidad
              Remitos = Remitos & vbCrLf & rs!NRO_REMITO
              rs.MoveNext
        Loop
        
        
BajasMensuales = cantidad
RemitosBajas = Remitos


End Function

Public Function Precintos(COD_CLIENTE As Long, fechadesde As String, FechaHasta As String) As Integer
Dim rs As New ADODB.Recordset
Dim Sql As String

Sql = " SELECT     IDREQUERIMIENTO, ID_CLIENTE, IDTIPOREQUERIMIENTO, CANTIDAD, ANULADO, FECHAENTREGA"
Sql = Sql & " From REQUERIMIENTO"
Sql = Sql & "  Where(IDTIPOREQUERIMIENTO = 23)"
Sql = Sql & vbCrLf & "  AND FECHARECEPCION BETWEEN " & FechaFormato(fechadesde)
Sql = Sql & vbCrLf & "  AND " & FechaFormato(FechaHasta)
Sql = Sql & vbCrLf & " and  ID_CLIENTE = " & COD_CLIENTE
rs.Open Sql, ConActiva, 0, 1

Precintos = 0
        
        
         Do While Not rs.EOF
            Precintos = Precintos + rs!cantidad
            rs.MoveNext
         Loop
         
End Function

Public Sub Facturacion_Mensual()
    Dim rsFacturacion As New ADODB.Recordset
    Dim rsFletes As New ADODB.Recordset
    Dim RsDigital As New ADODB.Recordset
    Dim rsMaxRemito As New ADODB.Recordset
    Dim Sql As String
    Dim Cantidades As Long
    Dim RemitosVacias As String
    Dim CantidadImagen As Integer
    Dim RemitosCrecimiento As String
    Dim RemitosOrden As String
    Dim RemitosImagenes As String
    Dim RemitosBajas As String
    
    Dim TITULO_SQL As String
    Dim GRUPO_SQL As Integer
    Dim COD_CLIENTE_SQL As Integer
    Dim RAZON_SOCIAL_SQL   As String
    Dim CAJAS_CANON_SQL As String
    Dim CAJAS_CRECIMIENTO_SQL As String
    Dim CAJAS_VACIAS_SQL As String
    Dim CAJAS_BAJAS_SQL As String
    Dim CAJAS_CAMBIO_SQL As String
    Dim LIBROS_SQL As String
    Dim LEGAJOS_ACUMULADOS_SQL As String
    Dim LEGAJOS_CRECIMIENTOS_SQL As String
    Dim CONSULTAS_CAJAS_SQL As String
    Dim CONSULTAS_LEGAJOS_SQL As String
    Dim CONSULTAS_LIBROS_SQL As String
    Dim CONSULTAS_PLANTA_SQL As String
    Dim CONSULTAS_DIGITALES_SQL As String
    Dim CONSULTA_IMAGENES_SQL As String
    Dim FLETES_NORMALES_SQL
    Dim FLETES_URGENTES_SQL As String
    Dim REARCHIVO_FISICO_SQL  As String
    Dim REARCHIVO_LOTE_SQL As String
    Dim IMAGENES_SQL As String
    Dim PRECINTOS_SQL As String
    Dim HORAS_ARCHIVISTA_PLANTA_SQL As String
    Dim HORAS_ARCHIVISTA_CLIENTE_SQL As String
    
    
    
    
    
    

    
    
    
    CantidadImagen = 0
On Error GoTo salir:

Dim R As Integer
Dim C As Integer
        
    Dim Filtro As String
   
    Sql = " SELECT     COD_CLIENTE_CABECERA , ID_CLIENTE, RAZON_SOCIAL, NOFACTURAR ,DETALLE_FACTURACION"
    Sql = Sql & " From Clientes "
    Sql = Sql & " where NOFACTURAR is null  "
    
    If MsgBox("Facturar Los clientes B Custodia ", vbYesNo) = vbYes Then
       
    
    Filtro = "34, 294, 321, 1022, 1025, 1028, 1036, 1038, 1048, 1049, 1056, 1060, 1108, 1132, 1134, 1141, 1143, 1149, 1152, 1156, 1164, 1205, 1295"
    If Filtro <> "0" Then
            Sql = Sql & " AND ID_CLIENTE In ( " & Filtro & ")"
        End If
    Else
        Filtro = InputBox("Ingrese los Nº de clientes separados por ," & vbCrLf & "Para todos los clientes 0", "Filtro Cliente", 0)
        If Filtro <> "0" Then
            Sql = Sql & " AND ID_CLIENTE In ( " & Filtro & ")"
        End If
    End If
    
    
    Sql = Sql & " Order by  ID_CLIENTE "
    
    Set rsFacturacion = New ADODB.Recordset
    rsFacturacion.Open Sql, strConBasa, 0, 1

    
    FECHA_INICIO = mskFecha_Desde.Text
    FECHA_FIN = mskFecha_Hasta.Text
    
    
    Sql = " SELECT    NRO_REMITO, FECHA, ID_CLIENTE , COD_FLETE"
    Sql = Sql & vbCrLf & " From REMITOS_CUERPO"
    Sql = Sql & vbCrLf & " Where (Cod_Flete Is Null)"
    Sql = Sql & vbCrLf & " And ANULADO Is Null"
    Sql = Sql & vbCrLf & "  AND FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    Sql = Sql & vbCrLf & "  AND (REMITOS_CUERPO.TIPO = 1) AND"
    Sql = Sql & vbCrLf & "  (REMITOS_CUERPO.OPERACION = 1) "
    
   
    
    
     Set rsFletes = New ADODB.Recordset
    
    Dim flete As String
    rsFletes.Open Sql, ConActiva, 0, 1
    
    If Not rsFletes.EOF Then
                    flete = flete & " ; " & rsFletes!Cod_Flete
          rsFletes.MoveNext
    Else
       MsgBox "Atencion  faltan procesar los fletes los datos seran copiados a memoria" & vbCrLf & flete, vbCritical
        flete = "NRO_REMITO" & vbTab & "FECHA" & vbTab & "ID_CLIENTE"
        Do While Not rsFletes.EOF
                    
            
             flete = flete & " vbcrlf  " & rsFletes!NRO_REMITO & vbTab & rsFletes!fecha & vbTab & rsFletes!id_cliente
       rsFletes.MoveNext
        Loop
        Clipboard.Clear
        Clipboard.SetText flete
        
     If MsgBox("Usted Quiere continuar sin procesar los fletes", vbYesNo) = vbNo Then
     Exit Sub
     
     Else
     End If
     
    End If
    
    
    
    
    
       
        
Sql = "    SELECT     REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.ID_CLIENTE, REQUERIMIENTO.IDTIPOREQUERIMIENTO,"
Sql = Sql & vbCrLf & " REQUERIMIENTO.Cantidad_Imagenes , REQUERIMIENTO.FECHARECEPCION, Clientes.NOFACTURAR"
Sql = Sql & vbCrLf & " FROM         REQUERIMIENTO INNER JOIN"
Sql = Sql & vbCrLf & " CLIENTES ON REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE"
Sql = Sql & vbCrLf & " WHERE     (REQUERIMIENTO.IDTIPOREQUERIMIENTO IN (13, 14)) AND (REQUERIMIENTO.CANTIDAD_IMAGENES IS NULL) "
 If Filtro <> "0" Then
         Sql = Sql & " AND REQUERIMIENTO.ID_CLIENTE In ( " & Filtro & ")"
         Else
         Sql = Sql & vbCrLf & " AND ( CLIENTES.NOFACTURAR IS NULL)"
        End If

Sql = Sql & vbCrLf & " AND (REQUERIMIENTO.ANULADO IS NULL) AND REQUERIMIENTO.FECHARECEPCION BETWEEN " & FechaFormato("01/09/2010") & " AND " & FechaFormato("30/09/2010")

        Sql = Sql & vbCrLf & "  AND FECHARECEPCION BETWEEN " & FechaFormato(FECHA_INICIO)
        Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
        
        Set RsDigital = New ADODB.Recordset
        
        RsDigital.Open Sql, strConBasa, 0, 1
        Dim ErrorDigital As String
        
        ErrorDigital = ""
        
        
         Do While Not RsDigital.EOF
            ErrorDigital = ErrorDigital & " " & RsDigital!IDREQUERIMIENTO
            RsDigital.MoveNext
         Loop
         
         If ErrorDigital <> "" Then
            MsgBox "Falta procesar los requerimientos digitales Nº :" & ErrorDigital
            Exit Sub
         End If
    
        
        
    
    If DateDiff("d", mskFecha_Hasta.Text, Now) < 1 Then
        MsgBox "Atención NO se puede procesar por que pueden existir errores" & vbCrLf & "Por favor Modifique la Fecha Hasta", vbCritical
        Exit Sub
    End If
    
    Sql = " SELECT     MAX(NRO_REMITO) AS MaxRemito From REMITOS_CUERPO "
    Sql = Sql & vbCrLf & " WHERE FECHA BETWEEN " & FechaFormato(FECHA_INICIO)
    Sql = Sql & vbCrLf & "  AND " & FechaFormato(FECHA_FIN)
    
    Set rsMaxRemito = New ADODB.Recordset
    
    
    rsMaxRemito.Open Sql, ConActiva, 0, 1
    
    
    
            GRUPO_SQL = 0
            COD_CLIENTE_SQL = 0
            RAZON_SOCIAL_SQL = ""
            CAJAS_CANON_SQL = ""
            CAJAS_CRECIMIENTO_SQL = ""
            CAJAS_VACIAS_SQL = ""
            CAJAS_BAJAS_SQL = ""
            CAJAS_CAMBIO_SQL = ""
            LIBROS_SQL = ""
            LEGAJOS_ACUMULADOS_SQL = ""
            LEGAJOS_CRECIMIENTOS_SQL = ""
            CONSULTAS_CAJAS_SQL = ""
            CONSULTAS_PLANTA_SQL = ""
            CONSULTAS_LIBROS_SQL = ""
            
            CONSULTAS_DIGITALES_SQL = ""
            CONSULTA_IMAGENES_SQL = ""
            REARCHIVO_FISICO_SQL = ""
            REARCHIVO_LOTE_SQL = ""
            IMAGENES_SQL = ""
            HORAS_ARCHIVISTA_PLANTA_SQL = ""
            HORAS_ARCHIVISTA_CLIENTE_SQL = ""
            PRECINTOS_SQL = ""
    

  
TITULO_SQL = "Movimientos del " & FECHA_INICIO & " hasta " & FECHA_FIN
TITULO_SQL = TITULO_SQL & vbCrLf & " Ultimo Remito " & rsMaxRemito!MaxRemito
 ExecutarSql " Delete From basasql.dbo.FACTURACION"

 Do While Not rsFacturacion.EOF
    If IsNull(rsFacturacion!COD_CLIENTE_CABECERA) Then
        GRUPO_SQL = 0
    Else
        GRUPO_SQL = rsFacturacion!COD_CLIENTE_CABECERA
    End If
    COD_CLIENTE_SQL = rsFacturacion!id_cliente
    RAZON_SOCIAL_SQL = Mid(Trim(rsFacturacion!RAZON_SOCIAL), 1, 30)
     If rsFacturacion!id_cliente = 39 Then
        CAJAS_CANON_SQL = 16955
     Else
        CAJAS_CANON_SQL = CAJAS_CANTIDAD(rsFacturacion!id_cliente) + CajasSumaResta(rsFacturacion!id_cliente)
     End If
    CAJAS_CRECIMIENTO_SQL = CAJAS_CRECIMIENTO_MES(rsFacturacion!id_cliente, RemitosCrecimiento)
    If RemitosCrecimiento <> "" Then
        Rem CAJAS_CRECIMIENTO_SQL = CAJAS_CRECIMIENTO_SQL & " \ " & "RM:" & " \ " & Trim(RemitosCrecimiento)
         CAJAS_CRECIMIENTO_SQL = CAJAS_CRECIMIENTO_SQL & vbCrLf & "RM:" & " \ " & Trim(RemitosCrecimiento)
    End If
    
    CAJAS_VACIAS_SQL = CAJAS_VACIAS(rsFacturacion!id_cliente, RemitosVacias)
    If RemitosVacias <> "" Then
        CAJAS_VACIAS_SQL = CAJAS_VACIAS_SQL & " - " & "RS:" & " \ " & Trim(RemitosVacias)
    End If
    
    CAJAS_BAJAS_SQL = BajasMensuales(rsFacturacion!id_cliente, RemitosBajas)
    If RemitosBajas <> "" Then
        CAJAS_BAJAS_SQL = CAJAS_BAJAS_SQL & vbCrLf & "RS:" & Trim(RemitosBajas)
     End If
    CAJAS_CAMBIO_SQL = RECAMBIO_CAJAS(rsFacturacion!id_cliente)
    LIBROS_SQL = LIBROS_CANTIDAD(rsFacturacion!id_cliente)
    LEGAJOS_ACUMULADOS_SQL = LEGAJOS_CANTIDAD(rsFacturacion!id_cliente)
    LEGAJOS_CRECIMIENTOS_SQL = LEGAJOS_CARGA(rsFacturacion!id_cliente)
    CONSULTAS_PLANTA_SQL = CONSULTAS_EN_PLANTA(rsFacturacion!id_cliente, "")
    CONSULTAS_CAJAS_SQL = ConsultasCajas(rsFacturacion!id_cliente)
    CONSULTAS_LEGAJOS_SQL = ConsultasLegajos(rsFacturacion!id_cliente)
    CONSULTAS_LIBROS_SQL = ConsultasLibros(rsFacturacion!id_cliente)
   Rem  MsgBox ConsultasTodas(rsFacturacion!id_cliente)
    
    
    CONSULTAS_DIGITALES_SQL = CONSULTAS_DIGITALES(rsFacturacion!id_cliente, CantidadImagen)
    CONSULTA_IMAGENES_SQL = CantidadImagen
    FLETES_NORMALES_SQL = FLETES_NORMALES(rsFacturacion!id_cliente)
    FLETES_URGENTES_SQL = FLETES_URGENTES(rsFacturacion!id_cliente)
    REARCHIVO_FISICO_SQL = ORDEN_DOCUMENTACION(rsFacturacion!id_cliente, RemitosOrden, FISICO)
    If RemitosOrden <> "" Then
            REARCHIVO_FISICO_SQL = REARCHIVO_FISICO_SQL & " \ " & "RM:" & " \ " & RemitosOrden
    End If
        
    REARCHIVO_LOTE_SQL = ORDEN_DOCUMENTACION(rsFacturacion!id_cliente, RemitosOrden, lote)
    If RemitosOrden <> "" Then
        REARCHIVO_LOTE_SQL = REARCHIVO_LOTE_SQL & " \ " & "RM:" & " \ " & Trim(RemitosOrden)
    End If
    
    IMAGENES_SQL = IMAGENES(rsFacturacion!id_cliente, RemitosImagenes)
    If RemitosImagenes <> "" Then
        IMAGENES_SQL = IMAGENES_SQL & " \ " & "RM:" & " \ " & Trim(RemitosImagenes)
    End If
      
    PRECINTOS_SQL = Precintos(rsFacturacion!id_cliente, mskFecha_Desde, mskFecha_Hasta)
    
        Sql = " Insert "
        Sql = Sql & vbCrLf & " INTO FACTURACION("
        Sql = Sql & vbCrLf & " TITULO"
        Sql = Sql & vbCrLf & ",GRUPO"
        Sql = Sql & vbCrLf & ",COD_CLIENTE"
        Sql = Sql & vbCrLf & ",RAZON_SOCIAL"
        Sql = Sql & vbCrLf & ",CAJAS_CANON"
        Sql = Sql & vbCrLf & ",CAJAS_CRECIMIENTO"
        Sql = Sql & vbCrLf & ",CAJAS_VACIAS"
        Sql = Sql & vbCrLf & ",CAJAS_BAJAS"
        Sql = Sql & vbCrLf & ",CAJAS_CAMBIO"
        Sql = Sql & vbCrLf & ",LIBROS"
        Sql = Sql & vbCrLf & ",LEGAJOS_ACUMULADOS"
        Sql = Sql & vbCrLf & ",LEGAJOS_CRECIMIENTOS"
        Sql = Sql & vbCrLf & ",CONSULTAS_CAJAS"
        Sql = Sql & vbCrLf & ",CONSULTAS_LEGAJOS"
        Sql = Sql & vbCrLf & ",CONSULTAS_LIBROS"
        Sql = Sql & vbCrLf & ",CONSULTAS_PLANTA"
        Sql = Sql & vbCrLf & ",CONSULTAS_DIGITALES"
        Sql = Sql & vbCrLf & ",CONSULTA_IMAGENES"
        Sql = Sql & vbCrLf & ",FLETES_NORMALES"
        Sql = Sql & vbCrLf & ",FLETES_URGENTES"
        Sql = Sql & vbCrLf & ",REARCHIVO_FISICO"
        Sql = Sql & vbCrLf & ",REARCHIVO_LOTE"
        Sql = Sql & vbCrLf & ",IMAGENES"
        Sql = Sql & vbCrLf & ",HORAS_ARCHIVISTA_PLANTA"
        Sql = Sql & vbCrLf & ",HORAS_ARCHIVISTA_CLIENTE"
        Sql = Sql & vbCrLf & ",PRECINTOS)"
        Sql = Sql & vbCrLf & " VALUES   ("
        Sql = Sql & vbCrLf & "'" & TITULO_SQL
        Sql = Sql & vbCrLf & "'," & GRUPO_SQL
        Sql = Sql & vbCrLf & "," & COD_CLIENTE_SQL
        Sql = Sql & vbCrLf & ",'" & RAZON_SOCIAL_SQL
        Sql = Sql & vbCrLf & "','" & CAJAS_CANON_SQL
        Sql = Sql & vbCrLf & "','" & CAJAS_CRECIMIENTO_SQL
        Sql = Sql & vbCrLf & "','" & CAJAS_VACIAS_SQL
        Sql = Sql & vbCrLf & "','" & CAJAS_BAJAS_SQL
        Sql = Sql & vbCrLf & "','" & CAJAS_CAMBIO_SQL
        Sql = Sql & vbCrLf & "','" & LIBROS_SQL
        Sql = Sql & vbCrLf & "','" & LEGAJOS_ACUMULADOS_SQL
        Sql = Sql & vbCrLf & "','" & LEGAJOS_CRECIMIENTOS_SQL
        Sql = Sql & vbCrLf & "','" & CONSULTAS_CAJAS_SQL
        Sql = Sql & vbCrLf & "','" & CONSULTAS_LEGAJOS_SQL
        Sql = Sql & vbCrLf & "','" & CONSULTAS_LIBROS_SQL
        Sql = Sql & vbCrLf & "','" & CONSULTAS_PLANTA_SQL
        Sql = Sql & vbCrLf & "','" & CONSULTAS_DIGITALES_SQL
        Sql = Sql & vbCrLf & "','" & CONSULTA_IMAGENES_SQL
        Sql = Sql & vbCrLf & "','" & FLETES_NORMALES_SQL
        Sql = Sql & vbCrLf & "','" & FLETES_URGENTES_SQL
        Sql = Sql & vbCrLf & "','" & REARCHIVO_FISICO_SQL
        Sql = Sql & vbCrLf & "','" & REARCHIVO_LOTE_SQL
        Sql = Sql & vbCrLf & "','" & IMAGENES_SQL
        Sql = Sql & vbCrLf & "','" & HORAS_ARCHIVISTA_PLANTA_SQL
        Sql = Sql & vbCrLf & "','" & HORAS_ARCHIVISTA_CLIENTE_SQL
        Sql = Sql & vbCrLf & "','" & PRECINTOS_SQL & "')"
    
    ExecutarSql Sql
    
    
    rsFacturacion.MoveNext
 Loop

Dim rs As New ADODB.Recordset
        Sql = " SELECT   * "
        Sql = Sql & " From basasql.dbo.FACTURACION"
        Sql = Sql & " ORDER BY COD_CLIENTE"
      Rem  frmReportes.ImprimirReporte PasoReportes & "Facturacion_Jose.rpt", Sql, True
        Set rs = New ADODB.Recordset
        rs.Open Sql, strConBasa, 3, 2
        Set grdfactura.DataSource = rs.DataSource
        grdfactura.DataMember = rs.DataMember
        MsgBox "Operacion terminada"
      Exit Sub
        
salir:
     
     MsgBox "Error " & Err.Description
        
End Sub

Public Function CajasSumaResta(Cliente As Integer) As Long
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    
    
    Sql = " SELECT     ID_CLIENTE, CANTIDADCAJASUMARESTA"
Sql = Sql & " From basasql.dbo.Clientes "
Sql = Sql & "  Where id_cliente =" & Cliente

rs.Open Sql, strConBasa

If Not IsNull(rs!CANTIDADCAJASUMARESTA) Then
    CajasSumaResta = rs!CANTIDADCAJASUMARESTA
Else
    CajasSumaResta = 0

End If


End Function

Public Sub ActualizarFacturasCustodia(FacturaABC As String, NumeroFactura As Long)
       
            

End Sub

Public Sub InsertarFacturaElectronica(NumeroFactura As Long, IDCliente As Integer, FacturaABC As String, Cuit As String, Subtotal As String, IVAInscripto As String, TotalFacturado As String, MesFacturacion As Integer, AnoFacturacion As Integer, NombreCliente As String, FechaFacturacion As Long)
            Dim Sql As String
            Dim SqlFacturaCabecera As String
            Dim conData As New ADODB.Connection
            Dim ConFacturaElectronica As New ADODB.Connection
            Dim RsFactura As New ADODB.Recordset
            Dim rsTEM_IVA_DATA As New ADODB.Recordset
            Dim rsFACDET As New ADODB.Recordset
            Dim Empresa As String
            Dim TipoFactura As String
            Dim PuntoDeVenta As Integer
            Dim Nro_Factura As Long
            Dim MAX_ID_FACTURA As Long
            Dim Rs_MAX_ID_FACTURA As New ADODB.Recordset
            Dim ConDetalleFactura As New ADODB.Connection
            
             Rem ConPedro.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Serverbasa1\SistemasBasa\ClientesBases\cambio.mdb;Persist Security Info=False"
            
            If chkPasarTodas.value = 1 Then
                ConDetalleFactura.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=Z:\Sistemas\Datas\FACTURASMDB\facturas" & Month(Now) & ".mdb"
            End If
            
            
            
            Rem Z:\Sistemas\Datas\FACTURASMDB
            
                conData.Open "Provider=MSDASQL.1;Persist Security Info=False;Data Source=datas"
                Sql = " UPDATE factura SET DetallePago = 'PASO " & Now & "'"
                Sql = Sql & " Where FacturaABC =  '" & FacturaABC & "'"
                Sql = Sql & " and  NumeroFactura = " & NumeroFactura
               conData.Execute Sql
                Rem MsgBox strConBasa
                ConFacturaElectronica.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=False;User ID=sa;Initial Catalog=factura_electronica;Data Source=222.15.19.150"
                    PuntoDeVenta = 0
                                Nro_Factura = 0
                    
                   TipoFactura = ""
                   If FacturaABC = "A" Or FacturaABC = "B" Then
                                               
                        Empresa = "'Custodia'"
                        TipoFactura = "'" & FacturaABC & "'"
                        If NumeroFactura > 20000000 Then
                                PuntoDeVenta = 2
                                Nro_Factura = Mid(NumeroFactura, 4)
                        Else
                                PuntoDeVenta = 1
                                Nro_Factura = Mid(NumeroFactura, 4)
                        End If
                        
                        
                        
                   End If
                   If FacturaABC = "F" Or FacturaABC = "G" Then
                        Empresa = "'Basa'"
                        If FacturaABC = "F" Then
                            TipoFactura = "'A'"
                            If CLng(NumeroFactura) > 40000000 Then
                                PuntoDeVenta = 4
                                Nro_Factura = Mid(NumeroFactura, 4)
                            Else
                                PuntoDeVenta = 1
                                Nro_Factura = Mid(NumeroFactura, 4)
                            End If
                            
                        End If
                        If FacturaABC = "G" Then
                            TipoFactura = "'B'"
                            If CLng(NumeroFactura) > 40000000 Then
                                PuntoDeVenta = 4
                                Nro_Factura = Mid(NumeroFactura, 4)
                            Else
                                PuntoDeVenta = 1
                                Nro_Factura = Mid(NumeroFactura, 4)
                            End If
                            
                        End If
                        
                   End If
                    
                    Sql = "INSERT INTO factura_electronica.dbo.FACTURA ( "
                    Sql = Sql & vbCrLf & " NumeroFactura "
                    Sql = Sql & vbCrLf & " , IDCliente"
                    Sql = Sql & vbCrLf & " , FacturaABC"
                    Sql = Sql & vbCrLf & " , CUIT"
                    Sql = Sql & vbCrLf & " , Subtotal"
                    Sql = Sql & vbCrLf & " , IVAInscripto"
                    Sql = Sql & vbCrLf & " , TotalFacturado"
                    Sql = Sql & vbCrLf & " , MesFacturacion"
                    Sql = Sql & vbCrLf & " , AnoFacturacion"
                    Sql = Sql & vbCrLf & " , NombreCliente"
                    Sql = Sql & vbCrLf & " , Fecha"
                    Sql = Sql & vbCrLf & " , Empresa "
                    Sql = Sql & vbCrLf & " , TipoFactura "
                    Sql = Sql & vbCrLf & " , PuntoDeVenta "
                    Sql = Sql & vbCrLf & " , Nro_Factura "
                    Sql = Sql & vbCrLf & " )"
                    Sql = Sql & vbCrLf & " VALUES ("
                    Sql = Sql & vbCrLf & NumeroFactura
                    Sql = Sql & vbCrLf & " , " & IDCliente
                    Sql = Sql & vbCrLf & " , '" & Trim(FacturaABC) & "'"
                    Sql = Sql & vbCrLf & " , '" & Trim(Cuit) & "'"
                    Sql = Sql & vbCrLf & " , '" & Replace(Subtotal, ",", ".") & "'"
                    Sql = Sql & vbCrLf & " , '" & Replace(IVAInscripto, ",", ".") & "'"
                    Sql = Sql & vbCrLf & " , '" & Replace(TotalFacturado, ",", ".") & "'"
                    Sql = Sql & vbCrLf & " , " & MesFacturacion
                    Sql = Sql & vbCrLf & " , " & AnoFacturacion
                    Sql = Sql & vbCrLf & " , '" & NombreCliente & "'"
                    Sql = Sql & vbCrLf & " , '" & DateAdd("D", FechaFacturacion, "28/12/1800") & "'"
                    Sql = Sql & vbCrLf & " , " & Empresa
                    Sql = Sql & vbCrLf & " , " & TipoFactura
                    Sql = Sql & vbCrLf & " , " & PuntoDeVenta
                    Sql = Sql & vbCrLf & " , " & Nro_Factura
                    Sql = Sql & vbCrLf & " )"
                   SqlFacturaCabecera = Sql
                   
                   
                    
                    Sql = " SELECT MAX(ID_FACTURA) AS Rs_MAX_ID_FACTURA"
                    Sql = Sql & vbCrLf & "  From factura_electronica.dbo.FACTURA"
                    
                    Set Rs_MAX_ID_FACTURA = New ADODB.Recordset
                    
                    Rs_MAX_ID_FACTURA.Open Sql, ConFacturaElectronica
                    MAX_ID_FACTURA = Rs_MAX_ID_FACTURA!Rs_MAX_ID_FACTURA
                
                    Sql = " SELECT FACTURAABC , NUMEROFACTURA , CANTIDAD "
                    Sql = Sql & vbCrLf & " , PRECIOUNITARIO , PRECIOTOTAL "
                    Sql = Sql & vbCrLf & " , POSICION , DETALLE  "
                    Sql = Sql & vbCrLf & " FROM FACDET  "
                    Sql = Sql & vbCrLf & " WHERE FACTURAABC ='" & Trim(FacturaABC) & "'"
                    Sql = Sql & vbCrLf & " AND NUMEROFACTURA =" & NumeroFactura
                    Sql = Sql & vbCrLf & " ORDER BY POSICION "
                    
                    
                    Dim cantidad  As Long
                    Dim PRECIOUNITARIO As String
                    Dim PRECIOTOTAL As String
                    Dim Posicion As String
                    Dim detalle  As String
                    
                  Set rsFACDET = New ADODB.Recordset
                  rsFACDET.CursorLocation = adUseClient
                
                Rem ConDetalleFactura
                
                
                
                
                
                If chkPasarTodas.value = 0 Then
                   rsFACDET.Open Sql, conData, adOpenForwardOnly, adLockReadOnly
                Else
                    rsFACDET.Open Sql, ConDetalleFactura, adOpenForwardOnly, adLockReadOnly
                End If
                
                
             If rsFACDET.EOF Then
                MsgBox "No hay Detalle Factura La factura No se pasara"
                Exit Sub
            End If
            
            
            ConFacturaElectronica.Execute SqlFacturaCabecera
                
            
          MAX_ID_FACTURA = ObtenerIDFac(FacturaABC, NumeroFactura, CLng(IDCliente))
            
            If MAX_ID_FACTURA = 0 Then
            Exit Sub
            End If
            
            
        
            
            Do While Not rsFACDET.EOF
                cantidad = rsFACDET!cantidad
                PRECIOUNITARIO = "'" & Replace(rsFACDET!PRECIOUNITARIO, ",", ".") & "'"
                PRECIOTOTAL = "'" & Replace(rsFACDET!PRECIOTOTAL, ",", ".") & "'"
                Posicion = rsFACDET!Posicion
                detalle = "'" & Trim(rsFACDET!detalle) & "'"
                Sql = " INSERT INTO factura_electronica.dbo.FACDET"
                Sql = Sql & vbCrLf & " ( "
                Sql = Sql & vbCrLf & "  FK_ID_FACTURA"
                Sql = Sql & vbCrLf & " ,FACTURAABC"
                Sql = Sql & vbCrLf & " , NUMEROFACTURA"
                Sql = Sql & vbCrLf & " , CANTIDAD"
                Sql = Sql & vbCrLf & " , PRECIOUNITARIO"
                Sql = Sql & vbCrLf & " , PRECIOTOTAL"
                Sql = Sql & vbCrLf & " , POSICION"
                Sql = Sql & vbCrLf & " , DETALLE"
                Sql = Sql & vbCrLf & " )"
                Sql = Sql & vbCrLf & " VALUES ( "
                Sql = Sql & vbCrLf & MAX_ID_FACTURA
                Sql = Sql & vbCrLf & ", '" & FacturaABC & "'"
                Sql = Sql & vbCrLf & " , " & NumeroFactura
                Sql = Sql & vbCrLf & " , " & cantidad
                Sql = Sql & vbCrLf & " , " & PRECIOUNITARIO
                Sql = Sql & vbCrLf & " , " & PRECIOTOTAL
                Sql = Sql & vbCrLf & " , " & Posicion
                Sql = Sql & vbCrLf & " , " & detalle
                Sql = Sql & vbCrLf & " )"
                ConFacturaElectronica.Execute Sql
                rsFACDET.MoveNext
            Loop
            
            
            
'            INSERT INTO factura_electronica.dbo.FACDET
'                                     (FACTURAABC, FACTURAABC, CANTIDAD, PRECIOUNITARIO, PRECIOTOTAL, POSICION, DETALLE, id)
'            VALUES        (N'FACTURAABC', N'FACTURAABC', 0, 0, 0, 1, N'0',)

End Sub
 

Public Function ObtenerIDFac(FacturaABC As String, NumeroFactura As Long, IDCliente As Long) As Long

Dim Sql As String
Dim conFac As New ADODB.Connection

Dim rs As New ADODB.Recordset

conFac.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=False;User ID=sa;Initial Catalog=factura_electronica;Data Source=222.15.19.150"

Sql = " SELECT       ID_FACTURA, FacturaABC, NumeroFactura, FechaFacturacion, IDCliente"
Sql = Sql & " From FACTURA "
Sql = Sql & "  WHERE FacturaABC ='" & FacturaABC & "'"
Sql = Sql & "  AND NumeroFactura = " & NumeroFactura
Sql = Sql & "  AND IDCliente = " & IDCliente


rs.Open Sql, conFac
If rs.EOF Then
ObtenerIDFac = 0
Else
ObtenerIDFac = rs!ID_FACTURA
End If


End Function
