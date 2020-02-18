VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmImpresionRotulo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión de Rotulos"
   ClientHeight    =   7650
   ClientLeft      =   2025
   ClientTop       =   1290
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimirAna 
      Caption         =   "Impri_A"
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
      Left            =   7740
      TabIndex        =   31
      Top             =   5700
      Width           =   1035
   End
   Begin VB.CommandButton cmdImprimirMiguel 
      Caption         =   "Impri_M"
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
      Left            =   6540
      TabIndex        =   30
      Top             =   5700
      Width           =   1035
   End
   Begin VB.CommandButton cmdCordoba 
      Caption         =   "Córdoba"
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
      Left            =   3120
      TabIndex        =   29
      Top             =   5700
      Width           =   1450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Etiqueta lectura"
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
      Left            =   1620
      TabIndex        =   28
      Top             =   5700
      Width           =   1450
   End
   Begin VB.CommandButton cmdCopiarExcel 
      Caption         =   "Copiar Excel"
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
      Left            =   120
      TabIndex        =   27
      Top             =   5700
      Width           =   1450
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1095
      Left            =   120
      TabIndex        =   26
      Top             =   6360
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   1931
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
   Begin VB.Frame Frame2 
      Height          =   675
      Left            =   60
      TabIndex        =   22
      Top             =   1680
      Width           =   4335
      Begin VB.TextBox txtTomarLectura 
         BackColor       =   &H00C0FFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   1020
         PasswordChar    =   "*"
         TabIndex        =   24
         Top             =   240
         Width           =   1860
      End
      Begin VB.CommandButton cmdColector 
         Caption         =   "Colector"
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
         Left            =   3180
         TabIndex        =   23
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Lectura:"
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
         TabIndex        =   25
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   4620
      TabIndex        =   12
      Top             =   480
      Width           =   4335
      Begin VB.TextBox txtEstanteria 
         Alignment       =   2  'Center
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
         Left            =   1200
         TabIndex        =   21
         Top             =   240
         Width           =   2985
      End
      Begin VB.TextBox txtHorizontal 
         Alignment       =   2  'Center
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
         Left            =   1200
         TabIndex        =   20
         Top             =   600
         Width           =   2985
      End
      Begin VB.TextBox txtVertical 
         Alignment       =   2  'Center
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
         Left            =   1200
         TabIndex        =   19
         Top             =   960
         Width           =   2985
      End
      Begin VB.ComboBox cboFrenteAtras 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmImprecionRotulo.frx":0000
         Left            =   1200
         List            =   "frmImprecionRotulo.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1320
         Width           =   1740
      End
      Begin VB.CommandButton cmdInsertarCaja 
         Caption         =   "Insertar "
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
         Left            =   3000
         TabIndex        =   13
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label Label5 
         Caption         =   "Estanteria"
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
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label6 
         Caption         =   "Horizontal"
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
         TabIndex        =   17
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "Vertical"
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
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "AD/AT:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   720
      End
   End
   Begin VB.Frame fraClienteCaja 
      Height          =   1035
      Left            =   60
      TabIndex        =   6
      Top             =   540
      Width           =   4335
      Begin VB.TextBox txtCaja 
         Alignment       =   2  'Center
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
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   2205
      End
      Begin VB.CommandButton CmdInsertarCajaCliente 
         Caption         =   "Insertar "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3060
         TabIndex        =   7
         Top             =   600
         Width           =   1200
      End
      Begin Controles.cltGenerico ctlCliente 
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   661
      End
      Begin VB.Label Label1 
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
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label5 
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
         Index           =   3
         Left            =   60
         TabIndex        =   10
         Top             =   180
         Width           =   765
      End
   End
   Begin Controles.cltGenerico ctlPersonal 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
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
      Height          =   375
      Left            =   4620
      TabIndex        =   1
      Top             =   5700
      Width           =   1450
   End
   Begin MSFlexGridLib.MSFlexGrid grdImpresion 
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   2460
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   3
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
   Begin VB.Label lblFecha 
      Caption         =   "10/10/2000"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblEntrega 
      Caption         =   "Responsable : "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "frmImpresionRotulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cboCliente_Change()

End Sub

Private Sub cmdCancelar_Click()
    grdImpresion.Clear
    TituloGrilla
End Sub

Private Sub cmdColector_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim Caja As Long
    Dim Cliente As Integer
    Dim Posicion As Integer
        Sql = " SELECT NUMERO_LECTURA, CAJA, CLIENTE, ORDEN From LECTURACOLECTOR "
        Sql = Sql & "Where NUMERO_LECTURA = " & InputBox("Por Favor Ingrese el numero de Lectura ", "Lectura")
        Sql = Sql & " and CLIENTE < 8000"
        Sql = Sql & " ORDER BY ORDEN"
            rs.Open Sql, ConActiva, 0, 1
            Do While Not rs.EOF
                Posicion = CInt(rs!Orden)
                Caja = CLng(rs!Caja)
                Cliente = CInt(rs!Cliente)
                CargarGrilla Caja, Cliente
                rs.MoveNext
            Loop
End Sub

Private Sub cmdCopiarExcel_Click()
    Dim R As New ADODB.Recordset
        R.CursorLocation = adUseClient
        R.Open "SELECT     * From dbo.RECAMBIO_CAJA ORDER BY NRO_CAJA", ConActiva, 0, 1
        Set DataGrid1.DataSource = R.DataSource
        CopiarDatosGrilla DataGrid1
End Sub

Private Sub cmdCordoba_Click()
 Dim Filtro As String
    Dim R As Integer
    Dim C As Integer
    Dim Sql As String
        If grdImpresion.Rows = 2 Then
            Filtro = "    (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(1, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(1, 2) & " )"
            Sql = " INSERT INTO dbo.RECAMBIO_CAJA"
            Sql = Sql & " (FK_CLIENTE, NRO_CAJA, FECHA)"
            Sql = Sql & "  VALUES     (" & grdImpresion.TextMatrix(1, 0) & "," & grdImpresion.TextMatrix(1, 2) & "," & SysDate & ")"
            ExecutarSql Sql
            
        Else
            Filtro = "    (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(1, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(1, 2) & " )"
            For R = 2 To grdImpresion.Rows - 1
                If grdImpresion.TextMatrix(R, 0) <> "" Then
                    Filtro = Filtro & vbCrLf & " OR (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(R, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(R, 2) & " )"
                        Sql = " INSERT INTO dbo.RECAMBIO_CAJA"
                        Sql = Sql & " (FK_CLIENTE, NRO_CAJA, FECHA)"
                        Sql = Sql & "  VALUES     (" & grdImpresion.TextMatrix(R, 0) & "," & grdImpresion.TextMatrix(R, 2) & "," & SysDate & ")"
                        ExecutarSql Sql
                End If
             Next
        End If
        ImprimirRotulosCordoba Filtro
End Sub

Private Sub cmdImprimir_Click()
'    Dim Filtro As String
'    Dim R As Integer
'    Dim C As Integer
'    Dim Sql As String
'        If grdImpresion.Rows = 2 Then
'            Filtro = "    (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(1, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(1, 2) & " )"
'            Sql = " INSERT INTO dbo.RECAMBIO_CAJA"
'            Sql = Sql & " (FK_CLIENTE, NRO_CAJA, FECHA)"
'            Sql = Sql & "  VALUES     (" & grdImpresion.TextMatrix(1, 0) & "," & grdImpresion.TextMatrix(1, 2) & "," & SysDate & ")"
'            ExecutarSql Sql
'
'        Else
'            Filtro = "    (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(1, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(1, 2) & " )"
'            For R = 2 To grdImpresion.Rows - 1
'                If grdImpresion.TextMatrix(R, 0) <> "" Then
'                    Filtro = Filtro & vbCrLf & " OR (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(R, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(R, 2) & " )"
'                        Sql = " INSERT INTO dbo.RECAMBIO_CAJA"
'                        Sql = Sql & " (FK_CLIENTE, NRO_CAJA, FECHA)"
'                        Sql = Sql & "  VALUES     (" & grdImpresion.TextMatrix(R, 0) & "," & grdImpresion.TextMatrix(R, 2) & "," & SysDate & ")"
'                        ExecutarSql Sql
'                End If
'             Next
'        End If
'        ImprimirRotulos Filtro
End Sub

Private Sub cmdImprimirAna_Click()

Dim Filtro As String
    Dim R As Integer
    Dim C As Integer
    Dim Sql As String
        If grdImpresion.Rows = 2 Then
            Filtro = "    (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(1, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(1, 2) & " )"
            Sql = " INSERT INTO dbo.RECAMBIO_CAJA"
            Sql = Sql & " (FK_CLIENTE, NRO_CAJA, FECHA)"
            Sql = Sql & "  VALUES     (" & grdImpresion.TextMatrix(1, 0) & "," & grdImpresion.TextMatrix(1, 2) & "," & SysDate & ")"
            ExecutarSql Sql
            
        Else
            Filtro = "    (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(1, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(1, 2) & " )"
            For R = 2 To grdImpresion.Rows - 1
                If grdImpresion.TextMatrix(R, 0) <> "" Then
                    Filtro = Filtro & vbCrLf & " OR (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(R, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(R, 2) & " )"
                        Sql = " INSERT INTO dbo.RECAMBIO_CAJA"
                        Sql = Sql & " (FK_CLIENTE, NRO_CAJA, FECHA)"
                        Sql = Sql & "  VALUES     (" & grdImpresion.TextMatrix(R, 0) & "," & grdImpresion.TextMatrix(R, 2) & "," & SysDate & ")"
                        ExecutarSql Sql
                End If
             Next
        End If

ImprimirRotulosAna Filtro


End Sub

Private Sub cmdImprimirMiguel_Click()
Dim Filtro As String
    Dim R As Integer
    Dim C As Integer
    Dim Sql As String
        If grdImpresion.Rows = 2 Then
            Filtro = "    (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(1, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(1, 2) & " )"
            Sql = " INSERT INTO dbo.RECAMBIO_CAJA"
            Sql = Sql & " (FK_CLIENTE, NRO_CAJA, FECHA)"
            Sql = Sql & "  VALUES     (" & grdImpresion.TextMatrix(1, 0) & "," & grdImpresion.TextMatrix(1, 2) & "," & SysDate & ")"
            ExecutarSql Sql
            
        Else
            Filtro = "    (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(1, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(1, 2) & " )"
            For R = 2 To grdImpresion.Rows - 1
                If grdImpresion.TextMatrix(R, 0) <> "" Then
                    Filtro = Filtro & vbCrLf & " OR (CONTENEDOR.COD_CLIENTE = " & grdImpresion.TextMatrix(R, 0) & " AND CONTENEDOR.NRO_CAJA = " & grdImpresion.TextMatrix(R, 2) & " )"
                    Sql = " INSERT INTO dbo.RECAMBIO_CAJA"
                    Sql = Sql & " (FK_CLIENTE, NRO_CAJA, FECHA)"
                    Sql = Sql & "  VALUES     (" & grdImpresion.TextMatrix(R, 0) & "," & grdImpresion.TextMatrix(R, 2) & "," & SysDate & ")"
                    ExecutarSql Sql
                End If
             Next
        End If
        ImprimirRotulosMiguel Filtro
End Sub

Private Sub cmdInsertarCaja_Click()
    Dim Sql As String
    Dim Adelante_Atras As Integer
        If txtEstanteria <> "" And txtHorizontal <> "" And txtHorizontal <> "" And cboFrenteAtras.ListIndex <> -1 Then
                Select Case cboFrenteAtras.ListIndex
                Case Is = 1
                    Adelante_Atras = 1
                Case Is = 0
                    Adelante_Atras = 2
                End Select
                Sql = " SELECT * From CONTENEDOR WHERE "
                Sql = Sql & vbCrLf & " Estanteria =" & txtEstanteria
                Sql = Sql & vbCrLf & " AND HORIZONTAL =" & txtHorizontal
                Sql = Sql & vbCrLf & " AND VERTICAL = " & txtVertical
                Sql = Sql & vbCrLf & " AND ADELANTE_ATRAS =" & Adelante_Atras
                Dim rs As New ADODB.Recordset
                rs.Open Sql, ConActiva, 0, 1
                If Not rs.EOF Then
                    If Not IsNull(rs!COD_CLIENTE) Then
                        CargarGrilla CLng(rs!NRO_CAJA), CInt(rs!COD_CLIENTE)
                    Else
                        MsgBox "NO TIENE CAJA"
                    End If
                End If
        End If
End Sub

Private Sub CmdInsertarCajaCliente_Click()
    CargarGrilla txtCaja, ctlCliente.Valor
End Sub

Private Sub Form_Load()
    IdEntrega = 0
    idRecibe = 0
    IdClienteAnterior = 0
    ctlCliente.TipoControl = Cliente
    lblFecha = Format(SysDate2, "dd/mm/yyyy")
    TituloGrilla
    ctlPersonal.TipoControl = Personal
End Sub

Public Sub TituloGrilla()
    grdImpresion.ColWidth(0) = 100
    grdImpresion.ColWidth(1) = 7000
    grdImpresion.ColWidth(2) = 1000
    grdImpresion.ColAlignment(1) = 4
    grdImpresion.ColAlignment(2) = 4
    grdImpresion.TextMatrix(0, 0) = ""
    grdImpresion.TextMatrix(0, 1) = "Cliente"
    grdImpresion.TextMatrix(0, 2) = "Cajas"
    grdImpresion.Rows = 2
    grdImpresion.Cols = 3
End Sub
Public Sub CargarGrilla(Caja As Long, Cliente As Integer)
    Dim C As Integer
    Dim R As Integer
    Dim rsCliente As New ADODB.Recordset
'        For R = 1 To grdImpresion.Rows - 1
'            If grdImpresion.TextMatrix(R, 2) <> "" Then
'                If grdImpresion.TextMatrix(R, 2) = Caja Then
'
'                    MsgBox "La caja " & Caja & " esta reperida"
'                    Exit Sub
'                End If
'            End If
'        Next
        rsCliente.Open "Select * from Clientes where id_cliente= " & Cliente, ConActiva, 0, 1
        grdImpresion.TextMatrix(grdImpresion.Rows - 1, 0) = Cliente
        grdImpresion.TextMatrix(grdImpresion.Rows - 1, 1) = Trim(UCase(rsCliente!RAZON_SOCIAL))
        grdImpresion.TextMatrix(grdImpresion.Rows - 1, 2) = Caja
        grdImpresion.AddItem ""
        Set rsCliente = Nothing
End Sub
Public Function Validar() As Boolean
    Validar = True
    If Trim(lblIDCliente) = "" Then
        MsgBox "Falta  el cliente"
        Validar = False
        Exit Function
    End If
    If Trim(lblIDPersonal) = "" Then
        MsgBox "Falta el responsable"
        Validar = False
         Exit Function
    End If
    If grdGuardiayCustodia.Rows < 1 Then
        MsgBox "no tiene caja"
        Validar = False
         Exit Function
    End If
    If grdGuardiayCustodia.TextMatrix(0, 1) <> "NRO_CAJA" Then
        MsgBox "usted debe derificar las posiocnes"
        Validar = False
         Exit Function
    End If
End Function
Private Sub MMControl1_PlayClick(Cancel As Integer)
    Dim j
    j = 0
End Sub

Private Sub txtCaja_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       CmdInsertarCajaCliente_Click
       txtCaja = ""
    End If
End Sub

Private Sub txtTomarLectura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case UCase(Mid(txtTomarLectura.Text, 1, 3))
        Case "CAN"
'            If lblCantidadTotal <> "" Then
'               Rem Hablar "C" & lblCantidadTotal.Caption
'            Else
'              Rem   Hablar "C" & "0"
'            End If
        Case "P01"
'            Dim rsPersonal As OraDynaset
'                lblIDPersonal = CInt(Mid(txtTomarLectura, 4))
'                Set rsPersonal = CONBASA.CreateDynaset("Select * from Personal where idpersonal =" & CInt(lblIDPersonal), ORADYN_READONLY)
'                If Not rsPersonal.EOF Then
'                    lblEntregaNombre = UCase(CStr(rsPersonal!Apellido) & "  " & CStr(rsPersonal!Nombre))
'                End If
        Case Else
            Dim Cliente As Integer
            Dim Caja As Long
            If txtTomarLectura <> "" And Len(txtTomarLectura) > 16 Then
                Caja = Mid(txtTomarLectura.Text, Len(txtTomarLectura.Text) - 5)
                Cliente = Mid(txtTomarLectura.Text, Len(txtTomarLectura.Text) - 8, 3)
                CargarGrilla CLng(Caja), Cliente
            End If
        End Select
        txtTomarLectura = ""
        txtTomarLectura.SetFocus
   End If
End Sub
Public Sub ImprimirRotulosMiguel(Filtro As String)
   Dim Sql As String
   On Error GoTo LERROR
       Sql = "  SELECT * "
       Sql = Sql & vbCrLf & " From "
       Sql = Sql & vbCrLf & " CONTENEDOR "
       Sql = Sql & vbCrLf & " Where " & Filtro
       Sql = Sql & vbCrLf & "  order by estanteria , vertical , horizontal"
       frmReportes.ImprimirReporte PasoReportes & "rotulo.rpt", Sql, True
    Exit Sub
LERROR:
    MsgBox "REPITA LA OPERACION"
End Sub

Public Sub ImprimirRotulosAna(Filtro As String)
   Dim Sql As String
   On Error GoTo LERROR
       Sql = "  SELECT * "
       Sql = Sql & vbCrLf & " From "
       Sql = Sql & vbCrLf & " CONTENEDOR "
       Sql = Sql & vbCrLf & " Where " & Filtro
       Sql = Sql & vbCrLf & "  order by estanteria , vertical , horizontal"
       frmReportes.ImprimirReporte PasoReportes & "rotulo_ana.rpt", Sql, True
    Exit Sub
LERROR:
    MsgBox "REPITA LA OPERACION"
End Sub

Public Sub ImprimirRotulosCordoba(Filtro As String)
   Dim Sql As String
   On Error GoTo LERROR
       Sql = "  SELECT * "
       Sql = Sql & vbCrLf & " From "
       Sql = Sql & vbCrLf & " CONTENEDOR "
       Sql = Sql & vbCrLf & " Where " & Filtro
       Sql = Sql & vbCrLf & "  order by estanteria , vertical , horizontal"
       frmReportes.ImprimirReporte PasoReportes & "Rotulo_Etiqueta.rpt", Sql, True
    Exit Sub
LERROR:
    MsgBox "REPITA LA OPERACION"
End Sub
