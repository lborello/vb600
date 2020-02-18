VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmCajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAJAS"
   ClientHeight    =   7590
   ClientLeft      =   1845
   ClientTop       =   1290
   ClientWidth     =   11505
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11505
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   315
      Left            =   9000
      TabIndex        =   20
      Top             =   7140
      Width           =   1515
   End
   Begin VB.CommandButton cmdID 
      Caption         =   "..."
      Height          =   315
      Left            =   2340
      TabIndex        =   19
      Top             =   1740
      Width           =   315
   End
   Begin VB.TextBox txtID 
      BackColor       =   &H00FFC0FF&
      Height          =   330
      Left            =   840
      TabIndex        =   18
      Top             =   1740
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Lectura Orden"
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
      TabIndex        =   16
      Top             =   1080
      Width           =   1515
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
      Left            =   7080
      TabIndex        =   15
      Top             =   420
      Width           =   1095
   End
   Begin VB.CommandButton cmdDepositoCaja 
      Caption         =   "Marcar Deposito Caja"
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
      Left            =   6480
      TabIndex        =   14
      Top             =   6720
      Width           =   2115
   End
   Begin VB.CommandButton cmdMarcarDeposito 
      Caption         =   "Marcar Deposito"
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
      Left            =   4560
      TabIndex        =   13
      Top             =   6720
      Width           =   1755
   End
   Begin VB.ComboBox cboDeposito 
      BackColor       =   &H80000004&
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
      ItemData        =   "frmCajas.frx":0000
      Left            =   1560
      List            =   "frmCajas.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
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
      Left            =   8460
      TabIndex        =   10
      Top             =   420
      Width           =   1095
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   315
      Left            =   4200
      TabIndex        =   8
      Top             =   1740
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   556
   End
   Begin VB.CommandButton cmdLectura 
      Caption         =   "Lectura"
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
      TabIndex        =   7
      Top             =   420
      Width           =   1095
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
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
      Left            =   3900
      TabIndex        =   6
      Top             =   420
      Width           =   1095
   End
   Begin VB.CommandButton cmdCopiarExcel 
      Caption         =   "Copiar Excel"
      Height          =   315
      Left            =   8940
      TabIndex        =   4
      Top             =   6660
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid grdCajas 
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   2220
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7435
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
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
            Format          =   "0"
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
   Begin VB.TextBox txtCajaHasta 
      BackColor       =   &H00FFFFC0&
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
      Left            =   2040
      TabIndex        =   2
      Top             =   900
      Width           =   1455
   End
   Begin VB.TextBox txtCajaDesde 
      BackColor       =   &H00FFFFC0&
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
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
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
      Index           =   3
      Left            =   360
      TabIndex        =   17
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label Label1 
      Caption         =   "Depósito"
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
      Index           =   2
      Left            =   420
      TabIndex        =   12
      Top             =   6780
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   1
      Left            =   3300
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Caja Hasta:"
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
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Caja Desde :"
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
      Left            =   600
      TabIndex        =   0
      Top             =   420
      Width           =   1095
   End
End
Attribute VB_Name = "frmCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscar_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    On Error GoTo salir
    
        Sql = " SELECT   ID_CAJA ,   FK_CLIENTE, NRO_CAJA,ESTADO,  ESTANTERIA, UB_PROVISORIA, RESPONSABLE, DESCRIPCION, FK_LECTURA, FK_PERSONAL_PLANILLA,"
        Sql = Sql & "  FK_CLIENTES_USUARIO , FK_TIPO_REFERENCIA , "
        Sql = Sql & "  FK_TIPO_REFERENCIA_PERSONAL, TIPO_REFERENCIA_FECHA ,  ROLLO, DIGITO_VERIFICADOR  "
        Sql = Sql & "  From dbo.V_CAJAS"
        If txtCajaDesde.Text <> "" Then
        If txtCajaHasta.Text <> "" Then
                Sql = Sql & "  WHERE     NRO_CAJA BETWEEN " & txtCajaDesde.Text & " AND " & txtCajaHasta.Text
        Else
                    Sql = Sql & "  WHERE     NRO_CAJA IN (" & txtCajaDesde.Text & ")"
        End If
        End If
        
        Sql = Sql & "  ORDER BY NRO_CAJA"
       rs.CursorLocation = adUseClient
       rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
       
       Set grdCajas.DataSource = rs.DataSource
       grdCajas.Refresh
       
       
       
       
salir:



End Sub

Private Sub cmdCambioOsep_Click()
Dim Sql As String
On Error GoTo salir


Dim CAJAS As String

Dim concajas As New ADODB.Connection
concajas.Open strConBasa
CAJAS = InputBox("Ingrese las cajas separadas por ,", , 0)
Dim clienteInicial As Integer
Dim clienteFinal As Integer

clienteInicial = 20
clienteFinal = 77

Sql = " Update dbo.CONTENEDOR"
Sql = Sql & " Set COD_CLIENTE = " & clienteFinal
Sql = Sql & " Where COD_CLIENTE = " & clienteInicial
Sql = Sql & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute Sql



Sql = " Update dbo.cajas "
Sql = Sql & "  Set FK_CLIENTE = " & clienteFinal
Sql = Sql & "  WHERE  FK_CLIENTE =  " & clienteInicial
Sql = Sql & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute Sql


Sql = " Update dbo.REFERENCIAS"
Sql = Sql & " Set COD_CLIENTE = " & clienteFinal
Sql = Sql & " Where COD_CLIENTE = " & clienteInicial
Sql = Sql & " AND (NRO_CAJA IN (" & CAJAS & "))"
concajas.Execute Sql

Sql = "  Update dbo.MOV_CAJAS2 "
Sql = Sql & " Set id_cliente =  " & clienteFinal
Sql = Sql & " Where (Tipo_elemento = 0)"
Sql = Sql & " AND (ELEMENTO IN (" & CAJAS & "))"
Sql = Sql & " AND ID_CLIENTE = " & clienteInicial
concajas.Execute Sql

MsgBox "Terminado"
Exit Sub
salir:

End Sub

Private Sub cmdCopiarExcel_Click()
    CopiarDatosGrilla grdCajas
End Sub

Private Sub cmdDepositoCaja_Click()
'Dim RS As New ADODB.Recordset
'On Error GoTo salir:
'Dim SQL As String
'Dim Caja As Long
'
'Caja = InputBox("Ingrese una caja")
'  If Caja > 100000 Then
'    SQL = " UPDATE CAJAS SET    DEPOSITO = '" & UCase(Trim(cboDeposito.Text)) & "'"
'    SQL = SQL & " Where ID_CAJA = " & Caja
'    ExecutarSql SQL
' End If
'
'Exit Sub
'salir:

End Sub

Private Sub cmdID_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    On Error GoTo salir
    
        Sql = " SELECT   ID_CAJA ,   FK_CLIENTE, NRO_CAJA,ESTADO,  ESTANTERIA, UB_PROVISORIA, RESPONSABLE, DESCRIPCION, FK_LECTURA, FK_PERSONAL_PLANILLA,"
        Sql = Sql & "  FK_CLIENTES_USUARIO , FK_TIPO_REFERENCIA , "
        Sql = Sql & "  FK_TIPO_REFERENCIA_PERSONAL, TIPO_REFERENCIA_FECHA ,  ROLLO  "
        Sql = Sql & "  From dbo.V_CAJAS"
        If txtID.Text <> "" Then
        
                    Sql = Sql & "  WHERE     ID_CAJA IN (" & txtID.Text & ")"
       
        End If
        
        Sql = Sql & "  ORDER BY NRO_CAJA"
       rs.CursorLocation = adUseClient
       rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
       
       Set grdCajas.DataSource = rs.DataSource
       grdCajas.Refresh
       
       
       
       
salir:
End Sub

Private Sub cmdLectura_Click()
    Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim rsLectura As New ADODB.Recordset
    Dim Filtro As String
    
    
    Sql = " SELECT     NUMERO_LECTURA, CAJA, CLIENTE"
    Sql = Sql & "  From LECTURACOLECTOR"
    Sql = Sql & "  Where NUMERO_LECTURA in( " & InputBox("Ingrese el numero de lectura", "Lectura", 0) & ")"
    Sql = Sql & "  ORDER BY CAJA"
    rsLectura.Open Sql, ConActiva, 0, 1
    Do While Not rsLectura.EOF
        If rsLectura!Cliente < 9000 Then
            
                Filtro = Filtro & " Or (NRO_CAJA = " & rsLectura!Caja & " AND FK_CLIENTE = " & rsLectura!Cliente & ") " & vbCrLf
            
        End If
        rsLectura.MoveNext
    Loop
    Sql = "SELECT dbo.V_CAJAS.FK_CLIENTE, dbo.V_CAJAS.NRO_CAJA,dbo.V_CAJAS.ESTADO, "
    Sql = Sql & " dbo.V_CAJAS.ESTANTERIA, dbo.V_CAJAS.UB_PROVISORIA, dbo.V_CAJAS.RESPONSABLE, dbo.V_CAJAS.DESCRIPCION, dbo.V_CAJAS.FK_LECTURA,"
    Sql = Sql & " dbo.V_CAJAS.FK_PERSONAL_PLANILLA , dbo.V_CAJAS.FECHA_PLANILLA, dbo.V_CAJAS.FK_CLIENTES_USUARIO , TIPO_REFERENCIA , FK_ESTADO , FK_REMITO_BAJA"
    Sql = Sql & " FROM  dbo.V_CAJAS"
    Sql = Sql & " Where " & Mid(Trim(Filtro), 3)
    Sql = Sql & " ORDER BY NRO_CAJA "
    rs.CursorLocation = adUseClient
    rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
    Set grdCajas.DataSource = rs.DataSource
    grdCajas.Refresh
       

End Sub

Private Sub cmdLeer_Click()
Dim DATO As String
DATO = Clipboard.GetText
DATO = Replace(DATO, vbCrLf, ",")
txtCajaDesde.Text = Mid(DATO, 1, Len(DATO) - 1)
End Sub

Private Sub cmdMarcarDeposito_Click()
'Dim RS As New ADODB.Recordset
'On Error GoTo salir:
'Dim SQL As String
'
''INSERT INTO ENTRADA
''                      (ELEMENTO, COD_CLIENTE, FECHA, COD_ESTADO, DESCRIPCION, LOTE)
''SELECT      CAJA, CLIENTE, 14 / 10 / 2011 AS Expr1, 0 AS Expr2, 'pARA ENVIO ALSINA' AS Expr3, NUMERO_LECTURA
''From LECTURACOLECTOR
''Where (NUMERO_LECTURA = 1524)
'
'
'
'
'SQL = " SELECT CAJA, CLIENTE, NUMERO_LECTURA, ORDEN"
'SQL = SQL & " From LECTURACOLECTOR "
'SQL = SQL & " Where NUMERO_LECTURA =  " & InputBox("Ingrese el numero de lectura")
'SQL = SQL & " ORDER BY ORDEN"
' RS.Open SQL, strConBasa
'
' Do While Not RS.EOF
'    SQL = " UPDATE CAJAS SET    DEPOSITO = '" & UCase(Trim(cboDeposito.Text)) & "'"
'    SQL = SQL & " Where FK_CLIENTE = " & RS!Cliente
'    SQL = SQL & " And NRO_CAJA = " & RS!Caja
'    ExecutarSql SQL
'    RS.MoveNext
'  Loop
'Exit Sub
'salir:

End Sub

Private Sub Command1_Click()
        Dim rs As New ADODB.Recordset
        Dim Sql As String
        rs.CursorLocation = adUseClient
    On Error GoTo salir
    
        Sql = " SELECT CONTENEDOR.COD_CLIENTE, CONTENEDOR.NRO_CAJA, CAJAS_ESTADO.DESCRIPCION, CONTENEDOR.ESTANTERIA, CONTENEDOR.HORIZONTAL,"
        Sql = Sql & vbCrLf & " CONTENEDOR.Vertical , CONTENEDOR.Adelante_Atras "
        Sql = Sql & vbCrLf & " FROM CONTENEDOR INNER JOIN"
        Sql = Sql & vbCrLf & " CAJAS_ESTADO ON CONTENEDOR.ESTADO = CAJAS_ESTADO.ID_CAJAS_ESTADO"
        Sql = Sql & vbCrLf & " Where CONTENEDOR.COD_CLIENTE =" & ctlCliente.Valor
        Sql = Sql & vbCrLf & " ORDER BY CONTENEDOR.NRO_CAJA"
        
        rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
        Set grdCajas.DataSource = rs.DataSource
        grdCajas.Refresh
        
        Exit Sub
salir:
        MsgBox Err.Description
        
End Sub

Private Sub Command2_Click()
 Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim rsLectura As New ADODB.Recordset
    Dim Filtro As String
        Sql = " SELECT LECTURACOLECTOR.ORDEN , dbo.V_CAJAS.FK_CLIENTE, dbo.V_CAJAS.NRO_CAJA,dbo.V_CAJAS.ESTADO, "
        Sql = Sql & " dbo.V_CAJAS.ESTANTERIA, dbo.V_CAJAS.UB_PROVISORIA, dbo.V_CAJAS.RESPONSABLE, dbo.V_CAJAS.DESCRIPCION, dbo.V_CAJAS.FK_LECTURA,"
        Sql = Sql & " dbo.V_CAJAS.FK_PERSONAL_PLANILLA , dbo.V_CAJAS.FECHA_PLANILLA, dbo.V_CAJAS.FK_CLIENTES_USUARIO , TIPO_REFERENCIA "
        Sql = Sql & " FROM         V_CAJAS INNER JOIN"
        Sql = Sql & " LECTURACOLECTOR ON V_CAJAS.NRO_CAJA = LECTURACOLECTOR.CAJA AND V_CAJAS.FK_CLIENTE = LECTURACOLECTOR.CLIENTE"
        Sql = Sql & " Where LECTURACOLECTOR.NUMERO_LECTURA in( " & InputBox("Ingrese el numero de lectura") & ")"
        Sql = Sql & " ORDER BY LECTURACOLECTOR.ORDEN"
        rs.CursorLocation = adUseClient
        rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
        Set grdCajas.DataSource = rs.DataSource
        grdCajas.Refresh

End Sub

Private Sub Command3_Click()

'Rem Muestra todos los archivos y directorios
'
'Dim sVíaAcceso As String
'
'Dim sDir As String, sValue As String
'
'sDir = "Directorios:"
'
'
'sPath = "c:/lec/lectura/"
'
'sValue = Dir$(sPath + getPathSeparator + "*.*")
'
'Do
'
'If sValue <> "." And sValue <> ".." Then
'
'If (GetAttr(sPath + getPathSeparator + sValue) And 16) > 0 Then
'
'Rem Obtener los directorios
'
'sDir = sDir & Chr(13) & sValue
'
'End If
'
'End If
'
'sValue = Dir$
'
'Loop Until sValue = ""
'
'MsgBox sDir, 0, sPath
End Sub

Private Sub Form_Load()
ctlCliente.TipoControl = Cliente
End Sub

