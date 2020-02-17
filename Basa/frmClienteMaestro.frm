VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmClienteMaestro 
   Caption         =   "Archivo de Clientes"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   11460
   Begin VB.CommandButton Command2 
      Caption         =   "Refrescar"
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
      TabIndex        =   8
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmdBorrarCliente 
      Caption         =   "Borrar"
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
      Left            =   6000
      TabIndex        =   7
      Top             =   60
      Width           =   1275
   End
   Begin VB.CommandButton cmdCaratulaCliente 
      Caption         =   "Caratula"
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
      Left            =   10140
      TabIndex        =   6
      Top             =   60
      Width           =   1155
   End
   Begin VB.CommandButton cmdNuevoCliente 
      Caption         =   "Nuevo Cliente"
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
      Left            =   8700
      TabIndex        =   5
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
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
      TabIndex        =   4
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Nombre de Cliente"
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
      Left            =   2220
      TabIndex        =   3
      Top             =   60
      Width           =   1875
   End
   Begin VB.TextBox txtFiltroCliente 
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
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   11055
   End
   Begin VB.CommandButton cmdBuscar_id_Cliente 
      Caption         =   "ID Cliente"
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
      TabIndex        =   1
      Top             =   60
      Width           =   1875
   End
   Begin MSDataGridLib.DataGrid grdClientes 
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   900
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   13361
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   18
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
End
Attribute VB_Name = "frmClienteMaestro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Dim rs As New ADODB.Recordset



Private Sub cmdMo_Click()

End Sub

Private Sub cmdBorrarCliente_Click()
    Dim rs2 As New ADODB.Recordset
    Dim Sql As String
        Sql = " SELECT     NRO_REMITO, ID_CLIENTE, FECHA"
        Sql = Sql & " From REMITOS_CUERPO"
        Sql = Sql & "  Where id_cliente = " & rs.Fields.Item(0).value
        rs2.Open Sql, ConActiva, 0, 1
        If rs2.EOF Then
            If MsgBox("¿Esta usted seguro de Barra el cliente  " & vbCrLf & rs.Fields.Item(1).value & "?", vbYesNo) = vbYes Then
                Sql = " DELETE FROM CLIENTES "
                Sql = Sql & "  Where id_cliente =" & rs.Fields.Item(0).value
                ExecutarSql Sql
            End If
        Else
                MsgBox "El cliente ya tiene movimientos consulte con el administrador"
        End If
        rs.Requery
 
End Sub

Private Sub cmdCaratulaCliente_Click()
    Dim Sql As String
    Sql = "  SELECT * "
    Sql = Sql & " FROM   CLIENTES "
     
Sql = Sql & " WHERE     (CARATULA = N'1')"
Sql = Sql & " ORDER BY ID_CLIENTE"
    frmReportes.ImprimirReporte PasoReportes & "CaratulaClientes.rpt", Sql, True
End Sub

Private Sub cmdModificar_Click()
    
  frmCliente.ID_Cliente_Maestro = rs.Fields.Item(0).value
  frmCliente.Show


End Sub

Private Sub cmdNuevoCliente_Click()
    frmCliente.ID_Cliente_Maestro = 0
    frmCliente.Show
End Sub

Private Sub Command2_Click()
Dim Sql As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    Sql = " SELECT     ID_CLIENTE, RAZON_SOCIAL, CALLE, NUMERO, LOCALIDAD, ID_PROVINCIA, TELEFONOS, NRO_CUIT"
    Sql = Sql & " From Clientes "
    Sql = Sql & " ORDER BY ID_CLIENTE "
    
    rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
    rs.Requery
    Set grdClientes.DataSource = rs.DataSource
    grdClientes.DataMember = rs.DataMember
    grdClientes.Rebind
    grdClientes.Refresh

End Sub

Private Sub Form_Load()
    Dim Sql As String

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    
    Sql = " SELECT     ID_CLIENTE, RAZON_SOCIAL, CALLE, NUMERO, LOCALIDAD, ID_PROVINCIA, TELEFONOS, NRO_CUIT"
    Sql = Sql & " From Clientes "
    Sql = Sql & " ORDER BY ID_CLIENTE "
    
    rs.Open Sql, ConActiva, adOpenDynamic, adLockReadOnly
    rs.Requery
    Set grdClientes.DataSource = rs.DataSource
    grdClientes.DataMember = rs.DataMember
    grdClientes.Rebind
    grdClientes.Refresh

End Sub

Private Sub Form_Resize()
    grdClientes.Width = frmClienteMaestro.Width - 500
    grdClientes.Height = frmClienteMaestro.Height - 500 - grdClientes.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo salir
rs.Close
salir:

End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtFiltroCliente_Change()
On Error GoTo salir:

If txtFiltroCliente.Text <> "" Then

rs.Filter = " RAZON_SOCIAL like '%" & txtFiltroCliente.Text & "%'"
Else
    rs.Filter = ""
    rs.Requery
End If
salir:

End Sub
