VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAsignacionCajasCarga 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CAJAS DE LEGAJOS PARA CARGAR"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14805
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   14805
   Begin VB.CheckBox chkLimpiarUsuarioCarga 
      Caption         =   "Limpiar Usuario Carga"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdCopiar 
      Caption         =   "X"
      Height          =   315
      Left            =   10080
      TabIndex        =   13
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdLeer 
      Caption         =   "Leer"
      Height          =   315
      Left            =   12960
      TabIndex        =   11
      Top             =   120
      Width           =   1035
   End
   Begin VB.TextBox txtCajasLectura 
      Height          =   330
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   7935
   End
   Begin VB.CommandButton cmdReferenciasCargadas 
      Caption         =   "Ref. Carga"
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
      Left            =   2880
      TabIndex        =   9
      Top             =   1020
      Width           =   1035
   End
   Begin VB.CommandButton cmdOrdenCarro 
      Caption         =   "O. Carro"
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
      Left            =   11820
      TabIndex        =   8
      Top             =   120
      Width           =   1035
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
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
      Left            =   1740
      TabIndex        =   7
      Top             =   1020
      Width           =   1035
   End
   Begin VB.CommandButton cmdActualizarTipo 
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
      Height          =   315
      Left            =   10080
      TabIndex        =   6
      Top             =   540
      Width           =   1035
   End
   Begin VB.ComboBox cboTipoReferencia 
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
      ItemData        =   "frmAsignacionCajasCarga.frx":0000
      Left            =   1680
      List            =   "frmAsignacionCajasCarga.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   540
      Width           =   3435
   End
   Begin VB.CommandButton cmdActializar 
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
      Height          =   315
      Left            =   10680
      TabIndex        =   4
      Top             =   120
      Width           =   1035
   End
   Begin VB.TextBox txtOrden 
      BackColor       =   &H00C0FFC0&
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
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   555
   End
   Begin MSFlexGridLib.MSFlexGrid grdGrilla 
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   13785
      _Version        =   393216
      Cols            =   7
      BackColorSel    =   12648384
      MergeCells      =   1
      AllowUserResizing=   1
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
   Begin VB.CommandButton cmdRefrescar 
      Caption         =   "Actualizar Lista"
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
      Left            =   4020
      TabIndex        =   0
      Top             =   1020
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo Referencia"
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
      Left            =   180
      TabIndex        =   12
      Top             =   600
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Orden"
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
      Left            =   240
      TabIndex        =   3
      Top             =   180
      Width           =   615
   End
End
Attribute VB_Name = "frmAsignacionCajasCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActializar_Click()
    Dim I As Integer
    Dim SQL As String
    
    If txtCajasLectura.Text <> "" Then
                SQL = " Update basasql.dbo.CAJAS"
                SQL = SQL & " Set ORDEN_CARGA = " & TxtOrden.Text
                SQL = SQL & " Where ID_CAJA in( " & txtCajasLectura.Text & ")"
                ExecutarSql SQL
                txtCajasLectura.Text = ""
                MsgBox "Terminado"
                Exit Sub
    End If
    
    
    
    
    For I = 1 To grdGrilla.Rows - 1
    
     If grdGrilla.TextMatrix(I, 1) = "SI" Then
            If IsNumeric(TxtOrden.Text) Then
                SQL = " Update basasql.dbo.CAJAS"
                SQL = SQL & " Set ORDEN_CARGA = " & TxtOrden.Text
                SQL = SQL & " Where ID_CAJA = " & grdGrilla.TextMatrix(I, 0)
                ExecutarSql SQL
            End If
      End If
     
    
    Next
    MsgBox "Terminado"
End Sub

Private Sub cmdActualizarTipo_Click()
    Dim I As Integer
    Dim SQL As String
    
    If txtCajasLectura.Text <> "" Then
        SQL = " Update basasql.dbo.CAJAS "
        SQL = SQL & vbCrLf & " SET FK_TIPO_REFERENCIA = " & Mid(cboTipoReferencia.Text, 1, 4)
        SQL = SQL & vbCrLf & ", FK_TIPO_REFERENCIA_PERSONAL = " & MDIfrmInicio.StaInicio.Panels(2).Text
        SQL = SQL & vbCrLf & ", TIPO_REFERENCIA_FECHA = " & SysDate
        SQL = SQL & vbCrLf & " Where ID_CAJA in( " & txtCajasLectura.Text & ")"
        ExecutarSql SQL
        txtCajasLectura.Text = ""
        MsgBox "Terminado"
        Exit Sub
    End If
    
    
    
    For I = 1 To grdGrilla.Rows - 1
        If grdGrilla.TextMatrix(I, 1) = "SI" Then
            SQL = " Update basasql.dbo.CAJAS "
            SQL = SQL & vbCrLf & " SET FK_TIPO_REFERENCIA = " & Mid(cboTipoReferencia.Text, 1, 4)
            SQL = SQL & vbCrLf & " , FK_TIPO_REFERENCIA_PERSONAL = " & MDIfrmInicio.StaInicio.Panels(2).Text
            SQL = SQL & vbCrLf & " , TIPO_REFERENCIA_FECHA = " & SysDate
            If chkLimpiarUsuarioCarga.value = 1 Then
                SQL = SQL & vbCrLf & " , FK_PERSONAL_LEGAJO = NULL"
            End If
            SQL = SQL & vbCrLf & " Where ID_CAJA = " & grdGrilla.TextMatrix(I, 0)
            ExecutarSql SQL
        End If
    Next
    MsgBox "Terminado"
        
     


End Sub

Private Sub cmdCopiar_Click()
txtCajasLectura.Text = Replace(Clipboard.GetText, vbCrLf, ",")
txtCajasLectura.Text = Mid(txtCajasLectura.Text, 1, Len(txtCajasLectura.Text) - 1)
End Sub

Private Sub cmdExcel_Click()
    CopiarDatosGrillaMSg grdGrilla
End Sub

Private Sub cmdLeer_Click()
Dim L As String
Dim I As Integer
Dim DATO As String
Dim datoInicio As String
Dim espacio As Integer
Dim comienzo As Integer

On Error GoTo salir

L = Clipboard.GetText
L = Trim(L)
comienzo = 1
espacio = 1
L = Replace(L, vbCrLf, ",")
txtCajasLectura.Text = Mid(L, 1, Len(L) - 1)
Exit Sub
salir:
MsgBox Err.Description
End Sub

Private Sub cmdOrdenCarro_Click()
Dim I As Integer
    Dim SQL As String
    For I = 1 To grdGrilla.Rows - 1
        If IsNumeric(grdGrilla.TextMatrix(I, 9)) Then
            SQL = " Update basasql.dbo.CAJAS "
            SQL = SQL & vbCrLf & " SET Orden_Carro = " & grdGrilla.TextMatrix(I, 9)
            SQL = SQL & vbCrLf & " Where ID_CAJA = " & grdGrilla.TextMatrix(I, 0)
            ExecutarSql SQL
        Else
        SQL = " Update basasql.dbo.CAJAS "
            SQL = SQL & vbCrLf & " SET Orden_Carro = NULL"
            SQL = SQL & vbCrLf & " Where ID_CAJA = " & grdGrilla.TextMatrix(I, 0)
            ExecutarSql SQL
        End If
    Next
    MsgBox "Terminado"
End Sub

Private Sub cmdReferenciasCargadas_Click()
Dim SQL As String
Dim RS As New ADODB.Recordset
MousePointer = 11
SQL = " SELECT    CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"
SQL = SQL & " FROM         CAJAS INNER JOIN"
SQL = SQL & " REFERENCIAS ON CAJAS.FK_CLIENTE = REFERENCIAS.COD_CLIENTE AND CAJAS.NRO_CAJA = REFERENCIAS.NRO_CAJA"
SQL = SQL & " WHERE     (CAJAS.FK_TIPO_REFERENCIA IN (1001, 1002, 1004, 1005))"
SQL = SQL & " GROUP BY CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"

RS.Open SQL, strConBasa

Do While Not RS.EOF
    SQL = " Update basasql.dbo.CAJAS"
    SQL = SQL & "  Set FK_TIPO_REFERENCIA = 1006"
    SQL = SQL & " Where ID_CAJA = " & RS!ID_CAJA
    ExecutarSql SQL
    RS.MoveNext
Loop



SQL = " SELECT     CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"
SQL = SQL & " FROM         CAJAS INNER JOIN"
SQL = SQL & "                       LEGAJOS ON CAJAS.NRO_CAJA = LEGAJOS.NRO_CAJA AND CAJAS.FK_CLIENTE = LEGAJOS.COD_CLIENTE"
SQL = SQL & " WHERE     (CAJAS.FK_TIPO_REFERENCIA IN (1001, 1002, 1004, 1005))"
SQL = SQL & " GROUP BY CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CAJAS.NRO_CAJA"

Set RS = New ADODB.Recordset

RS.Open SQL, strConBasa
Do While Not RS.EOF
    SQL = " Update basasql.dbo.CAJAS"
    SQL = SQL & "  Set FK_TIPO_REFERENCIA = 1020"
    SQL = SQL & " Where ID_CAJA = " & RS!ID_CAJA
    ExecutarSql SQL
    RS.MoveNext
Loop

MousePointer = 0

End Sub

Private Sub cmdRefrescar_Click()

grdGrilla.Clear
grdGrilla.Rows = 1
grdGrilla.Cols = 13
grdGrilla.ColWidth(0) = 100
grdGrilla.ColWidth(1) = 500
grdGrilla.ColWidth(2) = 800
grdGrilla.ColWidth(3) = 4500
grdGrilla.ColWidth(4) = 800
grdGrilla.ColWidth(5) = 1000
grdGrilla.ColWidth(6) = 800
grdGrilla.ColWidth(7) = 800
grdGrilla.ColWidth(8) = 800
grdGrilla.ColWidth(9) = 800
grdGrilla.ColWidth(10) = 1400
grdGrilla.ColWidth(11) = 1400
grdGrilla.ColWidth(12) = 1900

grdGrilla.ColAlignment(0) = 1
grdGrilla.ColAlignment(1) = 1
grdGrilla.ColAlignment(2) = 1
grdGrilla.ColAlignment(3) = 1
grdGrilla.ColAlignment(4) = 1
grdGrilla.ColAlignment(5) = 1
grdGrilla.ColAlignment(6) = 1
grdGrilla.ColAlignment(7) = 1
grdGrilla.ColAlignment(8) = 1
grdGrilla.ColAlignment(9) = 1
grdGrilla.ColAlignment(10) = 1
grdGrilla.ColAlignment(11) = 1
grdGrilla.ColAlignment(12) = 1



grdGrilla.TextMatrix(0, 1) = "Selec."
grdGrilla.TextMatrix(0, 2) = "Estado"
grdGrilla.TextMatrix(0, 3) = "Cliente"

grdGrilla.TextMatrix(0, 4) = "Caja"
grdGrilla.TextMatrix(0, 5) = "Fecha"
grdGrilla.TextMatrix(0, 6) = "Orden"
grdGrilla.TextMatrix(0, 7) = "Personal"
grdGrilla.TextMatrix(0, 8) = "Lugar"
grdGrilla.TextMatrix(0, 9) = "Orden Carro"
grdGrilla.TextMatrix(0, 10) = "NRO_REM_PROV"
grdGrilla.TextMatrix(0, 11) = "FECHA_REMITO"
grdGrilla.TextMatrix(0, 12) = "APELLIDO_NOMBRE"


Dim SQL As String
Dim RS As New ADODB.Recordset




SQL = " SELECT CAJAS.ID_CAJA, CAJAS.FK_CLIENTE, CLIENTES.RAZON_SOCIAL, CAJAS.NRO_CAJA, CONVERT(CHAR, CAJAS.TIPO_REFERENCIA_FECHA, 103) AS FECHA_ASIG,"
SQL = SQL & vbCrLf & " CAJAS.ORDEN_CARGA, CONTENEDOR.ESTADO, CAJAS.FK_PERSONAL_LEGAJO, CAJAS.FK_TIPO_REFERENCIA, CAJAS.ORDEN_CARRO,"
SQL = SQL & vbCrLf & " REMITOS_CUERPO.NRO_REM_PROV, CONVERT(CHAR, REMITOS_CUERPO.FECHA, 103) AS FECHA_REMITO, CLIENTEUSUARIO.APELLIDO_NOMBRE"
SQL = SQL & vbCrLf & " FROM CAJAS INNER JOIN"
SQL = SQL & vbCrLf & " CLIENTES ON CAJAS.FK_CLIENTE = CLIENTES.ID_CLIENTE INNER JOIN"
SQL = SQL & vbCrLf & " CONTENEDOR ON CAJAS.FK_CLIENTE = CONTENEDOR.COD_CLIENTE AND CAJAS.NRO_CAJA = CONTENEDOR.NRO_CAJA LEFT OUTER JOIN"
SQL = SQL & vbCrLf & " REMITOS_CUERPO ON CAJAS.FK_REMITO_CUSTODIA = REMITOS_CUERPO.NRO_REMITO LEFT OUTER JOIN"
SQL = SQL & vbCrLf & " CLIENTEUSUARIO ON REMITOS_CUERPO.COD_USUARIO_CLIENTE = CLIENTEUSUARIO.ID_CLIENTEUSUARIO"
SQL = SQL & vbCrLf & " WHERE  "

If Mid(cboTipoReferencia, 1, 4) = "1010" Or Mid(cboTipoReferencia, 1, 4) = "1015" Then
    SQL = SQL & "   (CAJAS.FK_TIPO_REFERENCIA in( 1010  , 1015))"
Else
    SQL = SQL & "   CAJAS.FK_TIPO_REFERENCIA = " & Mid(cboTipoReferencia, 1, 4)
End If
Rem SQL = SQL & " AND  (CAJAS.FK_TIPO_REFERENCIA_PERSONAL IN (19, 69, 17, 47, 31, 84,83,46)) "
SQL = SQL & "  ORDER BY CAJAS.ORDEN_CARGA desc,FK_TIPO_REFERENCIA DESC , ORDEN_CARRO ,CLIENTES.ID_CLIENTE "

 RS.Open SQL, strConBasa
grdGrilla.Enabled = True
Do While Not RS.EOF
 
 
    grdGrilla.AddItem RS!ID_CAJA & vbTab & "NO" & vbTab & Trim(RS!estado) & vbTab & _
    Trim(RS!FK_CLIENTE) & " - " & Trim(RS!RAZON_SOCIAL) & vbTab & Trim(RS!NRO_CAJA) & _
    vbTab & Trim(RS!FECHA_ASIG) & vbTab & Trim(RS!ORDEN_CARGA) & vbTab & _
    Trim(RS!FK_PERSONAL_LEGAJO) & vbTab & Trim(RS!FK_TIPO_REFERENCIA) & _
    vbTab & RS!ORDEN_CARRO & vbTab & Trim(RS!NRO_REM_PROV) & vbTab & _
    Trim(RS!FECHA_REMITO) & vbTab & Trim(RS!APELLIDO_NOMBRE)

     If RS!estado <> 2 Then
 
        grdGrilla.Col = 2
        grdGrilla.Row = grdGrilla.Rows - 1
        grdGrilla.CellBackColor = &H8080FF
 
 End If
    
    RS.MoveNext
Loop


End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    CargarTipoReferencias
    cboTipoReferencia.ListIndex = 6
    cmdActializar.Enabled = True
If MDIfrmInicio.StaInicio.Panels(2).Text = 47 And MDIfrmInicio.StaInicio.Panels(2).Text = 48 Then
cmdActializar.Enabled = True
End If

End Sub
Public Sub CargarTipoReferencias()
    Dim SQL As String
    Dim RS As New ADODB.Recordset
            SQL = "SELECT      ID_PARAMETRO, DESCRIPCION, TABLA, CAMPO_NOMBRE"
            SQL = SQL & " From basasql.dbo.PARAMETROS "
            SQL = SQL & " WHERE     (CAMPO_NOMBRE = 'FK_TIPO_REFERENCIA')"
            SQL = SQL & " ORDER BY ID_PARAMETRO "
            RS.Open SQL, strConBasa
            Do While Not RS.EOF
                cboTipoReferencia.AddItem RS!ID_PARAMETRO & "-" & Trim(RS!Descripcion)
                RS.MoveNext
            Loop
End Sub

Private Sub Form_Resize()
On Error GoTo salir
   grdGrilla.Width = frmAsignacionCajasCarga.Width - 200
   grdGrilla.Height = frmAsignacionCajasCarga.Height - 2000
salir:
End Sub

Private Sub grdGrilla_DblClick()
grdGrilla.Col = 1
 If grdGrilla.Text = "NO" Then
    grdGrilla.Text = "SI"
 Else
    grdGrilla.Text = "NO"
 End If
 
End Sub

Private Sub grdGrilla_KeyDown(KeyCode As Integer, Shift As Integer)
If grdGrilla.Col = 9 Then
 If KeyCode = 8 Then
 grdGrilla.TextMatrix(grdGrilla.Row, 9) = ""
 Else
    grdGrilla.TextMatrix(grdGrilla.Row, 9) = grdGrilla.TextMatrix(grdGrilla.Row, 9) & Chr(KeyCode)
End If
End If

End Sub

