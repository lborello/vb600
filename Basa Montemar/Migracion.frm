VERSION 5.00
Object = "{1DB4A93D-668A-4FEF-8676-0C152947E154}#1.0#0"; "Controles2.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin Control.ctlVerImagenes ctlVerImagenes1 
      Height          =   1995
      Left            =   660
      TabIndex        =   4
      Top             =   1740
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   3519
   End
   Begin VB.CommandButton cmdEstadoCajas 
      Caption         =   "Actualizar Estado Cajas"
      Height          =   315
      Left            =   -60
      TabIndex        =   3
      Top             =   1140
      Width           =   3135
   End
   Begin VB.CommandButton cmdGuardaCustodia 
      Caption         =   "Actualizar Guarda y Custodia"
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   780
      Width           =   3135
   End
   Begin VB.CommandButton cmdDevolucionVacias 
      Caption         =   "Actualizar Devolucion Vacias"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   3135
   End
   Begin VB.CommandButton cmdActualizar_Bajas_Cajas 
      Caption         =   "Actualizar Bajas en cajas"
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Bajas_Cajas_Click()
Dim RsBajas As New ADODB.Recordset
Dim Sql As String
Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_DETALLE.DESDE, REMITOS_CUERPO.ID_CLIENTE"
Sql = Sql & " FROM REMITOS_CUERPO INNER JOIN "
Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO "
Sql = Sql & " WHERE (REMITOS_CUERPO.ANULADO IS NULL) "
Sql = Sql & " AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) "
Sql = Sql & " AND (REMITOS_CUERPO.TIPO = 3) "
Sql = Sql & " AND (REMITOS_CUERPO.ID_CLIENTE < 1000)"
Sql = Sql & " ORDER BY REMITOS_CUERPO.ID_CLIENTE, REMITOS_DETALLE.DESDE"

RsBajas.Open Sql, strConBasa

Do While Not RsBajas.EOF
    Sql = " Update Cajas"
    Sql = Sql & "  SET FK_ESTADO =1140 "
    Sql = Sql & "  , FK_REMITO_BAJA = " & RsBajas!NRO_REMITO
      Sql = Sql & "  , FECHA_MODIFICACION =" & RsBajas!Fecha
    Sql = Sql & "  Where FK_CLIENTE = " & RsBajas!id_cliente
     Sql = Sql & "  And NRO_CAJA = " & RsBajas!Desde

    ConBasa.Execute Sql
    RsBajas.MoveNext
Loop






End Sub

Private Sub Command1_Click()
Dim RsBajas As New ADODB.Recordset
Dim Sql As String
Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_DETALLE.DESDE, REMITOS_CUERPO.ID_CLIENTE"
Sql = Sql & " FROM REMITOS_CUERPO INNER JOIN "
Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO "
Sql = Sql & " WHERE (REMITOS_CUERPO.ANULADO IS NULL) "
Sql = Sql & " AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) "
Sql = Sql & " AND (REMITOS_CUERPO.TIPO = 3) "
Sql = Sql & " AND (REMITOS_CUERPO.ID_CLIENTE < 1000)"
Sql = Sql & " ORDER BY REMITOS_CUERPO.ID_CLIENTE, REMITOS_DETALLE.DESDE"

RsBajas.Open Sql, strConBasa

Do While Not RsBajas.EOF
    Sql = " Update Cajas"
    Sql = Sql & "  SET FK_ESTADO =1140 "
    Sql = Sql & "  , FK_REMITO_BAJA = " & RsBajas!NRO_REMITO
      Sql = Sql & "  , FECHA_MODIFICACION =" & RsBajas!Fecha
    Sql = Sql & "  Where FK_CLIENTE = " & RsBajas!id_cliente
     Sql = Sql & "  And NRO_CAJA = " & RsBajas!Desde

    ConBasa.Execute Sql
    RsBajas.MoveNext
Loop
End Sub

Private Sub cmdDevolucionVacias_Click()
Dim RsBajas As New ADODB.Recordset
Dim Sql As String
Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_DETALLE.DESDE, REMITOS_CUERPO.ID_CLIENTE"
Sql = Sql & " FROM REMITOS_CUERPO INNER JOIN "
Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO "
Sql = Sql & " WHERE (REMITOS_CUERPO.ANULADO IS NULL) "
Sql = Sql & " AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) "
Sql = Sql & " AND (REMITOS_CUERPO.TIPO = 4) "
Sql = Sql & " AND (REMITOS_CUERPO.ID_CLIENTE < 1000)"
Sql = Sql & " ORDER BY REMITOS_CUERPO.ID_CLIENTE, REMITOS_DETALLE.DESDE"

RsBajas.Open Sql, strConBasa

Do While Not RsBajas.EOF
    Sql = " Update Cajas"
    Sql = Sql & "  SET FK_ESTADO =1160 "
    Sql = Sql & "  , FK_DEVOLUCON_VACIAS  = " & RsBajas!NRO_REMITO
      Sql = Sql & "  , FECHA_MODIFICACION =" & Format(RsBajas!Fecha, "DD/MM/YYYY")
    Sql = Sql & "  Where FK_CLIENTE = " & RsBajas!id_cliente
     Sql = Sql & "  And NRO_CAJA = " & RsBajas!Desde

    ConBasa.Execute Sql
    RsBajas.MoveNext
Loop
End Sub

Private Sub cmdEstadoCajas_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String
Dim estado As Integer

Sql = " SELECT     COD_CLIENTE, NRO_CAJA, ESTADO, F_MODIFICACION"
Sql = Sql & "  From CONTENEDOR Where (Not (COD_CLIENTE Is Null))"
Sql = Sql & "  ORDER BY COD_CLIENTE, NRO_CAJA"

rs.Open Sql, strConBasa

Do While Not rs.EOF
    Select Case rs!estado
    Case 4
        estado = 1100
    Case 5
        estado = 1110
    Case 2
        estado = 1120
   Case 3
        estado = 1130
   Case Else
    estado = 0
         End Select
    
    If estado <> 0 Then
    Sql = " Update Cajas SET "
    Sql = Sql & "   FK_ESTADO  = " & estado
    If Not IsNull(rs!F_MODIFICACION) Then
        Sql = Sql & " , FECHA_MODIFICACION = " & Format(rs!F_MODIFICACION, "dd/mm/yyyy")
    End If
    
    Sql = Sql & "  Where FK_CLIENTE = " & rs!COD_CLIENTE
    Sql = Sql & "  And NRO_CAJA = " & rs!NRO_CAJA

    ConBasa.Execute Sql
    End If
    
    rs.MoveNext
Loop

End Sub

Private Sub cmdGuardaCustodia_Click()

Dim RsBajas As New ADODB.Recordset
Dim Sql As String
Sql = " SELECT     REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_DETALLE.DESDE,REMITOS_DETALLE.HASTA ,  REMITOS_CUERPO.ID_CLIENTE"
Sql = Sql & " FROM REMITOS_CUERPO INNER JOIN "
Sql = Sql & " REMITOS_DETALLE ON REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO "
Sql = Sql & " WHERE (REMITOS_CUERPO.ANULADO IS NULL) "
Sql = Sql & " AND (REMITOS_CUERPO.COD_TIPO_ALMACENAMIENTO = 0) "
Sql = Sql & " AND (REMITOS_CUERPO.TIPO = 0) "
Sql = Sql & " AND (REMITOS_CUERPO.ID_CLIENTE < 1000)"
Sql = Sql & " ORDER BY REMITOS_CUERPO.ID_CLIENTE, REMITOS_DETALLE.DESDE"

RsBajas.Open Sql, strConBasa

Do While Not RsBajas.EOF
    Sql = " Update Cajas SET "
    Sql = Sql & "   FK_REMITO_CUSTODIA  = " & RsBajas!NRO_REMITO
    Sql = Sql & "  Where FK_CLIENTE = " & RsBajas!id_cliente
    Sql = Sql & "  And NRO_CAJA BETWEEN " & RsBajas!Desde & " AND " & RsBajas!Hasta

    ConBasa.Execute Sql
    RsBajas.MoveNext
Loop
End Sub

Private Sub Form_Load()
Set ConBasa = New ADODB.Connection

ConBasa.Open "Provider=SQLOLEDB.1;Password=21877471;Persist Security Info=True;User ID=sa;Initial Catalog=BasaSistema; Data Source=server-cudea"
End Sub
