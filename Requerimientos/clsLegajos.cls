VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLegajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'Public Sub GuardaCustodia()
'End Sub
'
'Public Sub Consultas(ID_CLIENTE_LEGAJO As Collection, Cod_CLiente As Integer, TipoOperacion As Remito_Operacion, NRO_REMITO As Long, FechaRemito As String)
'    Dim SQL As String
'    SQL = "  Update LEGAJOS"
'    If TipoOperacion = Remito_Operacion.Salida Then
'        SQL = SQL & vbCrLf & " SET COD_ESTADO = 3 "
'        SQL = SQL & vbCrLf & " ,FECHA_ACTUALIZACION = TO_DATE('" & FechaRemito & "','DD/MM/YYYY')"
'        SQL = SQL & vbCrLf & " , COD_REMITO = " & NRO_REMITO
'        SQL = SQL & vbCrLf & " WHERE ID_CLIENTE_LEGAJO in (" & CollecionASql(ID_CLIENTE_LEGAJO) & ")"
'        SQL = SQL & vbCrLf & " AND COD_CLIENTE = " & Cod_CLiente
'        SQL = SQL & vbCrLf & " AND COD_ESTADO = 2 "
'    End If
'    If TipoOperacion = Remito_Operacion.ENTRADA Then
'        SQL = SQL & vbCrLf & " SET COD_ESTADO = 2 "
'        SQL = SQL & vbCrLf & " ,FECHA_ACTUALIZACION = TO_DATE('" & FechaRemito & "','DD/MM/YYYY')"
'        SQL = SQL & vbCrLf & " , COD_REMITO = " & NRO_REMITO
'        SQL = SQL & vbCrLf & " WHERE ID_CLIENTE_LEGAJO in (" & CollecionASql(ID_CLIENTE_LEGAJO) & ")"
'        SQL = SQL & vbCrLf & " AND COD_CLIENTE = " & Cod_CLiente
'        SQL = SQL & vbCrLf & " AND COD_ESTADO = 3 "
'    End If
'    ExecutarSql SQL
'    Movimiento NRO_REMITO, Consulta, TipoOperacion, Legajo, ID_CLIENTE_LEGAJO, Cod_CLiente, FechaRemito
'
'End Sub
'
'Public Sub Bajas(ID_CLIENTE_LEGAJO As Long, Cod_CLiente As Integer, NRO_REMITO As Long, COD_INDICE As String, COD_CLIENTE_USUARIO As Integer)
'   Dim rsBajas As ADODB.Recordset
'   Dim SQL As String
'        SQL = " INSERT INTO 'BAJASLEGAJOS'"
'        SQL = SQL & vbCrLf & " (ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE,"
'        SQL = SQL & vbCrLf & " CLIENTE_LEGAJO, DESCRIPCION, NRO_CAJA, COD_CLIENTE,"
'        SQL = SQL & vbCrLf & " COD_UBICACION, COD_ESTADO, NOMBRE, COD_REMITO,"
'        SQL = SQL & vbCrLf & " FECHA)"
'        SQL = SQL & vbCrLf & " SELECT ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE,"
'        SQL = SQL & vbCrLf & " CLIENTE_LEGAJO, DESCRIPCION, NRO_CAJA, COD_CLIENTE,"
'        SQL = SQL & vbCrLf & " COD_UBICACION, COD_ESTADO, NOMBRE, COD_REMITO, Fecha"
'        SQL = SQL & vbCrLf & " From Legajos"
'        SQL = SQL & vbCrLf & " Where ID_CLIENTE_LEGAJO in( " & CollecionASql(ID_CLIENTE_LEGAJO) & ")"
'        SQL = SQL & vbCrLf & " And COD_CLIENTE = " & Cod_CLiente
'        ExecutarSql SQL
'        SQL = "  Update LEGAJOS"
'        SQL = SQL & vbCrLf & " SET COD_ESTADO = 0 "
'        SQL = SQL & vbCrLf & " ,FECHA_ACTUALIZACION = TO_DATE('" & FechaRemito & "','DD/MM/YYYY')"
'        SQL = SQL & vbCrLf & " , COD_REMITO = " & NRO_REMITO
'        SQL = SQL & vbCrLf & " WHERE ID_CLIENTE_LEGAJO in (" & CollecionASql(ID_CLIENTE_LEGAJO) & ")"
'        SQL = SQL & vbCrLf & " AND COD_CLIENTE = " & Cod_CLiente
'        ExecutarSql SQL
'        Movimiento NRO_REMITO, Bajas, TipoOperacion, Legajo, ID_CLIENTE_LEGAJO, Cod_CLiente, FechaRemito
'
'
'
'End Sub
'
