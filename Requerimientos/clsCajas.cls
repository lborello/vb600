VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'Public Sub Vacias(Cajas As Collection, Cod_CLiente As Integer, NRO_REMITO As Long)
'    Dim SQL As String
'    Dim i As Integer
'        SQL = " Update CONTENEDOR"
'        SQL = SQL & vbCrLf & " SET ESTADO =5 , NRO_REMITO =" & NRO_REMITO & ", F_MODIFICACION =" & SysDate & " ,"
'        SQL = SQL & vbCrLf & " Where ESTADO =4  AND COD_CLIENTE = " & Cod_CLiente & "  And NRO_CAJA IN(" & CollecionASql(Cajas) & ")"
'        ExecutarSql SQL
'        Movimiento NRO_REMITO, Cajas_vacias, Salida, Caja, NRO_CAJA, Cod_CLiente, Date
'End Sub
'
'Public Sub DevolVacias(NRO_CAJA As Collection, Cod_CLiente As Integer, NRO_REMITO As Long)
'        Dim SQL As String
'        Dim i As Integer
'        SQL = " Update CONTENEDOR"
'        SQL = SQL & vbCrLf & " SET ESTADO =4 , NRO_REMITO =" & NRO_REMITO & ", F_MODIFICACION =" & SysDate & " ,"
'        SQL = SQL & vbCrLf & " Where ESTADO =5 AND COD_CLIENTE = " & Cod_CLiente & "  And NRO_CAJA IN(" & CollecionASql(Cajas) & ")"
'        ExecutarSql SQL
'        Movimiento NRO_REMITO, Devol_Cajas_Vacias, ENTRADA, Caja, NRO_CAJA, Cod_CLiente, Date
'End Sub
'
'Public Sub GuardaCustodia(NRO_CAJA As Collection, Cod_CLiente As Integer, NRO_REMITO As Long, COD_INDICE As String, COD_CLIENTE_USUARIO As Integer)
'    Dim SQL As String
'    Dim i As Integer
'        SQL = " Update CONTENEDOR"
'        SQL = SQL & vbCrLf & " SET ESTADO =2 , NRO_REMITO =" & NRO_REMITO & ", F_MODIFICACION =" & SysDate & " ,"
'        SQL = SQL & vbCrLf & " Where ESTADO =5 AND COD_CLIENTE = " & Cod_CLiente & "  And NRO_CAJA IN(" & CollecionASql(Cajas) & ")"
'        ExecutarSql SQL
'        Movimiento NRO_REMITO, Guarda_y_Custodia, TipoOperacion, Caja, NRO_CAJA, Cod_CLiente, Date
'End Sub
'
'Public Sub Consultas(Cajas As Collection, Cod_CLiente As Integer, TipoOperacion As Remito_Operacion, NRO_REMITO As Long)
'    Dim SQL As String
'    Dim i As Integer
'         If TipoOperacion = ENTRADA Then
'             SQL = " Update CONTENEDOR"
'             SQL = SQL & vbCrLf & " SET ESTADO = 2 , NRO_REMITO =" & NRO_REMITO & ", F_MODIFICACION =" & SysDate & " ,"
'             SQL = SQL & vbCrLf & " Where ESTADO =3 AND COD_CLIENTE = " & Cod_CLiente & "  And NRO_CAJA IN(" & CollecionASql(Cajas) & ")"
'        End If
'        If TipoOperacion = Salida Then
'             SQL = " Update CONTENEDOR"
'             SQL = SQL & vbCrLf & " SET ESTADO =3 , NRO_REMITO =" & NRO_REMITO & ", F_MODIFICACION =" & SysDate & " ,"
'             SQL = SQL & vbCrLf & " Where ESTADO =2 AND COD_CLIENTE = " & Cod_CLiente & "  And NRO_CAJA IN(" & CollecionASql(Cajas) & ")"
'        End If
'        ExecutarSql SQL
'        Movimiento NRO_REMITO, Consulta, TipoOperacion, Caja, NRO_CAJA, Cod_CLiente, Date
'End Sub
'
'Public Sub Bajas(Cajas As Collection, Cod_CLiente As Integer, NRO_REMITO As Long, COD_INDICE As String, COD_CLIENTE_USUARIO As Integer)
'    Dim SQL As String
'        SQL = "  INSERT INTO 'Baja_Cajas'"
'        SQL = SQL & vbCrLf & "  (ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
'        SQL = SQL & vbCrLf & " ESTADO, COD_CLIENTE, NRO_CAJA, NRO_REMITO)"
'        SQL = SQL & vbCrLf & " SELECT ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
'        SQL = SQL & vbCrLf & " Estado , COD_CLIENTE, NRO_CAJA, NRO_REMITO"
'        SQL = SQL & vbCrLf & " From CONTENEDOR"
'        SQL = SQL & vbCrLf & " Where COD_CLIENTE = " & Cod_CLiente
'        SQL = SQL & vbCrLf & " And NRO_CAJA IN(" & CollecionASql(Cajas) & ")"
'        ExecutarSql SQL
'
'        SQL = "  Update REFERENCIAS"
'        SQL = SQL & vbCrLf & "  Set BORRADO = 1"
'        SQL = SQL & vbCrLf & " Where COD_CLIENTE = " & Cod_CLiente
'        SQL = SQL & vbCrLf & " And NRO_CAJA IN(" & CollecionASql(Cajas) & ")"
'
'        ExecutarSql SQL
'        SQL = "  Update CONTENEDOR"
'        SQL = SQL & vbCrLf & " SET COD_CLIENTE = NULL, NRO_CAJA = NULL, ESTADO = 1,"
'        SQL = SQL & vbCrLf & " NRO_REMITO = NULL, IDREQUERIMIENTO = NULL, F_MODIFICACION = Null"
'        SQL = SQL & vbCrLf & " WHERE COD_CLIENTE = " & Cod_CLiente & " AND  NRO_CAJA IN(" & CollecionASql(Cajas) & ")"
'        ExecutarSql
'        Movimiento NRO_REMITO, Bajas, Salida, Caja, Cajas, Cod_CLiente, Date
'End Sub
'
'
