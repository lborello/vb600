VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
'Public Sub GuardaCustodia(NRO_LIBRO_INTERNO As Collection, Cod_CLiente As Integer, TipoOperacion As Remito_Operacion, NRO_REMITO As Long, FechaRemito As String)
'    Dim SQL As String
'        SQL = " Update libros "
'        SQL = SQL & vbCrLf & " Set Estado = 2 "
'        SQL = SQL & vbCrLf & " ,NRO_REMITO = " & NRO_REMITO
'        SQL = SQL & vbCrLf & " WHERE COD_CLIENTE =" & Cod_CLiente
'        SQL = SQL & vbCrLf & " AND ESTADO = 5 "
'        SQL = SQL & vbCrLf & " AND NRO_LIBRO_INTERNO in (" & CollecionASql(NRO_LIBRO_INTERNO) & ")"
'        ExecutarSql SQL
'        Movimiento NRO_REMITO, Guarda_y_Custodia, TipoOperacion, Libro, NRO_LIBRO_INTERNO, Cod_CLiente, FechaRemito
'
'End Sub
'
'Public Sub Consultas(NRO_LIBRO_INTERNO As Collection, Cod_CLiente As Integer, TipoOperacion As Remito_Operacion, NRO_REMITO As Long, FechaRemito As String)
'    Dim SQL As String
'        If TipoOperacion = Remito_Operacion.Salida Then
'            SQL = " Update libros "
'            SQL = SQL & vbCrLf & " Set Estado = 3 "  bb
'            SQL = SQL & vbCrLf & " , NRO_REMITO = " & NRO_REMITO
'            SQL = SQL & vbCrLf & " WHERE COD_CLIENTE =" & Cod_CLiente
'            SQL = SQL & vbCrLf & " AND ESTADO = 2"
'            SQL = SQL & vbCrLf & " AND NRO_LIBRO_INTERNO in (" & CollecionASql(NRO_LIBRO_INTERNO) & ")"
'        End If
'        If TipoOperacion = Remito_Operacion.ENTRADA Then
'            SQL = " Update libros "
'            SQL = SQL & vbCrLf & " Set Estado = 2 "
'            SQL = SQL & vbCrLf & " , NRO_REMITO = " & NRO_REMITO
'            SQL = SQL & vbCrLf & " WHERE COD_CLIENTE =" & Cod_CLiente
'            SQL = SQL & vbCrLf & " AND ESTADO = 3 "
'            SQL = SQL & vbCrLf & " AND NRO_LIBRO_INTERNO in (" & CollecionASql(NRO_LIBRO_INTERNO) & ")"
'        End If
'        ExecutarSql SQL
'        Movimiento NRO_REMITO, Consulta, TipoOperacion, Libro, NRO_LIBRO_INTERNO, Cod_CLiente, FechaRemito
'End Sub
'
'
