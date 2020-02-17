VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsCajas"
Attribute VB_Ext_KEY = "Member1" ,"clsLegajos"
Attribute VB_Ext_KEY = "Member2" ,"clsLibros"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarclsCajas As clsCajas
Private mvarclsLegajos As clsLegajos
Private mvarclsLibros As clsLibros


Public Property Get clsLibros() As clsLibros
   Dim clsLibros2   As New clsLibros

End Property


Public Property Set clsLibros(vData As clsLibros)
    Set mvarclsLibros = vData
End Property


   
Public Property Get clsLegajos() As clsLegajos
    Set clsLegajos = mvarclsLegajos
End Property


Public Property Set clsLegajos(vData As clsLegajos)
    Set mvarclsLegajos = vData
End Property



Private Sub Class_Initialize()
    'crear el objeto mclsCajas cuando se crea la clase clsRemitos
    Set mvarclsCajas = New clsCajas
  'crear el objeto mclsLegajos cuando se crea la clase clsRemitos
  Set mvarclsLegajos = New clsLegajos
  'crear el objeto mclsLibros cuando se crea la clase clsRemitos
  Set mvarclsLibros = New clsLibros
End Sub


Public Property Get clsCajas() As clsCajas
    Set clsCajas = mvarclsCajas
End Property


Public Property Set clsCajas(vData As clsCajas)
    Set mvarclsCajas = vData
End Property
Private Sub Class_Terminate()
  Set mvarclsLibros = Nothing
  Set mvarclsLegajos = Nothing
  Set mvarclsCajas = Nothing
End Sub
Public Sub Remto_ADD(NRO_REM_PROV As String, TIPO As Integer, OPERACION As Integer, ESTADO As Integer, _
Fecha As String, ID_CLIENTE As Integer, OBSERVACIONES As String, elementos As Collection, AUDIT_USUARIO As String, _
AUDIT_FECHA As String, COD_TIPO_ALMACENAMIENTO As Integer, COD_USUARIO_CLIENTE As Integer)
'R_cod_cliente As Integer, R_TipoAlmacenamiento As Tipo_Almacenamiento, R_Tipo_Remito As Tipo_Remito, R_TipoEstado As Estado_Sistema_Remito, R_elementos As Collection)
'INSERT INTO REMITOS_CUERPO
'    (NRO_REMITO, NRO_REM_PROV, TIPO, OPERACION, ESTADO,
'    FECHA, ID_CLIENTE, OBSERVACIONES, CANTIDAD,
'    AUDIT_USUARIO, AUDIT_FECHA,
'    COD_TIPO_ALMACENAMIENTO, COD_PERSONAL_ENTREGA,
'    COD_PERSONAL_RECIBE, COD_USUARIO_CLIENTE)
'VALUES (1, '1', 1, 1, 1, TO_DATE(1, 'J'), 1, '1', 1, '1', TO_DATE(1, 'J'),
'    1, 1, 1, 1)

End Sub
