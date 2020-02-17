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
Private mvarCajas As clsCajas
Private mvarLegajos As clsLegajos
Private mvarLibros As clsLibros


Public Property Get Libros() As Libros
    Set clsLibros = mvarLibros
End Property


Public Property Set Libros(vData As clsLibros)
    Set mvarLibros = vData
End Property



Public Property Get Legajos() As clsLegajos
    Set Legajos = mvarLegajos
End Property


Public Property Set Legajos(vData As clsLegajos)
    Set mvarLegajos = vData
End Property



Private Sub Class_Initialize()
    'crear el objeto mCajas cuando se crea la clase Remitos
    Set mvarCajas = New clsCajas
  'crear el objeto mLegajos cuando se crea la clase Remitos
  Set mvarLegajos = New clsLegajos
  'crear el objeto mLibros cuando se crea la clase Remitos
  Set mvarLibros = New clsLibros
End Sub


Public Property Get Cajas() As clsCajas
    Set Cajas = mvarCajas
End Property


Public Property Set Cajas(vData As clsCajas)
    Set mvarCajas = vData
End Property
Private Sub Class_Terminate()
  Set mvarLibros = Nothing
  Set mvarLegajos = Nothing
  Set mvarCajas = Nothing
End Sub



