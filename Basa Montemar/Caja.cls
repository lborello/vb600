VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Caja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String
 
'local variable(s) to hold property value(s)
Private mvarEstanteria As Long 'local copy
Private mvarHorizontal As Long 'local copy
Private mvarVertical As Long 'local copy
Private mvarAdelante_Atras As Integer 'local copy
Private mvarNRO_ESTANTE As Long 'local copy
Private mvarEstado As Integer 'local copy
Private mvarCod_Cliente As Integer 'local copy
Private mvarNRO_CAJA As Long 'local copy
Public Property Let NRO_CAJA(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NRO_CAJA = 5
    mvarNRO_CAJA = vData
End Property


Public Property Get NRO_CAJA() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NRO_CAJA
  NRO_CAJA = mvarNRO_CAJA
 End Property



Public Property Let COD_CLIENTE(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cod_Cliente = 5
    mvarCod_Cliente = vData
End Property


Public Property Get COD_CLIENTE() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cod_Cliente
    COD_CLIENTE = mvarCod_Cliente
End Property



Public Property Let estado(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Estado = 5
    mvarEstado = vData
End Property


Public Property Get estado() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Estado
    estado = mvarEstado
End Property



Public Property Let NRO_ESTANTE(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.NRO_ESTANTE = 5
    mvarNRO_ESTANTE = vData
End Property


Public Property Get NRO_ESTANTE() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.NRO_ESTANTE
    NRO_ESTANTE = mvarNRO_ESTANTE
End Property



Public Property Let Adelante_Atras(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Adelante_Atras = 5
    mvarAdelante_Atras = vData
End Property


Public Property Get Adelante_Atras() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Adelante_Atras
    Adelante_Atras = mvarAdelante_Atras
End Property



Public Property Let Vertical(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Vertical = 5
    mvarVertical = vData
End Property


Public Property Get Vertical() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Vertical
    Vertical = mvarVertical
End Property



Public Property Let Horizontal(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Horizontal = 5
    mvarHorizontal = vData
End Property


Public Property Get Horizontal() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Horizontal
    Horizontal = mvarHorizontal
End Property



Public Property Let Estanteria(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Estanteria = 5
    mvarEstanteria = vData
End Property


Public Property Get Estanteria() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Estanteria
    Estanteria = mvarEstanteria
End Property



