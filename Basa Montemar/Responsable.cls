VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResponsable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

'variables locales para almacenar los valores de las propiedades
Private mvarApellido As String 'copia local
Private mvarID_Personal As Integer 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarNombre As String 'copia local
Public Property Let Nombre(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Nombre = 5
    mvarNombre = vData
End Property


Public Property Get Nombre() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property



Public Property Let ID_Personal(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ID_Personal = 5
    mvarID_Personal = vData
End Property


Public Property Get ID_Personal() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ID_Personal
    ID_Personal = mvarID_Personal
End Property



Public Property Let Apellido(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Apellido = 5
    mvarApellido = vData
End Property


Public Property Get Apellido() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Apellido
    Apellido = mvarApellido
End Property



