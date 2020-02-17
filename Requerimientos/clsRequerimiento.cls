VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRequerimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"clsRequerimientos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarclsRequerimientos As clsRequerimientos
'local variable(s) to hold property value(s)
Private mvarFecha As Date 'local copy
Private mvarEstado As Integer 'local copy
Private mvarNumeroRequerimiento As Long 'local copy
Private mvarTipo As Integer 'local copy
Private mvarResponsable As Integer
Public Property Let TIPO(ByVal vData As Integer)
    mvarTipo = vData
End Property


Public Property Get TIPO() As Integer
    TIPO = mvarTipo
End Property
Public Property Let NumeroRequerimiento(ByVal vData As Long)
    mvarNumeroRequerimiento = vData
End Property
Public Property Get NumeroRequerimiento() As Long
    NumeroRequerimiento = mvarNumeroRequerimiento
End Property
Public Property Let ESTADO(ByVal vData As Integer)
    mvarEstado = vData
End Property
Public Property Get ESTADO() As Integer
    ESTADO = mvarEstado
End Property
Public Property Let Fecha(ByVal vData As Date)
    mvarFecha = vData
End Property
Public Property Get Fecha() As Date
    Fecha = mvarFecha
End Property
Public Property Get clsRequerimientos() As clsRequerimientos
    If mvarclsRequerimientos Is Nothing Then
        Set mvarclsRequerimientos = New clsRequerimientos
    End If
    Set clsRequerimientos = mvarclsRequerimientos
End Property
Public Property Set clsRequerimientos(vData As clsRequerimientos)
    Set mvarclsRequerimientos = vData
End Property
Private Sub Class_Terminate()
    Set mvarclsRequerimientos = Nothing
End Sub
'Public Property Get Responsable() As Integer
'    Responsable = mvarResponsable
'End Property
'
'Public Property Let Responsable(ByVal vNewValue As Integer)
'    mvarResponsable = vNewValue
'End Property
