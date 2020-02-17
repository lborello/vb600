VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Requerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Requerimiento"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarRequerimiento As Requerimiento


Public Property Get Requerimiento() As Requerimiento
    If mvarRequerimiento Is Nothing Then
        Set mvarRequerimiento = New Requerimiento
    End If


    Set Requerimiento = mvarRequerimiento
End Property


Public Property Set Requerimiento(vData As Requerimiento)
    Set mvarRequerimiento = vData
End Property
Private Sub Class_Terminate()
    Set mvarRequerimiento = Nothing
End Sub



