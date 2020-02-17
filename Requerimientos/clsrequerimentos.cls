VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsrequerimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsrequerimiento"
Attribute VB_Ext_KEY = "Member0" ,"clsrequerimiento"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public Sub Clear()
Dim i As Integer
    For i = 0 To Count - 1
        Remove (1)
    Next

End Sub

Public Sub LoadRequerimiento(IDquerimiento As Long)
 Dim i
 For i = 100 To 120
 Add "10/10/1999", 1, CStr(i)
 Next



End Sub


Public Function Add(Fecha As Date, Tipo As Variant, Optional sKey As String) As clsrequerimiento
    'create a new object
    
    Dim objNewMember As clsrequerimiento
    Set objNewMember = New clsrequerimiento


    'set the properties passed into the method
    objNewMember.Fecha = Fecha
    If IsObject(Tipo) Then
        Set objNewMember.Tipo = Tipo
    Else
        objNewMember.Tipo = Tipo
    End If
    Rem Set objNewMember.clsrequerimentos = clsrequerimentos
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsrequerimiento
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

