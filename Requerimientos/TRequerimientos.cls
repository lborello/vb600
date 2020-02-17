VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRequerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"TRequerimiento"
Attribute VB_Ext_KEY = "Member0" ,"TRequerimiento"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public Sub Clear()
 Dim i As Integer
For i = 1 To Count
    Remove (1)
Next
End Sub


Public Function Add(IDREQUERIMIENTO As Long, ID_CLIENTE As Long, IDPERSONAL As Integer, IDTIPORECEPCION As Integer, IDESTADO As Integer, IDTIPOREQUERIMIENTO As Integer, IDFAX As Long, Sector As String, TELEFONO As String, DESCRIPCION As String, SOLICITANTE As String, TOMO As Integer, FECHAENTREGA As Date, FECHALIMITE As Date, FECHARECEPCION As Date, CANTIDAD As Integer, IDREMITO As Long, TIEMPOTOTAL As String, Optional sKey As String) As TRequerimiento
    'create a new object
    Dim objNewMember As TRequerimiento
    Set objNewMember = New TRequerimiento


    'set the properties passed into the method
    objNewMember.IDREQUERIMIENTO = IDREQUERIMIENTO
    objNewMember.ID_CLIENTE = ID_CLIENTE
    objNewMember.IDPERSONAL = IDPERSONAL
    objNewMember.IDTIPORECEPCION = IDTIPORECEPCION
    objNewMember.IDESTADO = IDESTADO
    objNewMember.IDTIPOREQUERIMIENTO = IDTIPOREQUERIMIENTO
    objNewMember.IDFAX = IDFAX
    objNewMember.Sector = Sector
    objNewMember.TELEFONO = TELEFONO
    objNewMember.DESCRIPCION = DESCRIPCION
    objNewMember.SOLICITANTE = SOLICITANTE
    objNewMember.TOMO = TOMO
    objNewMember.FECHAENTREGA = FECHAENTREGA
    objNewMember.FECHALIMITE = FECHALIMITE
    objNewMember.FECHARECEPCION = FECHARECEPCION
    objNewMember.CANTIDAD = CANTIDAD
    objNewMember.IDREMITO = IDREMITO
    objNewMember.TIEMPOTOTAL = TIEMPOTOTAL
    Set objNewMember.TRequerimientos = TRequerimientos
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As TRequerimiento
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

