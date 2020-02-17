VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRequerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Collection" ,"clsRequerimiento"
Attribute VB_Ext_KEY = "Member0" ,"clsRequerimiento"
Option Explicit

'local variable to hold collection
Private mCol As Collection
Public Sub CambioEstado(IDEmpleado As Integer, ActualizaContador As Boolean, EstadoInicial As Integer, EstadoFinal As Integer, conexion As ADODB.Connection)
    Dim rs As ADODB.Recordset
    Dim RSH_ESTADO_REQUE As ADODB.Recordset
    Dim sql As String
    Dim FECHARECEPCION As Date
    Dim IDTIPOREQUERIMIENTO As Integer
    Dim i As Integer
    Dim CONTADOR As Integer
    
   
    For i = 1 To Count
            
            ' REQUERIMIENTO
            sql = " UPDATE REQUERIMIENTO SET "
            sql = sql & vbCrLf & " IDESTADO= " & EstadoFinal
            If IDEmpleado <> 0 Then
                sql = sql & vbCrLf & ", IDPERSONAL = " & IDEmpleado
            End If
            sql = sql & vbCrLf & " WHERE idRequerimiento IN  ( " & Item(i).NumeroRequerimiento & ")"
            sql = sql & vbCrLf & " AND IDESTADO = " & Item(i).ESTADO
            conexion.Execute (sql)
            
            ' CONTADOR
            sql = " SELECT max(Contador)AS CONTADOR From  H_ESTADO_REQUE  Where IDRequerimiento = " & Item(i).NumeroRequerimiento
            Set RSH_ESTADO_REQUE = New ADODB.Recordset
             RSH_ESTADO_REQUE.Open sql, ConActiva, 0, 1
            If Not RSH_ESTADO_REQUE.EOF Then
            If IsNull(RSH_ESTADO_REQUE!CONTADOR) Then
                CONTADOR = 1
                Else
                    If ActualizaContador Then
                        CONTADOR = CInt(RSH_ESTADO_REQUE!CONTADOR) + 1
                    Else
                        CONTADOR = CInt(RSH_ESTADO_REQUE!CONTADOR)
                    End If
                End If
            Else
                CONTADOR = 1
            End If
            
            ' H_ESTADO_REQUE
            sql = " INSERT INTO H_ESTADO_REQUE ("
            sql = sql & vbCrLf & " IDREQUERIMIENTO, IDESTADO, IDPERSONAL,"
            sql = sql & vbCrLf & " CONTADOR, FECHA )"
            sql = sql & vbCrLf & "  VALUES ("
            sql = sql & vbCrLf & Item(i).NumeroRequerimiento & "," & EstadoFinal & "," & IDEmpleado & ","
            sql = sql & vbCrLf & CONTADOR & "," & SysDate & ")"
            conexion.Execute (sql)
     Next
 End Sub

Public Sub Clear()
    Dim i As Integer
    For i = 1 To Count
     Remove (1)
    Next
End Sub


Public Function Add(Fecha As Date, ESTADO As Integer, IDREQUERIMIENTO As Long, TIPO As Integer, Optional sKey As String) As clsRequerimiento
    'create a new object
    Dim objNewMember As clsRequerimiento
    Set objNewMember = New clsRequerimiento


    'set the properties passed into the method
    Rem Set objNewMember.clsRequerimientos = clsRequerimientos
    objNewMember.Fecha = Fecha
    objNewMember.ESTADO = ESTADO
    objNewMember.NumeroRequerimiento = IDREQUERIMIENTO
    objNewMember.TIPO = TIPO
   Rem  objNewMember.Responsable = Responsable
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As clsRequerimiento
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

