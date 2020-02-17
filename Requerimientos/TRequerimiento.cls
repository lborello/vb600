VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TRequerimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"TRequerimientos"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarIDREQUERIMIENTO As Long 'local copy
Private mvarID_CLIENTE As Long 'local copy
Private mvarIDPERSONAL As Integer 'local copy
Private mvarIDTIPORECEPCION As Integer 'local copy
Private mvarIDESTADO As Integer 'local copy
Private mvarIDTIPOREQUERIMIENTO As Integer 'local copy
Private mvarIDFAX As Long 'local copy
Private mvarSECTOR As String 'local copy
Private mvarTELEFONO As String 'local copy
Private mvarDESCRIPCION As String 'local copy
Private mvarSOLICITANTE As String 'local copy
Private mvarTOMO As Integer 'local copy
Private mvarFECHAENTREGA As Date 'local copy
Private mvarFECHALIMITE As Date 'local copy
Private mvarFECHARECEPCION As Date 'local copy
Private mvarCANTIDAD As Integer 'local copy
Private mvarIDREMITO As Long 'local copy
Private mvarTIEMPOTOTAL As String 'local copy
Private mvarTRequerimientos As TRequerimientos

Public Property Get TRequerimientos() As TRequerimientos
    If mvarTRequerimientos Is Nothing Then
        Set mvarTRequerimientos = New TRequerimientos
    End If
    Set TRequerimientos = mvarTRequerimientos
End Property

Public Property Set TRequerimientos(vData As TRequerimientos)
    Set mvarTRequerimientos = vData
End Property

Private Sub Class_Terminate()
    Set mvarTRequerimientos = Nothing
End Sub

Public Property Let TIEMPOTOTAL(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TIEMPOTOTAL = 5
    mvarTIEMPOTOTAL = vData
End Property

Public Property Get TIEMPOTOTAL() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TIEMPOTOTAL
    TIEMPOTOTAL = mvarTIEMPOTOTAL
End Property

Public Property Let IDREMITO(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDREMITO = 5
    mvarIDREMITO = vData
End Property

Public Property Get IDREMITO() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDREMITO
    IDREMITO = mvarIDREMITO
End Property

Public Property Let CANTIDAD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CANTIDAD = 5
    mvarCANTIDAD = vData
End Property

Public Property Get CANTIDAD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CANTIDAD
    CANTIDAD = mvarCANTIDAD
End Property

Public Property Let FECHARECEPCION(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FECHARECEPCION = 5
    mvarFECHARECEPCION = vData
End Property

Public Property Get FECHARECEPCION() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FECHARECEPCION
    FECHARECEPCION = mvarFECHARECEPCION
End Property

Public Property Let FECHALIMITE(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FECHALIMITE = 5
    mvarFECHALIMITE = vData
End Property

Public Property Let FECHAENTREGA(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FECHAENTREGA = 5
    mvarFECHAENTREGA = vData
End Property


Public Property Get FECHAENTREGA() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FECHAENTREGA
    FECHAENTREGA = mvarFECHAENTREGA
End Property



Public Property Let TOMO(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TOMO = 5
    mvarTOMO = vData
End Property

Public Property Get TOMO() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TOMO
    TOMO = mvarTOMO
End Property

Public Property Let SOLICITANTE(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SOLICITANTE = 5
    mvarSOLICITANTE = vData
End Property

Public Property Get SOLICITANTE() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SOLICITANTE
    SOLICITANTE = mvarSOLICITANTE
End Property

Public Property Let DESCRIPCION(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DESCRIPCION = 5
    mvarDESCRIPCION = vData
End Property

Public Property Get DESCRIPCION() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DESCRIPCION
    DESCRIPCION = mvarDESCRIPCION
End Property

Public Property Let TELEFONO(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TELEFONO = 5
    mvarTELEFONO = vData
End Property

Public Property Get TELEFONO() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TELEFONO
    TELEFONO = mvarTELEFONO
End Property

Public Property Let Sector(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SECTOR = 5
    mvarSECTOR = vData
End Property

Public Property Get Sector() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SECTOR
    Sector = mvarSECTOR
End Property

Public Property Let IDFAX(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDFAX = 5
    mvarIDFAX = vData
End Property

Public Property Get IDFAX() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDFAX
    IDFAX = mvarIDFAX
End Property

Public Property Let IDTIPOREQUERIMIENTO(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDTIPOREQUERIMIENTO = 5
    mvarIDTIPOREQUERIMIENTO = vData
End Property


Public Property Get IDTIPOREQUERIMIENTO() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDTIPOREQUERIMIENTO
    IDTIPOREQUERIMIENTO = mvarIDTIPOREQUERIMIENTO
End Property

Public Property Let IDESTADO(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDESTADO = 5
    mvarIDESTADO = vData
End Property

Public Property Get IDESTADO() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDESTADO
    IDESTADO = mvarIDESTADO
End Property

Public Property Let IDTIPORECEPCION(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDTIPORECEPCION = 5
    mvarIDTIPORECEPCION = vData
End Property

Public Property Get IDTIPORECEPCION() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDTIPORECEPCION
    IDTIPORECEPCION = mvarIDTIPORECEPCION
End Property

Public Property Let IDPERSONAL(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDPERSONAL = 5
    mvarIDPERSONAL = vData
End Property

Public Property Get IDPERSONAL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDPERSONAL
    IDPERSONAL = mvarIDPERSONAL
End Property

Public Property Let ID_CLIENTE(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ID_CLIENTE = 5
    mvarID_CLIENTE = vData
End Property

Public Property Get ID_CLIENTE() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ID_CLIENTE
    ID_CLIENTE = mvarID_CLIENTE
End Property

Public Property Let IDREQUERIMIENTO(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IDREQUERIMIENTO = 5
    mvarIDREQUERIMIENTO = vData
End Property

Public Property Get IDREQUERIMIENTO() As Long
Attribute IDREQUERIMIENTO.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IDREQUERIMIENTO
    IDREQUERIMIENTO = mvarIDREQUERIMIENTO
End Property
