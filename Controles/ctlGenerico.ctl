VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl cltGenerico 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   ScaleHeight     =   315
   ScaleWidth      =   5895
   Begin VB.ComboBox cboDescripcion 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   540
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   5355
   End
   Begin MSMask.MaskEdBox mskID 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "cltGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_DESCRIPCION = "0"
Const m_def_Valor = Null
Const m_def_TipoControl = 0
'Const m_def_Descripcion = 0
Const m_def_ID = 0
'Const m_def_TipoControl = 0
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'Const m_def_Cod_Cliente = 0
'Const m_def_Razon_Social = "0"
'Property Variables:
Dim m_DESCRIPCION As String
Dim m_Valor As Variant
Dim m_TipoControl As Variant
'Dim m_Descripcion As Variant
Dim m_ID As Variant
Public Enum E_TipoControl
   Cliente = 0
   Personal = 1
   Tipo_Remito_Operacion = 2
   Tipo_Remito_Estados = 3
   TIPO_REMITO = 4
   Tipo_Remito_almacenamiento = 5
   Tipo_Requerimiento = 6
   Tipo_Recepcion = 7
   Tipo_Documento = 8
End Enum
'Dim m_TipoControl As E_TipoControl
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'Dim m_Cod_Cliente As Variant
'Dim m_Razon_Social As String
'Event Declarations:
Event Click() 'MappingInfo=cboDescripcion,cboDescripcion,-1,Click
Event DblClick() 'MappingInfo=cboDescripcion,cboDescripcion,-1,DblClick
'Event Click()
'Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Private Sub cboDescripcion_Click()
    If cboDescripcion.ListIndex <> -1 Then
      mskID = cboDescripcion.ItemData(cboDescripcion.ListIndex)
      Valor = mskID.Text
    End If
    RaiseEvent Click
End Sub

Private Sub cboDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub mskID_GotFocus()
    mskID.SelStart = 0
End Sub

Private Sub mskID_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Dim i As Integer
        
       
       If mskID.ClipText = "" Then
        MsgBox "Error en ingreso", vbCritical
        cboDescripcion.ListIndex = -1
        Exit Sub
       End If
        
        For i = 0 To cboDescripcion.ListCount - 1
        If mskID.Text = cboDescripcion.ItemData(i) Then
        cboDescripcion.ListIndex = i
            Valor = mskID.Text
        Exit For
        End If
        
        
        Next
        SendKeys vbTab
        
    End If
End Sub

Private Sub mskID_LostFocus()
        If mskID.Text <> "" Then
           ID = mskID.Text
        Else
           ID = Null
        End If
       
End Sub

Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    mskID.Enabled = m_Enabled
    cboDescripcion.Enabled = m_Enabled
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    mskID.Enabled = m_Enabled
    cboDescripcion.Enabled = m_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub UserControl_Initialize()
    Inicio
End Sub

Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
'    m_Cod_Cliente = m_def_Cod_Cliente
'    m_Razon_Social = m_def_Razon_Social
'    m_Descripcion = m_def_Descripcion
    m_ID = m_def_ID
'    m_TipoControl = m_def_TipoControl
    m_TipoControl = m_def_TipoControl
    m_Valor = m_def_Valor
    m_DESCRIPCION = m_def_DESCRIPCION

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
  Rem  Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
'   m_Descripcion = PropBag.ReadProperty("Descripcion", m_def_Descripcion)
    m_ID = PropBag.ReadProperty("ID", m_def_ID)
'Rem    m_TipoControl = PropBag.ReadProperty("TipoControl", m_def_TipoControl)
 Rem    m_TipoControl = PropBag.ReadProperty("TipoControl", m_def_TipoControl)
    m_Valor = PropBag.ReadProperty("Valor", m_def_Valor)
    m_DESCRIPCION = PropBag.ReadProperty("DESCRIPCION", m_def_DESCRIPCION)
End Sub

Private Sub UserControl_Resize()
    cboDescripcion.width = UserControl.width - mskID.width
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
'    Call PropBag.WriteProperty("Descripcion", m_Descripcion, m_def_Descripcion)
    Call PropBag.WriteProperty("ID", m_ID, m_def_ID)
'Rem     Call PropBag.WriteProperty("TipoControl", m_TipoControl, m_def_TipoControl)
    Call PropBag.WriteProperty("TipoControl", m_TipoControl, m_def_TipoControl)
    Call PropBag.WriteProperty("Valor", m_Valor, m_def_Valor)
    Call PropBag.WriteProperty("DESCRIPCION", m_DESCRIPCION, m_def_DESCRIPCION)
End Sub

Private Sub cboDescripcion_DblClick()
    RaiseEvent DblClick
End Sub
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=14,0,0,0
'Public Property Get DESCRIPCION() As Variant
'    DESCRIPCION = m_Descripcion
'End Property
'
'Public Property Let DESCRIPCION(ByVal New_Descripcion As Variant)
'    m_Descripcion = New_Descripcion
'    PropertyChanged "Descripcion"
'End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get ID() As Variant
    ID = m_ID
End Property

Public Property Let ID(ByVal New_ID As Variant)
    m_ID = New_ID
    m_Valor = New_ID
    PropertyChanged "ID"
End Property
Public Property Get TipoControl() As E_TipoControl
    TipoControl = m_TipoControl
End Property

Public Property Let TipoControl(ByVal New_TipoControl As E_TipoControl)
       
    m_TipoControl = New_TipoControl
     Dim rsCliente As New ADODB.Recordset
    Dim i As Integer
    Dim sql As String
    
    Select Case m_TipoControl
    
    Case E_TipoControl.Cliente
        sql = "SELECT ID_CLIENTE AS ID, Razon_Social AS Descripcion From Clientes ORDER BY Razon_social"
    Case E_TipoControl.Personal
        sql = " SELECT IDPERSONAL AS ID, RTRIM(NOMBRE) + ' ' + RTRIM(APELLIDO) AS DESCRIPCION, NOMBRE From PERSONAL  WHERE   (ACTIVO = 1) ORDER BY NOMBRE"
    Case E_TipoControl.Tipo_Remito_almacenamiento
        sql = " SELECT ID, DESCRIPCION From Tipo_Almacenamiento ORDER BY ID"
    Case E_TipoControl.TIPO_REMITO
        sql = " SELECT ID, DESCRIPCION From TIPO_REMITO ORDER BY ID"
    Case E_TipoControl.Tipo_Remito_Operacion
        sql = " SELECT ID, DESCRIPCION From TIPO_OPERACION ORDER BY ID"
    Case E_TipoControl.Tipo_Remito_Estados
        sql = " SELECT ID, DESCRIPCION From REMITO_ESTADOS ORDER BY ID"
    Case E_TipoControl.Tipo_Requerimiento
        sql = "SELECT IDTIPOREQUERIMIENTO AS ID, DESCRIPCION From TIPOREQUERIMIENTO ORDER BY IDTIPOREQUERIMIENTO"
    Case E_TipoControl.Tipo_Recepcion
        sql = " SELECT IDTIPORECEPCION AS ID, DESCRIPCION From TIPORECEPCION ORDER BY IDTIPORECEPCION"
    Case E_TipoControl.Tipo_Documento
        sql = "SELECT ID, DESCRIPCION From TIPO_DOCUMENTO ORDER BY ID "
   End Select
   rsCliente.CursorLocation = adUseClient
        rsCliente.Open sql, strConBasa, 0, 1
        Do While Not rsCliente.EOF
            cboDescripcion.AddItem UCase(Trim(rsCliente.Fields("Descripcion").Value))
            cboDescripcion.ItemData(i) = rsCliente.Fields("ID").Value
            rsCliente.MoveNext
            i = i + 1
        Loop
        cboDescripcion.width = UserControl.width - mskID.width
    
    
    PropertyChanged "TipoControl"
End Property



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get Valor() As Variant
    Valor = m_Valor
End Property

Public Property Let Valor(ByVal New_Valor As Variant)
    m_Valor = New_Valor
    Dim i As Integer
  If IsNull(New_Valor) Then
        cboDescripcion.ListIndex = -1
        mskID.Text = ""
        PropertyChanged "Valor"
        Exit Property
  End If
   
   For i = 0 To cboDescripcion.ListCount - 1
      If m_Valor = cboDescripcion.ItemData(i) Then
        cboDescripcion.ListIndex = i
        mskID.Text = Valor
      End If
    Next
    PropertyChanged "Valor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,1,1,0
Public Property Get Descripcion() As String
    Rem DESCRIPCION = m_DESCRIPCION
    Descripcion = cboDescripcion.Text
End Property

Public Property Let Descripcion(ByVal New_DESCRIPCION As String)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_DESCRIPCION = New_DESCRIPCION
    PropertyChanged "DESCRIPCION"
End Property

