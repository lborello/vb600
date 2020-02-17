VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.UserControl ctlClientes 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   ScaleHeight     =   360
   ScaleWidth      =   5895
   Begin VB.ComboBox cboRazon_Social 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   540
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   5355
   End
   Begin MSMask.MaskEdBox mskCod_Cliente 
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   635
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "ctlClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Cod_Cliente = 0
Const m_def_Razon_Social = "0"
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_Cod_Cliente As Variant
Dim m_Razon_Social As String
'Event Declarations:
Event Click() 'MappingInfo=cboRazon_Social,cboRazon_Social,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=cboRazon_Social,cboRazon_Social,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
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


Private Sub cboRazon_Social_Click()

    If cboRazon_Social.ListIndex <> -1 Then
      COD_CLIENTE = cboRazon_Social.ItemData(cboRazon_Social.ListIndex)
    End If
    RaiseEvent Click
End Sub

Private Sub cboRazon_Social_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub mskCod_Cliente_GotFocus()
    mskCod_Cliente.SelStart = 1
End Sub

Private Sub mskCod_Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub mskCod_Cliente_LostFocus()
        If mskCod_Cliente.Text <> "" Then
           COD_CLIENTE = mskCod_Cliente.Text
        Else
           COD_CLIENTE = Null
        End If
End Sub

Private Sub UserControl_Initialize()
    Dim rsCliente As New ADODB.Recordset
    Dim I As Integer
        rsCliente.Open "Select * from Clientes order by Razon_social", "Provider=MSDAORA.1;Password=1742;User ID=basa;Data Source=bpdc;Persist Security Info=True"
        Do While Not rsCliente.EOF
            cboRazon_Social.AddItem UCase(Trim(rsCliente.Fields("Razon_Social").Value))
            cboRazon_Social.ItemData(I) = rsCliente.Fields("ID_cliente").Value
            rsCliente.MoveNext
            I = I + 1
        Loop
        cboRazon_Social.Width = UserControl.Width - mskCod_Cliente.Width
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
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
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

Public Property Get COD_CLIENTE() As Variant
    COD_CLIENTE = m_Cod_Cliente
End Property

Public Property Let COD_CLIENTE(ByVal New_Cod_Cliente As Variant)
    Dim I As Integer
       On Error GoTo SALIR
        If Not IsNull(New_Cod_Cliente) Then
            For I = 0 To cboRazon_Social.ListCount
                If cboRazon_Social.ItemData(I) = New_Cod_Cliente Then
                    m_Cod_Cliente = New_Cod_Cliente
                    cboRazon_Social.ListIndex = I
                    mskCod_Cliente.Text = COD_CLIENTE
                    PropertyChanged "Cod_Cliente"
                    Exit Property
                End If
            Next
        Else
SALIR:
            mskCod_Cliente.Text = ""
            cboRazon_Social.ListIndex = -1
            m_Cod_Cliente = Null
            PropertyChanged "Cod_Cliente"
        End If
End Property

Public Property Get Razon_Social() As String
    Razon_Social = cboRazon_Social.Text
End Property

Public Property Let Razon_Social(ByVal New_Razon_Social As String)
    m_Razon_Social = New_Razon_Social
    PropertyChanged "Razon_Social"
End Property

Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_Cod_Cliente = m_def_Cod_Cliente
    m_Razon_Social = m_def_Razon_Social
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
  Rem  Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Cod_Cliente = PropBag.ReadProperty("Cod_Cliente", m_def_Cod_Cliente)
    m_Razon_Social = PropBag.ReadProperty("Razon_Social", m_def_Razon_Social)
End Sub

Private Sub UserControl_Resize()
    cboRazon_Social.Width = UserControl.Width - mskCod_Cliente.Width
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Cod_Cliente", m_Cod_Cliente, m_def_Cod_Cliente)
    Call PropBag.WriteProperty("Razon_Social", m_Razon_Social, m_def_Razon_Social)
End Sub

Private Sub cboRazon_Social_DblClick()
    RaiseEvent DblClick
End Sub

