VERSION 5.00
Begin VB.UserControl cboTipoRecepcion 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   ScaleHeight     =   345
   ScaleWidth      =   2490
   Begin VB.ComboBox cboTipoRecepcion 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "cboTipoRecepcion.ctx":0000
      Left            =   0
      List            =   "cboTipoRecepcion.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   2475
   End
End
Attribute VB_Name = "cboTipoRecepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Const m_def_Valor = 0
    Dim m_Valor As Variant
    Event Click()
Private Sub cboTipoRecepcion_Click()
    RaiseEvent Click
    If cboTipoRecepcion.ListIndex <> -1 Then
        Valor = CInt(cboTipoRecepcion.ItemData(cboTipoRecepcion.ListIndex))
    Else
        Valor = Null
    End If
End Sub

Private Sub UserControl_Initialize()
Dim rsTipoRecepcion As New ADODB.Recordset
 If Not CONBASA Is Nothing Then
    rsTipoRecepcion.Open "SELECT IDTIPORECEPCION, DESCRIPCION FROM TIPORECEPCION", CONBASA
    cboTipoRecepcion.Clear
    Do While Not rsTipoRecepcion.EOF
       cboTipoRecepcion.AddItem CStr(rsTipoRecepcion!IDTIPORECEPCION) & " - " & Trim(UCase(CStr(rsTipoRecepcion!DESCRIPCION)))
       cboTipoRecepcion.ItemData(cboTipoRecepcion.ListCount - 1) = rsTipoRecepcion!IDTIPORECEPCION
       rsTipoRecepcion.MoveNext
    Loop
    Set rsTipoRecepcion = Nothing
   End If
End Sub

Private Sub UserControl_Resize()
    cboTipoRecepcion.Width = UserControl.Width
End Sub
Public Property Get Font() As Font
    Set Font = cboTipoRecepcion.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cboTipoRecepcion.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get Valor() As Variant
    If cboTipoRecepcion.ListIndex <> -1 Then
        Valor = CInt(cboTipoRecepcion.ItemData(cboTipoRecepcion.ListIndex))
    Else
        Valor = Null
    End If
End Property

Public Property Let Valor(ByVal New_Valor As Variant)
    m_Valor = New_Valor
End Property

Private Sub UserControl_InitProperties()
    m_Valor = m_def_Valor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set cboTipoRecepcion.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Valor = PropBag.ReadProperty("Valor", m_def_Valor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", cboTipoRecepcion.Font, Ambient.Font)
    Call PropBag.WriteProperty("Valor", m_Valor, m_def_Valor)
End Sub



