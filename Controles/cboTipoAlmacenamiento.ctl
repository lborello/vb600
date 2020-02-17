VERSION 5.00
Begin VB.UserControl cboTipoAlmacenamiento 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3360
   ScaleHeight     =   375
   ScaleWidth      =   3360
   Begin VB.ComboBox cboTipoAlmacenamiento 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "cboTipoAlmacenamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Const m_def_Valor = 0
    Dim m_Valor As Variant
    Event Click()
Private Sub cboTipoAlmacenamiento_Click()
    RaiseEvent Click
    If cboTipoAlmacenamiento.ListIndex <> -1 Then
        Valor = CInt(cboTipoAlmacenamiento.ItemData(cboTipoAlmacenamiento.ListIndex))
    Else
        Valor = Null
    End If
End Sub

Private Sub UserControl_Initialize()
Dim rsTipoAlmacenamiento As New ADODB.Recordset
 If Not CONBASA Is Nothing Then
    rsTipoAlmacenamiento.Open "Select *  from Tipo_Almacenamiento", CONBASA
    cboTipoAlmacenamiento.Clear
    Do While Not rsTipoAlmacenamiento.EOF
       cboTipoAlmacenamiento.AddItem CStr(rsTipoAlmacenamiento!ID) & " - " & Trim(UCase(CStr(rsTipoAlmacenamiento!DESCRIPCION)))
       cboTipoAlmacenamiento.ItemData(cboTipoAlmacenamiento.ListCount - 1) = rsTipoAlmacenamiento!ID
       rsTipoAlmacenamiento.MoveNext
    Loop
    Set rsTipoAlmacenamiento = Nothing
   End If
End Sub

Private Sub UserControl_Resize()
    cboTipoAlmacenamiento.Width = UserControl.Width
End Sub
Public Property Get Font() As Font
    Set Font = cboTipoAlmacenamiento.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cboTipoAlmacenamiento.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get Valor() As Variant
    If cboTipoAlmacenamiento.ListIndex <> -1 Then
        Valor = CInt(cboTipoAlmacenamiento.ItemData(cboTipoAlmacenamiento.ListIndex))
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
    Set cboTipoAlmacenamiento.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Valor = PropBag.ReadProperty("Valor", m_def_Valor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", cboTipoAlmacenamiento.Font, Ambient.Font)
    Call PropBag.WriteProperty("Valor", m_Valor, m_def_Valor)
End Sub


