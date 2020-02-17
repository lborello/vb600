VERSION 5.00
Begin VB.UserControl cboTipoRequerimientos 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   ScaleHeight     =   390
   ScaleWidth      =   3375
   Begin VB.ComboBox cboTipoRequerimiento 
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
      ItemData        =   "cboTipoRequerimientos.ctx":0000
      Left            =   0
      List            =   "cboTipoRequerimientos.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "cboTipoRequerimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_Valor = 0
'Property Variables:
Dim m_Valor As Variant
Event Click()
Private Sub cboTipoRequerimiento_Click()
    RaiseEvent Click
    If cboTipoRequerimiento.ListIndex <> -1 Then
       Valor = CInt(cboTipoRequerimiento.ItemData(cboTipoRequerimiento.ListIndex))
    Else
       Valor = Null
    End If
End Sub

Private Sub UserControl_Resize()
    cboTipoRequerimiento.Width = UserControl.Width
End Sub
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = cboTipoRequerimiento.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set cboTipoRequerimiento.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get Valor() As Variant
    If cboTipoRequerimiento.ListIndex <> -1 Then
        Valor = CInt(cboTipoRequerimiento.ItemData(cboTipoRequerimiento.ListIndex))
    Else
        Valor = Null
    End If
End Property

Public Property Let Valor(ByVal New_Valor As Variant)
    m_Valor = New_Valor
    Dim I As Integer
    For I = 0 To cboTipoRequerimiento.ListCount - 1
        If cboTipoRequerimiento.ItemData(I) = New_Valor Then
            cboTipoRequerimiento.ListIndex = I
            Exit Property
            PropertyChanged "Valor"
        End If
    Next
    cboTipoRequerimiento.ListIndex = -1
End Property
Private Sub UserControl_InitProperties()
    m_Valor = m_def_Valor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set cboTipoRequerimiento.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Valor = PropBag.ReadProperty("Valor", m_def_Valor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", cboTipoRequerimiento.Font, Ambient.Font)
    Call PropBag.WriteProperty("Valor", m_Valor, m_def_Valor)
End Sub
Public Property Let CargarDatos(ByVal vNewValue As ADODB.Recordset)
    cboTipoRequerimiento.Clear
    Do While Not vNewValue.EOF
       cboTipoRequerimiento.AddItem CStr(vNewValue!ID_TIPO_REQUERIMIENTO) & " - " & Trim(UCase(CStr(vNewValue!DESCRIPCION)))
       cboTipoRequerimiento.ItemData(cboTipoRequerimiento.ListCount - 1) = vNewValue!ID_TIPO_REQUERIMIENTO
       vNewValue.MoveNext
    Loop
    Set vNewValue = Nothing


End Property
