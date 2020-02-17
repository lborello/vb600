VERSION 5.00
Begin VB.UserControl cboPersonal 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3405
   ScaleHeight     =   360
   ScaleWidth      =   3405
   Begin VB.ComboBox cboPersonal 
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
      ItemData        =   "cboPersonal.ctx":0000
      Left            =   0
      List            =   "cboPersonal.ctx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "cboPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
    Const m_def_Valor = 0
    Dim m_Valor As Variant
    Event Click()


Private Sub cboPersonal_Click()
    RaiseEvent Click
    If cboPersonal.ListIndex <> -1 Then
        Valor = CInt(cboPersonal.ItemData(cboPersonal.ListIndex))
    Else
        Valor = Null
    End If

End Sub

Private Sub cboPersonal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub UserControl_Initialize()
    Dim rsPersonal As New ADODB.Recordset
        If Not CONBASA Is Nothing Then
            rsPersonal.Open "SELECT IDPERSONAL, NOMBRE, APELLIDO FROM PERSONAL WHERE NAVES= '1'", CONBASA
            cboPersonal.Clear
            Do While Not rsPersonal.EOF
               cboPersonal.AddItem CStr(rsPersonal!idPersonal) & " - " & Trim(UCase(CStr(rsPersonal!Apellido))) & " " & Trim(UCase(CStr(rsPersonal!Nombre)))
               cboPersonal.ItemData(cboPersonal.ListCount - 1) = rsPersonal!idPersonal
               rsPersonal.MoveNext
            Loop
            Set rsPersonal = Nothing
        End If
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys vbTab
End If
End Sub

Private Sub UserControl_Resize()
    cboPersonal.Width = UserControl.Width
End Sub
Public Property Get Font() As Font
    Set Font = cboPersonal.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set cboPersonal.Font = New_Font
    PropertyChanged "Font"
End Property
Public Property Get Valor() As Variant
    If cboPersonal.ListIndex <> -1 Then
        Valor = CInt(cboPersonal.ItemData(cboPersonal.ListIndex))
    Else
        Valor = Null
    End If
End Property
Public Property Let Valor(ByVal New_Valor As Variant)
    m_Valor = New_Valor
    Dim I As Integer
    For I = 0 To cboPersonal.ListCount - 1
        If cboPersonal.ItemData(I) = New_Valor Then
            cboPersonal.ListIndex = I
            Exit Property
        End If
    Next
    PropertyChanged "Valor"
End Property
Private Sub UserControl_InitProperties()
    m_Valor = m_def_Valor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   Rem Set cboPersonal.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Valor = PropBag.ReadProperty("Valor", m_def_Valor)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", cboPersonal.Font, Ambient.Font)
    Call PropBag.WriteProperty("Valor", m_Valor, m_def_Valor)
End Sub



