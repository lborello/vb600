VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl cltIndice 
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2310
   ScaleHeight     =   3375
   ScaleWidth      =   2310
   Begin MSComctlLib.TreeView trvIndices 
      Height          =   3315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5847
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   265
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cltIndice.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cltIndice.ctx":0279
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cltIndice.ctx":1F83
            Key             =   "Casa"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cltIndice.ctx":229D
            Key             =   "Documento"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cltIndice.ctx":2523
            Key             =   "Sector"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cltIndice.ctx":28E0
            Key             =   "Documentos"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cltIndice.ctx":2B61
            Key             =   "Legajo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "cltIndice.ctx":343B
            Key             =   "Sucursal"
            Object.Tag             =   "Sucursal"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "cltIndice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
  Option Explicit
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_Appearance = 1
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_Appearance As AppearanceConstants
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
Event DblClick()
Attribute DblClick.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse y después lo vuelve a presionar y liberar sobre un objeto."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Ocurre cuando el usuario presiona una tecla mientras un objeto tiene el enfoque."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Ocurre cuando el usuario libera una tecla mientras un objeto tiene el enfoque."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."
Public Cod_cliente As Integer
Enum TipoIndice
   Sector = 0
   Documentos = 1
   Documento = 2
   Nulo = 3
End Enum


Private Sub trvIndice_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub trvIndices_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub trvIndices_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Initialize()
Inicio
End Sub

Private Sub UserControl_Resize()
    trvIndices.width = UserControl.width
    trvIndices.Height = UserControl.Height
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=28,0,0,1
Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Devuelve o establece si los controles, formularios y formularios MDI se dibujan en tiempo de ejecución con efectos 3D."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indica si un control Label o el color de fondo de un control Shape es transparente u opaco."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Obliga a volver a dibujar un objeto."
     
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_Appearance = m_def_Appearance
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle

End Sub
Public Function ExisItem(dato As String) As Boolean
    Dim s As String
    On Error GoTo ErrorHandler
        ExisItem = True
        s = trvIndices.Nodes.Item(dato)
    Exit Function
ErrorHandler:
    ExisItem = False
End Function
'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
End Sub
Public Function Item_Selecionado() As String
Dim Nodo_1 As Node
    Dim i As Integer
       With trvIndices.Nodes
           For i = 1 To .Count
              If .Item(i).Selected Then
                  Item_Selecionado = Mid(.Item(i).Key, 2)
                  Exit Function
              End If
           Next
        End With
  End Function
Public Function Index_Selecionado() As Integer
Dim Nodo_1 As Node
    Dim i As Integer
       With trvIndices.Nodes
           For i = 1 To .Count
              If .Item(i).Selected Then
                 Index_Selecionado = i
                  Exit Function
              End If
           Next
        End With
  End Function
  Public Function Descripcion() As String
Dim Nodo_1 As Node
    Dim i As Integer
       With trvIndices.Nodes
           For i = 1 To .Count
              If .Item(i).Selected Then
                 Descripcion = .Item(i).Text
                  Exit Function
              End If
           Next
        End With
  End Function

  
Public Function Indice() As String
Dim i As Integer
Dim ItemSel As String
ItemSel = Item_Selecionado
 For i = 1 To Len(ItemSel)
  If Mid(ItemSel, i, 1) = "-" Then
    Indice = Trim(Mid(ItemSel, 1, i - 1))
  End If
  
  
  Next
  End Function
    
  
  Public Function Cliente() As Integer
    Cliente = Cod_cliente
  End Function
Public Function Actualizar(Cod_cliente As Integer, Filtro As TipoIndice, ExpanderIndex As Integer, Optional Filtro_Indice As String)
    Dim Indice0 As String
    Dim KeyTreeView1 As String
    Dim Indice1 As String
    Dim Descripcion As String
    Dim nodX As Node
    Dim rsIndices As New ADODB.Recordset
    Dim sql As String
    Dim Tipo_Indice As String
   Cod_cliente = Cod_cliente
  If Filtro_Indice = "" Then
     sql = " SELECT * From INDICES  Where COD_CLIENTE =" & Cod_cliente
   Else
    sql = " SELECT * From INDICES  Where COD_CLIENTE =" & Cod_cliente
    sql = sql & " AND  (INDICE LIKE '" & Filtro_Indice & "%')"
   End If
   
    Select Case Filtro
    Case TipoIndice.Sector
        sql = sql & " AND TIPO_INDICE ='Sector' "
    End Select
    
    sql = sql & " ORDER BY INDICE"
      rsIndices.Open sql, strConBasa
        
        trvIndices.Nodes.Clear
        Set nodX = trvIndices.Nodes.Add(, , "RAIZ", "TODAS LAS CATEGORIAS", "Casa") ' Root
        trvIndices.Nodes.Item("RAIZ").Tag = "TODOS"
        Do While Not rsIndices.EOF
        Tipo_Indice = Trim(rsIndices!Tipo_Indice)
       
            If ExisItem("R" & Mid(rsIndices!Indice, 1, Len(rsIndices!Indice) - 3)) Then
                KeyTreeView1 = "R" & Mid(rsIndices!Indice, 1, Len(rsIndices!Indice) - 3)
                Descripcion = rsIndices!Indice & " - " & rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!Descripcion)
                Set nodX = trvIndices.Nodes.Add(KeyTreeView1, tvwChild, "R" & rsIndices!Indice, Descripcion, Tipo_Indice, Tipo_Indice)
                trvIndices.Nodes.Item("R" & rsIndices!Indice).Tag = rsIndices!Indice
            Else
                Descripcion = rsIndices!Indice & " - " & rsIndices!ID_CODIGO_DOCUMENTO & " // " & Trim(rsIndices!Descripcion)
                Set nodX = trvIndices.Nodes.Add(, , "R" & rsIndices!Indice, Descripcion, Tipo_Indice, Tipo_Indice) ' Root
                trvIndices.Nodes.Item("R" & rsIndices!Indice).Tag = rsIndices!Indice
            End If
            rsIndices.MoveNext
        Loop
        EXPANDER ExpanderIndex
End Function
Public Function EXPANDER(IndiceNumerico As Integer)
Dim i As Integer
Dim indexs As Integer
indexs = IndiceNumerico
On Error Resume Next
For i = 1 To 10
   indexs = trvIndices.Nodes.Item(indexs).Parent.Index
   trvIndices.Nodes.Item(indexs).Expanded = True
Next
trvIndices.Nodes.Item(IndiceNumerico).Expanded = True
trvIndices.Refresh
End Function
Private Function PonerImagen(Doc As Variant) As Integer
PonerImagen = 0
If Not IsNull(Doc) Then
    Select Case Trim(Doc)
    Case "Sector"
        PonerImagen = 3
    Case "Documento"
        PonerImagen = 2
    Case "Documentos"
        PonerImagen = 4
    Case "Legajo"
        PonerImagen = 5
        
    End Select
End If
End Function
'Public Sub MarcarIndiceFrase(Dato As String, Optional EXPANDER As Boolean)
'    Dim i  As Integer
'    Dim a As Integer
'    Dim B As Integer
'    Dim Indice As String
'
'
'        For i = 1 To trvIndices.Nodes.Count
'            trvIndices.Nodes.Item(i).BackColor = &H80000005
'            If Dato = "" Or Dato = " " Then
'            Else
'                B = InStr(UCase(trvIndices.Nodes.Item(i).Text), "-")
'                If UCase(trvIndices.Nodes.Item(i).Text) <> "TODAS LAS CATEGORIAS" Then
'                   If Mid(Dato, 1, 1) = "0" Then
'                       ' BUSCAR INDICE
'                       Indice = Mid(trvIndices.Nodes.Item(i).Text, 1, B - 2)
'                       If Indice = UCase(Dato) Then
'                            a = 1
'                       Else
'                            a = 0
'                       End If
'                   Else
'                        ' BUSCAR NOMBRE
'                        a = InStr(UCase(trvIndices.Nodes.Item(i).Text), UCase(Dato))
'                    End If
'                    If a = 0 Then
'                      If EXPANDER = True Then
'                        trvIndices.Nodes.Item(i).Expanded = False
'                      End If
'                    Else
''                        trvIndices.Nodes.Item(i).Expanded = True
''                        trvIndices.Nodes.Item(i).Selected = True
''                        trvIndices.Nodes.Item(i).BackColor = &H80000002
''                        trvIndices.Nodes.Item(i).Bold = True
'                    End If
'                End If
'            End If
'        Next
'End Sub
Public Sub BuscarIndice(dato As String, Optional EXPANDER As Boolean)
    Dim i  As Integer
    Dim a As Integer
    Dim B As Integer
    Dim Indice As String
        Dim ImagenLegajo As Integer
        
        For i = 1 To trvIndices.Nodes.Count
            trvIndices.Nodes.Item(i).BackColor = &H80000005
            trvIndices.Nodes.Item(i).ForeColor = &H80000008
            trvIndices.Nodes.Item(i).Bold = False
            If Trim(dato) = "" Then
                Exit Sub
            Else
                    a = InStr(UCase(trvIndices.Nodes.Item(i).Text), UCase(dato))
                    If a <> 0 Then
                              If EXPANDER = True Then
                                    trvIndices.Nodes.Item(i).ForeColor = &H80000002
                                    trvIndices.Nodes.Item(i).Bold = True
                                    trvIndices.Nodes.Item(i).Selected = True
                                    trvIndices.Nodes.Item(i).Expanded = True
                              Else
                                    trvIndices.Nodes.Item(i).ForeColor = &H80000002
                                    trvIndices.Nodes.Item(i).Bold = True
                                    trvIndices.Nodes.Item(i).Selected = True
                                    trvIndices.Nodes.Item(i).Expanded = True
                              End If
                    End If
             End If
                    
        Next
        trvIndices.Refresh
End Sub
Public Sub BuscarTipoIndice(dato As String, Optional EXPANDER As Boolean)
    Dim i  As Integer
    Dim a As Integer
    Dim B As Integer
    Dim Indice As String
        Dim ImagenLegajo As Integer
        
        For i = 1 To trvIndices.Nodes.Count
            trvIndices.Nodes.Item(i).BackColor = &H80000005
            trvIndices.Nodes.Item(i).ForeColor = &H80000008
            trvIndices.Nodes.Item(i).Bold = False
            If trvIndices.Nodes.Item(i).Image = dato Then
                  If EXPANDER = True Then
                        trvIndices.Nodes.Item(i).ForeColor = &H80000002
                        trvIndices.Nodes.Item(i).Bold = True
                        trvIndices.Nodes.Item(i).Selected = True
                        trvIndices.Nodes.Item(i).Expanded = True
                  Else
                        trvIndices.Nodes.Item(i).ForeColor = &H80000002
                        trvIndices.Nodes.Item(i).Bold = True
                        trvIndices.Nodes.Item(i).Selected = True
                        trvIndices.Nodes.Item(i).Expanded = True
                  End If
            End If
            
                    
        Next
        trvIndices.Refresh
End Sub
