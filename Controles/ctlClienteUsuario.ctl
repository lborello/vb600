VERSION 5.00
Begin VB.UserControl ctlClienteUsuario 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3420
   ScaleHeight     =   315
   ScaleWidth      =   3420
   Begin VB.ComboBox cboNombre 
      BackColor       =   &H00FFFFFF&
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
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.TextBox txtidUsuarios 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "ctlClienteUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Cliente As Integer
Public Conexion As ADODB.Connection
Event SectorEncontrado(Sector As String)
Event KeyPress(KeyAscii As Integer)
Event LostFocusClienteUsuario()
'Default Property Values:
Const m_def_Enabled = 0
Const m_def_Valor = Null
'Property Variables:
Dim m_Enabled As Boolean
Dim m_Valor As Variant






Public Function BuscarNombre(dato As String, Cliente As Integer) As Boolean
    Dim rsClienteUsuario As New ADODB.Recordset
    Dim sql As String
    ClearCbo
            sql = " SELECT ID_CLIENTEUSUARIO, COD_CLIENTE, APELLIDO_NOMBRE"
            sql = sql & vbCrLf & " From CLIENTEUSUARIO "
            sql = sql & vbCrLf & " Where (DESHABILITADO IS NULL) and COD_CLIENTE = " & Cliente
            sql = sql & vbCrLf & " AND  APELLIDO_NOMBRE LIKE '%" & UCase(dato) & "%'"
            
            rsClienteUsuario.CursorType = adOpenStatic
            rsClienteUsuario.CursorLocation = adUseClient
            rsClienteUsuario.Open sql, strConBasa
            Select Case rsClienteUsuario.RecordCount
            Case 1
                cboNombre.Text = Trim(rsClienteUsuario!APELLIDO_NOMBRE)
                txtidUsuarios.Text = rsClienteUsuario!ID_CLIENTEUSUARIO
                cboNombre.BackColor = &HC0FFC0
                BuscarNombre = True
               
            Case Is > 1
                
                Dim s As String
                s = cboNombre.Text
                cboNombre.Clear
                cboNombre.BackColor = &HC0FFFF
                Do While Not rsClienteUsuario.EOF
                    cboNombre.AddItem Trim(rsClienteUsuario!APELLIDO_NOMBRE)
                    cboNombre.ItemData(cboNombre.ListCount - 1) = rsClienteUsuario.Fields("ID_CLIENTEUSUARIO").Value
                    rsClienteUsuario.MoveNext
                Loop
                 cboNombre.Text = s
                 cboNombre.SelStart = Len(cboNombre.Text)
            Case 0
                cboNombre.BackColor = &HC0E0FF
                txtidUsuarios.Text = ""
                BuscarNombre = False
            End Select
    
End Function
Public Function LlenarConCliente(L_Cliente As Integer) As Boolean
    Dim rsClienteUsuario As New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Cliente = L_Cliente
    Clear
    cboNombre.BackColor = &HFFFFFF
    sql = " SELECT ID_CLIENTEUSUARIO, COD_CLIENTE, APELLIDO_NOMBRE"
    sql = sql & vbCrLf & " From CLIENTEUSUARIO "
    sql = sql & vbCrLf & " Where  (DESHABILITADO IS NULL) and  COD_CLIENTE = " & Cliente
    sql = sql & vbCrLf & " Order by APELLIDO_NOMBRE "
    rsClienteUsuario.CursorType = adOpenStatic
    rsClienteUsuario.CursorLocation = adUseClient
    rsClienteUsuario.Open sql, strConBasa
    Do While Not rsClienteUsuario.EOF
        cboNombre.AddItem Trim(rsClienteUsuario!APELLIDO_NOMBRE)
        cboNombre.ItemData(cboNombre.ListCount - 1) = rsClienteUsuario.Fields("ID_CLIENTEUSUARIO").Value
        rsClienteUsuario.MoveNext
    Loop
           
End Function



Public Sub Clear()
    txtidUsuarios.Text = ""
    cboNombre.Text = ""
    ClearCbo
End Sub
Private Sub ClearCbo()
   Dim i As Integer
   For i = 0 To cboNombre.ListCount - 1
    cboNombre.RemoveItem 0
   Next
End Sub

Public Property Let Valor2(ByVal New_Valor2 As Variant)
    
End Property

Public Property Get Valor2() As Variant
    If txtidUsuarios.Text = "" Then
        Valor = Null
    Else
        Valor = txtidUsuarios.Text
    End If
End Property
Public Property Get Descripcion() As String
    Descripcion = cboNombre.Text
End Property
Private Sub cboNombre_Click()
    If cboNombre.ListIndex <> -1 Then
        txtidUsuarios.Text = cboNombre.ItemData(cboNombre.ListIndex)
    End If
End Sub

Private Sub cboNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
            Exit Sub
        End If
        If KeyAscii = 13 Then
            
            If cboNombre.Text = "" Then
               LlenarConCliente Cliente
            Else
                SendKeys vbTab
            End If
        End If
        If Len(cboNombre.Text) > 2 Then
            If BuscarNombre(cboNombre.Text, Cliente) = True Then
                KeyAscii = 0
            End If
        Else
           Rem  cboNombre.BackColor = &HC0E0FF
            txtidUsuarios.Text = ""
        End If
End Sub

Private Sub cboNombre_LostFocus()
    RaiseEvent LostFocusClienteUsuario
End Sub

Private Sub txtidUsuarios_LostFocus()
    RaiseEvent LostFocusClienteUsuario
End Sub

Private Sub UserControl_Initialize()
Inicio
End Sub

Private Sub UserControl_LostFocus()
    RaiseEvent LostFocusClienteUsuario
End Sub

Private Sub UserControl_Resize()
    cboNombre.width = UserControl.width - txtidUsuarios.width
End Sub

Public Property Get Sector() As String
    Dim rsBuscarSector As New ADODB.Recordset
    Dim sql As String
        
   If txtidUsuarios.Text <> "" Then
   
        sql = " SELECT INDICES.DESCRIPCION,ID_CODIGO_DOCUMENTO, CLIENTEUSUARIO.COD_INDICE "
        sql = sql & vbCrLf & " From CLIENTEUSUARIO, INDICES"
        sql = sql & vbCrLf & " WHERE CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE (+)"
        sql = sql & vbCrLf & " AND CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE (+) AND"
        sql = sql & vbCrLf & " CLIENTEUSUARIO.ID_CLIENTEUSUARIO =  " & txtidUsuarios.Text
        
sql = "  SELECT     INDICES.DESCRIPCION, INDICES.ID_CODIGO_DOCUMENTO, CLIENTEUSUARIO.COD_INDICE"
sql = sql & vbCrLf & "  FROM         CLIENTEUSUARIO LEFT OUTER JOIN"
sql = sql & vbCrLf & "  INDICES ON CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE AND CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE"
sql = sql & vbCrLf & "  Where CLIENTEUSUARIO.ID_CLIENTEUSUARIO =  " & txtidUsuarios.Text

        
        rsBuscarSector.Open sql, strConBasa
        If Not rsBuscarSector.EOF Then
            If Not IsNull(rsBuscarSector!Descripcion) Then
                Sector = rsBuscarSector!ID_CODIGO_DOCUMENTO & " - " & UCase(Trim(rsBuscarSector!Descripcion))
            Else
                Sector = "No Registra Sector"
            End If
        Else
            Sector = ""
        End If
    Else
        Sector = ""
    End If
End Property
Public Property Get Indice() As String
    Dim rsBuscarSector As New ADODB.Recordset
    Dim sql As String
        
   If txtidUsuarios.Text <> "" Then
   
     
        
        sql = " SELECT     INDICES.DESCRIPCION, INDICES.ID_CODIGO_DOCUMENTO, CLIENTEUSUARIO.COD_INDICE"
 sql = sql & vbCrLf & "  FROM         CLIENTEUSUARIO LEFT OUTER JOIN"
  sql = sql & vbCrLf & "                       INDICES ON CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE AND CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE"
 sql = sql & vbCrLf & "  Where CLIENTEUSUARIO.ID_CLIENTEUSUARIO = " & txtidUsuarios.Text
        
        rsBuscarSector.Open sql, strConBasa
        If Not rsBuscarSector.EOF Then
            If Not IsNull(rsBuscarSector!COD_INDICE) Then
                Indice = rsBuscarSector!COD_INDICE
            Else
                Indice = "No Registra Sector"
            End If
        Else
            Indice = ""
        End If
    Else
        Indice = ""
    End If
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,Null
Public Property Get Valor() As Variant
    If IsNumeric(txtidUsuarios.Text) Then
        m_Valor = txtidUsuarios.Text
        Valor = m_Valor
    Else
        Valor = m_Valor
    End If
End Property

Public Property Let Valor(ByVal New_Valor As Variant)
    Dim i As Integer
    Dim rsNombre As New ADODB.Recordset
    Dim sql As String
   If Not IsNull(New_Valor) Then
        
        cboNombre.Clear
        txtidUsuarios.Text = New_Valor
        sql = " SELECT APELLIDO_NOMBRE"
        sql = sql & vbCrLf & " From CLIENTEUSUARIO "
        sql = sql & vbCrLf & " Where ID_CLIENTEUSUARIO = " & New_Valor
        rsNombre.Open sql, strConBasa
        If Not rsNombre.EOF Then
           cboNombre.AddItem rsNombre!APELLIDO_NOMBRE
           cboNombre.ItemData(0) = txtidUsuarios.Text
           cboNombre.ListIndex = 0
        End If
    Else
        txtidUsuarios.Text = ""
        cboNombre.Clear
    End If
    m_Valor = New_Valor
    PropertyChanged "Valor"
End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_Valor = m_def_Valor
    m_Enabled = m_def_Enabled

End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Valor = PropBag.ReadProperty("Valor", m_def_Valor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Valor", m_Valor, m_def_Valor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = m_Enabled
    cboNombre.Enabled = Enabled
    txtidUsuarios.Enabled = Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    cboNombre.Enabled = m_Enabled
    txtidUsuarios.Enabled = m_Enabled
    PropertyChanged "Enabled"
End Property

