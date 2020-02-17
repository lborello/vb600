VERSION 5.00
Begin VB.UserControl ctlClienteUsuario 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   ScaleHeight     =   405
   ScaleWidth      =   3480
   Begin VB.ComboBox cboNombre 
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
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.TextBox txtidUsuarios 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
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
Attribute VB_Exposed = False
Option Explicit
Public Cliente As Integer
Public Conexion As ADODB.Connection
Event SectorEncontrado(Sector As String)


Public Function BuscarNombre(dato As String) As Boolean
    Dim rsClienteUsuario As New ADODB.Recordset
    Dim SQL As String
        
            
            SQL = " SELECT ID_CLIENTEUSUARIO, COD_CLIENTE, APELLIDO_NOMBRE"
            SQL = SQL & vbCrLf & " From CLIENTEUSUARIO "
            SQL = SQL & vbCrLf & " Where COD_CLIENTE = " & Cliente
            SQL = SQL & vbCrLf & " AND  APELLIDO_NOMBRE LIKE '%" & UCase(dato) & "%'"
            
            rsClienteUsuario.CursorType = adOpenStatic
            rsClienteUsuario.CursorLocation = adUseClient
            rsClienteUsuario.Open SQL, Conexion
            Select Case rsClienteUsuario.RecordCount
            Case 1
                cboNombre.Text = Trim(rsClienteUsuario!APELLIDO_NOMBRE)
                txtidUsuarios.Text = rsClienteUsuario!ID_CLIENTEUSUARIO
                cboNombre.BackColor = &HC0FFC0
                BuscarNombre = True
                BuscarSector (rsClienteUsuario!ID_CLIENTEUSUARIO)
            Case Is > 1
                
                Dim s As String
                s = cboNombre.Text
                cboNombre.Clear
                cboNombre.BackColor = &HC0FFFF
                Do While Not rsClienteUsuario.EOF
                    cboNombre.AddItem Trim(rsClienteUsuario!APELLIDO_NOMBRE)
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

Public Sub Clear()
    txtidUsuarios.Text = ""
    cboNombre.Text = ""
End Sub

Public Property Get Valor() As Integer
    If txtidUsuarios.Text = "" Then
        Valor = 0
    Else
        Valor = txtidUsuarios.Text
    End If
End Property

Private Sub cboNombre_Click()
    BuscarNombre cboNombre.Text
End Sub

Private Sub cboNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
            Exit Sub
        End If
        If KeyAscii = 13 Then
           SendKeys vbTab
        End If
        If Len(cboNombre.Text) > 2 Then
            If BuscarNombre(cboNombre.Text) = True Then
                KeyAscii = 0
            End If
        Else
            cboNombre.BackColor = &HC0E0FF
            txtidUsuarios.Text = ""
        End If
End Sub

Private Sub UserControl_Resize()
    cboNombre.Width = UserControl.Width - txtidUsuarios.Width
End Sub

Private Sub BuscarSector(ID_CLIENTEUSUARIO As Integer)
Dim rsBuscarSector As New ADODB.Recordset
Dim SQL As String
SQL = " SELECT INDICES.DESCRIPCION"
SQL = SQL & vbCrLf & "  From CLIENTEUSUARIO, INDICES"
SQL = SQL & vbCrLf & "  WHERE CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE AND"
     SQL = SQL & vbCrLf & " CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE AND"
    SQL = SQL & vbCrLf & "  CLIENTEUSUARIO.ID_CLIENTEUSUARIO = " & ID_CLIENTEUSUARIO
With rsBuscarSector
    .Open SQL, Conexion
    If Not .EOF Then
         RaiseEvent SectorEncontrado(Trim(!DESCRIPCION))
    End If
End With


End Sub
