Attribute VB_Name = "Modulo2"
Option Explicit



Global OraSqlStmt As OraSqlStmt

Global EstadoConexion As Boolean
Global MdiChildCount As Integer
Global A_Tipo(4) As String
Global Ruta As String

'Connection Information
Global UserName$
Global Password$
Global DatabaseName$
Global Connect$
Global ConexionODBC As String

Global RemitoTipo As Integer
Global RemitoCod As Long
Global RemitoNuevo As Boolean
Global IndexPrn As Integer
Global Primero As Integer
Global Ultimo As Integer
Rem Global Prn As New cfgprn

Global RemitoModificado As Boolean
Global FormTitulo As String
Global FormularioActivo As Integer
Global Reportes(3) As String

' Show parameters
Global Const MARGEN_INFERIOR = 250
Global Const MARGEN_SUPERIOR = 50
Global Const MARGEN_DERECHO = 50
Global Const MARGEN_IZQUIERDO = 150
Global Const MARGEN_INTERNO = 25

Global Const CANCELAR = 0
Global Const ACEPTAR = 1
Global Const MODAL = 1
Global Const MODELESS = 0


Sub CenterForm(F As Form)

  ' Center the specified form within the screen

  F.Move (Screen.Width - F.Width) \ 2, (Screen.Height - F.Height) \ 2

End Sub

Function IsDba() As Boolean
  Dim OraRol As ADODB.Recordset
  IsDba = False
  Set OraRol = New ADODB.Recordset
  OraRol.Open "Select * from USER_ROLE_PRIVS WHERE GRANTED_ROLE='DBA'", conbasa
  If Not OraRol.EOF Then IsDba = True
End Function


