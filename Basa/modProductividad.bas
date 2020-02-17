Attribute VB_Name = "modProductividad"
Option Explicit

  Public Enum TipoTarea
    COLOCAR_Cajas = 1
    BUSCAR_Cajas = 2
    REFERENCIA_SISTEMA = 3
    REFERENCIA_MANUAL = 4
    BUSCAR_DOCUMENTACION = 5
    ESPECIAL = 6
    PEGADO_DE_ROTULOS = 7
    EXPEDIENTE_SISTEMA = 8
  End Enum

Public Sub InsertarProducion(ID_Personal As Integer, ID_TIPOTAREA As Integer, Elemento As String, cantidad As Integer, COD_CLIENTE As Integer, Optional Fecha As String)

Dim RS As ADODB.Recordset
Dim UNIDADPRODUCION As Double
Dim SQL As String
Set RS = New ADODB.Recordset
Dim Fecha1 As String

If (Fecha) <> "" Then
 Fecha1 = FechaServerTipo(Fecha)
Else
  Fecha1 = SysDate
End If

RS.Open "SELECT  FACTORMULTIPLICACION From TipoTarea Where ID_TIPOTAREA = " & ID_TIPOTAREA, ConActiva, 0, 1
If Not RS.EOF Then
    UNIDADPRODUCION = RS!FACTORMULTIPLICACION * cantidad
Else
    MsgBox "ERROR TAREA"
End If
    SQL = " INSERT INTO PRODUCION (ID_PERSONAL, ID_TIPOTAREA, ELEMENTO, FECHA,UNIDADPRODUCION,cod_cliente)"
    SQL = SQL & " VALUES (" & CInt(ID_Personal) & "," & ID_TIPOTAREA & ",'" & Elemento & "','" & Fecha1 & "','" & UNIDADPRODUCION & "'," & COD_CLIENTE & ")"
  Rem  ExecutarSql SQL
End Sub
