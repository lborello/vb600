Attribute VB_Name = "modGeneral"
Option Explicit
Public ConBasa As ADODB.Connection
Public BaseOracle As Boolean
Public Enum RemitoOperacion
   ENTRADA = 0
   Salida = 1
End Enum
Public Enum RemitoEstados
   Normal = 0
   Urgente = 1
End Enum
Enum RemitoTipo
   Guardia_Y_Custodia = 0
   Consultas = 1
   CAJAS_VACIAS = 2
   Bajas = 3
   Devolución_Cajas_Vacias = 4
End Enum
Enum Estado_Contenedor
   Disponible = 1
   Ocupado = 2
   Reservado = 3
   Cliente_Reservado = 4
End Enum
Enum tipo_almacenamiento
   Caja = 0
   Libro = 1
   Legajo = 2
   UIN = 3
End Enum
Enum Estado_Almacenamiento
    BAJA = 0
    En_Planta = 2
    CONSULTA = 3
    Reserva = 4
    Cliente_Nueva = 5
End Enum
Enum ABM
    Altas = 0
    Actualización = 1
    Bajas = 2
End Enum
Public PasoReportes As String
Public strPasoPlanillas  As String
Public PasoImagenes  As String
Public ClienteOsep  As String
Public ClienteEcogas  As String
Public strConBasa  As String
Public strConSoporte As String
Public Usuario As String
Public ID_Usuario As Integer
Public Sucursal As String
Public Responsables As clsResponsables
'Public oWordL  As Word.Application
'Public oTmpDocL As Word.Document
Public itemSelecionado As String
Public Nro_documento As String
Public Const ColorHabilitado = &HC0FFFF
Public Const ColorDesaHabilitado = &HC0C0FF

Public strConTangoCustodia As String
Public strConTangoBasa As String

Private Declare Function GetUserName& Lib "advapi32.dll" Alias "GetUserNameA" _
(ByVal lpBuffer As String, _
nSize As Long)
 
Public Function WindowsUserName() As String
     '   ---------------------------------------------
     '   Function to extract the name:
     '   ---------------------------------------------
    Dim szBuffer As String * 100
    Dim lBufferLen As Long
     
    lBufferLen = 100
     
    If CBool(GetUserName(szBuffer, lBufferLen)) Then
         
        WindowsUserName = Left$(szBuffer, lBufferLen - 1)
         
    Else
         
        WindowsUserName = CStr(Empty)
         
    End If
     
End Function
    Public Function ExisteArchivo(Paso As String) As Boolean

ExisteArchivo = False

If Dir(Paso) <> "" Then
ExisteArchivo = True
End If


End Function


Public Function Max_Cod_Id_Referencia() As Long
    Dim rsMax As New ADODB.Recordset
    rsMax.Open "SELECT MAX(COD_ID_REFERENCIA)AS MAX_COD_ID_REFERENCIA From REFERENCIAS ", ConActiva, 0, 1
    Max_Cod_Id_Referencia = rsMax!Max_Cod_Id_Referencia + 1
End Function

Public Sub TituloHerencia(Clientes As Integer)

        
        Dim RSLOOP As New ADODB.Recordset
        Dim rsCliente As New ADODB.Recordset
        Dim rs As ADODB.Recordset
        Dim sql As String
        Dim Cliente As Integer
        Dim conTitulo As New ADODB.Connection

            conTitulo.Open strConBasa
            conTitulo.CursorLocation = adUseClient

            If Clientes = 0 Then
                rsCliente.Open "Select * FROM clientes order by id_cliente ", strConBasa, 0, 1
            Else
                rsCliente.Open "Select * FROM clientes where id_cliente =" & Clientes & " order by id_cliente ", strConBasa, 0, 1
            End If
            
            Do While Not rsCliente.EOF
                
                Cliente = rsCliente!id_cliente
                
                ExecutarSql " Update INDICES Set TituloHerencia = Null Where COD_CLIENTE = " & Cliente
                sql = "SELECT COD_CLIENTE, INDICE, DESCRIPCION From INDICES WHERE COD_CLIENTE= " & Cliente & "  ORDER BY COD_CLIENTE, INDICE"
                Set RSLOOP = New ADODB.Recordset
                RSLOOP.Open sql, strConBasa, 0, 1
                Do While Not RSLOOP.EOF
                    Set rs = New ADODB.Recordset
                    sql = " SELECT TITULOHERENCIA"
                    sql = sql & vbCrLf & " From INDICES WHERE COD_CLIENTE =" & RSLOOP!COD_CLIENTE & " AND INDICE LIKE '" & RSLOOP!Indice & "%'"
                    rs.Open sql, ConActiva, adOpenKeyset, adLockOptimistic
                    Do While Not rs.EOF
                        If IsNull(rs!TituloHerencia) Then
                            rs!TituloHerencia = UCase(Trim(RSLOOP!Descripcion))
                        Else
                            rs!TituloHerencia = Mid(UCase(Trim(rs!TituloHerencia & " ** " & RSLOOP!Descripcion)), 1, 240)
                        End If
                        rs.Update
                        rs.MoveNext
                    Loop
                    RSLOOP.MoveNext
                Loop
                rsCliente.MoveNext
            Loop

End Sub
Public Function ConActiva() As ADODB.Connection

On Error GoTo salir:

'If IsObject(ConBasa) Then
'
'If ConBasa.State = 0 Then
    Set ConBasa = New ADODB.Connection
    ConBasa.CursorLocation = adUseClient
    ConBasa.CommandTimeout = 30000
    ConBasa.Open strConBasa
    Set ConActiva = ConBasa
'Else
' Set ConActiva = ConBasa
'End If
'Else
'Set ConBasa = New ADODB.Connection
'    ConBasa.Open strConBasa
'    Set ConActiva = ConBasa
'End If

    Exit Function
salir:
 Set ConBasa = New ADODB.Connection
    ConBasa.CursorLocation = adUseClient
    ConBasa.CommandTimeout = 30
    ConBasa.Open strConBasa
    Set ConActiva = ConBasa
End Function


Public Function SysDate2() As String
    
    
        
         Dim rsDate As New ADODB.Recordset
    Dim sql As String
    If BaseOracle = True Then
       sql = " SELECT "
        sql = sql & " TO_CHAR(SYSDATE, 'DD/MM/YYYY HH24:MI') as fecha "
        sql = sql & " FROM DUAL "
        rsDate.Open sql, ConActiva, 0, 1
        Rem SysDate2 = "'" & CStr(rsDate!Fecha) & "'"
        Else
        sql = " SELECT  GETDATE() AS FECHA"
        rsDate.Open sql, ConActiva, 0, 1
        SysDate2 = "'" & Format(rsDate!fecha, "dd/mm/yyyy hh:mm") & "'"
       End If
        
 End Function
 
 Public Function SysDateMinutoSegundo() As String
    
        
        
           Dim rsDate As New ADODB.Recordset
    Dim sql As String
    
        sql = " SELECT  GETDATE() AS FECHA"
        rsDate.Open sql, ConActiva, 0, 1
        SysDateMinutoSegundo = "CONVERT(DATETIME, '" & Format(rsDate!fecha, "YYYY-MM-DD hh:mm:ss") & "' , 102)"
    
        
 End Function

Public Function Legajos_RecalcularCaracteres_DescripcionRemito(SqlFiltro As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim CantidadCaracteres As Integer
    Dim DescripcionRemito As String
    Dim SumarCaracteres As Boolean
    
    SumarCaracteres = True
    
    
         '  3983 indice ecoga 4889 indice osep
        rs.CursorLocation = adUseClient
        
       sql = "  SELECT     ID_LEGAJO,DESCRIPCION_REMITO , FK_INDICES , CANTIDAD_CARACTERES, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION,FECHA_CREACION"
        sql = sql & " From LEGAJOS "
        sql = sql & SqlFiltro

        rs.Open sql, ConActiva, adOpenKeyset, adLockOptimistic


    Do While Not rs.EOF
        DescripcionRemito = ""
        CantidadCaracteres = Len(rs!ID_LEGAJO)
        
        If rs!FK_INDICES = 3983 Or rs!FK_INDICES = 4889 Then
            SumarCaracteres = False
        Else
           SumarCaracteres = True
        End If
        
        
        
        If Not IsNull(rs!NRO_DESDE) Then
           CantidadCaracteres = CantidadCaracteres + Len(rs!NRO_DESDE)
           DescripcionRemito = DescripcionRemito & rs!NRO_DESDE & " "
        End If
        
        If Not IsNull(rs!NRO_HASTA) Then
            If rs!NRO_DESDE <> rs!NRO_HASTA Then
                CantidadCaracteres = CantidadCaracteres + Len(rs!NRO_HASTA)
                DescripcionRemito = DescripcionRemito & rs!NRO_HASTA & " "
            End If
        End If
        
        
        If Not IsNull(rs!LETRA_DESDE) Then
        
            If SumarCaracteres Then
                CantidadCaracteres = CantidadCaracteres + Len(Trim(rs!LETRA_DESDE))
            End If
            DescripcionRemito = DescripcionRemito & Trim(rs!LETRA_DESDE) & " "
        End If
        
        If Not IsNull(rs!LETRA_HASTA) Then
            If rs!LETRA_HASTA <> rs!LETRA_DESDE Then
                If SumarCaracteres Then
                    CantidadCaracteres = CantidadCaracteres + Len(Trim(rs!LETRA_HASTA))
                End If
                DescripcionRemito = DescripcionRemito & Trim(rs!LETRA_HASTA) & " "
            End If
        End If
        
        
        If Not IsNull(rs!FECHA_DESDE) Then
            If Mid(rs!FECHA_DESDE, 1, 5) = "01/01" Then
                    CantidadCaracteres = CantidadCaracteres + 4
                    DescripcionRemito = DescripcionRemito & Mid(rs!FECHA_DESDE, 7) & " "
            Else
                    CantidadCaracteres = CantidadCaracteres + Len(rs!FECHA_DESDE)
            End If
        End If
        
        
        If Not IsNull(rs!FECHA_HASTA) Then
            If Mid(rs!FECHA_HASTA, 1, 5) = "31/12" Then
                
                
            Else
                CantidadCaracteres = CantidadCaracteres + Len(rs!FECHA_HASTA)
                DescripcionRemito = DescripcionRemito & Mid(rs!FECHA_HASTA, 7)
            End If
       End If
        
        
        
        
        If Not IsNull(rs!Descripcion) Then
             CantidadCaracteres = CantidadCaracteres + Len(rs!Descripcion)
             DescripcionRemito = DescripcionRemito & Trim(rs!Descripcion)
        End If
        rs!CANTIDAD_CARACTERES = CantidadCaracteres
        rs!DESCRIPCION_REMITO = DescripcionRemito
        CantidadCaracteres = 0
        DescripcionRemito = ""
        rs.Update
    
        rs.MoveNext
    Loop
  rs.Close


End Function

 Public Function BuscarIDDocumento(ID_CODIGO_DOCUMENTO As Long, COD_CLIENTE As Integer) As String
    Dim RsDocumento As ADODB.Recordset
    Set RsDocumento = New ADODB.Recordset
    Dim sSQL As String
        
        sSQL = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE,"
        sSQL = sSQL & vbCrLf & " DESCRIPCION, FECHA, NUMERO, LETRA, EXPEDIENTE,APELLIDO_NOMBRE"
        sSQL = sSQL & vbCrLf & " From INDICES"
        sSQL = sSQL & vbCrLf & " WHERE (COD_CLIENTE =" & COD_CLIENTE & ") AND  (ID_CODIGO_DOCUMENTO =" & ID_CODIGO_DOCUMENTO & ")"
        RsDocumento.Open sSQL, ConActiva, 0, 1
        If Not RsDocumento.EOF Then
            BuscarIDDocumento = Trim(RsDocumento!Indice)
        Else
            BuscarIDDocumento = "ERROR"
        End If
        
        
End Function
 
Public Function Buscar_ID_CODIGO_DOCUMENTO(Indice As String, COD_CLIENTE As Integer) As Long
    Dim RsDocumento As ADODB.Recordset
    Set RsDocumento = New ADODB.Recordset
    Dim sSQL As String
        
        sSQL = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE,"
        sSQL = sSQL & vbCrLf & " DESCRIPCION, FECHA, NUMERO, LETRA, EXPEDIENTE,APELLIDO_NOMBRE"
        sSQL = sSQL & vbCrLf & " From INDICES"
        sSQL = sSQL & vbCrLf & " WHERE (COD_CLIENTE =" & COD_CLIENTE & ") AND  (INDICE ='" & Indice & "')"
        RsDocumento.Open sSQL, ConActiva, 0, 1
        If Not RsDocumento.EOF Then
            Buscar_ID_CODIGO_DOCUMENTO = Trim(RsDocumento!ID_CODIGO_DOCUMENTO)
        Else
            Buscar_ID_CODIGO_DOCUMENTO = 0
        End If
        
        
End Function
 Public Function Buscar_Indice_ID(Cliente As Integer, Nro_documento As Integer) As Long
    Dim rs As New ADODB.Recordset
    Dim sql As String
    
    sql = " SELECT     ID"
sql = sql & " From INDICES"
sql = sql & "  Where COD_CLIENTE = " & Cliente
sql = sql & "  And ID_CODIGO_DOCUMENTO = " & Nro_documento


rs.Open sql, ConActiva, 0, 1

Buscar_Indice_ID = rs!ID


End Function

Public Function BuscarIndice(Cliente As Integer, NumeroDOc As Long) As String
Dim rs3 As New ADODB.Recordset
Dim CONB As New ADODB.Connection
CONB.Open strConBasa
Dim sql As String
sql = " SELECT        ID, INDICE"
sql = sql & " From INDICES"
sql = sql & "  Where COD_CLIENTE = " & Cliente
sql = sql & "  And ID_CODIGO_DOCUMENTO = " & NumeroDOc
rs3.Open sql, CONB

If Not rs3.EOF Then
    BuscarIndice = rs3!Indice
    Rem ID_indice = rs3!ID
 Else
    BuscarIndice = 0
   Rem  ID_indice = 0
End If
End Function
 
 Public Function SysDate() As String
    
    Dim rsDate As New ADODB.Recordset
    Dim sql As String
    
        sql = " SELECT  GETDATE() AS FECHA "
        rsDate.Open sql, ConActiva, 0, 1
        SysDate = FechaFormato(rsDate!fecha)
    
        
 End Function
 
 Public Function BuscarDirectorioPaso(DATO As Long) As String

Dim rs As New ADODB.Recordset
Dim sql As String
rs.CursorLocation = adUseClient

sql = " SELECT     ID,DIRECTORIO_PASO"
sql = sql & " From DIRECTORIOS_IMAGENES"
sql = sql & " WHERE " & DATO & " BETWEEN DESDE AND HASTA"
rs.Open sql, ConActiva, 0, 1
If Not rs.EOF Then
BuscarDirectorioPaso = rs!DIRECTORIO_PASO
Else
BuscarDirectorioPaso = ""
End If


End Function

 
Public Function Buscar_ID_Indice(ID_CODIGO_DOCUMENTO As Long, COD_CLIENTE As Integer) As Long
    Dim RsDocumento As ADODB.Recordset
    Set RsDocumento = New ADODB.Recordset
    Dim sSQL As String
        sSQL = " SELECT ID "
        sSQL = sSQL & vbCrLf & " From INDICES"
        sSQL = sSQL & vbCrLf & " WHERE (COD_CLIENTE =" & COD_CLIENTE & ") AND  (ID_CODIGO_DOCUMENTO =" & ID_CODIGO_DOCUMENTO & ")"
        RsDocumento.Open sSQL, strConBasa, 0, 1
        If Not RsDocumento.EOF Then
            Buscar_ID_Indice = RsDocumento!ID
        Else
            Buscar_ID_Indice = "0"
        End If
        
        
End Function

Public Function Buscar_ID_Indice_Por_indice(Indice As String, COD_CLIENTE As Integer) As Long
    Dim RsDocumento As ADODB.Recordset
    Set RsDocumento = New ADODB.Recordset
    Dim sSQL As String
    
   sSQL = " SELECT     ID, INDICE, COD_CLIENTE"
sSQL = sSQL & vbCrLf & " From INDICES"
sSQL = sSQL & vbCrLf & " WHERE INDICE = '" & Trim(Indice) & "'"
sSQL = sSQL & vbCrLf & "  AND COD_CLIENTE = " & COD_CLIENTE
    
        RsDocumento.Open sSQL, ConActiva, 0, 1
        If Not RsDocumento.EOF Then
            Buscar_ID_Indice_Por_indice = RsDocumento!ID
        Else
            Buscar_ID_Indice_Por_indice = "0"
        End If
        
        
End Function


'Public Function BuscarIndice(ID_CODIGO_DOCUMENTO As Long, COD_CLIENTE As Integer) As String
'    Dim RsDocumento As ADODB.Recordset
'    Set RsDocumento = New ADODB.Recordset
'    Dim sSQL As String
'        sSQL = " SELECT INDICE  "
'        sSQL = sSQL & vbCrLf & " From INDICES"
'        sSQL = sSQL & vbCrLf & " WHERE (COD_CLIENTE =" & COD_CLIENTE & ") AND  (ID_CODIGO_DOCUMENTO =" & ID_CODIGO_DOCUMENTO & ")"
'        RsDocumento.Open sSQL, ConActiva, 0, 1
'        If Not RsDocumento.EOF Then
'            BuscarIndice = RsDocumento!indice
'        Else
'            BuscarIndice = "001"
'            MsgBox "Se asignara un indice generico cliente : " & COD_CLIENTE & " documento : " & ID_CODIGO_DOCUMENTO
'        End If
'
'
'End Function

Public Function BuscarIndiceDescripcion(Indice As String, COD_CLIENTE As Integer) As String
    Dim RsDocumento As ADODB.Recordset
    Set RsDocumento = New ADODB.Recordset
    Dim sSQL As String
        
        sSQL = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE,"
        sSQL = sSQL & vbCrLf & " DESCRIPCION, FECHA, NUMERO, LETRA, EXPEDIENTE,APELLIDO_NOMBRE"
        sSQL = sSQL & vbCrLf & " From INDICES"
        sSQL = sSQL & vbCrLf & " WHERE (COD_CLIENTE =" & COD_CLIENTE & ") AND  (INDICE ='" & Indice & "')"
        RsDocumento.Open sSQL, ConActiva, 0, 1
        If Not RsDocumento.EOF Then
            BuscarIndiceDescripcion = Trim(RsDocumento!Descripcion)
        Else
            BuscarIndiceDescripcion = "ERROR"
        End If
        
        
End Function
Public Function BuscarID_IndiceDocumento_Indice(Documento As String, COD_CLIENTE As Integer) As Long
    Dim RsDocumento As ADODB.Recordset
    Set RsDocumento = New ADODB.Recordset
    Dim sSQL As String
        
        sSQL = " SELECT FK_INDICES "
        sSQL = sSQL & vbCrLf & " From INDICES"
        sSQL = sSQL & vbCrLf & " WHERE (COD_CLIENTE =" & COD_CLIENTE & ") AND  ID_CODIGO_DOCUMENTO =" & Documento
        RsDocumento.Open sSQL, ConActiva, 0, 1
        If Not RsDocumento.EOF Then
            BuscarID_IndiceDocumento_Indice = Trim(RsDocumento!FK_INDICES)
        Else
            BuscarID_IndiceDocumento_Indice = 0
        End If
        
        
End Function


Public Function BuscarIndiceDocumento_Indice(Documento As String, COD_CLIENTE As Integer) As String
    Dim RsDocumento As ADODB.Recordset
    Set RsDocumento = New ADODB.Recordset
    Dim sSQL As String
        
        sSQL = " SELECT COD_CLIENTE, ID_CODIGO_DOCUMENTO, INDICE,"
        sSQL = sSQL & vbCrLf & " DESCRIPCION, FECHA, NUMERO, LETRA, EXPEDIENTE,APELLIDO_NOMBRE"
        sSQL = sSQL & vbCrLf & " From INDICES"
        sSQL = sSQL & vbCrLf & " WHERE (COD_CLIENTE =" & COD_CLIENTE & ") AND  ID_CODIGO_DOCUMENTO =" & Documento
        RsDocumento.Open sSQL, ConActiva, 0, 1
        If Not RsDocumento.EOF Then
            BuscarIndiceDocumento_Indice = Trim(RsDocumento!Indice)
        Else
            BuscarIndiceDocumento_Indice = ""
        End If
        
        
End Function
Public Function BuscarMaskExpediente(Indice As String, COD_CLIENTE As Integer) As String
    Dim RsDocumento As ADODB.Recordset
    Set RsDocumento = New ADODB.Recordset
    Dim sSQL As String
        
        BuscarMaskExpediente = ""
        sSQL = " SELECT MASK_EXPEDIENTE "
        sSQL = sSQL & vbCrLf & "  From INDICES"
        sSQL = sSQL & vbCrLf & "  WHERE COD_CLIENTE = " & COD_CLIENTE
        sSQL = sSQL & vbCrLf & "  AND INDICE = '" & Indice & "'"
        RsDocumento.Open sSQL, ConActiva, 0, 1
        If Not RsDocumento.EOF Then
            If Not IsNull(RsDocumento!MASK_EXPEDIENTE) Then
                BuscarMaskExpediente = Trim(RsDocumento!MASK_EXPEDIENTE)
            End If
        Else
            BuscarMaskExpediente = "ERROR"
        End If
        
        
End Function

'Function Paradoja_Estado(Elementos As Variant, COD_CLIENTE As Integer, Para_Operacion As RemitoOperacion, _
'                        Para_Tipo As RemitoTipo, Para_Almacemamiento As tipo_almacenamiento, _
'                        Mensajes As Boolean, GrillaConsulta As MSFlexGrid, _
'                        Optional GrillaCustodia As MSFlexGrid, Optional lblCantidadConsulta As Label, _
'                        Optional lblCantidadCustodia As Label, Optional VerMovimientos As Boolean, _
'                        Optional ControlHablar As MMControl) As Boolean
'  Dim ColElementos As New Collection
'  Dim Elemento As Long
'  Dim CaptionElemento As String
'  Dim RsEstado As ADODB.Recordset
'  Dim Sql As String
'  Dim i As Integer
'
'        If IsObject(Elementos) Then
'            Set ColElementos = Elementos
'        Else
'            ColElementos.Add Elementos
'        End If
'        Paradoja_Estado = True
'            For i = 1 To ColElementos.Count
'                  Elemento = ColElementos.Item(i)
'                  Select Case Para_Almacemamiento
'                  Case tipo_almacenamiento.Caja
'                      Sql = "Select COD_Estado as EstadoElemento From Cajas Where Cliente_Caja = " & Elemento & " AND  COD_Cliente = " & COD_CLIENTE
'                      CaptionElemento = "Caja"
'                  Case tipo_almacenamiento.Legajo
'                      Sql = "Select cod_Estado as EstadoElemento From Legajos Where ID_Legajo = " & Elemento & " AND  COD_Cliente = " & COD_CLIENTE
'                      CaptionElemento = "Legajo"
'                  Case tipo_almacenamiento.Libro
'                      Sql = "Select Estado as EstadoElemento  From Libros Where NRO_LIBRO_INTERNO  = " & Elemento & " AND  COD_Cliente = " & COD_CLIENTE
'                      CaptionElemento = "Libro"
'                  End Select
'                  Set RsEstado = New ADODB.Recordset
'                  RsEstado.Open Sql, ConActiva, 0, 1
'                    If Not RsEstado.EOF Then
'                        Select Case Para_Tipo
'                        Case RemitoTipo.Bajas
'                            If Not (RsEstado.Fields("EstadoElemento").value = Estado_Almacenamiento.Reserva) Then
'                                If Mensajes Then
'                                    MsgBox " El Estado  de " & CaptionElemento & vbCrLf & " Numero " & Elemento & " ES INCORRECTO ", vbCritical
'                                End If
'                               Rem hablar "ESTADO INCORRECTO", ControlHablar
'                                Paradoja_Estado = False
'                            End If
'                        Case RemitoTipo.CAJAS_VACIAS
'                            If Not (RsEstado.Fields("EstadoElemento").value = Estado_Almacenamiento.Reserva) Then
'                                If Mensajes Then
'                                    MsgBox " El Estado  de " & CaptionElemento & vbCrLf & " Numero " & Elemento & " ES INCORRECTO ", vbCritical
'                                End If
'                              Rem hablar "ESTADO INCORRECTO", ControlHablar
'                                Paradoja_Estado = False
'                            End If
'                        Case RemitoTipo.Consultas
'                            If Para_Operacion = ENTRADA Then
'                                If Not (RsEstado.Fields("EstadoElemento").value = Estado_Almacenamiento.Consulta) Then
'                                    If Mensajes Then
'                                        MsgBox " El Estado  de " & CaptionElemento & vbCrLf & " Numero " & Elemento & " ES INCORRECTO ", vbCritical
'                                    End If
'                                  Rem hablar "ESTADO INCORRECTO", ControlHablar
'                                    Paradoja_Estado = False
'                                Else
'                                    CargarGrilla CStr(Elemento), GrillaConsulta, lblCantidadConsulta, False, ControlHablar
'                                End If
'                            Else
'                                If Not (RsEstado.Fields("EstadoElemento").value = Estado_Almacenamiento.En_Planta) Then
'                                        If Mensajes Then
'                                            MsgBox " El Estado  de " & CaptionElemento & vbCrLf & " Numero " & Elemento & " ES INCORRECTO ", vbCritical
'                                        End If
'                                      Rem hablar "ESTADO INCORRECTO", ControlHablar
'                                        Paradoja_Estado = False
'                                Else
'                                    CargarGrilla CStr(Elemento), GrillaConsulta, lblCantidadConsulta, False, ControlHablar
'                                End If
'                            End If
'                        Case RemitoTipo.Devolución_Cajas_Vacias
'                           If Not (RsEstado.Fields("EstadoElemento").value = Estado_Almacenamiento.Cliente_Nueva) Then
'                                If Mensajes Then
'                                    MsgBox " El Estado  de " & CaptionElemento & vbCrLf & " Numero " & Elemento & " ES INCORRECTO ", vbCritical
'                                End If
'                              Rem hablar "Estado Incorrecto", ControlHablar
'                                Paradoja_Estado = False
'                           Else
'                                CargarGrilla CStr(Elemento), GrillaConsulta, lblCantidadConsulta, False, ControlHablar
'                           End If
'                        Case RemitoTipo.Guardia_Y_Custodia
'                            If Not (RsEstado.Fields("EstadoElemento").value = Estado_Almacenamiento.Cliente_Nueva) Then
'                                If Mensajes Then
'                                   MsgBox " El Estado  de " & CaptionElemento & vbCrLf & " Numero " & Elemento & " ES INCORRECTO ", vbCritical
'                                End If
'                              Rem hablar "Estado Incorrecto", ControlHablar
'                                Paradoja_Estado = False
'                           End If
'                        End Select
'                     Else
'                        MsgBox " La/El  " & CaptionElemento & vbCrLf & " Numero " & Elemento & " NO EXISTE", vbCritical
'                        Paradoja_Estado = False
'                    End If
'            Next
'End Function
Function Paradoja_Estado_Simple(Elementos As Variant, COD_CLIENTE As Integer, Para_Operacion As RemitoOperacion, Para_Almacemamiento As tipo_almacenamiento) As Boolean
  Dim ColElementos As New Collection
  Dim Elemento As Long
  Dim CaptionElemento As String
  Dim RsEstado As ADODB.Recordset
  Dim sql As String
  Dim i As Integer
  
        If IsObject(Elementos) Then
            Set ColElementos = Elementos
        Else
            ColElementos.Add Elementos
        End If
        Paradoja_Estado_Simple = True
            For i = 1 To ColElementos.Count
                  Elemento = ColElementos.Item(i)
                  Select Case Para_Almacemamiento
                  Case tipo_almacenamiento.Caja
                      sql = " SELECT ESTADO as EstadoElemento From CONTENEDOR Where (COD_CLIENTE = " & COD_CLIENTE & ") And (NRO_CAJA = " & Elemento & ")"
                      CaptionElemento = "Caja"
                  Case tipo_almacenamiento.Legajo
                      sql = "Select cod_Estado as EstadoElemento From Legajos Where ID_CLIENTE_LEGAJO = " & Elemento & " AND  COD_Cliente = " & COD_CLIENTE
                      CaptionElemento = "Legajo"
                  Case tipo_almacenamiento.Libro
                      sql = "Select Estado as EstadoElemento  From Libros Where NRO_LIBRO_INTERNO  = " & Elemento & " AND  COD_Cliente = " & COD_CLIENTE
                      CaptionElemento = "Libro"
                  End Select
                  Set RsEstado = New ADODB.Recordset
                  RsEstado.Open sql, ConActiva, 0, 1
                    If Not RsEstado.EOF Then
                        If Para_Operacion = Salida Then
                            If RsEstado.Fields("EstadoElemento").value <> 2 Then
                               MsgBox " La/El  " & CaptionElemento & vbCrLf & " Numero " & Elemento & " No Tiene estado Correcto"
                               Paradoja_Estado_Simple = False
                            End If
                        Else
                            If RsEstado.Fields("EstadoElemento").value <> 3 Or RsEstado.Fields("EstadoElemento").value <> 5 Then
                               MsgBox " La/El  " & CaptionElemento & vbCrLf & " Numero " & Elemento & " No Tiene estado Correcto"
                               Paradoja_Estado_Simple = False
                            End If
                        End If
                    
                     Else
                        MsgBox " La/El  " & CaptionElemento & vbCrLf & " Numero " & Elemento & " NO EXISTE", vbCritical
                        Paradoja_Estado_Simple = False
                    End If
            Next
End Function




'Public Sub CargarGrilla(Valor As String, Grilla As MSFlexGrid, lblCantidad As Label, Mensaje As Boolean, Optional ControlHablar As MMControl)
'    Dim C As Integer
'    Dim R As Integer
'
'    For R = 1 To Grilla.Rows - 1
'        For C = 1 To Grilla.Cols - 1
'            If Grilla.TextMatrix(R, C) = Valor Then
'                If Mensaje Then
'                    MsgBox "Repetida", vbCritical
'                End If
'
'              Rem hablar "REPETIDA", ControlHablar
'                Exit Sub
'            End If
'        Next
'    Next
'    R = 1
'    C = 1
'    For R = 1 To Grilla.Rows - 1
'        For C = 1 To Grilla.Cols - 1
'            If Grilla.TextMatrix(R, C) = "" Then
'                Grilla.TextMatrix(R, C) = Valor
'              Rem hablar "ENTRADA", ControlHablar
'                ContarGrilla Grilla, lblCantidad
'                Exit Sub
'            End If
'        Next
'    Next
'    Grilla.AddItem "" & vbTab & Valor
'  Rem hablar "ENTRADA", ControlHablar
'    ContarGrilla Grilla, lblCantidad
'End Sub
Public Sub LimpiarMask(Control As MaskEdBox)
Dim Mask1 As String
Mask1 = Control.Mask
Control.Mask = ""
Control.Text = ""
Control.Mask = Mask1
End Sub
Public Function ContarGrilla(Grilla As MSFlexGrid, cantidad As Label)
    Dim i As Integer
    Dim R As Integer
    Dim C As Integer
   i = 0
    With Grilla
         For R = 1 To .Rows - 1
             For C = 1 To .Cols - 1
                 If .TextMatrix(R, C) <> "" Then
                     i = i + 1
                 End If
             Next
         Next
     End With
     cantidad.Caption = i
End Function
Public Function PosicionCaracter(DATO As String, Buscar As String) As Integer

Dim i As Integer


For i = 1 To Len(DATO)
 If Mid(DATO, i, 1) = Buscar Then
     PosicionCaracter = i
     Exit Function
 End If
Next
PosicionCaracter = 0
End Function


Public Sub DATOSGRILLA(Grilla As DataGrid, rs As ADODB.Recordset)
        Grilla.ClearFields
        Grilla.ClearSelCols
        Grilla.ScrollBars = dbgAutomatic
        Dim i As Integer
        For i = 0 To rs.Fields.Count - 1
            Debug.Print rs.Fields.Item(i).Name & "  " & rs.Fields.Item(i).Type
            Grilla.Columns.Add i
            Grilla.Columns.Item(i).DataField = rs.Fields(i).Name
            Grilla.Columns.Item(i).Caption = rs.Fields(i).Name
            Select Case rs.Fields.Item(i).Type
            Case "131" ' NUMERO
                Grilla.Columns.Item(i).Width = 500
            Case "200" 'TEXT
                Grilla.Columns.Item(i).Width = 1500
            Case "135" 'FECHA
                Grilla.Columns.Item(i).Width = 700
            End Select
        Next
        Set Grilla.DataSource = rs.DataSource
        Grilla.Refresh
End Sub
Public Sub CopiarDatosGrilla(Grilla As DataGrid)
Dim C As Integer
Dim R As Integer
Dim RSDATOS As ADODB.Recordset
Dim DATO As String
Dim ColGrilla As Integer
Set RSDATOS = New ADODB.Recordset

Set RSDATOS.DataSource = Grilla.DataSource
 On Error GoTo salir
 

For C = 0 To RSDATOS.Fields.Count - 1
    DATO = DATO & RSDATOS.Fields(C).Name & vbTab
 Next
    DATO = DATO & vbCrLf
    Do While Not RSDATOS.EOF
        For C = 0 To RSDATOS.Fields.Count - 1
            If Not IsNull(RSDATOS.Fields.Item(C).value) Then
                DATO = DATO & RSDATOS.Fields.Item(C).value & vbTab
            Else
                DATO = DATO & "" & vbTab
            End If
        Next
        RSDATOS.MoveNext
        DATO = DATO & vbCrLf
    Loop
 Clipboard.Clear
 Clipboard.SetText DATO
 MsgBox "LOS DATOS FUERON COPIADOS"
salir:
If Err.Number <> 0 Then
    MsgBox Err.Description
    Exit Sub
End If
 
End Sub
Public Sub InsertarImagenes(ID_SQL As Long, COD_CLIENTE As Integer, Elemento As Long, _
    TIPO_DOCUMENTO As Integer, fecha As String)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
        sql = "  SELECT ID_SQL, COD_CLIENTE, ELEMENTO "
        sql = sql & vbCrLf & " From IMAGENES "
        sql = sql & vbCrLf & " WHERE ID_SQL = " & ID_SQL
        sql = sql & vbCrLf & " AND COD_CLIENTE = " & COD_CLIENTE
        sql = sql & vbCrLf & " AND ELEMENTO = " & Elemento
     
        rs.Open sql, ConActiva, 0, 1
        
     
    Rem  MsgBox "errro1"
     If rs.EOF Then
        sql = " INSERT INTO IMAGENES ("
        sql = sql & vbCrLf & " ID, ID_SQL, COD_CLIENTE "
        sql = sql & vbCrLf & " , ELEMENTO, TIPO_DOCUMENTO,FECHA)"
        sql = sql & vbCrLf & " VALUES ("
        sql = sql & vbCrLf & maxImagen & "," & ID_SQL & "," & COD_CLIENTE
        sql = sql & vbCrLf & "," & Elemento & "," & TIPO_DOCUMENTO & "," & Format(fecha, "DD/MM/YYYY") & ")"
       Rem  Clipboard.SetText Sql
       Rem  MsgBox Sql
        ExecutarSql sql
    End If
    
End Sub

Public Function maxImagen() As Long
    Dim sql As String
    Dim rsMax As New ADODB.Recordset
        sql = " SELECT MAX(ID) AS maxImagen From IMAGENES "
        rsMax.Open sql, ConActiva, 0, 1
        maxImagen = rsMax!maxImagen + 1
        
End Function

Public Function ControldatoString(DATO As String) As String
    
    If Trim(DATO) = "" Then
        ControldatoString = "NULL"
    Else
        ControldatoString = "'" & UCase(Trim(DATO)) & "'"
   End If
    
End Function


Public Function FechaServerTipo(DATO As String) As String
    If BaseOracle Then
        FechaServerTipo = " TO_DATE('" & DATO & "', 'DD/MM/YYYY')"
    Else
        FechaServerTipo = " CONVERT(DATETIME, '" & Format(DATO, "dd/mm/yyyy") & "' , 103)"
         
    End If

 
End Function
Public Function FechaSegundoServerTipo(DATO As String) As String
    If BaseOracle = True Then

    FechaSegundoServerTipo = " TO_DATE('" & DATO & "', 'DD/MM/YYYY HH24:MI:SS')"
    
    Else
    
     FechaSegundoServerTipo = " CONVERT(DATETIME, '" & DATO & "', 102)"
    
    End If
    
 
End Function


Public Function BuscarID_legajo(Etiqueta As Long, COD_CLIENTE As Integer) As Long
Dim rs As New ADODB.Recordset
Dim sql As String

sql = " SELECT     ID_LEGAJO "
sql = sql & " From LEGAJOS "
sql = sql & "  Where ID_CLIENTE_LEGAJO = " & Etiqueta
sql = sql & "  And cod_cliente = " & COD_CLIENTE

rs.Open sql, ConActiva, 0, 1

If rs.EOF Then
    BuscarID_legajo = 0
Else
    BuscarID_legajo = rs!ID_LEGAJO
End If



End Function
Public Function Digito_Verificador(DATO As String) As Integer
 Digito_Verificador = BuscarDigitoVerificador(CLng(DATO))

'Dim i As Integer
'Dim Sumar As Integer
'
'    For i = 1 To Len(Trim(Dato))
'        Sumar = Sumar + Mid(Dato, i, 1)
'    Next
'      Digito_Verificador = Sumar
'
End Function

Function DigitoEAN13(RawString As String) As Integer
Dim Position As Integer
Dim CheckSum As Integer

CheckSum = 0
For Position = 2 To 12 Step 2
      CheckSum = CheckSum + Val(Mid$(RawString, Position, 1))
Next Position
CheckSum = CheckSum * 3
For Position = 1 To 11 Step 2
     CheckSum = CheckSum + Val(Mid$(RawString, Position, 1))
Next Position
CheckSum = CheckSum Mod 10
CheckSum = 10 - CheckSum
If CheckSum = 10 Then
     CheckSum = 0
End If
DigitoEAN13 = Format$(CheckSum, "0")
End Function

Public Function BuscarDigitoVerificador(ID_LEGAJO As Long) As Integer

Dim sql  As String
Dim rs As New ADODB.Recordset

sql = " SELECT     ID_LEGAJO, DIGITO_VERIFICADOR"
sql = sql & "  From basasql.dbo.LEGAJOS"
sql = sql & "  Where id_legajo = " & ID_LEGAJO

rs.Open sql, strConBasa

If IsNull(rs!Digito_Verificador) Then
Exit Function

End If


If Not rs.EOF Then
             BuscarDigitoVerificador = rs!Digito_Verificador
    
Else
    BuscarDigitoVerificador = 0
End If

End Function

Public Function BuscarDigitoVerificadorCajas(CAJAS As String) As Integer

Dim sql  As String
Dim rs As New ADODB.Recordset
 Dim i As Integer
Dim Sumar As Integer



sql = " SELECT  ID_CAJA, DIGITO_VERIFICADOR"
sql = sql & "   From basasql.dbo.Cajas"
sql = sql & "   Where ID_CAJA =" & CAJAS


rs.Open sql, strConBasa

If Not rs.EOF Then
 If Not IsNull(rs!Digito_Verificador) Then
    BuscarDigitoVerificadorCajas = rs!Digito_Verificador
Else
    For i = 1 To Len(Trim(CAJAS))
        Sumar = Sumar + Mid(CAJAS, i, 1)
    Next
    BuscarDigitoVerificadorCajas = Sumar
End If

Else
BuscarDigitoVerificadorCajas = 0
End If

End Function
Public Function inicio()
    
    
On Error GoTo salir
  Dim cad As String, i As Byte, s As Byte, var As Byte
  PasoReportes = "Z:\Sistemas\Basa\Reportes_Sistema\"
  strPasoPlanillas = "Z:\Sistemas\Basa\Planillas\"
'  ClienteOsep = "Z:\Sistemas\Basa\ClientesBases\OSEP.mdb    "
'  ClienteEcogas = "Z:\Sistemas\Basa\ClientesBases\Pig.mdb"
  
'   MsgBox WindowsUserName
'
''        If FileLen("\\222.15.19.251\basa\Sistemas\Basa\SistemaBasa.exe") <> FileLen(App.Path & "\SistemaBasa.exe") Then
''             MsgBox "Su sistema desactualizado", vbCritical
''         End
''        End If
    
  
  
  Rem   Shell "Regsvr32.exe /u /s " & "Z:\Sistemas\Basa\Controles\Controles4.ocx"
  Rem  Open "\\222.15.19.251\basa\Sistemas\Basa\Configuracion.txt" For Input As #1
     Open "z:\Sistemas\Basa\Configuracion.txt" For Input As #1
     
     While Not EOF(1) 'Recorre archivo hasta que termine
        Input #1, cad
        s = 1 'Controla inicio de cada cadena
        var = 1 'Control el Campo a asignar
        Select Case Trim(Mid(cad, 1, 24))
        Case "PasoImagenes"
            PasoImagenes = Trim(Mid(cad, 25))
        Case "strConBasa"
              strConBasa = Replace(Trim(Mid(cad, 25)), ":", ",")
       Case "Sucursal"
            Sucursal = Trim(Trim(Mid(cad, 25)))
        Case "strConTangoCustodia"
            strConTangoCustodia = Replace(Trim(Mid(cad, 25)), ":", ",")
       Case "strConTangoBasa"
            strConTangoBasa = Replace(Trim(Mid(cad, 25)), ":", ",")
        End Select
       
       Wend
       Rem ojoluis
       Rem strConBasa = "Provider=SQLOLEDB.1;Password=Basa2012;Persist Security Info=False;User ID=sa;Initial Catalog=basasql;Data Source=190.151.143.135,5555"
    BaseOracle = False
   Rem  MsgBox "1"
    Dim rs As ADODB.Recordset
    Dim sql As String

Rem MsgBox "2"
    
    Set rs = New ADODB.Recordset
    sql = " SELECT     IDPERSONAL, NOMBRE, APELLIDO, USUARIOSYS "
    sql = sql & " From dbo.Personal "
    sql = sql & " Where USUARIOSYS = '" & WindowsUserName & "'"


    rs.Open sql, ConActiva, adOpenStatic, adLockReadOnly
   Rem  MsgBox "3"

    If Not rs.EOF Then
        Usuario = rs!idPersonal
        MDIfrmInicio.StaInicio.Panels(2).Text = Usuario
        MDIfrmInicio.StaInicio.Panels(3).Text = Trim(rs!Nombre) & " " & Trim(rs!Apellido)
        Exit Function
    End If

    Usuario = InputBox("INGRESE EL Nº DE USUARIO")
    Set rs = New ADODB.Recordset
    sql = " SELECT     IDPERSONAL, NOMBRE, APELLIDO"
    sql = sql & " From dbo.Personal "
    sql = sql & " Where IDPERSONAL = " & Usuario
    rs.Open sql, ConActiva, adOpenStatic, adLockReadOnly
    If rs.EOF Then
        MsgBox "Usuario Incorrecto"
      End
    Else
    MDIfrmInicio.StaInicio.Panels.Item(2).Text = Usuario
       MDIfrmInicio.StaInicio.Panels(3).Text = Trim(rs!Nombre) & " " & Trim(rs!Apellido)
    End If

     Close #1

    Exit Function
  Exit Function
salir:
  MsgBox Err.Description
 Rem End
  
    
    
    
    
End Function

Public Sub CopiarDatosGrillaMSg(Grilla As MSFlexGrid)
    Dim s As String
    Dim C As Integer
    Dim i As Integer
        For i = 0 To Grilla.Rows - 1
           For C = 0 To Grilla.Cols - 1
                s = s & vbTab & Replace(Grilla.TextMatrix(i, C), vbCrLf, " ")
            Next
            C = 0
            s = s & vbCrLf
        Next
        Clipboard.Clear
        Clipboard.SetText s
        MsgBox "Datos copiados"

End Sub
Public Function ExecutarSql(sql As String) As Integer
On Error GoTo salir
    Dim con As New ADODB.Connection
    Dim Registros As Integer
        Set con = ConActiva
       con.Execute sql, Registros
        ExecutarSql = Registros
        con.Close
        Exit Function
salir:
MsgBox Err.Description
MsgBox sql
        ExecutarSql = 0

End Function

Public Function FechaFormato(fecha As Variant)
    If UCase(fecha) <> "NULL" Then
        FechaFormato = " CONVERT(DATETIME, '" & Format(fecha, "YYYY-MM-DD") & " 00:00:00', 102)"
    Else
        FechaFormato = "NULL"
    End If
End Function


Public Function FECHADATA_Fecha(NumeroDias) As Date
  FECHADATA_Fecha = DateAdd("d", NumeroDias, "28/12/1800")
End Function
Public Function FECHADATA_Dias(fecha As Date) As Long
 FECHADATA_Dias = DateDiff("d", "28/12/1800", fecha)
End Function



Public Function CrearCajas(Caja As Long, Cliente As Integer)
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim ID_CAJA As Long
    Dim Etiqueta As String
        Rem Cajas
        sql = " SELECT  ID_CAJA, FK_CLIENTE, NRO_CAJA, FK_CONTENEDOR, FK_ESTADO"
        sql = sql & " From basasql.dbo.Cajas"
        sql = sql & " Where FK_CLIENTE = " & Cliente
        sql = sql & " And NRO_CAJA = " & Caja
        rs.Open sql, strConBasa
        If rs.EOF Then
            sql = " SELECT  TOP (1) ID_CAJA"
            sql = sql & " From basasql.dbo.Cajas"
            sql = sql & "  WHERE     ( FK_CLIENTE  IS NULL)"
            sql = sql & "  AND ID_CAJA BETWEEN 384957  and    400000"
            sql = sql & "  ORDER BY ID_CAJA "
            Set rs2 = New ADODB.Recordset
            rs2.Open sql, strConBasa
            Etiqueta = 110000000000# + rs2!ID_CAJA
            sql = " UPDATE  basasql.dbo.CAJAS "
            sql = sql & "  SET "
            sql = sql & "  FK_CLIENTE  = " & Cliente
            sql = sql & " , NRO_CAJA = " & Caja
            sql = sql & " , FK_ESTADO =1120"
            sql = sql & " , FECHA_CREACION_CAJA  = " & SysDate
            sql = sql & " , FK_USUARIO_CREACION_CAJA =17 "
            sql = sql & " , DIGITO_VERIFICADOR = " & DigitoEAN13(Etiqueta)
            sql = sql & "  Where ID_CAJA = " & rs2!ID_CAJA
            ExecutarSql sql
        End If
    Rem Contenedor
    
        sql = "  SELECT     ESTADO, COD_CLIENTE, NRO_CAJA"
        sql = sql & "  From basasql.dbo.CONTENEDOR"
        sql = sql & "  Where COD_CLIENTE = " & Cliente
        sql = sql & "  And NRO_CAJA = " & Caja
        Set rs = New ADODB.Recordset
        rs.Open sql, strConBasa
        If rs.EOF Then
            sql = "  Update basasql.dbo.CONTENEDOR"
            sql = sql & "   SET COD_CLIENTE =" & Cliente
            sql = sql & "   , ESTADO =2"
            sql = sql & "   , NRO_CAJA =" & Caja
            sql = sql & "  Where ID_CONTENEDOR = " & PosicionLibre
            ExecutarSql sql
        End If
        

 
 
 
 
 
 
 


End Function

Public Function PosicionLibre() As Long
Dim sql As String
Dim rs As New ADODB.Recordset

sql = " SELECT TOP (1) ID_CONTENEDOR, ESTANTERIA "
sql = sql & " From basasql.dbo.CONTENEDOR"
sql = sql & "  WHERE     (COD_CLIENTE IS NULL) AND (ESTANTERIA BETWEEN 150 AND 160) AND (ESTADO = 1)"
rs.Open sql, strConBasa

PosicionLibre = rs!ID_CONTENEDOR


End Function

Public Sub MarcarTipoReferencia(Cliente As Integer, Caja As Long, TipoReferncia As Integer, ControlEstadoIngreso As Boolean)
    Dim rsControlContenedor As New ADODB.Recordset
    Dim sql As String
    Dim Actualizar As Boolean
    Actualizar = False
    If ControlEstadoIngreso = True Then
        sql = " SELECT     ESTADO From basasql.dbo.CONTENEDOR "
        sql = sql & " Where  (estado = 5)"
        sql = sql & " and COD_CLIENTE =  " & Cliente
        sql = sql & " And NRO_CAJA = " & Caja
        rsControlContenedor.Open sql, strConBasa
        
        If rsControlContenedor.EOF Then
            Actualizar = False
        Else
            Actualizar = True
        End If
    Else
        Actualizar = True
    End If
    
    If Actualizar Then
            sql = " Update basasql.dbo.CAJAS "
            sql = sql & vbCrLf & " SET FK_TIPO_REFERENCIA = " & TipoReferncia
            sql = sql & vbCrLf & ", FK_TIPO_REFERENCIA_PERSONAL = " & MDIfrmInicio.StaInicio.Panels(2).Text
            sql = sql & vbCrLf & ", TIPO_REFERENCIA_FECHA = " & SysDate
            sql = sql & vbCrLf & " WHERE FK_CLIENTE = " & Cliente
            sql = sql & vbCrLf & " AND NRO_CAJA = " & Caja
            ExecutarSql sql
    End If

End Sub
