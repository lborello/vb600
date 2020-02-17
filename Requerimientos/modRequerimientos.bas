Attribute VB_Name = "modRequerimientos"


Option Explicit
    Public Zoo As Integer
    Public IDREQUERIMIENTO As Long
    Public TipoConsulta As String
    Public CRequerimientos As New clsRequerimientos
    Public TRequerimientos As New TRequerimientos
    Public EstadoFinal As Integer
    Public IDOperador As Integer
    Public Pedido As Integer
    Public strConBasa  As String
    Public BaseOracle As Boolean
    
    
    Public PasoReportes As String
    Public strPasoPlanillas  As String
    Public PasoImagenes  As String
    Public ClienteOsep  As String
    Public ClienteEcogas  As String
    Public Sucursal As String

    Public strConSoporte As String
    Public ConBasa As ADODB.Connection
    Public Usuario As String
    Public ID_Usuario As Integer
    
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
    
    
    
Public Function FechaFormato(Fecha)
 FechaFormato = " CONVERT(DATETIME, '" & Format(Fecha, "YYYY-MM-DD") & " 00:00:00', 102)"
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
        
        
        
        
        If Not IsNull(rs!DESCRIPCION) Then
             CantidadCaracteres = CantidadCaracteres + Len(rs!DESCRIPCION)
             DescripcionRemito = DescripcionRemito & Trim(rs!DESCRIPCION)
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
Public Function SysDate() As String
    Dim rsDate As ADODB.Recordset
    Dim sql As String
      
        sql = "SELECT     GETDATE() AS FechaHora"
        Set rsDate = New ADODB.Recordset
        rsDate.Open sql, ConActiva, 0, 1
        
      If Not rsDate.EOF Then
        
        SysDate = FechaFormato(rsDate!FechaHora)
        Else
        Rem SysDate = Format(SysDate, "DD/MM/YYYY")
        
        End If
        
End Function

Public Function Inicio()
  On Error GoTo salir
  Dim cad As String, i As Byte, s As Byte, var As Byte
  PasoReportes = "Z:\Sistemas\Basa\Reportes_Sistema\"
  strPasoPlanillas = "Z:\Sistemas\Basa\Planillas\"
  ClienteOsep = "Z:\Sistemas\Basa\ClientesBases\"
  ClienteEcogas = "Z:\Sistemas\Basa\ClientesBases\"

'
'        If FileLen("\\222.15.19.251\basa\Sistemas\Basa\Requerimiento.exe") <> FileLen(App.Path & "\Requerimiento.exe") Then
'            MsgBox "Su sistema desactualizado", vbCritical
'            End
'        End If
'

   
   Open "Z:\Sistemas\Basa\Configuracion.txt" For Input As #1

   
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
        End Select
       
       Wend
       
        Close #1
       
    BaseOracle = False
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    sql = " SELECT     IDPERSONAL, NOMBRE, APELLIDO, USUARIOSYS "
    sql = sql & " From dbo.Personal "
    sql = sql & " Where USUARIOSYS = '" & WindowsUserName & "'"
    rs.Open sql, ConActiva, adOpenStatic, adLockReadOnly
    If Not rs.EOF Then
        Usuario = rs!IDPERSONAL
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
    
    Exit Function
  Exit Function
salir:
  MsgBox Err.Description
  End
  
    
End Function



Public Function SysDate_mm_ss() As String
    Dim rsDate As ADODB.Recordset
    Dim sql As String
      
        sql = "SELECT     GETDATE() AS FechaHora"
        Set rsDate = New ADODB.Recordset
        rsDate.Open sql, ConActiva, 0, 1
        
      If Not rsDate.EOF Then
        
        SysDate_mm_ss = "CONVERT(DATETIME, '" & Format(rsDate!FechaHora, "YYYY-MM-DD hh:mm:ss") & "', 102)"
        Else
        SysDate_mm_ss = "CONVERT(DATETIME, '" & Format(rsDate!FechaHora, "YYYY-MM-DD hh:mm:ss") & "', 102)"
        
        End If
        
End Function

Public Function FechaSolaString(NombreCampo As String) As String
    If BaseOracle = False Then
        FechaSolaString = " CONVERT(CHAR(10)," & NombreCampo & ", 103) "
    Else
        FechaSolaString = " TO_CHAR(" & NombreCampo & ",'DD/MM/YYYY') "
    End If
 End Function

Public Function SysDate_DD_MM_YYYY() As String
    Dim rsDate As ADODB.Recordset
    Dim sql As String
        sql = "SELECT     GETDATE() AS FechaHora"
        Set rsDate = New ADODB.Recordset
        rsDate.Open sql, ConActiva, 0, 1
        SysDate_DD_MM_YYYY = Format(CStr(rsDate!FechaHora), "DD/MM/YYYY")
End Function

Public Function SysDate_DD_MM_YYYY_mm_ss() As String
    Dim rsDate As ADODB.Recordset
    Dim sql As String
        sql = "SELECT     GETDATE() AS FechaHora"
        Set rsDate = New ADODB.Recordset
        rsDate.Open sql, ConActiva, 0, 1
        SysDate_DD_MM_YYYY_mm_ss = Format(CStr(rsDate!FechaHora), "DD/MM/YYYY hh:MM")
End Function
Public Function MaxIDRequerimiento() As Long
    Dim rsMaxRequerimiento As ADODB.Recordset
    Dim sql As String
            Set rsMaxRequerimiento = New ADODB.Recordset
            rsMaxRequerimiento.Open "Select max(idrequerimiento)as maxRequerimiento from Requerimiento ", ConActiva, 0, 1
            If Not rsMaxRequerimiento.EOF Then
               
               
               
                   MaxIDRequerimiento = CLng(rsMaxRequerimiento!maxRequerimiento)
              
            End If
            Set rsMaxRequerimiento = Nothing
End Function

Public Function MaxIDFax()
    Dim rsMaxFax As ADODB.Recordset
    Dim sql As String
            Set rsMaxFax = New ADODB.Recordset
            rsMaxFax.Open "Select max(idFax)as maxfax from Fax ", ConActiva, 0, 1
            If Not rsMaxFax.EOF Then
               If IsNull(rsMaxFax!MaxFax) Then
                   MaxIDFax = 1
               Else
                   MaxIDFax = CLng(rsMaxFax!MaxFax) + 1
               End If
            End If
            Set rsMaxFax = Nothing

End Function

Public Function SysDateCompare() As String
    Dim rsDate As ADODB.Recordset
    Dim sql As String
       If BaseOracle = True Then
        sql = " SELECT "
        sql = sql & " TO_CHAR(SYSDATE, 'DD/MM/YYYY HH24:MI') as SysDate "
        sql = sql & " FROM DUAL "
        Set rsDate = New ADODB.Recordset
        rsDate.Open sql, ConActiva, 0, 1
        SysDateCompare = Format(CStr(rsDate!SysDate), "DD/MM/YYYY HH:MM")
        Else
        Rem SysDateCompare = Format(SysDate, "DD/MM/YYYY HH:MM")
        End If
    
End Function

Public Sub CambirarEstado(IDREQUERIMIENTO As String, IDPERSONAL As Integer, EstadoOrigen As Integer, EstadoDestino As Integer, CambiarEstadoRequerimiento As Boolean)
    Dim rs As ADODB.Recordset
    Dim RSH_ESTADO_REQUE As ADODB.Recordset
    Dim sql As String
    Dim FECHARECEPCION As Date
    Dim IDTIPOREQUERIMIENTO As Integer
    Dim Filtro As String
    Dim CONTADOR As Integer
        If CambiarEstadoRequerimiento = True Then
            sql = " UPDATE REQUERIMIENTO SET"
            sql = sql & " IDESTADO=" & EstadoDestino
            sql = sql & " where idRequerimiento IN  " & IDREQUERIMIENTO
            ExecutarSql (sql)
            sql = " update requerimiento set idPersonal = " & IDPERSONAL
            sql = sql & " where idrequerimiento = " & IDREQUERIMIENTO
            ExecutarSql (sql)
        End If
        sql = " SELECT * From  REQUERIMIENTO  Where IDRequerimiento in " & IDREQUERIMIENTO
        Set rs = New ADODB.Recordset
        rs.Open sql, ConActiva, 0, 1
            Do While Not rs.EOF
                    sql = " SELECT max(Contador)AS CONTADOR From  H_ESTADO_REQUE  Where IDRequerimiento = " & CLng(rs!IDREQUERIMIENTO)
                    Set RSH_ESTADO_REQUE = New ADODB.Recordset
                     RSH_ESTADO_REQUE.Open sql, ConActiva, 0, 1
                    If Not RSH_ESTADO_REQUE.EOF Then
                        If IsNull(RSH_ESTADO_REQUE!CONTADOR) Then
                            CONTADOR = 1
                        Else
                          If CambiarEstadoRequerimiento Then
                            CONTADOR = CInt(RSH_ESTADO_REQUE!CONTADOR) + 1
                          Else
                            CONTADOR = CInt(RSH_ESTADO_REQUE!CONTADOR)
                          End If
                        End If
                    Else
                        CONTADOR = 1
                    End If
                        sql = " INSERT INTO H_ESTADO_REQUE ("
                        sql = sql & " IDREQUERIMIENTO, IDESTADO, IDPERSONAL,"
                        sql = sql & " CONTADOR, FECHA )"
                        sql = sql & "  VALUES ("
                        sql = sql & CLng(rs!IDREQUERIMIENTO) & "," & EstadoDestino & "," & IDPERSONAL & ","
                        sql = sql & CONTADOR & "," & SysDate & ")"
                        ExecutarSql (sql)
                        rs.MoveNext
               Loop
               frmControlEstados.CargarTree
End Sub

Public Sub CambiarEstadoTipoConsualta(TipoConsualta As String, IDPERSONAL As Integer, EstadoOrigen As Integer, EstadoDestino As Integer, CambiarEstadoRequerimiento As Boolean)
    Dim rs As ADODB.Recordset
    Dim RSH_ESTADO_REQUE As ADODB.Recordset
    Dim sql As String
    Dim FECHARECEPCION As Date
    Dim IDTIPOREQUERIMIENTO As Integer
    Dim Filtro As String
    Dim CONTADOR As Integer
        FECHARECEPCION = Mid(TipoConsualta, 3, 10)
        IDTIPOREQUERIMIENTO = Mid(TipoConsualta, 14)
        If CambiarEstadoRequerimiento Then
            sql = " select * from requerimiento where  " & FechaSolaString("FECHARECEPCION") & " = '" & FECHARECEPCION & "'"
            sql = sql & "' and  IDTIPOREQUERIMIENTO = " & IDTIPOREQUERIMIENTO
            sql = sql & " and idestado = " & EstadoOrigen
        Else
            sql = " select * from requerimiento where " & FechaSolaString("FECHARECEPCION") & " = '" & FECHARECEPCION & "'"
            sql = sql & "' and  IDTIPOREQUERIMIENTO = " & IDTIPOREQUERIMIENTO
            sql = sql & " and idestado = " & EstadoDestino
        End If
        Set rs = New ADODB.Recordset
         rs.Open sql, ConActiva, 0, 1
         Filtro = ""
        Do While Not rs.EOF
            Filtro = Filtro & "," & CStr(rs!IDREQUERIMIENTO)
            rs.MoveNext
        Loop
        Filtro = "(" & Mid(Filtro, 2) & ")"
        If Filtro <> "" And Filtro <> "()" Then
            If CambiarEstadoRequerimiento = True Then
                sql = " UPDATE REQUERIMIENTO SET"
                sql = sql & " IDESTADO=" & EstadoDestino
                sql = sql & " where idRequerimiento in  " & Filtro
                sql = sql & " and IDestado =" & EstadoOrigen
                ExecutarSql (sql)
            End If
        End If
        sql = " select * from requerimiento where  IDREQUERIMIENTO IN " & Filtro
        Set rs = New ADODB.Recordset
        rs.Open sql, ConActiva, 0, 1
            Do While Not rs.EOF
                    sql = " SELECT max(Contador)AS CONTADOR From  H_ESTADO_REQUE  Where IDRequerimiento =" & CLng(rs!IDREQUERIMIENTO)
                    Set RSH_ESTADO_REQUE = New ADODB.Recordset
                     RSH_ESTADO_REQUE.Open sql, ConActiva, 0, 1
                    If Not RSH_ESTADO_REQUE.EOF Then
                            If IsNull(RSH_ESTADO_REQUE!CONTADOR) Then
                                CONTADOR = 1
                            Else
                                If CambiarEstadoRequerimiento Then
                                    CONTADOR = CInt(RSH_ESTADO_REQUE!CONTADOR) + 1
                                Else
                                    CONTADOR = CInt(RSH_ESTADO_REQUE!CONTADOR)
                                End If
                            End If
                    Else
                        CONTADOR = 1
                    End If
                    sql = " INSERT INTO H_ESTADO_REQUE ("
                    sql = sql & " IDREQUERIMIENTO, IDESTADO, IDPERSONAL,"
                    sql = sql & " CONTADOR, FECHA )"
                    sql = sql & "  VALUES ("
                    sql = sql & CLng(rs!IDREQUERIMIENTO) & "," & EstadoDestino & "," & IDPERSONAL & ","
                    sql = sql & CONTADOR & "," & SysDate & ")"
                    ExecutarSql (sql)
                    rs.MoveNext
            Loop
            frmControlEstados.CargarTree
End Sub

Public Sub Insert_Requerimiento_Historico_Descripcion(FK_REQUERIMIENTO As Long, DESCRIPCION As String, FK_USUARIO As Integer, Fecha As String, conReq As ADODB.Connection)
    Dim sql As String
    
    sql = "INSERT INTO dbo.REQUERIMIENTO_DESCRIPCION_HISTORICO"
    sql = sql & " (FK_REQUERIMIENTO, DESCRIPCION, FK_USUARIO, FECHA)"
    sql = sql & " VALUES     "
    sql = sql & " (" & FK_REQUERIMIENTO & "," & Trim(DESCRIPCION) & "," & FK_USUARIO & "," & Fecha & ")"
    conReq.Execute sql



End Sub

Public Function ExecutarSql(sql As String) As Integer
Dim con As New ADODB.Connection
Dim Registros As Integer
On Error GoTo salir:

con.Open strConBasa
con.BeginTrans
con.Execute sql, Registros

ExecutarSql = Registros
con.CommitTrans
con.Close

Exit Function

salir:
con.RollbackTrans
con.Close
MsgBox Err.Description

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

Public Function ConActiva() As ADODB.Connection

On Error GoTo salir:

'If IsObject(ConBasa) Then
'    If ConBasa.State = 0 Then
'        Set ConBasa = New ADODB.Connection
        ConBasa.CursorLocation = adUseClient
        ConBasa.Open strConBasa
        Set ConActiva = ConBasa
'    Else
'     Set ConActiva = ConBasa
'    End If
'Else
'    Set ConBasa = New ADODB.Connection
'        ConBasa.Open strConBasa
'        Set ConActiva = ConBasa
'    End If
    Exit Function
salir:
    Set ConBasa = New ADODB.Connection
    ConBasa.CursorLocation = adUseClient
    ConBasa.Open strConBasa
    Set ConActiva = ConBasa
End Function
