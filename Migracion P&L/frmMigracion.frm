VERSION 5.00
Begin VB.Form frmMigracion 
   Caption         =   "Form4"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10860
   LinkTopic       =   "Form4"
   ScaleHeight     =   7845
   ScaleWidth      =   10860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   7800
      TabIndex        =   8
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MIgracionGodoy cruz"
      Height          =   855
      Left            =   5160
      TabIndex        =   7
      Top             =   4680
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   675
      Left            =   540
      TabIndex        =   6
      Top             =   4380
      Width           =   2535
   End
   Begin VB.CommandButton cmdControlFondo 
      Caption         =   "cmdControlFondo"
      Height          =   795
      Left            =   420
      TabIndex        =   3
      Top             =   1380
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   795
      Left            =   6420
      TabIndex        =   2
      Top             =   1260
      Width           =   2595
   End
   Begin VB.Frame frmMigracionCajas 
      Caption         =   "Migracion Cajas"
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8715
      Begin VB.CommandButton cmdMigracionCajas 
         Caption         =   "Migracion Cajas"
         Height          =   375
         Left            =   6840
         TabIndex        =   1
         Top             =   420
         Width           =   1575
      End
   End
   Begin VB.Label lblorden 
      Caption         =   "Label1"
      Height          =   675
      Left            =   4860
      TabIndex        =   5
      Top             =   3060
      Width           =   1875
   End
   Begin VB.Label lblcaja 
      Caption         =   "Label1"
      Height          =   735
      Left            =   1980
      TabIndex        =   4
      Top             =   2940
      Width           =   1815
   End
End
Attribute VB_Name = "frmMigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim strConAsp150 As String
    Dim conPYL As New ADODB.Connection
    Dim strConGodoyCruz As String
    Dim strBasa As String
    

    

Private Sub cmdControlFondo_Click()

    Dim ConAsp As New ADODB.Connection
    ConAsp.Open strConAsp150
    Dim RSASP As New ADODB.Recordset
    
    
    Dim conFondo As New ADODB.Connection
    conFondo.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
    Dim rsFondo As New ADODB.Recordset
    
    
    
    Dim sql As String
    Dim sqlASP As String
    
    
    
    
        sqlASP = " SELECT  id, codigo, estado, generaCanonMensual, clienteAsp_id, clienteEmp_id, contenedor_id, posicion_id, tipoElemento_id, depositoActual_id, codigoAlternativo,"
        sqlASP = sqlASP & " nroPrecinto, estadoTrabajo, fechaModificacion, fechaModificacionPrecinto, nroPrecinto1, tipoTrabajo, usuarioModificacion_id, usuarioModificacionPrecinto_id,"
        sqlASP = sqlASP & " ubicacionProvisoria , cerrado, habilitadoCerrar"
        sqlASP = sqlASP & " From elementos"
        sqlASP = sqlASP & " WHERE     (codigo BETWEEN '120006707861' AND '120006771860')"
        sqlASP = sqlASP & " ORDER BY codigo"

Set RSASP = New ADODB.Recordset
RSASP.Open sqlASP, ConAsp


sql = " SELECT     Id, LEGAJO_ASP, ID_LEGAJO_ASP"
sql = sql & " From REFERENCIAS_FONDO_COMPLETA"
sql = sql & "  ORDER BY Id"

rsFondo.Open sql, conFondo

    
    Do While Not RSASP.EOF
             sql = " Update REFERENCIAS_FONDO_COMPLETA"
            sql = sql & " SET LEGAJO_ASP =" & RSASP!codigo
            sql = sql & ", ID_LEGAJO_ASP =" & RSASP!ID
            sql = sql & " Where ID = " & rsFondo!ID
            conFondo.Execute sql
            rsFondo.MoveNext
            RSASP.MoveNext
    Loop
'
'        sql = " SELECT [Id] ,[EXPEDIENTES] ,[Caja]"
'        sql = sql & "  ,[NUMERO] ,[LETRA]  ,[Ambito]"
'        sql = sql & "   ,[AÑO] ,[Cuerpo]  ,[Descripcion]"
'        sql = sql & "  ,[ESTADO] ,[Nombre] ,[Expr1]"
'        sql = sql & "  ,[ISC] ,[CAJAPYL] ,[ID_EMPRESA_ASP]"
'        sql = sql & "  ,[CajaAzul] ,[Caja_Asp] ,[ID_ASP]"
'        sql = sql & "  From [P&LCUSTODIA].[dbo].[REFERENCIAS_FONDO_COMPLETA]"


End Sub

Private Sub cmdMigracionCajas_Click()

    Dim sql As String
    Dim IDEmpresa As Integer
    
    
    Dim ConAsp As New ADODB.Connection
    ConAsp.Open strConAsp150
    Dim RSASP As New ADODB.Recordset
    
    
    Dim rsPYL As New ADODB.Recordset
    Dim lote_referencia_id As Long
    



IDEmpresa = 20016

    
            
            
            
            sql = " SELECT Caja.Id, Caja.Numero, Caja.Fecha, Caja.IdEmpresa, "
            sql = sql & " Caja.CajaAzul , Caja.CAJA_ASP , caja.ID_ASP, caja.ID_CLIENTE_ASP,"
            sql = sql & " Empresa.Nombre, Empresa.Id AS IDEMPRESA, Empresa.ID_EMPRESA_ASP "
            sql = sql & " FROM Caja INNER JOIN Empresa ON Caja.IdEmpresa = Empresa.Id "
            sql = sql & " Where Empresa.ID = " & IDEmpresa
            sql = sql & " And (Not (caja.CAJA_ASP Is Null)) "
            
            rsPYL.Open sql, conPYL
        
        
        
        lote_referencia_id = InsertLoteReferencia(rsPYL!ID_EMPRESA_ASP)
        
            Do While Not rsPYL.EOF
                
                    sql = " SELECT ID, codigo, estado, generaCanonMensual, clienteAsp_id, clienteEmp_id, contenedor_id, posicion_id, tipoElemento_id, depositoActual_id, codigoAlternativo,"
                    sql = sql & " nroPrecinto, estadoTrabajo, fechaModificacion, fechaModificacionPrecinto, nroPrecinto1, tipoTrabajo, usuarioModificacion_id, usuarioModificacionPrecinto_id,"
                    sql = sql & " ubicacionProvisoria , cerrado, habilitadoCerrar"
                    sql = sql & " From elementos "
                    sql = sql & " WHERE (codigo LIKE '" & rsPYL!CAJA_ASP & "')"
                    
                    Set RSASP = New ADODB.Recordset
                    
                    RSASP.Open sql, ConAsp
                    
                    If Not RSASP.EOF Then
                       
                       If IsNull(RSASP!clienteEmp_id) Then
                      
                       
                            If Trim(RSASP!estado) = "En Guarda" Then
                                    sql = " Update elementos"
                                    sql = sql & " SET clienteEmp_id =" & rsPYL!ID_EMPRESA_ASP
                                    sql = sql & "   WHERE    ID = " & RSASP!ID
                            Else
                            
                            If Trim(RSASP!estado) = "Creado" Then
                                sql = " Update elementos "
                                sql = sql & " SET clienteEmp_id = " & rsPYL!ID_EMPRESA_ASP
                                sql = sql & " , estado ='En Guarda'"
                                sql = sql & "   WHERE    ID = " & RSASP!ID
                            Else
                            
                                sql = " Update elementos "
                                sql = sql & " SET clienteEmp_id = " & rsPYL!ID_EMPRESA_ASP
                                Rem sql = sql & " , estado ='En Guarda'"
                                sql = sql & "   WHERE    ID = " & RSASP!ID
                            End If
                            
                            
                            End If
                            ConAsp.Execute sql
                              End If
                                    sql = " INSERT INTO referencia"
                                    sql = sql & vbCrLf & " ( indice_individual"
                                    sql = sql & vbCrLf & " , prefijoCodigoTipoElemento"
                                    sql = sql & vbCrLf & " , texto1"
                                    sql = sql & vbCrLf & " , texto2"
                                    sql = sql & vbCrLf & " , clasificacion_documental_id"
                                    sql = sql & vbCrLf & " , elemento_id"
                                    sql = sql & vbCrLf & " , lote_referencia_id"
                                    sql = sql & vbCrLf & " )"
                                    sql = sql & vbCrLf & " VALUES  "
                                    sql = sql & vbCrLf & " (0"
                                    sql = sql & vbCrLf & " ,'11'"
                                    sql = sql & vbCrLf & " ,'" & rsPYL!numero & "'"
                                    sql = sql & vbCrLf & " ,'" & rsPYL!numero & "'"
                                    sql = sql & vbCrLf & " , " & clasificacion_documental_id_cajaspyl(rsPYL!ID_EMPRESA_ASP)
                                    sql = sql & vbCrLf & " , " & RSASP!ID
                                    sql = sql & vbCrLf & " , " & lote_referencia_id
                                    sql = sql & vbCrLf & ")"
                            
                               ConAsp.Execute sql
                        
                     
                       
                      
                         
                    End If
                    
                    
            
                rsPYL.MoveNext
            Loop
            
            
            
 MsgBox "TERMINADO"

End Sub

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Dim sql As String
Dim Descripcion As String
Dim fecha1 As String
Dim fecha2 As String
Dim numero1 As String
Dim numero2 As String
Dim texto1 As String
Dim texto2 As String
Dim clasificacion_documental_id As Integer
Dim elemento_id As Long
Dim lote_referencia_id As Long
Dim descripcion_rearchivo As String
Dim referencia_rearchivo_id As String
Dim prefijoCodigoTipoElemento As String
Dim ordenRearchivo As Integer
Dim pathArchivoDigital As String
Dim indice_individual As Integer
Dim ConAsp As New ADODB.Connection
Dim CajaAnterior As Long
Dim conBasa As New ADODB.Connection
Dim ID_ASP_CAJA As Long
Dim rsPYLDocumentos As New ADODB.Recordset
Dim conPYL  As New ADODB.Connection


conPYL.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"

ConAsp.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"



            sql = " SELECT     ID_LEGAJO, CODIGO_ASP_ELEMENTO_LEGAJO, CODIGO_ASP_ELEMENTO_CAJA, ID_ASP_LEGAJO, ID_ASP_CAJA"
            sql = sql & vbCrLf & " , LETRA_DESDE AS texto1 , LETRA_HASTA AS texto2, NRO_DESDE AS numero1"
            sql = sql & vbCrLf & ", NRO_HASTA AS numero2, CONVERT( CHAR ,FECHA_DESDE ,103) AS fecha1  "
            sql = sql & vbCrLf & " , CONVERT(CHAR, FECHA_HASTA, 103) AS fecha2 "
            sql = sql & vbCrLf & ", DESCRIPCION AS DESCRIPCION_ASP , NRO_CAJA, ID_ASP_REFERENCIAS"
            sql = sql & vbCrLf & " From LEGAJOS"
            sql = sql & vbCrLf & " Where (COD_CLIENTE = 128) and NRO_CAJA > 797282  "
            sql = sql & vbCrLf & " ORDER BY NRO_CAJA, ID_LEGAJO"
            
    
    sql = " SELECT     Empresa.Nombre, Caja.Id AS IDCAJA, Caja.Numero, Documento.Id AS IDDOCUMENTO, Documento.NUMERO1, Documento.NUMERO2, Documento.TEXTO1,"
    sql = sql & vbCrLf & " Documento.TEXTO2, CONVERT(char, Documento.FECHA1, 103) AS FECHA1, CONVERT(char, Documento.FECHA2, 103) AS FECHA2, Documento.DESCRIPCION_ASP,"
    sql = sql & vbCrLf & " Documento.CODIGO_ASP_ELEMENTO_LEGAJO , Documento.ID_ASP_LEGAJO ,  Empresa.Id AS IDEMPRESA "
    sql = sql & vbCrLf & " FROM         Documento INNER JOIN"
    sql = sql & vbCrLf & " Caja ON Documento.IdCaja = Caja.Id INNER JOIN"
    sql = sql & vbCrLf & " Empresa ON Caja.IdEmpresa = Empresa.Id"
    sql = sql & vbCrLf & " WHERE     (Empresa.Nombre LIKE N'%MUNI%')  AND (Documento.ID_ASP_REFERENCIAS IS NULL) "
    sql = sql & vbCrLf & " ORDER BY Documento.IdCaja, IDDOCUMENTO"
    
    
    
  
sql = "  SELECT     Id AS IDREFERENCIAS_FONDO_COMPLETA, EXPEDIENTES, Caja, NUMERO, LETRA, Ambito, AÑO, Cuerpo, Descripcion, ESTADO, Nombre, Expr1, ISC, CAJAPYL, ID_EMPRESA_ASP, CajaAzul, Caja_Asp,"
sql = sql & vbCrLf & "                      ID_ASP_CAJA as IDCAJA , LEGAJO_ASP as CODIGO_ASP_ELEMENTO_LEGAJO , ID_LEGAJO_ASP as ID_ASP_LEGAJO "
sql = sql & vbCrLf & " From [P&LCUSTODIA].dbo.REFERENCIAS_FONDO_COMPLETA"
sql = sql & vbCrLf & " ORDER BY ID_ASP_CAJA , Id "
    
    
    
    
    rs.Open sql, conPYL, adOpenStatic, adLockReadOnly
            
            
            
            
            
            
    Rem         rs.Open sql, strConBasa

           Rem conAsp.Open strConAsp


Do While Not rs.EOF

    If CajaAnterior <> rs!IDCAJA Then
        lote_referencia_id = InsertLoteReferencia(rs!ID_EMPRESA_ASP)
        ordenRearchivo = 0
        CajaAnterior = rs!IDCAJA
        sql = " Update elementos"
        sql = sql & vbCrLf & " Set cerrado = 1"
        sql = sql & vbCrLf & " Where ID = " & rs!IDCAJA
        ConAsp.Execute sql
       
    End If
    
'        sql = " Update Documento "
'        sql = sql & vbCrLf & "  Set ID_ASP_REFERENCIAS = 0109"
'        sql = sql & vbCrLf & "  Where ID = " & RS!IDDOCUMENTO
'        conPYL.Execute sql
        lblcaja.Caption = ID_ASP_CAJA
        lblcaja.Refresh
        lblorden.Caption = ordenRearchivo
        lblorden.Refresh
        indice_individual = 1
        clasificacion_documental_id = 20241
        elemento_id = rs!ID_ASP_LEGAJO
    
    
    sql = " Update elementos"
    sql = sql & vbCrLf & " Set contenedor_id = " & rs!IDCAJA
    sql = sql & vbCrLf & ", clienteEmp_id =" & rs!ID_EMPRESA_ASP
    sql = sql & vbCrLf & " Where ID = " & rs!ID_ASP_LEGAJO
    sql = sql & vbCrLf & " And (contenedor_id Is Null)"
    
    ConAsp.Execute sql
   
    prefijoCodigoTipoElemento = "'" & Mid(rs!CODIGO_ASP_ELEMENTO_LEGAJO, 1, 2) & "'"
    ordenRearchivo = ordenRearchivo + 1
    pathArchivoDigital = "Null"
    descripcion_rearchivo = "Null"
    referencia_rearchivo_id = "NULL"
    
    Descripcion = "'" & Trim(rs!Descripcion) & " CAJAP&L: " & rs!caja & "'"
    
'    If Not IsNull(RS!DESCRIPCION_ASP) And Len(Trim(RS!DESCRIPCION_ASP)) > 0 Then
'        Descripcion = "'" & Trim(RS!DESCRIPCION_ASP) & "'"
'    Else
'        Descripcion = "NULL"
'    End If
    
fecha1 = "NULL"
fecha2 = "NULL"

   If IsNumeric(rs!año) Then
       If Len(rs!año) = 4 Then
        fecha1 = FechaFormato("01/01/" & rs!año)
         fecha2 = FechaFormato("31/12/" & rs!año)
       
       End If
       
   End If
   
'
'    If Not IsNull(RS!fecha1) And IsDate(RS!fecha1) Then
'        fecha1 = FechaFormato(Trim(RS!fecha1))
'    Else
'        fecha1 = "NULL"
'    End If
'
'
'    If Not IsNull(RS!fecha2) And IsDate(RS!fecha2) Then
'        fecha2 = FechaFormato(Trim(RS!fecha2))
'    Else
'        fecha2 = "NULL"
'    End If
    
    
    If Not IsNull(rs!numero) And IsNumeric(rs!numero) Then
        numero1 = rs!numero
    Else
        numero1 = "Null"
    End If
    
    
    If Not IsNull(rs!Cuerpo) Then
        If IsNumeric(rs!Cuerpo) Then
            numero2 = rs!Cuerpo
      
        Else
             numero2 = "1"
        End If
    Else
        numero2 = "1"
    End If
    
    
'    If Not IsNull(RS!numero2) And IsNumeric(RS!numero2) Then
'        numero2 = RS!numero2
'    Else
'        numero2 = "Null"
'    End If
    
    If Not IsNull(rs!LETRA) And Len(Trim(rs!LETRA)) > 0 Then
        texto1 = "'" & Trim(rs!LETRA) & "'"
    Else
        texto1 = "NULL"
    End If
    
    If Not IsNull(rs!LETRA) And Len(Trim(rs!LETRA)) > 0 Then
        texto2 = "'" & Trim(rs!LETRA) & "'"
    Else
        texto2 = "NULL"
    End If




        sql = " INSERT INTO referencia"
        sql = sql & vbCrLf & "("
        sql = sql & vbCrLf & "descripcion"
        sql = sql & vbCrLf & ", fecha1"
        sql = sql & vbCrLf & ", fecha2"
        sql = sql & vbCrLf & ", indice_individual"
        sql = sql & vbCrLf & ", numero1"
        sql = sql & vbCrLf & ", numero2"
        sql = sql & vbCrLf & ", texto1"
        sql = sql & vbCrLf & ", texto2"
        sql = sql & vbCrLf & ", clasificacion_documental_id"
        sql = sql & vbCrLf & ", elemento_id"
        sql = sql & vbCrLf & ", lote_referencia_id"
        sql = sql & vbCrLf & ", descripcion_rearchivo"
        sql = sql & vbCrLf & ", referencia_rearchivo_id"
        sql = sql & vbCrLf & ", prefijoCodigoTipoElemento"
        sql = sql & vbCrLf & ", ordenRearchivo"
        sql = sql & vbCrLf & ", pathArchivoDigital"
        sql = sql & vbCrLf & ")"
        sql = sql & vbCrLf & " VALUES ("
        sql = sql & vbCrLf & Descripcion
        sql = sql & vbCrLf & "," & fecha1
        sql = sql & vbCrLf & "," & fecha2
        sql = sql & vbCrLf & "," & indice_individual
        sql = sql & vbCrLf & "," & numero1
        sql = sql & vbCrLf & "," & numero2
        sql = sql & vbCrLf & "," & texto1
        sql = sql & vbCrLf & "," & texto2
        sql = sql & vbCrLf & "," & clasificacion_documental_id
        sql = sql & vbCrLf & "," & elemento_id
        sql = sql & vbCrLf & "," & lote_referencia_id
        sql = sql & vbCrLf & "," & descripcion_rearchivo
        sql = sql & vbCrLf & "," & referencia_rearchivo_id
        sql = sql & vbCrLf & "," & prefijoCodigoTipoElemento
        sql = sql & vbCrLf & "," & ordenRearchivo
        sql = sql & vbCrLf & "," & pathArchivoDigital
        sql = sql & vbCrLf & ")"
    
    ConAsp.Execute sql
    
  sql = "   Update [P&LCUSTODIA].dbo.REFERENCIAS_FONDO_COMPLETA"
sql = sql & vbCrLf & " SET              PASADOASP = 'SI'"
sql = sql & vbCrLf & " Where ID = " & rs!IDREFERENCIAS_FONDO_COMPLETA
conPYL.Execute sql
    
    
    rs.MoveNext
Loop



End Sub
Public Function InsertLoteReferencia(cliente_emp_id As String) As Long
    Dim sql As String
    Dim RsMax As New ADODB.Recordset
    Dim MaxCodigo As Long
    
 Dim fecha_registro As String
 Dim cliente_asp_id As String

 Dim empresa_id As String
 Dim sucursal_id As String
 Dim habilitado As String
 Dim codigo As String
 Dim cargaPorRango As String
 Dim Conlotereferencia As New ADODB.Connection
    
    
    sql = " SELECT MAX(codigo) AS MaxCodigo"
    sql = sql & " From lotereferencia "
    
'Set RsMax = New ADODB.Recordset
'RsMax.Open sql, straso
'codigo = RsMax!MaxCodigo + 1



fecha_registro = FechaFormato(Now)
cliente_asp_id = "1"
 cliente_emp_id = cliente_emp_id
 empresa_id = "20004"
 sucursal_id = "30010"
 habilitado = "NULL"
 codigo = codigo
 cargaPorRango = "NULL"


sql = "INSERT INTO lotereferencia"
 sql = sql & vbCrLf & "("
 sql = sql & vbCrLf & "fecha_registro"
 sql = sql & vbCrLf & ", cliente_asp_id"
 sql = sql & vbCrLf & ", cliente_emp_id"
 sql = sql & vbCrLf & ", empresa_id"
 sql = sql & vbCrLf & ", sucursal_id"
 sql = sql & vbCrLf & ", habilitado"
 sql = sql & vbCrLf & ", codigo"
 sql = sql & vbCrLf & ", cargaPorRango"
 sql = sql & vbCrLf & ")"
 sql = sql & vbCrLf & " VALUES ("
  sql = sql & vbCrLf & fecha_registro
 sql = sql & vbCrLf & "," & cliente_asp_id
 sql = sql & vbCrLf & "," & cliente_emp_id
 sql = sql & vbCrLf & "," & empresa_id
 sql = sql & vbCrLf & "," & sucursal_id
 sql = sql & vbCrLf & "," & habilitado
 sql = sql & vbCrLf & "," & codigo
 sql = sql & vbCrLf & "," & cargaPorRango
 sql = sql & vbCrLf & ")"
 
 Conlotereferencia.Open strConAsp150
 Conlotereferencia.Execute sql
 
 
  
    
sql = " SELECT  MAX(id) AS MaxID "
sql = sql & "  From lotereferencia "
sql = sql & "  Where codigo = " & codigo
    
Set RsMax = New ADODB.Recordset
RsMax.Open sql, strConAsp150
InsertLoteReferencia = RsMax!MaxID
 

End Function

Public Function InsertLoteReferenciaGodoyCruz() As Long
    Dim sql As String
    Dim RsMax As New ADODB.Recordset
    Dim MaxCodigo As Long
    
 Dim fecha_registro As String
 Dim cliente_asp_id As String

 Dim empresa_id As String
 Dim sucursal_id As String
 Dim habilitado As String
 Dim codigo As String
 Dim cargaPorRango As String
 Dim Conlotereferencia As New ADODB.Connection
 Dim cliente_emp_id  As Long
    
    sql = " SELECT MAX(codigo) AS MaxCodigo"
    sql = sql & " From lotereferencia "
    
Set RsMax = New ADODB.Recordset
RsMax.Open sql, strConGodoyCruz
codigo = RsMax!MaxCodigo + 1



fecha_registro = FechaFormato(Now)
cliente_asp_id = "1"
 cliente_emp_id = 1
 empresa_id = "1"
 sucursal_id = "1"
 habilitado = "NULL"
 codigo = codigo
 cargaPorRango = "NULL"


sql = "INSERT INTO lotereferencia"
 sql = sql & vbCrLf & "("
 sql = sql & vbCrLf & "fecha_registro"
 sql = sql & vbCrLf & ", cliente_asp_id"
 sql = sql & vbCrLf & ", cliente_emp_id"
 sql = sql & vbCrLf & ", empresa_id"
 sql = sql & vbCrLf & ", sucursal_id"
 sql = sql & vbCrLf & ", habilitado"
 sql = sql & vbCrLf & ", codigo"
 sql = sql & vbCrLf & ", cargaPorRango"
 sql = sql & vbCrLf & ")"
 sql = sql & vbCrLf & " VALUES ("
  sql = sql & vbCrLf & fecha_registro
 sql = sql & vbCrLf & "," & cliente_asp_id
 sql = sql & vbCrLf & "," & cliente_emp_id
 sql = sql & vbCrLf & "," & empresa_id
 sql = sql & vbCrLf & "," & sucursal_id
 sql = sql & vbCrLf & "," & habilitado
 sql = sql & vbCrLf & "," & codigo
 sql = sql & vbCrLf & "," & cargaPorRango
 sql = sql & vbCrLf & ")"
 
 Conlotereferencia.Open strConGodoyCruz
 Conlotereferencia.Execute sql
 
 
  
    
sql = " SELECT  MAX(id) AS MaxID "
sql = sql & "  From lotereferencia "
sql = sql & "  Where codigo = " & codigo
    
Set RsMax = New ADODB.Recordset
RsMax.Open sql, strConGodoyCruz
InsertLoteReferenciaGodoyCruz = RsMax!MaxID
 

End Function



Public Function clasificacion_documental_id_cajaspyl(cliente_emp_id As Long) As Long
        
        
        Dim sql As String
        Dim rs As New ADODB.Recordset
        
        sql = " SELECT     id "
        sql = sql & " From clasificacionDocumental"
        sql = sql & " WHERE codigo='10' and "
        sql = sql & " cliente_emp_id = " & cliente_emp_id
        
        rs.Open sql, strConAsp150
        
        
        If Not rs.EOF Then
            clasificacion_documental_id_cajaspyl = rs!ID
        Else
            MsgBox "eRROR"
            clasificacion_documental_id_cajaspyl = 0
        End If
        




End Function

Private Sub Command2_Click()

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim RSASP As New ADODB.Recordset
    Dim Com150 As New ADODB.Connection
    Dim ComPYL As New ADODB.Connection
    
        Com150.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
        ComPYL.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
        
        
        sql = " SELECT  Caja_Asp, ID_ASP_CAJA "
        sql = sql & " From FondoNuevaRelacion "
        sql = sql & " GROUP BY Caja_Asp, ID_ASP_CAJA "
    
    rs.Open sql, ComPYL
    
    Do While Not rs.EOF
        sql = " SELECT id, codigo, estado, generaCanonMensual "
        sql = sql & " From elementos "
        sql = sql & " WHERE codigo = '" & rs!CAJA_ASP & "'"
        Set RSASP = New ADODB.Recordset
        RSASP.Open sql, Com150
        If Not RSASP.EOF Then
                sql = " Update FondoNuevaRelacion "
                sql = sql & " Set ID_ASP_CAJA = " & RSASP!ID
                sql = sql & " Where CAJA_ASP =" & RSASP!codigo
                ComPYL.Execute sql
        End If
        rs.MoveNext
    Loop

End Sub

Private Sub Command3_Click()
Dim rs As New ADODB.Recordset
Dim sql As String
Dim Descripcion As String
Dim fecha1 As String
Dim fecha2 As String
Dim numero1 As String
Dim numero2 As String
Dim texto1 As String
Dim texto2 As String
Dim clasificacion_documental_id As Integer
Dim elemento_id As Long
Dim lote_referencia_id As Long
Dim descripcion_rearchivo As String
Dim referencia_rearchivo_id As String
Dim prefijoCodigoTipoElemento As String
Dim ordenRearchivo As Integer
Dim pathArchivoDigital As String
Dim indice_individual As Integer

Dim CajaAnterior As Long
Dim conBasa As New ADODB.Connection
Dim ID_ASP_CAJA As Long
Dim ConGodoyCruz  As New ADODB.Connection

ConGodoyCruz.Open strConGodoyCruz


sql = " SELECT        LEGAJOS.LETRA_DESDE, LEGAJOS.LETRA_HASTA, LEGAJOS.NRO_DESDE, LEGAJOS.NRO_HASTA, LEGAJOS.FECHA_DESDE, LEGAJOS.FECHA_HASTA,"
sql = sql & vbCrLf & "                          LEGAJOS.DESCRIPCION, LEGAJOS.NRO_CAJA, LEGAJOS.COD_CLIENTE, INDICES.INDICE, INDICES.ID_GODOYCRUZ, LEGAJOS.ETIQUETA,"
sql = sql & vbCrLf & "                         LEGAJOS.ID_LEGAJO"
sql = sql & vbCrLf & "  FROM            LEGAJOS INNER JOIN"
sql = sql & vbCrLf & "                         INDICES ON LEGAJOS.FK_INDICES = INDICES.ID"
sql = sql & vbCrLf & "  WHERE        (LEGAJOS.COD_CLIENTE = 1156) AND (INDICES.INDICE LIKE '001001%')"
sql = sql & vbCrLf & "  ORDER BY LEGAJOS.NRO_CAJA, LEGAJOS.ID_LEGAJO"


    rs.CursorLocation = adUseClient
    
    
    rs.Open sql, strBasa
            
  Dim cantidadElemento As Long
            
            
            
         

Do While Not rs.EOF

    If CajaAnterior <> rs!NRO_CAJA Then
        lote_referencia_id = InsertLoteReferenciaGodoyCruz
        ordenRearchivo = 0
        CajaAnterior = rs!NRO_CAJA
        cantidadElemento = 0
     Else
     cantidadElemento = cantidadElemento + 1
      If cantidadElemento > 50 Then
      cantidadElemento = 0
      lote_referencia_id = InsertLoteReferenciaGodoyCruz
      End If
      
    End If
    

        indice_individual = 1
        clasificacion_documental_id = rs!ID_GODOYCRUZ
        
    
    
    sql = " Update elementos"
    sql = sql & vbCrLf & " Set contenedor_id = " & idElementoGodoyCruz("11000" & rs!NRO_CAJA)
   sql = sql & vbCrLf & " Where ID = " & idElementoGodoyCruz(rs!ETIQUETA)
    sql = sql & vbCrLf & " And (contenedor_id Is Null)"
    
    ConGodoyCruz.Execute sql
   
    prefijoCodigoTipoElemento = "'11'"
    ordenRearchivo = ordenRearchivo + 1
    pathArchivoDigital = "'c:\Archivos_Digitales\pdf\" & rs!ETIQUETA & ".pdf'"
    descripcion_rearchivo = "Null"
    referencia_rearchivo_id = "NULL"
    
    
    
     If Not IsNull(rs!Descripcion) And Len(Trim(rs!Descripcion)) > 0 Then
        Descripcion = "'" & Trim(rs!Descripcion) & "'"
    Else
        Descripcion = "NULL"
    End If
    

'
    If Not IsNull(rs!fecha_desde) And IsDate(rs!fecha_desde) Then
        fecha1 = FechaFormato(Trim(rs!fecha_desde))
    Else
        fecha1 = "NULL"
    End If


    If Not IsNull(rs!fecha_hasta) And IsDate(rs!fecha_hasta) Then
        fecha2 = FechaFormato(Trim(rs!fecha_hasta))
    Else
        fecha2 = "NULL"
    End If
    
    
    If Not IsNull(rs!nro_desde) And IsNumeric(rs!nro_desde) Then
        numero1 = rs!nro_desde
    Else
        numero1 = "Null"
    End If
    
    If Not IsNull(rs!nro_hasta) And IsNumeric(rs!nro_hasta) Then
        numero2 = rs!nro_hasta
    Else
        numero2 = "Null"
    End If
    
    
    
    If Not IsNull(rs!LETRA_desde) And Len(Trim(rs!LETRA_desde)) > 0 Then
        texto1 = "'" & Trim(rs!LETRA_desde) & "'"
    Else
        texto1 = "NULL"
    End If
    
    If Not IsNull(rs!LETRA_hasta) And Len(Trim(rs!LETRA_hasta)) > 0 Then
        texto2 = "'" & Trim(rs!LETRA_hasta) & "'"
    Else
        texto2 = "NULL"
    End If




        sql = " INSERT INTO referencia"
        sql = sql & vbCrLf & "("
        sql = sql & vbCrLf & "descripcion"
        sql = sql & vbCrLf & ", fecha1"
        sql = sql & vbCrLf & ", fecha2"
        sql = sql & vbCrLf & ", indice_individual"
        sql = sql & vbCrLf & ", numero1"
        sql = sql & vbCrLf & ", numero2"
        sql = sql & vbCrLf & ", texto1"
        sql = sql & vbCrLf & ", texto2"
        sql = sql & vbCrLf & ", clasificacion_documental_id"
        sql = sql & vbCrLf & ", elemento_id"
        sql = sql & vbCrLf & ", lote_referencia_id"
        sql = sql & vbCrLf & ", descripcion_rearchivo"
        sql = sql & vbCrLf & ", referencia_rearchivo_id"
        sql = sql & vbCrLf & ", prefijoCodigoTipoElemento"
        sql = sql & vbCrLf & ", ordenRearchivo"
        sql = sql & vbCrLf & ", pathLegajo"
        sql = sql & vbCrLf & ",fechaHora"
        sql = sql & vbCrLf & ")"
        sql = sql & vbCrLf & " VALUES ("
        sql = sql & vbCrLf & Descripcion
        sql = sql & vbCrLf & "," & fecha1
        sql = sql & vbCrLf & "," & fecha2
        sql = sql & vbCrLf & "," & indice_individual
        sql = sql & vbCrLf & "," & numero1
        sql = sql & vbCrLf & "," & numero2
        sql = sql & vbCrLf & "," & texto1
        sql = sql & vbCrLf & "," & texto2
        sql = sql & vbCrLf & "," & clasificacion_documental_id
        sql = sql & vbCrLf & "," & idElementoGodoyCruz(rs!ETIQUETA)
        sql = sql & vbCrLf & "," & lote_referencia_id
        sql = sql & vbCrLf & "," & descripcion_rearchivo
        sql = sql & vbCrLf & "," & referencia_rearchivo_id
        sql = sql & vbCrLf & "," & prefijoCodigoTipoElemento
        sql = sql & vbCrLf & "," & ordenRearchivo
        sql = sql & vbCrLf & "," & pathArchivoDigital
        sql = sql & vbCrLf & ",'09/05/2016 15:28:00'"
        sql = sql & vbCrLf & ")"
    
    ConGodoyCruz.Execute sql
    
  
rs.MoveNext
Loop

End Sub

Private Sub Command4_Click()
Dim conAsp150 As New ADODB.Connection
Dim RsAsp150 As New ADODB.Recordset
Dim sql As String
    sql = " SELECT ID_REFERENCIAS "
    sql = sql & "  From REFERENCIAS_FONDO_PARA_BORRAR_20160729"


    conAsp150.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"

    RsAsp150.Open sql, conAsp150

Do While RsAsp150.EOF
    sql = " Delete Top(1) From referencia WHERE id  = " & RsAsp150!ID_REFERENCIAS
    conAsp150.Execute sql
    RsAsp150.MoveNext
    
Loop


End Sub

Private Sub Form_Load()

 strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
    
    conPYL.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
   strConGodoyCruz = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=godoycruz;Data Source=222.15.19.134"
strBasa = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
End Sub

Public Function idElementoGodoyCruz(elemento As String) As Long
 Dim rs As New ADODB.Recordset
 
 Dim sql As String
 
 
 sql = " SELECT id From elementos WHERE codigo = '" & elemento & "'"
 rs.Open sql, strConGodoyCruz
 
 idElementoGodoyCruz = rs!ID


End Function
