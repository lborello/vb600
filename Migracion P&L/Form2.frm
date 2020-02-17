VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form2"
   ScaleHeight     =   7755
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   795
      Left            =   3540
      TabIndex        =   17
      Top             =   6720
      Width           =   1875
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Command11"
      Height          =   615
      Left            =   3720
      TabIndex        =   16
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Command10"
      Height          =   555
      Left            =   6600
      TabIndex        =   15
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   855
      Left            =   7380
      TabIndex        =   14
      Top             =   1140
      Width           =   1035
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   675
      Left            =   7440
      TabIndex        =   13
      Top             =   6180
      Width           =   915
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   675
      Left            =   6960
      TabIndex        =   12
      Top             =   3900
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   735
      Left            =   6540
      TabIndex        =   11
      Top             =   2700
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   735
      Left            =   4200
      TabIndex        =   10
      Top             =   5460
      Width           =   1995
   End
   Begin VB.CommandButton cmdMuniLujan2 
      Caption         =   "muni3"
      Height          =   1155
      Left            =   240
      TabIndex        =   9
      Top             =   5460
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   4500
      TabIndex        =   8
      Top             =   1500
      Width           =   2235
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   420
      TabIndex        =   7
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton cmdActualizacionLegajos 
      Caption         =   "Actualizacion Legajos"
      Height          =   1035
      Left            =   240
      TabIndex        =   4
      Top             =   3300
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   795
      Left            =   4080
      TabIndex        =   3
      Top             =   2580
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4620
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton CMDpLAZA 
      Caption         =   "pLAZA"
      Height          =   915
      Left            =   420
      TabIndex        =   1
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdMigracion 
      Caption         =   "Migracion_Muni "
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   540
      Width           =   2475
   End
   Begin VB.Label lblorden 
      Caption         =   "Label1"
      Height          =   435
      Left            =   5100
      TabIndex        =   6
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label lblcaja 
      Caption         =   "Label1"
      Height          =   435
      Left            =   4920
      TabIndex        =   5
      Top             =   4260
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim strConAsp As String
    Dim strConBasa As String
    Dim strConPyl As String
    
Private Sub cmdActualizacionLegajos_Click()
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
conPYL.Open strConPyl

ConAsp.Open strConAsp



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
    
    rs.Open sql, strConPyl, adOpenStatic, adLockReadOnly
            
            
            
            
            
            
    Rem         rs.Open sql, strConBasa

           Rem conAsp.Open strConAsp
CajaAnterior = 0
conBasa.Open strConBasa
ID_ASP_CAJA = 1122054

Do While Not rs.EOF

    If CajaAnterior <> rs!IDCAJA Then
        lote_referencia_id = InsertLoteReferencia(4)
        ordenRearchivo = 0
        CajaAnterior = rs!IDCAJA
        sql = " Update elementos"
        sql = sql & vbCrLf & " Set cerrado = 1"
        sql = sql & vbCrLf & " Where ID = " & ID_ASP_CAJA
        ConAsp.Execute sql
       
    End If
    
        sql = " Update Documento "
        sql = sql & vbCrLf & "  Set ID_ASP_REFERENCIAS = 0109"
        sql = sql & vbCrLf & "  Where ID = " & rs!IDDOCUMENTO
        conPYL.Execute sql
        lblcaja.Caption = ID_ASP_CAJA
        lblcaja.Refresh
        lblorden.Caption = ordenRearchivo
        lblorden.Refresh
        indice_individual = 1
        clasificacion_documental_id = 20216
        elemento_id = rs!ID_ASP_LEGAJO
    
    
    sql = " Update elementos"
    sql = sql & vbCrLf & " Set contenedor_id = " & ID_ASP_CAJA
    sql = sql & vbCrLf & " Where ID = " & rs!ID_ASP_LEGAJO
    sql = sql & vbCrLf & " And (contenedor_id Is Null)"
    
    ConAsp.Execute sql
   
    prefijoCodigoTipoElemento = "'" & Mid(rs!CODIGO_ASP_ELEMENTO_LEGAJO, 1, 2) & "'"
    ordenRearchivo = ordenRearchivo + 1
    pathArchivoDigital = "Null"
    descripcion_rearchivo = "Null"
    referencia_rearchivo_id = "NULL"
    
    If Not IsNull(rs!DESCRIPCION_ASP) And Len(Trim(rs!DESCRIPCION_ASP)) > 0 Then
        Descripcion = "'" & Trim(rs!DESCRIPCION_ASP) & "'"
    Else
        Descripcion = "NULL"
    End If
    

    If Not IsNull(rs!fecha1) And IsDate(rs!fecha1) Then
        fecha1 = FechaFormato(Trim(rs!fecha1))
    Else
        fecha1 = "NULL"
    End If
    
    
    If Not IsNull(rs!fecha2) And IsDate(rs!fecha2) Then
        fecha2 = FechaFormato(Trim(rs!fecha2))
    Else
        fecha2 = "NULL"
    End If
    
    
    If Not IsNull(rs!numero1) And IsNumeric(rs!numero1) Then
        numero1 = rs!numero1
    Else
        numero1 = "Null"
    End If
    
    
    If Not IsNull(rs!numero2) And IsNumeric(rs!numero2) Then
        numero2 = rs!numero2
    Else
        numero2 = "Null"
    End If
    
    If Not IsNull(rs!texto1) And Len(Trim(rs!texto1)) > 0 Then
        texto1 = "'" & Trim(rs!texto1) & "'"
    Else
        texto1 = "NULL"
    End If
    
    If Not IsNull(rs!texto2) And Len(Trim(rs!texto2)) > 0 Then
        texto2 = "'" & Trim(rs!texto2) & "'"
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
    
    
    rs.MoveNext
Loop



End Sub

Private Sub cmdMuniLujan2_Click()
    Dim sql As String
     Dim rsLegajos As New ADODB.Recordset
    Dim FormatoLegajo As String
    Dim ID_ASP_ELEMENTOS_LEGAJOS As String
     Dim ID_ASP_ELEMENTOS_CAJA As String
     
    Dim rsPYLDocumentos As New ADODB.Recordset
    

    
    Dim conBasa As New ADODB.Connection
    Dim ConPLY As New ADODB.Connection
    
    
    conBasa.Open strConBasa
    
    
    sql = " SELECT     Empresa.Nombre, Caja.Id AS IDCAJA, Caja.Numero, Documento.Id AS IDDOCUMENTO, Documento.NUMERO1, Documento.NUMERO2, Documento.TEXTO1,"
    sql = sql & vbCrLf & " Documento.TEXTO2, CONVERT(char, Documento.FECHA1, 103) AS FECHA1, CONVERT(char, Documento.FECHA2, 103) AS FECHA2, Documento.DESCRIPCION_ASP,"
    sql = sql & vbCrLf & " Documento.CODIGO_ASP_ELEMENTO_LEGAJO , Documento.ID_ASP_LEGAJO ,  Empresa.Id AS IDEMPRESA "
    sql = sql & vbCrLf & " FROM         Documento INNER JOIN"
    sql = sql & vbCrLf & " Caja ON Documento.IdCaja = Caja.Id INNER JOIN"
    sql = sql & vbCrLf & " Empresa ON Caja.IdEmpresa = Empresa.Id"
    sql = sql & vbCrLf & " WHERE     (Empresa.Nombre LIKE N'%MUNI%')"
    sql = sql & vbCrLf & " ORDER BY Documento.IdCaja, IDDOCUMENTO"
    
    rsPYLDocumentos.Open sql, strConPyl, adOpenDynamic, adLockReadOnly
    
sql = "  SELECT     ID_LEGAJO, CODIGO_ASP_ELEMENTO_CAJA, ID_ASP_LEGAJO, ID_ASP_CAJA, ID_ASP_REFERENCIAS, FK_PYL_CLIENTE, FK_PYL_DOCUMENTO"
sql = sql & vbCrLf & " From basasql.dbo.LEGAJOS"
sql = sql & vbCrLf & " WHERE     (ID_LEGAJO BETWEEN 6266260 AND 6368660) AND (COD_CLIENTE = 406) AND (FK_PYL_CLIENTE IS NULL)"
sql = sql & vbCrLf & " ORDER BY ID_LEGAJO"


rsLegajos.Open sql, strConBasa, adOpenDynamic, adLockReadOnly
ConPLY.Open strConPyl

Do While Not rsLegajos.EOF And Not rsPYLDocumentos.EOF
        sql = " Update basasql.dbo.LEGAJOS"
        sql = sql & vbCrLf & " SET FK_PYL_CLIENTE =" & rsPYLDocumentos!IDEmpresa
        sql = sql & vbCrLf & " , FK_PYL_DOCUMENTO =" & rsPYLDocumentos!IDDOCUMENTO
        sql = sql & vbCrLf & " Where ID_LEGAJO =  " & rsLegajos!ID_LEGAJO
        sql = sql & vbCrLf & " And (COD_CLIENTE = 406) And (FK_PYL_CLIENTE Is Null)"
        conBasa.Execute sql
        
        sql = " Update Documento "
        sql = sql & vbCrLf & "  SET ID_ASP_LEGAJO =" & rsLegajos!ID_ASP_LEGAJO
        sql = sql & vbCrLf & " Where ID = " & rsPYLDocumentos!IDDOCUMENTO
        ConPLY.Execute sql
        
        
        rsLegajos.MoveNext
        rsPYLDocumentos.MoveNext
Loop



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


    
    sql = " SELECT     Caja.Id, Caja.Numero, Empresa.Id AS IDEMPRESA, Empresa.Nombre, Documento.NUMERO1, Documento.NUMERO2, Documento.TEXTO1, Documento.TEXTO2,"
    sql = sql & vbCrLf & "  convert(char,Documento.FECHA1,103) as FECHA_DESDE  ,CONVERT(CHAR, Documento.FECHA2, 103) AS FECHA_HASTA , Documento.DESCRIPCION_ASP, DocumentoTipo.Descripcion as DocumentoTipoDesc "
    sql = sql & vbCrLf & " FROM caja INNER JOIN"
    sql = sql & vbCrLf & "  Documento ON Caja.Id = Documento.IdCaja INNER JOIN"
    sql = sql & vbCrLf & "  Empresa ON Caja.IdEmpresa = Empresa.Id INNER JOIN"
    sql = sql & vbCrLf & "  DocumentoTipo ON Documento.IdTipoDocumento = DocumentoTipo.Id"
    sql = sql & vbCrLf & " WHERE     (Empresa.Nombre LIKE N'%plaza%')"
    sql = sql & vbCrLf & " ORDER BY DocumentoTipo.Descripcion"

    rs.Open sql, strConPyl

ConAsp.Open strConAsp



Do While Not rs.EOF

    indice_individual = 0
    clasificacion_documental_id = 30249
    elemento_id = 1121413
    lote_referencia_id = 30086
    prefijoCodigoTipoElemento = "'11'"
    ordenRearchivo = ordenRearchivo + 1
    pathArchivoDigital = "Null"
    descripcion_rearchivo = "Null"
    referencia_rearchivo_id = "NULL"
    
    If Not IsNull(rs!DESCRIPCION_ASP) And Len(Trim(rs!DESCRIPCION_ASP)) > 0 Then
        Descripcion = "'" & Trim(rs!DocumentoTipoDesc) & " : " & Trim(rs!DESCRIPCION_ASP) & "'"
    Else
        Descripcion = "NULL"
    End If
    

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
    
    
    If Not IsNull(rs!numero1) And IsNumeric(rs!numero1) Then
        numero1 = rs!numero1
    Else
        numero1 = "Null"
    End If
    
    
    If Not IsNull(rs!numero2) And IsNumeric(rs!numero2) Then
        numero2 = rs!numero2
    Else
        numero2 = "Null"
    End If
    
    If Not IsNull(rs!texto1) And Trim(rs!texto1) Then
        texto1 = "'" & Trim(rs!texto1) & "'"
    Else
        texto1 = "NULL"
    End If
    
    If Not IsNull(rs!texto2) And Trim(rs!texto2) Then
        texto2 = "'" & Trim(rs!texto2) & "'"
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
    
    
    rs.MoveNext
Loop



End Sub

Private Sub cmdMigracion_Click()
    
    Dim rsLegajos As New ADODB.Recordset
    Dim FormatoLegajo As String
    Dim ID_ASP_ELEMENTOS_LEGAJOS As String
     Dim ID_ASP_ELEMENTOS_CAJA As String
     
    
    
    Dim sql As String
    
    Dim conBasa As New ADODB.Connection
    
    conBasa.Open strConBasa
    
    sql = " SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE, FK_INDICES, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA,  convert( char,FECHA_DESDE, 103) as FECHA_DESDE ,"
    sql = sql & " DESCRIPCION , NRO_CAJA, COD_CLIENTE"
    sql = sql & "  From basasql.dbo.LEGAJOS"
    sql = sql & "  Where (COD_CLIENTE = 128)"
    sql = sql & "  ORDER BY ID_CLIENTE_LEGAJO"
    
    rsLegajos.Open sql, strConBasa
    
    Do While Not rsLegajos.EOF
        If (rsLegajos!ID_CLIENTE_LEGAJO) < 3000 Then
            ID_ASP_ELEMENTOS_LEGAJOS = "14" & Format(rsLegajos!ID_CLIENTE_LEGAJO, "0000000000")
        Else
            ID_ASP_ELEMENTOS_LEGAJOS = "12" & Format(rsLegajos!ID_CLIENTE_LEGAJO, "0000000000")
        End If
        
        
        If (rsLegajos!NRO_CAJA) < 3000 Then
            ID_ASP_ELEMENTOS_CAJA = "13" & Format(rsLegajos!NRO_CAJA, "0000000000")
        Else
            ID_ASP_ELEMENTOS_CAJA = "11" & Format(rsLegajos!NRO_CAJA, "0000000000")
        End If
       
        
        
        
        
    If ID_ASP_ELEMENTOS_LEGAJOS = 0 Then
        MsgBox "no ENCONTRADOS"
    Else
        sql = "Update LEGAJOS"
        sql = sql & "  Set ID_ASP_ELEMENTO_LEGAJO = " & ID_ASP_ELEMENTOS_LEGAJOS
        sql = sql & "  Where ID_LEGAJO = " & rsLegajos!ID_LEGAJO
        conBasa.Execute sql

    End If
    
    
    
    If ID_ASP_ELEMENTOS_CAJA = 0 Then
        MsgBox "no ENCONTRADOS"
    Else
        sql = "Update LEGAJOS"
        sql = sql & "  Set ID_ASP_ELEMENTO_CAJA = " & ID_ASP_ELEMENTOS_CAJA
        sql = sql & "  Where ID_LEGAJO = " & rsLegajos!ID_LEGAJO
        conBasa.Execute sql

    End If
    rsLegajos.MoveNext
    Loop
    
    
    
End Sub



Private Sub CMDpLAZA_Click()


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


    
    sql = " SELECT     Caja.Id, Caja.Numero, Empresa.Id AS IDEMPRESA, Empresa.Nombre, Documento.NUMERO1, Documento.NUMERO2, Documento.TEXTO1, Documento.TEXTO2,"
    sql = sql & vbCrLf & "  convert(char,Documento.FECHA1,103) as FECHA_DESDE  ,CONVERT(CHAR, Documento.FECHA2, 103) AS FECHA_HASTA , Documento.DESCRIPCION_ASP, DocumentoTipo.Descripcion as DocumentoTipoDesc "
    sql = sql & vbCrLf & " FROM caja INNER JOIN"
    sql = sql & vbCrLf & "  Documento ON Caja.Id = Documento.IdCaja INNER JOIN"
    sql = sql & vbCrLf & "  Empresa ON Caja.IdEmpresa = Empresa.Id INNER JOIN"
    sql = sql & vbCrLf & "  DocumentoTipo ON Documento.IdTipoDocumento = DocumentoTipo.Id"
    sql = sql & vbCrLf & " WHERE     (Empresa.Nombre LIKE N'%plaza%')"
    sql = sql & vbCrLf & " ORDER BY DocumentoTipo.Descripcion"

    rs.Open sql, strConPyl

ConAsp.Open strConAsp



Do While Not rs.EOF

    indice_individual = 0
    clasificacion_documental_id = 30249
    elemento_id = 1121413
    lote_referencia_id = 30086
    prefijoCodigoTipoElemento = "'11'"
    ordenRearchivo = ordenRearchivo + 1
    pathArchivoDigital = "Null"
    descripcion_rearchivo = "Null"
    referencia_rearchivo_id = "NULL"
    
    If Not IsNull(rs!DESCRIPCION_ASP) And Len(Trim(rs!DESCRIPCION_ASP)) > 0 Then
        Descripcion = "'" & Trim(rs!DocumentoTipoDesc) & " : " & Trim(rs!DESCRIPCION_ASP) & "'"
    Else
        Descripcion = "NULL"
    End If
    

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
    
    
    If Not IsNull(rs!numero1) And IsNumeric(rs!numero1) Then
        numero1 = rs!numero1
    Else
        numero1 = "Null"
    End If
    
    
    If Not IsNull(rs!numero2) And IsNumeric(rs!numero2) Then
        numero2 = rs!numero2
    Else
        numero2 = "Null"
    End If
    
    If Not IsNull(rs!texto1) And Trim(rs!texto1) Then
        texto1 = "'" & Trim(rs!texto1) & "'"
    Else
        texto1 = "NULL"
    End If
    
    If Not IsNull(rs!texto2) And Trim(rs!texto2) Then
        texto2 = "'" & Trim(rs!texto2) & "'"
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
    
    
    rs.MoveNext
Loop




End Sub


Private Sub Command10_Click()
Dim MyName As String
Dim Nombre_Archivo As String
Dim Datos As String
Dim sql As String
Dim ConPYLS As New ADODB.Connection

MyName = Dir("\\222.15.19.150\Archivos_Procesos\reposicion\lecturas\error\*.txt")
ConPYLS.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"

Do While MyName <> ""   ' Start the loop.
Nombre_Archivo = MyName
        
  Open "\\222.15.19.150\Archivos_Procesos\reposicion\lecturas\error\" & Nombre_Archivo For Input As #1



Dim Linea As String, Total As String
Do Until EOF(1)
    Line Input #1, Linea
        
        If Trim(Linea) <> "" And Len(Trim(Linea)) < 15 Then
        
        sql = " Insert Into CONTROL_LECTURA_ERROR "
        sql = sql & "(  "
       sql = sql & " CAJAS"
       sql = sql & " , ARCHIVO"
       sql = sql & ")"
        sql = sql & " VALUES ("
        sql = sql & Linea
       sql = sql & " ,'" & Nombre_Archivo & "'"
       sql = sql & ")"
        ConPYLS.Execute sql
        End If
    Loop
Close #1
    
        
        MyName = Dir()   ' Get next entry.
Loop



End Sub

Private Sub Command11_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim strConAsp150 As String
strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"

    
    
    
    
    sql = "SELECT     CODIGO, ARCHIVO "
    sql = sql & " From CONTROL_LECTURAS"
    sql = sql & " Where (CODIGO > 120001145570)"
    sql = sql & " ORDER BY CODIGO DESC"
    rs.Open sql, strConAsp150

Do While Not rs.EOF
    
    
    If Dir("C:\Lectura\error\" & rs!ARCHIVO) <> "" Then
        FileCopy "C:\Lectura\error\" & rs!ARCHIVO, "C:\Lectura\txt\" & rs!codigo & DigitoEAN13(rs!codigo) & ".txt"
  
    
    End If
    
    
    
    rs.MoveNext
Loop






End Sub
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

Private Sub Command12_Click()
Dim MyName As String
Dim Nombre_Archivo As String
Dim Datos As String
Dim sql As String
Dim strConAsp150 As String
strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"

Dim conAsp150 As New ADODB.Connection


conAsp150.Open strConAsp150

MyName = Dir("\\222.15.19.150\Archivos_Procesos\reposicion\lecturas\error\*.txt")



Do While MyName <> ""   ' Start the loop.
Nombre_Archivo = MyName
        
  Open "\\222.15.19.150\Archivos_Procesos\reposicion\lecturas\error\" & Nombre_Archivo For Input As #1



Dim Linea As String, Total As String
Do Until EOF(1)
    Line Input #1, Linea
       If Trim(Linea) <> "" And Len(Trim(Linea)) < 20 Then
            sql = " INSERT INTO CONTROL_LECTURAS "
            sql = sql & " ( CODIGO "
            sql = sql & " , ARCHIVO )"
            sql = sql & " Values"
            sql = sql & " ('" & Linea & "'"
            sql = sql & ",'" & Nombre_Archivo & "')"
            conAsp150.Execute sql
        End If
    Loop
Close #1
    
        
        MyName = Dir()   ' Get next entry.
Loop



End Sub

Private Sub Command2_Click()
Dim rs As New ADODB.Recordset

Dim RsAspElemento As New ADODB.Recordset
Dim sql As String
Dim conBasa As New ADODB.Connection

conBasa.Open strConBasa
sql = " SELECT ID_LEGAJO, CODIGO_ASP_ELEMENTO_LEGAJO, "
sql = sql & " CODIGO_ASP_ELEMENTO_CAJA, ID_ASP_LEGAJO, ID_ASP_CAJA"
sql = sql & " From LEGAJOS"
sql = sql & " Where  (COD_CLIENTE = 128) AND (ID_ASP_LEGAJO IS NULL) "
sql = sql & " ORDER BY ID_LEGAJO"
rs.Open sql, strConBasa

Do While Not rs.EOF
    sql = " SELECT id, codigo, clienteAsp_id, clienteEmp_id, contenedor_id"
    sql = sql & "  From elementos"
    sql = sql & "  WHERE clienteAsp_id = 1 "
    sql = sql & "  AND codigo = '" & rs!CODIGO_ASP_ELEMENTO_LEGAJO & "'"
    
    Set RsAspElemento = New ADODB.Recordset
    
    RsAspElemento.Open sql, strConAsp
    
    If Not RsAspElemento.EOF Then
    
    sql = " Update LEGAJOS"
    sql = sql & "  Set ID_ASP_LEGAJO = " & RsAspElemento!ID
    sql = sql & "  Where ID_LEGAJO = " & rs!ID_LEGAJO
    conBasa.Execute sql
    
    End If
    
    sql = " SELECT id, codigo, clienteAsp_id, clienteEmp_id, contenedor_id"
    sql = sql & "  From elementos"
    sql = sql & "  WHERE clienteAsp_id = 1 "
    sql = sql & "  AND codigo = '" & rs!CODIGO_ASP_ELEMENTO_CAJA & "'"
    
    Set RsAspElemento = New ADODB.Recordset
    
    RsAspElemento.Open sql, strConAsp
    
    If Not RsAspElemento.EOF Then
    
    sql = " Update LEGAJOS"
    sql = sql & "  Set ID_ASP_CAJA = " & RsAspElemento!ID
    sql = sql & "  Where ID_LEGAJO = " & rs!ID_LEGAJO
    conBasa.Execute sql
    
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
Dim ConAsp As New ADODB.Connection


    
    sql = " SELECT     Caja.Id, Caja.Numero, Empresa.Id AS IDEMPRESA, Empresa.Nombre, Documento.NUMERO1, Documento.NUMERO2, Documento.TEXTO1, Documento.TEXTO2,"
    sql = sql & vbCrLf & "  convert(char,Documento.FECHA1,103) as FECHA_DESDE  ,CONVERT(CHAR, Documento.FECHA2, 103) AS FECHA_HASTA , Documento.DESCRIPCION_ASP, DocumentoTipo.Descripcion as DocumentoTipoDesc "
    sql = sql & vbCrLf & " FROM caja INNER JOIN"
    sql = sql & vbCrLf & "  Documento ON Caja.Id = Documento.IdCaja INNER JOIN"
    sql = sql & vbCrLf & "  Empresa ON Caja.IdEmpresa = Empresa.Id INNER JOIN"
    sql = sql & vbCrLf & "  DocumentoTipo ON Documento.IdTipoDocumento = DocumentoTipo.Id"
    sql = sql & vbCrLf & " WHERE     (Empresa.Nombre LIKE N'%plaza%')"
    sql = sql & vbCrLf & " ORDER BY DocumentoTipo.Descripcion"



    
    
    
    rs.Open sql, strConPyl

ConAsp.Open strConAsp



Do While Not rs.EOF

    indice_individual = 0
    clasificacion_documental_id = 30249
    elemento_id = 1121413
    lote_referencia_id = 30086
    prefijoCodigoTipoElemento = "'11'"
    ordenRearchivo = ordenRearchivo + 1
    pathArchivoDigital = "Null"
    descripcion_rearchivo = "Null"
    referencia_rearchivo_id = "NULL"
    
    If Not IsNull(rs!DESCRIPCION_ASP) And Len(Trim(rs!DESCRIPCION_ASP)) > 0 Then
        Descripcion = "'" & Trim(rs!DocumentoTipoDesc) & " : " & Trim(rs!DESCRIPCION_ASP) & "'"
    Else
        Descripcion = "NULL"
    End If
    

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
    
    
    If Not IsNull(rs!numero1) And IsNumeric(rs!numero1) Then
        numero1 = rs!numero1
    Else
        numero1 = "Null"
    End If
    
    
    If Not IsNull(rs!numero2) And IsNumeric(rs!numero2) Then
        numero2 = rs!numero2
    Else
        numero2 = "Null"
    End If
    
    If Not IsNull(rs!texto1) And Trim(rs!texto1) Then
        texto1 = "'" & Trim(rs!texto1) & "'"
    Else
        texto1 = "NULL"
    End If
    
    If Not IsNull(rs!texto2) And Trim(rs!texto2) Then
        texto2 = "'" & Trim(rs!texto2) & "'"
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
    
    
    rs.MoveNext
Loop




End Sub



Private Sub Command4_Click()


Dim ConID As New ADODB.Connection
ConID.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Elementos_asp.mdb;Persist Security Info=False"
Dim conBasa As New ADODB.Connection
Dim ID_BASA_LEGAJOS As Long
Dim sql As String
conBasa.Open strConBasa

Dim rs As New ADODB.Recordset

sql = "SELECT id, codigo"
sql = sql & " From ELEMENTOS_ASP"
sql = sql & " ORDER BY codigo ;"


rs.Open sql, ConID


Do While Not rs.EOF
        ID_BASA_LEGAJOS = Mid(rs!codigo, 3)
        
    sql = " Update LEGAJOS"
sql = sql & " Set ID_ASP_LEGAJO = " & rs!ID
sql = sql & " Where ID_LEGAJO = " & ID_BASA_LEGAJOS

conBasa.Execute sql


    rs.MoveNext
Loop




End Sub


Private Sub Command5_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim ConAsp As ADODB.Connection


sql = " SELECT      ID_ASP_LEGAJO"
sql = sql & " From LEGAJOS"
sql = sql = sql & "  (COD_CLIENTE = 128)"
sql = sql & "  And (COD_ESTADO <> 2)"

rs.Open sql, strConBasa

Do While Not rs.EOF
    sql = " Update elementos"
    sql = sql & " SET estado ='En el Cliente'"
    sql = sql & " WHERE elementos.id =  " & rs!ID_ASP_LEGAJO
    sql = sql & " AND (elementos.estado = 'Creado')"
    ConAsp.Execute sql
    rs.MoveNext
Loop



End Sub

Private Sub Command6_Click()
Dim MyName As String
Dim Nombre_Archivo As String
Dim Datos As String
Dim sql As String
Dim ConPYLS As New ADODB.Connection

MyName = Dir("X:\Lectura_\Lecturas Diarias\ASP\*.txt")
ConPYLS.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"

Do While MyName <> ""   ' Start the loop.
Nombre_Archivo = MyName
        
  Open "X:\Lectura_\Lecturas Diarias\ASP\" & Nombre_Archivo For Input As #1



Dim Linea As String, Total As String
Do Until EOF(1)
    Line Input #1, Linea
        
        If Trim(Linea) <> "" Then
        
        sql = " Insert Into CONTROL_LECTURA "
        sql = sql & "(  "
       sql = sql & " CAJAS"
       sql = sql & " , ARCHIVO"
       sql = sql & ")"
        sql = sql & " VALUES ("
        sql = sql & Linea
       sql = sql & " ,'" & Nombre_Archivo & "'"
       sql = sql & ")"
        ConPYLS.Execute sql
        End If
    Loop
Close #1
    
        
        MyName = Dir()   ' Get next entry.
Loop





End Sub

Private Sub Command7_Click()
Dim MyName As String
Dim Nombre_Archivo As String
Dim Datos As String
Dim sql As String
Dim ConPYLS As New ADODB.Connection
Dim ConASPI As New ADODB.Connection
Dim rs As New ADODB.Recordset

ConPYLS.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
ConASPI.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"

sql = " SELECT     CAJAS"
sql = sql & "  From [P&LCUSTODIA].dbo.CONTROL_LECTURA"
sql = sql & "  GROUP BY CAJAS"
sql = sql & "  Having (CAJAS < 120001125424)"
sql = sql & "  ORDER BY CAJAS"

rs.Open sql, ConPYLS

Do While Not rs.EOF
    sql = "Update elementos"
    sql = sql & " Set depositoActual_id = 4 "
    sql = sql & " WHERE (posicion_id IS NULL) and  elementos.codigo ='" & rs!cajas & "'"
    ConASPI.Execute sql
    rs.MoveNext
Loop


'


End Sub

Private Sub Command8_Click()
Dim MyName As String
Dim Nombre_Archivo As String
Dim Datos As String
Dim sql As String
Dim ConPYLS As New ADODB.Connection
Dim ConASPI As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim caja As Double
Dim i As Integer

ConPYLS.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
ConASPI.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"

For caja = 110001145555# To 110001149504#
    sql = "Update elementos"
    sql = sql & " Set depositoActual_id = 4 "
    sql = sql & " WHERE (posicion_id IS NULL) and  elementos.codigo ='" & caja & "'"
    ConASPI.Execute sql, i
    If i = 0 Then
    Debug.Print caja
        End If
    
Next

End Sub

Private Sub Command9_Click()
    Dim strPyl As String
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim conPYL As New ADODB.Connection
        strPyl = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
        sql = " SELECT bloque From BLOQUE_ERROR"
        conPYL.Open strPyl
        rs.Open sql, strPyl
        Do While Not rs.EOF
            sql = " Update [P&LCUSTODIA].dbo.CONTROL_LECTURA"
            sql = sql & " Set borrar ='" & Trim(rs!bloque) & "'"
            sql = sql & " WHERE     (ARCHIVO LIKE '%" & Trim(rs!bloque) & "%')"
            conPYL.Execute sql
            rs.MoveNext
        Loop
    

End Sub

Private Sub Form_Load()
    strConPyl = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
    strConAsp = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=190.151.143.135"
    strConBasa = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"

End Sub

Public Function ID_ASP_ELEMENTOS(elemento As String) As Long
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = " SELECT     id, codigo "
    sql = sql & " From elementos"
    sql = sql & " WHERE  codigo = '" & elemento & "'"

rs.Open sql, strConAsp

If Not rs.EOF Then
    ID_ASP_ELEMENTOS = rs!ID
Else
    ID_ASP_ELEMENTOS = 0
End If


End Function

Public Function InsertLoteReferencia(StrCodigo As String) As Long
    Dim sql As String
    Dim RsMax As New ADODB.Recordset
    Dim MaxCodigo As Long
    
 Dim fecha_registro As String
 Dim cliente_asp_id As String
 Dim cliente_emp_id As String
 Dim empresa_id As String
 Dim sucursal_id As String
 Dim habilitado As String
 Dim codigo As String
 Dim cargaPorRango As String
 Dim Conlotereferencia As New ADODB.Connection
    
    
    sql = " SELECT MAX(codigo) AS MaxCodigo"
    sql = sql & " From lotereferencia "
    
Set RsMax = New ADODB.Recordset
RsMax.Open sql, strConAsp
codigo = RsMax!MaxCodigo + 1



fecha_registro = FechaFormato(Now)
cliente_asp_id = "1"
 cliente_emp_id = "20006"
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
 
 Conlotereferencia.Open strConAsp
 Conlotereferencia.Execute sql
 
 
  
    
sql = " SELECT  MAX(id) AS MaxID "
sql = sql & "  From lotereferencia "
sql = sql & "  Where codigo = " & codigo
    
Set RsMax = New ADODB.Recordset
RsMax.Open sql, strConAsp
InsertLoteReferencia = RsMax!MaxID
 

End Function
