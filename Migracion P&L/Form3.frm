VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7260
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11445
   LinkTopic       =   "Form3"
   ScaleHeight     =   7260
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Consulta de Green"
      Height          =   495
      Left            =   8400
      TabIndex        =   11
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   735
      Left            =   8880
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   555
      Left            =   5700
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   615
      Left            =   5220
      TabIndex        =   8
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CONTROL_DIGITALIZACION"
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   540
      Width           =   2595
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copiar Excel"
      Height          =   315
      Left            =   60
      TabIndex        =   6
      Top             =   1020
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "pasar imagens de Green"
      Height          =   495
      Left            =   2100
      TabIndex        =   5
      Top             =   1320
      Width           =   3195
   End
   Begin VB.CommandButton cmdDirectorios 
      Caption         =   "cmdDirectorios"
      Height          =   435
      Left            =   5160
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "green lotes"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   3300
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtElemento 
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Text            =   "120000890288"
      Top             =   60
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4815
      Left            =   180
      TabIndex        =   0
      Top             =   2280
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8493
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDirectorios_Click()

    Dim sql As String
    
    Dim MyName As String
    
    Dim DIRECTORIO(300, 300, 300) As String
    Dim D As Integer
    Dim SD As Integer
    Dim A As Integer
    Dim P As Integer
    Dim PasoFin(9000) As String
    Dim NombreArchivo As String
    
    MyName = Dir("I:\0229-GREEN\DIGITALIZADAS\", vbDirectory)
        Do While MyName <> ""
            If Len(MyName) > 3 Then
            
                DIRECTORIO(D, 0, 0) = MyName
                D = D + 1
            End If
            MyName = Dir()
        Loop
    
   
    For D = 0 To 300
         SD = 1
        If DIRECTORIO(D, 0, 0) <> "" Then
            MyName = Dir("I:\0229-GREEN\DIGITALIZADAS\" & DIRECTORIO(D, 0, 0) & "\", vbDirectory)
            Do While MyName <> ""
            If Len(MyName) > 3 Then
                DIRECTORIO(D, SD, 0) = MyName
                SD = SD + 1
            End If
            MyName = Dir()
            Loop
         End If
    Next
    
  
    
     For D = 0 To 300
        If DIRECTORIO(D, 0, 0) <> "" Then
          For SD = 1 To 300
            A = 2
            MyName = Dir("I:\0229-GREEN\DIGITALIZADAS\" & DIRECTORIO(D, 0, 0) & "\" & DIRECTORIO(D, SD, 0) & "\*.PDF")
            Do While MyName <> ""
            If Len(MyName) > 3 Then
                DIRECTORIO(D, SD, A) = MyName
                A = A + 1
            End If
            MyName = Dir()
            Loop
         Next
      
         End If
    Next
    
    
    P = 0
    For D = 0 To 300
        For SD = 1 To 300
            For A = 2 To 300
                If DIRECTORIO(D, SD, A) <> "" Then
'                    MsgBox DIRECTORIO(D, SD, A)
'                    MsgBox DIRECTORIO(D, SD, 0)
'                    MsgBox DIRECTORIO(D, 0, 0)
                    PasoFin(P) = DIRECTORIO(D, 0, 0) & "\" & DIRECTORIO(D, SD, 0) & "\" & DIRECTORIO(D, SD, A)
                    NombreArchivo = Mid(DIRECTORIO(D, SD, A), 1, 10)
                    
                    FileCopy "I:\0229-GREEN\DIGITALIZADAS\" & PasoFin(P), "I:\0229-GREEN\TERMINADO\" & DIRECTORIO(D, SD, A)
                    
'                    MsgBox PasoFin(P)
                End If
            Next
         Next
    Next
    
    
    
    

End Sub

Private Sub Command1_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset

    Dim strConAsp As String
    Dim strConAsp150 As String
    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"

strConAsp = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=190.151.143.135"
        
        
        sql = " SELECT requerimiento.numero AS [N°REQUERIMINTO], operacion.id AS IDOPERACION, elementos.codigo, elementos.estado, operacion.codigo AS Expr1,"
        sql = sql & " operacion.estado AS Expr2"
        sql = sql & " FROM         requerimiento INNER JOIN"
        sql = sql & " operacion ON requerimiento.id = operacion.requerimiento_id INNER JOIN"
        sql = sql & " x_operacion_elemento ON operacion.id = x_operacion_elemento.operacion_id INNER JOIN"
        sql = sql & " elementos ON x_operacion_elemento.elemento_id = elementos.id"
        sql = sql & " WHERE (elementos.codigo = '" & Trim(txtElemento.Text) & "')"

'SELECT     requerimiento.numero AS [N°REQUERIMINTO], elementos.codigo, elementos.estado, elementos.id, referencia.elemento_id, referencia.numero1, referencia.texto1,
'YEAR(referencia.fecha1) AS Expr3, elementos.clienteEmp_id, elementos.clienteAsp_id, clientesEmp.razonSocial_id, personas_juridicas.razonSocial
'FROM         clientesEmp INNER JOIN
'requerimiento INNER JOIN
'operacion ON requerimiento.id = operacion.requerimiento_id INNER JOIN
'x_operacion_elemento ON operacion.id = x_operacion_elemento.operacion_id INNER JOIN
'elementos ON x_operacion_elemento.elemento_id = elementos.id INNER JOIN
'referencia ON elementos.id = referencia.elemento_id INNER JOIN
'lotereferencia ON referencia.lote_referencia_id = lotereferencia.id ON clientesEmp.id = lotereferencia.cliente_emp_id INNER JOIN
'personas_juridicas ON clientesEmp.razonSocial_id = personas_juridicas.id, empresas
'GROUP BY requerimiento.numero, elementos.codigo, elementos.estado, elementos.id, referencia.elemento_id, referencia.numero1, referencia.texto1, YEAR(referencia.fecha1),
'elementos.clienteEmp_id , elementos.clienteAsp_id, clientesEmp.razonSocial_id, personas_juridicas.razonSocial
'ORDER BY personas_juridicas.razonSocial


rs.CursorLocation = adUseClient
rs.Open sql, strConAsp150
Set DataGrid1.DataSource = rs.DataSource
    DataGrid1.Refresh





End Sub



Private Sub Command2_Click()
    Dim sql As String
    Dim rs As New ADODB.Recordset



Dim strConAsp150 As String


    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"


sql = "  SELECT     CONVERT(char, lotereferencia.fecha_registro, 103) AS fecha, lotereferencia.codigo, lotereferencia.cliente_asp_id, lotereferencia.cliente_emp_id, COUNT(*)"
sql = sql & " AS CANTIDAD, personas_juridicas.razonSocial, lotereferencia.empresa_id, lotereferencia.sucursal_id, sucursales.descripcion, elementos.codigo AS Etiqueta,"
sql = sql & " referencia.pathLegajo"
sql = sql & " FROM         elementos INNER JOIN"
sql = sql & " referencia ON elementos.id = referencia.elemento_id LEFT OUTER JOIN"
sql = sql & " lotereferencia INNER JOIN"
sql = sql & " personas_juridicas INNER JOIN"
sql = sql & " clientesEmp ON personas_juridicas.id = clientesEmp.razonSocial_id ON lotereferencia.cliente_emp_id = clientesEmp.id INNER JOIN"
sql = sql & " sucursales ON lotereferencia.sucursal_id = sucursales.id ON referencia.lote_referencia_id = lotereferencia.id"
sql = sql & " GROUP BY CONVERT(char, lotereferencia.fecha_registro, 103), lotereferencia.codigo, lotereferencia.cliente_asp_id, lotereferencia.cliente_emp_id,"
sql = sql & " personas_juridicas.razonSocial , lotereferencia.empresa_id, lotereferencia.sucursal_id, sucursales.descripcion, elementos.Codigo, referencia.pathLegajo"
sql = sql & " Having (lotereferencia.cliente_emp_id = 20026)"
sql = sql & " ORDER BY lotereferencia.codigo, Etiqueta"



rs.CursorLocation = adUseClient
rs.Open sql, strConAsp150
Set DataGrid1.DataSource = rs.DataSource
    DataGrid1.Refresh


End Sub

Private Sub Command3_Click()

 Dim MyName As String
 Dim sql As String
 Dim codigo As String
 Dim rs As New ADODB.Recordset
 Dim conAsp150 As New ADODB.Connection
 Dim PasoInicio As String
 Dim PasoFinal As String
   Rem MsgBox rs!ID
             PasoInicio = "I:\0229-GREEN\PARA ASP\"
             PasoFinal = "\\222.15.19.150\Archivos_Digitales\PDF\ConsultasDigitales\"
 
 
 
 Dim strConAsp150 As String
        strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
        conAsp150.Open strConAsp150
        MyName = Dir(PasoInicio & "*.PDF")
    Do While MyName <> ""
        If Len(MyName) > 3 Then
            codigo = Mid(MyName, 1, 12)
            sql = " SELECT id, codigo "
            sql = sql & " From elementos"
            sql = sql & " WHERE codigo = '" & codigo & "'"
            Set rs = New ADODB.Recordset
            rs.Open sql, strConAsp150
            If Not rs.EOF Then
                sql = " UPDATE referencia "
                sql = sql & " SET pathLegajo = '" & "C:\Archivos_Digitales\PDF\ConsultasDigitales\" & MyName & "'"
                sql = sql & " Where referencia.elemento_id = " & rs!ID
                sql = sql & " AND pathLegajo Is Null "
                conAsp150.Execute sql
                FileCopy PasoInicio & MyName, PasoFinal & MyName
                FileCopy PasoInicio & MyName, PasoInicio & "Subidas\" & MyName
                Kill PasoInicio & MyName
            Else
            
            End If
        
        End If
        MyName = Dir()
    Loop

End Sub

Private Sub Command4_Click()
    CopiarDatosGrilla DataGrid1
End Sub

Private Sub Command5_Click()

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim strConAsp150 As String
        strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
        sql = " SELECT     CONVERT(char, lotereferencia.fecha_registro, 103) AS FECHA, lotereferencia.codigo AS CODIGOLOTE,"
        sql = sql & "  elementos.codigo AS ETIQUETA, elementos_1.codigo AS CAJA"
        sql = sql & "  FROM  elementos INNER JOIN"
        sql = sql & "  referencia ON elementos.id = referencia.elemento_id INNER JOIN"
        sql = sql & " elementos elementos_1 ON elementos.contenedor_id = elementos_1.id LEFT OUTER JOIN"
        sql = sql & " lotereferencia INNER JOIN"
        sql = sql & " personas_juridicas INNER JOIN"
        sql = sql & " clientesEmp ON personas_juridicas.id = clientesEmp.razonSocial_id ON lotereferencia.cliente_emp_id = clientesEmp.id INNER JOIN"
        sql = sql & " sucursales ON lotereferencia.sucursal_id = sucursales.id ON referencia.lote_referencia_id = lotereferencia.id"
        sql = sql & "  Where (lotereferencia.cliente_emp_id = 20026) And (lotereferencia.Codigo > 4051) And (referencia.pathLegajo Is Null)"
        sql = sql & " ORDER BY CODIGOLOTE ,CAJA, ETIQUETA"
        rs.CursorLocation = adUseClient
        rs.Open sql, strConAsp150
        Set DataGrid1.DataSource = rs.DataSource
        DataGrid1.Refresh








End Sub

Private Sub Command6_Click()


    Dim sql As String
    
    Dim strConAsp150 As String
    Dim conAsp150 As New ADODB.Connection
    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"

    Dim conpyl1 As New ADODB.Connection
    conpyl1.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
    Dim rs As New ADODB.Recordset
Dim RSASP As New ADODB.Recordset
    
   
sql = " SELECT Id, Caja_Asp, ID_ASP "
sql = sql & "  From [P&LCUSTODIA].dbo.caja "
sql = sql & "  Where (Not (CAJA_ASP Is Null)) "
Rem AND (ID_ASP IS NULL) "
rs.CursorLocation = adUseClient
 rs.Open sql, conpyl1, adOpenForwardOnly, adLockOptimistic
Dim clienteEmp_id As Integer
Do While Not rs.EOF
    sql = "SELECT id, codigoBarras, observacion, orden, elemento_id, lectura_id "
    sql = sql & " From lecturaDetalle"
    sql = sql & " WHERE codigoBarras = '" & rs!CAJA_ASP & "'"
    
    sql = " SELECT     id , clienteEmp_id"
sql = sql & "  From elementos"
sql = sql & "  WHERE   codigo= '" & rs!CAJA_ASP & "'"
    
    Set RSASP = New ADODB.Recordset
    RSASP.Open sql, strConAsp150
    If Not RSASP.EOF Then
    
     If IsNull(RSASP!clienteEmp_id) Then
        clienteEmp_id = 0
     Else
        clienteEmp_id = RSASP!clienteEmp_id
     End If
     
        sql = " Update caja "
        sql = sql & " SET ID_CLIENTE_ASP = " & clienteEmp_id
   sql = sql & " , ID_ASP =" & RSASP!ID
        
        sql = sql & " Where ID =" & rs!ID
        conpyl1.Execute sql
    End If
    rs.MoveNext
Loop


End Sub

Private Sub Command7_Click()
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String
Dim i As Integer
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Z:\Tareas\Migracion  de P&L\Administracion\Proyecto\MIGRA.mdb"

rs.Open " SELECT Empresa_1Id , CajaId FROM CAJASCLIENTES", con

Do While Not rs.EOF
   MsgBox rs!Empresa_1Id
    rs.MoveNext
Loop

End Sub

Private Sub Command8_Click()


Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Rs2 As New ADODB.Recordset

    Dim strConAsp As String
    Dim strConAsp150 As String
    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"


sql = " SELECT        clasificacionDocumental_1.codigo AS CODIGO_SUCURSAL, clasificacionDocumental_2.codigo AS CODIGO_DOCUMENTO,"
sql = sql & vbCrLf & "                         clasificacionDocumental_2.nombre AS NOMBRE_DOCUMENTOS, clasificacionDocumental_2.id AS ID_DOCUMENTO"
sql = sql & vbCrLf & " FROM            clasificacionDocumental INNER JOIN"
                         sql = sql & vbCrLf & " clasificacionDocumental AS clasificacionDocumental_1 ON clasificacionDocumental.id = clasificacionDocumental_1.padre_id INNER JOIN"
sql = sql & vbCrLf & "                         clasificacionDocumental AS clasificacionDocumental_2 ON clasificacionDocumental_1.id = clasificacionDocumental_2.padre_id"
sql = sql & vbCrLf & " WHERE        (clasificacionDocumental.cliente_emp_id = 20042) AND (clasificacionDocumental.id IN (71054, 71054, 71055, 81267, 81371, 81374, 81438))"

rs.Open sql, strConAsp150

 Do While Not rs.EOF
    
    sql = " SELECT codigo"
    sql = sql & vbCrLf & "  From clasificacionDocumental"
    sql = sql & vbCrLf & "  Where (cliente_emp_id = 20042) And codigo = " & rs!CODIGO_SUCURSAL & "02"
    Rs2.Open sql, strConAsp150
    If Not rs.EOF Then
    Debug.Print Rs2!codigo
    End If
    rs.MoveNext
 Loop
 



End Sub

Private Sub Command9_Click()

Dim sql  As String
Dim rs As New ADODB.Recordset


Dim strConAsp150 As String
    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"




sql = "SELECT        TOP (1000) ETIQUETA, CAJA, ID_CLIENTE, RAZON_SOCIAL, TIPO_DOCUMENTO, NRO_DESDE, NRO_HASTA, LETRA_DESDE, LETRA_HASTA, DESCRIPCION, FECHA_DESDE, FECHA_HASTA, pathLegajo,"
sql = sql & " NumeroLote , clasificacion_documental_id, codigo, fechaHora"
sql = sql & " From VIEW_REFERENCIAS"
sql = sql & " WHERE        (ID_CLIENTE = 20026) AND (NRO_DESDE IN (139558, 140415, 140441, 141184, 141888, 142231, 142764, 142853, 142986, 143862, 144365, 144366, 144367, 144744, 144802, 145520, 145671, 145673, 145914,"
                         sql = sql & "146174, 146691, 146752, 146902, 147831, 147845, 147859, 147861, 148479, 149226, 149229, 149230, 149766, 150679, 151156, 151891, 152608, 152788, 153242, 153502, 153609, 154311, 154703, 155312,"
                         sql = sql & " 155656, 155769, 156000, 156256, 156848, 157154, 157366, 157556, 157839, 157917, 158524, 158691, 158736, 158737, 159481, 159594, 160095, 160338, 160669, 160723, 161063, 161229, 161662, 161704,"
                         sql = sql & " 162030, 162136, 162932, 163561, 163047, 163538, 163539, 163543, 163551, 163552, 163555, 163556, 163557, 163558, 163559, 163560, 163562, 163563, 163739, 163985, 164594, 164620, 164886, 165109,"
                         sql = sql & " 165487, 165488, 165685, 165972, 166036, 166457, 166574, 166945, 167152, 167405, 167406, 167407, 167442, 167444, 167486, 167604, 167614, 167745, 168311, 169470, 169991, 170305, 170608, 171915,"
                         sql = sql & " 172081, 172847, 172873, 173861, 173881, 174610, 174815, 175065, 175701, 175770, 175779, 175863, 176628, 176634, 176726, 177866, 177867, 178585, 178617, 178709, 178723))"


rs.Open sql, strConAsp150
Do While Not rs.EOF
    FileCopy Replace(rs!pathLegajo, "C:", "\\222.15.19.150"), "C:\Green\" & rs!nro_desde & " Etiqueta " & rs!ETIQUETA & ".PDF"
    


    rs.MoveNext
Loop


End Sub
