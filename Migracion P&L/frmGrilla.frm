VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmGrilla 
   Caption         =   "Grilla"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11790
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   11790
   Begin VB.CommandButton Command8 
      Caption         =   "CPOPIAR EXCEL"
      Height          =   375
      Left            =   9840
      TabIndex        =   15
      Top             =   1560
      Width           =   1635
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   435
      Left            =   7860
      TabIndex        =   14
      Top             =   1500
      Width           =   1515
   End
   Begin VB.CommandButton Command6 
      Caption         =   "poner estanteria"
      Height          =   495
      Left            =   5220
      TabIndex        =   13
      Top             =   1140
      Width           =   1995
   End
   Begin VB.CommandButton Command5 
      Caption         =   "LLena planilla General"
      Height          =   555
      Left            =   2640
      TabIndex        =   12
      Top             =   1140
      Width           =   2115
   End
   Begin VB.CommandButton cmdActualizarCarga 
      Caption         =   "Actualizar Carga"
      Height          =   615
      Left            =   9960
      TabIndex        =   11
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox txtFechaCargaASP 
      Height          =   315
      Left            =   8040
      TabIndex        =   9
      Top             =   960
      Width           =   1635
   End
   Begin VB.CommandButton cmdBuscarCajaAsp 
      Caption         =   "Buscar ASP"
      Height          =   315
      Left            =   8580
      TabIndex        =   7
      Top             =   180
      Width           =   1935
   End
   Begin VB.TextBox txtCajaASP 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   180
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "cOMPLETAR PLANILLA"
      Height          =   435
      Left            =   240
      TabIndex        =   5
      Top             =   1260
      Width           =   2235
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1755
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   675
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   1680
      TabIndex        =   2
      Top             =   180
      Width           =   1155
   End
   Begin VB.TextBox txtCajaPYL 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "0"
      Top             =   180
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid grdCajas 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   18615
      _ExtentX        =   32835
      _ExtentY        =   10821
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
   Begin VB.Label Label1 
      Caption         =   "Fecha Carga ASP"
      Height          =   195
      Left            =   8100
      TabIndex        =   10
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label lblID 
      Height          =   315
      Left            =   6060
      TabIndex        =   8
      Top             =   660
      Width           =   1695
   End
End
Attribute VB_Name = "frmGrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conPYL As New ADODB.Connection
Dim conBasa As New ADODB.Connection
Dim ValorAnteVideo  As Long

Private Sub cmdActualizarCarga_Click()
    Dim sql As String
        If lblID.Caption <> "" Then
           sql = " Update Caja "
           sql = sql & " Set CARGADA_ASP = '" & txtFechaCargaASP.Text & "'"
           sql = sql & " Where ID = " & lblID.Caption
           conPYL.Execute sql
           MsgBox "OK"
         Else
            MsgBox "NO se registro"
        End If
 
End Sub

Private Sub cmdBuscarCajaAsp_Click()
Dim sql As String
Dim rs As New ADODB.Recordset
    



sql = "SELECT  Caja.Id, Empresa.Nombre, Caja.Numero, Caja.Fecha, Caja.Pasillo, Caja.Estante, Caja.Modulo, Caja.Ubicacion, Caja.Caja_Asp, Caja.VIDEO, Caja.TIEMPO_NUMERO,"
 sql = sql & " Caja.CARGADA_ASP"
sql = sql & " FROM Caja LEFT OUTER JOIN"
sql = sql & " Empresa ON Caja.IdEmpresa = Empresa.Id"
sql = sql & " WHERE (Caja.Caja_Asp LIKE N'%" & Trim(txtCajaASP.Text) & "')"
sql = sql & " ORDER BY Caja.IdEmpresa"



rs.CursorLocation = adUseClient
      
       
rs.Open sql, conPYL

lblID.Caption = ""
Set grdCajas.DataSource = rs.DataSource
       grdCajas.Refresh

End Sub

Private Sub Command1_Click()

Dim sql As String
Dim rs As New ADODB.Recordset
    

sql = " SELECT     Caja.Id, Empresa.Nombre, Caja.Numero, Caja.Fecha, Caja.Pasillo, Caja.Estante, Caja.Modulo, Caja.Ubicacion, Caja.Caja_Asp, Caja.VIDEO, Caja.TIEMPO_NUMERO,CARGADA_ASP"
sql = sql & " FROM         Caja INNER JOIN"
sql = sql & "  Empresa ON Caja.IdEmpresa = Empresa.Id"
sql = sql & "  WHERE      (Caja.Numero LIKE '%" & Trim(txtCajaPYL.Text) & "%')"
sql = sql & " ORDER BY IdEmpresa"
rs.CursorLocation = adUseClient
      
       
rs.Open sql, conPYL

lblID.Caption = ""
Set grdCajas.DataSource = rs.DataSource
       grdCajas.Refresh

 

 
 
End Sub

Private Sub Command2_Click()
Dim i As Integer
Dim sql As String
Dim rs As New ADODB.Recordset
Dim numero1 As Long
Dim numero2 As Long
Dim texto1 As String
Dim texto2 As String
Dim fecha1 As String
Dim fecha2 As String
Dim DESCRIPCION_ASP As String

Dim conPYL2 As New ADODB.Connection
conPYL2.ConnectionString = conPYL.ConnectionString
conPYL2.Open

numero1 = 0
numero2 = 0
texto1 = ""
texto2 = ""
fecha1 = ""
fecha2 = ""



sql = " SELECT Documento.Comentario, Documento.ID , Documento.Orden , Documento.Numero , Documento.Letra, Documento.Anio, Documento.Dni, Documento.Nombre, Documento.Fojas, Documento.Descripcion, Documento.IdCaja, Documento.IdTipoDocumento,"
sql = sql & " Documento.IdUbicacion, Documento.Comentario, Documento.NUMERO1, Documento.NUMERO2, Documento.TEXTO1, Documento.TEXTO2, Documento.FECHA1,"
sql = sql & " Documento.FECHA2, Documento.DESCRIPCION_ASP, Caja.Numero AS CAJANRO"
sql = sql & " FROM  Documento LEFT OUTER JOIN"
sql = sql & " Caja ON Documento.IdCaja = Caja.Id"
sql = sql & " ORDER BY Caja.Id"


Set rs = New ADODB.Recordset

rs.Open sql, conPYL2, adOpenStatic, adLockReadOnly


Do While Not rs.EOF

numero1 = 0
numero2 = 0
texto1 = ""
texto2 = ""
fecha1 = ""
fecha2 = ""
DESCRIPCION_ASP = ""

If Trim(rs!numero) = Trim(rs!Anio) And Len(Trim(rs!Anio)) = 4 Then
    If IsNumeric(rs!Anio) And Len(Trim(rs!Anio)) = 4 Then
        fecha1 = "01/01/" & rs!Anio
        fecha2 = "31/12/" & rs!Anio
    End If
Else
    If IsNumeric(rs!numero) Then
        numero1 = rs!numero
        numero2 = rs!numero
    Else
        If IsNumeric(rs!Dni) Then
            numero1 = rs!Dni
            numero2 = rs!Dni
        End If
    End If
    If IsNumeric(rs!Anio) And Len(Trim(rs!Anio)) = 4 Then
        fecha1 = "01/01/" & rs!Anio
        fecha2 = "31/12/" & rs!Anio
    End If


End If

    If Len(Trim(rs!LETRA)) = 1 Then
       texto1 = rs!LETRA
       texto2 = rs!LETRA
    End If
    If Not IsNumeric(rs!numero) Then
        DESCRIPCION_ASP = " NRO: " & Trim(rs!numero)
    End If
    If Len(Trim(rs!LETRA)) <> 1 And Not IsNull(rs!LETRA) And Trim(rs!LETRA) <> "" Then
         DESCRIPCION_ASP = DESCRIPCION_ASP & " Letra: " & Trim(rs!LETRA)
    End If
    If Not IsNull(rs!Orden) And Trim(rs!Orden) <> "" Then
         DESCRIPCION_ASP = DESCRIPCION_ASP & " Orden: " & Trim(rs!Orden)
    End If
    If Not IsNumeric(rs!Anio) And Not IsNull(rs!Anio) And Trim(rs!Anio) <> "" Then
         DESCRIPCION_ASP = DESCRIPCION_ASP & " Año: " & Trim(rs!Anio)
    End If
    If Not IsNull(rs!Dni) And Not IsNull(rs!Dni) And Trim(rs!Dni) <> "" Then
     DESCRIPCION_ASP = DESCRIPCION_ASP & " Dni: " & Trim(rs!Dni)
    End If
    
    If Not IsNull(rs!Nombre) And Not IsNull(rs!Nombre) And Trim(rs!Nombre) <> "" Then
     DESCRIPCION_ASP = DESCRIPCION_ASP & " Nombre: " & Trim(rs!Nombre)
    End If
    
    If Not IsNull(rs!Fojas) And Trim(rs!Fojas) <> "" Then
     DESCRIPCION_ASP = DESCRIPCION_ASP & " Fojas: " & Trim(rs!Fojas)
    End If
    
    If Not IsNull(rs!Descripcion) And Trim(rs!Descripcion) <> "" Then
     DESCRIPCION_ASP = DESCRIPCION_ASP & " Descripcion: " & Trim(rs!Descripcion)
    End If
    If Not IsNull(rs!Descripcion) And Trim(rs!Descripcion) <> "" Then
     DESCRIPCION_ASP = DESCRIPCION_ASP & " Descripcion: " & Trim(rs!Descripcion)
    End If


    
    
    If Not IsNull(rs!Comentario) And Trim(rs!Comentario) <> 0 Then
        DESCRIPCION_ASP = DESCRIPCION_ASP & " Comentario: " & Trim(rs!Comentario)
    End If

    If Not IsNull(rs!CAJANRO) Then
        DESCRIPCION_ASP = DESCRIPCION_ASP & " CAJA: " & Trim(rs!CAJANRO)
    End If
 
      sql = " Update [P&LCUSTODIA].dbo.Documento"
      sql = sql & vbCrLf & " SET "

If numero1 <> 0 Then
    sql = sql & vbCrLf & "   NUMERO1 = " & numero1
Else
 sql = sql & vbCrLf & "   NUMERO1 = " & numero1
End If
If numero2 <> 0 Then
    sql = sql & vbCrLf & " ,  NUMERO2 = " & numero2
End If
If texto1 <> "" Then
    sql = sql & vbCrLf & " , TEXTO1 = '" & texto1 & "'"
End If

If texto2 <> "" Then
    sql = sql & vbCrLf & " , TEXTO2 = '" & texto2 & "'"
End If
If fecha1 <> "" Then
    sql = sql & vbCrLf & " , FECHA1 = '" & fecha1 & "'"
End If
If fecha2 <> "" Then
    sql = sql & vbCrLf & " , FECHA2 = '" & fecha2 & "'"
End If

If DESCRIPCION_ASP <> "" Then
 sql = sql & vbCrLf & " , DESCRIPCION_ASP = '" & Trim(Replace(DESCRIPCION_ASP, "'", "´")) & "'"
End If


sql = sql & vbCrLf & "  Where ID = " & rs!ID

 ejecutar sql
 

    rs.MoveNext
Loop


End Sub

Private Sub Command3_Click()
Dim sql As String
Dim rs As New ADODB.Recordset

sql = " SELECT  [Id]"
sql = sql & "      ,[SinReferencias]"
sql = sql & "      ,[IdCaja]"
sql = sql & "   From [P&LCUSTODIA].[dbo].[CAJAS_CON_ REFERENCIAS]"

rs.Open sql, conPYL

Do While Not rs.EOF
    sql = " Update [P&LCUSTODIA].dbo.Caja"
sql = sql & "  SET              CONREF ='CON REFERENCIA'"
sql = sql & "  Where ID =" & rs!ID
conPYL.Execute sql

    rs.MoveNext
    
Loop





End Sub

Private Sub Command4_Click()

        Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
        Dim P As Integer
        Dim rs As ADODB.Recordset
        Dim sql As String
        Dim cliente As Integer
        
Rem 32
        'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open("Z:\Tareas\Migracion  de P&L\Pedidos\buscar.xlsx")
        Set hojaEx = libroEx.Worksheets.Item(1)
       
       
       cliente = InputBox("Ingrese la id empresa")
       
        hojaEx.Cells(1, 4) = " Id"
        hojaEx.Cells(1, 5) = " Numero"
        hojaEx.Cells(1, 6) = " Pasillo"
        hojaEx.Cells(1, 7) = " Estante"
        hojaEx.Cells(1, 8) = " Modulo"
        hojaEx.Cells(1, 9) = " Ubicacion"
        hojaEx.Cells(1, 10) = " CAJA_ASP"
        hojaEx.Cells(1, 11) = " Video"
        
        
       
       For P = 1 To 100
            
            
            If hojaEx.Cells(P, 3) <> "" Then
                    Set rs = New ADODB.Recordset
                    rs.CursorLocation = adUseClient
                    sql = " SELECT     Id, Numero, Pasillo,"
                    sql = sql & " Estante , Modulo, Ubicacion, CAJA_ASP, Video"
                    sql = sql & " From [P&LCUSTODIA].dbo.Caja"
                    sql = sql & " WHERE     (Numero LIKE '%" & hojaEx.Cells(P, 3) & "') "
                    sql = sql & " AND IdEmpresa = " & cliente
                    rs.Open sql, conPYL
                    
                    If rs.RecordCount = 1 Then
                    

                        
                        hojaEx.Cells(P, 4) = rs!ID
                        hojaEx.Cells(P, 5) = rs!numero
                        hojaEx.Cells(P, 6) = rs!Pasillo
                        hojaEx.Cells(P, 7) = rs!estante
                        hojaEx.Cells(P, 8) = rs!Modulo
                        hojaEx.Cells(P, 9) = rs!Ubicacion
                        hojaEx.Cells(P, 10) = rs!CAJA_ASP
                        hojaEx.Cells(P, 11) = rs!Video
                    
                    Else
                    
                    hojaEx.Cells(P, 4) = "restros:" & rs.RecordCount
                    
                    End If
            
            End If
       Next
       
libroEx.Save
           libroEx.Close
           ApExcel.Quit
           Set ApExcel = Nothing
           Set libroEx = Nothing
End Sub

Private Sub Command5_Click()


        Dim ApExcel As Excel.Application
        Dim libroEx As Excel.Workbook
        Dim hojaEx As Excel.Worksheet
        Dim P As Integer
        Dim rs As ADODB.Recordset
        Dim sql As String
        Dim cliente As Integer
        
        Dim C_Caja_PYL As Integer
        Dim C_ID_Caja As Integer
        Dim C_CAJA_ASP As Integer
        Dim C_Nombre As Integer
        Dim C_Video As Integer
        Dim C_ESTANTERIA As Integer
        
        
        C_Caja_PYL = 8
        C_ID_Caja = 9
        C_CAJA_ASP = 10
        C_Nombre = 11
        C_Video = 12
        C_ESTANTERIA = 15
        
        
Rem 32
        'abrir hoja excel
        Set ApExcel = New Excel.Application
        Set libroEx = Excel.Workbooks.Open("Z:\Tareas\Migracion  de P&L\Planta\pedidos\Planilla General.xlsx")
        Set hojaEx = libroEx.Worksheets.Item(1)
        hojaEx.Unprotect 21877471
       
       For P = 2 To 2000
            
            
            If hojaEx.Cells(P, C_Caja_PYL) <> "" Then
                    Set rs = New ADODB.Recordset
                    rs.CursorLocation = adUseClient
                    sql = " SELECT     Id, Numero, Pasillo, "
                    sql = sql & "  CAJA_ASP, Video ,  ESTANTERIA "
                    sql = sql & " From [P&LCUSTODIA].dbo.Caja"
                    sql = sql & " WHERE     (Numero LIKE '%" & Trim(hojaEx.Cells(P, C_Caja_PYL)) & "%') "
                    rs.Open sql, conPYL
                    If rs.RecordCount = 1 Then
                        If CDbl(hojaEx.Cells(P, C_ID_Caja)) = 0 Then
                            hojaEx.Cells(P, C_ID_Caja) = rs!ID
                        End If
                   Else
                    If Trim(hojaEx.Cells(P, C_ID_Caja)) = "" Then
                        hojaEx.Cells(P, C_ID_Caja) = "C:" & rs.RecordCount
                    End If
                    
                   End If
                   End If
                   
                  
                  Debug.Print hojaEx.Cells(P, C_ID_Caja)

                   If Trim(hojaEx.Cells(P, C_ID_Caja)) <> "" Then
                            
                        If IsNumeric(hojaEx.Cells(P, C_ID_Caja)) Then
                            Set rs = New ADODB.Recordset
                            rs.CursorLocation = adUseClient
                            sql = " SELECT Id, Numero, Pasillo,"
                            sql = sql & " CAJA_ASP, Video , ESTANTERIA "
                            sql = sql & " From [P&LCUSTODIA].dbo.Caja"
                            sql = sql & " WHERE   Id = " & Trim(hojaEx.Cells(P, C_ID_Caja))
                            rs.Open sql, conPYL
                            If IsNull(rs!CAJA_ASP) Then
                                hojaEx.Cells(P, C_CAJA_ASP) = "NO TIENE"
                            Else
                                hojaEx.Cells(P, C_CAJA_ASP) = rs!CAJA_ASP & " " & BuscarEstanteria(rs!CAJA_ASP)
                             End If
                            hojaEx.Cells(P, C_Video) = rs!Video
                            hojaEx.Cells(P, C_ESTANTERIA) = rs!ESTANTERIA
                        End If
                    
                    End If
                    
                  
                
            
        
       Next
        hojaEx.Protect 21877471
           
libroEx.SaveAs "Z:\Tareas\Migracion  de P&L\Planta\pedidos\" & InputBox("Ingrese el nombre de la planilla") & ".xls"
           libroEx.Close
           ApExcel.Quit
           Set ApExcel = Nothing
           Set libroEx = Nothing
End Sub

Private Sub Command6_Click()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset

sql = "SELECT Id, Caja_Asp, ESTANTERIA"
sql = sql & " From Caja"
sql = sql & " Where (Not (CAJA_ASP Is Null))"
sql = sql & " ORDER BY Id"

rs.Open sql, conPYL

Do While Not rs.EOF
    sql = " SELECT     posiciones.id, posiciones.posHorizontal, posiciones.posVertical, elementos.id AS id_elementos, elementos.codigo, estanterias.codigo AS ESTANTERIA"
    sql = sql & " FROM         posiciones INNER JOIN"
    sql = sql & " elementos ON posiciones.id = elementos.posicion_id INNER JOIN"
    sql = sql & " estanterias ON posiciones.estante_id = estanterias.id"
    sql = sql & " WHERE elementos.codigo = '" & rs!CAJA_ASP & "'"
    Set Rs2 = New ADODB.Recordset
    Rs2.Open sql, "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=190.151.143.135"
    
    If Not Rs2.EOF Then
        sql = "  Update Caja"
        sql = sql & " SET ESTANTERIA ='E" & Rs2!ESTANTERIA & "  V:" & Rs2!posVertical & " H:" & Rs2!posHorizontal & "'"
        sql = sql & " Where (Not (CAJA_ASP Is Null)) "
        sql = sql & " And ID = " & rs!ID
        conPYL.Execute sql
    
    End If
    

    rs.MoveNext
Loop



End Sub

Private Sub Command7_Click()

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Rs2 As New ADODB.Recordset
    
    
    Set conPYL = New ADODB.Connection
    conPYL.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
    sql = " SELECT VIDEO"
    sql = sql & " From caja "
    sql = sql & " GROUP BY VIDEO"
    sql = sql & " Having (Not (Video Is Null))"
    sql = sql & " ORDER BY VIDEO"


    rs.Open sql, conPYL

    Do While Not rs.EOF
    
        sql = " UPDATE    [P&LCUSTODIA].dbo.VIDEONOMBRE"
        sql = sql & " SET              ENCONTRADO = 'SI'"
        sql = sql & " WHERE     (NombreVideo LIKE '" & Mid(rs!Video, 1, Len(rs!Video) - 4) & "%')"
   
    conPYL.Execute sql
        
    
    
    
    
'        If Dir("\\MIGUEL3-PC\Videos Procesados\" & Mid(rs!Video, 1, Len(rs!Video) - 4) & " CANT 1.MP4", vbArchive) = "" Then
'            If Dir("\\MIGUEL3-PC\Videos Procesados\" & rs!Video, vbArchive) <> "" Then
'                Rem FileCopy "\\MIGUEL3-PC\Videos Procesados\" & rs!Video, "\\MIGUEL3-PC\Videos Procesados\" & Mid(rs!Video, 1, Len(rs!Video) - 4) & " CANT 1.MP4"
'             Else
'                Debug.Print rs!Video
'            End If
'        End If
        rs.MoveNext
    Loop



'   Dim MyName As String
    
'   MyName = Dir("\\MIGUEL3-PC\Videos Procesados\*.MP4")
'        Do While MyName <> ""
'
'
'            sql = " INSERT INTO [P&LCUSTODIA].dbo.VIDEONOMBRE"
'            sql = sql & "( NombreVideo )"
'            sql = sql & " VALUES ('" & MyName & "')"
'            conPYL.Execute sql
'
'
''            sql = " INSERT INTO [P&LCUSTODIA].dbo.VIDEONOMBRE"
''            sql = sql & "( NombreVideo )"
''            sql = sql & " VALUES ('" & Mid(MyName, 1, Len(MyName) - 4) & " CANT 1.MP4" & "')"
''            conPYL.Execute sql
'
'
'
'            MyName = Dir()
'        Loop
'



End Sub

Private Sub Command8_Click()
 CopiarDatosGrilla grdCajas
End Sub

Private Sub Form_Load()

    Set conPYL = New ADODB.Connection
    conPYL.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
    
    Set conBasa = New ADODB.Connection
    conBasa.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"




End Sub

Private Sub grdCajas_DblClick()
grdCajas.Col = 0
lblID.Caption = grdCajas.Text
End Sub

Public Sub ejecutar(sql As String)

Dim conPyl3 As New ADODB.Connection

conPyl3.Open "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=P&LCUSTODIA;Data Source=222.15.19.150"
conPyl3.Execute sql
End Sub

Public Function BuscarEstanteria(CajaASP As String) As String


Dim sql As String
Dim rs As New ADODB.Recordset
  Dim strConAsp150 As String
strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"




sql = sql & "  SELECT     elementos.id, elementos.codigo, elementos.estado, elementos.posicion_id, estanterias.codigo AS estante, posiciones.posHorizontal, posiciones.posVertical"
sql = sql & " FROM         elementos INNER JOIN"
sql = sql & "                       posiciones ON elementos.posicion_id = posiciones.id INNER JOIN"
sql = sql & "                       estanterias ON posiciones.estante_id = estanterias.id"
sql = sql & "   WHERE     (elementos.codigo = '" & CajaASP & "')"


rs.Open sql, strConAsp150

If Not rs.EOF Then
BuscarEstanteria = "E:" & rs!estante & " H:" & rs!posHorizontal & " V:" & rs!posVertical
Else
BuscarEstanteria = " E:NOtiene"

End If



End Function
