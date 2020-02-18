VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMigracion 
   Caption         =   "Migración"
   ClientHeight    =   7350
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   10020
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   4440
      TabIndex        =   6
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   6600
      Width           =   1695
   End
   Begin VB.TextBox txtClienteCustodia 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   240
      Width           =   1275
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   1020
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   8493
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   16
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
         Name            =   "Calibri"
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
      Height          =   315
      Left            =   5280
      TabIndex        =   4
      Top             =   6000
      Width           =   1575
   End
End
Attribute VB_Name = "frmMigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim rs As New ADODB.Recordset
Dim Sql As String

Sql = " SELECT     ID, CLIENTE_CUSTODIA, NOMBRESUCURSAL, NOMBRETIPODOCUMENTO, CANTIDAD, CLIENTE_BASA, DOCUMENTO, COPIAR_DESCR"
Sql = Sql & " From INTERCAMBIO"
Sql = Sql & " Where CLIENTE_CUSTODIA = " & InputBox("Ingrese el cliente custodia")
Sql = Sql & " ORDER BY CANTIDAD"


    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open Sql, ConActiva, 2, 3
    
  
Set DataGrid1.DataSource = rs
DataGrid1.Rebind
DataGrid1.Refresh



End Sub

Private Sub Command2_Click()
Dim MyName As String
Dim ID As Long
Dim NRO_CAJA As Long
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rsCustodia As New ADODB.Recordset
Dim con As New ADODB.Connection

Dim COD_CLIENTE As Integer

Dim Indice As String
Dim FK_INDICES As Long
Dim Descripcion As String
Dim FECHA_DESDE As String
Dim FECHA_HASTA As String
Dim NRO_DESDE As String
Dim NRO_HASTA As String
Dim LETRA_DESDE As String
Dim LETRA_HASTA As String
Dim PASOARCHIVO As String
Dim FK_PERSONAL_CREACION As String
Dim FECHA_CREACION As String
Dim Descrip_Suc_Doc As String
Dim RsIntercambio As ADODB.Recordset
Dim ID_CUSTODIA As Long

Dim maxLegajos As Long

Set rs = New ADODB.Recordset


rs.Open " SELECT MAX(ID_LEGAJO) AS MaxLegajos From LEGAJOS", ConActiva, 0, 1

maxLegajos = rs!maxLegajos + 1



Dim CLIENTE_CUSTODIA, IDSUCURSAL, NOMBRESUCURSAL, IDTIPODOCUMENTO, NOMBRETIPODOCUMENTO, cantidad, CLIENTE_BASA As String

con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\DatCus\referencias.accdb;Persist Security Info=False"
Dim AntNOMBRESUCURSAL As String
Dim AntNOMBRETIPODOCUMENTO As String


          
          Sql = " SELECT IDDOCUMENTO, IDCAJA, NOMBRETIPODOCUMENTO,NOMBRESUCURSAL,"
                Sql = Sql & " Descripcion , fechadesde, FechaHasta, DESDENUMERO, HASTANUMERO"
                Sql = Sql & " From DCTO" & Format(txtClienteCustodia.Text, "0000")
                Sql = Sql & "  ORDER BY NOMBRETIPODOCUMENTO, NOMBRESUCURSAL , IDCAJA , IDDOCUMENTO ;"
          
          
            Set rs = New ADODB.Recordset
            rs.Open Sql, con
             Dim i As Long
             AntNOMBRESUCURSAL = "l"
             AntNOMBRETIPODOCUMENTO = "l"
             
  
  Do While Not rs.EOF
   i = i + 1
  
                If Not (AntNOMBRESUCURSAL = Trim(rs!NOMBRESUCURSAL) And AntNOMBRETIPODOCUMENTO = Trim(rs!NOMBRETIPODOCUMENTO) And (Trim(rs!NOMBRETIPODOCUMENTO) = "" Or Trim(rs!NOMBRESUCURSAL) = "")) Then
    
                        Sql = "  SELECT     CLIENTE_CUSTODIA, NOMBRESUCURSAL, NOMBRETIPODOCUMENTO, CLIENTE_BASA, DOCUMENTO, COPIAR_DESCR"
                        Sql = Sql & "  From INTERCAMBIO "
                        Sql = Sql & " WHERE   CLIENTE_CUSTODIA = " & txtClienteCustodia.Text
                        Sql = Sql & " AND NOMBRESUCURSAL = '" & Trim(Replace(Replace(UCase(Trim(rs!NOMBRESUCURSAL)), Chr(0), ""), Chr(16), "")) & "'"
                        Sql = Sql & " AND NOMBRETIPODOCUMENTO = '" & Trim(Replace(Replace(UCase(Trim(rs!NOMBRETIPODOCUMENTO)), Chr(0), ""), Chr(16), "")) & "'"
                        Set RsIntercambio = New ADODB.Recordset
                        RsIntercambio.Open Sql, ConActiva, 0, 1
                        If RsIntercambio.EOF Then
                            MsgBox "Error"
                        Else
                            AntNOMBRESUCURSAL = Trim(rs!NOMBRESUCURSAL)
                            AntNOMBRETIPODOCUMENTO = Trim(rs!NOMBRETIPODOCUMENTO)
                            COD_CLIENTE = RsIntercambio!CLIENTE_BASA
                             If BuscarIndice(RsIntercambio!Documento, RsIntercambio!CLIENTE_BASA) = "Error" Then
                             MsgBox "eRROR ENE EL PROCESO DOCUMENTO NO VALIDO " & RsIntercambio!Documento
                             Exit Sub
                             Else
                                 Indice = "'" & BuscarIndice(RsIntercambio!Documento, RsIntercambio!CLIENTE_BASA) & "'"
                                FK_INDICES = Buscar_ID_Indice(RsIntercambio!Documento, RsIntercambio!CLIENTE_BASA)
                            End If
                            FK_PERSONAL_CREACION = 99
                            FECHA_CREACION = FechaFormato(Now)
                            Descrip_Suc_Doc = UCase(Trim(rs!NOMBRESUCURSAL)) & "//" & Trim(Replace(Replace(UCase(Trim(rs!NOMBRETIPODOCUMENTO)), Chr(0), ""), Chr(16), "")) & "//"
                        End If
                   End If
                        ID_CUSTODIA = rs!IDDOCUMENTO
                        NRO_CAJA = rs!IDCaja
                        If RsIntercambio!COPIAR_DESCR = True Then
                            Descripcion = UCase(Descrip_Suc_Doc) & UCase(Trim(rs!Descripcion))
                        Else
                            Descripcion = UCase(Trim(rs!Descripcion))
                        End If
                        Descripcion = "'" & Replace(Descripcion, "'", " ") & "'"
                        
                        
                        If rs!fechadesde = 0 Then
                        FECHA_DESDE = "Null"
                        Else
                        FECHA_DESDE = FechaFormato(DateAdd("d", rs!fechadesde, "28/12/1800"))
                        End If
                        
                        If rs!FechaHasta = 0 Then
                          FECHA_HASTA = "NULL"
                        Else
                            FECHA_HASTA = FechaFormato(DateAdd("d", rs!FechaHasta, "28/12/1800"))
                        End If
                        
                        NRO_DESDE = rs!DESDENUMERO
                        NRO_HASTA = rs!HASTANUMERO
                        
                        
                        If rs!DESDENUMERO = rs!HASTANUMERO And rs!DESDENUMERO <> 0 Then
                               InsertarLegajos maxLegajos, COD_CLIENTE, NRO_CAJA, Indice, FK_INDICES, Descripcion, FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA, PASOARCHIVO, FK_PERSONAL_CREACION, FECHA_CREACION, ID_CUSTODIA, txtClienteCustodia.Text
                               maxLegajos = maxLegajos + 1
                         Else
                               InsertarReferencias COD_CLIENTE, NRO_CAJA, Indice, FK_INDICES, Descripcion, FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA, FK_PERSONAL_CREACION, FECHA_CREACION, CLng(ID_CUSTODIA), txtClienteCustodia.Text
                         End If
                        
                        
             Label1.Caption = i
             Label1.Refresh
             rs.MoveNext
             Loop
             
          MsgBox "Terminado"



End Sub

Private Sub Command3_Click()
Dim con As New ADODB.Connection
con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\DatCus\referencias.accdb;Persist Security Info=False"

Dim rs As New ADODB.Recordset

Dim Sql As String

Sql = " SELECT CAJAS.IDCliente, CAJAS.IDCaja, CAJAS.Estado"
Sql = Sql & "  From Cajas"
Sql = Sql & "   WHERE  (((CAJAS.IDCliente)<>0)) and     (((CAJAS.IDCliente) In (91    ,95 ,96 ,97 ,98 ,99 ,100    ,101    ,104    ,105    )))"
Sql = Sql & " ORDER BY CAJAS.IDCaja;"

Dim estado As Integer
 rs.Open Sql, con

Do While Not rs.EOF

Select Case UCase(Trim(rs!estado))
Case "EN TRANSITO"
estado = 3
Case Is = "ENVIO BASA"
estado = 0
Case "LIBRE"
estado = 0
Case "OCUPADA"
estado = 2
Case "RESERVA"
   Case Else
    MsgBox "eRROR"
    estado = 0

End Select



If estado <> 0 Then

Sql = " Update CONTENEDOR"
Sql = Sql & " Set estado =  " & estado
Sql = Sql & " Where NRO_CAJA = " & rs!IDCaja
Sql = Sql & "  And COD_CLIENTE = 172 "

 ExecutarSql Sql
End If


    rs.MoveNext
Loop



End Sub

Private Sub Command4_Click()
Dim MyName As String
Dim ID As Long
Dim NRO_CAJA As Long
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rsCustodia As New ADODB.Recordset
Dim con As New ADODB.Connection

Dim COD_CLIENTE As Integer

Dim Indice As String
Dim FK_INDICES As Long
Dim Descripcion As String
Dim FECHA_DESDE As String
Dim FECHA_HASTA As String
Dim NRO_DESDE As String
Dim NRO_HASTA As String
Dim LETRA_DESDE As String
Dim LETRA_HASTA As String
Dim PASOARCHIVO As String
Dim FK_PERSONAL_CREACION As String
Dim FECHA_CREACION As String
Dim Descrip_Suc_Doc As String
Dim RsIntercambio As ADODB.Recordset
Dim ID_CUSTODIA As Long

Dim maxLegajos As Long

Set rs = New ADODB.Recordset


rs.Open " SELECT MAX(ID_LEGAJO) AS MaxLegajos From LEGAJOS", ConActiva, 0, 1

maxLegajos = rs!maxLegajos + 1



Dim CLIENTE_CUSTODIA, IDSUCURSAL, NOMBRESUCURSAL, IDTIPODOCUMENTO, NOMBRETIPODOCUMENTO, cantidad, CLIENTE_BASA As String

con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:DISCO.accdb;Persist Security Info=False"
Dim AntNOMBRESUCURSAL As String
Dim AntNOMBRETIPODOCUMENTO As String


          
          Sql = " SELECT IDDOCUMENTO, IDCAJA, NOMBRETIPODOCUMENTO,NOMBRESUCURSAL,"
                Sql = Sql & " Descripcion , fechadesde, FechaHasta, DESDENUMERO, HASTANUMERO"
                Sql = Sql & " From DISCO"
                Sql = Sql & "  ORDER BY NOMBRETIPODOCUMENTO, NOMBRESUCURSAL , IDCAJA , IDDOCUMENTO ;"
          
          
            Set rs = New ADODB.Recordset
            rs.Open Sql, con
             Dim i As Long
             AntNOMBRESUCURSAL = "l"
             AntNOMBRETIPODOCUMENTO = "l"
             

   
  Do While Not rs.EOF
   
                        ID_CUSTODIA = rs!IDDOCUMENTO
                        NRO_CAJA = rs!IDCaja
                       Descripcion = UCase(Trim(rs!NOMBRESUCURSAL))
                         Descripcion = Descripcion & " // " & Mid(UCase(Trim(rs!NOMBRETIPODOCUMENTO)), 1, 50)
'                         For I = 1 To Len(rs!NOMBRETIPODOCUMENTO)
'                         Debug.Print Mid(UCase(Trim(rs!NOMBRETIPODOCUMENTO)), I, 1)
'
'                         Debug.Print I
'                         Debug.Print Asc(Mid(UCase(Trim(rs!NOMBRETIPODOCUMENTO)), I, 1))
'                         Next
                         
                         
                         Descripcion = Descripcion & " // " & UCase(Trim(rs!Descripcion))
                         Descripcion = Replace(Descripcion, Chr(0), "")
                        Descripcion = Replace(Descripcion, "", "")
                        Descripcion = Replace(Descripcion, vbCrLf, "")
                         Descripcion = Replace(Descripcion, vbCr, "")
                        Descripcion = Replace(Descripcion, "-", " ")
                         Descripcion = Replace(Descripcion, "  ", " ")
                          Descripcion = Replace(Descripcion, "  ", " ")
                           Descripcion = Replace(Descripcion, "  ", " ")
                            Descripcion = Replace(Descripcion, "  ", " ")
                        
                        If rs!fechadesde = 0 Then
                        FECHA_DESDE = "Null"
                        Else
                        FECHA_DESDE = "'" & DateAdd("d", rs!fechadesde, "28/12/1800") & "'"
                        End If
                        
                        If rs!FechaHasta = 0 Then
                          FECHA_HASTA = "NULL"
                        Else
                            FECHA_HASTA = "'" & DateAdd("d", rs!FechaHasta, "28/12/1800") & "'"
                        End If
                        
                        NRO_DESDE = rs!DESDENUMERO
                        NRO_HASTA = rs!HASTANUMERO
                     

              Sql = " INSERT INTO FINAL (CAJA,DETALLE,N_DESDE,N_HASTA,FECHA_DESDE,FECHA_HASTA)"
    Sql = Sql & " VALUES (" & NRO_CAJA & ",'" & Trim(Descripcion) & "'," & NRO_DESDE & "," & NRO_HASTA & "," & FECHA_DESDE & "," & FECHA_HASTA & ")"
  
  con.Execute Sql
  
             rs.MoveNext
             Loop
             
          MsgBox "Terminado"


End Sub

Private Sub Command5_Click()
Dim MyName As String
Dim ID As Long
Dim NRO_CAJA As Long
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rsCustodia As New ADODB.Recordset
Dim con As New ADODB.Connection

Dim COD_CLIENTE As Integer

Dim Indice As String
Dim FK_INDICES As Long
Dim Descripcion As String
Dim FECHA_DESDE As String
Dim FECHA_HASTA As String
Dim NRO_DESDE As String
Dim NRO_HASTA As String
Dim LETRA_DESDE As String
Dim LETRA_HASTA As String
Dim PASOARCHIVO As String
Dim FK_PERSONAL_CREACION As String
Dim FECHA_CREACION As String
Dim Descrip_Suc_Doc As String
Dim RsIntercambio As ADODB.Recordset
Dim ID_CUSTODIA As Long

Dim maxLegajos As Long

Set rs = New ADODB.Recordset


rs.Open " SELECT MAX(ID_LEGAJO) AS MaxLegajos From LEGAJOS", ConActiva, 0, 1

maxLegajos = rs!maxLegajos + 1



Dim CLIENTE_CUSTODIA, IDSUCURSAL, NOMBRESUCURSAL, IDTIPODOCUMENTO, NOMBRETIPODOCUMENTO, cantidad, CLIENTE_BASA As String

con.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\DatCus\datas.accdb;Persist Security Info=False"
Dim AntNOMBRESUCURSAL As String
Dim AntNOMBRETIPODOCUMENTO As String


          
          Sql = " SELECT IDDOCUMENTO, IDCAJA, NOMBRETIPODOCUMENTO,NOMBRESUCURSAL,"
                Sql = Sql & " Descripcion , fechadesde, FechaHasta, DESDENUMERO, HASTANUMERO"
                Sql = Sql & " From DCTO" & Format(txtClienteCustodia.Text, "0000")
                Sql = Sql & "  ORDER BY NOMBRETIPODOCUMENTO, NOMBRESUCURSAL , IDCAJA , IDDOCUMENTO ;"
          
          
            Set rs = New ADODB.Recordset
            rs.Open Sql, con
             Dim i As Long
             AntNOMBRESUCURSAL = "l"
             AntNOMBRETIPODOCUMENTO = "l"
             
  
  Do While Not rs.EOF
                
                
                COD_CLIENTE = txtClienteCustodia.Text + 1000
                             
                                 Indice = "'001'"
                                FK_INDICES = 0
               
                            FK_PERSONAL_CREACION = 99
                            FECHA_CREACION = FechaFormato(Now)
                            Descrip_Suc_Doc = UCase(Trim(rs!NOMBRESUCURSAL)) & "//" & Trim(Replace(Replace(UCase(Trim(rs!NOMBRETIPODOCUMENTO)), Chr(0), ""), Chr(16), "")) & "//"
                    
                        ID_CUSTODIA = rs!IDDOCUMENTO
                        NRO_CAJA = rs!IDCaja
                       
                            Descripcion = Descrip_Suc_Doc & UCase(Trim(rs!Descripcion))
                        Descripcion = "'" & Replace(Descripcion, "'", " ") & "'"
                        
                        
                        If rs!fechadesde = 0 Then
                        FECHA_DESDE = "Null"
                        Else
                        FECHA_DESDE = FechaFormato(DateAdd("d", rs!fechadesde, "28/12/1800"))
                        End If
                        
                        If rs!FechaHasta = 0 Then
                          FECHA_HASTA = "NULL"
                        Else
                            FECHA_HASTA = FechaFormato(DateAdd("d", rs!FechaHasta, "28/12/1800"))
                        End If
                        
                        NRO_DESDE = rs!DESDENUMERO
                        NRO_HASTA = rs!HASTANUMERO
                        
                        
                               InsertarReferencias COD_CLIENTE, NRO_CAJA, Indice, FK_INDICES, Descripcion, FECHA_DESDE, FECHA_HASTA, NRO_DESDE, NRO_HASTA, FK_PERSONAL_CREACION, FECHA_CREACION, CLng(ID_CUSTODIA), txtClienteCustodia.Text
                        
                        
             Label1.Caption = i
             Label1.Refresh
             rs.MoveNext
             Loop
             
          MsgBox "Terminado"

End Sub

Private Sub Form_Resize()
DataGrid1.Left = 0
DataGrid1.Width = frmMigracion.Width - 300
End Sub

Public Sub InsertarLegajos(ID_LEGAJO As Long, COD_CLIENTE As Integer, NRO_CAJA As Long, Indice As String, FK_INDICES As Long, Descripcion As String, _
FECHA_DESDE As String, FECHA_HASTA As String, NRO_DESDE As String, NRO_HASTA As String, PASOARCHIVO As String, FK_PERSONAL_CREACION As String, FECHA_CREACION As String, ID_CUSTODIA As Long, COD_CLIENTE_CUSTODIA As Long)
    
    Dim Sql As String
    Dim FECHA_MODIFICACION As String
    Dim COD_ID_REFERENCIA As Long
    FECHA_MODIFICACION = FECHA_CREACION
    FK_PERSONAL_MODIFICACION = FK_PERSONAL_CREACION
    
    ID_CLIENTE_LEGAJO = ID_LEGAJO
        Sql = "  INSERT INTO LEGAJOS "
       Sql = Sql & vbCrLf & " (  ID_CLIENTE_LEGAJO , ID_LEGAJO"
        Sql = Sql & vbCrLf & "  , COD_CLIENTE, NRO_CAJA "
        Sql = Sql & vbCrLf & " , COD_INDICE, DESCRIPCION, FK_INDICES"
        Sql = Sql & vbCrLf & " , FECHA_DESDE, FECHA_HASTA"
        Sql = Sql & vbCrLf & " , NRO_DESDE, NRO_HASTA"
        Sql = Sql & vbCrLf & " , FK_PERSONAL_CREACION, FK_PERSONAL_ACTUALIZACION"
        Sql = Sql & vbCrLf & " , FECHA_ACTUALIZACION , FECHA_CREACION  "
        Sql = Sql & vbCrLf & " , COD_ESTADO, ID_CUSTODIA, COD_CLIENTE_CUSTODIA )"
        Sql = Sql & vbCrLf & " VALUES  "
        Sql = Sql & vbCrLf & " (" & ID_CLIENTE_LEGAJO & "," & ID_LEGAJO
        Sql = Sql & vbCrLf & " ," & COD_CLIENTE & "," & NRO_CAJA
        Sql = Sql & vbCrLf & " ," & Indice & "," & Descripcion & "," & FK_INDICES
        Sql = Sql & vbCrLf & " ," & FECHA_DESDE & "," & FECHA_HASTA
        Sql = Sql & vbCrLf & " ," & NRO_DESDE & "," & NRO_HASTA
        Sql = Sql & vbCrLf & " ," & FK_PERSONAL_CREACION & "," & FK_PERSONAL_MODIFICACION
        Sql = Sql & vbCrLf & " ," & FECHA_MODIFICACION & "," & FECHA_CREACION
        Sql = Sql & vbCrLf & " ,2," & ID_CUSTODIA & ", " & COD_CLIENTE_CUSTODIA & ")"
        ExecutarSql Sql
        
        
        
End Sub


Public Function InsertarReferencias(COD_CLIENTE As Integer, NRO_CAJA As Long, Indice As String, FK_INDICES As Long, Descripcion As String, _
FECHA_DESDE As String, FECHA_HASTA As String, NRO_DESDE As String, NRO_HASTA As String, FK_PERSONAL_CREACION As String, FECHA_CREACION As String, ID_CUSTODIA As Long, COD_CLIENTE_CUSTODIA As Long)
    
    Dim Sql As String
    Dim FECHA_MODIFICACION As String
    Dim COD_ID_REFERENCIA As Long
    FECHA_MODIFICACION = FECHA_CREACION
    FK_PERSONAL_MODIFICACION = FK_PERSONAL_CREACION
        Sql = "  INSERT INTO REFERENCIAS"
        Sql = Sql & vbCrLf & " ( COD_CLIENTE, NRO_CAJA "
        Sql = Sql & vbCrLf & " ,COD_TIPO_ALMACENAMIENTO, ITEM"
        Sql = Sql & vbCrLf & " , INDICE, DESCRIPCION, FK_INDICES"
        Sql = Sql & vbCrLf & " , FECHA_DESDE, FECHA_HASTA"
        Sql = Sql & vbCrLf & " , NRO_DESDE, NRO_HASTA"
        Sql = Sql & vbCrLf & " , FK_PERSONAL_CREACION, FK_PERSONAL_MODIFICACION"
        Sql = Sql & vbCrLf & " , BORRADO  "
        Sql = Sql & vbCrLf & " , FECHA_MODIFICACION, FECHA_CREACION,ID_CUSTODIA , COD_CLIENTE_CUSTODIA )"
        Sql = Sql & vbCrLf & " VALUES  "
        Sql = Sql & vbCrLf & " (" & COD_CLIENTE & "," & NRO_CAJA
        Sql = Sql & vbCrLf & ",0, 1"
        Sql = Sql & vbCrLf & " ," & Indice & "," & Descripcion & "," & FK_INDICES
        Sql = Sql & vbCrLf & " ," & FECHA_DESDE & "," & FECHA_HASTA
        Sql = Sql & vbCrLf & " ," & NRO_DESDE & "," & NRO_HASTA
        Sql = Sql & vbCrLf & " ," & FK_PERSONAL_CREACION & "," & FK_PERSONAL_MODIFICACION
        Sql = Sql & vbCrLf & " ,0"
        Sql = Sql & vbCrLf & " ," & FECHA_MODIFICACION & "," & FECHA_CREACION & "," & ID_CUSTODIA & "," & COD_CLIENTE_CUSTODIA & ")"
        ExecutarSql Sql
    
End Function
