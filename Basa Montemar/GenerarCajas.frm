VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmGeneracionCajas 
   Caption         =   "Generacion de Cajas"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   9060
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   2820
      TabIndex        =   27
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Impresión en termica "
      Height          =   495
      Left            =   6240
      TabIndex        =   26
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtCantidad 
      Height          =   375
      Left            =   7200
      TabIndex        =   25
      Top             =   240
      Width           =   1575
   End
   Begin MSComctlLib.ProgressBar pbsEstado 
      Height          =   435
      Left            =   240
      TabIndex        =   24
      Top             =   1020
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   300
      TabIndex        =   23
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdCajas 
      Caption         =   "Rollo"
      Height          =   435
      Left            =   2040
      TabIndex        =   22
      Top             =   180
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1260
      TabIndex        =   21
      Top             =   7680
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Impresion Custodia Todas"
      Height          =   375
      Left            =   1020
      TabIndex        =   20
      Top             =   7140
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.CommandButton cmdImpresionCustodia 
      Caption         =   "Impresión Custodia"
      Height          =   375
      Left            =   780
      TabIndex        =   19
      Top             =   6600
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CommandButton cmdImpresionModulos 
      Caption         =   "Impresion Modulos"
      Height          =   435
      Left            =   4920
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtCajaFin 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   5760
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txtCajaInicio 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmdImpresion 
      Caption         =   "Impresión"
      Height          =   375
      Left            =   2940
      TabIndex        =   15
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ComboBox cboSucursal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "GenerarCajas.frx":0000
      Left            =   2520
      List            =   "GenerarCajas.frx":0010
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.CommandButton cmdGenerarCajas 
      Caption         =   "Generar Cajas"
      Height          =   435
      Left            =   5280
      TabIndex        =   12
      Top             =   180
      Width           =   1455
   End
   Begin VB.TextBox txtColumnas 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Text            =   "8"
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFilas 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Text            =   "5"
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCantidadModulos 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Sucursal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      TabIndex        =   14
      Top             =   5100
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Caja Final"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4020
      TabIndex        =   11
      Top             =   4560
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblCajaFinal 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5580
      TabIndex        =   10
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Caja Inicial"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4020
      TabIndex        =   9
      Top             =   4020
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblCajaInicial 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCantidadCajas 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5580
      TabIndex        =   7
      Top             =   3480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Cantidad Cajas:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4020
      TabIndex        =   6
      Top             =   3540
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Columnas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Filas:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      TabIndex        =   2
      Top             =   4020
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Cantidad Modulos :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   780
      TabIndex        =   0
      Top             =   3540
      Visible         =   0   'False
      Width           =   1635
   End
End
Attribute VB_Name = "frmGeneracionCajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCajas_Click()
 Dim sql As String
 sql = " SELECT ID_CAJA,DIGITO_VERIFICADOR,ROLLO "
 sql = sql & " FROM   CAJAS "
 sql = sql & " WHERE ROLLO = " & InputBox("Ingrese el rollo")
 sql = sql & " ORDER BY ID_CAJA desc"
 
frmReportes.ImprimirReporte PasoReportes + "cajasbasaEtiquetas.rpt", sql, True

End Sub

Private Sub Command1_Click()
 Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim Modulo As String
        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=E:\datas10112010\Movimiento\Mov10112010.mdb"
        sql = " SELECT ID , IDCAJA, CAJAS.UBICACION, CAJAS.UBICACION_TEXT , "
        sql = sql & " cajas.Digito_Verificador , BARRA_CAJA, CAJA_TEXTO , "
        sql = sql & " MODULO,MODULO_TEXT , DIGITO_TEXT  , HORIZONTAL, VERTICAL"
        sql = sql & " FROM CAJAS "
        sql = sql & " order by ID "
        rs.CursorLocation = adUseClient
        rs.Open sql, con, adOpenKeyset, adLockOptimistic
        Do While Not rs.EOF
'            rs!UBICACION_TEXT =
'            rs!Modulo = CInt(Mid(rs!UBICACION, 1, 6))
'            rs!MODULO_TEXT = Format(CInt(Mid(rs!UBICACION, 1, 6)), "00000")
'            rs!CAJA_TEXTO = rs!IDCAJA
'            rs!Digito_Verificador = Digito_Verificador(rs!IDCAJA)
'            rs!DIGITO_TEXT = Format(rs!Digito_Verificador, "00")
'            rs!BARRA_CAJA = "C6" & Format(CLng(rs!IDCAJA), "0000000") & Format(rs!Digito_Verificador, "00")
            rs!Horizontal = Mid(rs!Ubicacion, 7, 2)
            rs!Vertical = Mid(rs!Ubicacion, 9, 2)
            rs.Update
            rs.MoveNext
         Loop
         MsgBox "Terminado"

End Sub

Private Sub cmdGenerarCajas_Click()
        
    Dim C As Long
    Dim sql As String
    Dim MaxCaja As Long
    Dim rsMaxCajas As New ADODB.Recordset
    Dim MaxRollo As Integer
    Dim rsMaxRollo As New ADODB.Recordset
    Dim CantRollo As Integer
    Dim Rollo As Integer
    Dim cantCajas As Long
    Dim fecha As String
fecha = SysDate

    
    
        MousePointer = 11
        
        rsMaxCajas.Open "SELECT MAX(ID_CAJA) AS MaxCajas FROM CAJAS ", strConBasa
        rsMaxRollo.Open " SELECT MAX(ROLLO) AS Rollo FROM CAJAS ", strConBasa
        MaxRollo = rsMaxRollo!Rollo
        MaxCaja = rsMaxCajas!MaxCajas
        cantCajas = txtCantidad.Text * 3950
        pbsEstado.Max = cantCajas
        pbsEstado.value = 0
        For Rollo = 1 To txtCantidad.Text
            MaxRollo = MaxRollo + 10
            For cantCajas = 1 To 3950
                pbsEstado.value = pbsEstado + 1
                MaxCaja = MaxCaja + 1
                Etiqueta = 110000000000# + MaxCaja
                
                    sql = " INSERT INTO CAJAS "
                    sql = sql & "  (ID_CAJA, NRO_CAJA,  FK_ESTADO "
                    sql = sql & "  , FECHA_CREACION_CAJA, FK_USUARIO_CREACION_CAJA, DIGITO_VERIFICADOR , ETIQUETA, ROLLO ) "
                    sql = sql & "  VALUES    "
                    sql = sql & "( " & MaxCaja & "," & MaxCaja & ", 4 "
                    sql = sql & "," & fecha & ",99," & DigitoEAN13(Trim(Str(Etiqueta))) & ",'" & Etiqueta & "'," & MaxRollo & ")"
                    
                    
                    ExecutarSql sql
                
    Next
    Next
       
        MousePointer = 0
        MsgBox "Operacion terminada"



End Sub


Private Sub cmdImpresion_Click()



Dim sql As String
        sql = " SELECT dbo.CAJAS.ID_CAJA, dbo.CAJAS.DIGITO_VERIFICADOR, dbo.MODULOS.MODULO, dbo.MODULOS.COLUMNAS, dbo.MODULOS.FILAS"
        sql = sql & " FROM   dbo.CAJAS INNER JOIN "
        sql = sql & " dbo.MODULOS ON dbo.CAJAS.FK_MODULO = dbo.MODULOS.ID_MODULOS "
        sql = sql & " Where (dbo.Cajas.FK_CLIENTE Is Null) "
        sql = sql & " And dbo.Cajas.ID_CAJA between  " & txtCajaInicio.Text & " AND " & txtCajaFin.Text
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Modulo As String
rs.Open sql, ConActiva, 0, 1
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=\\Serverbasa1\SistemasBasa\Etiquetas\CAJAS.mdb"


con.Execute " DELETE * FROM CAJAS"

Do While Not rs.EOF
    Modulo = Format(rs!Modulo, "000000") & "F" & Format(rs!Filas, "00") & "C" & Format(rs!Columnas, "00")
    sql = " INSERT INTO CAJAS ( ID_CAJA, DIGITO_VERIFICADOR, MODULO , ID_CAJA_TEXT,BARRA_CAJA)"
    sql = sql & "VALUES (" & rs!ID_CAJA & "," & rs!Digito_Verificador & ",'" & Modulo & "','" & CStr(rs!ID_CAJA) & "','" & "C5" & Format(rs!ID_CAJA, "0000000") & Format(rs!Digito_Verificador, "00") & "') "
    con.Execute sql
    
    rs.MoveNext
Loop


MsgBox "Terminado"




End Sub

Private Sub cmdImpresionCustodia_Click()
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim Modulo As String
    
        con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=\\Serverbasa1\SistemasBasa\Etiquetas\CAJASCustodia.mdb"
        sql = " SELECT CAJAS.ID_CAJA, CAJAS.UBICACION, CAJAS.UBICACION_TEXT, "
        sql = sql & "  cajas.Digito_Verificador , cajas.BARRA_CAJA, CAJA_TEXTO "
        sql = sql & "  FROM CAJAS "
        sql = sql & " WHERE  Digito_Verificador = 0"
        sql = sql & " order by ID_CAJA "
        rs.CursorLocation = adUseClient
        
        sql = " SELECT CAJAS.ID_CAJA, CAJAS.UBICACION, CAJAS.UBICACION_TEXT, "
        sql = sql & "  cajas.Digito_Verificador , cajas.BARRA_CAJA, CAJA_TEXTO "
        sql = sql & "  FROM CAJAS "
        sql = sql & " WHERE  Digito_Verificador = 0"
        sql = sql & " "
        
        
        sql = " SELECT ROLLO.Rollo, ROLLO.Caja_Desde, ROLLO.Caja_Hasta, ROLLO.Cantidad"
        sql = sql & "  From rollo "
        sql = sql & "  ORDER BY ROLLO.Rollo "

        
        
        rs.Open sql, con, adOpenKeyset, adLockOptimistic
        
        
        
        
        
        Do While Not rs.EOF
        
         sql = " UPDATE CAJAS "
         sql = sql = "   SET CAJAS.ROLLO = " & rs!Rollo
         sql = sql & " WHERE CAJAS.ID_CAJA "
         sql = sql & "  Between " & rs!Caja_Desde
         sql = sql & "   And " & rs!Caja_Hasta

            con.Execute sql
            
            
'
'            rs!UBICACION_TEXT = CInt(Mid(rs!Ubicacion, 1, 6)) & "-" & Mid(rs!Ubicacion, 7, 2) & "-" & Mid(rs!Ubicacion, 9, 2)
'            rs!CAJA_TEXTO = rs!ID_CAJA
'            rs!Digito_Verificador = Digito_Verificador(rs!ID_CAJA)
'            rs!BARRA_CAJA = "C5" & Format(rs!ID_CAJA, "0000000") & rs!Digito_Verificador
'            rs.Update
            rs.MoveNext
        Loop
        MsgBox "Terminado"

End Sub


Private Sub cmdImpresionModulos_Click()
Dim sql As String
Dim rs As ADODB.Recordset
Dim ConModulos As New ADODB.Connection

sql = " SELECT     dbo.MODULOS.MODULO "
sql = sql & " FROM         dbo.CAJAS INNER JOIN "
sql = sql & " dbo.MODULOS ON dbo.CAJAS.FK_MODULO = dbo.MODULOS.ID_MODULOS "
sql = sql & "  WHERE     (dbo.CAJAS.ID_CAJA BETWEEN " & txtCajaInicio.Text & " AND " & txtCajaFin.Text & ")"
sql = sql & "  GROUP BY dbo.MODULOS.MODULO "
sql = sql & "  ORDER BY dbo.MODULOS.MODULO "

ConModulos.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=\\Serverbasa1\SistemasBasa\Etiquetas\Modulos.mdb"


ConModulos.Execute " DELETE * FROM MODULOS ; "
Dim Modulo As Long
Dim MODULO_TEXT   As String
Dim BARRA_MODULO As String
Set rs = New ADODB.Recordset
rs.Open sql, ConActiva, 0, 1

Do While Not rs.EOF
    
Modulo = rs!Modulo
MODULO_TEXT = "'" & rs!Modulo & "'"
BARRA_MODULO = "'M1" & Format(rs!Modulo, "000000") & "'"
sql = " INSERT INTO MODULOS ( MODULO, MODULO_TEXT, BARRA_MODULO )"
sql = sql & " values(" & Modulo & "," & MODULO_TEXT & "," & BARRA_MODULO & ")"
ConModulos.Execute sql
    rs.MoveNext
Loop
MsgBox "Terminado"

End Sub

Private Sub Command3_Click()
Dim rs As New ADODB.Recordset


Dim sql As String

sql = "SELECT     CAJA, CLIENTE, GALPON"
sql = sql & " From GALPONES"
sql = sql & " ORDER BY CAJA, CLIENTE"

rs.Open sql, strConBasa

Do While Not rs.EOF

If Not IsNull(rs!Cliente) Then
sql = "  Update Cajas"
sql = sql & " SET  DEPOSITO = '" & Trim(rs!GALPON) & "'"
sql = sql & " Where FK_CLIENTE = " & rs!Cliente
sql = sql & "  And NRO_CAJA = " & rs!Caja
 ExecutarSql sql
 End If
    rs.MoveNext

Loop


End Sub

Private Sub Command4_Click()
Dim ConEtiquetas As New ADODB.Connection
ConEtiquetas.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\cajas.mdb;Persist Security Info=False"

Dim rs As New ADODB.Recordset
Dim sql As String
Dim C As Long
Dim Rollo As Integer
Dim DIGITOVERIFICADOR As Integer




sql = " SELECT     ID_CAJA, ETIQUETA, DIGITO_VERIFICADOR, ROLLO"
sql = sql & "  From basasql.dbo.CAJAS"
sql = sql & "  Where (ROLLO > 1120)"
sql = sql & "  ORDER BY ROLLO, ID_CAJA DESC"




'Sql = " SELECT     ID_CAJA, ETIQUETA, DIGITO_VERIFICADOR, ROLLO"
'Sql = Sql & "  From basasql.dbo.CAJAS"
'Sql = Sql & "  WHERE     (ID_CAJA BETWEEN 891613  and 891695 ) "
'Sql = Sql & "  ORDER BY ROLLO, ID_CAJA DESC"





Rollo = 0
rs.Open sql, strConBasa



Do While Not rs.EOF
Etiqueta = 110000000000# + rs!ID_CAJA

            If Rollo <> rs!Rollo Then
                    Rollo = rs!Rollo
                    sql = " SELECT ID , ROLLO, ID_CAJA ,ETIQUETA,DIGITO, CAJA INTO " & Rollo & " FROM ETIQUETA;"
                    ConEtiquetas.Execute sql
             End If
             
              

               
              sql = " INSERT INTO " & Rollo
               sql = sql & "( ROLLO , "
               sql = sql & " ID_CAJA, "
               sql = sql & " ETIQUETA, "
               sql = sql & " DIGITO  "
               sql = sql & "  )"
               sql = sql & " VALUES ("
               sql = sql & rs!Rollo & ", '"
               sql = sql & rs!ID_CAJA & "', "
               sql = sql & "'" & Etiqueta & "', "
               sql = sql & rs!Digito_Verificador
               sql = sql & " )"


               
               
               
                ConEtiquetas.Execute sql
                    
                   
                    Debug.Print Rollo
            

        rs.MoveNext
        
        
        Loop
        
        MsgBox "Operacion terminada"



End Sub

Private Sub Command5_Click()

Dim ConEtiquetas As New ADODB.Connection
ConEtiquetas.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\CAJAS.mdb;Persist Security Info=False"

Dim rs As New ADODB.Recordset
Dim sql As String




sql = " SELECT   ROLLO, MIN(ID_CAJA) AS Minimo, MAX(ID_CAJA) AS Maximo ,  (MIN(ID_CAJA) - MAX(ID_CAJA) ) as Cantidad"
sql = sql & "  From basasql.dbo.CAJAS"
sql = sql & "  GROUP BY ROLLO"
sql = sql & "  Having (ROLLO > 1020 )"
sql = sql & "  ORDER BY ROLLO"


Rollo = 0
rs.Open sql, strConBasa
ConEtiquetas.Execute "DELETE * FROM ROLLOS "



Do While Not rs.EOF

                    
                    sql = " INSERT INTO ROLLOS ( ROLLO, DESDE, HASTA )"
                    sql = sql & " VALUES (" & rs!Rollo & "," & rs!Minimo & "," & rs!Maximo & " )"
                    ConEtiquetas.Execute sql
                    sql = " INSERT INTO ROLLOS ( ROLLO, DESDE, HASTA )"
                    sql = sql & " VALUES (" & rs!Rollo & "," & rs!Minimo & "," & rs!Maximo & " )"
                    ConEtiquetas.Execute sql
              
          
        
        rs.MoveNext
        
        
        Loop
        
        MsgBox "Operacion terminada"




End Sub


Private Sub txtCantidadModulos_Change()
CalculoCantidad
End Sub

Private Sub txtColumnas_Change()
CalculoCantidad
End Sub

Public Sub CalculoCantidad()
On Error GoTo salir

lblCantidadCajas.Caption = txtCantidadModulos.Text * txtFilas.Text * txtColumnas.Text
Exit Sub

salir:
lblCantidadCajas.Caption = 0

End Sub

Private Sub txtFilas_Change()
CalculoCantidad
End Sub

