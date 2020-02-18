VERSION 5.00
Object = "{C981C8C8-C8F3-471A-A947-0318B0DF45F0}#1.0#0"; "Controles4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEtiquetasLegajos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiquetas"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "Etiquetas Legajos Miguel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   25
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rollos"
      Height          =   375
      Left            =   3360
      TabIndex        =   24
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "mdb"
      Height          =   375
      Left            =   5100
      TabIndex        =   23
      Top             =   5880
      Width           =   1395
   End
   Begin VB.CommandButton cmdl 
      Caption         =   "Command2"
      Height          =   375
      Left            =   1740
      TabIndex        =   22
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Pasar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtetiqueta 
      Height          =   375
      Left            =   3540
      TabIndex        =   19
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtcliente 
      Height          =   375
      Left            =   1140
      TabIndex        =   17
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtLegajos 
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   6975
   End
   Begin VB.CommandButton cmdImprimirEtiquetas 
      Caption         =   "Etiquetas Legajos Ana"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   13
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Asignar Responsable"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   7035
      Begin VB.CommandButton cmdAsignar 
         Caption         =   "Asignar"
         Height          =   375
         Left            =   3420
         TabIndex        =   12
         Top             =   1380
         Width           =   1515
      End
      Begin Controles.cltGenerico ctlPersonal 
         Height          =   435
         Left            =   1020
         TabIndex        =   10
         Top             =   960
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   767
      End
      Begin VB.TextBox txtDesdeAsignacion 
         Height          =   345
         Left            =   1020
         TabIndex        =   7
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox txtHastaAsignacion 
         Height          =   345
         Left            =   3300
         TabIndex        =   6
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Personal:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label3 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   2700
         TabIndex        =   8
         Top             =   420
         Width           =   615
      End
   End
   Begin MSComctlLib.ProgressBar pbsEstado 
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   660
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1
      Max             =   6600
   End
   Begin VB.TextBox txtCantidad 
      Height          =   330
      Left            =   1680
      TabIndex        =   2
      Text            =   "0"
      Top             =   120
      Width           =   555
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar Etiquetas"
      Height          =   375
      Left            =   2340
      TabIndex        =   0
      Top             =   60
      Width           =   1755
   End
   Begin VB.Label Label7 
      Caption         =   "Nº de  Etiqueta"
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   18
      Top             =   3180
      Width           =   1395
   End
   Begin VB.Label Label7 
      Caption         =   "Cliente"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   3180
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "Ingrese los Nº de Legajos separados por coma"
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
      Left            =   180
      TabIndex        =   14
      Top             =   3780
      Width           =   4275
   End
   Begin VB.Label Label5 
      Caption         =   "Estado:"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   660
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad Rollo :"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   1335
   End
End
Attribute VB_Name = "frmEtiquetasLegajos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAsignar_Click()
Dim Sql As String
Dim cant As Long

Sql = " UPDATE    dbo.LEGAJOS"
Sql = Sql & "  SET FK_PERSONAL_ASIGNACION =" & ctlPersonal.Valor
Sql = Sql & "  WHERE   ID_LEGAJO BETWEEN " & txtDesdeAsignacion.Text & " AND " & txtHastaAsignacion.Text
 cant = ExecutarSql(Sql)

MsgBox "La acutlizacion se realizo con exito cantidad de registros " & cant


End Sub

Private Sub cmdl_Click()
MsgBox Append_EAN_Checksum("123456789129")
End Sub

Private Sub cmdReparacionLegajos_Click()
Dim rs As New ADODB.Recordset
Dim MaxLegajo As Long


Sql = " SELECT     ID_LEGAJO, ID_CLIENTE_LEGAJO, COD_INDICE, ROLLO, NRO_CAJA, COD_CLIENTE, COD_ESTADO"
Sql = Sql & "  From basasql.dbo.LEGAJOS"
Sql = Sql & " Where (ID_LEGAJO > 4784873)"
rs.Open Sql, strConBasa, 2, 3
MaxLegajo = 4784873

Do While Not rs.EOF
    MaxLegajo = MaxLegajo + 1
    
    rs!ID_LEGAJO = MaxLegajo
    rs!ID_CLIENTE_LEGAJO = MaxLegajo
    
    rs.Update
    
    rs.MoveNext
Loop



End Sub

Private Sub Command1_Click()
 
 
 
 
 
 
 
 
 
 
 Dim Sql As String
 Dim Sql2 As String
 Dim ConEtiquetas As New ADODB.Connection


 ConEtiquetas.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Serverbasa1\SistemasBasa\Etiquetas\Etiquetas.mdb;Persist Security Info=False"

 Dim rs As New ADODB.Recordset

 Sql = " SELECT     ID_LEGAJO, DIGITO_VERIFICADOR , ROLLO "
Sql = Sql & "  From dbo.LEGAJOS"
Sql = Sql & "   WHERE  ID_LEGAJO BETWEEN " & txtDesde.Text & " AND " & txtHasta.Text
Sql = Sql & " ORDER BY ID_LEGAJO DESC "



rs.Open Sql, ConActiva, 0, 1


ConEtiquetas.Execute " DELETE * FROM ETIQUETAS "

Do While Not rs.EOF

'SQL2 = " INSERT INTO ETIQUETAS ( ID_LEGAJO, DIGITO_VERIFICADOR, BARRA , ID_LEGAJO_TEXTO )"
'SQL2 = SQL2 & " values (" & rs!ID_LEGAJO & "," & rs!Digito_Verificador & ",'" & "L2" & Format(rs!ID_LEGAJO, "0000000") & "','" & CStr(rs!ID_LEGAJO) & "')"

Sql2 = " INSERT INTO ETIQUETAS ( ID_LEGAJO, DIGITO_VERIFICADOR, BARRA , ID_LEGAJO_TEXTO, ROLLO )"
Sql2 = Sql2 & " values (" & rs!ID_LEGAJO & "," & rs!Digito_Verificador & ",'" & "L2" & Format(rs!ID_LEGAJO, "0000000") & "','" & CStr(rs!ID_LEGAJO) & "' ," & rs!Rollo & ")"



  
  ConEtiquetas.Execute Sql2
    rs.MoveNext
Loop

MsgBox "terminado"


'
''SELECT ETIQUETAS.Id, ETIQUETAS.ORDEN, ETIQUETAS.ID_LEGAJO, ETIQUETAS.DIGITO_VERIFICADOR, ETIQUETAS.BARRA, ETIQUETAS.ID_LEGAJO_TEXTO INTO 2100
''From ETIQUETAS
''Where (((ETIQUETAS.ROLLO) = 100))
''ORDER BY ETIQUETAS.Id DESC;

'Dim SQL As String
' Dim SQL2 As String
' Dim ConEtiquetas As New ADODB.Connection
' Dim i As Long
' Dim C As Long
'
' ConEtiquetas.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Serverbasa1\SistemasBasa\Etiquetas\ECI.mdb;Persist Security Info=False"
'C = 100000
' For i = 46299 To 50499
'
'    SQL = " INSERT INTO ECI ( ECI_NUMERO, ECI_BARRA )"
'    SQL = SQL & " values (" & i & ",'" & Format(i, "000000") & "')"
'    ConEtiquetas.Execute SQL
'
'  Next
'
'
'MsgBox "terminado"



End Sub

Private Sub cmdBuscar_Click()

Dim Sql As String
Dim rs As New ADODB.Recordset
Sql = "  SELECT     ID_CLIENTE_LEGAJO, COD_CLIENTE, ID_LEGAJO"
Sql = Sql & " From LEGAJOS"
Sql = Sql & "  Where ID_CLIENTE_LEGAJO = " & txtEtiqueta.Text
Sql = Sql & "  And COD_CLIENTE = " & txtCliente.Text
rs.Open Sql, ConActiva

If Not rs.EOF Then
    txtLegajos.Text = txtLegajos.Text & rs!ID_LEGAJO & ","
End If


End Sub

Private Sub cmdGenerar_Click()
Dim C As Long
    Dim Sql As String
    Dim MaxID_LEGAJO As Long
    Dim MaxID_Cliente_Legajo As Long
    Dim COD_CLIENTE As Integer
    Dim sSQL As String
    Dim rsMaxID_LEGAJO As New ADODB.Recordset
    Dim rsMaxID_CLIENTE_LEGAJO As New ADODB.Recordset
    Dim rsMaxRollo As New ADODB.Recordset
    Dim MaxRollo As Integer
    Dim CantRollo As Integer
    Dim Rollo As Integer
    Dim CantidadEtiquetas As Integer
    Dim Etiqueta As Double
        MousePointer = 11
        rsMaxID_LEGAJO.Open "SELECT MAX(ID_LEGAJO) AS MaxID_LEGAJO  From LEGAJOS", ConActiva, 0, 1
        rsMaxRollo.Open " SELECT MAX(ROLLO) AS Rollo From LEGAJOS ", strConBasa
        MaxRollo = rsMaxRollo!Rollo
        MaxID_LEGAJO = rsMaxID_LEGAJO!MaxID_LEGAJO
       
'       pbsEstado.Max = (txtCantidad.Text * 6400) + 1
       pbsEstado.value = 1
       For Rollo = 1 To txtCantidad.Text
            MaxRollo = MaxRollo + 10
            For CantidadEtiquetas = 1 To 10901
               Rem  pbsEstado.value = pbsEstado + 1
                MaxID_LEGAJO = MaxID_LEGAJO + 1
                Etiqueta = 120000000000# + MaxID_LEGAJO
                Sql = "INSERT INTO LEGAJOS (ID_LEGAJO, ID_CLIENTE_LEGAJO, DIGITO_VERIFICADOR , ROLLO , Etiqueta)"
                Sql = Sql & " VALUES ( " & MaxID_LEGAJO & "," & MaxID_LEGAJO & "," & DigitoEAN13(CStr(Etiqueta)) & ", " & MaxRollo & ",'" & Trim(Etiqueta) & "' )"
                
                ExecutarSql (Sql)
'                If MaxID_LEGAJO >= 7409888 Then
'                    MsgBox "LLEGO"
'                End If
                
           Next
    Next
       
        MousePointer = 0
        MsgBox "Operacion terminada"
End Sub

Function Append_EAN_Checksum(RawString As String) As Integer
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
Append_EAN_Checksum = RawString & Format$(CheckSum, "0")
End Function


Private Sub cmdImpr_Click()
Dim Sql As String
Sql = " SELECT LEGAJO_1, LEGAJO_2, ORDEN "
 Sql = Sql & " , DIGITO_1, DIGITO_2, BARRA_1, "
 Sql = Sql & " TEM_LEGAJOS.BARRA_2 "
 Sql = Sql & "  From TEM_LEGAJOS "
 Sql = Sql & "  ORDER BY ORDEN"
 frmReportes.ImprimirReporte PasoReportes & "Etiquetas_Legajos.rpt", Sql, True

End Sub

Private Sub cmdImprimir_Click()
        Dim Sql As String
        Dim Rollo As String
        Dim intROLLO As String
        Dim LEGAJO_1   As String
        Dim LEGAJO_2 As String
        Dim BARRA_1 As String
        Dim BARRA_2  As String
        Dim DIGITO_1 As String
        Dim DIGITO_2 As String
Dim rs As New ADODB.Recordset


Sql = " SELECT     ID_LEGAJO, ROLLO, DIGITO_VERIFICADOR"
Sql = Sql & "  From LEGAJOS "
Sql = Sql & "  Where  Rollo = " & InputBox("Ingrese el rollo")
Sql = Sql & "  ORDER BY ID_LEGAJO DESC"



Rem   '" & "L2" & Format(rs!ID_LEGAJO, "0000000") & "','" & CStr(rs!ID_LEGAJO) & "' ," & rs
  
  rs.Open Sql, strConBasa
  Dim Orden As Integer
  Orden = 1
  
    ExecutarSql "DELETE FROM TEM_LEGAJOS "
    pbsEstado.value = 1
    Do While Not rs.EOF
        frmEtiquetasLegajos.Refresh
        pbsEstado = pbsEstado.value + 1
        pbsEstado.Refresh
        Rollo = "'" & rs!Rollo & "'"
        intROLLO = rs!Rollo
        LEGAJO_1 = "'" & rs!ID_LEGAJO & "'"
        BARRA_1 = "'" & "*L2" & Format(rs!ID_LEGAJO, "0000000") & "*'"
        DIGITO_1 = "'" & rs!Digito_Verificador & "'"

    
   
    
    
    rs.MoveNext
    If Not rs.EOF Then
            LEGAJO_2 = "'" & rs!ID_LEGAJO & "'"
            BARRA_2 = "'" & "*L2" & Format(rs!ID_LEGAJO, "0000000") & "*'"
            DIGITO_2 = "'" & rs!Digito_Verificador & "'"
            Rollo = "'" & rs!Rollo & "'"
            Sql = " INSERT INTO TEM_LEGAJOS (ORDEN, ROLLO, LEGAJO_1, LEGAJO_2, BARRA_1, BARRA_2, DIGITO_1, DIGITO_2) "
            Sql = Sql & " VALUES     (" & Orden & "," & Rollo & "," & LEGAJO_1 & "," & LEGAJO_2 & "," & BARRA_1 & "," & BARRA_2 & "," & DIGITO_1 & "," & DIGITO_2 & ") "
            ExecutarSql Sql
            rs.MoveNext
    Else
    
           
            Sql = " INSERT INTO TEM_LEGAJOS (ORDEN, ROLLO, LEGAJO_1, BARRA_1, DIGITO_1 ) "
            Sql = Sql & " VALUES ( " & Orden & "," & Rollo & "," & LEGAJO_1 & "," & BARRA_1 & "," & DIGITO_1 & ")"
            ExecutarSql Sql

    
    End If
  
    Orden = Orden + 1
  
  
  Loop
  
  Set rs = New ADODB.Recordset
   
  Sql = " SELECT     MAX(ID_LEGAJO) AS MaxLegajo, MIN(ID_LEGAJO) AS MinLegajo From LEGAJOS  Where ROLLO = " & Rollo
  
  If Rollo = "" Then
    MsgBox "EL ROLLO NO EXISTE"
    pbsEstado.value = 1
    Exit Sub
  End If
  
  rs.Open Sql, strConBasa
   
   Sql = " INSERT INTO TEM_LEGAJOS (ORDEN, LEGAJO_1, LEGAJO_2 ) "
            Sql = Sql & " VALUES ( " & Orden & ",'ROLLO:" & intROLLO & " D:" & rs!MinLegajo & "','H:" & rs!MaxLegajo & "')"
            ExecutarSql Sql
            ExecutarSql Sql
            ExecutarSql Sql
            ExecutarSql Sql
   
   MsgBox "TERMINADO"

   
   
   
 Sql = " SELECT LEGAJO_1, LEGAJO_2, ORDEN "
 Sql = Sql & " , DIGITO_1, DIGITO_2, BARRA_1, "
 Sql = Sql & " TEM_LEGAJOS.BARRA_2 "
 Sql = Sql & "  From TEM_LEGAJOS "
 Sql = Sql & "  ORDER BY ORDEN"
 frmReportes.ImprimirReporte PasoReportes & "Etiquetas_Legajos.rpt", Sql, True
End Sub

Private Sub cmdImprimirPapel_Click()
Dim Sql As String
            Sql = " SELECT * "
            Sql = Sql & vbCrLf & " FROM LEGAJOS "
            Sql = Sql & vbCrLf & " where ID_LEGAJO  BETWEEN " & txtDesde.Text
            Sql = Sql & vbCrLf & "  AND  " & txtHasta.Text
            Sql = Sql & vbCrLf & " ORDER BY ID_LEGAJO "
             frmReportes.ImprimirReporte PasoReportes & "rptEtiquetasChicasLegajosOK2.rpt", Sql, True
            
End Sub

Private Sub cmdIngresoNumero_Click()
    If InputBox("Ingrese la clave") = "21877471" Then
        txtDesde.Enabled = True
        txtHasta.Enabled = True
    End If
End Sub

Private Sub Command2_Click()

Dim ConEtiquetas As New ADODB.Connection
ConEtiquetas.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Etiquetas.mdb;Persist Security Info=False"

Dim rs As New ADODB.Recordset
Dim Sql As String
Dim C As Long
Dim Rollo As Integer
Dim DIGITOVERIFICADOR As Integer




Sql = " SELECT     ID_LEGAJO, ETIQUETA, DIGITO_VERIFICADOR, ROLLO"
Sql = Sql & "  From basasql.dbo.LEGAJOS"
Sql = Sql & "  Where (ROLLO > " & InputBox("Ingrese el rollo") & ")"
Sql = Sql & "  ORDER BY ROLLO, ID_LEGAJO DESC"


Rollo = 0
rs.Open Sql, strConBasa



Do While Not rs.EOF
            If Rollo <> rs!Rollo Then
                    Rollo = rs!Rollo
                    Sql = " SELECT  ID, ROLLO, ID_ETIQUETA, TEXT_ID_LEGAJO,ETIQUETA, DIGITO  INTO " & Rollo & " FROM ETIQUETA;"
                    ConEtiquetas.Execute Sql
             End If
             
             
                    
               Sql = " INSERT INTO " & Rollo
               Sql = Sql & "( ROLLO , "
               Sql = Sql & " TEXT_ID_LEGAJO , "
               Sql = Sql & " ETIQUETA, "
               Sql = Sql & "  DIGITO , "
               Sql = Sql & " ID_ETIQUETA )"
               Sql = Sql & " VALUES ("
               Sql = Sql & rs!Rollo & ", '"
               Sql = Sql & rs!ID_LEGAJO & "', "
               Sql = Sql & "'" & rs!Etiqueta & "', "
               Sql = Sql & rs!Digito_Verificador & ","
               Sql = Sql & rs!ID_LEGAJO & " )"
               ConEtiquetas.Execute Sql
               Debug.Print Rollo
            

        rs.MoveNext
        
        
        Loop
        
        MsgBox "Operacion terminada"




End Sub

Private Sub Command3_Click()
    Dim ConEtiquetas As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim Sql As String
        ConEtiquetas.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Etiquetas.mdb;Persist Security Info=False"
        Sql = " SELECT   ROLLO, MIN(ID_LEGAJO) AS Minimo, MAX(ID_LEGAJO) AS Maximo ,  (MIN(ID_LEGAJO) - MAX(ID_LEGAJO) ) as Cantidad"
        Sql = Sql & "  From basasql.dbo.LEGAJOS"
        Sql = Sql & "  GROUP BY ROLLO"
        Sql = Sql & "  Having (ROLLO > " & InputBox("Ingrese el rollo de inicio") & ")"
        Sql = Sql & "  ORDER BY ROLLO"
        Rollo = 0
        rs.Open Sql, strConBasa
        ConEtiquetas.Execute "DELETE * FROM ROLLOS "
        Do While Not rs.EOF
            Sql = " INSERT INTO ROLLOS ( ROLLO, DESDE, HASTA )"
            Sql = Sql & " VALUES (" & rs!Rollo & "," & rs!Minimo & "," & rs!Maximo & " )"
            ConEtiquetas.Execute Sql
            Sql = " INSERT INTO ROLLOS ( ROLLO, DESDE, HASTA )"
            Sql = Sql & " VALUES (" & rs!Rollo & "," & rs!Minimo & "," & rs!Maximo & " )"
            ConEtiquetas.Execute Sql
            rs.MoveNext
        Loop
        MsgBox "Operacion terminada"
End Sub



Private Sub cmdImprimirEtiquetas_Click()

        Dim Sql As String
        Dim Rollo As String
        Dim intROLLO As String
        Dim LEGAJO_1   As String
        Dim LEGAJO_2 As String
        Dim BARRA_1 As String
        Dim BARRA_2  As String
        Dim DIGITO_1 As String
        Dim DIGITO_2 As String
        Dim rs As New ADODB.Recordset
        Dim DATO As String
        
        DATO = Mid(txtLegajos.Text, 1, Len(txtLegajos.Text) - 1)
        
      Sql = "  SELECT   Digito_Verificador, ID_LEGAJO,ID_CLIENTE_LEGAJO, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION, NRO_CAJA,"
      Sql = Sql & vbCrLf & " COD_CLIENTE"
Sql = Sql & vbCrLf & " From LEGAJOS"
Sql = Sql & vbCrLf & " Where ID_LEGAJO in(" & DATO & ")"
  
  
  Dim Orden As Integer
  Dim Descripcion As String
  Orden = 1
  
   ExecutarSql "DELETE FROM TEM_LEGAJOS "
   
   rs.Open Sql, strConBasa, 0, 1
  Do While Not rs.EOF
    LEGAJO_1 = "'" & rs!ID_CLIENTE_LEGAJO & "'"
    BARRA_1 = "'" & "12" & Format(rs!ID_LEGAJO, "0000000000") & "'"
    If rs!ID_LEGAJO < 4794261 Then
        DIGITO_1 = "'" & rs!Digito_Verificador & "'"
    Else
        DIGITO_1 = "'" & DigitoEAN13("12" & Format(rs!ID_LEGAJO, "0000000000")) & "'"
    
    End If
    
    Descripcion = rs!NRO_DESDE & " " & Trim(rs!LETRA_DESDE) & " " & Format(rs!FECHA_DESDE, "yyyy") & " " & Trim(rs!Descripcion)
    Sql = " INSERT INTO TEM_LEGAJOS (ORDEN,  LEGAJO_1,  BARRA_1,  DIGITO_1, DESCRIPCION) "
    Sql = Sql & " VALUES     (" & Orden & "," & LEGAJO_1 & "," & BARRA_1 & "," & DIGITO_1 & ",'" & Descripcion & "') "
    ExecutarSql Sql
    Orden = Orden + 1
    rs.MoveNext
  Loop
     
 Sql = " SELECT * "
  Sql = Sql & "  From TEM_LEGAJOS "
 Sql = Sql & "  ORDER BY ORDEN"
 frmReportes.ImprimirReporte PasoReportes & "rpt_Etiquetas_Legajos_Ana.rpt", Sql, True

End Sub

Private Sub Command4_Click()
   Dim Sql As String
        Dim Rollo As String
        Dim intROLLO As String
        Dim LEGAJO_1   As String
        Dim LEGAJO_2 As String
        Dim BARRA_1 As String
        Dim BARRA_2  As String
        Dim DIGITO_1 As String
        Dim DIGITO_2 As String
        Dim rs As New ADODB.Recordset
        
        Dim DATO As String
        
        DATO = Mid(txtLegajos.Text, 1, Len(txtLegajos.Text) - 1)
        
      Sql = "  SELECT   Digito_Verificador, ID_LEGAJO,ID_CLIENTE_LEGAJO, LETRA_DESDE, LETRA_HASTA, NRO_DESDE, NRO_HASTA, FECHA_DESDE, FECHA_HASTA, DESCRIPCION, NRO_CAJA,"
      Sql = Sql & vbCrLf & " COD_CLIENTE"
Sql = Sql & vbCrLf & " From LEGAJOS"
Sql = Sql & vbCrLf & " Where ID_LEGAJO in(" & DATO & ")"
  
  
  Dim Orden As Integer
  Dim Descripcion As String
  Orden = 1
  
   ExecutarSql "DELETE FROM TEM_LEGAJOS "
   
   rs.Open Sql, strConBasa, 0, 1
  Do While Not rs.EOF
    LEGAJO_1 = "'" & rs!ID_CLIENTE_LEGAJO & "'"
    BARRA_1 = "'" & "12" & Format(rs!ID_LEGAJO, "0000000000") & "'"
    If rs!ID_LEGAJO < 4794261 Then
        DIGITO_1 = "'" & rs!Digito_Verificador & "'"
    Else
        DIGITO_1 = "'" & DigitoEAN13("12" & Format(rs!ID_LEGAJO, "0000000000")) & "'"
    
    End If
    
    Descripcion = rs!NRO_DESDE & " " & Trim(rs!LETRA_DESDE) & " " & Format(rs!FECHA_DESDE, "yyyy") & " " & Trim(rs!Descripcion)
    Sql = " INSERT INTO TEM_LEGAJOS (ORDEN,  LEGAJO_1,  BARRA_1,  DIGITO_1, DESCRIPCION) "
    Sql = Sql & " VALUES     (" & Orden & "," & LEGAJO_1 & "," & BARRA_1 & "," & DIGITO_1 & ",'" & Descripcion & "') "
    ExecutarSql Sql
    Orden = Orden + 1
    rs.MoveNext
  Loop
     
 Sql = " SELECT * "
  Sql = Sql & "  From TEM_LEGAJOS "
 Sql = Sql & "  ORDER BY ORDEN"
 frmReportes.ImprimirReporte PasoReportes & "rpt_Etiquetas_Legajos_Miguel_1.rpt", Sql, True

End Sub


Private Sub Form_Load()
ctlPersonal.TipoControl = Personal
End Sub

