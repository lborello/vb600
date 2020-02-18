VERSION 5.00
Begin VB.Form frmEstanterias 
   Caption         =   "Estanterias"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4155
   ScaleWidth      =   4995
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   435
      Left            =   4380
      TabIndex        =   24
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   435
      Left            =   1680
      TabIndex        =   23
      Top             =   4740
      Width           =   1875
   End
   Begin VB.CommandButton cmdRotulosEstanteria 
      Caption         =   "Rotulos Estanteria"
      Height          =   615
      Left            =   2580
      TabIndex        =   22
      Top             =   6660
      Width           =   3135
   End
   Begin VB.CommandButton cmdActualizacion 
      Caption         =   "Actualizacion Lectura"
      Height          =   615
      Left            =   900
      TabIndex        =   21
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   435
      Left            =   5760
      TabIndex        =   20
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   435
      Left            =   3360
      TabIndex        =   19
      Top             =   5820
      Width           =   2715
   End
   Begin VB.ComboBox cboTipoCaja 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmEstanterias.frx":0000
      Left            =   1680
      List            =   "frmEstanterias.frx":000A
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   300
      Width           =   2475
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   4680
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   5220
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1740
      TabIndex        =   14
      Top             =   5340
      Width           =   1035
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   3420
      Width           =   1035
   End
   Begin VB.TextBox txtVerticalHasta 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3300
      TabIndex        =   10
      Text            =   "5"
      Top             =   2820
      Width           =   1035
   End
   Begin VB.TextBox txtHorinzontalHasta 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3300
      TabIndex        =   9
      Text            =   "22"
      Top             =   2400
      Width           =   1035
   End
   Begin VB.TextBox txtEstanteriaHasta 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3300
      TabIndex        =   8
      Text            =   "4"
      Top             =   1980
      Width           =   1035
   End
   Begin VB.TextBox txtEstado 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   7
      Text            =   "1"
      Top             =   3240
      Width           =   1035
   End
   Begin VB.TextBox txtVerticalDesde 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   5
      Text            =   "1"
      Top             =   2820
      Width           =   1035
   End
   Begin VB.TextBox txtHorinzontalDesde 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   3
      Text            =   "22"
      Top             =   2400
      Width           =   1035
   End
   Begin VB.TextBox txtEstanteriaDesde 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   1
      Text            =   "4"
      Top             =   1980
      Width           =   1035
   End
   Begin VB.Label Label7 
      Caption         =   "Tipo de Caja :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3300
      TabIndex        =   12
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   1035
   End
   Begin VB.Label Label4 
      Caption         =   "Estado :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Vertical :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2820
      Width           =   1515
   End
   Begin VB.Label Label2 
      Caption         =   "Horizontal :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Estanteria : "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1980
      Width           =   1515
   End
End
Attribute VB_Name = "frmEstanterias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdActualizacion_Click()
    Dim Sql As String
    Dim SqlLectura As String
    Dim disP As String
    Dim rsLectura As ADODB.Recordset
    Dim RsDisponibles As ADODB.Recordset
        SqlLectura = " SELECT     CAJA, CLIENTE"
        SqlLectura = SqlLectura & " From LECTURACOLECTOR"
        SqlLectura = SqlLectura & " Where (NUMERO_LECTURA = 10093)"
        SqlLectura = SqlLectura & " ORDER BY CAJA"
        
        Dim ConBasa As New ADODB.Connection
        ConBasa.Open strConBasa
        
    
    Set rsLectura = New ADODB.Recordset
    rsLectura.Open SqlLectura, ConActiva, 0, 1
    
    Do While Not rsLectura.EOF
        Sql = " Update CONTENEDOR"
        Sql = Sql & " SET ESTADO =1 "
        Sql = Sql & " , COD_CLIENTE = null "
        Sql = Sql & " , NRO_CAJA = null "
        Sql = Sql & " Where COD_CLIENTE = " & rsLectura!Cliente
        Sql = Sql & " And NRO_CAJA = " & rsLectura!Caja
        ExecutarSql Sql
        rsLectura.MoveNext
    Loop
 
    disP = "  SELECT  TOP 200 ESTADO, COD_CLIENTE, NRO_CAJA, ESTANTERIA, HORIZONTAL, "
    disP = disP & " VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE"
    disP = disP & " From CONTENEDOR "
    disP = disP & " Where ( Estanteria > 120 )"
    disP = disP & " And ( Estado = 1 )"
    disP = disP & " And ( COD_CLIENTE Is Null ) "
    disP = disP & " And ( NRO_CAJA Is Null ) "
    disP = disP & " ORDER BY ESTANTERIA "
    Set RsDisponibles = New ADODB.Recordset
    RsDisponibles.CursorLocation = adUseClient
  
    RsDisponibles.Open disP, ConActiva, adOpenDynamic, adLockPessimistic
    
     Set rsLectura = New ADODB.Recordset
    rsLectura.Open SqlLectura, ConActiva, 0, 1
    
    Do While Not rsLectura.EOF
        RsDisponibles!COD_CLIENTE = rsLectura!Cliente
        RsDisponibles!NRO_CAJA = rsLectura!Caja
        RsDisponibles!estado = 5
        RsDisponibles.Update
        RsDisponibles.MoveNext
        rsLectura.MoveNext
    Loop
    
    
    
    
    
    
    

End Sub

Private Sub cmdCrear_Click()
Dim Estanteria As Integer
Dim Horizontal As Integer
Dim Vertical As Integer
Dim NRO_ESTANTE  As Integer
Dim estado As Integer
Dim sSQL As String
Dim Modulo_V As String
Dim Modulo_H As String
Dim Modulo As Long
Dim Sql As String
Dim rs As New ADODB.Recordset

Rem  Inicio
    Dim ConBasa As New ADODB.Connection
    
    
    
  Rem  Inicio
        
        ConBasa.Open strConBasa
     
     For Estanteria = txtEstanteriaDesde To txtEstanteriaHasta
        For Vertical = txtVerticalDesde To txtVerticalHasta
            For Horizontal = txtHorinzontalDesde To txtHorinzontalHasta
                Select Case Horizontal
                Case 16, 17, 18
                    NRO_ESTANTE = 7
                Case 19, 20, 21
                    NRO_ESTANTE = 8
                End Select
                
                If cboTipoCaja.Text = "Chica" Then
                    estado = 1
                Else
                  estado = EstadoTipoCaja(Vertical, False)
                End If
                
                 estado = 1
                
                Select Case Vertical
                Case 1, 2, 3, 4, 5, 6, 7, 8
                    Modulo_V = 1
                Case 9, 10, 11, 12, 13, 14, 15, 16
                   Modulo_V = 2
                Case 17, 18, 19, 20, 21, 22, 23, 24
                   Modulo_V = 3
                Case 25, 26, 27, 28, 29, 30, 31, 32
                   Modulo_V = 4
                Case 33, 34, 35, 36, 37, 38, 39, 40
                   Modulo_V = 5
                Case 41, 42, 43, 44, 45, 46, 47, 48
                   Modulo_V = 6
                Case 49, 50, 51, 52, 53, 54, 55, 56
                   Modulo_V = 7
                Case 57, 58, 59, 60, 61, 62, 63, 64
                   Modulo_V = 8
                Case 65, 66, 67, 68, 69, 70, 71, 72
                   Modulo_V = 9
                Case 73, 74, 75, 76, 77, 78, 79, 80
                   Modulo_V = 10
                Case 81, 82, 83, 84, 85, 86, 87, 88
                   Modulo_V = 11
                Case 89, 90, 91, 92, 93, 94, 95, 96
                   Modulo_V = 12
                Case 97, 98, 99, 100, 101, 102, 103, 104
                   Modulo_V = 13
                 Case 105, 106, 107, 108, 109, 110, 111, 112
                   Modulo_V = 14
                 Case 113, 114, 115, 116, 117, 118, 119, 120
                   Modulo_V = 15
                 Case 121, 122, 123, 124, 125, 126, 127, 128
                   Modulo_V = 16
                 Case 129, 130, 131, 132, 133, 134, 135, 136
                   Modulo_V = 17
                 Case 137, 138, 139, 140, 141, 142, 143, 144
                   Modulo_V = 18
                 Case 145, 146, 147, 148, 149, 150, 151, 152
                    Modulo_V = 19
                 Case 153, 154, 155, 156, 157, 158, 159, 160
                    Modulo_V = 20
                 Case 161, 162, 163, 164, 165, 166, 167, 168
                    Modulo_V = 21
                 Case 169, 170, 171, 172, 173, 174, 175, 176
                    Modulo_V = 22
                  Case 177, 178, 179, 180, 181, 182, 183, 184
                    Modulo_V = 23
                  Case 185, 186, 187, 188, 189, 190, 191, 192
                    Modulo_V = 24
                  Case 193, 194, 195, 196, 197, 198, 199, 200
                    Modulo_V = 25
                  Case 201, 202, 203, 204, 205, 206, 207, 208
                    Modulo_V = 26
                  Case 209, 210, 211, 212, 213, 214, 215, 216
                    Modulo_V = 27
                    Case 217, 218, 219, 220, 221, 222, 223, 224
                    Modulo_V = 28
                End Select
                
                Select Case Horizontal
                Case 1, 2, 3, 4, 5
                    Modulo_H = 1
                Case 6, 7, 8, 9, 10
                    Modulo_H = 2
                Case 11, 12, 13, 14, 15
                    Modulo_H = 3
               Case 16, 17, 18, 19, 20
                    Modulo_H = 4
               Case 21, 22, 23, 24, 25
                    Modulo_H = 5
                End Select
                
'
                
                Set rs = New ADODB.Recordset
                
                Sql = "  SELECT     ESTANTERIA, VERTICAL, HORIZONTAL, ADELANTE_ATRAS"
                Sql = Sql & " From CONTENEDOR"
                Sql = Sql & " Where Estanteria = " & Estanteria
                Sql = Sql & " And Vertical = " & Vertical
                Sql = Sql & " And Horizontal = " & Horizontal

                
                rs.Open Sql, ConBasa, 0, 1
                
               Rem  Debug.Assert Estanteria <> 5272
                If rs.EOF Then
           
                    Modulo = Str(Estanteria) & Modulo_V & Modulo_H & 1
                    sSQL = "INSERT INTO CONTENEDOR "
                    sSQL = sSQL & vbCrLf & " (ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
                    sSQL = sSQL & vbCrLf & " NRO_ESTANTE, ESTADO, MODULO_V, MODULO_H,NRO_CAJA)"
                    sSQL = sSQL & vbCrLf & " VALUES (" & Estanteria & "," & Horizontal & "," & Vertical & ",1 ," & NRO_ESTANTE & "," & estado & "," & Modulo_V & "," & Modulo_H & "," & Estanteria & Horizontal & Vertical & " )"
                    ExecutarSql (sSQL)
                  End If
 
 
               Rem  Debug.Print "Estanteria : " & Estanteria & " Horizontal : " & Horizontal & " Vertical : " & Vertical & " Adelante = 1"
              
             Next
        Next
    Next
    MsgBox "terminado"
End Sub

Private Sub cmdRotulosEstanteria_Click()
    Dim Sql As String
    Dim rs As New ADODB.Recordset
    Dim conAcces As New ADODB.Connection
    conAcces.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\Serverbasa1\SistemasBasa\Etiquetas\ESTANTERIAS.mdb;Persist Security Info=False"
    

    Sql = " SELECT     ESTANTERIA, MODULO_V, MODULO_H"
    Sql = Sql & " From dbo.CONTENEDOR"
    Sql = Sql & " GROUP BY ESTANTERIA, MODULO_V, MODULO_H"
    Sql = Sql & " Having (Estanteria > 4999)"
    Sql = Sql & " ORDER BY ESTANTERIA, MODULO_V DESC, MODULO_H DESC"

conAcces.Execute "DELETE * FROM ESTANTERIA"
Rem Inicio


rs.Open Sql, ConActiva, 0, 1
Dim BARRA As String


Do While Not rs.EOF

    BARRA = "E" & Format(rs!Estanteria, "0000") & Format(rs!Modulo_V, "00") & Format(rs!Modulo_H, "00")
    Sql = " INSERT INTO ESTANTERIA "
    Sql = Sql & "( ESTANTERIA , V, H, BARRA )"
    Sql = Sql & " VALUES("
    Sql = Sql & rs!Estanteria & "," & rs!Modulo_V & "," & rs!Modulo_H & ",'" & BARRA & "')"
    conAcces.Execute Sql
    rs.MoveNext
Loop




Dim RsSql As ADODB.Recordset
'
'
'        Dim sql As String
'
'
'
'
'
'    Dim MaxLegajo As Integer
'    Set rs = New ADODB.Recordset
'
'        sql = " SELECT max(Legajos.id_Legajo) as MaxLegajo FROM Legajos "
'        rs.Open sql, conAcces
'
'   If Not rs.EOF Then
'        MaxLegajo = rs!MaxLegajo + 1
'   Else
'
'
'   End If


End Sub

Private Sub Command4_Click()







Dim rs As New ADODB.Recordset
Dim Sql As String
    Dim ConBasa As New ADODB.Connection
        ConBasa.Open strConBasa
Dim i As Long
rs.CursorLocation = adUseClient
Sql = " SELECT     ESTADO, NRO_CAJA, COD_CLIENTE, UB_PROVISORIA, ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS"
Sql = Sql & "  From CONTENEDOR"
Sql = Sql & "  WHERE     (COD_CLIENTE IS NULL) AND (ESTADO = 1) AND (ESTANTERIA BETWEEN 150 AND 200)"
Sql = Sql & "  ORDER BY ESTANTERIA"


rs.Open Sql, ConActiva, adOpenKeyset, adLockOptimistic
i = 38628


Do While Not rs.EOF
If i > 38657 Then
    Exit Sub
End If

    rs!NRO_CAJA = i
    rs!COD_CLIENTE = 20
    rs!estado = 5
    rs.Update
    i = i + 1
    rs.MoveNext
Loop



End Sub

Private Sub Command1_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String
 Dim ConBasa As New ADODB.Connection
  ConBasa.Open "Provider=MSDAORA.1;Password=1742;User ID=basa;Data Source=bpdc;Persist Security Info=True"

 rs.CursorLocation = adUseClient
 
 
 Sql = "  SELECT MODULO_V, MODULO_H, MODULO, ADELANTE_ATRAS,"
Sql = Sql & "    ESTANTERIA, ID_CONTENEDOR, HORIZONTAL,   Vertical"
Sql = Sql & " From CONTENEDOR"
Sql = Sql & " Where (Estanteria = 330) And (Not (Modulo_V Is Null))"
rs.Open Sql, ConActiva, 3, 2

Do While Not rs.EOF
   rs!Modulo = CLng(Str(rs!Estanteria) & rs!Modulo_V & rs!Modulo_H & rs!Adelante_Atras)
    ExecutarSql Sql
    rs.MoveNext
Loop


End Sub

Public Function EstadoTipoCaja(Vertical As Integer, TipoCajaChica As Boolean) As Integer
EstadoTipoCaja = 1
If TipoCajaChica = True Then
Else
Select Case Vertical
Case 7, 8, 15, 16, 23, 24, 32, 31, 39, 40, 48, 47, 55, 56, 64, 63
    EstadoTipoCaja = 0
End Select
End If

End Function

Private Sub Command3_Click()
 Dim ConBasa As New ADODB.Connection
        ConBasa.Open "Provider=MSDAORA.1;Password=1742;User ID=basa;Data Source=bpdc;Persist Security Info=True"
     Dim rs As New ADODB.Recordset

Dim Sql As String
    Sql = " SELECT CLIENTEUSUARIO.ID_CLIENTEUSUARIO,"
    Sql = Sql & " CLIENTEUSUARIO.COD_CLIENTE,"
    Sql = Sql & "  CLIENTEUSUARIO.APELLIDO_NOMBRE,"
    Sql = Sql & " CLIENTEUSUARIO.CORREO, CLIENTEUSUARIO.COD_INDICE,"
    Sql = Sql & "  CLIENTEUSUARIO.DOCUMENTO,   CLIENTEUSUARIO.TELEFONOS,"
    Sql = Sql & "  CLIENTEUSUARIO.REFERENCIAS , INDICES.Indice"
    Sql = Sql & "  From CLIENTEUSUARIO, INDICES"
    Sql = Sql & "  WHERE CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE AND"
    Sql = Sql & " CLIENTEUSUARIO.DOCUMENTO = INDICES.ID_CODIGO_DOCUMENTO"
    Sql = Sql & "  AND (CLIENTEUSUARIO.COD_CLIENTE = 04)"
    Sql = Sql & "  ORDER BY CLIENTEUSUARIO.APELLIDO_NOMBRE"
    
    rs.Open Sql, ConActiva, 0, 1
    
    
    Do While rs.EOF
        Sql = " Update CLIENTEUSUARIO "
        Sql = Sql & " SET COD_INDICE = " & rs!INDICES
        Sql = Sql & " Where ID_CLIENTEUSUARIO = " & rs!ID_CLIENTEUSUARIO
        Sql = Sql & "  And COD_CLIENTE = 4 "
        ExecutarSql Sql
        rs.MoveNext
    Loop
    
    
    
    
End Sub

Private Sub Command5_Click()
Dim Sql As String
Dim rs As New ADODB.Recordset
 Dim ConBasa As New ADODB.Connection
        ConBasa.Open strConBasa

Sql = " SELECT     CAMBIOPOSICION.ESTANTERIA, CAMBIOPOSICION.HORIZONTAL, CAMBIOPOSICION.VERTICAL, CAMBIOPOSICION.ADELANTE_ATRAS,"
Sql = Sql & " CAMBIOPOSICION.NRO_ESTANTE, CAMBIOPOSICION.ESTADO, CAMBIOPOSICION.COD_CLIENTE, CAMBIOPOSICION.NRO_CAJA,"
Sql = Sql & " CAMBIOPOSICION.Fecha , CAMBIOPOSICION.ID_Personal, LECTURACOLECTOR.NUMERO_LECTURA"
Sql = Sql & " FROM         CAMBIOPOSICION INNER JOIN"
Sql = Sql & " LECTURACOLECTOR ON CAMBIOPOSICION.NRO_CAJA = LECTURACOLECTOR.CAJA AND"
Sql = Sql & " CAMBIOPOSICION.COD_CLIENTE = LECTURACOLECTOR.Cliente"
Sql = Sql & "  WHERE     (CAMBIOPOSICION.FECHA = CONVERT(DATETIME, '2008-10-02 00:00:00', 102)) AND (LECTURACOLECTOR.NUMERO_LECTURA = 9605)"
Sql = Sql & "  ORDER BY CAMBIOPOSICION.COD_CLIENTE, CAMBIOPOSICION.NRO_CAJA"
 
 Dim i As Integer
 rs.Open Sql, ConActiva, 0, 1

Do While Not rs.EOF
'        Sql = " Update CONTENEDOR"
'        Sql = Sql & " SET ESTADO =" & rs!Estado
'        Sql = Sql & " , COD_CLIENTE =" & rs!COD_CLIENTE
'        Sql = Sql & ", NRO_CAJA =" & rs!NRO_CAJA
'        Sql = Sql & "  Where Estanteria = " & rs!Estanteria
'        Sql = Sql & " And Horizontal = " & rs!Horizontal
'        Sql = Sql & " And Vertical = " & rs!Vertical
'        Sql = Sql & " And Adelante_Atras =" & rs!Adelante_Atras
        ExecutarSql Sql

i = i + 1
    rs.MoveNext
Loop

MsgBox i

End Sub

Private Sub Command6_Click()

Dim rs As New ADODB.Recordset
Dim Sql As String
Dim V As Integer
Dim E As Integer

Dim h As Integer
Dim i As Integer

rs.CursorLocation = adUseClient


For E = 2000 To 2029
    For V = 1 To 17
        For h = 1 To 3
            Sql = " SELECT     ESTANTERIA, MODULO_V, MODULO_H, ID_CONTENEDOR, VERTICAL, HORIZONTAL, ORDEN"
            Sql = Sql & " From CONTENEDOR "
            Sql = Sql & " Where Estanteria = " & E & " And Modulo_V = " & V & " And Modulo_H = " & h
            Sql = Sql & " ORDER BY ID_CONTENEDOR "
            
            Set rs = New ADODB.Recordset
            
            rs.Open Sql, ConActiva, 2, 3
                i = 1
            Do While Not rs.EOF
                
                rs!Orden = i
                i = i + 1
                rs.Update
                rs.MoveNext
            Loop
            
            
    
        Next
    Next
Next

End Sub

Private Sub Command7_Click()
Dim Estanteria As Integer
Dim Horizontal As Integer
Dim Vertical As Integer
Dim NRO_ESTANTE  As Integer
Dim estado As Integer
Dim sSQL As String
Dim Modulo_V As String
Dim Modulo_H As String
Dim Modulo As Long
Dim rs As ADODB.Recordset
Dim Sql As String

   Dim ConBasa As New ADODB.Connection
        ConBasa.Open strConBasa
     
     For Estanteria = txtEstanteriaDesde To txtEstanteriaHasta
        For Vertical = txtVerticalDesde To txtVerticalHasta
            For Horizontal = txtHorinzontalDesde To txtHorinzontalHasta
                If cboTipoCaja.Text = "Chica" Then
                  estado = 1
                Else
                  estado = EstadoTipoCaja(Vertical, False)
                End If
                
                 estado = 1
                
                Select Case Vertical
                Case 1, 2, 3, 4, 5, 6, 7, 8
                    Modulo_V = 1
                Case 9, 10, 11, 12, 13, 14, 15, 16
                   Modulo_V = 2
                Case 17, 18, 19, 20, 21, 22, 23, 24
                   Modulo_V = 3
                Case 25, 26, 27, 28, 29, 30, 31, 32
                   Modulo_V = 4
                Case 33, 34, 35, 36, 37, 38, 39, 40
                   Modulo_V = 5
                Case 41, 42, 43, 44, 45, 46, 47, 48
                   Modulo_V = 6
                Case 49, 50, 51, 52, 53, 54, 55, 56
                   Modulo_V = 7
                Case 57, 58, 59, 60, 61, 62, 63, 64
                   Modulo_V = 8
                Case 65, 66, 67, 68, 69, 70, 71, 72
                   Modulo_V = 9
                Case 73, 74, 75, 76, 77, 78, 79, 80
                   Modulo_V = 10
                Case 81, 82, 83, 84, 85, 86, 87, 88
                   Modulo_V = 11
                Case 89, 90, 91, 92, 93, 94, 95, 96
                   Modulo_V = 12
                Case 97, 98, 99, 100, 101, 102, 103, 104
                   Modulo_V = 13
                 Case 105, 106, 107, 108, 109, 110, 111, 112
                   Modulo_V = 14
                 Case 113, 114, 115, 116, 117, 118, 119, 120
                   Modulo_V = 15
                 Case 121, 122, 123, 124, 125, 126, 127, 128
                   Modulo_V = 16
                 Case 129, 130, 131, 132, 133, 134, 135, 136
                   Modulo_V = 17
                 Case 137, 138, 139, 140, 141, 142, 143, 144
                   Modulo_V = 18
                 Case 145, 146, 147, 148, 149, 150, 151, 152
                    Modulo_V = 19
                 Case 153, 154, 155, 156, 157, 158, 159, 160
                    Modulo_V = 20
                 Case 161, 162, 163, 164, 165, 166, 167, 168
                    Modulo_V = 21
                 Case 169, 170, 171, 172, 173, 174, 175, 176
                    Modulo_V = 22
                  Case 177, 178, 179, 180, 181, 182, 183, 184
                    Modulo_V = 23
                  Case 185, 186, 187, 188, 189, 190, 191, 192
                    Modulo_V = 24
                  Case 193, 194, 195, 196, 197, 198, 199, 200
                    Modulo_V = 25
                  Case 201, 202, 203, 204, 205, 206, 207, 208
                    Modulo_V = 26
                End Select
               Select Case Horizontal
               Case 1, 2, 3, 4, 5
                    Modulo_H = 1
               Case 6, 7, 8, 9, 10
                    Modulo_H = 2
               Case 11, 12, 13, 14, 15
                    Modulo_H = 3
               Case 16, 17, 18, 19, 20
                    Modulo_H = 4
               Case 21, 22, 23, 24, 25
                    Modulo_H = 5
               End Select
               Set rs = New ADODB.Recordset
               
               Sql = " Select * FROM CONTENEDOR where "
               Sql = Sql & " Estanteria  = " & Estanteria
               Sql = Sql & " AND Horizontal  = " & Horizontal
               Sql = Sql & " AND Vertical  = " & Vertical
          
                rs.Open Sql, strConBasa
               
               If rs.EOF Then
                    Modulo = Str(Estanteria) & Modulo_V & Modulo_H & 2
                    sSQL = "INSERT INTO CONTENEDOR "
                    sSQL = sSQL & vbCrLf & " (ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS,"
                    sSQL = sSQL & vbCrLf & " NRO_ESTANTE, ESTADO,MODULO_V, MODULO_H)"
                    sSQL = sSQL & vbCrLf & " VALUES (" & Estanteria & "," & Horizontal & "," & Vertical & ", 2 ," & NRO_ESTANTE & "," & estado & "," & Modulo_V & "," & Modulo_H & ")"
                    ExecutarSql (sSQL)
               End If
            Next
        Next
    Next
End Sub

Private Sub Form_Load()
inicio
End Sub
