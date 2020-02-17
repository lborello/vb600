VERSION 5.00
Begin VB.Form frmActualizarAlsina 
   Caption         =   "Actualizar Alsina"
   ClientHeight    =   4305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   8325
   Begin VB.TextBox txtCantidadRegistros 
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdAcutlizacionBase 
      Caption         =   "Actualizacion"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.ComboBox cboPaso 
      Height          =   315
      ItemData        =   "frmActualizarAlsina.frx":0000
      Left            =   1560
      List            =   "frmActualizarAlsina.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Parar a los "
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblHecho 
      Caption         =   "Label3"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblCantidadTotal 
      Caption         =   "Label3"
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad TOtal"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Paso"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmActualizarAlsina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdAcutlizacionBase_Click()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim RsCust As New ADODB.Recordset
Dim rsContenedor As ADODB.Recordset
Dim rsContenedor25 As ADODB.Recordset
Dim conAlsinaSQL As New ADODB.Connection
Dim conAlsinaAccess As New ADODB.Connection

conAlsinaAccess.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & cboPaso.Text

sql = " SELECT   count(*) as cant  "
sql = sql & " From CONTENEDOR25 "
sql = sql & " WHERE Not (CONTENEDOR25.NRO_CAJA) Is Null AND (CONTENEDOR25.ACTUALIZACIONBASA) Is Null "


Set rsContenedor25 = New ADODB.Recordset
rsContenedor25.Open sql, conAlsinaAccess
Dim cant As Long

lblCantidadTotal.Caption = rsContenedor25!cant




sql = " SELECT     COD_CLIENTE, NRO_CAJA, ID_CONTENEDOR, EMPRESA, ACTUALIZACIONBASA "
sql = sql & " From CONTENEDOR25 "
sql = sql & " WHERE Not (CONTENEDOR25.NRO_CAJA) Is Null AND (CONTENEDOR25.ACTUALIZACIONBASA) Is Null "
sql = sql & " ORDER BY ID_CONTENEDOR "




Dim Cliente_Caja As Integer

Set rsContenedor25 = New ADODB.Recordset
rsContenedor25.Open sql, conAlsinaAccess
conAlsinaSQL.Open strConBasa
Do While Not rsContenedor25.EOF


If txtCantidadRegistros.Text < cant Then
conAlsinaAccess.Close
Exit Do
End If


        
       If rsContenedor25!NRO_CAJA < 100000 Then
             If rsContenedor25!EMPRESA = "VCUS" Or rsContenedor25!EMPRESA = "CUST" Then
                 sql = " SELECT ID_CAJA, FK_CLIENTE, DEPOSITO From Cajas "
                 sql = sql & "  Where ID_CAJA =  " & rsContenedor25!NRO_CAJA
             Else
                 sql = " SELECT ID_CAJA, FK_CLIENTE, DEPOSITO From Cajas "
                 sql = sql & "  Where NRO_CAJA =  " & rsContenedor25!NRO_CAJA
                 sql = sql & " AND  FK_CLIENTE = " & rsContenedor25!Cod_cliente
            End If
       Else
            sql = " SELECT ID_CAJA, FK_CLIENTE, DEPOSITO From Cajas "
            sql = sql & "  Where ID_CAJA =  " & rsContenedor25!NRO_CAJA
       End If

                    Set rs = New ADODB.Recordset
                    rs.Open sql, conAlsinaSQL, 0, 1
                    
                    If Not rs.EOF Then
                        If IsNull(rs!FK_CLIENTE) Then
                            GoTo Proximo:
                        Else
                            Cliente_Caja = rs!FK_CLIENTE
                        End If
                    End If
                        sql = " SELECT     ID_CONTENEDOR, COD_CLIENTE, NRO_CAJA , ESTADO "
                        sql = sql & "  From CONTENEDOR "
                        sql = sql & " Where  COD_CLIENTE = " & Cliente_Caja
                        sql = sql & " AND NRO_CAJA = " & rsContenedor25!NRO_CAJA
                       Set rsContenedor = New ADODB.Recordset
                        
                        rsContenedor.Open sql, conAlsinaSQL, 0, 1
                    If rsContenedor.EOF Then
                        sql = " Update CONTENEDOR "
                        sql = sql & "   SET ESTADO = 2"
                        sql = sql & "  , NRO_CAJA =" & rsContenedor25!NRO_CAJA
                        sql = sql & "  , COD_CLIENTE =" & Cliente_Caja
                        sql = sql & "  Where ID_CONTENEDOR = " & rsContenedor25!ID_CONTENEDOR
                        sql = sql & "  and COD_CLIENTE  is null "
                        conAlsinaSQL.Execute sql
                        sql = " Update Cajas "
                        sql = sql & " SET DEPOSITO = 'ALSINA'"
                        sql = sql & " Where FK_CLIENTE = " & Cliente_Caja
                        sql = sql & " And NRO_CAJA = " & rsContenedor25!NRO_CAJA
                        conAlsinaSQL.Execute sql
                        sql = " UPDATE CONTENEDOR25 SET ACTUALIZACIONBASA = '" & Format(Now, "DD/MM/YYYY") & "'"
                        sql = sql & " WHERE ID_CONTENEDOR=" & rsContenedor25!ID_CONTENEDOR
                        
                        
                    Else
                        sql = "INSERT INTO CAMBIOPOSICION "
                        sql = sql & vbCrLf & " (ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA, FECHA, ID_PERSONAL)"
                        sql = sql & vbCrLf & " SELECT ESTANTERIA, HORIZONTAL, VERTICAL, ADELANTE_ATRAS, NRO_ESTANTE, ESTADO, COD_CLIENTE, NRO_CAJA, '10/11/2011' AS FECHA, 99 AS PERSONAL"
                        sql = sql & vbCrLf & " From CONTENEDOR "
                        sql = sql & vbCrLf & " Where ID_CONTENEDOR = " & rsContenedor!ID_CONTENEDOR
                        conAlsinaSQL.Execute sql
                        sql = " Update CONTENEDOR "
                        sql = sql & vbCrLf & " SET COD_CLIENTE = NULL, NRO_CAJA = NULL, ESTADO = 1"
                        sql = sql & " Where COD_CLIENTE = " & Cliente_Caja
                        sql = sql & " And NRO_CAJA = " & rsContenedor25!NRO_CAJA
                        conAlsinaSQL.Execute sql
                        sql = "Update CONTENEDOR "
                        sql = sql & " SET  ESTADO =" & rsContenedor!ESTADO
                        sql = sql & " , NRO_CAJA =" & rsContenedor25!NRO_CAJA
                        sql = sql & " , COD_CLIENTE =" & Cliente_Caja
                        sql = sql & "  Where ID_CONTENEDOR = " & rsContenedor25!ID_CONTENEDOR
                        sql = sql & " and NRO_CAJA is null"
                        conAlsinaSQL.Execute sql
                        sql = " UPDATE CONTENEDOR25 SET ACTUALIZACIONBASA = '" & Format(Now, "DD/MM/YYYY") & "'"
                        sql = sql & " WHERE ID_CONTENEDOR=" & rsContenedor25!ID_CONTENEDOR
                        conAlsinaAccess.Execute sql
                    End If

      GoTo OK
      
Proximo:

OK:

cant = cant + 1
lblHecho.Caption = cant
lblHecho.Refresh
frmActualizarAlsina.Refresh
rsContenedor25.MoveNext
    


Loop



MsgBox "Terminado"



End Sub

Private Sub Form_Load()

End Sub
