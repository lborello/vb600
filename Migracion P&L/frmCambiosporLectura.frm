VERSION 5.00
Begin VB.Form frmCambiosporLectura 
   Caption         =   "Cambios Lecturas"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10545
   LinkTopic       =   "Form4"
   ScaleHeight     =   6105
   ScaleWidth      =   10545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCambioeEstadoAutomatico 
      Caption         =   "Cambio de estado Automatico"
      Height          =   615
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdCambioCliente 
      Caption         =   "Cambio Cliente"
      Height          =   435
      Left            =   8520
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.ComboBox cboClientes 
      Height          =   315
      Left            =   2940
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1080
      Width           =   5115
   End
   Begin VB.ComboBox cboEstados 
      Height          =   315
      ItemData        =   "frmCambiosporLectura.frx":0000
      Left            =   2880
      List            =   "frmCambiosporLectura.frx":0016
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   300
      Width           =   1695
   End
   Begin VB.CommandButton cmdEstado 
      Caption         =   "Cambio Estado"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   300
      Width           =   1455
   End
   Begin VB.TextBox txtLectura 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   300
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Lectura"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   855
   End
End
Attribute VB_Name = "frmCambiosporLectura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    Dim conestado As New ADODB.Connection



Dim strConAsp150 As String


    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
conestado.Open strConAsp150

        
        SQL = " SELECT lecturas.codigo, lecturaDetalle.elemento_id"
        SQL = SQL & " FROM            lecturas INNER JOIN "
        SQL = SQL & "  lecturaDetalle ON lecturas.id = lecturaDetalle.lectura_id"
        SQL = SQL & "  Where lecturas.CODIGO = " & txtLectura.Text
        SQL = SQL & "  ORDER BY lecturas.id DESC"
        rs.Open SQL, strConAsp150
        
        Do While Not rs.EOF
            SQL = " Update elementos"
            SQL = SQL & "  SET estado = '" & cboEstados.Text & "'"
            SQL = SQL & " Where ID = " & rs!elemento_id
            conestado.Execute SQL
            rs.MoveNext
        Loop


End Sub

Private Sub cmdCambioCliente_Click()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    
    Dim strConAsp150 As String
Dim conCliente As New ADODB.Connection

    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
conCliente.Open strConAsp150

Dim LecturaCodigo As String
Dim cliente As Long


Rem "Control que no existan legajos cargados"

        LecturaCodigo = txtLectura.Text
        SQL = " SELECT lecturas.codigo, lecturaDetalle.elemento_id, elementos.contenedor_id "
        SQL = SQL & " FROM lecturas INNER JOIN "
        SQL = SQL & " lecturaDetalle ON lecturas.id = lecturaDetalle.lectura_id INNER JOIN"
        SQL = SQL & " elementos ON lecturaDetalle.elemento_id = elementos.id"
        SQL = SQL & " Where lecturas.CODIGO = " & LecturaCodigo
        SQL = SQL & " And (Not (elementos.contenedor_id Is Null))"
        SQL = SQL & " ORDER BY lecturas.id DESC"
        Set rs = New ADODB.Recordset
        rs.Open SQL, strConAsp150
         If Not rs.EOF Then
            MsgBox " Se se puede cambiar el cliente existesn legajos cargados "
            Exit Sub
         End If
 

           
            
             SQL = " SELECT  lecturas.codigo, lecturaDetalle.elemento_id, referencia.id, lotereferencia.cliente_emp_id AS lotereferencia_cliente_emp_id, "
            SQL = SQL & " elementos.clienteEmp_id AS elementos_clienteEmp_id, referencia.clasificacion_documental_id, referencia.indice_individual, "
            SQL = SQL & " clasificacionDocumental.codigo AS Expr1, clasificacionDocumental.nombre, lotereferencia.id AS lotereferencia_ID "
            SQL = SQL & " FROM lecturas INNER JOIN "
            SQL = SQL & " lecturaDetalle ON lecturas.id = lecturaDetalle.lectura_id INNER JOIN "
            SQL = SQL & " referencia ON lecturaDetalle.elemento_id = referencia.elemento_id INNER JOIN "
            SQL = SQL & " lotereferencia ON referencia.lote_referencia_id = lotereferencia.id INNER JOIN "
            SQL = SQL & " elementos ON lecturaDetalle.elemento_id = elementos.id AND referencia.elemento_id = elementos.id INNER JOIN "
            SQL = SQL & " clasificacionDocumental ON referencia.clasificacion_documental_id = clasificacionDocumental.id "
            SQL = SQL & " Where lecturas.codigo = " & LecturaCodigo
            SQL = SQL & " And (clasificacionDocumental.codigo <> 10) "
            SQL = SQL & " ORDER BY lecturas.id DESC "
          
            
            Set rs = New ADODB.Recordset
            rs.Open SQL, strConAsp150
            If Not rs.EOF Then
                MsgBox " Se se puede cambiar el cliente existe referencia cargada "
                Exit Sub
            End If



            SQL = " SELECT  lecturas.codigo, lecturaDetalle.elemento_id, referencia.id AS referenciaid , lotereferencia.cliente_emp_id AS lotereferencia_cliente_emp_id, "
            SQL = SQL & " elementos.clienteEmp_id AS elementos_clienteEmp_id, referencia.clasificacion_documental_id, referencia.indice_individual, "
            SQL = SQL & " clasificacionDocumental.codigo AS clasificacionDocumental_codigo, clasificacionDocumental.nombre, lotereferencia.id AS lotereferencia_ID "
            SQL = SQL & " FROM lecturas INNER JOIN "
            SQL = SQL & " lecturaDetalle ON lecturas.id = lecturaDetalle.lectura_id INNER JOIN "
            SQL = SQL & " referencia ON lecturaDetalle.elemento_id = referencia.elemento_id INNER JOIN "
            SQL = SQL & " lotereferencia ON referencia.lote_referencia_id = lotereferencia.id INNER JOIN "
            SQL = SQL & " elementos ON lecturaDetalle.elemento_id = elementos.id AND referencia.elemento_id = elementos.id INNER JOIN "
            SQL = SQL & " clasificacionDocumental ON referencia.clasificacion_documental_id = clasificacionDocumental.id "
            SQL = SQL & " Where lecturas.codigo = " & LecturaCodigo
            SQL = SQL & " And (clasificacionDocumental.codigo = 10) "
            SQL = SQL & " ORDER BY lecturas.id DESC "
          
            Set rs = New ADODB.Recordset
            rs.Open SQL, strConAsp150
            Do While Not rs.EOF
               cliente = Mid(cboClientes.Text, 1, 6)
                SQL = " Update elementos Set clienteEmp_id = " & cliente
                SQL = SQL & " Where ID = " & rs!elemento_id
                 conCliente.Execute SQL
                
                SQL = " UPDATE  lotereferencia SET cliente_emp_id = " & cliente
                SQL = SQL & " Where ID = " & rs!lotereferencia_ID
                conCliente.Execute SQL
                
                 SQL = " UPDATE referencia Set "
                 SQL = SQL & " clasificacion_documental_id = " & BuscarIndiceNuevo(CInt(cliente), rs!clasificacionDocumental_codigo)
                 SQL = SQL & " Where ID = " & rs!referenciaid
                 conCliente.Execute SQL
                 
                rs.MoveNext
            Loop
            


End Sub

Private Sub cmdCambioeEstadoAutomatico_Click()
Dim rs As New ADODB.Recordset
    Dim SQL As String
   
    Dim conestado As New ADODB.Connection



Dim strConAsp150 As String


    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
conestado.Open strConAsp150

        
       
        
        SQL = " SELECT lecturaDetalle.elemento_id, COUNT(*) AS cantidad, elementos.estado"
        SQL = SQL & " FROM lecturas INNER JOIN"
        SQL = SQL & " lecturaDetalle ON lecturas.id = lecturaDetalle.lectura_id INNER JOIN"
        SQL = SQL & " lecturaDetalle AS lecturaDetalle_1 ON lecturaDetalle.elemento_id = lecturaDetalle_1.elemento_id INNER JOIN"
        SQL = SQL & " lecturas AS lecturas_1 ON lecturaDetalle_1.lectura_id = lecturas_1.id INNER JOIN"
        SQL = SQL & " elementos ON lecturaDetalle.elemento_id = elementos.id"
        SQL = SQL & " Where lecturas.codigo = " & InputBox("Lectura ")
        SQL = SQL & " GROUP BY lecturaDetalle.elemento_id, lecturaDetalle_1.codigoBarras, elementos.estado"
        SQL = SQL & " ORDER BY lecturaDetalle_1.codigoBarras"
        
        
        
        rs.Open SQL, strConAsp150
        
        
        
        Do While Not rs.EOF
         If rs!cantidad > 1 Then
        
            SQL = " Update elementos"
            SQL = SQL & "  SET estado = 'En Consulta'"
            SQL = SQL & " Where ID = " & rs!elemento_id
            conestado.Execute SQL
        Else
            SQL = " Update elementos"
            SQL = SQL & "  SET estado = 'En el Cliente'"
            SQL = SQL & " Where ID = " & rs!elemento_id
            conestado.Execute SQL
     
        End If
        
            rs.MoveNext
        
        Loop
        
        
        
        
End Sub

Private Sub cmdEstado_Click()
Dim rs As New ADODB.Recordset
    Dim SQL As String
   
    Dim conestado As New ADODB.Connection



Dim strConAsp150 As String


    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
conestado.Open strConAsp150

        
        SQL = " SELECT lecturas.codigo, lecturaDetalle.elemento_id"
        SQL = SQL & " FROM            lecturas INNER JOIN "
        SQL = SQL & "  lecturaDetalle ON lecturas.id = lecturaDetalle.lectura_id"
        SQL = SQL & "  Where lecturas.CODIGO = " & txtLectura.Text
        SQL = SQL & "  ORDER BY lecturas.id DESC"
        rs.Open SQL, strConAsp150
        
        Do While Not rs.EOF
            SQL = " Update elementos"
            SQL = SQL & "  SET estado = '" & Trim(cboEstados.Text) & "'"
            SQL = SQL & " Where ID = " & rs!elemento_id
            conestado.Execute SQL
            rs.MoveNext
        Loop
End Sub

Private Sub Form_Load()


Dim rs As New ADODB.Recordset
    Dim SQL As String
   
    Dim conestado As New ADODB.Connection



Dim strConAsp150 As String


    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"
conestado.Open strConAsp150
        
        SQL = " SELECT clientesEmp.id, clientesEmp.codigo, personas_juridicas.razonSocial, clientesEmp.empresa_id"
        SQL = SQL & " FROM            clientesEmp INNER JOIN"
        SQL = SQL & " personas_juridicas ON clientesEmp.razonSocial_id = personas_juridicas.id"
        SQL = SQL & " Where (clientesEmp.empresa_id = 20004)"
        SQL = SQL & " Order By  personas_juridicas.razonSocial"
        rs.Open SQL, strConAsp150
        
        Do While Not rs.EOF
            cboClientes.AddItem Format(rs!ID, "000000") & " " & Format(rs!codigo, "00000") & " " & Trim(rs!razonSocial)
            rs.MoveNext
        Loop


End Sub

Public Function BuscarIndiceNuevo(cliente As Integer, codigo As String) As Long
  Dim SQL As String
  Dim rs As New ADODB.Recordset
  
Dim strConAsp150 As String


    strConAsp150 = "Provider=SQLOLEDB.1;Password=Sicuyo123;Persist Security Info=True;User ID=sa;Initial Catalog=basa;Data Source=222.15.19.150"


SQL = "  SELECT        id, cliente_asp_id, cliente_emp_id, padre_id, codigo"
SQL = SQL & " From clasificacionDocumental"
SQL = SQL & " Where codigo = " & codigo
SQL = SQL & " And cliente_emp_id = " & cliente

rs.Open SQL, strConAsp150

  BuscarIndiceNuevo = rs!ID

End Function
