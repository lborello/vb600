VERSION 5.00
Begin VB.Form frmCajasVacias 
   Caption         =   "Cajas Vacias"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   9960
   Begin VB.CommandButton cmdGeneracionCajasVaciasLectura 
      Caption         =   "Proceso por Lectura"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdGenerarCajasVacias 
      Caption         =   "Proceso por cajas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7920
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtDigito_Verificador 
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
      Left            =   6780
      TabIndex        =   4
      Top             =   540
      Width           =   555
   End
   Begin VB.TextBox txtCajaInicio 
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
      Left            =   5460
      TabIndex        =   3
      Top             =   540
      Width           =   1275
   End
   Begin VB.Label lbltipo 
      Caption         =   "Label6"
      Height          =   15
      Left            =   3120
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Requerimiento:"
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
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente:"
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
      Left            =   120
      TabIndex        =   11
      Top             =   180
      Width           =   1155
   End
   Begin VB.Label lblRequerimiento 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1500
      TabIndex        =   10
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label lbl_FK_Cliente 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   1500
      TabIndex        =   9
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label lblClienteRazonSocial 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblCliente"
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
      Left            =   2580
      TabIndex        =   8
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label lblCajaFin 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   8460
      TabIndex        =   6
      Top             =   540
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Caja Fin:"
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
      Left            =   7500
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblCantidadCajas 
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3540
      TabIndex        =   2
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Cantidad:"
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
      Left            =   2580
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Caja Inicio:"
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
      Left            =   4320
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmCajasVacias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub cmdGeneracionCajasVaciasLectura_Click()

On Error GoTo salir:
    Dim RsControlCajas As New ADODB.Recordset
    Dim SqlLectura As String
    Dim sql As String
    Dim CajasLectura As String
    Dim CANTIDAD As Integer
    Dim Lectura As New ADODB.Recordset
    
    
    SqlLectura = " SELECT     NUMERO_LECTURA, CAJA, CLIENTE"
    SqlLectura = SqlLectura & " From dbo.LECTURACOLECTOR"
    SqlLectura = SqlLectura & " Where Cliente = 0 "
    SqlLectura = SqlLectura & " and NUMERO_LECTURA = " & InputBox("Ingrese el N de Lectura", "", 0)
    RsControlCajas.Open SqlLectura, ConActiva, 0, 1
    
    Do While Not RsControlCajas.EOF
        CANTIDAD = CANTIDAD + 1
        CajasLectura = CajasLectura & "," & RsControlCajas!Caja
        RsControlCajas.MoveNext
    Loop
    
    CajasLectura = Mid(CajasLectura, 2)
    ' Control de Cajas
    sql = " SELECT ID_CAJA, FK_CLIENTE "
    sql = sql & " From dbo.Cajas"
    sql = sql & " WHERE  ID_CAJA in(" & CajasLectura & " )"
   sql = sql & " AND (NOT (FK_CLIENTE IS NULL)) "
    sql = sql & " ORDER BY ID_CAJA"
    Set RsControlCajas = New ADODB.Recordset
    RsControlCajas.Open sql, ConActiva, 0, 1
    
    
    If Not RsControlCajas.EOF Then
        MsgBox "Atencion las cajas ya estan en uso"
        Do While Not RsControlCajas.EOF
            MsgBox "La caja : " & RsControlCajas!ID_CAJA & " pertenece al cliente " & RsControlCajas!FK_CLIENTE
            RsControlCajas.MoveNext
        Loop
        Exit Sub
    End If
    
    
     Dim Caja As Long
     
     Dim rsContenedor As New ADODB.Recordset
        rsContenedor.CursorLocation = adUseClient
        sql = " SELECT  TOP " & CANTIDAD & " ID_CONTENEDOR, NRO_CAJA, COD_CLIENTE, ESTADO, F_MODIFICACION , FK_CAJAS "
        sql = sql & "  From CONTENEDOR "
        sql = sql & "  WHERE ESTADO = 1 AND COD_CLIENTE IS NULL AND"
        sql = sql & "  ESTANTERIA BETWEEN 150 AND 190 "
        rsContenedor.Open sql, strConBasa, adOpenKeyset, adLockPessimistic
    
    If rsContenedor.EOF Then
    
        MsgBox "No hay estanterias diponibles"
        Exit Sub
    
    End If
    
    sql = " SELECT ID_CAJA, FK_CLIENTE "
    sql = sql & " From dbo.Cajas"
    sql = sql & " WHERE  ( (FK_CLIENTE IS NULL)) "
    sql = sql & " AND ID_CAJA in(" & CajasLectura & " )"
    sql = sql & " ORDER BY ID_CAJA"
    Set RsControlCajas = New ADODB.Recordset
    RsControlCajas.Open sql, strConBasa, 0, 1
    Dim fechaModifi
    
    Do While Not RsControlCajas.EOF
        sql = " INSERT INTO dbo.REQUELIBOSCAJAS "
        sql = sql & " (IDREQUERIMIENTOS, CAJASLIBROS, FK_CAJAS)"
        sql = sql & "  VALUES (" & lblRequerimiento.Caption & "," & RsControlCajas!ID_CAJA & "," & RsControlCajas!ID_CAJA & ")"
        ExecutarSql sql
        rsContenedor!NRO_CAJA = RsControlCajas!ID_CAJA
        rsContenedor!FK_CAJAS = RsControlCajas!ID_CAJA
        rsContenedor!Cod_cliente = lbl_FK_Cliente.Caption
        If lbltipo.Caption = 26 Then
            rsContenedor!ESTADO = 5
        Else
            rsContenedor!ESTADO = 4
        End If
        rsContenedor!F_MODIFICACION = SysDate_DD_MM_YYYY
        rsContenedor.Update
        rsContenedor.MoveNext
        sql = " Update dbo.Cajas"
        sql = sql & " SET FK_CLIENTE = " & lbl_FK_Cliente.Caption & " ," & " NRO_CAJA =" & RsControlCajas!ID_CAJA
        sql = sql & " , FK_ESTADO =1100 "
        sql = sql & " Where ID_CAJA = " & RsControlCajas!ID_CAJA
        ExecutarSql sql
        RsControlCajas.MoveNext
    Loop
    
        sql = " Update Requerimiento "
        sql = sql & vbCrLf & "  SET IDESTADO =" & 3
        sql = sql & vbCrLf & "  Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
        ExecutarSql sql

        frmControlEstados.CargarTree
            Unload Me

     Exit Sub
salir:
    Rem conVacias.RollbackTrans
    MsgBox "Error en la generacion de cajas"

End Sub

Private Sub cmdGenerarCajasVacias_Click()


On Error GoTo salir:
    Dim RsControlCajas As New ADODB.Recordset
    Dim sql As String
    
    If Trim(txtDigito_Verificador.Text) = "" Then
    
             MsgBox "Ingrese el digito verificador"
             Exit Sub
        End If
    
    If (txtCajaInicio.Text) > 920404 Then
            If DigitoEAN13(CStr(110000000000# + txtCajaInicio.Text)) <> txtDigito_Verificador.Text Then
                MsgBox "El numero de caja No es el correcto ", vbCritical
                Exit Sub
            
            End If
   Else
            If Digito_Verificador(txtCajaInicio.Text) <> txtDigito_Verificador.Text Then
                MsgBox "El numero de caja No es el correcto ", vbCritical
                Exit Sub
            End If
    End If
    lblCajaFin.Caption = CLng(txtCajaInicio.Text) + CLng(lblCantidadcajas.Caption) - 1
    
'    If CLng(txtCajaInicio.Text) < 500000 Then
'        MsgBox "Error en Numero caja "
'        Exit Sub
'    End If
    
    ' Control de Cajas
    sql = " SELECT ID_CAJA, FK_CLIENTE "
    sql = sql & " From dbo.Cajas"
    sql = sql & " WHERE  (NOT (FK_CLIENTE IS NULL)) "
    sql = sql & " AND ID_CAJA BETWEEN " & txtCajaInicio.Text & " AND " & lblCajaFin.Caption
    sql = sql & " ORDER BY ID_CAJA"
    RsControlCajas.Open sql, ConActiva, 0, 1
    
    
    If Not RsControlCajas.EOF Then
        MsgBox "Atencion las cajas ya estan en uso"
        Do While Not RsControlCajas.EOF
            MsgBox "La caja : " & RsControlCajas!ID_CAJA & " pertenece al cliente " & RsControlCajas!FK_CLIENTE
            RsControlCajas.MoveNext
        Loop
        Exit Sub
    End If
    
    
     Dim Caja As Long
     
     Dim rsContenedor As New ADODB.Recordset
        rsContenedor.CursorLocation = adUseClient
        sql = " SELECT  TOP " & lblCantidadcajas.Caption & " ID_CONTENEDOR, NRO_CAJA, COD_CLIENTE, ESTADO, FK_CAJAS,F_MODIFICACION"
        sql = sql & "  From CONTENEDOR"
        sql = sql & "  WHERE ESTADO = 1 AND COD_CLIENTE IS NULL AND"
        sql = sql & "  ESTANTERIA BETWEEN 150 AND 190 "
        rsContenedor.Open sql, ConBasa, adOpenKeyset, adLockPessimistic
    
    
    For Caja = txtCajaInicio.Text To lblCajaFin.Caption
        sql = " INSERT INTO dbo.REQUELIBOSCAJAS "
        sql = sql & " (IDREQUERIMIENTOS, CAJASLIBROS, FK_CAJAS)"
        sql = sql & "  VALUES (" & lblRequerimiento.Caption & "," & Caja & "," & Caja & ")"
        ExecutarSql sql
        rsContenedor!NRO_CAJA = Caja
        rsContenedor!FK_CAJAS = Caja
        rsContenedor!Cod_cliente = lbl_FK_Cliente.Caption
       If lbltipo.Caption = 26 Then
            rsContenedor!ESTADO = 5
       Else
            rsContenedor!ESTADO = 4
       End If
        rsContenedor!F_MODIFICACION = SysDate_DD_MM_YYYY
        rsContenedor.Update
        rsContenedor.MoveNext
        sql = " Update dbo.Cajas"
        sql = sql & " SET FK_CLIENTE = " & lbl_FK_Cliente.Caption & " ," & " NRO_CAJA =" & Caja
        
        
        If lbltipo.Caption = 26 Then
            sql = sql & " , FK_ESTADO =1500 "
       Else
            sql = sql & " , FK_ESTADO =1100 "
        End If
        
        sql = sql & " Where ID_CAJA = " & Caja
        ExecutarSql sql
    Next
        sql = " Update Requerimiento "
        
        
        
       If lbltipo.Caption = 26 Then
       sql = sql & vbCrLf & "  SET IDESTADO =" & 6
       Else
        sql = sql & vbCrLf & "  SET IDESTADO =" & 3
        End If
  
        
        
        sql = sql & vbCrLf & "  Where IDREQUERIMIENTO = " & lblRequerimiento.Caption
        ExecutarSql sql
  
        frmControlEstados.CargarTree
        MsgBox "Operacion Completa", vbInformation
        Unload Me
     Exit Sub
salir:
    Rem conVacias.RollbackTrans
    MsgBox "Error en la generacion de cajas"
End Sub

Private Sub Form_Load()
    Dim rs As New ADODB.Recordset
    Dim sql As String
sql = " SELECT  dbo.REQUERIMIENTO.IDREQUERIMIENTO, dbo.REQUERIMIENTO.ID_CLIENTE, dbo.CLIENTES.RAZON_SOCIAL, dbo.REQUERIMIENTO.DESCRIPCION,"
sql = sql & " dbo.Requerimiento.CANTIDAD , IDTIPOREQUERIMIENTO "
sql = sql & " FROM         dbo.REQUERIMIENTO INNER JOIN"
sql = sql & " dbo.CLIENTES ON dbo.REQUERIMIENTO.ID_CLIENTE = dbo.CLIENTES.ID_CLIENTE"
sql = sql & " Where dbo.Requerimiento.IDREQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
rs.Open sql, ConActiva, 0, 1
If Not rs.EOF Then
    lbl_FK_Cliente.Caption = rs!ID_CLIENTE
    lblClienteRazonSocial.Caption = Trim(rs!Razon_Social)
    lblRequerimiento.Caption = rs!IDREQUERIMIENTO
    lblCantidadcajas.Caption = rs!CANTIDAD
    lbltipo.Caption = rs!IDTIPOREQUERIMIENTO
End If




End Sub

Private Sub txtCajaInicio_LostFocus()
If txtCajaInicio.Text <> "" Then
    lblCajaFin.Caption = CLng(txtCajaInicio.Text) + CLng(lblCantidadcajas.Caption) - 1
End If
End Sub
