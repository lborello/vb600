VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmRemitoDevolucion 
   Caption         =   "Remito"
   ClientHeight    =   6660
   ClientLeft      =   -2505
   ClientTop       =   7935
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   9900
      TabIndex        =   21
      Top             =   1260
      Width           =   435
   End
   Begin VB.Timer Timer1 
      Left            =   3720
      Top             =   7620
   End
   Begin VB.TextBox txtCaja 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   16
      Top             =   2640
      Width           =   1875
   End
   Begin VB.TextBox txtTomarLectura 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   60
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   6240
      Width           =   1200
   End
   Begin VB.Frame fraCliente 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   795
      Left            =   120
      TabIndex        =   3
      Top             =   1740
      Width           =   10155
      Begin VB.Label lblCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   300
         Width           =   5535
      End
      Begin VB.Label lblIDCliente 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   6240
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid grdCajasLibros 
      Height          =   3015
      Left            =   60
      TabIndex        =   1
      Top             =   3180
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   6
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MCI.MMControl MMControl5 
      Height          =   435
      Left            =   60
      TabIndex        =   20
      Top             =   7620
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   767
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   "C:\WINNT\Media\chimes.wav"
   End
   Begin VB.Label lblCantidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   8160
      TabIndex        =   19
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Cantidad : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6660
      TabIndex        =   18
      Top             =   2700
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Caja :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   2820
      TabIndex        =   17
      Top             =   2700
      Width           =   795
   End
   Begin VB.Label lblFecha 
      Caption         =   "10/10/2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   6540
      TabIndex        =   15
      Top             =   1260
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label lblRecibeNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   6540
      TabIndex        =   13
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label lblRecibe 
      Caption         =   "Recibe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   5280
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lblNumeroRemito 
      Caption         =   "JUAN PEREZ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   1260
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Remito : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label lblEntregaNombre 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   780
      Width           =   3615
   End
   Begin VB.Label lblEntrega 
      Caption         =   "Entrega: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   780
      Width           =   1095
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MOVIMIENTO DE DEVOLUCION DE CAJAS VACIAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   180
      TabIndex        =   7
      Top             =   120
      Width           =   10095
   End
End
Attribute VB_Name = "frmRemitoDevolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents mm As MMControl
Attribute mm.VB_VarHelpID = -1


Private Sub cmdAceptar_Click()
    If Validar Then
         Guardar_Remito
    End If
End Sub

Private Sub cmdCancelar_Click()
Dim a
a = 0

End Sub



Private Sub Form_Load()
    IdEntrega = 0
    idRecibe = 0
    IdClienteAnterior = 0

    lblFecha = Format(SysDate2, "dd/mm/yyyy")
   
    grdCajasLibros.ColWidth(0) = 100
    grdCajasLibros.ColWidth(1) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(2) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(3) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(4) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(5) = (grdCajasLibros.Width - 210) / 5
    
    grdCajasLibros.ColAlignment(1) = 4
    grdCajasLibros.ColAlignment(2) = 4
    grdCajasLibros.ColAlignment(3) = 4
    grdCajasLibros.ColAlignment(4) = 4
    grdCajasLibros.ColAlignment(5) = 4
    
    
    grdCajasLibros.TextMatrix(0, 1) = "CAJA"
    grdCajasLibros.TextMatrix(0, 2) = "CAJA"
    grdCajasLibros.TextMatrix(0, 3) = "CAJA"
    grdCajasLibros.TextMatrix(0, 4) = "CAJA"
    grdCajasLibros.TextMatrix(0, 5) = "CAJA"


End Sub
Function ProximoRemito() As Long
  Dim Sql As String
  Dim OraMax As OraDynaset
  Sql = "Select Max(Nro_Remito) Maximo From Remitos_Cuerpo"
  Set OraMax = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
  If IsNull(OraMax("Maximo")) Then ProximoRemito = 1: Exit Function
  ProximoRemito = Val(OraMax("Maximo")) + 1
End Function
Public Sub Guardar_Remito()
Dim Sql As String
Dim R As Integer
Dim c As Integer
Dim oradyn As OraDynaset
Dim Proximo_Nro_Remito As Long

On Error GoTo OraError

    If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
            Screen.MousePointer = 11
            OraSession.BeginTrans
            Proximo_Nro_Remito = ProximoRemito
            
            ' INSERTAR EN REMITO CUERPO
            Sql = "Insert into Remitos_Cuerpo (Nro_Remito,NRO_REM_PROV, Tipo, Operacion,"
            Sql = Sql & vbCrLf & " Estado, Fecha, Id_Cliente, Observaciones, Cantidad, "
            Sql = Sql & vbCrLf & " Audit_Usuario, Audit_Fecha, Fecha_Ingreso,Fecha_Error)"
            Sql = Sql & vbCrLf & " Values (" & CLng(Proximo_Nro_Remito) & ",'" & lblNumeroRemito & "',"       ' Nro Remito
            Sql = Sql & 4 & ","                  ' Tipo
            Sql = Sql & 0 & ","           ' Operacion
            Sql = Sql & vbCrLf & "  0," ' ESTADO
            Sql = Sql & SysDate3 & ","           ' Fecha
            Sql = Sql & CInt(lblIDCliente.Caption) & ","                 ' Id Cliente
            Sql = Sql & " ''" & ","                             ' Observaciones
            Sql = Sql & lblCantidad.Caption & ","                    ' Cantidad
            Sql = Sql & vbCrLf & "  '" & UCase(UserName$) & "',"                          ' Usuario
            Sql = Sql & SysDate & ","  ' Fecha y Hora
            Sql = Sql & SysDate & ","
            Sql = Sql & 0 & ")"
            Debug.Print Sql
            OraDatabase.ExecuteSQL Sql
            
                        
            'INSERTAR EN REMITO DELTALLE
            For R = 1 To grdCajasLibros.Rows - 1
               For c = 1 To grdCajasLibros.Cols - 1
                    If grdCajasLibros.TextMatrix(R, c) <> "" Then
                        Sql = "Insert into Remitos_Detalle(Nro_Remito, Desde, Hasta,"
                        Sql = Sql & vbCrLf & " Tipo_Almacenado, Detalle, Audit_Usuario, Audit_Fecha)"
                        Sql = Sql & vbCrLf & " Values (" + Format(Proximo_Nro_Remito) + ","
                        Sql = Sql & grdCajasLibros.TextMatrix(R, c) & ","
                        Sql = Sql & grdCajasLibros.TextMatrix(R, c) & ","
                        Sql = Sql & vbCrLf & 0 & ","
                        Sql = Sql & "'',"
                        Sql = Sql & " '" + UCase(UserName$) + "', "                                ' Usuario
                        Sql = Sql & SysDate & " )"
                        'MsgBox Sql
                        Debug.Print Sql
                        OraDatabase.ExecuteSQL Sql
                        GrabarMovHistorico Proximo_Nro_Remito, grdCajasLibros.TextMatrix(R, c), grdCajasLibros.TextMatrix(R, c), lblIDCliente, 0, 4, 0, SysDate
                    End If
                 Next
             Next
            
            'INSERTAR EN MOVIMIENTOS
            Sql = "Insert into Movimientos(Id_cliente, Fecha, Nro_Remito,"
            Sql = Sql & vbCrLf & " Tipo_Movim, Oper_Movim, Cantidad, Audit_Usuario,"
            Sql = Sql & vbCrLf & " Audit_Fecha)"
            Sql = Sql & vbCrLf & " Values ( " & CInt(lblIDCliente.Caption) & "," ' Id Cliente
            Sql = Sql & SysDate & "," ' Fecha
            Sql = Sql & CLng(Proximo_Nro_Remito) & ","  ' nro remito
            Sql = Sql & vbCrLf & "        " & 4 & "," ' Tipo
            Sql = Sql & 0 & ","  ' Operacion
            Sql = Sql & CLng(lblCantidad.Caption) & ","   ' Cantidad
            Sql = Sql & " '" + UCase(UserName$) + "',"                           ' Usuario
            Sql = Sql & vbCrLf & "        " & SysDate3 & ")"  ' Fecha de cargar
            Debug.Print Sql
             OraDatabase.ExecuteSQL Sql
            'MOVIMIENTO EN TABLA CONTENEDO
            For R = 1 To grdCajasLibros.Rows - 1
               For c = 1 To grdCajasLibros.Cols - 1
                   If grdCajasLibros.TextMatrix(R, c) <> "" Then
                        Sql = "UPDATE CONTENEDOR SET "
                        Sql = Sql & vbCrLf & " ESTADO = 1 "
                        Sql = Sql & vbCrLf & " , COD_CLIENTE = NULL "
                        Sql = Sql & vbCrLf & " , NRO_CAJA = NULL "
                        Sql = Sql & ", NRO_REMITO = " & Proximo_Nro_Remito
                        Sql = Sql & ", F_MODIFICACION = " & SysDate
                        Sql = Sql & vbCrLf & " WHERE "
                        Sql = Sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption)
                        Sql = Sql & " AND NRO_CAJA = " & CLng(grdCajasLibros.TextMatrix(R, c))
                        Sql = Sql & " AND ESTADO = 5 "
                        Debug.Print Sql
                        OraDatabase.ExecuteSQL Sql
                   End If
               Next
            Next
        OraSession.CommitTrans
        MsgBox "El remito fue grabado con exito", vbExclamation, "Remito"
        MsgBox "NUMERO DE MOVIMIENTO ES " & Proximo_Nro_Remito
        Screen.MousePointer = 0
        On Error GoTo ErrorPrn
    End If
Exit Sub
OraError:
    Screen.MousePointer = 0
    OraSession.Rollback
    frmLogOraError.Show MODAL
    Exit Sub
ErrorPrn:
    MsgBox Error
    Exit Sub
End Sub
Sub GrabarMovHistorico(mov_nrorem, mov_desde, mov_hasta, _
mov_cliente, mov_elem, mov_tipo, mov_oper, mov_fecha)
Dim R As Single
Dim Sql As String
Dim oradyn As OraDynaset
    
    Sql = "Select * from Mov_Cajas"
    Set oradyn = OraDatabase.CreateDynaset(Sql, ORADYN_DEFAULT)
    For R = mov_desde To mov_hasta
        oradyn.AddNew
        oradyn!NRO_REMITO = mov_nrorem
        oradyn!NRO_CAJA = R
        oradyn!ID_CLIENTE = mov_cliente
        oradyn!ELEMENTO = mov_elem
        oradyn!Tipo = mov_tipo
        oradyn!OPERACION = mov_oper
        oradyn!Fecha_Movimiento = Format(SysDate2, "DD/MM/YYYY")
        oradyn!Anulado = 0
        oradyn!AUDIT_USUARIO = UserName
        oradyn!AUDIT_FECHA = SysDate2
        oradyn.Update
    Next
    
End Sub

Public Sub TituloGrilla(Titulo)
    grdCajasLibros.ColWidth(0) = 100
    grdCajasLibros.ColWidth(1) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(2) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(3) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(4) = (grdCajasLibros.Width - 210) / 5
    grdCajasLibros.ColWidth(5) = (grdCajasLibros.Width - 210) / 5
    
    grdCajasLibros.ColAlignment(1) = 4
    grdCajasLibros.ColAlignment(2) = 4
    grdCajasLibros.ColAlignment(3) = 4
    grdCajasLibros.ColAlignment(4) = 4
    grdCajasLibros.ColAlignment(5) = 4
    
    
    grdCajasLibros.TextMatrix(0, 1) = Titulo
    grdCajasLibros.TextMatrix(0, 2) = Titulo
    grdCajasLibros.TextMatrix(0, 3) = Titulo
    grdCajasLibros.TextMatrix(0, 4) = Titulo
    grdCajasLibros.TextMatrix(0, 5) = Titulo
End Sub
Public Sub CargarGrilla(VALOR As String)
Dim c As Integer
Dim R As Integer
Dim RsEstadoCaja As OraDynaset
    
    
    
    Set RsEstadoCaja = OraDatabase.CreateDynaset("Select * from contenedor where cod_cliente= " & CInt(lblIDCliente) & " and nro_caja = " & VALOR, ORADYN_READONLY)
    

If Not RsEstadoCaja.EOF Then
            If CInt(RsEstadoCaja!Estado) <> 5 Then
                MsgBox "La caja " & VALOR & " NO TIENE UN ESTADO VALIDO"
                Exit Sub
            End If
    Else
        MsgBox "EL CLIENTE NO TIENE ESA CAJA"
        Exit Sub
    End If
    
    
    
    
    For R = 1 To grdCajasLibros.Rows - 1
        For c = 1 To grdCajasLibros.Cols - 1
            If grdCajasLibros.TextMatrix(R, c) = VALOR Then
                    Hablar "REPETIDA", MMControl5
                    Rem MsgBox "La Caja " & valor & " ya esta Cargada", vbInformation
                Exit Sub
            End If
            If grdCajasLibros.TextMatrix(R, c) = "" Then
                grdCajasLibros.TextMatrix(R, c) = VALOR
                   Hablar "ENTRADA", MMControl5
                Exit Sub
            End If
        Next
    Next
    grdCajasLibros.AddItem ""
    grdCajasLibros.TextMatrix(grdCajasLibros.Rows - 1, 1) = VALOR
    Hablar "ENTRADA", MMControl5
End Sub
Public Function Validar() As Boolean
 Dim RS As OraDynaset
 Dim R As Integer
 Dim c As Integer
 Dim Filtro As String
 Dim Sql As String
 Dim Bandera As Boolean
Dim I As Integer
 Bandera = False
   Validar = True
    For R = 1 To grdCajasLibros.Rows - 1
        For c = 1 To grdCajasLibros.Cols - 1
            If grdCajasLibros.TextMatrix(R, c) <> "" Then
                Filtro = Filtro & grdCajasLibros.TextMatrix(R, c) & ","
            End If
        Next
    Next
   Filtro = Mid(Filtro, 1, Len(Filtro) - 1)
   
            Sql = " SELECT ESTADO FROM CONTENEDOR WHERE "
            Sql = Sql & "COD_CLIENTE = " & CInt(lblIDCliente.Caption) 'CAJA
            Sql = Sql & " AND NRO_CAJA IN  (" & Filtro & ")"
   
  
   
    
   
   
   Set RS = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
        

                Do While Not RS.EOF
                    If CInt(RS!Estado) <> 5 Then
                        MsgBox "La Caja No tiene el estado Correcto"
                        Validar = False
                    End If
                RS.MoveNext
                Loop
        
        
    


End Function









Private Sub TXTcAJA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
          Dim cliente As Integer
          Dim Caja As Long
            If txtCaja <> "" And IsNumeric(txtCaja) Then
                 Caja = CInt(txtCaja)
                 CargarGrilla CLng(Caja)
            End If
            txtCaja = ""
            txtTomarLectura.SetFocus
   End If
End Sub

Private Sub txtTomarLectura_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then
             Select Case UCase(Mid(txtTomarLectura.Text, 1, 3))
             Case "R10"
                lblNumeroRemito = Mid(txtTomarLectura, 3)
             Case "P01"
                 Dim rsPersonal As OraDynaset
                 If lblEntregaNombre = "" Then
                     IdEntrega = CInt(Mid(txtTomarLectura, 4))
                     Set rsPersonal = OraDatabase.CreateDynaset("Select * from Personal where idpersonal =" & IdEntrega, ORADYN_READONLY)
                     If Not rsPersonal.EOF Then
                         lblEntregaNombre = UCase(CStr(rsPersonal!APELLIDO) & "  " & CStr(rsPersonal!NOMBRE))
                     End If
                 Else
                     idRecibe = CInt(Mid(txtTomarLectura, 4))
                     Set rsPersonal = OraDatabase.CreateDynaset("Select * from Personal where idpersonal =" & idRecibe, ORADYN_READONLY)
                     If Not rsPersonal.EOF Then
                         lblRecibeNombre = UCase(CStr(rsPersonal!APELLIDO) & "  " & CStr(rsPersonal!NOMBRE))
                     End If
                 End If
             Case Else
                Dim cliente As Integer
                Dim Caja As Long
                If txtTomarLectura <> "" And Len(txtTomarLectura) > 16 Then
                    Caja = Mid(txtTomarLectura.Text, Len(txtTomarLectura.Text) - 5)
                    cliente = Mid(txtTomarLectura.Text, Len(txtTomarLectura.Text) - 8, 3)
                     If IdClienteAnterior = cliente Or IdClienteAnterior = 0 Then
                        If IdClienteAnterior = 0 Then
                            IdClienteAnterior = cliente
                        End If
                        Dim RS As OraDynaset
                        Set RS = OraDatabase.CreateDynaset("SELECT * FROM CLIENTES WHERE ID_CLIENTE = " & cliente, ORADYN_READONLY)
                        If Not RS.EOF Then
                            lblIDCliente = RS!ID_CLIENTE
                            lblCliente = Trim(UCase(RS!RAZON_SOCIAL))
                        End If
                        CargarGrilla CLng(Caja)
                        lblCantidad = ContarGrilla
                      Else
                        Hablar "CLIENTE", MMControl5
                      End If
                End If
             End Select
             txtTomarLectura = ""
             txtTomarLectura.SetFocus
        End If

End Sub



Public Function ContarGrilla() As Integer
Dim I As Integer
Dim R As Integer
Dim c As Integer
 For R = 1 To grdCajasLibros.Rows - 1
        For c = 1 To grdCajasLibros.Cols - 1
            If grdCajasLibros.TextMatrix(R, c) <> "" Then
                I = I + 1
            End If
        Next
    Next
  ContarGrilla = I

End Function


