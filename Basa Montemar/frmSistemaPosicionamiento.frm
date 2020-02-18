VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmPosicionamiento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Posicionamiento"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "MS Sans Serif"
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
   ScaleHeight     =   5445
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
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
      Left            =   2460
      TabIndex        =   9
      Top             =   4920
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdPosicionamiento 
      Height          =   4275
      Left            =   180
      TabIndex        =   8
      Top             =   540
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   7541
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Controles.cltGenerico ctlResponsable 
      Height          =   375
      Left            =   1260
      TabIndex        =   7
      Top             =   60
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdColector 
      Caption         =   "Colector"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5700
      TabIndex        =   5
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4440
      TabIndex        =   4
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5700
      TabIndex        =   3
      Top             =   4920
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6960
      TabIndex        =   2
      Top             =   4920
      Width           =   1200
   End
   Begin VB.TextBox txtTomarLectura 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   4260
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo"
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
      Left            =   7020
      TabIndex        =   6
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label lblEntrega 
      Caption         =   "Responsable :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPosicionamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Public Sub Hablar(Data As String, MMControl5 As MMControl, Optional Paso As String)
'    MMControl5.Command = "close"
'    MMControl5.DeviceType = "WaveAudio"
'    If Paso = "" Then
'        MMControl5.FileName = "\\Server1basa\Sistemas\wav\" & Data & ".wav"
'    Else
'        MMControl5.FileName = "\\Server1basa\Sistemas\wav\" & Paso & Data & ".wav"
'    End If
'    MMControl5.Command = "open"
'    MMControl5.Command = "Prev"
'    MMControl5.Command = "Play"
'End Sub


Private Sub cmdAceptar_Click()
MousePointer = 11
    Dim sql1 As String
    Dim Sql As String
    Dim Filtro As String
    Dim R As Integer
    Dim Caja As Long
    Dim Cliente As Integer
    Dim Responsable As Integer
    If IsNull(ctlResponsable.Valor) Then
        MsgBox "Ingrese el responsable", vbInformation
        Exit Sub
    Else
        Responsable = ctlResponsable.Valor
    End If
    
    
    For R = 1 To grdPosicionamiento.Rows - 1
        Caja = grdPosicionamiento.TextMatrix(R, 1)
        Cliente = grdPosicionamiento.TextMatrix(R, 2)
        Filtro = Filtro & "( NRO_CAJA = " & Caja & " and Cod_Cliente = " & Cliente & ") OR "
        Sql = "INSERT INTO PRODUCION  (ID_PERSONAL, ID_TIPOTAREA, UNIDADPRODUCION,ELEMENTO, FECHA,COD_CLIENTE)"
        Sql = Sql & " VALUES (" & ctlResponsable.Valor & ", 1,1,'" & Caja & " - " & Cliente & "' ," & SysDate & "," & Cliente & ")"
        ExecutarSql Sql
        Sql = "  UPDATE CONTENEDOR"
        Sql = Sql & " Set COD_RESPONSABLE_POSICION = " & Responsable
        Sql = Sql & " Where Cod_Cliente = " & Cliente
        Sql = Sql & " And NRO_CAJA = " & Caja
        ExecutarSql Sql
    Next

Filtro = Mid(Filtro, 1, Len(Filtro) - 3)

    Sql = "  SELECT *"
    Sql = Sql & " From V_POSICIONAMIENTO"
    Sql = Sql & " Where " & Filtro

frmReportes.ImprimirReporte PasoReportes & "RptPosicionamiento.rpt", Sql, True

    TituloGrilla
    MousePointer = 0
End Sub

Private Sub cmdCancelar_Click()
    TituloGrilla
End Sub

Private Sub cmdColector_Click()
    Dim rs2 As New ADODB.Recordset
    Dim Sql As String
    Dim Lectura As Integer
    Dim Caja As Long
    Dim Cliente As Integer
    Dim fecha As String
    fecha = SysDate
   On Error GoTo salir
        Lectura = InputBox("Por Favor Ingrese el numero de Lectura ", "Lectura", 0)
        Sql = " SELECT NUMERO_LECTURA, CAJA, CLIENTE, ORDEN From LECTURACOLECTOR "
        Sql = Sql & "Where NUMERO_LECTURA = " & Lectura
        Sql = Sql & " ORDER BY ORDEN "
        rs2.Open Sql, ConActiva, 0, 1
        Do While Not rs2.EOF
           Caja = CLng(rs2!Caja)
            Cliente = CInt(rs2!Cliente)
            Sql = " INSERT INTO LECTURA_CAJAS (NRO_CAJA, COD_CLIENTE, ID_PERSONAL, FECHA) "
            Sql = Sql & " VALUES (" & Caja & "," & Cliente & "," & ctlResponsable.Valor & "," & fecha & ")"
            ExecutarSql (Sql)
            rs2.MoveNext
            CargarGrilla Cliente, Caja
       Loop
salir:

End Sub



Private Sub Form_Load()
    TituloGrilla
    ctlResponsable.TipoControl = PERSONAL
End Sub

Private Sub txtTomarLectura_KeyPress(KeyAscii As Integer)
   Dim rsPersonal As ADODB.Recordset
If KeyAscii = 13 Then
        If UCase(txtTomarLectura.Text) = "ACEPTAR" Then
            cmdAceptar_Click
            txtTomarLectura = ""
            Exit Sub
        End If
        If UCase(txtTomarLectura.Text) = "CANCELAR" Then
            cmdCancelar_Click
            txtTomarLectura = ""
            Exit Sub
        End If
        
        
             Select Case Len(txtTomarLectura)
               Case 17, 18 'caja
                    Dim Cliente As Integer
                    Dim Caja As Long
                    Dim Sql As String
                    Dim Personal1 As String
                    
                       If txtTomarLectura <> "" And Len(txtTomarLectura) > 16 Then
                        Caja = Mid(txtTomarLectura.Text, Len(txtTomarLectura.Text) - 5)
                        Cliente = Mid(txtTomarLectura.Text, Len(txtTomarLectura.Text) - 8, 3)
                        If IsNull(ctlResponsable.Valor) = "" Then
                           MsgBox "Atención Usted debe ingresar el responsable", vbCritical
                           Exit Sub
                        Else
                            
                            Sql = " INSERT INTO LECTURA_CAJAS (NRO_CAJA, COD_CLIENTE, ID_PERSONAL, FECHA , ID_PERSONAL1, TIPO_LECTURA) "
                            Sql = Sql & " VALUES (" & Caja & "," & Cliente & "," & lblIDPersonalEntrega & "," & SysDate & "," & ctlResponsable.Valor & " ,'POSICIONAMIENTO' )"
                            ExecutarSql (Sql)
                        End If
                        CargarGrilla Cliente, CLng(Caja)
                       End If
                Case 15 'Libro
'                    Tipo_Almacenado = 1
'                    TituloGrilla "Libros"
'                    lblCajaLibro = "Libro"
'                    Dim Libro As Long
'                    If txtTomarLectura <> "" Then
'                        Libro = Mid(txtTomarLectura.Text, 6, 5)
'                        Cliente = Mid(txtTomarLectura.Text, 11, 5)
'                        If IdClienteAnterior = Cliente Or IdClienteAnterior = 0 Then
'                            If IdClienteAnterior = 0 Then
'                                IdClienteAnterior = Cliente
'                            End If
'                            RS.Open "SELECT * FROM CLIENTES WHERE ID_CLIENTE = " & Cliente, strConBasa , 0 ,1
'                            If Not RS.EOF Then
'                                lblIDCliente = RS!id_cliente
'                                lblCliente = Trim(UCase(RS!razon_social))
'                            End If
'                            CargarGrilla CLng(Libro)
'                        Else
'                            Hablar "CLIENTE", MMControl5
'                        End If
'                    End If
                End Select

        txtTomarLectura = ""
        txtTomarLectura.SetFocus
   End If
End Sub


Public Sub InsertarLectura()

End Sub

Public Sub CargarGrilla(Cliente As Integer, Caja As Long)
    Dim R As Integer
    If grdPosicionamiento.Rows = 2 And grdPosicionamiento.TextMatrix(1, 1) = "" Then
        grdPosicionamiento.Rows = 1
    End If
    For R = 1 To grdPosicionamiento.Rows - 1
        If grdPosicionamiento.TextMatrix(R, 1) = CStr(Caja) And grdPosicionamiento.TextMatrix(R, 2) = CStr(Cliente) Then
            MsgBox "Caja Repetida"
            Exit Sub
        End If
     Next
     grdPosicionamiento.AddItem "" & vbTab & Caja & vbTab & Cliente
End Sub

Public Sub TituloGrilla()
    lblIDPersonalEntrega = ""
    lblEntregaNombre = ""
    
    grdPosicionamiento.Clear
    grdPosicionamiento.Rows = 2
    grdPosicionamiento.Cols = 4
    grdPosicionamiento.ColWidth(0) = 100
    grdPosicionamiento.ColWidth(1) = 1500
    grdPosicionamiento.ColWidth(2) = 1500
    grdPosicionamiento.ColWidth(3) = 3500
    
    grdPosicionamiento.TextMatrix(0, 1) = "Nro Caja"
    grdPosicionamiento.TextMatrix(0, 2) = "Cliente"
    grdPosicionamiento.TextMatrix(0, 3) = "Razon Social"
End Sub
