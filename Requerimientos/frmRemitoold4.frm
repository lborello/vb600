VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRemitoViejo 
   Caption         =   "Remito"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9990
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9990
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7380
      TabIndex        =   23
      Top             =   6780
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
      Left            =   180
      TabIndex        =   20
      Top             =   60
      Width           =   6555
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
         Left            =   960
         TabIndex        =   22
         Top             =   360
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
         TabIndex        =   21
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Requerimiento"
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
      Left            =   6840
      TabIndex        =   18
      Top             =   60
      Width           =   3075
      Begin VB.Label lbRequerimiento 
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
         Height          =   435
         Left            =   240
         TabIndex        =   19
         Top             =   300
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   8700
      TabIndex        =   16
      Top             =   6780
      Width           =   1200
   End
   Begin VB.TextBox txtObservaciones 
      Height          =   735
      Left            =   60
      TabIndex        =   15
      Top             =   5940
      Width           =   9855
   End
   Begin VB.ListBox lstPersonal 
      Height          =   2535
      ItemData        =   "frmRemitoold4.frx":0000
      Left            =   6900
      List            =   "frmRemitoold4.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Frame fraRemito 
      Caption         =   "Remito"
      Enabled         =   0   'False
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
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   9795
      Begin VB.ComboBox cboTipo_Almacenado 
         Height          =   315
         ItemData        =   "frmRemitoold4.frx":0004
         Left            =   1560
         List            =   "frmRemitoold4.frx":0011
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cboTipoRemito 
         Height          =   315
         ItemData        =   "frmRemitoold4.frx":002A
         Left            =   1560
         List            =   "frmRemitoold4.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboRemito_Estados 
         Height          =   315
         ItemData        =   "frmRemitoold4.frx":0070
         Left            =   4920
         List            =   "frmRemitoold4.frx":007A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1755
      End
      Begin VB.ComboBox cboRemito_Operacion 
         Height          =   315
         ItemData        =   "frmRemitoold4.frx":008F
         Left            =   8100
         List            =   "frmRemitoold4.frx":0099
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1515
      End
      Begin MSMask.MaskEdBox maskFechaRemito 
         Height          =   330
         Left            =   8100
         TabIndex        =   14
         Top             =   840
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6840
         TabIndex        =   13
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblCantidad 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5340
         TabIndex        =   12
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label Label5 
         Caption         =   "Almac."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   4020
         TabIndex        =   9
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Remito :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   4020
         TabIndex        =   6
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Operación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   6840
         TabIndex        =   5
         Top             =   420
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdCajasLibros 
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   2460
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4577
      _Version        =   393216
      Cols            =   6
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox cryRemito 
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   24
      Top             =   0
      Width           =   1200
   End
   Begin VB.Label lblDescripcionRequerimiento 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   60
      TabIndex        =   17
      Top             =   5100
      Width           =   9855
   End
End
Attribute VB_Name = "frmRemitoViejo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAceptar_Click()
    If Validar Then
         Guardar_Remito
    End If
End Sub

Public Sub CambioEstadoRemito(IDEmpleado As Integer, ActualizaContador As Boolean, Optional EstadoInicial As Integer, Optional EstadoFinal As Integer, Optional Requerimiento As Long)
    Dim RS As ADODB.Recordset
    Dim RSH_ESTADO_REQUE As ADODB.Recordset
    Dim Sql As String
    Dim FECHARECEPCION As Date
    Dim IDTIPOREQUERIMIENTO As Integer
    Dim I As Integer
    Dim CONTADOR As Integer
    
   
    
            
            ' REQUERIMIENTO
            Sql = " UPDATE REQUERIMIENTO SET "
            Sql = Sql & vbCrLf & " IDESTADO= " & EstadoFinal
            Sql = Sql & vbCrLf & ", IDPERSONAL = " & IDEmpleado
            Sql = Sql & vbCrLf & " WHERE idRequerimiento IN  " & Requerimiento
            Sql = Sql & vbCrLf & " AND IDESTADO = " & EstadoInicial
            ConBasa.Execute (Sql)
            
            ' CONTADOR
            Sql = " SELECT max(Contador)AS CONTADOR From  H_ESTADO_REQUE  Where IDRequerimiento = " & Requerimiento
            Set RSH_ESTADO_REQUE = New ADODB.Recordset
            RSH_ESTADO_REQUE.Open Sql, ConBasa
            If Not RSH_ESTADO_REQUE.EOF Then
            If IsNull(RSH_ESTADO_REQUE!CONTADOR) Then
                CONTADOR = 1
                Else
                    If ActualizaContador Then
                        CONTADOR = CInt(RSH_ESTADO_REQUE!CONTADOR) + 1
                    Else
                        CONTADOR = CInt(RSH_ESTADO_REQUE!CONTADOR)
                    End If
                End If
            Else
                CONTADOR = 1
            End If
            
            ' H_ESTADO_REQUE
            Sql = " INSERT INTO H_ESTADO_REQUE ("
            Sql = Sql & vbCrLf & " IDREQUERIMIENTO, IDESTADO, IDPERSONAL,"
            Sql = Sql & vbCrLf & " CONTADOR, FECHA )"
            Sql = Sql & vbCrLf & "  VALUES ("
            Sql = Sql & vbCrLf & Requerimiento & "," & EstadoFinal & "," & IDEmpleado & ","
            Sql = Sql & vbCrLf & CONTADOR & "," & SysDate & ")"
            ConBasa.Execute (Sql)
   
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim rsPersonal As ADODB.Recordset
    Set rsPersonal = New ADODB.Recordset
    rsPersonal.Open "Select * from Personal WHERE NAVES=1 ", ConBasa
    Do While Not rsPersonal.EOF
        lstPersonal.AddItem CStr(rsPersonal!IDPERSONAL) & " - " & Trim(CStr(rsPersonal!NOMBRE)) & " " & Trim(CStr(rsPersonal!APELLIDO))
        rsPersonal.MoveNext
    Loop
    CargarRemito
End Sub
Function ProximoRemito() As Long
  Dim Sql As String
  Dim OraMax As ADODB.Recordset
  Sql = "Select Max(Nro_Remito) Maximo From Remitos_Cuerpo"
  Set OraMax = New ADODB.Recordset
   OraMax.Open Sql, ConBasa
  If IsNull(OraMax("Maximo")) Then ProximoRemito = 1: Exit Function
  ProximoRemito = Val(OraMax("Maximo")) + 1
End Function
Sub GrabarMovHistorico(NRO_REMITO As Long, NRO_CAJA As Long, ID_CLIENTE As Integer, ELEMENTO As Long, TIPO As Integer, OPERACION As Integer, FECHA_MOVIMIENTO As String, TIPO_ELEMENTO As Integer, AUDIT_USUARIO As String, AUDIT_FECHA As String)
    Dim R As Single
    Dim Sql As String
    Sql = " INSERT INTO MOV_CAJAS2 "
    Sql = Sql & vbCrLf & "(NRO_REMITO, NRO_CAJA, ID_CLIENTE, ELEMENTO, TIPO,"
    Sql = Sql & vbCrLf & " OPERACION, FECHA_MOVIMIENTO, TIPO_ELEMENTO,"
    Sql = Sql & vbCrLf & " AUDIT_USUARIO, AUDIT_FECHA)"
    Sql = Sql & vbCrLf & " VALUES (" & NRO_REMITO & "," & NRO_CAJA & "," & ID_CLIENTE & "," & ELEMENTO & "," & TIPO & ","
    Sql = Sql & vbCrLf & OPERACION & "," & FECHA_MOVIMIENTO & "," & TIPO_ELEMENTO & ","
    Sql = Sql & vbCrLf & AUDIT_USUARIO & "," & AUDIT_FECHA & ")"
    ConBasa.Execute Sql

End Sub

Public Sub CargarRemito()
    Dim Sql As String
    Dim rsRequerimiento As ADODB.Recordset
    Dim rsRcajas As ADODB.Recordset
        Sql = "SELECT REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.SECTOR,"
        Sql = Sql & vbCrLf & " REQUERIMIENTO.TELEFONO, REQUERIMIENTO.ID_CLIENTE,"
        Sql = Sql & vbCrLf & " REQUERIMIENTO.DESCRIPCION, REQUERIMIENTO.SOLICITANTE, REQUERIMIENTO.TOMO,"
        Sql = Sql & vbCrLf & " REQUERIMIENTO.FECHALIMITE, "
        Sql = Sql & vbCrLf & " REQUERIMIENTO.FECHARECEPCION, REQUERIMIENTO.IDTIPORECEPCION, "
        Sql = Sql & vbCrLf & " REQUERIMIENTO.CANTIDAD, REQUERIMIENTO.IDESTADO,"
        Sql = Sql & vbCrLf & " REQUERIMIENTO.IDTIPOREQUERIMIENTO, Clientes.razon_social "
        Sql = Sql & vbCrLf & " From REQUERIMIENTO , Clientes"
        Sql = Sql & vbCrLf & " WHERE "
        Sql = Sql & vbCrLf & " REQUERIMIENTO.id_Cliente = Clientes.ID_Cliente and "
        Sql = Sql & vbCrLf & " REQUERIMIENTO.IDREQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
       
       Set rsRequerimiento = New ADODB.Recordset
      rsRequerimiento.Open Sql, ConBasa
    If Not rsRequerimiento.EOF Then
       With rsRequerimiento
           Select Case !IDTIPOREQUERIMIENTO
           Case 1, 8
                cboTipoRemito.ListIndex = 1
                cboRemito_Operacion.ListIndex = 1
                cboRemito_Estados.ListIndex = 0
                cboTipo_Almacenado.ListIndex = 0
                TituloGrilla "Cajas"
           Case 3
                cboTipoRemito.ListIndex = 1
                cboRemito_Operacion.ListIndex = 1
                cboRemito_Estados.ListIndex = 1
                cboTipo_Almacenado.ListIndex = 0
                TituloGrilla "Cajas"
           Case 2
                cboTipoRemito.ListIndex = 1
                cboRemito_Operacion.ListIndex = 1
                cboRemito_Estados.ListIndex = 0
                cboTipo_Almacenado.ListIndex = 1
                TituloGrilla "Libros"
           Case 4
                cboTipoRemito.ListIndex = 1
                cboRemito_Operacion.ListIndex = 1
                cboRemito_Estados.ListIndex = 1
                cboTipo_Almacenado.ListIndex = 1
                TituloGrilla "Libros"
           Case 7
                cboTipoRemito.ListIndex = 2
                cboRemito_Operacion.ListIndex = 1
                cboRemito_Estados.ListIndex = 0
                cboTipo_Almacenado.ListIndex = 0
                TituloGrilla "Cajas"
           Case 10
                cboTipoRemito.ListIndex = 1
                cboRemito_Operacion.ListIndex = 1
                cboRemito_Estados.ListIndex = 0
                cboTipo_Almacenado.ListIndex = 2
                TituloGrilla "LEGAJO"
           Case 11
                cboTipoRemito.ListIndex = 1
                cboRemito_Operacion.ListIndex = 1
                cboRemito_Estados.ListIndex = 1
                cboTipo_Almacenado.ListIndex = 2
                TituloGrilla "LEGAJO"
           End Select
            maskFechaRemito.Text = Format(SysDateCompare, "dd/mm/yyyy")
            lblCliente.Caption = Trim(UCase(!Razon_Social))
            lblIDCliente.Caption = !ID_CLIENTE
            lbRequerimiento.Caption = UCase(!IDREQUERIMIENTO)
            lblCantidad.Caption = CInt(!CANTIDAD)
            If Not IsNull(!DESCRIPCION) Then
                lblDescripcionRequerimiento = UCase(!DESCRIPCION)
            End If
                Sql = " SELECT "
                Sql = Sql & vbCrLf & " REQ.IDRequerimientos , REQ.CAJASLIBROS"
                Sql = Sql & vbCrLf & " From"
                Sql = Sql & vbCrLf & " REQUELIBOSCAJAS REQ"
                Sql = Sql & vbCrLf & " Where REQ.IDRequerimientos = " & CRequerimientos.Item(1).NumeroRequerimiento
                Sql = Sql & vbCrLf & " ORDER BY REQ.CAJASLIBROS"
               
                   Set rsRcajas = New ADODB.Recordset
                  rsRcajas.Open Sql, ConBasa
                   
           Do While Not rsRcajas.EOF
                If Not IsNull(rsRcajas!CAJASLIBROS) Then
                     CargarGrilla CStr(rsRcajas!CAJASLIBROS)
                End If
                rsRcajas.MoveNext
           Loop
        End With
    End If
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
Public Sub CargarGrilla(Valor As String)
Dim c As Integer
Dim R As Integer


For R = 1 To grdCajasLibros.Rows - 1

    For c = 1 To grdCajasLibros.Cols - 1
        
        If grdCajasLibros.TextMatrix(R, c) = Valor Then
            MsgBox "La Caja " & Valor & " ya esta Cargada", vbInformation
            Exit Sub
        End If
        If grdCajasLibros.TextMatrix(R, c) = "" Then
            grdCajasLibros.TextMatrix(R, c) = Valor
            Exit Sub
        End If
    Next
Next
grdCajasLibros.AddItem ""
grdCajasLibros.TextMatrix(grdCajasLibros.Rows - 1, 1) = Valor
End Sub
Public Sub ImprimirRemito(NumeroRemito As Long)
    Dim Sql As String
    Dim SQL1 As String
    Dim Bandera As Boolean
    Dim Responsables As String
    Dim RS  As ADODB.Recordset
    Dim ANTERIOR As Long
    Dim NombreReporte As String
   On Error GoTo LERROR
    
    If NumeroRemito = 0 Or IsNull(NumeroRemito) Then
        Exit Sub
    End If
    
    cryRemito.Connect = "DSN = bpdc;UID = " & UserName & ";PWD = " & Password
    
    Select Case CRequerimientos.Item(1).TIPO
    Case 1, 3, 7, 8  ' CAJAS
        Sql = "    SELECT"
        Sql = Sql & vbCrLf & "   REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.OBSERVACIONES,"
        Sql = Sql & vbCrLf & "   REMITOS_DETALLE.DESDE,"
        Sql = Sql & vbCrLf & "   REQUERIMIENTO.SECTOR, REQUERIMIENTO.SOLICITANTE, REQUERIMIENTO.FECHARECEPCION,"
        Sql = Sql & vbCrLf & "   REMITO_TIPO.DESCRIPCION,"
        Sql = Sql & vbCrLf & "   REMITO_OPERACION.DESCRIPCION,"
        Sql = Sql & vbCrLf & "   REMITO_ESTADOS.DESCRIPCION,"
        Sql = Sql & vbCrLf & "   clientes.id_cliente , clientes.RAZON_SOCIAL, clientes.CALLE, clientes.NUMERO, clientes.LOCALIDAD"
        Sql = Sql & vbCrLf & "   From "
        Sql = Sql & vbCrLf & "   BASA.REMITOS_CUERPO REMITOS_CUERPO,"
        Sql = Sql & vbCrLf & "   BASA.REMITOS_DETALLE REMITOS_DETALLE,"
        Sql = Sql & vbCrLf & "   BASA.REQUERIMIENTO REQUERIMIENTO,"
        Sql = Sql & vbCrLf & "   BASA.REMITO_TIPO REMITO_TIPO,"
        Sql = Sql & vbCrLf & "   BASA.REMITO_OPERACION REMITO_OPERACION,"
        Sql = Sql & vbCrLf & "   BASA.REMITO_ESTADOS REMITO_ESTADOS,"
        Sql = Sql & vbCrLf & "   BASA.clientes clientes"
        Sql = Sql & vbCrLf & " Where"
        Sql = Sql & vbCrLf & "   REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO AND"
        Sql = Sql & vbCrLf & "   REMITOS_CUERPO.NRO_REMITO = REQUERIMIENTO.IDREMITO AND"
        Sql = Sql & vbCrLf & "   REMITOS_CUERPO.TIPO = REMITO_TIPO.ID AND"
        Sql = Sql & vbCrLf & "   REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID AND"
        Sql = Sql & vbCrLf & "   REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
        Sql = Sql & vbCrLf & "   REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
        Sql = Sql & vbCrLf & "   REMITOS_CUERPO.NRO_REMITO =" & NumeroRemito
        NombreReporte = PASOREMITOS & "\remito.rpt"
    Case 2, 4 ' LIBRO
            Sql = " SELECT"
            Sql = Sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.OBSERVACIONES,"
            Sql = Sql & vbCrLf & " REMITOS_DETALLE.DESDE,"
            Sql = Sql & vbCrLf & " REQUERIMIENTO.SECTOR, REQUERIMIENTO.SOLICITANTE, REQUERIMIENTO.FECHARECEPCION,"
            Sql = Sql & vbCrLf & " REMITO_TIPO.DESCRIPCION,"
            Sql = Sql & vbCrLf & " REMITO_OPERACION.DESCRIPCION,"
            Sql = Sql & vbCrLf & " REMITO_ESTADOS.DESCRIPCION,"
            Sql = Sql & vbCrLf & " LIBROS.NRO_LIBRO, LIBROS.NRO_LIBRO_INTERNO, LIBROS.COD_CLIENTE, LIBROS.ESTADO, LIBROS.REFERENCIA, LIBROS.AUDIT_USUARIO,"
            Sql = Sql & vbCrLf & " clientes.id_cliente , clientes.RAZON_SOCIAL, clientes.CALLE, clientes.NUMERO, clientes.LOCALIDAD"
            Sql = Sql & vbCrLf & " From"
            Sql = Sql & vbCrLf & " BASA.REMITOS_CUERPO REMITOS_CUERPO,"
            Sql = Sql & vbCrLf & " BASA.REMITOS_DETALLE REMITOS_DETALLE,"
            Sql = Sql & vbCrLf & " BASA.REQUERIMIENTO REQUERIMIENTO,"
            Sql = Sql & vbCrLf & " BASA.REMITO_TIPO REMITO_TIPO,"
            Sql = Sql & vbCrLf & " BASA.REMITO_OPERACION REMITO_OPERACION,"
            Sql = Sql & vbCrLf & " BASA.REMITO_ESTADOS REMITO_ESTADOS,"
            Sql = Sql & vbCrLf & " BASA.LIBROS LIBROS,"
            Sql = Sql & vbCrLf & " BASA.clientes clientes"
            Sql = Sql & vbCrLf & " Where"
            Sql = Sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO AND"
            Sql = Sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO = REQUERIMIENTO.IDREMITO AND"
            Sql = Sql & vbCrLf & " REMITOS_CUERPO.TIPO = REMITO_TIPO.ID AND"
            Sql = Sql & vbCrLf & " REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID AND"
            Sql = Sql & vbCrLf & " REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
            Sql = Sql & vbCrLf & " REMITOS_DETALLE.DESDE = LIBROS.NRO_LIBRO_INTERNO AND"
            Sql = Sql & vbCrLf & " REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
            Sql = Sql & vbCrLf & " CLIENTES.ID_CLIENTE = LIBROS.COD_CLIENTE AND"
            Sql = Sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO =" & NumeroRemito
            Sql = Sql & vbCrLf & " Order By"
            Sql = Sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO Asc"
            NombreReporte = PASOREMITOS & "\REMITOLIBROS.rpt"
        Case 10, 11  'LEGAJO
                Sql = "SELECT REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.OBSERVACIONES, REMITOS_DETALLE.DESDE, REMITO_TIPO.DESCRIPCION, REMITO_OPERACION.DESCRIPCION, REMITO_ESTADOS.DESCRIPCION, CLIENTES.ID_CLIENTE, CLIENTES.RAZON_SOCIAL, CLIENTES.CALLE, CLIENTES.NUMERO, CLIENTES.LOCALIDAD, LEGAJOS.CLIENTE_LEGAJO, LEGAJOS.DESCRIPCION "
                Sql = Sql & vbCrLf & " From  LEGAJOS, REMITOS_DETALLE REMITOS_DETALLE,REMITOS_CUERPO REMITOS_CUERPO,REMITO_TIPO REMITO_TIPO,REMITO_OPERACION REMITO_OPERACION,REMITO_ESTADOS REMITO_ESTADOS,CLIENTES CLIENTES"
                Sql = Sql & vbCrLf & " Where Legajos.ID_CLIENTE_LEGAJO = REMITOS_DETALLE.DESDE And REMITOS_DETALLE.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO "
                Sql = Sql & vbCrLf & " And Legajos.Cod_Cliente = REMITOS_CUERPO.ID_CLIENTE And REMITOS_CUERPO.Tipo = REMITO_TIPO.ID And "
                Sql = Sql & vbCrLf & " REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID AND REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND "
                Sql = Sql & vbCrLf & " REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND REMITOS_CUERPO.NRO_REMITO = " & NumeroRemito
                NombreReporte = PASOREMITOS & "\remitolegajos.rpt"
        End Select
        SQL1 = "        SELECT"
        SQL1 = SQL1 & vbCrLf & "    H_ESTADO_REQUE.IDREQUERIMIENTO,"
        SQL1 = SQL1 & vbCrLf & "    H_ESTADO_REQUE.IDESTADO,"
        SQL1 = SQL1 & vbCrLf & "    H_ESTADO_REQUE.CONTADOR,"
        SQL1 = SQL1 & vbCrLf & "    PERSONAL.NOMBRE,PERSONAL.APELLIDO"
        SQL1 = SQL1 & vbCrLf & "  From"
        SQL1 = SQL1 & vbCrLf & "    BASA.H_ESTADO_REQUE , PERSONAL, Requerimiento"
        SQL1 = SQL1 & vbCrLf & "  Where"
        SQL1 = SQL1 & vbCrLf & "    H_ESTADO_REQUE.idPersonal = PERSONAL.idPersonal"
        SQL1 = SQL1 & vbCrLf & "    AND H_ESTADO_REQUE.IDESTADO = REQUERIMIENTO.IDESTADO"
        SQL1 = SQL1 & vbCrLf & "    AND H_ESTADO_REQUE.IDREQUERIMIENTO = REQUERIMIENTO.IDREQUERIMIENTO"
        SQL1 = SQL1 & vbCrLf & "    AND H_ESTADO_REQUE.IDREQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
        SQL1 = SQL1 & vbCrLf & "    Order By"
        SQL1 = SQL1 & vbCrLf & "    H_ESTADO_REQUE.IDREQUERIMIENTO Asc"
        Set RS = New ADODB.Recordset
        RS.Open SQL1, ConBasa
        Do While Not RS.EOF
            If CLng(RS!IDREQUERIMIENTO) = ANTERIOR Then
                Responsables = Responsables & " , " & Trim(UCase(RS!NOMBRE)) & " " & Trim(UCase(RS!APELLIDO))
            Else
                If Bandera = False Then
                    ANTERIOR = RS!IDREQUERIMIENTO
                    Bandera = True
                    Responsables = Trim(UCase(RS!NOMBRE)) & " " & Trim(UCase(RS!APELLIDO))
                Else
                    Exit Do
                End If
            End If
            RS.MoveNext
        Loop
            
           frmReportes.ImprimirReporte NombreReporte, Sql, "f ='" & " : " & Responsables & "'", "COPIA ='" & "ORIGINAL" & " '"
           frmReportes.ImprimirReporte NombreReporte, Sql, "f ='" & " : " & Responsables & "'", "COPIA ='" & "DUPLICADO" & " '"
'
'
'
'
'
'
'
'            cryRemito.DiscardSavedData = True
'            cryRemito.Formulas(0) = "f ='" & " : " & Responsables & "'"
'            cryRemito.Formulas(1) = "COPIA ='" & "ORIGINAL" & " '"
'            cryRemito.SQLQuery = Sql
'            cryRemito.Destination = 1
'            cryRemito.Action = 1
'
'            cryRemito.DiscardSavedData = True
'            cryRemito.Formulas(0) = "f ='" & " : " & Responsables & "'"
'            cryRemito.Formulas(1) = "COPIA ='" & "DUPLICADO" & " '"
'            cryRemito.SQLQuery = Sql
'            cryRemito.Destination = 1
'            cryRemito.Action = 1
     Exit Sub
LERROR:
    MsgBox "Atencion error al imprimir el remito " & vbCrLf & "Por favor intentolo desde la aplicacion de control de estados", vbInformation, "Error de Impresion"
End Sub


Public Function Validar() As Boolean
Dim RS As ADODB.Recordset
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
   Select Case cboTipo_Almacenado.ItemData(cboTipo_Almacenado.ListIndex)
   Case 0
        Sql = " SELECT * FROM CONTENEDOR WHERE "
        Sql = Sql & "COD_CLIENTE = " & CInt(lblIDCliente.Caption) 'CAJA
        Sql = Sql & " AND NRO_CAJA IN  (" & Filtro & ")"
   Case Is = 1
        Sql = " SELECT ESTADO FROM LIBROS WHERE "
        Sql = Sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption) 'LIBRO
        Sql = Sql & " AND NRO_LIBRO_INTERNO  IN  (" & Filtro & ")"
   Case Is = 3
      Sql = " SELECT COD_ESTADO as ESTADO , ID_CLIENTE_LEGAJO"
      Sql = Sql & " From LEGAJOS Where ID_CLIENTE_LEGAJO IN  (" & Filtro & ") And COD_CLIENTE = " & CInt(lblIDCliente.Caption)
   End Select
   
   
    For I = 0 To lstPersonal.ListCount - 1
        If lstPersonal.Selected(I) Then
            Bandera = True
        End If
    Next
    If Not Bandera Then
        MsgBox "Quien Transporta?", vbQuestion
        Validar = False
        Exit Function
    End If
   
   
   Set RS = New ADODB.Recordset
   RS.Open Sql, ConBasa
        
       If cboTipoRemito.ItemData(cboTipoRemito.ListIndex) = 1 Then
            If cboRemito_Operacion.ItemData(cboRemito_Operacion.ListIndex) = 1 Then
                Do While Not RS.EOF
                    If CInt(RS!ESTADO) <> 2 Then
                        MsgBox "La Caja/Libro" & CStr(RS!NRO_CAJA) & "No tiene el estado Correcto"
                        Validar = False
                    End If
                RS.MoveNext
                Loop
            End If
        End If
        
        If cboTipoRemito.ItemData(cboTipoRemito.ListIndex) = 2 Then
                Do While Not RS.EOF
                    If CInt(RS!ESTADO) <> 4 Then
                        MsgBox "La Caja/Libro No tiene el estado Correcto"
                        Validar = False
                    End If
                RS.MoveNext
                Loop
            
        End If
        
    


End Function



Private Sub lblDescripcionRequerimiento_DblClick()
    txtObservaciones.Text = Trim(UCase(lblDescripcionRequerimiento.Caption))
End Sub
Public Sub Guardar_Remito()
Dim Sql As String
Dim R As Integer
Dim c As Integer
Dim oradyn As New ADODB.Recordset
Dim Proximo_Nro_Remito As Long

On Error GoTo OraError

    If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
            Screen.MousePointer = 11
            ConBasa.BeginTrans
            Proximo_Nro_Remito = ProximoRemito
            ' INSERTAR EN REMITO CUERPO
            Dim COD_TIPO_ALMACENAMIENTO As Integer
            Dim NRO_REMITO As Long
            Dim TIPO As Integer
            Dim OPERACION As Integer
            Dim ESTADO As Integer
            Dim Fecha As String
            Dim ID_CLIENTE As Integer
            Dim OBSERVACIONES As String
            Dim CANTIDAD As Integer
            Dim AUDIT_USUARIO As String
            Dim AUDIT_FECHA As String
            
            
            
            
             Select Case CRequerimientos.Item(1).TIPO
             Case 1, 3, 8, 9
                COD_TIPO_ALMACENAMIENTO = 0
             Case 2, 4
                COD_TIPO_ALMACENAMIENTO = 1
             Case 10, 11
                COD_TIPO_ALMACENAMIENTO = 3
             End Select
            
            NRO_REMITO = Proximo_Nro_Remito
            TIPO = cboTipoRemito.ItemData(cboTipoRemito.ListIndex)
            OPERACION = cboRemito_Operacion.ItemData(cboRemito_Operacion.ListIndex)
            ESTADO = cboRemito_Estados.ItemData(cboRemito_Estados.ListIndex)
            Fecha = " TO_DATE('" & maskFechaRemito.Text & "','DD/MM/YYYY')"
            ID_CLIENTE = lblIDCliente.Caption
            If txtObservaciones.Text <> "" Then
                OBSERVACIONES = "'" & UCase(txtObservaciones.Text) & "'"
            Else
                OBSERVACIONES = "NULL"
            End If
            CANTIDAD = lblCantidad.Caption
            AUDIT_USUARIO = "'" & UserName & "'"
            AUDIT_FECHA = SysDate
            REMITO_CUERPO_ADD NRO_REMITO, TIPO, OPERACION, ESTADO, Fecha, ID_CLIENTE, _
            OBSERVACIONES, CANTIDAD, AUDIT_USUARIO, AUDIT_FECHA, COD_TIPO_ALMACENAMIENTO


            ' CAMBIO DE ESTADO REQUERIMIENTO
            Dim IDPERSONAL As Integer
            Dim I As Integer
            Dim Bandera As Boolean
            Bandera = False
            Dim RS As New ADODB.Recordset
            Dim Filtro As String
            Dim FECHARECEPCION As Date
            Dim IDTIPOREQUERIMIENTO As Integer
            For I = 0 To lstPersonal.ListCount - 1
                IDPERSONAL = Mid(lstPersonal.List(I), 1, 2)
                If lstPersonal.Selected(I) Then
                    If Bandera = True Then
                        CambioEstadoRemito IDPERSONAL, False, 4, 6, lbRequerimiento
                    Else
                        CambioEstadoRemito IDPERSONAL, True, 4, 6, lbRequerimiento
                        Bandera = True
                    End If
                End If
            Next
            Sql = "Update requerimiento set idremito = " & Proximo_Nro_Remito
            Sql = Sql & vbCrLf & " where idrequerimiento =" & CRequerimientos.Item(1).NumeroRequerimiento
            ConBasa.Execute Sql
            
            'INSERTAR EN REMITO DELTALLE
            
             With grdCajasLibros
                For R = 1 To .Rows - 1
                    For c = 1 To .Cols - 1
                        If .TextMatrix(R, c) <> "" Then
                            REMITO_DETALLE_ADD NRO_REMITO, .TextMatrix(R, c), COD_TIPO_ALMACENAMIENTO
                            GrabarMovHistorico Proximo_Nro_Remito, .TextMatrix(R, c), lblIDCliente.Caption, .TextMatrix(R, c), TIPO, OPERACION, Fecha, COD_TIPO_ALMACENAMIENTO, AUDIT_USUARIO, SysDate
                        End If
                    Next
                Next
             
            
           'MOVIMIENTO EN TABLA CONTENEDO
           Select Case cboTipo_Almacenado.Text
           Case "CAJA"
                    Select Case cboTipoRemito.ItemData(cboTipoRemito.ListIndex)
                    Case 1 'CONSULTA
                        For R = 1 To grdCajasLibros.Rows - 1
                            For c = 1 To grdCajasLibros.Cols - 1
                                If grdCajasLibros.TextMatrix(R, c) <> "" Then
                                    Sql = "UPDATE CONTENEDOR SET "
                                    Sql = Sql & vbCrLf & " ESTADO = 3 "
                                    Sql = Sql & ", NRO_REMITO = " & Proximo_Nro_Remito
                                    Sql = Sql & ", F_MODIFICACION = " & SysDate
                                    Sql = Sql & vbCrLf & " WHERE "
                                    Sql = Sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption)
                                    Sql = Sql & " AND NRO_CAJA = " & CLng(grdCajasLibros.TextMatrix(R, c))
                                    Sql = Sql & " AND ESTADO = 2 "
                                    ConBasa.Execute Sql
                                End If
                            Next
                        Next
                    Case 2 'CAJAS VACIAS
                        For R = 1 To grdCajasLibros.Rows - 1
                            For c = 1 To grdCajasLibros.Cols - 1
                                If grdCajasLibros.TextMatrix(R, c) <> "" Then
                                    Sql = "UPDATE CONTENEDOR SET "
                                    Sql = Sql & vbCrLf & " ESTADO = 5 "
                                    Sql = Sql & vbCrLf & ", NRO_REMITO = " & Proximo_Nro_Remito
                                    Sql = Sql & vbCrLf & ", F_MODIFICACION = " & SysDate
                                    Sql = Sql & vbCrLf & " WHERE "
                                    Sql = Sql & vbCrLf & " COD_CLIENTE = " & CLng(lblIDCliente.Caption)
                                    Sql = Sql & vbCrLf & " AND NRO_CAJA = " & CLng(grdCajasLibros.TextMatrix(R, c))
                                    Sql = Sql & vbCrLf & " AND ESTADO = 4 "
                                    ConBasa.Execute Sql
                                End If
                            Next
                        Next
                    End Select
            Case "LIBRO"
                    For R = 1 To grdCajasLibros.Rows - 1
                        For c = 1 To grdCajasLibros.Cols - 1
                            If grdCajasLibros.TextMatrix(R, c) <> "" Then
                                Sql = "UPDATE LIBROS SET "
                                Sql = Sql & vbCrLf & " ESTADO = 3 "
                                Sql = Sql & ", NRO_REMITO = " & Proximo_Nro_Remito
                                Sql = Sql & ", AUDIT_FECHA = " & SysDate
                                Sql = Sql & ", AUDIT_USUARIO = '" & UserName
                                Sql = Sql & vbCrLf & "' WHERE "
                                Sql = Sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption)
                                Sql = Sql & " AND NRO_LIBRO_INTERNO = " & CLng(grdCajasLibros.TextMatrix(R, c))
                                Sql = Sql & " AND ESTADO = 2 "
                                ConBasa.Execute Sql
                            End If
                        Next
                    Next
             Case "LEGAJO"
                  For R = 1 To grdCajasLibros.Rows - 1
                        For c = 1 To grdCajasLibros.Cols - 1
                            If grdCajasLibros.TextMatrix(R, c) <> "" Then
                                    Sql = " Update LEGAJOS"
                                    Sql = Sql & vbCrLf & " SET COD_ESTADO = 3,"
                                    Sql = Sql & vbCrLf & "  COD_REMITO = " & Proximo_Nro_Remito
                                    Sql = Sql & vbCrLf & " , FECHA = " & SysDate
                                    Sql = Sql & vbCrLf & " Where COD_CLIENTE = " & CInt(lblIDCliente.Caption)
                                    Sql = Sql & vbCrLf & " And ID_CLIENTE_LEGAJO = " & CLng(grdCajasLibros.TextMatrix(R, c))
                                    Sql = Sql & vbCrLf & " AND COD_ESTADO = 2 "
                                    ConBasa.Execute Sql
                            End If
                        Next
                    Next
            End Select
            End With
            ConBasa.CommitTrans
            MsgBox "El remito fue grabado con exito", vbExclamation, "Remito"
            ImprimirRemito CLng(Proximo_Nro_Remito)
            Screen.MousePointer = 0
            On Error GoTo ErrorPrn
            frmControlEstados.CargarTree
            Unload Me
    End If
Exit Sub
OraError:
    Screen.MousePointer = 0
    ConBasa.RollbackTrans
    frmLogOraError.Show MODAL
    Exit Sub
    
ErrorPrn:
    MsgBox ERROR
    Exit Sub
    
End Sub


Public Sub REMITO_CUERPO_ADD(NRO_REMITO As Long, TIPO As Integer, OPERACION As Integer, ESTADO As Integer, Fecha As String, ID_CLIENTE As Integer, _
OBSERVACIONES As String, CANTIDAD As Integer, AUDIT_USUARIO As String, AUDIT_FECHA As String, COD_TIPO_ALMACENAMIENTO As Integer)
Dim Sql As String
    Sql = "INSERT INTO REMITOS_CUERPO (NRO_REMITO , TIPO , OPERACION, ESTADO, FECHA, ID_CLIENTE ,"
    Sql = Sql & vbCrLf & " OBSERVACIONES , CANTIDAD , AUDIT_USUARIO , AUDIT_FECHA , COD_TIPO_ALMACENAMIENTO )"
    Sql = Sql & vbCrLf & " VALUES (" & NRO_REMITO & "," & TIPO & "," & OPERACION & "," & ESTADO & "," & Fecha & "," & ID_CLIENTE & ","
    Sql = Sql & vbCrLf & OBSERVACIONES & "," & CANTIDAD & "," & AUDIT_USUARIO & "," & AUDIT_FECHA & "," & COD_TIPO_ALMACENAMIENTO & ")"
    ConBasa.Execute Sql

End Sub

Public Sub REMITO_DETALLE_ADD(NRO_REMITO As Long, DESDE As Long, TIPO_ALMACENADO As Integer)
    Dim Sql As String
    Sql = "INSERT INTO REMITOS_DETALLE  (NRO_REMITO, DESDE, HASTA,  NRO_CAJA, TIPO_ALMACENADO ,AUDIT_USUARIO, AUDIT_FECHA)"
    Sql = Sql & " VALUES (" & NRO_REMITO & "," & DESDE & "," & DESDE & "," & DESDE & "," & TIPO_ALMACENADO & ",'basa'," & SysDate & ")"
    ConBasa.Execute Sql
 End Sub
