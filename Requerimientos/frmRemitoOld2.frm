VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form frmRemito 
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
      ItemData        =   "frmRemitoOld2.frx":0000
      Left            =   6900
      List            =   "frmRemitoOld2.frx":0002
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
         ItemData        =   "frmRemitoOld2.frx":0004
         Left            =   1560
         List            =   "frmRemitoOld2.frx":0011
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cboTipoRemito 
         Height          =   315
         ItemData        =   "frmRemitoOld2.frx":002A
         Left            =   1560
         List            =   "frmRemitoOld2.frx":003A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboRemito_Estados 
         Height          =   315
         ItemData        =   "frmRemitoOld2.frx":0070
         Left            =   4920
         List            =   "frmRemitoOld2.frx":007A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1755
      End
      Begin VB.ComboBox cboRemito_Operacion 
         Height          =   315
         ItemData        =   "frmRemitoOld2.frx":008F
         Left            =   8100
         List            =   "frmRemitoOld2.frx":0099
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
   Begin Crystal.CrystalReport cryRemito 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   """\\Server1basa\Sistemas\Requerimientos\remito.rpt"""
      Connect         =   """DSN = bpdc;UID = "" & UserName & "";PWD = "" & Password"
      UserName        =   "UserName "
      PrintFileLinesPerPage=   60
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
Attribute VB_Name = "frmRemito"
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

Public Sub CambioEstadoRemito(IDEmpleado As Integer, ActualizaContador As Boolean, Optional EstadoInicial As Integer, Optional EstadoFinal As Integer, Optional Requerimiento As Integer)
    Dim rs As ADODB.Recordset
    Dim RSH_ESTADO_REQUE As ADODB.Recordset
    Dim sql As String
    Dim FECHARECEPCION As Date
    Dim IDTIPOREQUERIMIENTO As Integer
    Dim i As Integer
    Dim CONTADOR As Integer
    
   
    
            
            ' REQUERIMIENTO
            sql = " UPDATE REQUERIMIENTO SET "
            sql = sql & vbCrLf & " IDESTADO= " & EstadoFinal
            sql = sql & vbCrLf & ", IDPERSONAL = " & IDEmpleado
            sql = sql & vbCrLf & " WHERE idRequerimiento IN  " & Requerimiento
            sql = sql & vbCrLf & " AND IDESTADO = " & EstadoInicial
            conbasa.Execute (sql)
            
            ' CONTADOR
            sql = " SELECT max(Contador)AS CONTADOR From  H_ESTADO_REQUE  Where IDRequerimiento = " & Requerimiento
            Set RSH_ESTADO_REQUE = New ADODB.Recordset
            RSH_ESTADO_REQUE.Open sql, conbasa
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
            sql = " INSERT INTO H_ESTADO_REQUE ("
            sql = sql & vbCrLf & " IDREQUERIMIENTO, IDESTADO, IDPERSONAL,"
            sql = sql & vbCrLf & " CONTADOR, FECHA )"
            sql = sql & vbCrLf & "  VALUES ("
            sql = sql & vbCrLf & Requerimiento & "," & EstadoFinal & "," & IDEmpleado & ","
            sql = sql & vbCrLf & CONTADOR & "," & SysDate & ")"
            conbasa.Execute (sql)
   
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim rsPersonal As ADODB.Recordset
    Set rsPersonal = New ADODB.Recordset
    rsPersonal.Open "Select * from Personal WHERE NAVES=1 ", conbasa
    Do While Not rsPersonal.EOF
        lstPersonal.AddItem CStr(rsPersonal!IDPERSONAL) & " - " & Trim(CStr(rsPersonal!Nombre)) & " " & Trim(CStr(rsPersonal!Apellido))
        rsPersonal.MoveNext
    Loop
    CargarRemito
End Sub
Function ProximoRemito() As Long
  Dim sql As String
  Dim OraMax As ADODB.Recordset
  sql = "Select Max(Nro_Remito) Maximo From Remitos_Cuerpo"
  Set OraMax = New ADODB.Recordset
   OraMax.Open sql, conbasa
  If IsNull(OraMax("Maximo")) Then ProximoRemito = 1: Exit Function
  ProximoRemito = Val(OraMax("Maximo")) + 1
End Function
Sub GrabarMovHistorico(NRO_REMITO As Long, NRO_CAJA As Long, ID_CLIENTE As Integer, ELEMENTO As Long, Tipo As Integer, OPERACION As Integer, FECHA_MOVIMIENTO As String, TIPO_ELEMENTO As Integer, AUDIT_USUARIO As String, AUDIT_FECHA As String)
    Dim r As Single
    Dim sql As String
    sql = " INSERT INTO MOV_CAJAS2 "
    sql = sql & vbCrLf & "(NRO_REMITO, NRO_CAJA, ID_CLIENTE, ELEMENTO, TIPO,"
    sql = sql & vbCrLf & " OPERACION, FECHA_MOVIMIENTO, TIPO_ELEMENTO,"
    sql = sql & vbCrLf & " AUDIT_USUARIO, AUDIT_FECHA)"
    sql = sql & vbCrLf & " VALUES (" & NRO_REMITO & "," & NRO_CAJA & "," & ID_CLIENTE & "," & ELEMENTO & "," & Tipo & ","
    sql = sql & vbCrLf & OPERACION & "," & FECHA_MOVIMIENTO & "," & TIPO_ELEMENTO & ","
    sql = sql & vbCrLf & AUDIT_USUARIO & "," & AUDIT_FECHA & ")"
    conbasa.Execute sql

End Sub

Public Sub CargarRemito()
    Dim sql As String
    Dim rsRequerimiento As ADODB.Recordset
    Dim rsRcajas As ADODB.Recordset
        sql = "SELECT REQUERIMIENTO.IDREQUERIMIENTO, REQUERIMIENTO.SECTOR,"
        sql = sql & vbCrLf & " REQUERIMIENTO.TELEFONO, REQUERIMIENTO.ID_CLIENTE,"
        sql = sql & vbCrLf & " REQUERIMIENTO.DESCRIPCION, REQUERIMIENTO.SOLICITANTE, REQUERIMIENTO.TOMO,"
        sql = sql & vbCrLf & " REQUERIMIENTO.FECHALIMITE, "
        sql = sql & vbCrLf & " REQUERIMIENTO.FECHARECEPCION, REQUERIMIENTO.IDTIPORECEPCION, "
        sql = sql & vbCrLf & " REQUERIMIENTO.CANTIDAD, REQUERIMIENTO.IDESTADO,"
        sql = sql & vbCrLf & " REQUERIMIENTO.IDTIPOREQUERIMIENTO, Clientes.razon_social "
        sql = sql & vbCrLf & " From REQUERIMIENTO , Clientes"
        sql = sql & vbCrLf & " WHERE "
        sql = sql & vbCrLf & " REQUERIMIENTO.id_Cliente = Clientes.ID_Cliente and "
        sql = sql & vbCrLf & " REQUERIMIENTO.IDREQUERIMIENTO = " & CRequerimientos.Item(1).NumeroRequerimiento
       
       Set rsRequerimiento = New ADODB.Recordset
      rsRequerimiento.Open sql, conbasa
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
                sql = " SELECT "
                sql = sql & vbCrLf & " REQ.IDRequerimientos , REQ.CAJASLIBROS"
                sql = sql & vbCrLf & " From"
                sql = sql & vbCrLf & " REQUELIBOSCAJAS REQ"
                sql = sql & vbCrLf & " Where REQ.IDRequerimientos = " & CRequerimientos.Item(1).NumeroRequerimiento
                sql = sql & vbCrLf & " ORDER BY REQ.CAJASLIBROS"
               
                   Set rsRcajas = New ADODB.Recordset
                  rsRcajas.Open sql, conbasa
                   
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
Public Sub CargarGrilla(valor As String)
Dim C As Integer
Dim r As Integer


For r = 1 To grdCajasLibros.Rows - 1

    For C = 1 To grdCajasLibros.Cols - 1
        
        If grdCajasLibros.TextMatrix(r, C) = valor Then
            MsgBox "La Caja " & valor & " ya esta Cargada", vbInformation
            Exit Sub
        End If
        If grdCajasLibros.TextMatrix(r, C) = "" Then
            grdCajasLibros.TextMatrix(r, C) = valor
            Exit Sub
        End If
    Next
Next
grdCajasLibros.AddItem ""
grdCajasLibros.TextMatrix(grdCajasLibros.Rows - 1, 1) = valor
End Sub
Public Sub ImprimirRemito(NumeroRemito As Long)
    Dim sql As String
    Dim SQL1 As String
    Dim Bandera As Boolean
    Dim Responsables As String
    Dim rs  As ADODB.Recordset
    Dim ANTERIOR As Long
   On Error GoTo LERROR
    
    If NumeroRemito = 0 Or IsNull(NumeroRemito) Then
        Exit Sub
    End If
    
    cryRemito.Connect = "DSN = bpdc;UID = " & UserName & ";PWD = " & Password
    
    Select Case CRequerimientos.Item(1).Tipo
    Case 1, 3, 7, 8  ' CAJAS
        sql = "    SELECT"
        sql = sql & vbCrLf & "   REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.OBSERVACIONES,"
        sql = sql & vbCrLf & "   REMITOS_DETALLE.DESDE,"
        sql = sql & vbCrLf & "   REQUERIMIENTO.SECTOR, REQUERIMIENTO.SOLICITANTE, REQUERIMIENTO.FECHARECEPCION,"
        sql = sql & vbCrLf & "   REMITO_TIPO.DESCRIPCION,"
        sql = sql & vbCrLf & "   REMITO_OPERACION.DESCRIPCION,"
        sql = sql & vbCrLf & "   REMITO_ESTADOS.DESCRIPCION,"
        sql = sql & vbCrLf & "   clientes.id_cliente , clientes.RAZON_SOCIAL, clientes.CALLE, clientes.NUMERO, clientes.LOCALIDAD"
        sql = sql & vbCrLf & "   From "
        sql = sql & vbCrLf & "   BASA.REMITOS_CUERPO REMITOS_CUERPO,"
        sql = sql & vbCrLf & "   BASA.REMITOS_DETALLE REMITOS_DETALLE,"
        sql = sql & vbCrLf & "   BASA.REQUERIMIENTO REQUERIMIENTO,"
        sql = sql & vbCrLf & "   BASA.REMITO_TIPO REMITO_TIPO,"
        sql = sql & vbCrLf & "   BASA.REMITO_OPERACION REMITO_OPERACION,"
        sql = sql & vbCrLf & "   BASA.REMITO_ESTADOS REMITO_ESTADOS,"
        sql = sql & vbCrLf & "   BASA.clientes clientes"
        sql = sql & vbCrLf & " Where"
        sql = sql & vbCrLf & "   REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO AND"
        sql = sql & vbCrLf & "   REMITOS_CUERPO.NRO_REMITO = REQUERIMIENTO.IDREMITO AND"
        sql = sql & vbCrLf & "   REMITOS_CUERPO.TIPO = REMITO_TIPO.ID AND"
        sql = sql & vbCrLf & "   REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID AND"
        sql = sql & vbCrLf & "   REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
        sql = sql & vbCrLf & "   REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
        sql = sql & vbCrLf & "   REMITOS_CUERPO.NRO_REMITO =" & NumeroRemito
        cryRemito.ReportFileName = PASOREMITOS & "\remito.rpt"
    Case 2, 4 ' LIBRO
            sql = " SELECT"
            sql = sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.OBSERVACIONES,"
            sql = sql & vbCrLf & " REMITOS_DETALLE.DESDE,"
            sql = sql & vbCrLf & " REQUERIMIENTO.SECTOR, REQUERIMIENTO.SOLICITANTE, REQUERIMIENTO.FECHARECEPCION,"
            sql = sql & vbCrLf & " REMITO_TIPO.DESCRIPCION,"
            sql = sql & vbCrLf & " REMITO_OPERACION.DESCRIPCION,"
            sql = sql & vbCrLf & " REMITO_ESTADOS.DESCRIPCION,"
            sql = sql & vbCrLf & " LIBROS.NRO_LIBRO, LIBROS.NRO_LIBRO_INTERNO, LIBROS.COD_CLIENTE, LIBROS.ESTADO, LIBROS.REFERENCIA, LIBROS.AUDIT_USUARIO,"
            sql = sql & vbCrLf & " clientes.id_cliente , clientes.RAZON_SOCIAL, clientes.CALLE, clientes.NUMERO, clientes.LOCALIDAD"
            sql = sql & vbCrLf & " From"
            sql = sql & vbCrLf & " BASA.REMITOS_CUERPO REMITOS_CUERPO,"
            sql = sql & vbCrLf & " BASA.REMITOS_DETALLE REMITOS_DETALLE,"
            sql = sql & vbCrLf & " BASA.REQUERIMIENTO REQUERIMIENTO,"
            sql = sql & vbCrLf & " BASA.REMITO_TIPO REMITO_TIPO,"
            sql = sql & vbCrLf & " BASA.REMITO_OPERACION REMITO_OPERACION,"
            sql = sql & vbCrLf & " BASA.REMITO_ESTADOS REMITO_ESTADOS,"
            sql = sql & vbCrLf & " BASA.LIBROS LIBROS,"
            sql = sql & vbCrLf & " BASA.clientes clientes"
            sql = sql & vbCrLf & " Where"
            sql = sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO = REMITOS_DETALLE.NRO_REMITO AND"
            sql = sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO = REQUERIMIENTO.IDREMITO AND"
            sql = sql & vbCrLf & " REMITOS_CUERPO.TIPO = REMITO_TIPO.ID AND"
            sql = sql & vbCrLf & " REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID AND"
            sql = sql & vbCrLf & " REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND"
            sql = sql & vbCrLf & " REMITOS_DETALLE.DESDE = LIBROS.NRO_LIBRO_INTERNO AND"
            sql = sql & vbCrLf & " REQUERIMIENTO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND"
            sql = sql & vbCrLf & " CLIENTES.ID_CLIENTE = LIBROS.COD_CLIENTE AND"
            sql = sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO =" & NumeroRemito
            sql = sql & vbCrLf & " Order By"
            sql = sql & vbCrLf & " REMITOS_CUERPO.NRO_REMITO Asc"
            cryRemito.ReportFileName = PASOREMITOS & "\REMITOLIBROS.rpt"
        Case 10, 11  'LEGAJO
                sql = "SELECT REMITOS_CUERPO.NRO_REMITO, REMITOS_CUERPO.FECHA, REMITOS_CUERPO.OBSERVACIONES, REMITOS_DETALLE.DESDE, REMITO_TIPO.DESCRIPCION, REMITO_OPERACION.DESCRIPCION, REMITO_ESTADOS.DESCRIPCION, CLIENTES.ID_CLIENTE, CLIENTES.RAZON_SOCIAL, CLIENTES.CALLE, CLIENTES.NUMERO, CLIENTES.LOCALIDAD, LEGAJOS.CLIENTE_LEGAJO, LEGAJOS.DESCRIPCION "
                sql = sql & vbCrLf & " From  LEGAJOS, REMITOS_DETALLE REMITOS_DETALLE,REMITOS_CUERPO REMITOS_CUERPO,REMITO_TIPO REMITO_TIPO,REMITO_OPERACION REMITO_OPERACION,REMITO_ESTADOS REMITO_ESTADOS,CLIENTES CLIENTES"
                sql = sql & vbCrLf & " Where Legajos.ID_CLIENTE_LEGAJO = REMITOS_DETALLE.DESDE And REMITOS_DETALLE.NRO_REMITO = REMITOS_CUERPO.NRO_REMITO "
                sql = sql & vbCrLf & " And Legajos.Cod_Cliente = REMITOS_CUERPO.ID_CLIENTE And REMITOS_CUERPO.Tipo = REMITO_TIPO.ID And "
                sql = sql & vbCrLf & " REMITOS_CUERPO.OPERACION = REMITO_OPERACION.ID AND REMITOS_CUERPO.ESTADO = REMITO_ESTADOS.ID AND "
                sql = sql & vbCrLf & " REMITOS_CUERPO.ID_CLIENTE = CLIENTES.ID_CLIENTE AND REMITOS_CUERPO.NRO_REMITO = " & NumeroRemito
                cryRemito.ReportFileName = PASOREMITOS & "\remitolegajos.rpt"
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
        Set rs = New ADODB.Recordset
        rs.Open SQL1, conbasa
        Do While Not rs.EOF
            If CLng(rs!IDREQUERIMIENTO) = ANTERIOR Then
                Responsables = Responsables & " , " & Trim(UCase(rs!Nombre)) & " " & Trim(UCase(rs!Apellido))
            Else
                If Bandera = False Then
                    ANTERIOR = rs!IDREQUERIMIENTO
                    Bandera = True
                    Responsables = Trim(UCase(rs!Nombre)) & " " & Trim(UCase(rs!Apellido))
                Else
                    Exit Do
                End If
            End If
            rs.MoveNext
        Loop
            cryRemito.DiscardSavedData = True
            cryRemito.Formulas(0) = "f ='" & " : " & Responsables & "'"
            cryRemito.Formulas(1) = "COPIA ='" & "ORIGINAL" & " '"
            cryRemito.SQLQuery = sql
            cryRemito.Destination = 1
            cryRemito.Action = 1
            
            cryRemito.DiscardSavedData = True
            cryRemito.Formulas(0) = "f ='" & " : " & Responsables & "'"
            cryRemito.Formulas(1) = "COPIA ='" & "DUPLICADO" & " '"
            cryRemito.SQLQuery = sql
            cryRemito.Destination = 1
            cryRemito.Action = 1
     Exit Sub
LERROR:
    MsgBox "Atencion error al imprimir el remito " & vbCrLf & "Por favor intentolo desde la aplicacion de control de estados", vbInformation, "Error de Impresion"
End Sub


Public Function Validar() As Boolean
Dim rs As ADODB.Recordset
 Dim r As Integer
 Dim C As Integer
 Dim Filtro As String
 Dim sql As String
 Dim Bandera As Boolean
Dim i As Integer
 Bandera = False
   Validar = True
    For r = 1 To grdCajasLibros.Rows - 1
        For C = 1 To grdCajasLibros.Cols - 1
            If grdCajasLibros.TextMatrix(r, C) <> "" Then
                Filtro = Filtro & grdCajasLibros.TextMatrix(r, C) & ","
            End If
        Next
    Next
   Filtro = Mid(Filtro, 1, Len(Filtro) - 1)
   Select Case cboTipo_Almacenado.ItemData(cboTipo_Almacenado.ListIndex)
   Case 0
        sql = " SELECT * FROM CONTENEDOR WHERE "
        sql = sql & "COD_CLIENTE = " & CInt(lblIDCliente.Caption) 'CAJA
        sql = sql & " AND NRO_CAJA IN  (" & Filtro & ")"
   Case Is = 1
        sql = " SELECT ESTADO FROM LIBROS WHERE "
        sql = sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption) 'LIBRO
        sql = sql & " AND NRO_LIBRO_INTERNO  IN  (" & Filtro & ")"
   Case Is = 3
      sql = " SELECT COD_ESTADO as ESTADO , ID_CLIENTE_LEGAJO"
      sql = sql & " From LEGAJOS Where ID_CLIENTE_LEGAJO IN  (" & Filtro & ") And COD_CLIENTE = " & CInt(lblIDCliente.Caption)
   End Select
   
   
    For i = 0 To lstPersonal.ListCount - 1
        If lstPersonal.Selected(i) Then
            Bandera = True
        End If
    Next
    If Not Bandera Then
        MsgBox "Quien Transporta?", vbQuestion
        Validar = False
        Exit Function
    End If
   
   
   Set rs = New ADODB.Recordset
   rs.Open sql, conbasa
        
       If cboTipoRemito.ItemData(cboTipoRemito.ListIndex) = 1 Then
            If cboRemito_Operacion.ItemData(cboRemito_Operacion.ListIndex) = 1 Then
                Do While Not rs.EOF
                    If CInt(rs!ESTADO) <> 2 Then
                        MsgBox "La Caja/Libro" & CStr(rs!NRO_CAJA) & "No tiene el estado Correcto"
                        Validar = False
                    End If
                rs.MoveNext
                Loop
            End If
        End If
        
        If cboTipoRemito.ItemData(cboTipoRemito.ListIndex) = 2 Then
                Do While Not rs.EOF
                    If CInt(rs!ESTADO) <> 4 Then
                        MsgBox "La Caja/Libro No tiene el estado Correcto"
                        Validar = False
                    End If
                rs.MoveNext
                Loop
            
        End If
        
    


End Function



Private Sub lblDescripcionRequerimiento_DblClick()
    txtObservaciones.Text = Trim(UCase(lblDescripcionRequerimiento.Caption))
End Sub
Public Sub Guardar_Remito()
Dim sql As String
Dim r As Integer
Dim C As Integer
Dim oradyn As New ADODB.Recordset
Dim Proximo_Nro_Remito As Long

On Error GoTo OraError

    If MsgBox("Usted quiere grabar el remito", vbQuestion + vbYesNo, "Atención") = vbYes Then
            Screen.MousePointer = 11
            conbasa.BeginTrans
            Proximo_Nro_Remito = ProximoRemito
            ' INSERTAR EN REMITO CUERPO
            Dim COD_TIPO_ALMACENAMIENTO As Integer
            Dim NRO_REMITO As Long
            Dim Tipo As Integer
            Dim OPERACION As Integer
            Dim ESTADO As Integer
            Dim Fecha As String
            Dim ID_CLIENTE As Integer
            Dim OBSERVACIONES As String
            Dim CANTIDAD As Integer
            Dim AUDIT_USUARIO As String
            Dim AUDIT_FECHA As String
            
            
            
            
             Select Case CRequerimientos.Item(1).Tipo
             Case 1, 3, 8, 9
                COD_TIPO_ALMACENAMIENTO = 0
             Case 2, 4
                COD_TIPO_ALMACENAMIENTO = 1
             Case 10, 11
                COD_TIPO_ALMACENAMIENTO = 3
             End Select
            
            NRO_REMITO = Proximo_Nro_Remito
            Tipo = cboTipoRemito.ItemData(cboTipoRemito.ListIndex)
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
            REMITO_CUERPO_ADD NRO_REMITO, Tipo, OPERACION, ESTADO, Fecha, ID_CLIENTE, _
            OBSERVACIONES, CANTIDAD, AUDIT_USUARIO, AUDIT_FECHA, COD_TIPO_ALMACENAMIENTO


            ' CAMBIO DE ESTADO REQUERIMIENTO
            Dim IDPERSONAL As Integer
            Dim i As Integer
            Dim Bandera As Boolean
            Bandera = False
            Dim rs As New ADODB.Recordset
            Dim Filtro As String
            Dim FECHARECEPCION As Date
            Dim IDTIPOREQUERIMIENTO As Integer
            For i = 0 To lstPersonal.ListCount - 1
                IDPERSONAL = Mid(lstPersonal.List(i), 1, 2)
                If lstPersonal.Selected(i) Then
                    If Bandera = True Then
                        CambioEstadoRemito IDPERSONAL, False, 4, 6, lbRequerimiento
                    Else
                        CambioEstadoRemito IDPERSONAL, True, 4, 6, lbRequerimiento
                        Bandera = True
                    End If
                End If
            Next
            sql = "Update requerimiento set idremito = " & Proximo_Nro_Remito
            sql = sql & vbCrLf & " where idrequerimiento =" & CRequerimientos.Item(1).NumeroRequerimiento
            conbasa.Execute sql
            
            'INSERTAR EN REMITO DELTALLE
            
             With grdCajasLibros
                For r = 1 To .Rows - 1
                    For C = 1 To .Cols - 1
                        If .TextMatrix(r, C) <> "" Then
                            REMITO_DETALLE_ADD NRO_REMITO, .TextMatrix(r, C), COD_TIPO_ALMACENAMIENTO
                            GrabarMovHistorico Proximo_Nro_Remito, .TextMatrix(r, C), lblIDCliente.Caption, .TextMatrix(r, C), Tipo, OPERACION, Fecha, COD_TIPO_ALMACENAMIENTO, AUDIT_USUARIO, SysDate
                        End If
                    Next
                Next
             
            
           'MOVIMIENTO EN TABLA CONTENEDO
           Select Case cboTipo_Almacenado.Text
           Case "CAJA"
                    Select Case cboTipoRemito.ItemData(cboTipoRemito.ListIndex)
                    Case 1 'CONSULTA
                        For r = 1 To grdCajasLibros.Rows - 1
                            For C = 1 To grdCajasLibros.Cols - 1
                                If grdCajasLibros.TextMatrix(r, C) <> "" Then
                                    sql = "UPDATE CONTENEDOR SET "
                                    sql = sql & vbCrLf & " ESTADO = 3 "
                                    sql = sql & ", NRO_REMITO = " & Proximo_Nro_Remito
                                    sql = sql & ", F_MODIFICACION = " & SysDate
                                    sql = sql & vbCrLf & " WHERE "
                                    sql = sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption)
                                    sql = sql & " AND NRO_CAJA = " & CLng(grdCajasLibros.TextMatrix(r, C))
                                    sql = sql & " AND ESTADO = 2 "
                                    conbasa.Execute sql
                                End If
                            Next
                        Next
                    Case 2 'CAJAS VACIAS
                        For r = 1 To grdCajasLibros.Rows - 1
                            For C = 1 To grdCajasLibros.Cols - 1
                                If grdCajasLibros.TextMatrix(r, C) <> "" Then
                                    sql = "UPDATE CONTENEDOR SET "
                                    sql = sql & vbCrLf & " ESTADO = 5 "
                                    sql = sql & vbCrLf & ", NRO_REMITO = " & Proximo_Nro_Remito
                                    sql = sql & vbCrLf & ", F_MODIFICACION = " & SysDate
                                    sql = sql & vbCrLf & " WHERE "
                                    sql = sql & vbCrLf & " COD_CLIENTE = " & CLng(lblIDCliente.Caption)
                                    sql = sql & vbCrLf & " AND NRO_CAJA = " & CLng(grdCajasLibros.TextMatrix(r, C))
                                    sql = sql & vbCrLf & " AND ESTADO = 4 "
                                    conbasa.Execute sql
                                End If
                            Next
                        Next
                    End Select
            Case "LIBRO"
                    For r = 1 To grdCajasLibros.Rows - 1
                        For C = 1 To grdCajasLibros.Cols - 1
                            If grdCajasLibros.TextMatrix(r, C) <> "" Then
                                sql = "UPDATE LIBROS SET "
                                sql = sql & vbCrLf & " ESTADO = 3 "
                                sql = sql & ", NRO_REMITO = " & Proximo_Nro_Remito
                                sql = sql & ", AUDIT_FECHA = " & SysDate
                                sql = sql & ", AUDIT_USUARIO = '" & UserName
                                sql = sql & vbCrLf & "' WHERE "
                                sql = sql & " COD_CLIENTE = " & CInt(lblIDCliente.Caption)
                                sql = sql & " AND NRO_LIBRO_INTERNO = " & CLng(grdCajasLibros.TextMatrix(r, C))
                                sql = sql & " AND ESTADO = 2 "
                                conbasa.Execute sql
                            End If
                        Next
                    Next
             Case "LEGAJO"
                  For r = 1 To grdCajasLibros.Rows - 1
                        For C = 1 To grdCajasLibros.Cols - 1
                            If grdCajasLibros.TextMatrix(r, C) <> "" Then
                                    sql = " Update LEGAJOS"
                                    sql = sql & vbCrLf & " SET COD_ESTADO = 3,"
                                    sql = sql & vbCrLf & "  COD_REMITO = " & Proximo_Nro_Remito
                                    sql = sql & vbCrLf & " , FECHA = " & SysDate
                                    sql = sql & vbCrLf & " Where COD_CLIENTE = " & CInt(lblIDCliente.Caption)
                                    sql = sql & vbCrLf & " And ID_CLIENTE_LEGAJO = " & CLng(grdCajasLibros.TextMatrix(r, C))
                                    sql = sql & vbCrLf & " AND COD_ESTADO = 2 "
                                    conbasa.Execute sql
                            End If
                        Next
                    Next
            End Select
            End With
            conbasa.CommitTrans
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
    conbasa.RollbackTrans
    frmLogOraError.Show MODAL
    Exit Sub
    
ErrorPrn:
    MsgBox ERROR
    Exit Sub
    
End Sub


Public Sub REMITO_CUERPO_ADD(NRO_REMITO As Long, Tipo As Integer, OPERACION As Integer, ESTADO As Integer, Fecha As String, ID_CLIENTE As Integer, _
OBSERVACIONES As String, CANTIDAD As Integer, AUDIT_USUARIO As String, AUDIT_FECHA As String, COD_TIPO_ALMACENAMIENTO As Integer)
Dim sql As String
    sql = "INSERT INTO REMITOS_CUERPO (NRO_REMITO , TIPO , OPERACION, ESTADO, FECHA, ID_CLIENTE ,"
    sql = sql & vbCrLf & " OBSERVACIONES , CANTIDAD , AUDIT_USUARIO , AUDIT_FECHA , COD_TIPO_ALMACENAMIENTO )"
    sql = sql & vbCrLf & " VALUES (" & NRO_REMITO & "," & Tipo & "," & OPERACION & "," & ESTADO & "," & Fecha & "," & ID_CLIENTE & ","
    sql = sql & vbCrLf & OBSERVACIONES & "," & CANTIDAD & "," & AUDIT_USUARIO & "," & AUDIT_FECHA & "," & COD_TIPO_ALMACENAMIENTO & ")"
    conbasa.Execute sql

End Sub

Public Sub REMITO_DETALLE_ADD(NRO_REMITO As Long, DESDE As Long, TIPO_ALMACENADO As Integer)
    Dim sql As String
    sql = "INSERT INTO REMITOS_DETALLE  (NRO_REMITO, DESDE, HASTA,  NRO_CAJA, TIPO_ALMACENADO ,AUDIT_USUARIO, AUDIT_FECHA)"
    sql = sql & " VALUES (" & NRO_REMITO & "," & DESDE & "," & DESDE & "," & DESDE & "," & TIPO_ALMACENADO & ",'basa'," & SysDate & ")"
    conbasa.Execute sql
 End Sub
