VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D59D5BAF-9D93-48D8-8248-71EA7498F357}#1.0#0"; "Controles4.ocx"
Begin VB.Form frmUsuariosClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuarios Clientes"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   11445
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10440
      TabIndex        =   30
      Top             =   600
      Width           =   795
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8580
      TabIndex        =   29
      Top             =   600
      Width           =   795
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9480
      TabIndex        =   28
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      TabIndex        =   27
      Top             =   600
      Width           =   795
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6780
      TabIndex        =   26
      Top             =   600
      Width           =   795
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5760
      TabIndex        =   24
      Top             =   600
      Width           =   915
   End
   Begin Controles.ctlClienteUsuario ctlClienteUsuario 
      Height          =   315
      Left            =   6900
      TabIndex        =   18
      Top             =   60
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   556
   End
   Begin Controles.cltIndice ctlIndiceUsuario 
      Height          =   4095
      Left            =   60
      TabIndex        =   16
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7223
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataGridLib.DataGrid grdNivel 
      Height          =   2775
      Left            =   60
      TabIndex        =   11
      Top             =   4680
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   4895
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   18
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   5820
      TabIndex        =   0
      Top             =   1080
      Width           =   5355
      Begin VB.TextBox txtfecha_Nac 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   3
         EndProperty
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
         Left            =   1200
         TabIndex        =   25
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox txtTelefono 
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
         Left            =   1200
         TabIndex        =   22
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox txtUsuario 
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
         Left            =   1200
         TabIndex        =   15
         Top             =   2820
         Width           =   3855
      End
      Begin VB.TextBox txtReferencias 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1980
         Width           =   435
      End
      Begin VB.TextBox txtCod_Indice 
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1980
         Width           =   1635
      End
      Begin VB.TextBox txtCorreo 
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
         Left            =   1200
         TabIndex        =   7
         Top             =   1140
         Width           =   3855
      End
      Begin VB.TextBox txtDocumento 
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
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtApellidoNombre 
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
         Left            =   1200
         TabIndex        =   3
         Top             =   300
         Width           =   3855
      End
      Begin VB.Label Label13 
         Caption         =   "Telefono:"
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
         Left            =   60
         TabIndex        =   23
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label Label9 
         Caption         =   "Usuario"
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
         Left            =   60
         TabIndex        =   14
         Top             =   2820
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Referencia Envio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3120
         TabIndex        =   12
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "Nivel"
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
         Left            =   60
         TabIndex        =   9
         Top             =   2040
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Nac:"
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
         Left            =   60
         TabIndex        =   8
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Correo:"
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
         Left            =   60
         TabIndex        =   6
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "Documento:"
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
         Left            =   60
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Apellido y Nombre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label1 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   3795
      End
   End
   Begin Controles.cltGenerico ctlCliente 
      Height          =   315
      Left            =   840
      TabIndex        =   17
      Top             =   60
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   556
   End
   Begin VB.Label Label12 
      Caption         =   "Resultado Busqueda"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   21
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Responsable:"
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
      Left            =   5640
      TabIndex        =   20
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label10 
      Caption         =   "Cliente:"
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
      Left            =   120
      TabIndex        =   19
      Top             =   60
      Width           =   615
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuVertodas 
         Caption         =   "VerTodas"
      End
      Begin VB.Menu mnuVerNivel 
         Caption         =   "Ver Nivel"
      End
      Begin VB.Menu mnuPosteriores 
         Caption         =   "Todas Posteriores"
      End
      Begin VB.Menu MaxNivel 
         Caption         =   "Max Nivel"
      End
      Begin VB.Menu nmuCopiarNivel 
         Caption         =   "CopiarNivel"
      End
      Begin VB.Menu mnuUsuarioPassWord 
         Caption         =   "Cambio de Usuario y Pasword"
      End
      Begin VB.Menu mnuTomarCOrreos 
         Caption         =   "Tomar Correos"
      End
      Begin VB.Menu mnuReporte 
         Caption         =   "Reporte"
      End
      Begin VB.Menu mnuBuscar 
         Caption         =   "Buscar"
      End
   End
   Begin VB.Menu mnuGrilla 
      Caption         =   "Grilla"
      Begin VB.Menu mnuGrillaCopiar 
         Caption         =   "Copiar"
      End
   End
End
Attribute VB_Name = "frmUsuariosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SqlBase As String
Dim RsUsuario As ADODB.Recordset
Dim RsNuevo As ADODB.Recordset
Rem Dim WithEvents adoPrimaryRS As Recordset

Private Sub ctlIndice1_PopupMenuAction()
    PopupMenu mnuMenu
End Sub
Private Sub cmdAceptar_Click()
    RsUsuario.Update
    RsUsuario.Requery
End Sub

Private Sub cmdBuscar_Click()
        Dim Sql As String
        Sql = SqlBase
        Sql = Sql & vbCrLf & " WHERE ID_CLIENTEUSUARIO = " & ctlClienteUsuario.Valor
        Sql = Sql & vbCrLf & " AND  COD_CLIENTE =" & ctlCliente.Valor
        Sql = Sql & vbCrLf & " ORDER BY APELLIDO_NOMBRE "
        Set RsUsuario = New ADODB.Recordset
        RsUsuario.CursorLocation = adUseClient
        RsUsuario.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        DATOSGRILLA grdNivel, RsUsuario
        SetControles RsUsuario
End Sub

Private Sub cmdCancelar_Click()
    RsUsuario.CancelUpdate
End Sub

Private Sub cmdExcel_Click()
CopiarDatosGrilla grdNivel
End Sub

Private Sub cmdNuevo_Click()
        Dim rsMax As New ADODB.Recordset
        Dim MaxID As Integer
        Dim Cliente As String
        Dim correo As String
        Dim Indice As String
        Dim ApellidoNombre As String
        Dim TELEFONO As String
        Dim Sql As String
        
        Sql = "SELECT MAX(ID_CLIENTEUSUARIO) AS maxID From CLIENTEUSUARIO"
        rsMax.Open Sql, ConActiva, 0, 1
        MaxID = rsMax!MaxID + 1
        If IsNull(ctlCliente.Valor) Then
            MsgBox "Ingrese el cliente"
            Exit Sub
        Else
            Cliente = "'" & ctlCliente.Valor & "'"
        End If
        
        If Len(txtApellidoNombre.Text) < 8 Then
            MsgBox "Nombre Incorrecto"
            Exit Sub
        Else
            ApellidoNombre = "'" & (txtApellidoNombre.Text) & "'"
        End If
        
        
        If Trim(txtCod_Indice.Text) = "" Then
            MsgBox "Indice Incorrecto"
            Exit Sub
       Else
            If Trim(txtCod_Indice.Text) = "AIZ" Then
             MsgBox "ERROR EN INDICE"
             Exit Sub
            Else
                Indice = "'" & txtCod_Indice.Text & "'"
            End If
            
        End If
        
              
        If txtCorreo.Text = "" Then
            correo = "Null"
        Else
            correo = "'" & Trim(txtCorreo.Text) & "'"
        End If
        
        
        If txtTelefono.Text = "" Then
            TELEFONO = "NULL"
        Else
            TELEFONO = "'" & txtTelefono.Text & "'"
        End If
        
        Sql = " INSERT INTO CLIENTEUSUARIO "
        Sql = Sql & vbCrLf & " (ID_CLIENTEUSUARIO, COD_CLIENTE"
        Sql = Sql & vbCrLf & " , APELLIDO_NOMBRE,CORREO "
        Sql = Sql & vbCrLf & " , COD_INDICE, TELEFONOS) "
        Sql = Sql & vbCrLf & " VALUES "
        Sql = Sql & vbCrLf & " (" & MaxID & "," & Cliente
        Sql = Sql & vbCrLf & "," & ApellidoNombre & "," & correo
        Sql = Sql & vbCrLf & "," & Indice & "," & TELEFONO & ")"
        ExecutarSql Sql
        Unload Me
        End Sub

Private Sub Command2_Click()
    RsNuevo.Update
End Sub



Private Sub ctlCliente_Click()
    If Not IsNull(ctlCliente.Valor) Then
        ctlIndiceUsuario.Actualizar ctlCliente.Valor, Sector, 0
        grdNivel.ClearFields
        ctlClienteUsuario.LlenarConCliente ctlCliente.Valor
         ctlClienteUsuario.LlenarConCliente ctlCliente.Valor
    End If
End Sub

Private Sub ctlClienteUsuario1_SectorEncontrado(Sector As String)

End Sub

Private Sub ctlIndiceUsuario_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
        PopupMenu mnuMenu
   End If
End Sub

Private Sub Form_Load()
    ctlCliente.TipoControl = Cliente
    
    SqlBase = " SELECT ID_CLIENTEUSUARIO as ID , APELLIDO_NOMBRE as NOMBRE, CORREO, TELEFONOS ,COD_INDICE AS INDICE, REFERENCIAS AS ENVIO_REF "
    
     SqlBase = " SELECT ID_CLIENTEUSUARIO, COD_CLIENTE , APELLIDO_NOMBRE, correo, Cod_Indice,"
     SqlBase = SqlBase & " DOCUMENTO, TELEFONOS, FECHA_NAC,    REFERENCIAS , USUARIO  ,  DESHABILITADO, FECHA_ENVIO_REFERENCIAS, CONOCIMIENTO_DICCIONARIO "
       
    SqlBase = SqlBase & vbCrLf & " From CLIENTEUSUARIO "
End Sub

Private Sub grdNivel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnuGrilla
End If
 
End Sub

Private Sub mnuBuscar_Click()
ctlIndiceUsuario.BuscarIndice InputBox("Ingrese la Busqueda"), True
End Sub

Private Sub mnuGrillaCopiar_Click()
Rem CopiarDatosGrilla grdNivel
End Sub

Private Sub mnuPosteriores_Click()
        Dim Sql As String
        Sql = SqlBase & vbCrLf & " WHERE COD_INDICE like '" & ctlIndiceUsuario.Item_Selecionado & "%'"
        Sql = Sql & vbCrLf & " AND  COD_CLIENTE =" & ctlCliente.Valor
        Sql = Sql & vbCrLf & " Order BY APELLIDO_NOMBRE "
        Set RsUsuario = New ADODB.Recordset
        RsUsuario.CursorLocation = adUseClient
        RsUsuario.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        DATOSGRILLA grdNivel, RsUsuario
        SetControles RsUsuario
End Sub

Private Sub mnuReporte_Click()
Dim Sql As String
Sql = " SELECT     CLIENTEUSUARIO.ID_CLIENTEUSUARIO, CLIENTEUSUARIO.COD_CLIENTE, INDICES.ID_CODIGO_DOCUMENTO, INDICES.DESCRIPCION,"
Sql = Sql & vbCrLf & " CLIENTEUSUARIO.APELLIDO_NOMBRE , CLIENTEUSUARIO.correo, CLIENTEUSUARIO.REFERENCIAS, CLIENTEUSUARIO.DESHABILITADO"
Sql = Sql & vbCrLf & " FROM CLIENTEUSUARIO INNER JOIN"
Sql = Sql & vbCrLf & " INDICES ON CLIENTEUSUARIO.COD_CLIENTE = INDICES.COD_CLIENTE AND CLIENTEUSUARIO.COD_INDICE = INDICES.INDICE"
Sql = Sql & vbCrLf & " Where CLIENTEUSUARIO.COD_CLIENTE = " & ctlCliente.Valor
Sql = Sql & vbCrLf & " ORDER BY CLIENTEUSUARIO.APELLIDO_NOMBRE "
Set RsUsuario = New ADODB.Recordset

RsUsuario.Open Sql, ConActiva, adOpenStatic, adLockReadOnly
 DATOSGRILLA grdNivel, RsUsuario
 
End Sub

Private Sub mnuTomarCOrreos_Click()
 Dim rs As New ADODB.Recordset
    Dim Sql As String
    Dim correo As String
        Sql = "   SELECT COD_CLIENTE, APELLIDO_NOMBRE, CORREO,"
        Sql = Sql & " Cod_Indice , ID_CLIENTEUSUARIO"
        Sql = Sql & "  From CLIENTEUSUARIO"
        Sql = Sql & vbCrLf & " WHERE COD_INDICE like '" & ctlIndiceUsuario.Item_Selecionado & "%'"
        Sql = Sql & vbCrLf & " AND  COD_CLIENTE =" & ctlCliente.Valor
        rs.Open Sql, ConActiva, 0, 1
        Do While Not rs.EOF
            correo = correo & rs!correo & " ; "
            rs.MoveNext
        Loop
        Clipboard.Clear
        Clipboard.SetText correo
        MsgBox "Correos Copiados"
        
End Sub

Private Sub mnuUsuarioPassWord_Click()
        Dim rs As New ADODB.Recordset
        rs.CursorLocation = adUseClient
        Dim Sql As String
        Sql = " SELECT ID_CLIENTEUSUARIO, APELLIDO_NOMBRE, USUARIO, Password  FROM CLIENTEUSUARIO  "
        Sql = Sql & " WHERE "
        Sql = Sql & vbCrLf & "  COD_CLIENTE =" & ctlCliente.Valor
        Sql = Sql & vbCrLf & " Order BY APELLIDO_NOMBRE "
        rs.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
        DATOSGRILLA grdNivel, rs
        
       
End Sub

Private Sub mnuVerNivel_Click()
        
        Dim Sql As String
        Sql = SqlBase & vbCrLf & " WHERE COD_INDICE like '" & ctlIndiceUsuario.Item_Selecionado & "'"
        Sql = Sql & vbCrLf & " AND  COD_CLIENTE =" & ctlCliente.Valor
        Sql = Sql & vbCrLf & " Order BY APELLIDO_NOMBRE "
Set RsUsuario = New ADODB.Recordset
RsUsuario.CursorLocation = adUseClient
RsUsuario.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
DATOSGRILLA grdNivel, RsUsuario
SetControles RsUsuario

End Sub

Private Sub mnuVertodas_Click()
Dim Sql As String

Sql = SqlBase & vbCrLf & " WHERE COD_CLIENTE =" & ctlCliente.Valor
Sql = Sql & vbCrLf & " Order BY APELLIDO_NOMBRE "
Set RsUsuario = New ADODB.Recordset
RsUsuario.CursorLocation = adUseClient
RsUsuario.Open Sql, ConActiva, adOpenDynamic, adLockOptimistic
DATOSGRILLA grdNivel, RsUsuario
SetControles RsUsuario

End Sub

Private Sub nmuCopiarNivel_Click()
    txtCod_Indice.Text = ctlIndiceUsuario.Item_Selecionado
    txtCod_Indice.SetFocus
End Sub
Public Sub DATOSGRILLA(Grilla As DataGrid, rs As ADODB.Recordset)
Grilla.ClearFields
Grilla.ClearSelCols
Grilla.ScrollBars = dbgAutomatic
Dim i As Integer
For i = 0 To rs.Fields.Count - 1
    
    Debug.Print rs.Fields.Item(i).Name & "  " & rs.Fields.Item(i).Type
    
    Grilla.Columns.Add i
    Grilla.Columns.Item(i).DataField = rs.Fields(i).Name
    Grilla.Columns.Item(i).Caption = rs.Fields(i).Name
    Select Case rs.Fields.Item(i).Type
    Case "131" ' NUMERO
        Grilla.Columns.Item(i).Width = 500
    Case "200" 'TEXT
        Grilla.Columns.Item(i).Width = 1500
    Case "135" 'FECHA
        Grilla.Columns.Item(i).Width = 700
    End Select
    
Next

Set Grilla.DataSource = rs.DataSource
Grilla.Refresh


End Sub


Public Function ControlDuplicidad() As Boolean
    Dim rsDuplicidad As ADODB.Recordset
    Set rsDuplicidad = New ADODB.Recordset
    Dim Sql As String
    ControlDuplicidad = True
    Sql = " SELECT * "
    Sql = Sql & vbCrLf & " From CLIENTEUSUARIO"
    Sql = Sql & vbCrLf & " where APELLIDO_NOMBRE like '%" & Trim(txtApellidoNombre.Text) & "%'"
    Sql = Sql & vbCrLf & " AND COD_CLIENTE =" & ctlCliente.Valor
    rsDuplicidad.Open Sql, ConActiva, 0, 1
    Do While Not rsDuplicidad.EOF
       If MsgBox("Ya existe " & rsDuplicidad!APELLIDO_NOMBRE & vbCrLf & "Usted desea continuar", vbYesNo) = vbYes Then
       Else
            ControlDuplicidad = False
            Exit Function
       End If
       rsDuplicidad.MoveNext
    Loop
    

End Function


Public Sub SetControles(rs As ADODB.Recordset)
        Set txtApellidoNombre.DataSource = rs.DataSource
        txtApellidoNombre.DataField = "APELLIDO_NOMBRE"
        
        Set txtCorreo.DataSource = rs.DataSource
        txtCorreo.DataField = "CORREO"
        
        Set txtCod_Indice.DataSource = rs.DataSource
        txtCod_Indice.DataField = "COD_INDICE"
        
        Set txtDocumento.DataSource = rs.DataSource
        txtDocumento.DataField = "DOCUMENTO"
        
        Set txtTelefono.DataSource = rs.DataSource
        txtTelefono.DataField = "TELEFONOS"
        
        
        Set txtfecha_Nac.DataSource = rs.DataSource
        txtfecha_Nac.DataField = "FECHA_NAC"
            
        Set txtUsuario.DataSource = rs.DataSource
        txtUsuario.DataField = "USUARIO"
        
        Set txtReferencias.DataSource = rs.DataSource
            txtReferencias.DataField = "REFERENCIAS"
            
        
End Sub
