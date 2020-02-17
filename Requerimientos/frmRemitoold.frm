VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK~1.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRemito 
   Caption         =   "Remito"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Observaciones 
      Enabled         =   0   'False
      Height          =   732
      Left            =   240
      TabIndex        =   21
      Top             =   5880
      Width           =   8415
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   8760
      TabIndex        =   20
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   8760
      TabIndex        =   19
      Top             =   6300
      Width           =   1215
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
      TabIndex        =   17
      Top             =   60
      Width           =   3135
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
         Left            =   360
         TabIndex        =   18
         Top             =   300
         Width           =   2535
      End
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
      TabIndex        =   9
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
         Height          =   435
         Left            =   840
         TabIndex        =   14
         Top             =   300
         Width           =   5535
      End
   End
   Begin VB.ListBox lstPersonal 
      Height          =   2535
      ItemData        =   "frmRemito.frx":0000
      Left            =   7020
      List            =   "frmRemito.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   2460
      Width           =   2895
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
         ItemData        =   "frmRemito.frx":0004
         Left            =   1560
         List            =   "frmRemito.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cboTipoRemito 
         Height          =   315
         ItemData        =   "frmRemito.frx":001F
         Left            =   1560
         List            =   "frmRemito.frx":002F
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox cboRemito_Estados 
         Height          =   315
         ItemData        =   "frmRemito.frx":0065
         Left            =   4920
         List            =   "frmRemito.frx":006F
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1755
      End
      Begin VB.ComboBox cboRemito_Operacion 
         Height          =   315
         ItemData        =   "frmRemito.frx":0084
         Left            =   8100
         List            =   "frmRemito.frx":008E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   1515
      End
      Begin MSMask.MaskEdBox maskFechaRemito 
         Height          =   330
         Left            =   8100
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblCantidad 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   5340
         TabIndex        =   13
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
         TabIndex        =   11
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
         TabIndex        =   10
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
      Height          =   2475
      Left            =   180
      TabIndex        =   0
      Top             =   2460
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4366
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
   Begin VB.Label lblDescripcionRequerimiento 
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Left            =   240
      TabIndex        =   22
      Top             =   5100
      Width           =   9735
   End
End
Attribute VB_Name = "frmRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
MsgBox cboRemito_Estados.List(cboRemito_Estados.ListIndex)
MsgBox cboRemito_Estados.ItemData(cboRemito_Estados.ListIndex)
End Sub

Private Sub cmdAceptar_Click()
 If Not ValidForm Then
        MsgBox "Hay error en el remito y no pudo ser guardado", vbExclamation, "ERROR"
        Exit Sub
    End If
    Guardar_Remito
ImprimirRemito Proximo_Nro_Remito
frmControlEstados.CargarTree
Unload Me
End Sub

Private Sub Form_Load()
    Dim rsPersonal As OraDynaset
    Set rsPersonal = OraDatabase.CreateDynaset("Select * from Personal WHERE NAVES=1   ", ORADYN_READONLY)
    Do While Not rsPersonal.EOF
        lstPersonal.AddItem CStr(rsPersonal!IDPERSONAL) & " - " & CStr(rsPersonal!Nombre)
        rsPersonal.MoveNext
    Loop
   

CargarRemito
End Sub

Public Sub CargarRemito()
 Dim Sql As String
    Dim rsRequerimiento As OraDynaset
    Dim rsRcajas As OraDynaset
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
    Set rsRequerimiento = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
    
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
           End Select
            maskFechaRemito.Text = Format(SysDateCompare, "dd/mm/yyyy")
            lblCliente.Caption = Trim(UCase(!RAZON_SOCIAL))
            lbRequerimiento.Caption = UCase(!IDREQUERIMIENTO)
            If Not IsNull(!DESCRIPCION) Then
               ' lblDescripcionRequerimiento = UCase(!DESCRIPCION)
            End If
                Sql = " SELECT "
                Sql = Sql & vbCrLf & " REQ.IDRequerimientos , REQ.CAJASLIBROS"
                Sql = Sql & vbCrLf & " From"
                Sql = Sql & vbCrLf & " REQUELIBOSCAJAS REQ"
                Sql = Sql & vbCrLf & " Where REQ.IDRequerimientos = " & CRequerimientos.Item(1).NumeroRequerimiento
                Sql = Sql & vbCrLf & " ORDER BY REQ.CAJASLIBROS"
           
           Set rsRcajas = OraDatabase.CreateDynaset(Sql, ORADYN_READONLY)
           Do While Not rsRcajas.EOF
           
                If Not IsNull(rsRcajas!CAJASLIBROS) Then
                 CargarGrilla CStr(rsRcajas!CAJASLIBROS)
'                    Grilla.TextMatrix(Grilla.Rows - 1, 2) = CStr(rsRcajas!CAJASLIBROS)
'                    Grilla.TextMatrix(Grilla.Rows - 1, 3) = CStr(rsRcajas!CAJASLIBROS)
'                    Grilla.AddItem ("")
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

